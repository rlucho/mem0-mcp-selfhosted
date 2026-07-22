"""Minimal MS-CFB (OLE2) reader + writer.

Rebuilds a compound file preserving every directory entry verbatim (names,
object types, tree sibling/child links, CLSIDs, timestamps) and only relaying
stream data + patching each entry's Starting-Sector and Stream-Size fields.
Because the directory topology is untouched, no red-black-tree surgery is
needed. Validated by identity round-trip against the original.
"""
import struct

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
DIFSECT    = 0xFFFFFFFC
SECT = 512
MINI = 64
MINI_CUTOFF = 4096

# ---------------------------------------------------------------- reader ----
class Cfb:
    def __init__(self, blob: bytes):
        self.blob = blob
        assert blob[:8] == bytes.fromhex("d0cf11e0a1b11ae1"), "not a CFB"
        (self.minor, self.major, self.bo, self.ssz, self.mssz) = struct.unpack_from("<HHHHH", blob, 0x18)
        assert self.ssz == 9, "only 512-byte sectors supported"
        (self.n_dir, self.n_fat, self.first_dir, self.trans,
         self.mini_cutoff, self.first_minifat, self.n_minifat,
         self.first_difat, self.n_difat) = struct.unpack_from("<IIIIIIIII", blob, 0x28)
        self.difat = list(struct.unpack_from("<109I", blob, 0x4C))
        self._read_fat()
        self._read_minifat()
        self._read_dir()

    def _sector(self, n):
        off = 512 + n * SECT
        return self.blob[off:off + SECT]

    def _read_fat(self):
        # gather all FAT sector numbers (109 in header DIFAT, then DIFAT chain)
        fat_sects = [s for s in self.difat if s != FREESECT]
        ds = self.first_difat
        while ds != ENDOFCHAIN and ds != FREESECT:
            data = self._sector(ds)
            vals = list(struct.unpack("<128I", data))
            fat_sects += [s for s in vals[:127] if s != FREESECT]
            ds = vals[127]
        fat = []
        for s in fat_sects:
            fat += list(struct.unpack("<128I", self._sector(s)))
        self.fat = fat

    def _chain(self, start):
        out = []
        s = start
        while s != ENDOFCHAIN and s != FREESECT:
            out.append(s)
            s = self.fat[s]
        return out

    def _read_stream_fat(self, start, size):
        data = b"".join(self._sector(s) for s in self._chain(start))
        return data[:size]

    def _read_minifat(self):
        data = b"".join(self._sector(s) for s in self._chain(self.first_minifat)) if self.first_minifat != ENDOFCHAIN else b""
        self.minifat = list(struct.unpack("<%dI" % (len(data)//4), data)) if data else []

    def _read_dir(self):
        raw = b"".join(self._sector(s) for s in self._chain(self.first_dir))
        self.dir_raw = raw
        self.entries = []
        for i in range(0, len(raw), 128):
            e = raw[i:i+128]
            if len(e) < 128:
                break
            self.entries.append(bytearray(e))
        # root entry (type 5) holds the mini-stream in the FAT
        self.root_idx = next(i for i,e in enumerate(self.entries) if e[0x42] == 5)
        r = self.entries[self.root_idx]
        r_start = struct.unpack_from("<I", r, 0x74)[0]
        r_size  = struct.unpack_from("<Q", r, 0x78)[0]
        self.ministream = self._read_stream_fat(r_start, r_size) if r_size else b""

    def _read_mini(self, start, size):
        out = bytearray()
        s = start
        while s != ENDOFCHAIN and s != FREESECT:
            out += self.ministream[s*MINI:(s+1)*MINI]
            s = self.minifat[s]
        return bytes(out[:size])

    def entry_name(self, e):
        nlen = struct.unpack_from("<H", e, 0x40)[0]
        if nlen <= 0: return ""
        return e[:nlen-2].decode("utf-16-le")

    def entry_data(self, i):
        e = self.entries[i]
        typ = e[0x42]
        if typ != 2:            # only streams carry data here
            return None
        start = struct.unpack_from("<I", e, 0x74)[0]
        size  = struct.unpack_from("<Q", e, 0x78)[0]
        if size == 0:
            return b""
        if size < self.mini_cutoff:
            return self._read_mini(start, size)
        return self._read_stream_fat(start, size)


# ---------------------------------------------------------------- writer ----
def build_cfb(template: "Cfb", new_data: dict) -> bytes:
    """new_data: {entry_index: bytes} overrides for stream entries.
    Everything else is taken from template.entries / entry_data."""
    entries = [bytearray(e) for e in template.entries]
    n = len(entries)

    # collect per-entry data (streams only)
    data = {}
    for i, e in enumerate(entries):
        if e[0x42] == 2:
            data[i] = new_data[i] if i in new_data else template.entry_data(i)

    sectors = []           # list of 512-byte bytearrays (sector index = position)
    fat = []               # parallel FAT values

    def add_chain(blob):
        """append blob as new 512-byte sectors, link FAT chain, return first idx"""
        if not blob:
            return ENDOFCHAIN
        first = len(sectors)
        nsec = (len(blob) + SECT - 1) // SECT
        blob = blob + b"\x00" * (nsec*SECT - len(blob))
        for k in range(nsec):
            sectors.append(bytearray(blob[k*SECT:(k+1)*SECT]))
            fat.append(0)   # placeholder
        for k in range(nsec):
            fat[first+k] = (first+k+1) if k < nsec-1 else ENDOFCHAIN
        return first

    # 1) mini-stream + mini-FAT (streams with 0 < size < cutoff)
    ministream = bytearray()
    minifat = []
    for i, e in enumerate(entries):
        if e[0x42] == 2 and 0 < len(data[i]) < MINI_CUTOFF:
            d = data[i]
            nmini = (len(d) + MINI - 1) // MINI
            first_mini = len(minifat)
            d = d + b"\x00" * (nmini*MINI - len(d))
            ministream += d
            for k in range(nmini):
                minifat.append((first_mini+k+1) if k < nmini-1 else ENDOFCHAIN)
            struct.pack_into("<I", e, 0x74, first_mini)
            struct.pack_into("<Q", e, 0x78, len(data[i]))
        elif e[0x42] == 2 and len(data[i]) == 0:
            struct.pack_into("<I", e, 0x74, ENDOFCHAIN)
            struct.pack_into("<Q", e, 0x78, 0)

    # 2) big streams (size >= cutoff) -> FAT chains
    for i, e in enumerate(entries):
        if e[0x42] == 2 and len(data[i]) >= MINI_CUTOFF:
            start = add_chain(data[i])
            struct.pack_into("<I", e, 0x74, start)
            struct.pack_into("<Q", e, 0x78, len(data[i]))

    # 3) mini-FAT sectors, then the mini-stream container (owned by root)
    minifat_blob = b"".join(struct.pack("<I", v) for v in minifat)
    # pad minifat sector(s) with FREESECT
    if minifat_blob:
        pad = (-len(minifat_blob)) % SECT
        minifat_blob += b"\xFF" * pad
    first_minifat = add_chain(minifat_blob) if minifat_blob else ENDOFCHAIN
    n_minifat = (len(minifat_blob)//SECT) if minifat_blob else 0

    root_start = add_chain(bytes(ministream)) if ministream else ENDOFCHAIN
    ri = template.root_idx
    struct.pack_into("<I", entries[ri], 0x74, root_start if ministream else ENDOFCHAIN)
    struct.pack_into("<Q", entries[ri], 0x78, len(ministream))

    # 4) directory sectors (entries padded to whole sector with free entries)
    dir_blob = b"".join(bytes(e) for e in entries)
    if len(dir_blob) % SECT:
        free = bytearray(128); struct.pack_into("<I", free, 0x74, ENDOFCHAIN)
        # unused entry: name len 0, type 0; startSect free
        need = (SECT - (len(dir_blob) % SECT))
        filler = bytearray()
        while len(filler) < need:
            fe = bytearray(128)
            struct.pack_into("<I", fe, 0x44, FREESECT)  # left
            struct.pack_into("<I", fe, 0x48, FREESECT)  # right
            struct.pack_into("<I", fe, 0x4C, FREESECT)  # child
            filler += fe
        dir_blob += bytes(filler[:need])
    first_dir = add_chain(dir_blob)
    # NOTE: v3 header stores n_dir = 0
    n_dir_sectors = len(dir_blob)//SECT

    # 5) FAT + DIFAT sizing (iterate to fixpoint)
    D = len(sectors)
    F = 1
    while True:
        G = 0 if F <= 109 else (F - 109 + 126)//127
        total = D + F + G
        need_F = (total + 127)//128
        if need_F <= F:
            break
        F = need_F
    G = 0 if F <= 109 else (F - 109 + 126)//127

    # reserve indices for FAT sectors and DIFAT sectors
    fat_sect_idx  = [D + k for k in range(F)]
    difat_sect_idx = [D + F + k for k in range(G)]
    # extend sectors/fat placeholders for FAT + DIFAT sectors
    for _ in range(F + G):
        sectors.append(bytearray(SECT)); fat.append(FREESECT)
    for s in fat_sect_idx:  fat[s] = FATSECT
    for s in difat_sect_idx: fat[s] = DIFSECT

    # build DIFAT array
    header_difat = [FREESECT]*109
    for k in range(min(109, F)):
        header_difat[k] = fat_sect_idx[k]
    # remaining FAT sector pointers go into DIFAT sectors
    rest = fat_sect_idx[109:]
    for gi, ds in enumerate(difat_sect_idx):
        block = rest[gi*127:(gi+1)*127]
        vals = list(block) + [FREESECT]*(127-len(block))
        nxt = difat_sect_idx[gi+1] if gi+1 < len(difat_sect_idx) else ENDOFCHAIN
        vals.append(nxt)
        sectors[ds] = bytearray(struct.pack("<128I", *vals))

    # pad FAT to whole sectors with FREESECT and write into FAT sectors
    fat_full = fat + [FREESECT]*(F*128 - len(fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat_full)
    for k, s in enumerate(fat_sect_idx):
        sectors[s] = bytearray(fat_bytes[k*SECT:(k+1)*SECT])

    # 6) header
    hdr = bytearray(512)
    hdr[0:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<HHHHH", hdr, 0x18, template.minor, template.major, 0xFFFE, 9, 6)
    struct.pack_into("<I", hdr, 0x28, 0)                 # n dir sectors (0 for v3)
    struct.pack_into("<I", hdr, 0x2C, F)                 # n FAT sectors
    struct.pack_into("<I", hdr, 0x30, first_dir)
    struct.pack_into("<I", hdr, 0x34, 0)                 # transaction
    struct.pack_into("<I", hdr, 0x38, MINI_CUTOFF)
    struct.pack_into("<I", hdr, 0x3C, first_minifat)
    struct.pack_into("<I", hdr, 0x40, n_minifat)
    struct.pack_into("<I", hdr, 0x44, difat_sect_idx[0] if difat_sect_idx else ENDOFCHAIN)
    struct.pack_into("<I", hdr, 0x48, G)
    struct.pack_into("<109I", hdr, 0x4C, *header_difat)

    return bytes(hdr) + b"".join(bytes(s) for s in sectors)
