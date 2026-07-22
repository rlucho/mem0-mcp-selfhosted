# MS-OVBA compression via all-literal COMPRESSED chunks (flag=1) + offset finder.
import olefile
from oletools.olevba import decompress_stream

CHUNK_D = 2048   # decompressed bytes per chunk (multiple of 8; keeps compressedData < 4096)

def compress_ovba(data: bytes) -> bytes:
    out = bytearray([0x01])                     # CompressedContainer SignatureByte
    for i in range(0, len(data), CHUNK_D):
        chunk = data[i:i+CHUNK_D]
        cdata = bytearray()
        for j in range(0, len(chunk), 8):
            cdata.append(0x00)                  # FlagByte: 8 literal tokens
            cdata += chunk[j:j+8]
        size_field = (len(cdata) + 2) - 3       # CompressedChunkSize - 3
        header = 0xB000 | (size_field & 0x0FFF) # flag=1(0x8000) | sig=0b011(0x3000)
        out += header.to_bytes(2, "little"); out += cdata
    return bytes(out)

def _b(x): return x if isinstance(x, bytes) else x.encode("latin-1")

def find_source_offset(raw: bytes):
    for off in range(len(raw) - 2):
        if raw[off] != 0x01: continue
        hdr = raw[off+1] | (raw[off+2] << 8)
        if (hdr >> 12) & 0x07 != 0x03: continue
        try: dec = _b(decompress_stream(bytearray(raw[off:])))
        except Exception: continue
        if dec[:13] == b"Attribute VB_": return off, dec
    return None, None

if __name__ == "__main__":
    ole = olefile.OleFileIO("extracted/xl/vbaProject.bin")
    ok = True
    print("Module offset + compressor round-trip:")
    for m in ("GlobalModule","Closing","Printing"):
        raw = ole.openstream(f"VBA/{m}").read()
        off, src_b = find_source_offset(raw)
        recomp = compress_ovba(src_b)
        rt = _b(decompress_stream(bytearray(recomp)))
        match = rt == src_b
        txt = src_b.decode("cp1252")
        crlf_ok = txt.replace("\r\n","\n").replace("\n","\r\n").encode("cp1252") == src_b
        ok = ok and match and crlf_ok and off is not None
        print(f"  VBA/{m:12s} offset={off:6d} src={len(src_b):7d} container={len(recomp):7d} roundtrip={'OK' if match else 'FAIL'} crlf={'OK' if crlf_ok else 'FAIL'}")
    ole.close()
    print("ALL_GOOD" if ok else "PROBLEM")
