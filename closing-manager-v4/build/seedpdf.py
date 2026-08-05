"""Reference implementation of the minimal PDF the VBA CM_SeedPdf writes.
Kept here so the byte-exactness of the VBA version can be proved."""

def seed_pdf(line1, line2):
    NL = "\n"
    body = ("BT" + NL +
            "/F1 16 Tf" + NL +
            "60 760 Td" + NL +
            "(" + line1 + ") Tj" + NL +
            "0 -28 Td" + NL +
            "(" + line2 + ") Tj" + NL +
            "ET" + NL)

    objs = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
        "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        "<< /Length " + str(len(body)) + " >>" + NL + "stream" + NL + body + "endstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    out = "%PDF-1.4" + NL
    offs = []
    for i, o in enumerate(objs):
        offs.append(len(out))
        out += str(i + 1) + " 0 obj" + NL + o + NL + "endobj" + NL

    xref_at = len(out)
    out += "xref" + NL + "0 " + str(len(objs) + 1) + NL
    out += "0000000000 65535 f " + NL
    for off in offs:
        out += "%010d 00000 n " % off + NL
    out += ("trailer" + NL + "<< /Size " + str(len(objs) + 1) + " /Root 1 0 R >>" + NL +
            "startxref" + NL + str(xref_at) + NL + "%%EOF" + NL)
    return out.encode("ascii")

if __name__ == "__main__":
    import io, sys
    from pypdf import PdfReader, PdfWriter

    a = seed_pdf("Closing Manager - preflight test", "PDF merge test page 1")
    b = seed_pdf("Closing Manager - preflight test", "PDF merge test page 2")
    open("seed1.pdf", "wb").write(a)
    open("seed2.pdf", "wb").write(b)
    print("seed sizes:", len(a), len(b))

    for n, data in (("seed1", a), ("seed2", b)):
        r = PdfReader(io.BytesIO(data))
        txt = (r.pages[0].extract_text() or "").replace("\n", " / ")
        print("  %s: pages=%d  mediabox=%s  text=%r" % (n, len(r.pages), r.pages[0].mediabox, txt))

    w = PdfWriter()
    for data in (a, b):
        for p in PdfReader(io.BytesIO(data)).pages:
            w.add_page(p)
    buf = io.BytesIO(); w.write(buf)
    m = PdfReader(io.BytesIO(buf.getvalue()))
    print("merged: pages=%d size=%d" % (len(m.pages), len(buf.getvalue())))
    open("seed_merged.pdf", "wb").write(buf.getvalue())
    assert len(m.pages) == 2
    print("VALID + MERGEABLE")
