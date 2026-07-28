from playwright.sync_api import sync_playwright
import pathlib, sys

src = pathlib.Path("Closing-Manager-Briefing-DSSmith.html").resolve()
out = pathlib.Path("Closing-Manager-Briefing-DSSmith.pdf").resolve()

# The published HTML has no <html>/<head>; wrap it so fonts + charset resolve.
raw = src.read_text(encoding="utf-8")
doc = ('<!doctype html><html><head><meta charset="utf-8">'
       '<meta name="viewport" content="width=device-width,initial-scale=1"></head>'
       '<body>' + raw + '</body></html>')
tmp = src.with_name("_print.html"); tmp.write_text(doc, encoding="utf-8")

footer = """
<div style="width:100%;font-size:7pt;font-family:'Liberation Sans',Arial,sans-serif;
            color:#93999E;padding:0 15mm;display:flex;justify-content:space-between;
            border-top:0.5pt solid #DEE1E3;padding-top:3mm;margin:0 0 4mm 0;">
  <span style="color:#4D4D4D;">DS Smith &nbsp;·&nbsp; Record to Report &nbsp;·&nbsp; Closing Manager dependency review</span>
  <span>Page <span class="pageNumber"></span> of <span class="totalPages"></span></span>
</div>"""

with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/opt/pw-browsers/chromium")
    pg = b.new_page()
    pg.goto(tmp.as_uri(), wait_until="networkidle")
    pg.emulate_media(media="print")
    pg.pdf(path=str(out), format="A4", print_background=True,
           display_header_footer=True,
           header_template="<div></div>",
           footer_template=footer,
           margin={"top":"16mm","bottom":"20mm","left":"15mm","right":"15mm"},
           prefer_css_page_size=False)
    b.close()
tmp.unlink()
print("wrote", out, out.stat().st_size, "bytes")
