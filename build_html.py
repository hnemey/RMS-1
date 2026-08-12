#!/usr/bin/env python3
"""Inline the JS libraries + core logic into template.html -> ppr_to_pptx.html"""
import pathlib

base = pathlib.Path("/home/user/RMS-1")
tpl = (base / "template.html").read_text()
xlsx = (base / "node_modules/xlsx/dist/xlsx.full.min.js").read_text()
pptx = (base / "node_modules/pptxgenjs/dist/pptxgen.bundle.js").read_text()
core = (base / "core.js").read_text()

def wrap(js):
    # guard against an accidental </script> inside library source
    return "<script>\n" + js.replace("</script>", "<\\/script>") + "\n</script>"

html = (tpl
        .replace("<!--XLSX_LIB-->", wrap(xlsx))
        .replace("<!--PPTX_LIB-->", wrap(pptx))
        .replace("<!--CORE_LIB-->", wrap(core)))

out = base / "ppr_to_pptx.html"
out.write_text(html)
print("wrote", out, f"({len(html)/1024:.0f} KB)")
