"""Build the two deliverables from src/app.html.

  python3 excel/make_xlsx.py --clean   # regenerate the Excel template (only when the template changes)
  python3 build.py                     # writes index.html and dist/artifact.html

- index.html          -> the standalone app served by GitHub Pages (full HTML document + PWA tags).
- dist/artifact.html  -> body-only version published as the claude.ai artifact (the viewer adds the skeleton).

Both embed excel/template_clean.xlsx as base64 in place of the /*TPL*/"" placeholder.
"""
import base64
import os

HERE = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(HERE, "src", "app.html"), encoding="utf-8").read()
tpl = base64.b64encode(open(os.path.join(HERE, "excel", "template_clean.xlsx"), "rb").read()).decode()
assert '/*TPL*/""' in src, "template placeholder missing from src/app.html"
body = src.replace('/*TPL*/""', '"' + tpl + '"')

os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
with open(os.path.join(HERE, "dist", "artifact.html"), "w", encoding="utf-8") as f:
    f.write(body)

HEAD = """<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#0E6A4E">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="دفتر التفتيش">
<link rel="manifest" href="manifest.webmanifest">
<link rel="apple-touch-icon" href="apple-touch-icon.png">
<link rel="icon" href="icon-192.png">
<style>:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>
</head>
<body>
"""
with open(os.path.join(HERE, "index.html"), "w", encoding="utf-8") as f:
    f.write(HEAD + body + "\n</body>\n</html>\n")
print("built index.html and dist/artifact.html")
