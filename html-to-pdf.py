#!/usr/bin/env python3
"""Render the Scheduling and Use Policy to PDF from index.html.

Usage:
  python3 html-to-pdf.py                     # index.html -> canonical iCloud PDF
  python3 html-to-pdf.py index.html out.pdf  # explicit paths

The PDF is the artifact owners and lessees actually sign (sent via DocuSign),
so it is generated straight from the published HTML rather than by way of the
docx. That removes Word from the chain entirely — there is no Word on this
machine to inspect the intermediate with, and Word, Pages and DocuSign are
three different renderers. Going HTML -> PDF means the file that ships is the
one the print stylesheet produced.

Uses headless Chrome, so the page's JS runs and the version spans
(#printVersion, #ackVersion) are populated exactly as on the site.
"""
import re, subprocess, sys, pathlib

CHROME = pathlib.Path(
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
)
REPO = pathlib.Path(__file__).resolve().parent
CANONICAL_PDF = pathlib.Path(
    "~/Library/Mobile Documents/com~apple~CloudDocs/Bop Aero/BopAeroMediaAssets/"
    "Scheduling and Use Policy/bop Aero® Scheduling and Use Policy.pdf"
).expanduser()

if len(sys.argv) > 3:
    sys.exit(__doc__)
src = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "index.html"
out = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else CANONICAL_PDF
src = src.resolve()

if not src.is_file():
    sys.exit(f"ERROR: no such file: {src}")
if not CHROME.exists():
    sys.exit(f"ERROR: Chrome not found at {CHROME}\n"
             "Install Google Chrome, or generate the PDF by opening the site "
             "and using the Print / Save PDF button.")

version = None
m = re.search(r"\{\s*version:\s*'(v[\d\-\.]+)'[^}]*?archived:\s*null",
              src.read_text(encoding="utf-8"), re.S)
if m:
    version = m.group(1)
else:
    sys.exit("ERROR: could not read CURRENT_VERSION from index.html "
             "(no VERSIONS entry with archived: null)")

out.parent.mkdir(parents=True, exist_ok=True)
result = subprocess.run(
    [str(CHROME), "--headless", "--disable-gpu", "--no-sandbox",
     # let the page settle so JS has filled the version spans
     "--virtual-time-budget=8000",
     "--run-all-compositor-stages-before-draw",
     "--no-pdf-header-footer",          # the page prints its own header
     f"--print-to-pdf={out}",
     src.as_uri()],
    capture_output=True, text=True,
)
if result.returncode != 0 or not out.is_file():
    sys.exit(f"ERROR: Chrome failed to render the PDF\n{result.stderr.strip()[:800]}")

data = out.read_bytes()
pages = data.count(b"/Type /Page") - data.count(b"/Type /Pages")
if pages < 2:
    sys.exit(f"ERROR: {out} came out with {pages} page(s) — the render is wrong. "
             "Check that the print stylesheet still applies.")

print(f"wrote {out}  ({version}, {pages} pages, {len(data):,} bytes)")
