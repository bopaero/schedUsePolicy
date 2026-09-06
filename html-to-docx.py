#!/usr/bin/env python3
"""Regenerate the Scheduling and Use Policy docx from index.html.

Usage:
  python3 html-to-docx.py                     # index.html -> canonical iCloud docx
  python3 html-to-docx.py index.html out.docx # explicit paths

The site is the source of truth for this direction: it rebuilds the docx from
published HTML. Use it when the docx has drifted (or lost formatting in a Pages
round-trip) rather than hand-reconciling the two.

CAUTION: regenerating changes list markers and unwraps rule-box/callout
structure, so the next `update-policy.py` run will see those as content changes
and over-report changed paragraphs. Expect to audit the change bars afterwards
and strip false positives in a follow-up commit.

Recipe (order matters):
  1. pull hero title/intro + all <section class="guide-section"> blocks
  2. strip callout-icon divs FIRST, so the callout wrapper regex sees no nested divs
  3. unwrap remaining presentational divs (callout, rule-box)
  4. drop class/style attributes
  5. pandoc -f html -t docx   (NO table of contents - a Word TOC field
     expands into ~80 junk paragraphs under `pandoc -t plain` and destroys
     the change-bar diff in update-policy.py)
"""
import re, subprocess, sys, pathlib, tempfile

REPO = pathlib.Path(__file__).resolve().parent
CANONICAL_DOCX = pathlib.Path(
    "~/Library/Mobile Documents/com~apple~CloudDocs/Bop Aero/BopAeroMediaAssets/"
    "Scheduling and Use Policy/bop Aero\u00ae Scheduling and Use Policy.docx"
).expanduser()

if len(sys.argv) > 3:
    sys.exit(__doc__)
src = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "index.html"
out = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else CANONICAL_DOCX
if not src.is_file():
    sys.exit(f"ERROR: no such file: {src}")
html = src.read_text(encoding="utf-8")

# --- hero title + intro -------------------------------------------------
hero = re.search(r'<div class="hero">(.*?)<div class="hero-meta">', html, re.S).group(1)
h1 = re.search(r"<h1>(.*?)</h1>", hero, re.S).group(1)
h1 = re.sub(r"<br\s*/?>", " — ", h1)
h1 = re.sub(r"</?span[^>]*>", "", h1).strip()
intros = [m.strip() for m in re.findall(r"<p[^>]*>(.*?)</p>", hero, re.S)]
if not intros:
    sys.exit("ERROR: no intro paragraph found in the hero block")

# --- version stamp -----------------------------------------------------
# The docx is signed by owners/lessees, so it must state the version it is
# being signed against. Mirrors the site's print-only header (.print-header)
# so the printed page and the docx carry identical identification.
vm = re.search(r"\{\s*version:\s*'(v[\d\-\.]+)'[^}]*?archived:\s*null", html, re.S)
if not vm:
    sys.exit("ERROR: could not read CURRENT_VERSION from index.html "
             "(no VERSIONS entry with archived: null)")
version = vm.group(1)

parts = [
    "<p><strong>bop Aero Services LLC</strong></p>",
    f"<p>Aircraft Ownership Programs \u2014 Scheduling and Use Policy | {version}</p>",
    f"<h1>{h1}</h1>",
] + [f"<p>{t}</p>" for t in intros]
parts += re.findall(r'<section class="guide-section"[^>]*>(.*?)</section>', html, re.S)
body = "\n".join(parts)

# --- 2. callout icons out first ----------------------------------------
body = re.sub(r'<div class="callout-icon">.*?</div>', "", body, flags=re.S)

# --- 3. unwrap presentational divs (now free of nested divs) -----------
for _ in range(4):
    new = re.sub(r'<div class="(?:callout[^"]*|rule-box)"\s*>(.*?)</div>', r"\1", body, flags=re.S)
    if new == body:
        break
    body = new

if "<div" in body:
    sys.exit("ERROR: unhandled <div> left in body:\n"
             + "\n".join(sorted(set(re.findall(r"<div[^>]*>", body)))))

# --- 4. drop presentation attributes ----------------------------------
body = re.sub(r'\s(?:class|style|id)="[^"]*"', "", body)

doc = "<html><body>\n" + body + "\n</body></html>"
tmp = pathlib.Path(tempfile.gettempdir()) / "policy-gen.html"
tmp.write_text(doc, encoding="utf-8")

subprocess.run(["pandoc", "-f", "html", "-t", "docx", str(tmp), "-o", str(out)], check=True)
print(f"wrote {out}  ({version})")
