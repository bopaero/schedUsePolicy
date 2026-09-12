#!/usr/bin/env python3
"""
update-policy.py — Automated policy update for schedUsePolicy.bopaero.com

Usage:
  python3 update-policy.py /path/to/updated.docx
  python3 update-policy.py /path/to/updated.docx --label "Updated cancellation fees"

What it does:
  1. Converts the new .docx to plain text via pandoc
  2. Diffs against the last-published version (stored in .reference.txt)
  3. Marks changed paragraphs with a visible change bar in the HTML
  4. Archives the current version to versions/
  5. Bumps the version number and updates the version history dropdown
  6. Commits and pushes to GitHub (live in ~30 seconds)
"""

import sys, os, re, shutil, subprocess, argparse, tempfile
from difflib import SequenceMatcher
from datetime import date

REPO        = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML  = os.path.join(REPO, 'index.html')
VERSIONS_DIR = os.path.join(REPO, 'versions')
REFERENCE   = os.path.join(REPO, '.reference.txt')
HTML_TO_DOCX = os.path.join(REPO, 'html-to-docx.py')
HTML_TO_PDF  = os.path.join(REPO, 'html-to-pdf.py')

# Matches the version stamp line the docx carries, e.g.
# "Aircraft Ownership Programs — Scheduling and Use Policy | v2026-09-06.1"
VERSION_STAMP_RE = re.compile(r'v\d{4}-\d{2}-\d{2}\.\d+')


# ── Helpers ──────────────────────────────────────────────────────────────────

def run(cmd, **kwargs):
    subprocess.run(cmd, check=True, cwd=REPO, **kwargs)

def pandoc_to_paragraphs(docx_path):
    """Convert a .docx to a list of normalized plain-text paragraphs."""
    result = subprocess.run(
        ['pandoc', docx_path, '-t', 'plain', '--wrap=none'],
        capture_output=True, text=True, check=True
    )
    paras = []
    for line in result.stdout.splitlines():
        line = re.sub(r'\s+', ' ', line).strip()
        if len(line) <= 15:      # skip headings, labels, blank lines
            continue
        # The docx carries a version stamp (see html-to-docx.py) so that a
        # signed copy states which version was signed. It changes every
        # release by definition — diffing it would flag a bogus change and
        # send the fuzzy matcher hunting for a paragraph that isn't policy.
        if VERSION_STAMP_RE.search(line):
            continue
        paras.append(line)
    return paras

def load_reference():
    if not os.path.exists(REFERENCE):
        return []
    with open(REFERENCE) as f:
        return [l.rstrip('\n') for l in f if l.strip()]

def save_reference(paras):
    with open(REFERENCE, 'w') as f:
        f.write('\n'.join(paras) + '\n')

def find_changed_texts(old_paras, new_paras):
    """Return set of paragraph texts that are new or modified in new_paras."""
    changed = set()
    sm = SequenceMatcher(None, old_paras, new_paras, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ('replace', 'insert'):
            for p in new_paras[j1:j2]:
                changed.add(p)
    return changed

def strip_tags(text):
    return re.sub(r'<[^>]+>', '', text)

def norm(text):
    return re.sub(r'\s+', ' ', strip_tags(text)).strip()

def matches_changed(element_text, changed_texts):
    n = norm(element_text)
    if len(n) < 15:
        return False
    for c in changed_texts:
        ratio = SequenceMatcher(None, n.lower(), c.lower()).ratio()
        if ratio > 0.80:
            return True
    return False


# ── HTML patching ─────────────────────────────────────────────────────────────

def add_class_to_tag(tag_str, cls):
    """Add a CSS class to an opening HTML tag string."""
    if 'class="' in tag_str:
        return re.sub(r'class="([^"]*)"', lambda m: f'class="{m.group(1).strip()} {cls}"', tag_str)
    return re.sub(r'^(<\w+)', r'\1' + f' class="{cls}"', tag_str)

def apply_change_bars(html, changed_texts):
    """Add class="changed" to block elements whose text matches a changed paragraph.
    Works line-by-line so multiline elements don't confuse the regex."""

    # Strip previous change bars — only from class="" attributes, never from CSS
    html = re.sub(r'(class="[^"]*)\bchanged\b\s*([^"]*")', r'\1\2', html)
    html = re.sub(r' class="\s*"', '', html)

    if not changed_texts:
        return html

    lines = html.split('\n')
    result = []
    i = 0
    BLOCK_TAGS = ('p', 'li', 'td', 'h3', 'h4')
    open_tag_re = re.compile(r'^(\s*)<(' + '|'.join(BLOCK_TAGS) + r')(\b[^>]*)>(.*)', re.DOTALL)

    while i < len(lines):
        line = lines[i]
        m = open_tag_re.match(line)
        if m:
            indent, tag, attrs, rest = m.group(1), m.group(2), m.group(3), m.group(4)
            # Gather full element content (may span multiple lines)
            content_lines = [rest]
            close = f'</{tag}>'
            j = i
            while close not in '\n'.join(content_lines) and j < len(lines) - 1:
                j += 1
                content_lines.append(lines[j])
            full_content = '\n'.join(content_lines)
            if matches_changed(full_content, changed_texts):
                tag_open = f'<{tag}{attrs}>'
                tag_open = add_class_to_tag(tag_open, 'changed')
                result.append(f'{indent}{tag_open}{full_content}')
                i = j + 1
                continue
            else:
                result.append(line)
                if j > i:
                    result.extend(lines[i+1:j+1])
                    i = j + 1
                    continue
        else:
            result.append(line)
        i += 1

    html = '\n'.join(result)

    # Second pass: <p> elements that do NOT start a line.
    #
    # The loop above anchors on ^\s*<tag>, so a paragraph wrapped in a one-line
    # callout never matches and never gets a bar:
    #
    #   <div class="callout note"><div class="callout-icon">i</div><p>text</p></div>
    #
    # The line begins with <div>, so the <p> is invisible to it. This was hit on
    # three consecutive releases (v2026-09-06.4, v2026-09-12.1, v2026-09-12.2) and
    # fixed by hand each time. It is fully deterministic, so handle it here.
    #
    # Deliberately <p> only: <td>/<th> are also inline within <tr>, but no table
    # cell in this document has ever carried a change bar (including the fuel and
    # runway tables) — the paragraph introducing a table is barred instead.
    def _bar_inline_p(m):
        attrs, content = m.group(1), m.group(2)
        if 'changed' in (attrs or ''):
            return m.group(0)
        if not matches_changed(content, changed_texts):
            return m.group(0)
        return add_class_to_tag(f'<p{attrs}>', 'changed') + content + '</p>'

    html = re.sub(r'(?<!^)(?<!\n)<p(\b[^>]*)?>(.*?)</p>',
                  _bar_inline_p, html, flags=re.MULTILINE)

    return html


# ── Version management ────────────────────────────────────────────────────────

def get_current_version(html):
    m = re.search(r"\{\s*version:\s*'(v[\d\-\.]+)'[^}]*?archived:\s*null", html, re.DOTALL)
    return m.group(1) if m else None

def next_version(current):
    today = date.today().strftime('%Y-%m-%d')
    m = re.match(r'v(\d{4}-\d{2}-\d{2})\.(\d+)', current)
    if m and m.group(1) == today:
        return f'v{today}.{int(m.group(2)) + 1}'
    return f'v{today}.1'

def update_versions_array(html, old_version, new_version, label):
    # Archive the current (old) entry
    html = re.sub(
        r"(version:\s*'" + re.escape(old_version) + r"'.*?archived:\s*)null",
        r"\1'versions/" + old_version + ".html'",
        html,
        flags=re.DOTALL
    )
    # Append new entry before closing ];
    # Escape for a single-quoted JS string literal — an apostrophe in the label
    # (e.g. "aircraft's operating account") otherwise breaks the whole <script>
    # block, blanking the version badge and history dropdown.
    js_label = label.replace('\\', '\\\\').replace("'", "\\'")
    new_entry = f"  {{ version: '{new_version}', label: '{js_label}', archived: null }},\n"
    html = re.sub(r'(\];\s*\nconst CURRENT_VERSION)', new_entry + r'\1', html)
    return html


def check_script_blocks(html):
    """Syntax-check every inline <script> block with `node --check`.

    update-policy.py writes a free-text label straight into a single-quoted JS
    string literal. A stray character there kills the whole <script> block, and
    because VERSIONS and the main UI block are coupled, that silently blanks the
    version badge, the history dropdown, the footer/print version and the nav
    scrollspy — on a page the script has already committed and pushed.

    Returns (ok, message). Skips the check (ok=True) if node is unavailable, so
    a missing toolchain never blocks a publish.
    """
    if not shutil.which('node'):
        return True, 'node not found — skipping JS syntax check'

    blocks = re.findall(r'<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>', html, re.DOTALL)
    if not blocks:
        return True, 'no inline <script> blocks found'

    for i, block in enumerate(blocks, 1):
        with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False,
                                         encoding='utf-8') as f:
            f.write(block)
            tmp = f.name
        try:
            result = subprocess.run(['node', '--check', tmp],
                                    capture_output=True, text=True)
        finally:
            os.unlink(tmp)
        if result.returncode != 0:
            detail = (result.stderr or '').strip()
            # node reports the resolved path (/private/var/... on macOS)
            for path in (os.path.realpath(tmp), tmp):
                detail = detail.replace(path, f'<script block {i}>')
            # keep the pointed-at source line and the error, drop node's stack
            lines = []
            for line in detail.splitlines():
                if line.startswith('    at ') or line.startswith('Node.js v'):
                    break
                lines.append(line)
            detail = '\n'.join(lines).strip()
            return False, f'script block {i} of {len(blocks)} has a syntax error:\n{detail}'

    return True, f'{len(blocks)} inline script block(s) OK'



def verify_live(new_version, timeout_s=300, interval_s=15):
    """Report what the SITE is actually serving — not merely what we pushed.

    The old message claimed the version was 'live' the instant git push returned.
    That is false: GitHub Pages still has to build and deploy, and on 2026-09-12 a
    build sat queued for over ten minutes while the site served the previous
    version. Reporting a push as a publish sends you off hard-refreshing a page
    that was never going to change.
    """
    import urllib.request, urllib.error, time

    url = 'https://schedusepolicy.bopaero.com/'
    try:
        local = open(INDEX_HTML, 'rb').read()
    except Exception:
        local = None

    print(f'Waiting for GitHub Pages to deploy (up to {timeout_s // 60} min)...')
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            req = urllib.request.Request(
                f'{url}?cb={int(time.time() * 1000)}',
                headers={'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
            with urllib.request.urlopen(req, timeout=30) as r:
                body = r.read()
            if new_version.encode() in body and (local is None or len(body) == len(local)):
                print(f'LIVE       {new_version} at {url}')
                return True
        except Exception:
            pass
        time.sleep(interval_s)

    print(f'NOT LIVE   the site is still serving an older version after '
          f'{timeout_s // 60} min.')
    print('  Your commit is safe on GitHub; only the Pages deploy is behind.')
    print('  Check the build:  https://github.com/bopaero/schedUsePolicy/actions')
    print('  A run stuck in "queued" clears with Cancel workflow, then Re-run.')
    print('  Also confirm Settings -> Pages still reads '
          '"Deploy from a branch: main / (root)".')
    return False


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Publish a policy update to schedUsePolicy.bopaero.com')
    parser.add_argument('docx', help='Path to the updated .docx exported from Pages')
    parser.add_argument('--label', default='', help='Short description of what changed (shown in version history)')
    parser.add_argument('--no-push', action='store_true',
                        help='Commit locally but do NOT push. Use this so the change-bar audit '
                             'can be folded into the SAME commit: a second push landing seconds '
                             'after the first cancels the in-flight GitHub Pages build and the '
                             'replacement can sit queued indefinitely (seen 2026-09-12).')
    args = parser.parse_args()

    docx_path = os.path.expanduser(args.docx)
    if not os.path.exists(docx_path):
        print(f'Error: file not found: {docx_path}')
        sys.exit(1)

    # ── Parse new document
    print('Parsing updated document...')
    new_paras = pandoc_to_paragraphs(docx_path)
    old_paras = load_reference()

    # ── First-time setup: just save reference, nothing to diff
    if not old_paras:
        print('No previous reference found — saving reference snapshot.')
        save_reference(new_paras)
        print('Done. Next time you run this script with an updated docx, changes will be detected.')
        return

    # ── Diff
    print(f'Comparing {len(old_paras)} → {len(new_paras)} paragraphs...')
    changed_texts = find_changed_texts(old_paras, new_paras)

    if not changed_texts:
        print('No content changes detected. Nothing to publish.')
        print('(If you expected changes, make sure you saved the Pages file and exported a fresh .docx.)')
        return

    print(f'\nDetected {len(changed_texts)} changed paragraph(s):')
    for t in sorted(changed_texts):
        preview = t[:90] + '...' if len(t) > 90 else t
        print(f'  • {preview}')

    # ── Label
    label = args.label.strip()
    if not label:
        label = input('\nShort description for version history (e.g. "Updated cancellation fees"): ').strip()
        if not label:
            label = 'Policy update'

    # ── Load HTML
    with open(INDEX_HTML) as f:
        html = f.read()

    current_version = get_current_version(html)
    if not current_version:
        print('Error: could not detect current version in index.html')
        sys.exit(1)

    new_version = next_version(current_version)
    print(f'\nVersion: {current_version} → {new_version}')

    # ── Archive current version
    os.makedirs(VERSIONS_DIR, exist_ok=True)
    archive_path = os.path.join(VERSIONS_DIR, f'{current_version}.html')
    shutil.copy(INDEX_HTML, archive_path)
    print(f'Archived: versions/{current_version}.html')

    # ── Patch HTML
    html = apply_change_bars(html, changed_texts)
    html = update_versions_array(html, current_version, new_version, label)

    # ── Verify the patched JS still parses BEFORE anything touches disk or git.
    # A bad label would otherwise be committed and pushed onto a broken page.
    print('\nChecking JavaScript syntax...')
    ok, message = check_script_blocks(html)
    print(f'  {message}')
    if not ok:
        print('\nABORTED — index.html was NOT modified and nothing was committed.')
        print('The version history label is the usual culprit; re-run with a')
        print('simpler --label, or copy the text from the docx rather than retyping it.')
        os.remove(archive_path)
        sys.exit(1)

    with open(INDEX_HTML, 'w') as f:
        f.write(html)

    # ── Rebuild the docx from the just-published HTML.
    # The version is only assigned here, so the docx that was fed in still
    # carries the previous one. Owners and lessees sign the docx, so it must
    # state the version it is being signed against — regenerate it now, then
    # take the reference snapshot from that file so HTML, docx and reference
    # are one lineage with one version.
    print('\nRegenerating the source docx from the published HTML...')
    subprocess.run([sys.executable, HTML_TO_DOCX, INDEX_HTML, docx_path], check=True)
    new_paras = pandoc_to_paragraphs(docx_path)

    # ── Render the signing PDF. This is the artifact owners and lessees
    # actually sign, generated straight from the published HTML so Word never
    # enters the chain. Non-fatal: a missing Chrome must not strand a release
    # that is otherwise complete — the site and docx are already correct.
    pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    print('Rendering the signing PDF...')
    pdf = subprocess.run([sys.executable, HTML_TO_PDF, INDEX_HTML, pdf_path],
                         capture_output=True, text=True)
    if pdf.returncode == 0:
        print(f'  {pdf.stdout.strip()}')
    else:
        print('  WARNING: PDF generation failed — site and docx are fine, but')
        print('  the signing PDF was not refreshed. Re-run: python3 html-to-pdf.py')
        print(f'  {pdf.stderr.strip()[:300]}')

    # ── Archive the signed artifacts for this version. versions/ held only
    # HTML; the canonical docx and PDF are overwritten every release, so a
    # signed version's exact files would otherwise be unrecoverable.
    for path in (docx_path, pdf_path):
        if os.path.exists(path):
            ext = os.path.splitext(path)[1]
            dest = os.path.join(VERSIONS_DIR, f'{new_version}{ext}')
            shutil.copy(path, dest)
            print(f'Archived: versions/{new_version}{ext}')

    # ── Save new reference
    save_reference(new_paras)

    # ── Git
    print('\nCommitting...')
    run(['git', 'pull'])
    run(['git', 'add', 'index.html',
         f'versions/{current_version}.html',
         '.reference.txt'])
    run(['git', 'add', '--', VERSIONS_DIR])
    run(['git', 'commit', '-m',
         f'Policy update {new_version}: {label}\n\n'
         'Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>'])

    if args.no_push:
        print(f'\nCOMMITTED  {new_version} (local only — not pushed)')
        print('  Audit the change bars now, amend or add a follow-up commit, then:')
        print('    git push')
        print('  Pushing once keeps this release to a single GitHub Pages build.')
        return

    run(['git', 'push'])
    print(f'\nPUSHED     {new_version}')
    verify_live(new_version)


if __name__ == '__main__':
    main()
