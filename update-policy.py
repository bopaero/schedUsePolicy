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

import sys, os, re, shutil, subprocess, argparse
from difflib import SequenceMatcher
from datetime import date

REPO        = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML  = os.path.join(REPO, 'index.html')
VERSIONS_DIR = os.path.join(REPO, 'versions')
REFERENCE   = os.path.join(REPO, '.reference.txt')


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
        if len(line) > 15:   # skip headings, labels, blank lines
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

    return '\n'.join(result)


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
    new_entry = f"  {{ version: '{new_version}', label: '{label}', archived: null }},\n"
    html = re.sub(r'(\];\s*\nconst CURRENT_VERSION)', new_entry + r'\1', html)
    return html


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Publish a policy update to schedUsePolicy.bopaero.com')
    parser.add_argument('docx', help='Path to the updated .docx exported from Pages')
    parser.add_argument('--label', default='', help='Short description of what changed (shown in version history)')
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

    with open(INDEX_HTML, 'w') as f:
        f.write(html)

    # ── Save new reference
    save_reference(new_paras)

    # ── Git
    print('\nPushing to GitHub...')
    run(['git', 'pull'])
    run(['git', 'add', 'index.html',
         f'versions/{current_version}.html',
         '.reference.txt'])
    run(['git', 'commit', '-m',
         f'Policy update {new_version}: {label}\n\n'
         'Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>'])
    run(['git', 'push'])

    print(f'\nDone! {new_version} is live at https://schedusepolicy.bopaero.com (updates in ~30s)')


if __name__ == '__main__':
    main()
