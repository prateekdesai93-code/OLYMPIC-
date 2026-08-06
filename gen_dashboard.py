#!/usr/bin/env python3
"""
gen_dashboard.py — publish the Clocking Dashboard to GitHub Pages
=====================================================================

Takes the accumulated raw-punch master store, embeds it into the
dashboard's index.html template as a compact inline JSON payload, and
(optionally) commits + pushes the result to a GitHub Pages repo — so the
published URL opens pre-loaded with the latest data, exactly like the
original site, while the same index.html can still be handed out
stand-alone (open it directly and use "Import Punch Data" — the embedded
payload is just an optional head start).

Usage:
    python gen_dashboard.py --master master_raw_data.xlsx \
        --template index.html \
        --repo /path/to/github-pages-repo \
        [--out-name index.html] [--branch main] \
        [--message "Update clocking dashboard"] [--no-push]

  --master    Accumulated raw-punch store (as produced by build_report.py
              / process_inbox.py).
  --template  The dashboard source file (the one Claude built) — this is
              never modified; a copy with data embedded is written out.
  --repo      Path to the local clone of the GitHub Pages repo. The
              generated file is written inside it and, unless --no-push
              is given, committed and pushed from there.
  --out-name  Filename to write inside --repo. Defaults to index.html.
  --branch    Branch to push. Defaults to main (adjust if your Pages
              site is served from gh-pages or docs/ on another branch).
  --no-push   Write the file and stage it locally, but skip git commit/push
              (useful for reviewing the diff first).
"""

import argparse
import json
import re
import subprocess
import sys
from datetime import date
from pathlib import Path

import build_report

EMBED_MARKER_RE = re.compile(
    r'(<script type="application/json" id="embeddedData">)(.*?)(</script>)',
    re.S,
)


def build_embedded_payload(punches):
    """Compact representation matching index.html's loadEmbeddedData() format:
       {id, f: firstName, l: lastName, dept, d: 'YYYY/MM/DD', m: minutes-of-day}"""
    compact = [{
        'id': p['id'],
        'f': p['first_name'],
        'l': p['last_name'],
        'dept': p['department'],
        'd': build_report.date_key(p['date']),
        'm': p['minutes'],
    } for p in punches]
    return {'punches': compact, 'generated': build_report.date_short(date.today())}


def embed_into_template(template_html: str, payload: dict) -> str:
    payload_json = json.dumps(payload, separators=(',', ':'), ensure_ascii=False)
    # Escape "</script" sequences so an embedded value can never prematurely close the tag
    payload_json = payload_json.replace('</script', '<\\/script')

    new_html, count = EMBED_MARKER_RE.subn(lambda m: m.group(1) + payload_json + m.group(3), template_html)
    if count == 0:
        raise RuntimeError(
            'Could not find the embeddedData placeholder in the template. '
            'Expected a line like: <script type="application/json" id="embeddedData">null</script>'
        )
    return new_html


def run_git(args, cwd, log):
    result = subprocess.run(['git'] + args, cwd=cwd, capture_output=True, text=True)
    log(f'$ git {" ".join(args)}')
    if result.stdout.strip():
        log(result.stdout.strip())
    if result.returncode != 0:
        log(result.stderr.strip(), err=True)
    return result.returncode == 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--master', required=True, help='Accumulated raw-punch store (xlsx).')
    ap.add_argument('--template', required=True, help='Dashboard source index.html (not modified).')
    ap.add_argument('--repo', required=True, help='Local GitHub Pages repo checkout to publish into.')
    ap.add_argument('--out-name', default='index.html', help='Filename to write inside --repo.')
    ap.add_argument('--branch', default='main', help='Branch to push.')
    ap.add_argument('--message', default=None, help='Commit message. Defaults to a generated summary.')
    ap.add_argument('--no-push', action='store_true', help='Write + stage only; skip commit and push.')
    args = ap.parse_args(argv)

    def log(msg, err=False):
        print(msg, file=sys.stderr if err else sys.stdout)

    template_path = Path(args.template)
    repo_path = Path(args.repo)
    if not template_path.exists():
        log(f'Template not found: {template_path}', err=True)
        return 1
    if not repo_path.exists():
        log(f'Repo path not found: {repo_path} (clone your GitHub Pages repo there first)', err=True)
        return 1

    punches = build_report.load_master_punches(args.master)
    if not punches:
        log(f'No punches found in master store: {args.master}', err=True)
        return 1
    log(f'Loaded {len(punches):,} punch(es) from {args.master}')

    payload = build_embedded_payload(punches)
    template_html = template_path.read_text(encoding='utf-8')
    published_html = embed_into_template(template_html, payload)

    out_path = repo_path / args.out_name
    out_path.write_text(published_html, encoding='utf-8')
    size_kb = out_path.stat().st_size / 1024
    log(f'Wrote {out_path} ({size_kb:.0f} KB, {len(punches):,} punches embedded)')

    if args.no_push:
        log('--no-push passed — file written, skipping git commit/push.')
        return 0

    # Confirm this checkout is actually a git repo before attempting anything
    check = subprocess.run(['git', 'rev-parse', '--is-inside-work-tree'], cwd=repo_path, capture_output=True, text=True)
    if check.returncode != 0:
        log(f'{repo_path} is not a git repository — file was written but not published. '
            f'Run with --no-push if this is intentional, or point --repo at your Pages clone.', err=True)
        return 1

    message = args.message or f'Update clocking dashboard — {len(punches):,} punches, generated {payload["generated"]}'
    ok = run_git(['add', args.out_name], repo_path, log)
    if not ok:
        return 1

    diff = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=repo_path)
    if diff.returncode == 0:
        log('No changes to publish (dashboard content is identical to what is already live).')
        return 0

    if not run_git(['commit', '-m', message], repo_path, log):
        return 1
    if not run_git(['push', 'origin', args.branch], repo_path, log):
        log('Commit succeeded locally but push failed — push manually once the issue is resolved '
            '(check network/auth; nothing was lost).', err=True)
        return 1

    log('Published.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
