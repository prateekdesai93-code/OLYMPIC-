#!/usr/bin/env python3
"""
process_inbox.py — Olympic Paints Clocking pipeline orchestrator
====================================================================

Watches an Inbox folder for new "Transaction*.xlsx" biometric exports,
merges them into the master punch store via build_report.py, produces
"Clocking Report YTD.xlsx", archives the processed input files, and
(optionally) emails the result.

This is meant to run on a schedule (cron / Task Scheduler) or be run by
hand after exporting from Advius.

Folder layout (created automatically on first run):

    <root>/
      Inbox/              <- drop Transaction*.xlsx exports here
      Archive/YYYY-MM/    <- processed inputs are moved here, timestamped
      master_raw_data.xlsx           <- accumulated punch history (do not edit by hand)
      Clocking Report YTD.xlsx       <- latest generated report (this is what gets emailed)
      process_inbox.log              <- run history

Usage:
    python process_inbox.py                       # process + email (if configured)
    python process_inbox.py --no-email             # process only, skip email
    python process_inbox.py --root /path/to/folder # use a different working folder
    python process_inbox.py --start 2026-06-25 --end 2026-07-25   # payroll-cycle report

Email is configured via environment variables (never hard-code credentials):
    CLOCKING_SMTP_HOST, CLOCKING_SMTP_PORT, CLOCKING_SMTP_USER, CLOCKING_SMTP_PASS,
    CLOCKING_MAIL_FROM, CLOCKING_MAIL_TO (comma-separated)
If CLOCKING_SMTP_HOST is not set, emailing is skipped automatically (with a notice)
even if --no-email wasn't passed.
"""

import argparse
import logging
import os
import shutil
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import build_report

INBOX_DIRNAME = 'Inbox'
ARCHIVE_DIRNAME = 'Archive'
MASTER_FILENAME = 'master_raw_data.xlsx'
REPORT_FILENAME = 'Clocking Report YTD.xlsx'
LOG_FILENAME = 'process_inbox.log'


def setup_logging(root: Path):
    log_path = root / LOG_FILENAME
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s  %(levelname)-7s  %(message)s',
        handlers=[logging.FileHandler(log_path, encoding='utf-8'), logging.StreamHandler(sys.stdout)],
    )
    return logging.getLogger('process_inbox')


def find_inbox_files(inbox_dir: Path):
    return sorted([p for p in inbox_dir.glob('*.xlsx') if not p.name.startswith('~$')])


def archive_files(files, root: Path, log):
    archive_month_dir = root / ARCHIVE_DIRNAME / datetime.now().strftime('%Y-%m')
    archive_month_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    for f in files:
        dest = archive_month_dir / f'{f.stem}_{stamp}{f.suffix}'
        shutil.move(str(f), str(dest))
        log.info(f'Archived {f.name} -> {dest.relative_to(root)}')


def send_email(report_path: Path, stats: dict, log):
    host = os.environ.get('CLOCKING_SMTP_HOST')
    if not host:
        log.info('CLOCKING_SMTP_HOST not set — skipping email (report was still generated).')
        return False

    port = int(os.environ.get('CLOCKING_SMTP_PORT', '587'))
    user = os.environ.get('CLOCKING_SMTP_USER')
    password = os.environ.get('CLOCKING_SMTP_PASS')
    mail_from = os.environ.get('CLOCKING_MAIL_FROM', user)
    mail_to = [addr.strip() for addr in os.environ.get('CLOCKING_MAIL_TO', '').split(',') if addr.strip()]

    if not mail_to:
        log.warning('CLOCKING_MAIL_TO is not set — skipping email.')
        return False

    msg = EmailMessage()
    msg['Subject'] = f"Clocking Report YTD — {stats['period']}"
    msg['From'] = mail_from
    msg['To'] = ', '.join(mail_to)
    msg.set_content(
        f"Olympic Paints Clocking Report attached.\n\n"
        f"{stats['period']}\n"
        f"Employees: {stats['employees']} (Olympic Paints: {stats['op_employees']} | Primeserve: {stats['ps_employees']})\n"
        f"Shift-days: {stats['records']}\n"
        f"Total hours: {stats['total_hours']}\n"
        f"Missing clock records: {stats['missing']} ({stats['missed_out']} clock-out, {stats['missed_in']} clock-in)\n\n"
        f"Generated automatically by process_inbox.py.\n"
    )
    with open(report_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application',
                            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            filename=report_path.name)

    try:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.starttls()
            if user and password:
                server.login(user, password)
            server.send_message(msg)
        log.info(f'Email sent to {", ".join(mail_to)}')
        return True
    except Exception as e:
        log.error(f'Email failed: {e}')
        return False


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--root', default='.', help='Working folder containing Inbox/, Archive/, master store, and report.')
    ap.add_argument('--no-email', action='store_true', help='Generate the report but skip emailing it.')
    ap.add_argument('--start', default=None, help='Optional start date (YYYY-MM-DD) to filter the report.')
    ap.add_argument('--end', default=None, help='Optional end date (YYYY-MM-DD) to filter the report.')
    args = ap.parse_args(argv)

    root = Path(args.root).resolve()
    inbox_dir = root / INBOX_DIRNAME
    inbox_dir.mkdir(parents=True, exist_ok=True)
    (root / ARCHIVE_DIRNAME).mkdir(parents=True, exist_ok=True)
    master_path = root / MASTER_FILENAME
    report_path = root / REPORT_FILENAME

    log = setup_logging(root)
    log.info('=' * 70)
    log.info(f'Run started. Root: {root}')

    new_files = find_inbox_files(inbox_dir)
    if not new_files:
        log.info('No new files in Inbox — nothing to merge. Regenerating report from master store only.')
    else:
        log.info(f'Found {len(new_files)} file(s) in Inbox: {", ".join(f.name for f in new_files)}')

    punches = build_report.load_master_punches(master_path)
    log.info(f'Master store: {len(punches):,} punch(es) on file.')

    for f in new_files:
        try:
            new_punches = build_report.extract_punches_from_workbook(f)
        except Exception as e:
            log.error(f'Failed to read {f.name}: {e} — leaving it in Inbox for review.')
            new_files.remove(f)
            continue
        if not new_punches:
            log.warning(f'{f.name} produced no recognizable punch rows — leaving it in Inbox for review.')
            new_files.remove(f)
            continue
        punches.extend(new_punches)
        log.info(f'Merged {len(new_punches):,} punch(es) from {f.name}')

    punches = build_report.dedupe_punches(punches)
    build_report.write_master(master_path, punches)
    log.info(f'Master store updated: {len(punches):,} punch(es) total.')

    records = build_report.build_records(punches)
    start = build_report.parse_date(args.start) if args.start else None
    end = build_report.parse_date(args.end) if args.end else None
    filtered = build_report.filter_records(records, start, end)

    if not filtered:
        log.error('No records in the selected range — aborting before overwriting the report.')
        return 1

    stats = build_report.write_report(str(report_path), filtered, punches)
    log.info(f"Report written -> {report_path.name}")
    log.info(f"  {stats['period']} | {stats['employees']} employees "
              f"(OP: {stats['op_employees']} | Primeserve: {stats['ps_employees']}) "
              f"| {stats['records']} shift-days | {stats['total_hours']} total hours")
    log.info(f"  Missing: {stats['missing']} ({stats['missed_out']} clock-out, {stats['missed_in']} clock-in)")

    if new_files:
        archive_files(new_files, root, log)

    if not args.no_email:
        send_email(report_path, stats, log)
    else:
        log.info('--no-email passed — skipping email.')

    log.info('Run complete.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
