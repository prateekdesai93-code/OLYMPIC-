#!/usr/bin/env python3
"""
build_report.py — Olympic Paints Clocking Report engine
=========================================================

Computes attendance, hours, and missing-punch data from raw biometric
punch exports and writes a multi-sheet "Clocking Report YTD.xlsx" workbook.

RULES (reverse-engineered from the original reference workbook and validated
to reproduce it exactly — employer totals, missed-in/out split, and grand
total hours all matched to the minute against a 13,705-punch sample):

  - Break:      a flat 45 minutes is deducted from every shift that has
                both a clock-in and a clock-out.
  - Snap-in:    a clock-in at or before 07:20 (time-of-day) is recorded
                as 07:15.
  - Snap-out:   a clock-out between 16:50 and 17:10 (inclusive) is
                recorded as 17:00.
  - Employer:   an Employee ID starting with "SD" is Primeserve; every
                other ID is Olympic Paints.
  - Per employee per day: punches are sorted by time. The first punch is
    the clock-in candidate, the last is the clock-out candidate. Punches
    in between are ignored for hours but still counted in Total Punches.
  - Single-punch days: ASSUMPTION — a lone punch before 15:00 is treated
    as a clock-in only (clock-out missing); at/after 15:00 it's treated
    as a clock-out only (clock-in missing). The source data shows missed
    clock-out punches topping out at 14:06 and missed clock-in punches
    starting at 15:41, so the true cutoff sits somewhere in that gap —
    15:00 is a reasonable midpoint. Adjust SINGLE_PUNCH_CUTOFF_MIN below
    if a more precise boundary is ever confirmed.

Usage:
    python build_report.py --master master_raw_data.xlsx \
        [--new Inbox/Transaction_2026-08-05.xlsx ...] \
        --out "Clocking Report YTD.xlsx" \
        [--start 2026-03-26] [--end 2026-07-31]

  --master   Path to the accumulated raw-punch store (created if absent).
             New punches from --new files are merged into it and it is
             rewritten, so it always holds full history.
  --new      One or more raw Transaction*.xlsx exports to merge in. Can be
             repeated or given as a glob the shell expands.
  --out      Output workbook path. Defaults to "Clocking Report YTD.xlsx".
  --start/--end  Optional YYYY-MM-DD bounds. Restricts every sheet to this
             date range (the master store itself is never filtered — only
             the generated report is). Omit for the full period on file.
"""

import argparse
import sys
from collections import defaultdict
from datetime import datetime, date, timedelta
from pathlib import Path

import openpyxl
import pandas as pd

# --------------------------------------------------------------------------
# Rules
# --------------------------------------------------------------------------
SNAP_IN_MAX_MIN = 7 * 60 + 20       # 07:20
SNAP_IN_TO_MIN = 7 * 60 + 15        # 07:15
SNAP_OUT_MIN_MIN = 16 * 60 + 50     # 16:50
SNAP_OUT_MAX_MIN = 17 * 60 + 10     # 17:10
SNAP_OUT_TO_MIN = 17 * 60           # 17:00
SINGLE_PUNCH_CUTOFF_MIN = 15 * 60   # 15:00 — see assumption above
BREAK_MINUTES = 45

WEEKDAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

RAW_COLUMNS = ['First Name', 'Last Name', 'ID', 'Department', 'Date', 'Time', 'Weekday',
               'Data Source', 'Device Name', 'Device Serial No.', 'Punch State', 'Location', 'Remarks']


# --------------------------------------------------------------------------
# Time / date helpers
# --------------------------------------------------------------------------
def parse_clock(raw):
    """'07:15' / '7:15:00' / datetime.time -> minutes since 00:00, or None."""
    if raw is None or raw == '':
        return None
    if hasattr(raw, 'hour'):  # datetime.time or datetime.datetime
        return raw.hour * 60 + raw.minute
    s = str(raw).strip()
    parts = s.split(':')
    if len(parts) < 2:
        return None
    try:
        h, m = int(parts[0]), int(parts[1])
        return h * 60 + m
    except ValueError:
        return None


def fmt_clock(mins):
    """minutes-of-day -> '07:15' (zero-padded hour)."""
    if mins is None:
        return ''
    h, m = divmod(round(mins), 60)
    return f'{h:02d}:{m:02d}'


def fmt_duration(mins):
    """duration-in-minutes -> '9:45' (no leading zero on hours, can exceed 24h)."""
    if mins is None:
        return ''
    neg = mins < 0
    mins = round(abs(mins))
    h, m = divmod(mins, 60)
    return f'{"-" if neg else ""}{h}:{m:02d}'


def to_decimal_hours(mins):
    if mins is None:
        return None
    return round(mins / 60, 2)


def parse_date(raw):
    """Accepts 'YYYY/MM/DD', 'YYYY-MM-DD', or a date/datetime -> date, or None."""
    if raw is None or raw == '':
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw).strip()
    for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def date_key(d):
    return d.strftime('%Y/%m/%d')


def date_short(d):
    return f'{d.day} {MONTH_NAMES[d.month-1]} {d.year}'


def week_start(d):
    """Most recent Wednesday on/before d (weeks run Wed-Tue)."""
    dow = d.weekday()  # Mon=0 .. Sun=6
    diff = (dow - 2) % 7  # days since most recent Wednesday (Wed=2)
    return d - timedelta(days=diff)


def week_label(d):
    ws = week_start(d)
    we = ws + timedelta(days=6)
    return f'{ws.day} {MONTH_NAMES[ws.month-1]}–{we.day} {MONTH_NAMES[we.month-1]}'


def month_key(d):
    return f'{MONTH_NAMES[d.month-1]} {d.year}'


def _clean_id(raw):
    """Normalizes an Employee ID cell. Long numeric IDs (e.g. biometric
    device serials like 7403056768186) are sometimes read back from Excel
    as floats rather than ints — str(7403056768186.0) would otherwise
    produce '7403056768186.0' and silently split one employee into two
    different IDs across imports. Whole-number floats are coerced to int
    first so the ID stays stable no matter how the source file typed it."""
    if isinstance(raw, float) and raw.is_integer():
        return str(int(raw))
    return str(raw).strip()


def employer_from_id(emp_id):
    return 'Primeserve' if _clean_id(emp_id).upper().startswith('SD') else 'Olympic Paints'


def snap_in(mins):
    return SNAP_IN_TO_MIN if mins <= SNAP_IN_MAX_MIN else mins


def snap_out(mins):
    return SNAP_OUT_TO_MIN if SNAP_OUT_MIN_MIN <= mins <= SNAP_OUT_MAX_MIN else mins


# --------------------------------------------------------------------------
# Raw punch import (auto-detects header row + column order)
# --------------------------------------------------------------------------
def _normalize(h):
    return ''.join(ch for ch in str(h or '').lower() if ch.isalnum())


def extract_punches_from_workbook(path):
    """Reads every sheet of an xlsx export and returns a list of punch dicts."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    punches = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        header_idx, header = None, None
        for i, row in enumerate(rows[:10]):
            norm = [_normalize(c) for c in row]
            if 'id' in norm and 'date' in norm and ('time' in norm or 'clockin' in norm):
                header_idx, header = i, row
                break
        if header_idx is None:
            continue

        def find_col(*names):
            norm = [_normalize(c) for c in header]
            for n in names:
                nn = _normalize(n)
                if nn in norm:
                    return norm.index(nn)
            return None

        c_first = find_col('First Name', 'FirstName')
        c_last = find_col('Last Name', 'LastName')
        c_id = find_col('ID', 'Employee ID', 'EmployeeID')
        c_dept = find_col('Department')
        c_date = find_col('Date')
        c_time = find_col('Time')
        if c_id is None or c_date is None or c_time is None:
            continue

        for row in rows[header_idx + 1:]:
            if row is None or c_id >= len(row):
                continue
            emp_id = row[c_id]
            if emp_id is None or str(emp_id).strip() == '':
                continue
            d = parse_date(row[c_date])
            t = parse_clock(row[c_time])
            if d is None or t is None:
                continue
            punches.append({
                'id': _clean_id(emp_id),
                'first_name': str(row[c_first]).strip() if c_first is not None and row[c_first] else '',
                'last_name': str(row[c_last]).strip() if c_last is not None and row[c_last] else '',
                'department': str(row[c_dept]).strip() if c_dept is not None and row[c_dept] else 'Unassigned',
                'date': d,
                'minutes': t,
            })
    wb.close()
    return punches


def load_master_punches(master_path):
    if not Path(master_path).exists():
        return []
    return extract_punches_from_workbook(master_path)


def dedupe_punches(punches):
    seen = set()
    out = []
    for p in punches:
        key = (p['id'], p['date'], p['minutes'], p['department'])
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def write_master(master_path, punches):
    """Writes the accumulated raw-punch store as a single 'Raw Data' sheet."""
    rows = []
    for p in sorted(punches, key=lambda p: (p['date'], p['id'], p['minutes'])):
        rows.append([
            p['first_name'], p['last_name'], p['id'], p['department'],
            date_key(p['date']), fmt_clock(p['minutes']), WEEKDAY_NAMES[p['date'].weekday()],
            'Device', '', '', '--', '--', '--',
        ])
    df = pd.DataFrame(rows, columns=RAW_COLUMNS)
    with pd.ExcelWriter(master_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)


# --------------------------------------------------------------------------
# Core engine — builds one record per employee/day
# --------------------------------------------------------------------------
def build_records(punches):
    groups = defaultdict(list)
    for p in punches:
        key = (p['id'], p['date'])
        groups[key].append(p)

    records = []
    for (emp_id, d), day_punches in groups.items():
        # Sort the day's punches by time — this determines clock-in/out AND which
        # punch's name/department apply. A day can have conflicting department
        # tags across punches (an employee transferred departments mid-period);
        # the clock-in punch's tag is the one that counts, matching payroll
        # practice (confirmed against the reference workbook: an employee who
        # moved from "Drivers" to "Monthly Employees" mid-year, with a stray
        # early punch still tagged with the old department, was filed under
        # the old department for that day — i.e. by clock-in, not last-seen).
        day_punches.sort(key=lambda p: p['minutes'])
        times = [p['minutes'] for p in day_punches]
        first, last = day_punches[0], day_punches[-1]
        total_punches = len(times)
        clock_in = clock_out = miss_type = None

        if total_punches == 1:
            t = times[0]
            if t < SINGLE_PUNCH_CUTOFF_MIN:
                clock_in, miss_type = snap_in(t), 'out'
            else:
                clock_out, miss_type = snap_out(t), 'in'
        else:
            clock_in = snap_in(times[0])
            clock_out = snap_out(times[-1])

        gross_min = break_min = net_min = None
        if clock_in is not None and clock_out is not None:
            g = clock_out - clock_in
            if g < 0:
                g += 24 * 60  # guard against an overnight shift crossing midnight
            gross_min = g
            break_min = BREAK_MINUTES
            net_min = max(0, gross_min - break_min)

        records.append({
            'date': d, 'weekday': WEEKDAY_NAMES[d.weekday()],
            'first_name': first['first_name'] or last['first_name'],
            'last_name': first['last_name'] or last['last_name'],
            'id': emp_id,
            'department': first['department'] or last['department'] or 'Unassigned',
            'employer': employer_from_id(emp_id),
            'clock_in': clock_in, 'clock_out': clock_out,
            'gross_min': gross_min, 'break_min': break_min, 'net_min': net_min,
            'total_punches': total_punches, 'miss_type': miss_type,
        })

    records.sort(key=lambda r: (r['date'], r['id']))
    return records


def filter_records(records, start=None, end=None):
    return [r for r in records if (start is None or r['date'] >= start) and (end is None or r['date'] <= end)]


# --------------------------------------------------------------------------
# Aggregations
# --------------------------------------------------------------------------
def employer_summary(records):
    buckets = {'Olympic Paints': defaultdict(lambda: 0), 'Primeserve': defaultdict(lambda: 0)}
    employees = {'Olympic Paints': set(), 'Primeserve': set()}
    for r in records:
        b = buckets[r['employer']]
        employees[r['employer']].add(r['id'])
        b['shift_days'] += 1
        if r['miss_type']:
            b['missing'] += 1
        if r['net_min'] is not None:
            b['net_min'] += r['net_min']
    out = []
    for employer in ('Olympic Paints', 'Primeserve'):
        b = buckets[employer]
        shift_days = b['shift_days']
        out.append({
            'employer': employer, 'employees': len(employees[employer]), 'shift_days': shift_days,
            'missing': b['missing'], 'net_min': b['net_min'],
            'avg_per_shift': b['net_min'] / shift_days if shift_days else 0,
            'pct_missing': (b['missing'] / shift_days * 100) if shift_days else 0,
        })
    return out


def daily_summary(records):
    buckets = {}
    for r in records:
        k = r['date']
        if k not in buckets:
            buckets[k] = {'date': k, 'weekday': r['weekday'], 'employees': set(), 'punches': 0, 'missing': 0, 'net_min': 0}
        b = buckets[k]
        b['employees'].add(r['id'])
        b['punches'] += r['total_punches']
        if r['miss_type']:
            b['missing'] += 1
        if r['net_min'] is not None:
            b['net_min'] += r['net_min']
    out = []
    for b in sorted(buckets.values(), key=lambda x: x['date']):
        n = len(b['employees'])
        out.append({**b, 'employees': n, 'avg_per_person': (b['net_min'] / n) if n else 0})
    return out


def department_summary(records):
    buckets = {}
    for r in records:
        k = (r['employer'], r['department'])
        if k not in buckets:
            buckets[k] = {'employer': r['employer'], 'department': r['department'], 'employees': set(), 'shift_days': 0, 'missing': 0, 'net_min': 0}
        b = buckets[k]
        b['employees'].add(r['id'])
        b['shift_days'] += 1
        if r['miss_type']:
            b['missing'] += 1
        if r['net_min'] is not None:
            b['net_min'] += r['net_min']
    out = []
    for b in buckets.values():
        n = len(b['employees'])
        out.append({**b, 'employees': n,
                     'avg_per_shift': (b['net_min'] / b['shift_days']) if b['shift_days'] else 0,
                     'pct_missing': (b['missing'] / b['shift_days'] * 100) if b['shift_days'] else 0})
    out.sort(key=lambda d: -d['net_min'])
    return out


def weekly_summary(records):
    by_employee = {}
    week_keys, month_keys = {}, {}
    for r in records:
        if r['net_min'] is None:
            continue
        wk = week_start(r['date'])
        week_keys.setdefault(wk, week_label(r['date']))
        mk = month_key(r['date'])
        month_keys.setdefault(mk, r['date'].replace(day=1))

        e = by_employee.setdefault(r['id'], {
            'employer': r['employer'], 'id': r['id'], 'first_name': r['first_name'], 'last_name': r['last_name'],
            'weeks': defaultdict(float), 'months': defaultdict(float), 'total': 0.0,
        })
        e['weeks'][wk] += r['net_min']
        e['months'][mk] += r['net_min']
        e['total'] += r['net_min']

    week_order = sorted(week_keys.items(), key=lambda kv: kv[0])
    month_order = sorted(month_keys.items(), key=lambda kv: kv[1])
    employees = sorted(by_employee.values(), key=lambda e: (e['employer'], e['first_name']))
    return employees, week_order, month_order


def missing_log(records):
    out = []
    for r in records:
        if not r['miss_type']:
            continue
        out.append({
            'date': r['date'], 'weekday': r['weekday'], 'first_name': r['first_name'], 'last_name': r['last_name'],
            'id': r['id'], 'department': r['department'], 'employer': r['employer'],
            'issue': 'Missed Clock-In' if r['miss_type'] == 'in' else 'Missed Clock-Out',
            'punch_time': fmt_clock(r['clock_out'] if r['miss_type'] == 'in' else r['clock_in']),
            'total_punches': r['total_punches'],
        })
    out.sort(key=lambda m: m['date'])
    return out


# --------------------------------------------------------------------------
# Workbook writer — matches the reference file's sheet names, headers, and
# plain-value convention (no live formulas — this is a computed snapshot).
# --------------------------------------------------------------------------
def write_report(out_path, records, punches, period_note_extra=''):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    employees = {r['id'] for r in records}
    op_employees = {r['id'] for r in records if r['employer'] == 'Olympic Paints'}
    ps_employees = {r['id'] for r in records if r['employer'] == 'Primeserve'}
    dates = [r['date'] for r in records]
    period_str = f'Period: {date_short(min(dates))} – {date_short(max(dates))}' if dates else 'Period: —'

    def write_sheet(name, title, subtitle_lines, header, rows):
        ws = wb.create_sheet(name)
        ws.append([title])
        for line in subtitle_lines:
            ws.append([line])
        ws.append([])
        ws.append(header)
        for row in rows:
            ws.append(row)
        ws.freeze_panes = 'A5'
        for col_cells in ws.columns:
            length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 42)

    # 1. Clocking Report (master)
    cr_header = ['Date', 'Day', 'First Name', 'Last Name', 'Employee ID', 'Department', 'Employer',
                 'Clock In', 'Clock Out', 'Gross Hours', 'Break Deducted', 'Hours Worked', 'Total Punches']
    cr_rows = [[
        date_key(r['date']), r['weekday'], r['first_name'], r['last_name'], r['id'], r['department'], r['employer'],
        fmt_clock(r['clock_in']), fmt_clock(r['clock_out']),
        fmt_duration(r['gross_min']), fmt_duration(r['break_min']), fmt_duration(r['net_min']), r['total_punches'],
    ] for r in records]
    write_sheet('Clocking Report', 'Olympic Paints — Clocking Report', [
        f'{period_str}    |    Employees: {len(employees)} (Olympic Paints: {len(op_employees)}  |  Primeserve: {len(ps_employees)})    |    Records: {len(records)}    |    Last updated: {date_short(date.today())}{period_note_extra}',
        'Hours Worked is NET of a fixed 45 min break deduction per day (lunch + tea). Gross Hours shows the snapped clock-in to clock-out span. '
        'Clock-in ≤ 07:20 is recorded as 07:15. Clock-out 16:50–17:10 is recorded as 17:00. Times outside these windows are recorded as-is.',
    ], cr_header, cr_rows)

    # 2. Summary by Date
    daily = daily_summary(records)
    sd_header = ['Date', 'Weekday', 'Employees', 'Total Punches', 'Missing Clock Out', 'Total Hours', 'Avg Hrs/Person']
    sd_rows = [[date_key(d['date']), d['weekday'], d['employees'], d['punches'], d['missing'],
                fmt_duration(d['net_min']), fmt_duration(d['avg_per_person'])] for d in daily]
    sd_rows.append(['TOTAL / AVERAGE', '', len(employees), sum(d['punches'] for d in daily),
                     sum(d['missing'] for d in daily), fmt_duration(sum(d['net_min'] for d in daily)), ''])
    write_sheet('Summary by Date', 'Olympic Paints — Daily Summary',
                [f'{period_str}   |   Hours are net of 45 min/day break (lunch + tea)'], sd_header, sd_rows)

    # 3. Summary by Employer
    emp = employer_summary(records)
    se_header = ['Employer', 'Unique Employees', 'Shift-Days', 'Missing Clock Out', 'Total Hours', 'Avg Hrs/Shift', '% Missing']
    se_rows = [[e['employer'], e['employees'], e['shift_days'], e['missing'],
                fmt_duration(e['net_min']), fmt_duration(e['avg_per_shift']), f"{e['pct_missing']:.1f}%"] for e in emp]
    write_sheet('Summary by Employer', 'Olympic Paints — Employer Summary',
                [f'{period_str}   |   Hours are net of 45 min/day break (lunch + tea)'], se_header, se_rows)

    # 4. Weekly Summary
    wk_employees, week_order, month_order = weekly_summary(records)
    ws_header = ['Employer', 'Employee ID', 'First Name', 'Last Name'] + \
                [label for _, label in week_order] + [mk for mk, _ in month_order] + ['Grand Total']
    ws_rows = []
    for e in wk_employees:
        row = [e['employer'], e['id'], e['first_name'], e['last_name']]
        row += [(to_decimal_hours(e['weeks'][wk]) if wk in e['weeks'] else '') for wk, _ in week_order]
        row += [(to_decimal_hours(e['months'][mk]) if mk in e['months'] else '') for mk, _ in month_order]
        row.append(to_decimal_hours(e['total']))
        ws_rows.append(row)
    totals_row = ['TOTAL', '', '', '']
    totals_row += [to_decimal_hours(sum(e['weeks'].get(wk, 0) for e in wk_employees)) for wk, _ in week_order]
    totals_row += [to_decimal_hours(sum(e['months'].get(mk, 0) for e in wk_employees)) for mk, _ in month_order]
    totals_row.append(to_decimal_hours(sum(e['total'] for e in wk_employees)))
    ws_rows.append(totals_row)
    write_sheet('Weekly Summary', 'Olympic Paints — Weekly Hours Summary by Employee', [
        f'Date filter: {date_short(min(dates))} → {date_short(max(dates))}   |   Weeks shown: {len(week_order)}   |   To filter: use --start / --end flags when running the script' if dates else '',
        'Hours shown are NET (45 min break deducted per day).  Filter column A (Employer) to compare Primeserve vs Olympic Paints.',
    ], ws_header, ws_rows)

    # 5. Summary by Department
    depts = department_summary(records)
    dd_header = ['Employer', 'Department', 'Unique Employees', 'Shift-Days', 'Missing Clock Out', 'Total Hours', 'Avg Hrs/Shift', '% Missing']
    dd_rows = [[d['employer'], d['department'], d['employees'], d['shift_days'], d['missing'],
                fmt_duration(d['net_min']), fmt_duration(d['avg_per_shift']), f"{d['pct_missing']:.1f}%"] for d in depts]
    dd_rows.append(['', 'TOTAL', len(employees), sum(d['shift_days'] for d in depts),
                     sum(d['missing'] for d in depts), fmt_duration(sum(d['net_min'] for d in depts)), '', ''])
    write_sheet('Summary by Department', 'Olympic Paints — Department Summary',
                [f'{period_str}   |   Hours are net of 45 min/day break (lunch + tea)'], dd_header, dd_rows)

    # 6. Missing Clock Out
    log = missing_log(records)
    missed_out = sum(1 for m in log if m['issue'] == 'Missed Clock-Out')
    missed_in = sum(1 for m in log if m['issue'] == 'Missed Clock-In')
    mc_header = ['Date', 'Day', 'First Name', 'Last Name', 'Employee ID', 'Department', 'Employer', 'Miss Type', 'Clock Time', 'Total Punches']
    mc_rows = [[date_key(m['date']), m['weekday'], m['first_name'], m['last_name'], m['id'], m['department'],
                m['employer'], m['issue'], m['punch_time'], m['total_punches']] for m in log]
    write_sheet('Missing Clock Out', 'Olympic Paints — Missing Clocking Records',
                [f'{len(log)} missing clocking record(s): {missed_out} missed clock-out, {missed_in} missed clock-in.'], mc_header, mc_rows)

    # 7. Raw Data (full accumulated history — not filtered by --start/--end)
    rd_rows = [[p['first_name'], p['last_name'], p['id'], p['department'], date_key(p['date']),
                fmt_clock(p['minutes']), WEEKDAY_NAMES[p['date'].weekday()], 'Device', '', '', '--', '--', '--']
               for p in sorted(punches, key=lambda p: (p['date'], p['id'], p['minutes']))]
    write_sheet('Raw Data', f'Accumulated Raw Punch Data — {period_str}', [], RAW_COLUMNS, rd_rows)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return {
        'records': len(records), 'employees': len(employees),
        'op_employees': len(op_employees), 'ps_employees': len(ps_employees),
        'missing': len(log), 'missed_in': missed_in, 'missed_out': missed_out,
        'total_hours': fmt_duration(sum(r['net_min'] or 0 for r in records)),
        'period': period_str,
    }


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--master', required=True, help='Accumulated raw-punch store (xlsx). Created if it does not exist.')
    ap.add_argument('--new', nargs='*', default=[], help='New raw Transaction*.xlsx exports to merge in.')
    ap.add_argument('--out', default='Clocking Report YTD.xlsx', help='Output report workbook path.')
    ap.add_argument('--start', default=None, help='Optional start date (YYYY-MM-DD) to filter the report.')
    ap.add_argument('--end', default=None, help='Optional end date (YYYY-MM-DD) to filter the report.')
    args = ap.parse_args(argv)

    punches = load_master_punches(args.master)
    print(f'Master store: {len(punches):,} punch(es) on file.')

    for path in args.new:
        new_punches = extract_punches_from_workbook(path)
        punches.extend(new_punches)
        print(f'Merged {len(new_punches):,} punch(es) from {path}')

    punches = dedupe_punches(punches)
    write_master(args.master, punches)
    print(f'Master store updated: {len(punches):,} punch(es) total -> {args.master}')

    records = build_records(punches)
    start = parse_date(args.start) if args.start else None
    end = parse_date(args.end) if args.end else None
    filtered = filter_records(records, start, end)
    if not filtered:
        print('No records in range — nothing to write.', file=sys.stderr)
        return 1

    stats = write_report(args.out, filtered, punches)
    print(f"Report written -> {args.out}")
    print(f"  {stats['period']}")
    print(f"  {stats['employees']} employees (OP: {stats['op_employees']} | Primeserve: {stats['ps_employees']}) "
          f"| {stats['records']} shift-days | {stats['total_hours']} total hours")
    print(f"  Missing: {stats['missing']} ({stats['missed_out']} clock-out, {stats['missed_in']} clock-in)")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
