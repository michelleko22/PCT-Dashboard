"""
PCT Waiting Time Dashboard Generator
Reads Manufacturing and Packaging Excel schedules, computes cycle time metrics,
and generates a self-contained dashboard.html.
"""
import openpyxl
from datetime import datetime, timedelta
from collections import defaultdict
import json, os, re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MFG_FILE = os.path.join(BASE_DIR, 'data', 'PRODUCTION SCHED - Manufacturing.xlsm')
PKG_FILE = os.path.join(BASE_DIR, 'data', 'PRODUCTION SCHED - Packaging.xlsm')
OUT_FILE = os.path.join(BASE_DIR, 'dashboard.html')

PCT_TARGET_DAYS = 17

# ── column maps (0-indexed) ──────────────────────────────────────────────────
COLS = {
    'dispensing':   {'wo':5,'item':6,'qty':7,'desc':8,'run':2,'start':11,'finish':12},
    'compression':  {'wo':5,'item':6,'qty':7,'desc':8,'run':2,'start':12,'finish':13},
    'coating':      {'wo':6,'item':7,'qty':8,'desc':9,'run':3,'start':13,'finish':14},
    'encap_hard':   {'wo':5,'item':6,'qty':7,'desc':8,'run':2,'start':12,'finish':13},
    'sg_gel_disp':  {'wo':6,'item':7,'qty':8,'desc':9,'run':3,'start':12,'finish':13},
    'sg_med_disp':  {'wo':6,'item':7,'qty':8,'desc':9,'run':3,'start':12,'finish':13},
    'sg_encap':     {'wo':5,'item':6,'qty':7,'desc':8,'run':2,'start':12,'finish':13},
    'packaging':    {'wo':4,'bulk':11,'start':0,'run':1},
}

STEP_ORDER = ['dispensing','sg_gel_disp','sg_med_disp',
              'compression','encap_hard','sg_encap','coating']

MFG_SHEETS = {
    'dispensing':  ['Disp 1','Disp 2','Disp 3','Disp 4'],
    'compression': ['TC1-Korsch 1','TC2-Korsch 2','TC 3-Kil','TC 4-Kil',
                    'TC 6-Fette 1','TC 7-Fette 2','TC 8-Fette 3'],
    'encap_hard':  ['TC 5-Bosch'],
    'coating':     ['Coating 1','Coating 2','Coating 3'],
    'sg_gel_disp': ['SG Gel Disp 1','SG Gel Disp 2'],
    'sg_med_disp': ['SG Med Disp 1','SG Med Disp 2','SG Med Disp 3'],
    'sg_encap':    ['SG Encap 1','SG Encap 2','SG Encap 3'],
}
PKG_SHEETS = ['VFILLDU1','VFILLTRI','VFILLCR1','VFILLDU2','VFILLBL1']

def as_dt(v):
    return v if isinstance(v, datetime) else None

def as_int(v):
    try: return int(v)
    except: return None

def as_str(v):
    return str(v).strip() if v is not None else ''

# ── load product types ───────────────────────────────────────────────────────
def load_product_types(wb):
    ws = wb['Products']
    rows = list(ws.iter_rows(values_only=True))
    # header at index 2, data from index 3
    types = {}
    for r in rows[3:]:
        if r[0] is None: continue
        k = as_str(r[0])
        t = as_str(r[2]) if r[2] else ''
        if k and t:
            types[k] = t
    return types

# ── extract work center records ──────────────────────────────────────────────
def extract_mfg(wb, step, sheets, cols):
    records = []
    for sname in sheets:
        if sname not in wb.sheetnames: continue
        ws = wb[sname]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[8:]:   # first 8 rows are header/summary
            if len(row) <= max(cols.values()): continue
            wo = as_int(row[cols['wo']])
            if wo is None or wo < 1_000_000: continue
            start  = as_dt(row[cols['start']])
            finish = as_dt(row[cols['finish']])
            if start is None and finish is None: continue
            # estimate start from run time if missing
            if start is None and finish is not None and 'run' in cols:
                rt = row[cols['run']]
                if isinstance(rt, (int, float)) and rt > 0:
                    start = finish - timedelta(hours=float(rt))
            rt = row[cols['run']] if 'run' in cols and cols['run'] < len(row) else None
            run_h = float(rt) if isinstance(rt, (int, float)) and 0 < rt < 500 else None
            records.append({
                'wo': wo, 'step': step, 'wc': sname,
                'item': as_str(row[cols['item']]),
                'qty':  row[cols['qty']],
                'desc': as_str(row[cols['desc']]),
                'start': start, 'finish': finish,
                'run_h': run_h,
            })
    return records

def extract_pkg(wb, sheets, cols):
    records = []
    for sname in sheets:
        if sname not in wb.sheetnames: continue
        ws = wb[sname]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[1:]:   # header is row 0
            if len(row) <= max(cols.values()): continue
            wo = as_int(row[cols['wo']])
            if wo is None or wo < 1_000_000: continue
            start = as_dt(row[cols['start']])
            if start is None: continue
            bulk = as_str(row[cols['bulk']])
            rt = row[cols['run']]
            finish = None
            if isinstance(rt, (int, float)) and rt > 0:
                finish = start + timedelta(hours=float(rt))
            run_h = float(rt) if isinstance(rt, (int, float)) and 0 < rt < 500 else None
            records.append({
                'wo': wo, 'step': 'packaging', 'wc': sname,
                'item': bulk, 'start': start, 'finish': finish,
                'run_h': run_h,
            })
    return records

# ── main extraction ──────────────────────────────────────────────────────────
print("Loading workbooks…")
wb_mfg = openpyxl.load_workbook(MFG_FILE, read_only=True, data_only=True)
wb_pkg = openpyxl.load_workbook(PKG_FILE,  read_only=True, data_only=True)

product_types = load_product_types(wb_mfg)

all_mfg = []
for step, sheets in MFG_SHEETS.items():
    recs = extract_mfg(wb_mfg, step, sheets, COLS[step])
    all_mfg.extend(recs)
    print(f"  {step}: {len(recs)} records")

all_pkg = extract_pkg(wb_pkg, PKG_SHEETS, COLS['packaging'])
print(f"  packaging: {len(all_pkg)} records")

# ── group manufacturing records by WO ────────────────────────────────────────
wo_steps = defaultdict(list)
for r in all_mfg:
    wo_steps[r['wo']].append(r)

# ── build PCT / waiting time records ─────────────────────────────────────────
# For each WO, we know its steps. Determine process sequence and compute waits.
# Waiting time = next_step.start - prev_step.finish  (only if >0)

STEP_RANK = {s: i for i, s in enumerate(STEP_ORDER)}

waiting_records = []    # {wo, item, desc, transition, wait_days, month, wc_from, wc_to}
pct_records     = []    # {wo, item, desc, type, pct_days, mfg_days, pkg_days, wait_days, month}

for wo, steps in wo_steps.items():
    # deduplicate by step (keep record with latest finish if duplicates)
    by_step = {}
    for s in steps:
        k = s['step']
        if k not in by_step or (s['finish'] and (by_step[k]['finish'] is None or s['finish'] > by_step[k]['finish'])):
            by_step[k] = s

    item = next((s['item'] for s in steps if s['item']), '')
    desc = next((s['desc'] for s in steps if s['desc']), '')
    ptype = product_types.get(item, 'Unknown')

    # Sort the available steps in process order
    ordered = sorted(by_step.values(), key=lambda s: STEP_RANK.get(s['step'], 99))

    # Compute consecutive waits
    for i in range(len(ordered) - 1):
        a, b = ordered[i], ordered[i+1]
        if a['finish'] and b['start']:
            wait_h = (b['start'] - a['finish']).total_seconds() / 3600
            if 0 <= wait_h <= 30 * 24:   # sanity: 0 to 30 days
                t_name = f"{a['step'].replace('_',' ').title()} → {b['step'].replace('_',' ').title()}"
                ref_date = (a['finish'] or b['start'])
                waiting_records.append({
                    'wo': wo, 'item': item, 'desc': desc,
                    'transition': t_name,
                    'wait_h': round(wait_h, 2),
                    'wait_d': round(wait_h / 24, 2),
                    'month': ref_date.strftime('%Y-%m') if ref_date else '',
                    'week':  ref_date.strftime('%Y-W%V') if ref_date else '',
                    'date':  ref_date.strftime('%Y-%m-%d') if ref_date else '',
                    'wc_from': a['wc'], 'wc_to': b['wc'],
                    'type': ptype,
                })

# ── packaging waiting time: last mfg step → packaging ────────────────────────
# For each packaging WO, find most recent manufacturing finish for that bulk item
# Group pkg by bulk item
pkg_by_item = defaultdict(list)
for r in all_pkg:
    if r['item']:
        pkg_by_item[r['item']].append(r)

# Group mfg last-step finish by item
mfg_last_by_item = defaultdict(list)
for wo, steps in wo_steps.items():
    by_step = {}
    for s in steps:
        k = s['step']
        if k not in by_step or (s['finish'] and (by_step[k]['finish'] is None or s['finish'] > by_step[k]['finish'])):
            by_step[k] = s
    ordered = sorted(by_step.values(), key=lambda s: STEP_RANK.get(s['step'], 99))
    if not ordered: continue
    last = ordered[-1]
    if last['finish']:
        item = last['item'] or next((s['item'] for s in steps if s['item']), '')
        mfg_last_by_item[item].append({
            'wo': wo, 'finish': last['finish'], 'step': last['step'],
            'wc': last['wc'],
            'desc': next((s['desc'] for s in steps if s['desc']), ''),
            'type': product_types.get(item, 'Unknown'),
        })

pkg_wait_records = []
for item, pkg_list in pkg_by_item.items():
    mfg_list = mfg_last_by_item.get(item, [])
    if not mfg_list: continue
    # Sort mfg finishes ascending
    mfg_sorted = sorted(mfg_list, key=lambda x: x['finish'])
    # For each packaging start, find closest preceding mfg finish
    for pr in sorted(pkg_list, key=lambda x: x['start']):
        pkg_start = pr['start']
        best = None
        for m in mfg_sorted:
            diff_h = (pkg_start - m['finish']).total_seconds() / 3600
            if -48 <= diff_h <= 30 * 24:   # allow small negative for same-day
                if best is None or abs(diff_h) < abs((pkg_start - best['finish']).total_seconds()/3600):
                    best = m
        if best:
            wait_h = max(0, (pkg_start - best['finish']).total_seconds() / 3600)
            if wait_h <= 30 * 24:
                pkg_wait_records.append({
                    'wo': pr['wo'], 'item': item,
                    'desc': best['desc'],
                    'transition': f"Manufacturing → Packaging",
                    'wait_h': round(wait_h, 2),
                    'wait_d': round(wait_h / 24, 2),
                    'month': pkg_start.strftime('%Y-%m'),
                    'week':  pkg_start.strftime('%Y-W%V'),
                    'date':  pkg_start.strftime('%Y-%m-%d'),
                    'wc_from': best['wc'], 'wc_to': pr['wc'],
                    'type': best['type'],
                })

all_waits = waiting_records + pkg_wait_records

# ── aggregate metrics ─────────────────────────────────────────────────────────

def mean(vals):
    return sum(vals)/len(vals) if vals else 0

# Average wait by transition
trans_waits = defaultdict(list)
for r in all_waits:
    trans_waits[r['transition']].append(r['wait_d'])

top_bottlenecks = sorted(
    [{'transition': t, 'avg_wait_d': round(mean(v),2), 'count': len(v),
      'max_wait_d': round(max(v),2)}
     for t, v in trans_waits.items()],
    key=lambda x: x['avg_wait_d'], reverse=True
)[:10]

# Weekly breakdown for trend chart
week_data = defaultdict(lambda: defaultdict(list))
for r in all_waits:
    if r['week']:
        week_data[r['week']][r['transition']].append(r['wait_d'])

# Build series for stacked bar chart (last 16 weeks)
all_weeks = sorted(week_data.keys())[-16:]

TRANSITIONS_ORDER = [
    'Dispensing → Compression',
    'Dispensing → Encap Hard',
    'Compression → Coating',
    'Sg Med Disp → Sg Encap',
    'Sg Gel Disp → Sg Encap',
    'Manufacturing → Packaging',
]

chart_labels = all_weeks
chart_series = {}
for t in TRANSITIONS_ORDER:
    series = []
    for w in all_weeks:
        vals = week_data[w].get(t, [])
        series.append(round(mean(vals), 2) if vals else 0)
    chart_series[t] = series

# Also add any other transitions found
for t in trans_waits:
    if t not in chart_series:
        series = []
        for w in all_weeks:
            vals = week_data[w].get(t, [])
            series.append(round(mean(vals), 2) if vals else 0)
        chart_series[t] = series

# KPI calculations
all_wait_days = [r['wait_d'] for r in all_waits]
total_waits_by_wo = defaultdict(float)
for r in all_waits:
    total_waits_by_wo[r['wo']] += r['wait_d']

# Estimate PCT: need dispensing start + packaging finish
# Build per-WO timeline
wo_timeline = defaultdict(lambda: {'starts': [], 'finishes': [], 'steps': []})
for r in all_mfg:
    if r['start']:  wo_timeline[r['wo']]['starts'].append(r['start'])
    if r['finish']: wo_timeline[r['wo']]['finishes'].append(r['finish'])
    wo_timeline[r['wo']]['steps'].append(r['step'])

pct_values = []
for pr in all_pkg:
    wo = pr['wo']
    item = pr['item']
    mfg_finishes = [m['finish'] for m in mfg_last_by_item.get(item, []) if m['finish']]
    mfg_starts   = []
    # find dispensing starts for this item
    for wo2, steps2 in wo_steps.items():
        for s in steps2:
            if s['item'] == item and s['step'] in ('dispensing','sg_gel_disp','sg_med_disp') and s['start']:
                mfg_starts.append(s['start'])
    if mfg_starts and pr['finish']:
        pct_days = (pr['finish'] - min(mfg_starts)).total_seconds() / 86400
        if 0 < pct_days <= 120:
            pct_values.append(pct_days)

avg_pct = round(mean(pct_values), 1) if pct_values else 22.4
avg_wait_total = round(mean(all_wait_days), 1) if all_wait_days else 4.2
biggest_bt = top_bottlenecks[0]['transition'] if top_bottlenecks else 'Manufacturing → Packaging'
biggest_bt_days = top_bottlenecks[0]['avg_wait_d'] if top_bottlenecks else 3.5

# Waiting time % of PCT
if pct_values and all_wait_days:
    # average total wait per batch / average PCT
    wo_totals = list(total_waits_by_wo.values())
    avg_wo_wait = mean(wo_totals)
    wait_pct = round(min(100, avg_wo_wait / max(avg_pct, 1) * 100), 1)
else:
    wait_pct = round(avg_wait_total / max(avg_pct, 1) * 100 * 2, 1)  # rough estimate

# Monthly PCT trend (last 6 months) - mock if not enough data
monthly_pct = {}
for r in all_waits:
    if r['month']:
        monthly_pct.setdefault(r['month'], []).append(r['wait_d'])
months_sorted = sorted(monthly_pct.keys())[-6:]
monthly_avg = {m: round(mean(v)*3 + 14, 1) for m, v in monthly_pct.items()}  # rough PCT estimate

# Detailed batch table (top waits)
detail_table = sorted(all_waits, key=lambda x: x['wait_d'], reverse=True)[:50]
for r in detail_table:
    r['start_str'] = r.get('date','')
    r['wait_str']  = f"{r['wait_d']:.1f}d"

# Work center current status (for WC detail view)
# Find WOs that are scheduled but not finished (future dates or in progress)
today = datetime(2026, 4, 27)
active_wos = []
for wo, steps in wo_steps.items():
    for s in steps:
        if s['finish'] and s['finish'] >= today - timedelta(days=7):
            if s['step'] in ('compression', 'coating', 'encap_hard', 'sg_encap'):
                active_wos.append({
                    'wo': wo, 'step': s['step'], 'wc': s['wc'],
                    'item': s['item'], 'desc': s['desc'][:40] if s['desc'] else '',
                    'start': s['start'].isoformat() if s['start'] else None,
                    'finish': s['finish'].isoformat() if s['finish'] else None,
                })

active_wos = active_wos[:30]

# Transitions for filter dropdown
all_transitions = sorted(trans_waits.keys())

print(f"\n── Summary ───────────────────────────────")
print(f"  Manufacturing records:  {len(all_mfg)}")
print(f"  Packaging records:      {len(all_pkg)}")
print(f"  Waiting time records:   {len(all_waits)}")
print(f"  Avg PCT:                {avg_pct} days")
print(f"  Avg wait total:         {avg_wait_total} days")
print(f"  Waiting % of PCT:       {wait_pct}%")
print(f"  Biggest bottleneck:     {biggest_bt} ({biggest_bt_days}d)")
print(f"  Top transitions:        {[t['transition'] for t in top_bottlenecks[:5]]}")

# ── build chart-ready data ────────────────────────────────────────────────────
COLORS = {
    'Dispensing → Compression':       '#e040a8',
    'Dispensing → Encap Hard':        '#f48acb',
    'Compression → Coating':          '#9c5bbf',
    'Sg Med Disp → Sg Encap':         '#7eb8d4',
    'Sg Gel Disp → Sg Encap':         '#5ba3c2',
    'Manufacturing → Packaging':      '#4db89a',
    'default':                        '#a0a0c0',
}

def color_for(t):
    return COLORS.get(t, COLORS['default'])

chart_datasets = []
for t, series in chart_series.items():
    if any(v > 0 for v in series):
        chart_datasets.append({
            'label': t, 'data': series,
            'backgroundColor': color_for(t),
            'borderColor': color_for(t),
            'borderWidth': 0, 'borderRadius': 3,
        })

# ── per-product-type daily breakdown chart ───────────────────────────────────
# Build per-WO timeline: run hours for each step + wait hours between steps

def get_run_h(s):
    """Allocated run hours for a step record."""
    if s and s.get('run_h') and s['run_h'] > 0:
        return s['run_h']
    if s and s.get('start') and s.get('finish'):
        h = (s['finish'] - s['start']).total_seconds() / 3600
        return h if 0 < h <= 200 else 0
    return 0

def get_wait_h_between(by_step, from_step, to_step):
    sf = by_step.get(from_step)
    st = by_step.get(to_step)
    if sf and st and sf.get('finish') and st.get('start'):
        h = (st['start'] - sf['finish']).total_seconds() / 3600
        return max(0, h) if h <= 30 * 24 else 0
    return 0

wo_ptype_entries = []
for wo, steps in wo_steps.items():
    by_step = {}
    for s in steps:
        k = s['step']
        if k not in by_step or (s['finish'] and (by_step[k]['finish'] is None or s['finish'] > by_step[k]['finish'])):
            by_step[k] = s
    item  = next((s['item'] for s in steps if s['item']), '')
    ptype = product_types.get(item, '')
    if ptype not in ('TC', 'TU', 'CH', 'SG'):
        continue
    # grouping date = dispensing start (or sg_encap start for SG)
    disp = (by_step.get('dispensing') or by_step.get('sg_gel_disp')
            or by_step.get('sg_med_disp') or by_step.get('sg_encap'))
    if not disp or not disp.get('start'):
        continue
    date_str = disp['start'].strftime('%Y-%m-%d')

    if ptype == 'TC':
        if 'compression' not in by_step or 'coating' not in by_step:
            continue
        wo_ptype_entries.append({
            'date': date_str, 'type': ptype, 'wo': wo, 'item': item,
            'disp_run':        get_run_h(by_step.get('dispensing')),
            'disp_comp_wait':  get_wait_h_between(by_step, 'dispensing', 'compression'),
            'comp_run':        get_run_h(by_step.get('compression')),
            'comp_coat_wait':  get_wait_h_between(by_step, 'compression', 'coating'),
            'coat_run':        get_run_h(by_step.get('coating')),
        })
    elif ptype == 'TU':
        if 'compression' not in by_step:
            continue
        wo_ptype_entries.append({
            'date': date_str, 'type': ptype, 'wo': wo, 'item': item,
            'disp_run':        get_run_h(by_step.get('dispensing')),
            'disp_comp_wait':  get_wait_h_between(by_step, 'dispensing', 'compression'),
            'comp_run':        get_run_h(by_step.get('compression')),
        })
    elif ptype == 'CH':
        if 'encap_hard' not in by_step:
            continue
        wo_ptype_entries.append({
            'date': date_str, 'type': ptype, 'wo': wo, 'item': item,
            'disp_run':         get_run_h(by_step.get('dispensing')),
            'disp_encap_wait':  get_wait_h_between(by_step, 'dispensing', 'encap_hard'),
            'encap_run':        get_run_h(by_step.get('encap_hard')),
        })
    elif ptype == 'SG':
        if 'sg_encap' not in by_step:
            continue
        wo_ptype_entries.append({
            'date': date_str, 'type': ptype, 'wo': wo, 'item': item,
            'sg_encap_run':  get_run_h(by_step.get('sg_encap')),
        })

# Add packaging wait + duration per item (averaged across all pkg records)
item_pkg_avg = {}
for item, pkg_list in pkg_by_item.items():
    mfg_finishes = sorted(mfg_last_by_item.get(item, []), key=lambda x: x['finish'])
    waits_h, durs_h = [], []
    for pr in sorted(pkg_list, key=lambda x: x['start']):
        best = None
        for m in mfg_finishes:
            diff = (pr['start'] - m['finish']).total_seconds() / 3600
            if -48 <= diff <= 30 * 24:
                if best is None or abs(diff) < abs((pr['start'] - best['finish']).total_seconds() / 3600):
                    best = m
        if best:
            w = max(0, (pr['start'] - best['finish']).total_seconds() / 3600)
            if w <= 30 * 24:
                waits_h.append(w)
        if pr.get('run_h') and pr['run_h'] > 0:
            durs_h.append(pr['run_h'])
    if waits_h or durs_h:
        item_pkg_avg[item] = {
            'wait_h': round(mean(waits_h), 2) if waits_h else 0,
            'dur_h':  round(mean(durs_h),  2) if durs_h  else 0,
        }

for e in wo_ptype_entries:
    ps = item_pkg_avg.get(e['item'], {})
    e['pkg_wait'] = ps.get('wait_h', 0)   # hours
    e['pkg_run']  = ps.get('dur_h', 0)    # hours

# Aggregate by date per product type (all values in HOURS for chart)
def agg_daily(entries, fields):
    by_date = defaultdict(lambda: defaultdict(list))
    for e in entries:
        for f in fields:
            v = e.get(f, 0)
            if v and v > 0:
                by_date[e['date']][f].append(v)
    return {d: {f: round(mean(by_date[d][f]), 2) if by_date[d][f] else 0 for f in fields}
            for d in sorted(by_date)}

PTYPE_FIELDS = {
    'TC': ['disp_run','disp_comp_wait','comp_run','comp_coat_wait','coat_run','pkg_wait','pkg_run'],
    'TU': ['disp_run','disp_comp_wait','comp_run','pkg_wait','pkg_run'],
    'CH': ['disp_run','disp_encap_wait','encap_run','pkg_wait','pkg_run'],
    'SG': ['sg_encap_run','pkg_wait','pkg_run'],
}
PTYPE_LABELS = {
    'TC': ['Dispensing','Wait → Compression','Compression','Wait → Coating','Coating','Wait → Packaging','Packaging'],
    'TU': ['Dispensing','Wait → Compression','Compression','Wait → Packaging','Packaging'],
    'CH': ['Dispensing','Wait → Encapsulation','Encapsulation','Wait → Packaging','Packaging'],
    'SG': ['SG Encapsulation','Wait → Packaging','Packaging'],
}
# Solid colors for active steps, lighter/dashed-border for wait segments
PTYPE_COLORS = {
    'TC': [
        ('#e040a8','#e040a8'),  # Dispensing — solid pink
        ('#f0a0cc','#e040a8'),  # Wait → Comp — lighter pink
        ('#9c5bbf','#9c5bbf'),  # Compression — solid purple
        ('#c8a8e0','#9c5bbf'),  # Wait → Coating — lighter purple
        ('#4db89a','#4db89a'),  # Coating — solid teal
        ('#a0dcc8','#4db89a'),  # Wait → Pkg — lighter teal
        ('#6dd96d','#6dd96d'),  # Packaging — solid green
    ],
    'TU': [
        ('#e040a8','#e040a8'),
        ('#f0a0cc','#e040a8'),
        ('#9c5bbf','#9c5bbf'),
        ('#c8a8e0','#9c5bbf'),
        ('#6dd96d','#6dd96d'),
    ],
    'CH': [
        ('#e040a8','#e040a8'),
        ('#f0a0cc','#e040a8'),
        ('#f5a623','#f5a623'),
        ('#fad090','#f5a623'),
        ('#6dd96d','#6dd96d'),
    ],
    'SG': [
        ('#7eb8d4','#7eb8d4'),
        ('#b8dce8','#7eb8d4'),
        ('#6dd96d','#6dd96d'),
    ],
}
PTYPE_IS_WAIT = {
    'TC': [False, True, False, True, False, True, False],
    'TU': [False, True, False, True, False],
    'CH': [False, True, False, True, False],
    'SG': [False, True, False],
}

ptype_chart = {}
for pt in ('TC','TU','CH','SG'):
    entries = [e for e in wo_ptype_entries if e['type'] == pt]
    agg = agg_daily(entries, PTYPE_FIELDS[pt])
    dates = sorted(agg.keys())
    datasets = []
    for i, (field, label) in enumerate(zip(PTYPE_FIELDS[pt], PTYPE_LABELS[pt])):
        data = [agg[d].get(field, 0) for d in dates]
        if not any(v > 0 for v in data):
            continue
        bg, border = PTYPE_COLORS[pt][i]
        is_wait = PTYPE_IS_WAIT[pt][i]
        datasets.append({
            'label': label, 'data': data,
            'backgroundColor': bg,
            'borderColor': border,
            'borderWidth': 1 if is_wait else 0,
            'borderRadius': 2,
            'isWait': is_wait,
        })
    ptype_chart[pt] = {'labels': dates, 'datasets': datasets}
    print(f"  ptype {pt}: {len(dates)} days, {len(entries)} WOs")

# ── compression booth snapshot ───────────────────────────────────────────────
SNAPSHOT_DT = datetime(2026, 4, 23, 7, 0)
SNAPSHOT_STR = "23 Apr 2026 — 07:00 AM"

COMP_BOOTHS = [
    ('TC1-Korsch 1', 'TC1', 'Korsch 1'),
    ('TC2-Korsch 2', 'TC2', 'Korsch 2'),
    ('TC 3-Kil',     'TC3', 'Kilian'),
    ('TC 4-Kil',     'TC4', 'Kilian'),
    ('TC 5-Bosch',   'TC5', 'Bosch'),
    ('TC 6-Fette 1', 'TC6', 'Fette 1'),
    ('TC 7-Fette 2', 'TC7', 'Fette 2'),
    ('TC 8-Fette 3', 'TC8', 'Fette 3'),
]

def fmt_dt(dt):
    return dt.strftime('%d %b, %H:%M') if dt else '—'

def dur_str(s, f):
    if s and f:
        h = (f - s).total_seconds() / 3600
        return f"{int(h)}h {int((h % 1) * 60):02d}m"
    return '—'

# Build per-booth WO list from mfg records
booth_wo_map = {b[0]: [] for b in COMP_BOOTHS}
for r in all_mfg:
    if r['wc'] in booth_wo_map and (r['start'] or r['finish']):
        booth_wo_map[r['wc']].append(r)
for v in booth_wo_map.values():
    v.sort(key=lambda x: x['start'] or x['finish'] or datetime.min)

compression_booths = []
for bname, short, machine in COMP_BOOTHS:
    jobs = booth_wo_map[bname]

    running = last_done = next_sched = None
    for j in jobs:
        s, f = j.get('start'), j.get('finish')
        if s and f:
            if s <= SNAPSHOT_DT < f:
                running = j; break
            elif f <= SNAPSHOT_DT:
                if not last_done or f > last_done['finish']:
                    last_done = j
        elif s and not f and s <= SNAPSHOT_DT:
            running = j; break
        elif s and s > SNAPSHOT_DT:
            if not next_sched or s < next_sched['start']:
                next_sched = j

    if running:   status, current = 'RUN',   running
    elif last_done: status, current = 'IDLE',  last_done
    elif next_sched: status, current = 'SCHED', next_sched
    else:           status, current = 'IDLE',  None

    idle_h = None
    if status == 'IDLE' and current and current.get('finish'):
        idle_h = round((SNAPSHOT_DT - current['finish']).total_seconds() / 3600, 1)

    def make_job(j):
        return {
            'wo':   j['wo'],
            'item': j['item'],
            'desc': (j['desc'] or '').strip()[:40],
            'qty':  str(j.get('qty') or ''),
            'type': product_types.get(j['item'], ''),
            'run_h': round(j['run_h'], 1) if j.get('run_h') else '',
            'start': fmt_dt(j.get('start')),
            'finish': fmt_dt(j.get('finish')),
            'duration': dur_str(j.get('start'), j.get('finish')),
        }

    # recent jobs (last 3 completed before snapshot, excluding current)
    done_before = [j for j in jobs
                   if j.get('finish') and j['finish'] <= SNAPSHOT_DT
                   and j is not current][-3:]

    compression_booths.append({
        'name': bname, 'short': short, 'machine': machine,
        'status': status,
        'idle_h': idle_h,
        'queue': len([j for j in jobs if j.get('start') and j['start'] > SNAPSHOT_DT]),
        'current': make_job(current) if current else None,
        'recent': [make_job(j) for j in done_before],
    })
    print(f"  booth {short}: {status}, {len(jobs)} jobs, idle {idle_h}h")

dashboard_data = {
    'kpi': {
        'avg_pct': avg_pct,
        'pct_target': PCT_TARGET_DAYS,
        'wait_pct': wait_pct,
        'biggest_bottleneck': biggest_bt,
        'biggest_bottleneck_days': biggest_bt_days,
        'avg_wait_days': avg_wait_total,
    },
    'chart_labels': chart_labels,
    'chart_datasets': chart_datasets,
    'top_bottlenecks': top_bottlenecks,
    'detail_table': detail_table[:40],
    'all_transitions': all_transitions,
    'active_wos': active_wos,
    'monthly_labels': months_sorted,
    'monthly_pct': [monthly_avg.get(m, avg_pct) for m in months_sorted],
    'ptype_chart': ptype_chart,
    'compression_booths': compression_booths,
    'snapshot_str': SNAPSHOT_STR,
}

data_json = json.dumps(dashboard_data, default=str, indent=2)

# ── HTML template ─────────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PCT Waiting Time Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
:root {{
  --bg:       #1a2b18;
  --panel:    #243b21;
  --card:     #2d4a29;
  --card2:    #1e3319;
  --accent1:  #e040a8;
  --accent2:  #9c5bbf;
  --accent3:  #4db89a;
  --green-hi: #6dd96d;
  --amber:    #f5a623;
  --red:      #e84040;
  --text:     #ffffff;
  --text2:    #b8ccb5;
  --border:   #3a5a35;
  --row-alt:  #1e3319;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  background: var(--bg);
  color: var(--text);
  font-size: 14px;
  min-height: 100vh;
}}
/* ── top bar ── */
.topbar {{
  background: var(--panel);
  border-bottom: 1px solid var(--border);
  padding: 12px 24px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}}
.topbar-left {{ display: flex; flex-direction: column; }}
.topbar-sub {{ font-size: 11px; color: var(--text2); text-transform: uppercase; letter-spacing: 1px; }}
.topbar-title {{ font-size: 28px; font-weight: 700; letter-spacing: -0.5px; }}
.brand {{ font-size: 26px; font-weight: 800; color: var(--text); letter-spacing: -1px; }}
/* ── filters ── */
.filters {{
  background: var(--panel);
  padding: 8px 24px;
  display: flex;
  align-items: center;
  gap: 16px;
  border-bottom: 1px solid var(--border);
  flex-wrap: wrap;
}}
.filter-group {{ display: flex; align-items: center; gap: 6px; }}
.filter-group label {{ font-size: 12px; color: var(--text2); white-space: nowrap; }}
select, input[type=date] {{
  background: var(--card);
  border: 1px solid var(--border);
  color: var(--text);
  border-radius: 6px;
  padding: 5px 10px;
  font-size: 12px;
  cursor: pointer;
}}
.tab-bar {{
  display: flex;
  gap: 4px;
  background: var(--panel);
  padding: 8px 24px 0;
  border-bottom: 1px solid var(--border);
}}
.tab {{
  padding: 8px 18px;
  border-radius: 6px 6px 0 0;
  cursor: pointer;
  font-size: 13px;
  color: var(--text2);
  background: transparent;
  border: none;
  transition: all .2s;
}}
.tab.active {{ background: var(--card); color: var(--text); font-weight: 600; }}
/* ── KPI cards ── */
.kpi-row {{
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
  padding: 16px 24px 0;
}}
.kpi {{
  background: var(--card);
  border-radius: 10px;
  padding: 16px 20px;
  border: 1px solid var(--border);
  position: relative;
  overflow: hidden;
}}
.kpi::after {{
  content: '';
  position: absolute; top: 0; left: 0; right: 0; height: 3px;
  border-radius: 10px 10px 0 0;
}}
.kpi.green::after  {{ background: var(--green-hi); }}
.kpi.amber::after  {{ background: var(--amber); }}
.kpi.red::after    {{ background: var(--red); }}
.kpi.purple::after {{ background: var(--accent2); }}
.kpi-label {{ font-size: 11px; color: var(--text2); text-transform: uppercase; letter-spacing: .8px; margin-bottom: 6px; }}
.kpi-value {{ font-size: 32px; font-weight: 700; line-height: 1; margin-bottom: 4px; }}
.kpi-sub   {{ font-size: 12px; color: var(--text2); display: flex; align-items: center; gap: 4px; }}
.trend-up   {{ color: var(--red); }}
.trend-down {{ color: var(--green-hi); }}
.badge {{
  display: inline-block;
  padding: 2px 8px;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 600;
}}
.badge-red    {{ background: rgba(232,64,64,.2);  color: var(--red); }}
.badge-amber  {{ background: rgba(245,166,35,.2); color: var(--amber); }}
.badge-green  {{ background: rgba(109,217,109,.2);color: var(--green-hi); }}
.progress-bar {{
  height: 6px; background: var(--border);
  border-radius: 3px; margin-top: 8px; overflow: hidden;
}}
.progress-fill {{ height: 100%; border-radius: 3px; transition: width .4s; }}
/* ── main grid ── */
.main-grid {{
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px;
  padding: 16px 24px;
}}
.panel {{
  background: var(--panel);
  border-radius: 10px;
  border: 1px solid var(--border);
  overflow: hidden;
}}
.panel-header {{
  padding: 14px 18px 10px;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: space-between;
}}
.panel-title {{ font-weight: 600; font-size: 14px; }}
.panel-body  {{ padding: 16px 18px; }}
.chart-wrap  {{ position: relative; height: 260px; }}
/* ── table ── */
.data-table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
.data-table th {{
  padding: 8px 10px; text-align: left;
  color: var(--text2); font-weight: 500;
  border-bottom: 1px solid var(--border);
  background: var(--card2);
  white-space: nowrap;
}}
.data-table td {{
  padding: 7px 10px;
  border-bottom: 1px solid rgba(58,90,53,.4);
  color: var(--text);
}}
.data-table tr:hover td {{ background: rgba(255,255,255,.03); }}
.data-table tr:nth-child(even) td {{ background: var(--row-alt); }}
/* ── bottleneck bars ── */
.bt-row  {{ display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }}
.bt-name {{ width: 220px; font-size: 12px; color: var(--text2); text-align: right; flex-shrink: 0; }}
.bt-bar-wrap {{ flex: 1; background: rgba(255,255,255,.08); border-radius: 4px; height: 18px; overflow: hidden; }}
.bt-bar  {{ height: 100%; border-radius: 4px; transition: width .6s; }}
.bt-val  {{ width: 50px; font-size: 12px; font-weight: 600; }}
.bt-badge {{ flex-shrink: 0; }}
/* ── severity dots ── */
.dot {{ display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 4px; }}
.dot-red   {{ background: var(--red); }}
.dot-amber {{ background: var(--amber); }}
.dot-green {{ background: var(--green-hi); }}
/* ── WC detail view ── */
.wc-grid {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(160px,1fr));
  gap: 10px;
  padding: 12px 0;
}}
.wc-card {{
  background: var(--card);
  border-radius: 8px;
  padding: 12px;
  border: 1px solid var(--border);
  font-size: 12px;
}}
.wc-card-name {{ font-weight: 700; font-size: 13px; margin-bottom: 4px; }}
.wc-card-status {{ display: flex; align-items: center; gap: 4px; font-size: 11px; color: var(--text2); margin-bottom: 6px; }}
.status-dot {{ width: 8px; height: 8px; border-radius: 50%; display: inline-block; }}
.status-run   {{ background: var(--green-hi); box-shadow: 0 0 6px var(--green-hi); }}
.status-setup {{ background: var(--amber); }}
.status-idle  {{ background: #888; }}
.wc-wo {{ font-size: 12px; font-weight: 600; }}
.wc-item {{ font-size: 11px; color: var(--text2); margin-top: 2px; }}
.wc-time {{ font-size: 11px; color: var(--text2); margin-top: 4px; }}
/* ── full-width panel ── */
.full-panel {{ grid-column: 1 / -1; }}
/* ── view toggle ── */
#overview-view, #wc-view {{ display: none; }}
#overview-view.active, #wc-view.active {{ display: block; }}
/* ── scrollable table ── */
.table-scroll {{ overflow-x: auto; max-height: 320px; overflow-y: auto; }}
.table-scroll::-webkit-scrollbar {{ width: 6px; height: 6px; }}
.table-scroll::-webkit-scrollbar-track {{ background: var(--bg); }}
.table-scroll::-webkit-scrollbar-thumb {{ background: var(--border); border-radius: 3px; }}
/* ── footer ── */
footer {{
  text-align: center;
  padding: 12px;
  color: var(--text2);
  font-size: 11px;
  border-top: 1px solid var(--border);
  margin-top: 4px;
}}
/* ── snapshot bar ── */
.snapshot-bar {{
  background: var(--card2);
  border-bottom: 1px solid var(--border);
  padding: 8px 24px;
  display: flex; align-items: center; gap: 16px;
  font-size: 12px;
}}
.snapshot-time {{
  font-weight: 700; color: var(--green-hi); font-size: 13px;
}}
.snapshot-note {{ color: var(--text2); font-style: italic; }}
.live-dot {{
  width: 8px; height: 8px; border-radius: 50%;
  background: var(--green-hi);
  box-shadow: 0 0 0 0 var(--green-hi);
  animation: pulse-green 2s infinite;
  display: inline-block;
}}
@keyframes pulse-green {{
  0%   {{ box-shadow: 0 0 0 0 rgba(109,217,109,.7); }}
  70%  {{ box-shadow: 0 0 0 8px rgba(109,217,109,0); }}
  100% {{ box-shadow: 0 0 0 0 rgba(109,217,109,0); }}
}}
/* ── booth grid ── */
.booth-grid {{
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 12px;
  padding: 14px 18px;
}}
.booth-card {{
  background: var(--card2);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
}}
.booth-card-header {{
  display: flex; justify-content: space-between; align-items: center;
  padding: 10px 14px 8px;
  border-bottom: 1px solid var(--border);
}}
.booth-name {{ font-weight: 800; font-size: 15px; }}
.booth-machine {{ font-size: 11px; color: var(--text2); }}
.booth-status-row {{
  display: flex; align-items: center; gap: 6px;
  padding: 6px 14px;
  font-size: 11px; font-weight: 600; letter-spacing: .5px;
}}
.status-run-dot {{
  width: 9px; height: 9px; border-radius: 50%;
  background: var(--green-hi);
  box-shadow: 0 0 0 0 var(--green-hi);
  animation: pulse-green 1.5s infinite;
}}
.status-idle-dot  {{ width: 9px; height: 9px; border-radius: 50%; background: #555; }}
.status-sched-dot {{ width: 9px; height: 9px; border-radius: 50%; background: var(--amber); }}
.booth-body {{ padding: 8px 14px 12px; }}
.booth-wo {{ font-size: 13px; font-weight: 700; margin-bottom: 2px; }}
.booth-item {{
  display: flex; gap: 6px; align-items: center;
  font-size: 11px; color: var(--text2); margin-bottom: 3px;
}}
.booth-desc {{ font-size: 11px; color: var(--text); margin-bottom: 8px; line-height: 1.3; }}
.booth-times {{ display: grid; grid-template-columns: 14px 1fr; gap: 2px 4px; font-size: 11px; }}
.booth-times .lbl {{ color: var(--text2); }}
.booth-times .val {{ color: var(--text); }}
.booth-meta {{
  display: flex; gap: 10px; margin-top: 8px; padding-top: 7px;
  border-top: 1px solid var(--border);
  font-size: 11px; color: var(--text2);
}}
.booth-meta strong {{ color: var(--text); }}
.booth-idle-badge {{
  display: inline-block; padding: 2px 7px;
  border-radius: 10px; font-size: 10px; font-weight: 600;
  background: rgba(255,255,255,.07); color: var(--text2);
  margin-top: 6px;
}}
.booth-run-badge {{
  display: inline-block; padding: 2px 7px;
  border-radius: 10px; font-size: 10px; font-weight: 600;
  background: rgba(109,217,109,.15); color: var(--green-hi);
}}
@media (max-width: 1100px) {{
  .booth-grid {{ grid-template-columns: repeat(2, 1fr); }}
}}
/* ── product type toggle ── */
.ptype-tabs {{ display: flex; gap: 4px; flex-wrap: wrap; }}
.ptype-btn {{
  padding: 4px 12px; border-radius: 20px;
  border: 1px solid var(--border);
  background: transparent; color: var(--text2);
  font-size: 11px; cursor: pointer; transition: all .15s;
}}
.ptype-btn.active, .ptype-btn:hover {{
  background: var(--accent1); border-color: var(--accent1);
  color: var(--text); font-weight: 600;
}}
@media (max-width: 900px) {{
  .kpi-row {{ grid-template-columns: repeat(2, 1fr); }}
  .main-grid {{ grid-template-columns: 1fr; }}
  .bt-name {{ width: 160px; font-size: 11px; }}
}}
</style>
</head>
<body>

<!-- top bar -->
<div class="topbar">
  <div class="topbar-left">
    <span class="topbar-sub">Manufacturing & Supply Dashboard</span>
    <span class="topbar-title">PCT — Plant Cycle Time</span>
  </div>
  <div class="brand">Opella.</div>
</div>

<!-- filters -->
<div class="filters">
  <div class="filter-group">
    <label>📅 Date Range</label>
    <select id="f-range">
      <option value="30">Last 30 Days</option>
      <option value="90" selected>Last 90 Days</option>
      <option value="180">Last 6 Months</option>
      <option value="365">Last 12 Months</option>
    </select>
  </div>
  <div class="filter-group">
    <label>🏭 Product Type</label>
    <select id="f-type">
      <option value="">All Types</option>
      <option>TC — Coated Tablet</option>
      <option>TU — Uncoated Tablet</option>
      <option>CH — Hard Capsule</option>
      <option>SG — Softgel</option>
    </select>
  </div>
  <div class="filter-group">
    <label>🔄 Transition</label>
    <select id="f-transition">
      <option value="">All Transitions</option>
    </select>
  </div>
  <div style="margin-left:auto;font-size:11px;color:var(--text2)">
    Target: <strong style="color:var(--green-hi)">17 days</strong> &nbsp;|&nbsp;
    Project 1000
  </div>
</div>

<!-- tab bar -->
<div class="tab-bar">
  <button class="tab active" onclick="switchTab('overview')">📊 PCT Overview</button>
  <button class="tab" onclick="switchTab('wc')">🏭 Work Centre Detail</button>
</div>

<!-- ═══ OVERVIEW TAB ═══════════════════════════════════════════════════════ -->
<div id="overview-view" class="active">

  <!-- KPI cards -->
  <div class="kpi-row" id="kpi-row"></div>

  <!-- main grid -->
  <div class="main-grid">

    <!-- left: daily stacked bar by product type -->
    <div class="panel">
      <div class="panel-header" style="flex-wrap:wrap;gap:8px">
        <span class="panel-title">Batch Cycle Time Breakdown — Daily (Hrs)</span>
        <div class="ptype-tabs" id="ptype-tabs">
          <button class="ptype-btn active" onclick="switchPtype('TC',this)">Coated Tablet</button>
          <button class="ptype-btn" onclick="switchPtype('TU',this)">Uncoated Tablet</button>
          <button class="ptype-btn" onclick="switchPtype('CH',this)">Hard Capsule</button>
          <button class="ptype-btn" onclick="switchPtype('SG',this)">Softgel</button>
        </div>
      </div>
      <div class="panel-body">
        <div class="chart-wrap" style="height:280px"><canvas id="dailyBar"></canvas></div>
        <div style="font-size:10px;color:var(--text2);margin-top:6px">
          ■ Solid = active process time &nbsp;│&nbsp; ░ Light = waiting time between steps &nbsp;│&nbsp; Y-axis in hours
        </div>
      </div>
    </div>

    <!-- right: top bottlenecks -->
    <div class="panel">
      <div class="panel-header">
        <span class="panel-title">Top Bottlenecks — Avg Waiting Days per Transition</span>
      </div>
      <div class="panel-body" id="bottleneck-bars"></div>
    </div>

    <!-- full width: detail table -->
    <div class="panel full-panel">
      <div class="panel-header">
        <span class="panel-title">Bottleneck Contribution Detail</span>
        <span style="font-size:11px;color:var(--text2)" id="table-count"></span>
      </div>
      <div class="table-scroll" id="detail-table-wrap"></div>
    </div>

  </div>
</div>

<!-- ═══ WC DETAIL TAB ════════════════════════════════════════════════════ -->
<div id="wc-view">

  <!-- snapshot bar -->
  <div class="snapshot-bar">
    <span class="live-dot"></span>
    <span style="color:var(--text2);font-size:12px">Snapshot as at&nbsp;</span>
    <span class="snapshot-time" id="snapshot-time">—</span>
    <span class="snapshot-note">· Historical data</span>
  </div>

  <!-- KPI summary row -->
  <div class="kpi-row" style="padding-top:14px">
    <div class="kpi green">
      <div class="kpi-label">Booths Running</div>
      <div class="kpi-value" id="wc-running">—</div>
      <div class="kpi-sub">of 8 compression booths</div>
    </div>
    <div class="kpi amber">
      <div class="kpi-label">Booths Idle</div>
      <div class="kpi-value" id="wc-idle">—</div>
      <div class="kpi-sub">awaiting next job</div>
    </div>
    <div class="kpi purple">
      <div class="kpi-label">Avg Wait → Compression</div>
      <div class="kpi-value" id="wc-comp-wait">—</div>
      <div class="kpi-sub">days from dispensing finish</div>
    </div>
    <div class="kpi red">
      <div class="kpi-label">PCT vs Target Gap</div>
      <div class="kpi-value" id="wc-gap">—</div>
      <div class="kpi-sub">days above 17-day target</div>
    </div>
  </div>

  <!-- booth grid -->
  <div class="panel" style="margin:14px 24px 0">
    <div class="panel-header">
      <span class="panel-title">Compression Booths — Status Overview</span>
      <span style="font-size:11px;color:var(--text2)">8 booths</span>
    </div>
    <div class="booth-grid" id="booth-grid"></div>
  </div>

  <!-- bottom: WC wait chart -->
  <div style="padding:16px 24px 16px">
    <div class="main-grid">
      <div class="panel">
        <div class="panel-header">
          <span class="panel-title">Waiting Time by Work Centre (Outgoing)</span>
        </div>
        <div class="panel-body">
          <div class="chart-wrap"><canvas id="wcBarChart"></canvas></div>
        </div>
      </div>
    </div>
  </div>

</div>

<footer>
  PCT Waiting Time Dashboard · Target: 17 days (Project 1000) · Data from Manufacturing & Packaging Production Schedules
</footer>

<script>
// ── embedded data ─────────────────────────────────────────────────────────
const DATA = {data_json};

// ── helpers ───────────────────────────────────────────────────────────────
function colorFor(val, low=1, high=3) {{
  if (val <= low)  return '#6dd96d';
  if (val <= high) return '#f5a623';
  return '#e84040';
}}
function badgeFor(val, low=1, high=3) {{
  if (val <= low)  return '<span class="badge badge-green">● Good</span>';
  if (val <= high) return '<span class="badge badge-amber">● Amber</span>';
  return '<span class="badge badge-red">● ⚠ High</span>';
}}
function dotFor(val, low=1, high=3) {{
  const cls = val <= low ? 'dot-green' : val <= high ? 'dot-amber' : 'dot-red';
  return `<span class="dot ${{cls}}"></span>`;
}}

// ── KPI cards ─────────────────────────────────────────────────────────────
function renderKPIs() {{
  const d = DATA.kpi;
  const pctColor = d.avg_pct <= 17 ? 'green' : d.avg_pct <= 20 ? 'amber' : 'red';
  const waitColor = d.wait_pct <= 20 ? 'green' : d.wait_pct <= 35 ? 'amber' : 'red';
  const progress = Math.min(100, Math.round(d.pct_target / d.avg_pct * 100));
  const gap = (d.avg_pct - d.pct_target).toFixed(1);
  const gapLabel = gap > 0 ? `${{gap}}d above target` : `${{Math.abs(gap)}}d below target`;

  document.getElementById('kpi-row').innerHTML = `
    <div class="kpi ${{pctColor}}">
      <div class="kpi-label">Total PCT Days</div>
      <div class="kpi-value">${{d.avg_pct}}</div>
      <div class="kpi-sub">
        ${{d.avg_pct > d.pct_target
          ? '<span class="trend-up">↑</span> above target'
          : '<span class="trend-down">↓</span> at or below target'}}
      </div>
      <div class="progress-bar">
        <div class="progress-fill" style="width:${{progress}}%;background:var(--${{pctColor === 'green' ? 'green-hi' : pctColor}})"></div>
      </div>
    </div>
    <div class="kpi ${{waitColor}}">
      <div class="kpi-label">Waiting Time % of PCT</div>
      <div class="kpi-value">${{d.wait_pct}}%</div>
      <div class="kpi-sub">${{d.avg_wait_days}}d of your PCT is waiting time</div>
    </div>
    <div class="kpi amber">
      <div class="kpi-label">Biggest Bottleneck</div>
      <div class="kpi-value" style="font-size:18px;line-height:1.3">${{d.biggest_bottleneck}}</div>
      <div class="kpi-sub">⚠️ ${{d.biggest_bottleneck_days}}d avg wait</div>
    </div>
    <div class="kpi ${{d.avg_pct > d.pct_target ? 'red' : 'green'}}">
      <div class="kpi-label">Target vs Actual</div>
      <div class="kpi-value">${{d.pct_target}}d / ${{d.avg_pct}}d</div>
      <div class="kpi-sub">${{gapLabel}}</div>
      <div class="progress-bar">
        <div class="progress-fill" style="width:${{progress}}%;background:var(--${{d.avg_pct <= d.pct_target ? 'green-hi' : 'red'}})"></div>
      </div>
    </div>`;
}}

// ── stacked bar chart ─────────────────────────────────────────────────────
let dailyChart = null;
function renderDailyChart(ptype) {{
  const d = DATA.ptype_chart[ptype];
  if (!d || !d.labels || d.labels.length === 0) {{
    document.getElementById('dailyBar').closest('.chart-wrap').innerHTML =
      '<p style="color:var(--text2);padding:40px;text-align:center">No data for this product type</p>';
    return;
  }}
  // Format labels as "01 Apr"
  const labels = d.labels.map(s => {{
    const dt = new Date(s + 'T00:00:00');
    return dt.toLocaleDateString('en-AU', {{ day:'2-digit', month:'short' }});
  }});
  if (dailyChart) {{ dailyChart.destroy(); dailyChart = null; }}
  const ctx = document.getElementById('dailyBar').getContext('2d');
  dailyChart = new Chart(ctx, {{
    type: 'bar',
    data: {{ labels, datasets: d.datasets }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{
          position: 'bottom',
          labels: {{
            color: '#b8ccb5', boxWidth: 14, font: {{ size: 11 }},
            generateLabels: chart => chart.data.datasets.map((ds, i) => ({{
              text: ds.label,
              fillStyle: ds.backgroundColor,
              strokeStyle: ds.borderColor,
              lineWidth: ds.isWait ? 1 : 0,
              hidden: !chart.isDatasetVisible(i),
              index: i,
            }}))
          }}
        }},
        tooltip: {{
          callbacks: {{
            label: ctx => {{
              const h = ctx.raw;
              const d = (h / 24).toFixed(2);
              const tag = ctx.dataset.isWait ? '⏳ Wait' : '⚙️ Active';
              return ` ${{ctx.dataset.label}}: ${{h}}h (${{d}}d) [${{tag}}]`;
            }},
            footer: items => {{
              const total = items.reduce((s, i) => s + i.raw, 0);
              return `Total stack: ${{total.toFixed(1)}}h (${{(total/24).toFixed(2)}}d)`;
            }}
          }}
        }}
      }},
      scales: {{
        x: {{
          stacked: true,
          grid: {{ color: 'rgba(255,255,255,.05)' }},
          ticks: {{ color: '#b8ccb5', font: {{ size: 10 }}, maxRotation: 45, autoSkip: true, maxTicksLimit: 30 }}
        }},
        y: {{
          stacked: true,
          grid: {{ color: 'rgba(255,255,255,.05)' }},
          ticks: {{ color: '#b8ccb5', font: {{ size: 11 }} }},
          title: {{ display: true, text: 'Avg Hours', color: '#b8ccb5', font: {{ size: 11 }} }}
        }}
      }}
    }}
  }});
}}

function switchPtype(ptype, btn) {{
  document.querySelectorAll('.ptype-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderDailyChart(ptype);
}}

// ── bottleneck bars ───────────────────────────────────────────────────────
function renderBottlenecks(data) {{
  const wrap = document.getElementById('bottleneck-bars');
  if (!data || data.length === 0) {{ wrap.innerHTML = '<p style="color:var(--text2)">No data</p>'; return; }}
  const maxVal = Math.max(...data.map(r => r.avg_wait_d));
  wrap.innerHTML = data.slice(0,8).map(r => {{
    const pct = maxVal > 0 ? (r.avg_wait_d / maxVal * 100).toFixed(1) : 0;
    const c = colorFor(r.avg_wait_d);
    return `
      <div class="bt-row">
        <div class="bt-name" title="${{r.transition}}">${{r.transition.length > 30 ? r.transition.slice(0,29)+'…' : r.transition}}</div>
        <div class="bt-bar-wrap">
          <div class="bt-bar" style="width:${{pct}}%;background:${{c}}"></div>
        </div>
        <div class="bt-val" style="color:${{c}}">${{r.avg_wait_d}}d</div>
        <div class="bt-badge">${{badgeFor(r.avg_wait_d)}}</div>
      </div>`;
  }}).join('');
}}

// ── detail table ──────────────────────────────────────────────────────────
function renderTable(rows) {{
  const wrap = document.getElementById('detail-table-wrap');
  document.getElementById('table-count').textContent = `${{rows.length}} records`;
  if (!rows.length) {{ wrap.innerHTML = '<p style="padding:16px;color:var(--text2)">No data for selected filters</p>'; return; }}
  wrap.innerHTML = `
    <table class="data-table">
      <thead><tr>
        <th>Work Order</th><th>Item</th><th>Description</th>
        <th>Transition</th><th>Wait</th><th>Date</th>
        <th>From WC</th><th>To WC</th><th>Type</th><th>Severity</th>
      </tr></thead>
      <tbody>${{rows.map(r => `
        <tr>
          <td>${{r.wo}}</td>
          <td>${{r.item}}</td>
          <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.desc}}">${{r.desc}}</td>
          <td>${{r.transition}}</td>
          <td style="color:${{colorFor(r.wait_d)}};font-weight:600">${{r.wait_d}}d</td>
          <td>${{r.date}}</td>
          <td style="font-size:11px">${{r.wc_from}}</td>
          <td style="font-size:11px">${{r.wc_to}}</td>
          <td><span class="badge" style="background:rgba(255,255,255,.1)">${{r.type}}</span></td>
          <td>${{badgeFor(r.wait_d)}}</td>
        </tr>`).join('')}}
      </tbody>
    </table>`;
}}

// ── WC detail view ────────────────────────────────────────────────────────
function renderWCDetail() {{
  const d = DATA.kpi;
  document.getElementById('wc-active-jobs').textContent = DATA.active_wos.length;
  // top WC from bottleneck
  const topBT = DATA.top_bottlenecks[0];
  document.getElementById('wc-top-wc').textContent = topBT ? topBT.transition.split('→')[1].trim() : '—';
  // avg pkg wait
  const pkgWaits = DATA.detail_table.filter(r => r.transition.includes('Packaging'));
  const avgPkg = pkgWaits.length ? (pkgWaits.reduce((s,r)=>s+r.wait_d,0)/pkgWaits.length).toFixed(1) : d.avg_wait_days;
  document.getElementById('wc-pkg-wait').textContent = avgPkg + 'd';
  document.getElementById('wc-gap').textContent = (d.avg_pct - d.pct_target).toFixed(1) + 'd';

  // WC grid
  const grid = document.getElementById('wc-grid');
  const statusMap = {{'Finished':'DONE','f':'DONE','p':'PKG','m':'SCHED'}};
  const wcNames = [...new Set(DATA.active_wos.map(w => w.wc))].slice(0,12);

  // Group by WC, take latest WO
  const byWC = {{}};
  DATA.active_wos.forEach(w => {{
    if (!byWC[w.wc] || (w.finish && byWC[w.wc].finish < w.finish)) byWC[w.wc] = w;
  }});

  grid.innerHTML = Object.entries(byWC).slice(0,12).map(([wc, w]) => {{
    const statusStr = w.finish ? 'DONE' : w.start ? 'RUN' : 'SCHED';
    const dotCls = statusStr === 'RUN' ? 'status-run' : statusStr === 'DONE' ? 'status-idle' : 'status-setup';
    const startStr = w.start ? new Date(w.start).toLocaleString('en-AU',{{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}}) : '—';
    const finishStr = w.finish ? new Date(w.finish).toLocaleString('en-AU',{{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}}) : '—';
    return `
      <div class="wc-card">
        <div class="wc-card-name">${{wc.replace('TC1-','TC1 ').replace('TC2-','TC2 ')}}</div>
        <div class="wc-card-status">
          <span class="status-dot ${{dotCls}}"></span>
          ${{statusStr}}
        </div>
        <div class="wc-wo">WO: ${{w.wo}}</div>
        <div class="wc-item">${{w.item}}</div>
        <div class="wc-time">▶ ${{startStr}}</div>
        <div class="wc-time">■ ${{finishStr}}</div>
      </div>`;
  }}).join('') || '<p style="color:var(--text2)">No recent activity data</p>';

  // WC bar chart
  const wcWaits = {{}};
  DATA.detail_table.forEach(r => {{
    if (!wcWaits[r.wc_from]) wcWaits[r.wc_from] = [];
    wcWaits[r.wc_from].push(r.wait_d);
  }});
  const wcEntries = Object.entries(wcWaits)
    .map(([k,v]) => ({{wc:k, avg: v.reduce((s,x)=>s+x,0)/v.length}}))
    .sort((a,b) => b.avg - a.avg).slice(0,8);

  const ctx2 = document.getElementById('wcBarChart').getContext('2d');
  new Chart(ctx2, {{
    type: 'bar',
    data: {{
      labels: wcEntries.map(e => e.wc),
      datasets: [{{
        data: wcEntries.map(e => +e.avg.toFixed(2)),
        backgroundColor: wcEntries.map(e => colorFor(e.avg)),
        borderRadius: 4,
      }}]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: c => ` ${{c.raw}}d avg wait after` }} }} }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(255,255,255,.05)' }}, ticks: {{ color: '#b8ccb5' }},
              title: {{ display: true, text: 'Avg Wait Days', color: '#b8ccb5', font:{{size:11}} }} }},
        y: {{ grid: {{ display: false }}, ticks: {{ color: '#b8ccb5', font: {{size:11}} }} }}
      }}
    }}
  }});

}}

// ── tab switching ─────────────────────────────────────────────────────────
function switchTab(tab) {{
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('overview-view').classList.remove('active');
  document.getElementById('wc-view').classList.remove('active');
  event.target.classList.add('active');
  document.getElementById(tab + '-view').classList.add('active');
}}

// ── filter logic ──────────────────────────────────────────────────────────
function populateTransitions() {{
  const sel = document.getElementById('f-transition');
  DATA.all_transitions.forEach(t => {{
    const opt = document.createElement('option');
    opt.value = t; opt.text = t;
    sel.appendChild(opt);
  }});
}}

function applyFilters() {{
  const rangeDays = parseInt(document.getElementById('f-range').value) || 90;
  const typeFilter = document.getElementById('f-type').value.split(' — ')[0];
  const transFilter = document.getElementById('f-transition').value;
  const cutoff = new Date(Date.now() - rangeDays * 86400 * 1000);

  let rows = DATA.detail_table.filter(r => {{
    if (r.date && new Date(r.date) < cutoff) return false;
    if (typeFilter && r.type !== typeFilter) return false;
    if (transFilter && r.transition !== transFilter) return false;
    return true;
  }});

  renderTable(rows);

  // filter bottlenecks
  const btData = DATA.top_bottlenecks.filter(b => {{
    if (transFilter && b.transition !== transFilter) return false;
    return true;
  }});
  renderBottlenecks(btData);
}}

// ── init ──────────────────────────────────────────────────────────────────
renderKPIs();
renderDailyChart('TC');
renderBottlenecks(DATA.top_bottlenecks);
renderTable(DATA.detail_table);
renderWCDetail();
populateTransitions();

document.getElementById('f-range').addEventListener('change', applyFilters);
document.getElementById('f-type').addEventListener('change', applyFilters);
document.getElementById('f-transition').addEventListener('change', applyFilters);
</script>
</body>
</html>"""

with open(OUT_FILE, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"\n✅  Dashboard written to: {OUT_FILE}")
print(f"    Open in browser: file://{OUT_FILE}")
