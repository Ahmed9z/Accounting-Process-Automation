#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════╗
║          BANK RECONCILIATION TOOL  v1.0              ║
║   Matches bank deposits against Odoo payment records  ║
╚══════════════════════════════════════════════════════╝

REQUIREMENTS (install once):
    pip install flask openpyxl

RUN:
    python3 BankRecon.py

Then open your browser at:  http://localhost:5050
"""

"""
Bank ↔ Odoo Reconciliation Engine
Works for any month - auto-detects sheet pairs by normalising bank/currency names.
"""
import re, datetime, shutil, os
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── constants ─────────────────────────────────────────────────────────────────
AMOUNT_TOL    = 0.01
TYPO_MAX_DIST = 2
TYPO_MIN_LEN  = 6

F_MATCHED    = PatternFill('solid', fgColor='C6EFCE')
F_TENTATIVE  = PatternFill('solid', fgColor='FFEB9C')
F_FLAG       = PatternFill('solid', fgColor='FFCC99')
F_UNMATCHED  = PatternFill('solid', fgColor='FFC7CE')
F_HDR        = PatternFill('solid', fgColor='1F4E78')
F_TOT_MATCH  = PatternFill('solid', fgColor='375623')
F_TOT_UNMATCH= PatternFill('solid', fgColor='C00000')
F_GBLU       = PatternFill('solid', fgColor='D9E1F2')

THIN  = Side(style='thin',   color='B7B7B7')
MED   = Side(style='medium', color='555555')
BRD   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BRD_T = Border(left=THIN, right=THIN, top=MED,  bottom=THIN)

FONT_HDR  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
FONT_BODY = Font(name='Calibri', size=10)
FONT_TOT  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
FONT_BOLD = Font(name='Calibri', bold=True, size=10)

STATUS_FILL  = {'MATCHED': F_MATCHED, 'TENTATIVE': F_TENTATIVE,
                'FLAG': F_FLAG, 'UNMATCHED': F_UNMATCHED}
STATUS_LABEL = {'MATCHED':   'Matched',
                'TENTATIVE': 'Tentative – verify',
                'FLAG':      'Ref match, amount differs',
                'UNMATCHED': 'No match found'}


# ── helpers ───────────────────────────────────────────────────────────────────

def _norm_sheet_key(name):
    """'CIB USD' / 'CIB-USD' / 'CIB,USD' → 'CIBUSD'"""
    return re.sub(r'[^A-Z0-9]', '', name.upper())

def parse_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date):     return v
    s = str(v).strip() if v is not None else ''
    for fmt in ('%d/%m/%Y','%Y-%m-%d','%m/%d/%Y'):
        try: return datetime.datetime.strptime(s, fmt).date()
        except ValueError: pass
    return None

def parse_amount(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(',','').replace(' ','')
    if s in ('','--','- -','-'): return None
    try: return float(s)
    except ValueError: return None

def normalize_tokens(raw):
    if raw is None: return set()
    s = str(raw).strip()
    if not s or s.lower() in ('none','nan','not found','notfound'): return set()
    tokens = set()
    for part in re.split(r'[\s/|,]+', s):
        part = part.strip().strip('\\').upper()
        part = part.split('\\')[0]
        if len(part) < 3: continue
        tokens.add(part)
        if len(part) > 2 and part[0].isdigit() and part[1].isalpha():
            tokens.add(part[1:])
    return tokens

def levenshtein(a, b):
    if a == b: return 0
    la, lb = len(a), len(b)
    if la == 0: return lb
    if lb == 0: return la
    prev = list(range(lb + 1))
    for ca in a:
        cur = [prev[0]+1] + [0]*lb
        for j, cb in enumerate(b, 1):
            cur[j] = min(prev[j]+1, cur[j-1]+1, prev[j-1]+(0 if ca==cb else 1))
        prev = cur
    return prev[lb]


# ── data classes ──────────────────────────────────────────────────────────────

class BankRow:
    __slots__ = ('sheet','excel_row','date','ref_raw','tokens','amount')
    def __init__(self, s, r, d, rr, t, a):
        self.sheet=s; self.excel_row=r; self.date=d
        self.ref_raw=rr; self.tokens=t; self.amount=a

class OdooRow:
    __slots__ = ('sheet','excel_row','date','number','ref_raw','ref_fb_raw','tokens','amount','used')
    def __init__(self, s, r, d, n, rr, rfb, t, a):
        self.sheet=s; self.excel_row=r; self.date=d; self.number=n
        self.ref_raw=rr; self.ref_fb_raw=rfb; self.tokens=t; self.amount=a; self.used=False


# ── auto-detect sheet structure ───────────────────────────────────────────────

def detect_bank_cfg(ws):
    """Return dict(header_row, date_col, ref_col, dep_col) for a bank sheet, or None."""
    for hrow in (1, 2):
        row = [str(v).strip().lower().replace('\t','') if v else ''
               for v in next(ws.iter_rows(min_row=hrow, max_row=hrow, values_only=True))]
        date_col = next((i+1 for i,v in enumerate(row)
                         if ('posting' in v or ('date' in v and 'value' not in v and i < 3))), None)
        if not date_col:
            date_col = next((i+1 for i,v in enumerate(row) if 'transaction date' in v or 'date' in v), 1)
        ref_col  = next((i+1 for i,v in enumerate(row)
                         if 'backoffice' in v or 'reference no' in v
                         or ('reference' in v and 'transaction' not in v and 'bank' not in v)), None)
        if not ref_col:
            ref_col = next((i+1 for i,v in enumerate(row) if 'reference' in v), None)
        dep_col  = next((i+1 for i,v in enumerate(row) if v in ('deposit','credit',' deposit ',' credit ')), None)
        if date_col and dep_col:
            return dict(header_row=hrow, date_col=date_col, ref_col=ref_col, dep_col=dep_col)
    return None

def detect_odoo_cfg(ws):
    """Return dict(header_row, date_col, number_col, ref_col, ref_fb, amount_col) or None."""
    row = [str(v).strip().lower() if v else ''
           for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    date_col   = next((i+1 for i,v in enumerate(row) if 'date' in v), None)
    number_col = next((i+1 for i,v in enumerate(row) if 'number' in v), None)
    ref_col    = next((i+1 for i,v in enumerate(row) if v == 'reference'), None)
    ref_fb     = next((i+1 for i,v in enumerate(row) if 'reference bank' in v or 'bank ref' in v), None)
    amount_col = next((i+1 for i,v in enumerate(row) if 'amount' in v and 'currency' in v), None)
    if not amount_col:
        amount_col = next((i+1 for i,v in enumerate(row) if 'amount' in v), None)
    if date_col and amount_col:
        return dict(header_row=1, date_col=date_col, number_col=number_col or 2,
                    ref_col=ref_col or 4, ref_fb=ref_fb, amount_col=amount_col)
    return None


# ── loaders ───────────────────────────────────────────────────────────────────

def load_bank_sheet(wb, sheet):
    ws  = wb[sheet]
    cfg = detect_bank_cfg(ws)
    if not cfg:
        return [], None
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=cfg['header_row']+1,
                                          max_row=ws.max_row, values_only=True),
                             start=cfg['header_row']+1):
        date   = parse_date(row[cfg['date_col']-1] if cfg['date_col']-1 < len(row) else None)
        amount = parse_amount(row[cfg['dep_col']-1]  if cfg['dep_col']-1  < len(row) else None)
        if date is None or amount is None or amount <= 0:
            continue
        ref_raw = row[cfg['ref_col']-1] if cfg['ref_col'] and cfg['ref_col']-1 < len(row) else None
        rows.append(BankRow(sheet, i, date, ref_raw, normalize_tokens(ref_raw), amount))
    return rows, cfg

def load_odoo_sheet(wb, sheet):
    ws  = wb[sheet]
    cfg = detect_odoo_cfg(ws)
    if not cfg:
        return [], None
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=cfg['header_row']+1,
                                          max_row=ws.max_row, values_only=True),
                             start=cfg['header_row']+1):
        date   = parse_date(row[cfg['date_col']-1]    if cfg['date_col']-1    < len(row) else None)
        amount = parse_amount(row[cfg['amount_col']-1] if cfg['amount_col']-1  < len(row) else None)
        if date is None or amount is None or amount <= 0:
            continue
        number  = row[cfg['number_col']-1] if cfg['number_col']-1 < len(row) else None
        ref_raw = row[cfg['ref_col']-1]    if cfg['ref_col']-1    < len(row) else None
        ref_fb  = row[cfg['ref_fb']-1]     if cfg['ref_fb'] and cfg['ref_fb']-1 < len(row) else None
        toks    = normalize_tokens(ref_raw) | normalize_tokens(ref_fb)
        rows.append(OdooRow(sheet, i, date, number, ref_raw, ref_fb, toks, amount))
    return rows, cfg


# ── pair sheets ───────────────────────────────────────────────────────────────

def pair_sheets(bank_names, odoo_names):
    """
    Match bank sheets to odoo sheets by normalised key (strip punctuation/spaces, uppercase).
    Returns list of (bank_sheet, odoo_sheet) and list of unpaired bank sheets.
    """
    odoo_map = {_norm_sheet_key(n): n for n in odoo_names}
    pairs, no_pair = [], []
    for b in bank_names:
        key = _norm_sheet_key(b)
        # direct match
        if key in odoo_map:
            pairs.append((b, odoo_map[key]))
            continue
        # try stripping trailing 'EGP'/'USD'/'EUR' suffix variations
        found = False
        for ok, on in odoo_map.items():
            if key == ok or key.startswith(ok) or ok.startswith(key):
                # check if the non-common suffix is a currency tag
                pairs.append((b, on)); found = True; break
        if not found:
            no_pair.append(b)
    return pairs, no_pair


# ── matching ──────────────────────────────────────────────────────────────────

def reconcile(bank_path, odoo_path):
    bank_wb = openpyxl.load_workbook(bank_path, read_only=True, data_only=True)
    odoo_wb = openpyxl.load_workbook(odoo_path, read_only=True, data_only=True)

    bank_data, bank_cfgs = {}, {}
    for s in bank_wb.sheetnames:
        rows, cfg = load_bank_sheet(bank_wb, s)
        if cfg:
            bank_data[s]  = rows
            bank_cfgs[s]  = cfg

    odoo_data, odoo_cfgs = {}, {}
    for s in odoo_wb.sheetnames:
        rows, cfg = load_odoo_sheet(odoo_wb, s)
        if cfg:
            odoo_data[s]  = rows
            odoo_cfgs[s]  = cfg

    pairs, no_pair = pair_sheets(list(bank_data.keys()), list(odoo_data.keys()))
    match_results  = {}

    for bsheet, osheet in pairs:
        brows = bank_data[bsheet]
        orows = odoo_data.get(osheet, [])

        tok_idx = defaultdict(list)
        for o in orows:
            for t in o.tokens:
                tok_idx[t].append(o)

        results = {}

        # Stage 1 – exact reference
        stage1_unresolved = []
        for b in brows:
            cands = {id(o): o for t in b.tokens for o in tok_idx.get(t, []) if not o.used}
            if not cands:
                stage1_unresolved.append(b); continue
            cl = list(cands.values())
            amt_ok = [o for o in cl if abs(o.amount - b.amount) <= AMOUNT_TOL]
            if amt_ok:
                chosen = min(amt_ok, key=lambda o: abs((o.date - b.date).days))
                chosen.used = True
                results[id(b)] = dict(brow=b, status='MATCHED', odoo=chosen, note='')
            else:
                chosen = min(cl, key=lambda o: abs(o.amount - b.amount))
                chosen.used = True
                results[id(b)] = dict(brow=b, status='FLAG', odoo=chosen,
                    note=(f'Reference matches Odoo {chosen.number}, but amount differs: '
                          f'bank={b.amount:,.2f} vs odoo={chosen.amount:,.2f}'))
                stage1_unresolved.append(b)

        # Stage 2 – amount + date fallback
        handled = {id(r['brow']) for r in results.values() if r['status'] == 'MATCHED'}
        by_amt  = defaultdict(list)
        for o in orows:
            if not o.used: by_amt[round(o.amount,2)].append(o)
        for b in sorted([b for b in brows if id(b) not in handled], key=lambda b: b.date):
            cands = [o for o in by_amt.get(round(b.amount,2),[]) if not o.used]
            if not cands: continue
            chosen = min(cands, key=lambda o: abs((o.date - b.date).days))
            chosen.used = True
            results[id(b)] = dict(brow=b, status='TENTATIVE', odoo=chosen,
                note=(f'No reference match; matched by identical amount ({b.amount:,.2f}) '
                      f'and nearest date (bank {b.date} vs odoo {chosen.date}). Please verify.'))

        for b in brows:
            if id(b) not in results:
                results[id(b)] = dict(brow=b, status='UNMATCHED', odoo=None, note='')

        match_results[bsheet] = list(results.values())

    # Stage 3 – same-sheet typo hints
    for bsheet, osheet in pairs:
        orphans = [o for o in odoo_data.get(osheet,[]) if not o.used]
        for r in match_results[bsheet]:
            if r['status'] not in ('UNMATCHED','FLAG'): continue
            b    = r['brow']
            best = None
            for bt in b.tokens:
                if len(bt) < TYPO_MIN_LEN: continue
                for o in orphans:
                    for ot in o.tokens:
                        if abs(len(ot)-len(bt)) > TYPO_MAX_DIST: continue
                        d = levenshtein(bt, ot)
                        if 0 < d <= TYPO_MAX_DIST and (best is None or d < best[0]):
                            best = (d, bt, ot, o)
            if best:
                d,bt,ot,o = best
                amt_note = ('amount matches' if abs(o.amount-b.amount)<=AMOUNT_TOL
                            else f'amount differs (bank={b.amount:,.2f} vs odoo={o.amount:,.2f})')
                r['typo'] = dict(dist=d, bank_tok=bt, odoo_tok=ot, odoo_row=o, amount_note=amt_note)

    # Stage 4 – cross-bank hints
    global_orphan = defaultdict(lambda: defaultdict(list))
    for os, orows in odoo_data.items():
        for o in orows:
            if not o.used:
                for t in o.tokens:
                    global_orphan[os][t].append(o)

    for bsheet, osheet in pairs:
        for r in match_results[bsheet]:
            if r['status'] not in ('UNMATCHED','FLAG'): continue
            b = r['brow']
            exact, fuzzy = [], []
            for os, tidx in global_orphan.items():
                if os == osheet: continue
                for bt in b.tokens:
                    for o in tidx.get(bt,[]): exact.append((bt,bt,o))
                    for ot, ol in tidx.items():
                        if abs(len(ot)-len(bt))<=TYPO_MAX_DIST and len(bt)>=TYPO_MIN_LEN:
                            d = levenshtein(bt,ot)
                            if 0 < d <= TYPO_MAX_DIST:
                                for o in ol: fuzzy.append((d,bt,ot,o))
            if exact:
                bt,ot,o = exact[0]
                amt_note = ('amount matches' if abs(o.amount-b.amount)<=AMOUNT_TOL
                            else f'bank={b.amount:,.2f} vs odoo={o.amount:,.2f}')
                r['cross_bank'] = dict(kind='exact',dist=0,bank_tok=bt,odoo_tok=ot,
                                       odoo_row=o,amount_note=amt_note)
            elif fuzzy:
                fuzzy.sort(key=lambda x:x[0])
                d,bt,ot,o = fuzzy[0]
                amt_note = ('amount matches' if abs(o.amount-b.amount)<=AMOUNT_TOL
                            else f'bank={b.amount:,.2f} vs odoo={o.amount:,.2f}')
                r['cross_bank'] = dict(kind='fuzzy',dist=d,bank_tok=bt,odoo_tok=ot,
                                       odoo_row=o,amount_note=amt_note)

    odoo_orphans = {os: [o for o in orows if not o.used]
                    for os, orows in odoo_data.items()}

    return dict(bank_data=bank_data, bank_cfgs=bank_cfgs,
                odoo_data=odoo_data, odoo_cfgs=odoo_cfgs,
                match_results=match_results, odoo_orphans=odoo_orphans,
                pairs=pairs, no_pair=no_pair)


# ── output builder ────────────────────────────────────────────────────────────

def _note(bsheet, r, bank_file_label, odoo_file_label):
    parts = []
    if r.get('note'): parts.append(r['note'])
    o = r.get('odoo')
    if o and r['status'] in ('MATCHED','TENTATIVE','FLAG'):
        parts.append(f"Odoo: {odoo_file_label} > '{o.sheet}' row {o.excel_row} (entry {o.number}).")
    ty = r.get('typo')
    if ty:
        o2 = ty['odoo_row']
        parts.append(f"Possible typo: '{ty['bank_tok']}' ≈ '{ty['odoo_tok']}' "
                     f"(dist {ty['dist']}, {ty['amount_note']}) — "
                     f"'{o2.sheet}' row {o2.excel_row} entry {o2.number}.")
    cb = r.get('cross_bank')
    if cb:
        o2 = cb['odoo_row']
        tag = 'exact' if cb['kind']=='exact' else f"fuzzy dist {cb['dist']}"
        parts.append(f"Wrong-bank candidate ({tag}): '{cb['bank_tok']}' found in "
                     f"'{o2.sheet}' row {o2.excel_row} entry {o2.number} ({cb['amount_note']}).")
    return '  '.join(parts)


def annotate_bank_sheet(wb, bsheet, results, cfg, bank_file_label, odoo_file_label):
    ws   = wb[bsheet]
    hrow = cfg['header_row']
    dep_col = cfg['dep_col']
    ref_col = cfg['ref_col']

    # clear any pre-existing stale XLOOKUP column (col 8 on CIB EGP style sheets)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                            min_col=ws.max_column, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, str) and cell.value in ('notfound','not found'):
                cell.value = None

    last_col = 1
    for c in range(1, ws.max_column + 2):
        if ws.cell(row=hrow, column=c).value not in (None, ''):
            last_col = c

    scol = last_col + 2
    ocol = scol + 1
    ncol = scol + 2
    ws.column_dimensions[get_column_letter(scol)].width = 22
    ws.column_dimensions[get_column_letter(ocol)].width = 22
    ws.column_dimensions[get_column_letter(ncol)].width = 50

    for c, lbl in [(scol,'Match Status'),(ocol,'Matched Odoo Entry'),(ncol,'Reconciliation Notes')]:
        cell = ws.cell(row=hrow, column=c, value=lbl)
        cell.fill=F_HDR; cell.font=FONT_HDR; cell.border=BRD
        cell.alignment=Alignment(horizontal='center', vertical='center', wrap_text=True)

    by_row = {r['brow'].excel_row: r for r in results}
    for erow, r in by_row.items():
        fill  = STATUS_FILL[r['status']]
        label = STATUS_LABEL[r['status']]
        note  = _note(bsheet, r, bank_file_label, odoo_file_label)
        o     = r.get('odoo')
        for c, val in [(scol,label),(ocol,o.number if o else ''),(ncol,note)]:
            cell = ws.cell(row=erow, column=c, value=val)
            cell.fill=fill; cell.font=FONT_BODY; cell.border=BRD
            cell.alignment=Alignment(vertical='top', wrap_text=True)
        for c in ([ref_col] if ref_col else []) + [dep_col]:
            ws.cell(row=erow, column=c).fill = fill

    # totals rows
    matched_rs   = [r for r in results if r['status'] in ('MATCHED','FLAG','TENTATIVE')]
    unmatched_rs = [r for r in results if r['status'] == 'UNMATCHED']
    matched_amt  = sum(r['brow'].amount for r in matched_rs)
    unmatched_amt= sum(r['brow'].amount for r in unmatched_rs)

    last_erow = max((r['brow'].excel_row for r in results), default=hrow)
    for extra in range(last_erow+1, last_erow+5):
        if ws.cell(row=extra, column=1).value not in (None,''):
            last_erow = extra
    t1 = last_erow + 2
    t2 = last_erow + 3

    def write_tot(trow, label, count, amount, fill):
        if dep_col > 1:
            try: ws.merge_cells(start_row=trow,start_column=1,end_row=trow,end_column=dep_col-1)
            except: pass
        lc = ws.cell(row=trow, column=1, value=label)
        lc.fill=fill; lc.font=FONT_TOT; lc.border=BRD_T
        lc.alignment=Alignment(horizontal='right',vertical='center')
        ac = ws.cell(row=trow, column=dep_col, value=amount)
        ac.fill=fill; ac.font=FONT_TOT; ac.border=BRD_T
        ac.number_format='#,##0.00'
        ac.alignment=Alignment(horizontal='right',vertical='center')
        for c in range(dep_col+1, scol):
            cl=ws.cell(row=trow,column=c); cl.fill=fill; cl.border=BRD_T
        cc = ws.cell(row=trow, column=scol, value=f'{count} record(s)')
        cc.fill=fill; cc.font=FONT_TOT; cc.border=BRD_T
        cc.alignment=Alignment(horizontal='center',vertical='center')
        for c in range(scol+1, ncol+1):
            cl=ws.cell(row=trow,column=c); cl.fill=fill; cl.border=BRD_T

    write_tot(t1, f'✔  MATCHED TOTAL ({bsheet})',   len(matched_rs),   matched_amt,   F_TOT_MATCH)
    write_tot(t2, f'✘  UNMATCHED TOTAL ({bsheet})', len(unmatched_rs), unmatched_amt, F_TOT_UNMATCH)
    ws.freeze_panes = ws.cell(row=hrow+1, column=1)


def build_report_sheet(wb, data, bank_file_label, odoo_file_label, month_label):
    if 'Reconciliation Report' in wb.sheetnames:
        del wb['Reconciliation Report']
    ws = wb.create_sheet('Reconciliation Report')
    ws.sheet_view.showGridLines = False

    cur = [1]
    def cr(): return cur[0]
    def nr(n=1): cur[0] += n

    def title(txt):
        ws.merge_cells(start_row=cr(),start_column=1,end_row=cr(),end_column=10)
        c=ws.cell(row=cr(),column=1,value=txt)
        c.font=Font(name='Calibri',bold=True,size=15,color='1F4E78'); nr(2)

    def sub(txt):
        for col in range(1,11): ws.cell(row=cr(),column=col).fill=F_HDR
        ws.merge_cells(start_row=cr(),start_column=1,end_row=cr(),end_column=10)
        c=ws.cell(row=cr(),column=1,value=txt); c.fill=F_HDR; c.font=FONT_HDR; nr(1)

    def note(txt):
        ws.merge_cells(start_row=cr(),start_column=1,end_row=cr(),end_column=10)
        c=ws.cell(row=cr(),column=1,value=txt)
        c.font=Font(name='Calibri',size=9,italic=True,color='595959')
        c.alignment=Alignment(wrap_text=True,vertical='top')
        ws.row_dimensions[cr()].height=26; nr(1)

    def hdr(cols, widths=None):
        for i,h in enumerate(cols,1):
            c=ws.cell(row=cr(),column=i,value=h)
            c.fill=F_HDR; c.font=FONT_HDR; c.border=BRD
            c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if widths:
            for i,w in enumerate(widths,1):
                ws.column_dimensions[get_column_letter(i)].width=w
        nr(1)

    def drow(vals, fill=None, bold=False, ncols=None):
        fnt=Font(name='Calibri',size=10,bold=bold)
        for i,v in enumerate(vals,1):
            c=ws.cell(row=cr(),column=i,value=v)
            c.font=fnt; c.border=BRD
            c.alignment=Alignment(vertical='top',wrap_text=True)
            if fill: c.fill=fill
            if ncols and i in ncols: c.number_format='#,##0.00'
        nr(1)

    mr  = data['match_results']
    od  = data['odoo_data']
    orp = data['odoo_orphans']
    pairs = data['pairs']

    title(f'Bank ↔ Odoo Reconciliation — {month_label}')
    note(f'Bank file: {bank_file_label}   |   Odoo file: {odoo_file_label}   |   '
         f'Matching: exact reference first, amount+date fallback where no reference exists. '
         f'Every row below includes exact cell locations.')
    nr(1)

    # 1a bank totals
    sub('1a.  Bank-Side Totals')
    note('"Matched" = Exact + Flag + Tentative.  "Unmatched" = no counterpart found at all.')
    hdr(['Bank Sheet','Odoo Sheet','Deposits (#)','Deposits (amt)',
         'Matched (#)','Matched (amt)','Unmatched (#)','Unmatched (amt)'],
        widths=[13,12,11,16,11,16,12,16])
    for bs, osheet in pairs:
        rs=mr.get(bs,[])
        tn=len(rs); ta=sum(r['brow'].amount for r in rs)
        mn=sum(1 for r in rs if r['status'] in ('MATCHED','FLAG','TENTATIVE'))
        ma=sum(r['brow'].amount for r in rs if r['status'] in ('MATCHED','FLAG','TENTATIVE'))
        un=sum(1 for r in rs if r['status']=='UNMATCHED')
        ua=sum(r['brow'].amount for r in rs if r['status']=='UNMATCHED')
        drow([bs,osheet,tn,round(ta,2),mn,round(ma,2),un,round(ua,2)],ncols={4,6,8})
    for bs in data['no_pair']:
        rs=mr.get(bs,[])
        tn=len(rs); ta=sum(r['brow'].amount for r in rs)
        drow([bs,'— no Odoo sheet —',tn,round(ta,2),0,'—',tn,round(ta,2)],
             fill=F_UNMATCHED,ncols={4,8})
    gt=sum(len(mr.get(bs,[])) for bs,_ in pairs)+sum(len(mr.get(bs,[])) for bs in data['no_pair'])
    gm=sum(sum(1 for r in mr.get(bs,[]) if r['status'] in ('MATCHED','FLAG','TENTATIVE'))
           for bs,_ in pairs)
    gu=gt-gm
    drow(['ALL SHEETS (row counts only — do not sum $ across currencies)','',
          gt,'—',gm,'—',gu,'—'],fill=F_GBLU,bold=True); nr(1)

    # 1b odoo totals
    sub('1b.  Odoo-Side Totals')
    note('"Orphan" = in Odoo but no bank deposit found. Full list in Section 7.')
    hdr(['Odoo Sheet','Bank Sheet','Entries (#)','Entries (amt)',
         'Matched (#)','Matched (amt)','Orphan (#)','Orphan (amt)'],
        widths=[12,13,11,16,11,16,10,15])
    o2b = {osheet: bs for bs, osheet in pairs}
    for os, orows in od.items():
        tn=len(orows); ta=sum(o.amount for o in orows)
        on=len(orp.get(os,[])); oa=sum(o.amount for o in orp.get(os,[]))
        mn=tn-on; ma=ta-oa
        drow([os,o2b.get(os,'—'),tn,round(ta,2),mn,round(ma,2),on,round(oa,2)],ncols={4,6,8})
    nr(1)

    # 1c breakdown
    sub('1c.  Match-Confidence Breakdown')
    hdr(['Bank Sheet','Exact (#)','Exact (amt)','Flag (#)','Flag (amt)',
         'Tentative (#)','Tentative (amt)','Unmatched (#)','Unmatched (amt)'],
        widths=[13,10,14,8,14,12,14,12,15])
    for bs, osheet in pairs:
        rs=mr.get(bs,[])
        def c_(s): return sum(1 for r in rs if r['status']==s)
        def a_(s): return round(sum(r['brow'].amount for r in rs if r['status']==s),2)
        drow([bs,c_('MATCHED'),a_('MATCHED'),c_('FLAG'),a_('FLAG'),
              c_('TENTATIVE'),a_('TENTATIVE'),c_('UNMATCHED'),a_('UNMATCHED')],
             ncols={3,5,7,9})
    nr(1)

    # 3 unmatched
    sub('3.  Unmatched Bank Deposits  (light-red on each sheet)')
    hdr(['Bank Sheet','Date','Reference','Amount','Location','Notes'],
        widths=[12,11,20,13,22,40])
    any3=False
    for bs, osheet in pairs:
        for r in sorted(mr.get(bs,[]),key=lambda r:r['brow'].excel_row):
            if r['status']!='UNMATCHED': continue
            any3=True; b=r['brow']
            cfg=data['bank_cfgs'][bs]
            col=cfg['ref_col'] or cfg['dep_col']
            loc=f"'{bs}'!{get_column_letter(col)}{b.excel_row}"
            nt=_note(bs,r,bank_file_label,odoo_file_label) or '—'
            drow([bs,str(b.date),str(b.ref_raw or '(no ref)'),round(b.amount,2),loc,nt],
                 fill=F_UNMATCHED,ncols={4})
    for bs in data['no_pair']:
        for r in mr.get(bs,[]):
            any3=True; b=r['brow']
            cfg=data['bank_cfgs'][bs]
            col=cfg['ref_col'] or cfg['dep_col']
            loc=f"'{bs}'!{get_column_letter(col)}{b.excel_row}"
            drow([bs,str(b.date),str(b.ref_raw or'(no ref)'),round(b.amount,2),loc,
                  'No matching Odoo sheet found for this bank/currency.'],
                 fill=F_UNMATCHED,ncols={4})
    if not any3: drow(['— none —']+['']*5)
    nr(1)

    # 4 flag
    sub('4.  Reference Matched but Amount Differs')
    hdr(['Bank Sheet','Date','Reference','Bank Amt','Odoo Amt','Diff','Bank Location','Odoo Location'],
        widths=[12,11,18,13,13,11,20,24])
    any4=False
    for bs, _ in pairs:
        for r in sorted(mr.get(bs,[]),key=lambda r:r['brow'].excel_row):
            if r['status']!='FLAG': continue
            any4=True; b=r['brow']; o=r['odoo']
            cfg=data['bank_cfgs'][bs]
            col=cfg['ref_col'] or cfg['dep_col']
            bloc=f"'{bs}'!{get_column_letter(col)}{b.excel_row}"
            oloc=f"'{o.sheet}' row {o.excel_row} ({o.number})"
            drow([bs,str(b.date),str(b.ref_raw),round(b.amount,2),round(o.amount,2),
                  round(b.amount-o.amount,2),bloc,oloc],fill=F_FLAG,ncols={4,5,6})
    if not any4: drow(['— none found —']+['']*7)
    nr(1)

    # 5 typos
    sub('5.  Possible Typos — Close Reference Within Same Sheet')
    hdr(['Bank Sheet','Date','Bank Ref','Closest Odoo Ref','Dist','Amount Check',
         'Bank Location','Odoo Location'],widths=[12,11,18,18,6,22,20,24])
    any5=False
    for bs, _ in pairs:
        for r in sorted(mr.get(bs,[]),key=lambda r:r['brow'].excel_row):
            ty=r.get('typo')
            if not ty: continue
            any5=True; b=r['brow']; o2=ty['odoo_row']
            cfg=data['bank_cfgs'][bs]
            col=cfg['ref_col'] or cfg['dep_col']
            bloc=f"'{bs}'!{get_column_letter(col)}{b.excel_row}"
            oloc=f"'{o2.sheet}' row {o2.excel_row} ({o2.number})"
            drow([bs,str(b.date),ty['bank_tok'],ty['odoo_tok'],ty['dist'],
                  ty['amount_note'],bloc,oloc])
    if not any5: drow(['— none found —']+['']*7)
    nr(1)

    # 6 cross-bank
    sub('6.  Possible Wrong-Bank Entries')
    note('Unmatched deposit whose reference appears as an orphan in a DIFFERENT bank\'s Odoo sheet.')
    hdr(['Bank Sheet','Date','Bank Ref','Amount','Bank Location',
         'Found in Odoo Sheet','Odoo Entry','Type','Amount Check'],
        widths=[14,11,16,13,20,18,18,8,20])
    any6=False
    for bs, _ in pairs:
        for r in sorted(mr.get(bs,[]),key=lambda r:r['brow'].excel_row):
            cb=r.get('cross_bank')
            if not cb: continue
            any6=True; b=r['brow']; o2=cb['odoo_row']
            cfg=data['bank_cfgs'][bs]
            col=cfg['ref_col'] or cfg['dep_col']
            bloc=f"'{bs}'!{get_column_letter(col)}{b.excel_row}"
            mtype='Exact' if cb['kind']=='exact' else f"Fuzzy d={cb['dist']}"
            drow([bs,str(b.date),cb['bank_tok'],round(b.amount,2),bloc,
                  o2.sheet,o2.number,mtype,cb['amount_note']],ncols={4})
    if not any6: drow(['— none found —']+['']*8)
    nr(1)

    # 7 orphans
    sub('7.  Odoo Entries With No Matching Bank Deposit (Orphans)')
    note('Recorded in Odoo as a deposit but no bank-side match found. May be wrong bank, timing, or data-entry error.')
    hdr(['Odoo Sheet','Entry Number','Date','Reference','Amount','Location'],
        widths=[12,20,11,24,13,22])
    any7=False
    for os, orows in orp.items():
        for o in sorted(orows,key=lambda o:o.excel_row):
            any7=True
            ref_d=o.ref_raw if o.ref_raw not in (None,'') else (o.ref_fb_raw or '(no ref)')
            cfg=data['odoo_cfgs'].get(os,{})
            rcol=cfg.get('ref_col',4)
            loc=f"'{os}'!{get_column_letter(rcol)}{o.excel_row}"
            drow([os,o.number,str(o.date),str(ref_d),round(o.amount,2),loc],ncols={5})
    if not any7: drow(['— none found —']+['']*5)

    ws.freeze_panes='A4'


def build_output(bank_path, odoo_path, out_path, month_label=''):
    data = reconcile(bank_path, odoo_path)
    shutil.copy(bank_path, out_path)
    wb   = openpyxl.load_workbook(out_path)

    # clear stale XLOOKUP residue on any sheet that has it
    for sname in wb.sheetnames:
        ws = wb[sname]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=ws.max_column, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.lower() in ('notfound','not found'):
                    cell.value = None

    bank_label = os.path.basename(bank_path)
    odoo_label = os.path.basename(odoo_path)

    for bsheet, results in data['match_results'].items():
        if bsheet in wb.sheetnames:
            annotate_bank_sheet(wb, bsheet, results, data['bank_cfgs'][bsheet],
                                bank_label, odoo_label)

    build_report_sheet(wb, data, bank_label, odoo_label, month_label or bank_label)
    wb.save(out_path)

    # summary stats for the UI
    total_bank = sum(len(rs) for rs in data['match_results'].values())
    matched    = sum(sum(1 for r in rs if r['status'] in ('MATCHED','FLAG','TENTATIVE'))
                     for rs in data['match_results'].values())
    unmatched  = total_bank - matched
    flags      = sum(sum(1 for r in rs if r['status']=='FLAG')
                     for rs in data['match_results'].values())
    orphans    = sum(len(v) for v in data['odoo_orphans'].values())
    sheet_pairs= [(bs, os) for bs, os in data['pairs']]
    no_pair    = data['no_pair']

    return dict(total=total_bank, matched=matched, unmatched=unmatched,
                flags=flags, orphans=orphans,
                sheet_pairs=sheet_pairs, no_pair=no_pair)


import os, uuid, threading, webbrowser, time, base64
from flask import Flask, request, jsonify, send_file, Response

app   = Flask(__name__)
jobs  = {}
UPLOAD_DIR = os.path.join(os.path.expanduser("~"), ".bankrecon", "uploads")
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), ".bankrecon", "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

_HTML_B64 = "PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04Ij4KPG1ldGEgbmFtZT0idmlld3BvcnQiIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAiPgo8dGl0bGU+QmFuayBSZWNvbmNpbGlhdGlvbiBUb29sPC90aXRsZT4KPHN0eWxlPgogIDpyb290IHsKICAgIC0taW5rOiAgICAgICMwZjE5MjM7CiAgICAtLWluay1zb2Z0OiAjNGE1NTY4OwogICAgLS1ydWxlOiAgICAgI2UyZThmMDsKICAgIC0tYmc6ICAgICAgICNmN2Y5ZmM7CiAgICAtLXdoaXRlOiAgICAjZmZmZmZmOwogICAgLS1ibHVlOiAgICAgIzFmNGU3ODsKICAgIC0tYmx1ZS1sdDogICMyYjZjYjA7CiAgICAtLWdyZWVuOiAgICAjMjc2NzQ5OwogICAgLS1ncmVlbi1sdDogI2M2ZWZjZTsKICAgIC0tcmVkOiAgICAgICM5YjIzMzU7CiAgICAtLXJlZC1sdDogICAjZmZjN2NlOwogICAgLS1hbWJlcjogICAgIzkyNjAwYTsKICAgIC0tYW1iZXItbHQ6ICNmZmViOWM7CiAgICAtLW9yYW5nZS1sdDojZmZjYzk5OwogICAgLS1yYWRpdXM6ICAgMTBweDsKICAgIC0tc2hhZG93OiAgIDAgMnB4IDE2cHggcmdiYSgxNSwyNSwzNSwuMDgpOwogIH0KCiAgKiwgKjo6YmVmb3JlLCAqOjphZnRlciB7IGJveC1zaXppbmc6IGJvcmRlci1ib3g7IG1hcmdpbjogMDsgcGFkZGluZzogMDsgfQoKICBib2R5IHsKICAgIGZvbnQtZmFtaWx5OiAnU2Vnb2UgVUknLCBzeXN0ZW0tdWksIHNhbnMtc2VyaWY7CiAgICBiYWNrZ3JvdW5kOiB2YXIoLS1iZyk7CiAgICBjb2xvcjogdmFyKC0taW5rKTsKICAgIG1pbi1oZWlnaHQ6IDEwMHZoOwogICAgZGlzcGxheTogZmxleDsKICAgIGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47CiAgfQoKICAvKiDilIDilIAgaGVhZGVyIOKUgOKUgCAqLwogIGhlYWRlciB7CiAgICBiYWNrZ3JvdW5kOiB2YXIoLS1ibHVlKTsKICAgIGNvbG9yOiAjZmZmOwogICAgcGFkZGluZzogMjJweCA0MHB4OwogICAgZGlzcGxheTogZmxleDsKICAgIGFsaWduLWl0ZW1zOiBjZW50ZXI7CiAgICBnYXA6IDE2cHg7CiAgfQogIGhlYWRlciBzdmcgeyBmbGV4LXNocmluazogMDsgfQogIGhlYWRlciBoMSAgeyBmb250LXNpemU6IDEuMjVyZW07IGZvbnQtd2VpZ2h0OiA3MDA7IGxldHRlci1zcGFjaW5nOiAtLjAxZW07IH0KICBoZWFkZXIgcCAgIHsgZm9udC1zaXplOiAuODJyZW07IG9wYWNpdHk6IC43NTsgbWFyZ2luLXRvcDogMnB4OyB9CgogIC8qIOKUgOKUgCBtYWluIGxheW91dCDilIDilIAgKi8KICBtYWluIHsKICAgIGZsZXg6IDE7CiAgICBtYXgtd2lkdGg6IDgyMHB4OwogICAgd2lkdGg6IDEwMCU7CiAgICBtYXJnaW46IDQwcHggYXV0bzsKICAgIHBhZGRpbmc6IDAgMjBweCA2MHB4OwogIH0KCiAgLyog4pSA4pSAIGNhcmQg4pSA4pSAICovCiAgLmNhcmQgewogICAgYmFja2dyb3VuZDogdmFyKC0td2hpdGUpOwogICAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzKTsKICAgIGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdyk7CiAgICBwYWRkaW5nOiAzNnB4IDQwcHg7CiAgICBtYXJnaW4tYm90dG9tOiAyNHB4OwogIH0KICAuY2FyZC10aXRsZSB7CiAgICBmb250LXNpemU6IC43cmVtOwogICAgZm9udC13ZWlnaHQ6IDcwMDsKICAgIGxldHRlci1zcGFjaW5nOiAuMWVtOwogICAgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsKICAgIGNvbG9yOiB2YXIoLS1ibHVlKTsKICAgIG1hcmdpbi1ib3R0b206IDIwcHg7CiAgfQoKICAvKiDilIDilIAgc3RlcCByb3cg4pSA4pSAICovCiAgLnN0ZXAgewogICAgZGlzcGxheTogZmxleDsKICAgIGdhcDogMThweDsKICAgIGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0OwogICAgbWFyZ2luLWJvdHRvbTogMjRweDsKICB9CiAgLnN0ZXA6bGFzdC1jaGlsZCB7IG1hcmdpbi1ib3R0b206IDA7IH0KICAuc3RlcC1udW0gewogICAgd2lkdGg6IDI4cHg7IGhlaWdodDogMjhweDsKICAgIGJvcmRlci1yYWRpdXM6IDUwJTsKICAgIGJhY2tncm91bmQ6IHZhcigtLWJsdWUpOwogICAgY29sb3I6ICNmZmY7CiAgICBmb250LXNpemU6IC43OHJlbTsKICAgIGZvbnQtd2VpZ2h0OiA3MDA7CiAgICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKICAgIGZsZXgtc2hyaW5rOiAwOwogICAgbWFyZ2luLXRvcDogMnB4OwogIH0KICAuc3RlcC1ib2R5IHsgZmxleDogMTsgfQogIC5zdGVwLWJvZHkgbGFiZWwgewogICAgZGlzcGxheTogYmxvY2s7CiAgICBmb250LXNpemU6IC44NXJlbTsKICAgIGZvbnQtd2VpZ2h0OiA2MDA7CiAgICBtYXJnaW4tYm90dG9tOiA2cHg7CiAgICBjb2xvcjogdmFyKC0taW5rKTsKICB9CiAgLnN0ZXAtYm9keSAuaGludCB7CiAgICBmb250LXNpemU6IC43OHJlbTsKICAgIGNvbG9yOiB2YXIoLS1pbmstc29mdCk7CiAgICBtYXJnaW4tYm90dG9tOiAxMHB4OwogIH0KCiAgLyog4pSA4pSAIGZpbGUgZHJvcCB6b25lIOKUgOKUgCAqLwogIC5kcm9wLXpvbmUgewogICAgYm9yZGVyOiAycHggZGFzaGVkIHZhcigtLXJ1bGUpOwogICAgYm9yZGVyLXJhZGl1czogOHB4OwogICAgcGFkZGluZzogMjJweCAxOHB4OwogICAgdGV4dC1hbGlnbjogY2VudGVyOwogICAgY3Vyc29yOiBwb2ludGVyOwogICAgdHJhbnNpdGlvbjogYm9yZGVyLWNvbG9yIC4xOHMsIGJhY2tncm91bmQgLjE4czsKICAgIHBvc2l0aW9uOiByZWxhdGl2ZTsKICB9CiAgLmRyb3Atem9uZTpob3ZlciwgLmRyb3Atem9uZS5kcmFnLW92ZXIgewogICAgYm9yZGVyLWNvbG9yOiB2YXIoLS1ibHVlLWx0KTsKICAgIGJhY2tncm91bmQ6ICNlZWY0ZmI7CiAgfQogIC5kcm9wLXpvbmUgaW5wdXRbdHlwZT1maWxlXSB7CiAgICBwb3NpdGlvbjogYWJzb2x1dGU7IGluc2V0OiAwOyBvcGFjaXR5OiAwOyBjdXJzb3I6IHBvaW50ZXI7IHdpZHRoOiAxMDAlOyBoZWlnaHQ6IDEwMCU7CiAgfQogIC5kcm9wLXpvbmUgLmR6LWljb24geyBjb2xvcjogdmFyKC0tYmx1ZS1sdCk7IG1hcmdpbi1ib3R0b206IDhweDsgfQogIC5kcm9wLXpvbmUgLmR6LXRleHQgeyBmb250LXNpemU6IC44M3JlbTsgY29sb3I6IHZhcigtLWluay1zb2Z0KTsgfQogIC5kcm9wLXpvbmUgLmR6LXRleHQgc3Ryb25nIHsgY29sb3I6IHZhcigtLWJsdWUtbHQpOyB9CiAgLmRyb3Atem9uZS5oYXMtZmlsZSB7IGJvcmRlci1jb2xvcjogdmFyKC0tZ3JlZW4pOyBiYWNrZ3JvdW5kOiAjZjBmYWY0OyB9CiAgLmRyb3Atem9uZS5oYXMtZmlsZSAuZHotaWNvbiB7IGNvbG9yOiB2YXIoLS1ncmVlbik7IH0KICAuZmlsZS1uYW1lIHsKICAgIGZvbnQtc2l6ZTogLjhyZW07IGZvbnQtd2VpZ2h0OiA2MDA7CiAgICBjb2xvcjogdmFyKC0tZ3JlZW4pOwogICAgbWFyZ2luLXRvcDogOHB4OwogICAgZGlzcGxheTogbm9uZTsKICB9CiAgLmRyb3Atem9uZS5oYXMtZmlsZSAuZmlsZS1uYW1lIHsgZGlzcGxheTogYmxvY2s7IH0KICAuZHJvcC16b25lLmhhcy1maWxlIC5kei10ZXh0ICAgeyBkaXNwbGF5OiBub25lOyB9CgogIC8qIOKUgOKUgCB0ZXh0IGlucHV0IOKUgOKUgCAqLwogIGlucHV0W3R5cGU9dGV4dF0gewogICAgd2lkdGg6IDEwMCU7CiAgICBwYWRkaW5nOiA5cHggMTRweDsKICAgIGJvcmRlcjogMS41cHggc29saWQgdmFyKC0tcnVsZSk7CiAgICBib3JkZXItcmFkaXVzOiA3cHg7CiAgICBmb250LXNpemU6IC44OHJlbTsKICAgIGNvbG9yOiB2YXIoLS1pbmspOwogICAgb3V0bGluZTogbm9uZTsKICAgIHRyYW5zaXRpb246IGJvcmRlci1jb2xvciAuMThzOwogICAgZm9udC1mYW1pbHk6IGluaGVyaXQ7CiAgfQogIGlucHV0W3R5cGU9dGV4dF06Zm9jdXMgeyBib3JkZXItY29sb3I6IHZhcigtLWJsdWUtbHQpOyB9CgogIC8qIOKUgOKUgCBzdWJtaXQgYnV0dG9uIOKUgOKUgCAqLwogIC5idG4tcnVuIHsKICAgIHdpZHRoOiAxMDAlOwogICAgcGFkZGluZzogMTRweDsKICAgIGJhY2tncm91bmQ6IHZhcigtLWJsdWUpOwogICAgY29sb3I6ICNmZmY7CiAgICBib3JkZXI6IG5vbmU7CiAgICBib3JkZXItcmFkaXVzOiA4cHg7CiAgICBmb250LXNpemU6IDFyZW07CiAgICBmb250LXdlaWdodDogNzAwOwogICAgY3Vyc29yOiBwb2ludGVyOwogICAgbGV0dGVyLXNwYWNpbmc6IC4wMWVtOwogICAgdHJhbnNpdGlvbjogYmFja2dyb3VuZCAuMThzLCB0cmFuc2Zvcm0gLjFzOwogICAgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsganVzdGlmeS1jb250ZW50OiBjZW50ZXI7IGdhcDogMTBweDsKICB9CiAgLmJ0bi1ydW46aG92ZXI6bm90KDpkaXNhYmxlZCkgeyBiYWNrZ3JvdW5kOiB2YXIoLS1ibHVlLWx0KTsgfQogIC5idG4tcnVuOmFjdGl2ZTpub3QoOmRpc2FibGVkKXsgdHJhbnNmb3JtOiBzY2FsZSguOTkpOyB9CiAgLmJ0bi1ydW46ZGlzYWJsZWQgeyBvcGFjaXR5OiAuNTsgY3Vyc29yOiBub3QtYWxsb3dlZDsgfQoKICAvKiDilIDilIAgcHJvZ3Jlc3Mg4pSA4pSAICovCiAgI3Byb2dyZXNzLXNlY3Rpb24geyBkaXNwbGF5OiBub25lOyB9CiAgLnByb2dyZXNzLWJhci13cmFwIHsKICAgIGhlaWdodDogOHB4OwogICAgYmFja2dyb3VuZDogdmFyKC0tcnVsZSk7CiAgICBib3JkZXItcmFkaXVzOiA5OXB4OwogICAgb3ZlcmZsb3c6IGhpZGRlbjsKICAgIG1hcmdpbjogMTZweCAwOwogIH0KICAucHJvZ3Jlc3MtYmFyIHsKICAgIGhlaWdodDogMTAwJTsKICAgIGJhY2tncm91bmQ6IHZhcigtLWJsdWUtbHQpOwogICAgYm9yZGVyLXJhZGl1czogOTlweDsKICAgIHdpZHRoOiAwJTsKICAgIHRyYW5zaXRpb246IHdpZHRoIC40cyBlYXNlOwogICAgYW5pbWF0aW9uOiBzbGlkZSAxLjRzIGVhc2UtaW4tb3V0IGluZmluaXRlOwogIH0KICBAa2V5ZnJhbWVzIHNsaWRlIHsKICAgIDAlICAgeyB3aWR0aDogMTUlOyBtYXJnaW4tbGVmdDogMDsgfQogICAgNTAlICB7IHdpZHRoOiA0NSU7IG1hcmdpbi1sZWZ0OiAzMCU7IH0KICAgIDEwMCUgeyB3aWR0aDogMTUlOyBtYXJnaW4tbGVmdDogODUlOyB9CiAgfQogIC5wcm9ncmVzcy1iYXIuZG9uZSB7IGFuaW1hdGlvbjogbm9uZTsgd2lkdGg6IDEwMCU7IGJhY2tncm91bmQ6IHZhcigtLWdyZWVuKTsgfQogIC5wcm9ncmVzcy1sYWJlbCB7CiAgICBmb250LXNpemU6IC44NXJlbTsgY29sb3I6IHZhcigtLWluay1zb2Z0KTsgdGV4dC1hbGlnbjogY2VudGVyOwogIH0KCiAgLyog4pSA4pSAIHJlc3VsdHMg4pSA4pSAICovCiAgI3Jlc3VsdC1zZWN0aW9uIHsgZGlzcGxheTogbm9uZTsgfQoKICAuc3RhdHMtZ3JpZCB7CiAgICBkaXNwbGF5OiBncmlkOwogICAgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxNDBweCwgMWZyKSk7CiAgICBnYXA6IDE0cHg7CiAgICBtYXJnaW4tYm90dG9tOiAyNHB4OwogIH0KICAuc3RhdC1ib3ggewogICAgYm9yZGVyLXJhZGl1czogOHB4OwogICAgcGFkZGluZzogMTZweCAxNHB4OwogICAgdGV4dC1hbGlnbjogY2VudGVyOwogIH0KICAuc3RhdC1ib3ggLnN0YXQtbnVtIHsKICAgIGZvbnQtc2l6ZTogMnJlbTsgZm9udC13ZWlnaHQ6IDgwMDsgbGluZS1oZWlnaHQ6IDE7CiAgICBtYXJnaW4tYm90dG9tOiA0cHg7CiAgfQogIC5zdGF0LWJveCAuc3RhdC1sYmwgewogICAgZm9udC1zaXplOiAuNzRyZW07IGZvbnQtd2VpZ2h0OiA2MDA7CiAgICB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogLjA2ZW07CiAgfQogIC5zdGF0LXRvdGFsICAgIHsgYmFja2dyb3VuZDogI2U4ZjBmYTsgY29sb3I6IHZhcigtLWJsdWUpOyB9CiAgLnN0YXQtbWF0Y2hlZCAgeyBiYWNrZ3JvdW5kOiB2YXIoLS1ncmVlbi1sdCk7IGNvbG9yOiB2YXIoLS1ncmVlbik7IH0KICAuc3RhdC11bm1hdGNoZWR7IGJhY2tncm91bmQ6IHZhcigtLXJlZC1sdCk7IGNvbG9yOiB2YXIoLS1yZWQpOyB9CiAgLnN0YXQtZmxhZyAgICAgeyBiYWNrZ3JvdW5kOiB2YXIoLS1vcmFuZ2UtbHQpOyBjb2xvcjogdmFyKC0tYW1iZXIpOyB9CiAgLnN0YXQtb3JwaGFuICAgeyBiYWNrZ3JvdW5kOiB2YXIoLS1hbWJlci1sdCk7IGNvbG9yOiB2YXIoLS1hbWJlcik7IH0KCiAgLnNoZWV0LXBhaXJzIHsKICAgIG1hcmdpbi1ib3R0b206IDIwcHg7CiAgfQogIC5zaGVldC1wYWlycy10aXRsZSB7CiAgICBmb250LXNpemU6IC43OHJlbTsgZm9udC13ZWlnaHQ6IDcwMDsKICAgIHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAuMDdlbTsKICAgIGNvbG9yOiB2YXIoLS1pbmstc29mdCk7IG1hcmdpbi1ib3R0b206IDEwcHg7CiAgfQogIC5wYWlyLXJvdyB7CiAgICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDhweDsKICAgIGZvbnQtc2l6ZTogLjgzcmVtOyBwYWRkaW5nOiA1cHggMDsKICAgIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ydWxlKTsKICB9CiAgLnBhaXItcm93Omxhc3QtY2hpbGQgeyBib3JkZXItYm90dG9tOiBub25lOyB9CiAgLnBhaXItYmFuayAgeyBmb250LXdlaWdodDogNjAwOyBtaW4td2lkdGg6IDEyMHB4OyB9CiAgLnBhaXItYXJyb3cgeyBjb2xvcjogdmFyKC0taW5rLXNvZnQpOyB9CiAgLnBhaXItb2RvbyAgeyBjb2xvcjogdmFyKC0tYmx1ZS1sdCk7IH0KICAucGFpci1ub3BhaXJ7IGNvbG9yOiB2YXIoLS1yZWQpOyBmb250LXN0eWxlOiBpdGFsaWM7IH0KCiAgLmJ0bi1kb3dubG9hZCB7CiAgICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZ2FwOiAxMHB4OwogICAgd2lkdGg6IDEwMCU7CiAgICBwYWRkaW5nOiAxNXB4OwogICAgYmFja2dyb3VuZDogdmFyKC0tZ3JlZW4pOwogICAgY29sb3I6ICNmZmY7CiAgICBib3JkZXI6IG5vbmU7CiAgICBib3JkZXItcmFkaXVzOiA4cHg7CiAgICBmb250LXNpemU6IDFyZW07IGZvbnQtd2VpZ2h0OiA3MDA7CiAgICBjdXJzb3I6IHBvaW50ZXI7CiAgICB0ZXh0LWRlY29yYXRpb246IG5vbmU7CiAgICB0cmFuc2l0aW9uOiBiYWNrZ3JvdW5kIC4xOHM7CiAgfQogIC5idG4tZG93bmxvYWQ6aG92ZXIgeyBiYWNrZ3JvdW5kOiB2YXIoLS1ncmVlbi1sdCk7IGNvbG9yOiB2YXIoLS1ncmVlbik7IH0KCiAgLyog4pSA4pSAIGVycm9yIOKUgOKUgCAqLwogICNlcnJvci1zZWN0aW9uIHsKICAgIGRpc3BsYXk6IG5vbmU7CiAgICBiYWNrZ3JvdW5kOiAjZmZmNWY1OwogICAgYm9yZGVyOiAxLjVweCBzb2xpZCB2YXIoLS1yZWQtbHQpOwogICAgYm9yZGVyLXJhZGl1czogOHB4OwogICAgcGFkZGluZzogMThweCAyMHB4OwogICAgY29sb3I6IHZhcigtLXJlZCk7CiAgICBmb250LXNpemU6IC44OHJlbTsKICB9CgogIC8qIOKUgOKUgCBsZWdlbmQg4pSA4pSAICovCiAgLmxlZ2VuZCB7CiAgICBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogMTBweDsKICAgIG1hcmdpbi10b3A6IDZweDsKICB9CiAgLmxlZ2VuZC1pdGVtIHsKICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogNnB4OwogICAgZm9udC1zaXplOiAuNzhyZW07IGNvbG9yOiB2YXIoLS1pbmstc29mdCk7CiAgfQogIC5sZWdlbmQtc3dhdGNoIHsKICAgIHdpZHRoOiAxNHB4OyBoZWlnaHQ6IDE0cHg7CiAgICBib3JkZXItcmFkaXVzOiAzcHg7IGZsZXgtc2hyaW5rOiAwOwogIH0KCiAgLmJ0bi1hZ2FpbiB7CiAgICBtYXJnaW4tdG9wOiAxNHB4OwogICAgYmFja2dyb3VuZDogbm9uZTsKICAgIGJvcmRlcjogMS41cHggc29saWQgdmFyKC0tcnVsZSk7CiAgICBib3JkZXItcmFkaXVzOiA3cHg7CiAgICBwYWRkaW5nOiAxMHB4IDE4cHg7CiAgICBmb250LXNpemU6IC44NXJlbTsKICAgIGN1cnNvcjogcG9pbnRlcjsKICAgIGNvbG9yOiB2YXIoLS1pbmstc29mdCk7CiAgICBmb250LWZhbWlseTogaW5oZXJpdDsKICAgIHRyYW5zaXRpb246IGJvcmRlci1jb2xvciAuMThzOwogICAgd2lkdGg6IDEwMCU7CiAgfQogIC5idG4tYWdhaW46aG92ZXIgeyBib3JkZXItY29sb3I6IHZhcigtLWJsdWUtbHQpOyBjb2xvcjogdmFyKC0tYmx1ZS1sdCk7IH0KCiAgZm9vdGVyIHsKICAgIHRleHQtYWxpZ246IGNlbnRlcjsKICAgIGZvbnQtc2l6ZTogLjc1cmVtOwogICAgY29sb3I6IHZhcigtLWluay1zb2Z0KTsKICAgIHBhZGRpbmc6IDIwcHg7CiAgICBib3JkZXItdG9wOiAxcHggc29saWQgdmFyKC0tcnVsZSk7CiAgfQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5PgoKPGhlYWRlcj4KICA8c3ZnIHdpZHRoPSIzNiIgaGVpZ2h0PSIzNiIgdmlld0JveD0iMCAwIDM2IDM2IiBmaWxsPSJub25lIj4KICAgIDxyZWN0IHdpZHRoPSIzNiIgaGVpZ2h0PSIzNiIgcng9IjgiIGZpbGw9InJnYmEoMjU1LDI1NSwyNTUsLjE1KSIvPgogICAgPHJlY3QgeD0iNyIgeT0iMTAiIHdpZHRoPSIyMiIgaGVpZ2h0PSIzIiByeD0iMS41IiBmaWxsPSJ3aGl0ZSIvPgogICAgPHJlY3QgeD0iNyIgeT0iMTYiIHdpZHRoPSIxNCIgaGVpZ2h0PSIzIiByeD0iMS41IiBmaWxsPSJ3aGl0ZSIvPgogICAgPHJlY3QgeD0iNyIgeT0iMjIiIHdpZHRoPSIxOCIgaGVpZ2h0PSIzIiByeD0iMS41IiBmaWxsPSJ3aGl0ZSIvPgogICAgPGNpcmNsZSBjeD0iMjgiIGN5PSIyNSIgcj0iNiIgZmlsbD0iIzcwQUQ0NyIvPgogICAgPHBhdGggZD0iTTI1IDI1bDIgMiA0LTQiIHN0cm9rZT0id2hpdGUiIHN0cm9rZS13aWR0aD0iMS41IiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KICA8L3N2Zz4KICA8ZGl2PgogICAgPGgxPkJhbmsgUmVjb25jaWxpYXRpb24gVG9vbDwvaDE+CiAgICA8cD5NYXRjaCBiYW5rIHN0YXRlbWVudCBkZXBvc2l0cyBhZ2FpbnN0IE9kb28gcGF5bWVudCByZWNvcmRzIGF1dG9tYXRpY2FsbHk8L3A+CiAgPC9kaXY+CjwvaGVhZGVyPgoKPG1haW4+CgogIDwhLS0gVXBsb2FkIGZvcm0gLS0+CiAgPGRpdiBpZD0iZm9ybS1zZWN0aW9uIj4KICAgIDxkaXYgY2xhc3M9ImNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJjYXJkLXRpdGxlIj5VcGxvYWQgRmlsZXM8L2Rpdj4KCiAgICAgIDxkaXYgY2xhc3M9InN0ZXAiPgogICAgICAgIDxkaXYgY2xhc3M9InN0ZXAtbnVtIj4xPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic3RlcC1ib2R5Ij4KICAgICAgICAgIDxsYWJlbD5CYW5rIFN0YXRlbWVudCBXb3JrYm9vazwvbGFiZWw+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJoaW50Ij5UaGUgRXhjZWwgZmlsZSBleHBvcnRlZCBmcm9tIHlvdXIgYmFuayAoYWxsIGN1cnJlbmN5IHNoZWV0cyBpbiBvbmUgd29ya2Jvb2spPC9kaXY+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJkcm9wLXpvbmUiIGlkPSJiYW5rLXpvbmUiPgogICAgICAgICAgICA8aW5wdXQgdHlwZT0iZmlsZSIgaWQ9ImJhbmstZmlsZSIgYWNjZXB0PSIueGxzeCwueGxzIj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iZHotaWNvbiI+CiAgICAgICAgICAgICAgPHN2ZyB3aWR0aD0iMjgiIGhlaWdodD0iMjgiIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMS41Ij4KICAgICAgICAgICAgICAgIDxwYXRoIGQ9Ik0xNCAySDZhMiAyIDAgMDAtMiAydjE2YTIgMiAwIDAwMiAyaDEyYTIgMiAwIDAwMi0yVjh6Ii8+CiAgICAgICAgICAgICAgICA8cG9seWxpbmUgcG9pbnRzPSIxNCAyIDE0IDggMjAgOCIvPgogICAgICAgICAgICAgICAgPGxpbmUgeDE9IjEyIiB5MT0iMTgiIHgyPSIxMiIgeTI9IjEyIi8+CiAgICAgICAgICAgICAgICA8bGluZSB4MT0iOSIgeTE9IjE1IiB4Mj0iMTUiIHkyPSIxNSIvPgogICAgICAgICAgICAgIDwvc3ZnPgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iZHotdGV4dCI+PHN0cm9uZz5DbGljayB0byBjaG9vc2U8L3N0cm9uZz4gb3IgZHJhZyAmIGRyb3A8YnI+QmFua19TdGF0ZW1lbnRzX1tNb250aF0ueGxzeDwvZGl2PgogICAgICAgICAgICA8ZGl2IGNsYXNzPSJmaWxlLW5hbWUiIGlkPSJiYW5rLW5hbWUiPjwvZGl2PgogICAgICAgICAgPC9kaXY+CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgoKICAgICAgPGRpdiBjbGFzcz0ic3RlcCI+CiAgICAgICAgPGRpdiBjbGFzcz0ic3RlcC1udW0iPjI8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJzdGVwLWJvZHkiPgogICAgICAgICAgPGxhYmVsPk9kb28gRXhwb3J0IFdvcmtib29rPC9sYWJlbD4KICAgICAgICAgIDxkaXYgY2xhc3M9ImhpbnQiPlRoZSBFeGNlbCBmaWxlIGV4cG9ydGVkIGZyb20gT2RvbyAob25lIHNoZWV0IHBlciBiYW5rIGFjY291bnQpPC9kaXY+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJkcm9wLXpvbmUiIGlkPSJvZG9vLXpvbmUiPgogICAgICAgICAgICA8aW5wdXQgdHlwZT0iZmlsZSIgaWQ9Im9kb28tZmlsZSIgYWNjZXB0PSIueGxzeCwueGxzIj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iZHotaWNvbiI+CiAgICAgICAgICAgICAgPHN2ZyB3aWR0aD0iMjgiIGhlaWdodD0iMjgiIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMS41Ij4KICAgICAgICAgICAgICAgIDxwYXRoIGQ9Ik0xNCAySDZhMiAyIDAgMDAtMiAydjE2YTIgMiAwIDAwMiAyaDEyYTIgMiAwIDAwMi0yVjh6Ii8+CiAgICAgICAgICAgICAgICA8cG9seWxpbmUgcG9pbnRzPSIxNCAyIDE0IDggMjAgOCIvPgogICAgICAgICAgICAgICAgPGxpbmUgeDE9IjEyIiB5MT0iMTgiIHgyPSIxMiIgeTI9IjEyIi8+CiAgICAgICAgICAgICAgICA8bGluZSB4MT0iOSIgeTE9IjE1IiB4Mj0iMTUiIHkyPSIxNSIvPgogICAgICAgICAgICAgIDwvc3ZnPgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iZHotdGV4dCI+PHN0cm9uZz5DbGljayB0byBjaG9vc2U8L3N0cm9uZz4gb3IgZHJhZyAmIGRyb3A8YnI+T2Rvb19CYW5rLnhsc3g8L2Rpdj4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iZmlsZS1uYW1lIiBpZD0ib2Rvby1uYW1lIj48L2Rpdj4KICAgICAgICAgIDwvZGl2PgogICAgICAgIDwvZGl2PgogICAgICA8L2Rpdj4KCiAgICAgIDxkaXYgY2xhc3M9InN0ZXAiPgogICAgICAgIDxkaXYgY2xhc3M9InN0ZXAtbnVtIj4zPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic3RlcC1ib2R5Ij4KICAgICAgICAgIDxsYWJlbD5Nb250aCAvIFBlcmlvZCBMYWJlbCA8c3BhbiBzdHlsZT0iZm9udC13ZWlnaHQ6NDAwO2NvbG9yOnZhcigtLWluay1zb2Z0KSI+KG9wdGlvbmFsKTwvc3Bhbj48L2xhYmVsPgogICAgICAgICAgPGRpdiBjbGFzcz0iaGludCI+VXNlZCBpbiB0aGUgcmVwb3J0IGhlYWRlci4gRS5nLiAiSnVseSAyMDI2IiBvciAiUTMgMjAyNiI8L2Rpdj4KICAgICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0ibW9udGgtbGFiZWwiIHBsYWNlaG9sZGVyPSJlLmcuIEp1bHkgMjAyNiI+CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0iY2FyZCIgc3R5bGU9InBhZGRpbmc6MjRweCA0MHB4Ij4KICAgICAgPGRpdiBjbGFzcz0ibGVnZW5kIj4KICAgICAgICA8ZGl2IGNsYXNzPSJsZWdlbmQtaXRlbSI+PGRpdiBjbGFzcz0ibGVnZW5kLXN3YXRjaCIgc3R5bGU9ImJhY2tncm91bmQ6I0M2RUZDRSI+PC9kaXY+TWF0Y2hlZCAoZXhhY3QpPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ibGVnZW5kLWl0ZW0iPjxkaXYgY2xhc3M9ImxlZ2VuZC1zd2F0Y2giIHN0eWxlPSJiYWNrZ3JvdW5kOiNGRkVCOUMiPjwvZGl2PlRlbnRhdGl2ZSAoYW1vdW50K2RhdGUgb25seSk8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJsZWdlbmQtaXRlbSI+PGRpdiBjbGFzcz0ibGVnZW5kLXN3YXRjaCIgc3R5bGU9ImJhY2tncm91bmQ6I0ZGQ0M5OSI+PC9kaXY+UmVmZXJlbmNlIG1hdGNoZXMsIGFtb3VudCBkaWZmZXJzPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ibGVnZW5kLWl0ZW0iPjxkaXYgY2xhc3M9ImxlZ2VuZC1zd2F0Y2giIHN0eWxlPSJiYWNrZ3JvdW5kOiNGRkM3Q0UiPjwvZGl2Pk5vIG1hdGNoIGZvdW5kPC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CgogICAgPGJ1dHRvbiBjbGFzcz0iYnRuLXJ1biIgaWQ9InJ1bi1idG4iIGRpc2FibGVkIG9uY2xpY2s9InN0YXJ0Sm9iKCkiPgogICAgICA8c3ZnIHdpZHRoPSIxOCIgaGVpZ2h0PSIxOCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIj4KICAgICAgICA8cG9seWdvbiBwb2ludHM9IjUgMyAxOSAxMiA1IDIxIDUgMyIvPgogICAgICA8L3N2Zz4KICAgICAgUnVuIFJlY29uY2lsaWF0aW9uCiAgICA8L2J1dHRvbj4KICA8L2Rpdj4KCiAgPCEtLSBQcm9ncmVzcyAtLT4KICA8ZGl2IGlkPSJwcm9ncmVzcy1zZWN0aW9uIj4KICAgIDxkaXYgY2xhc3M9ImNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJjYXJkLXRpdGxlIj5Qcm9jZXNzaW5n4oCmPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InByb2dyZXNzLWJhci13cmFwIj4KICAgICAgICA8ZGl2IGNsYXNzPSJwcm9ncmVzcy1iYXIiIGlkPSJwcm9ncmVzcy1iYXIiPjwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0icHJvZ3Jlc3MtbGFiZWwiIGlkPSJwcm9ncmVzcy1sYWJlbCI+UmVhZGluZyB3b3JrYm9va3MgYW5kIG1hdGNoaW5nIHJlY29yZHPigKY8L2Rpdj4KICAgIDwvZGl2PgogIDwvZGl2PgoKICA8IS0tIFJlc3VsdHMgLS0+CiAgPGRpdiBpZD0icmVzdWx0LXNlY3Rpb24iPgogICAgPGRpdiBjbGFzcz0iY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9ImNhcmQtdGl0bGUiPlJlc3VsdHM8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic3RhdHMtZ3JpZCIgaWQ9InN0YXRzLWdyaWQiPjwvZGl2PgoKICAgICAgPGRpdiBjbGFzcz0ic2hlZXQtcGFpcnMiIGlkPSJzaGVldC1wYWlycy13cmFwIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzaGVldC1wYWlycy10aXRsZSI+U2hlZXQgUGFpcnMgUmVjb25jaWxlZDwvZGl2PgogICAgICAgIDxkaXYgaWQ9InNoZWV0LXBhaXJzLWxpc3QiPjwvZGl2PgogICAgICA8L2Rpdj4KCiAgICAgIDxhIGNsYXNzPSJidG4tZG93bmxvYWQiIGlkPSJkb3dubG9hZC1idG4iIGhyZWY9IiMiPgogICAgICAgIDxzdmcgd2lkdGg9IjE4IiBoZWlnaHQ9IjE4IiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiPgogICAgICAgICAgPHBhdGggZD0iTTIxIDE1djRhMiAyIDAgMDEtMiAySDVhMiAyIDAgMDEtMi0ydi00Ii8+CiAgICAgICAgICA8cG9seWxpbmUgcG9pbnRzPSI3IDEwIDEyIDE1IDE3IDEwIi8+CiAgICAgICAgICA8bGluZSB4MT0iMTIiIHkxPSIxNSIgeDI9IjEyIiB5Mj0iMyIvPgogICAgICAgIDwvc3ZnPgogICAgICAgIERvd25sb2FkIFJlY29uY2lsaWF0aW9uIFdvcmtib29rCiAgICAgIDwvYT4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuLWFnYWluIiBvbmNsaWNrPSJyZXNldEZvcm0oKSI+4oaQIFJ1biBhbm90aGVyIHJlY29uY2lsaWF0aW9uPC9idXR0b24+CiAgICA8L2Rpdj4KICA8L2Rpdj4KCiAgPCEtLSBFcnJvciAtLT4KICA8ZGl2IGlkPSJlcnJvci1zZWN0aW9uIj4KICAgIDxzdHJvbmc+U29tZXRoaW5nIHdlbnQgd3Jvbmc8L3N0cm9uZz48YnI+CiAgICA8c3BhbiBpZD0iZXJyb3ItbXNnIj48L3NwYW4+CiAgICA8YnI+PGJyPgogICAgPGJ1dHRvbiBjbGFzcz0iYnRuLWFnYWluIiBvbmNsaWNrPSJyZXNldEZvcm0oKSI+4oaQIFRyeSBhZ2FpbjwvYnV0dG9uPgogIDwvZGl2PgoKPC9tYWluPgoKPGZvb3Rlcj5CYW5rIFJlY29uY2lsaWF0aW9uIFRvb2wgwrcgTWF0Y2hlcyBiYW5rIGRlcG9zaXRzIHRvIE9kb28gcmVjb3JkcyBieSByZWZlcmVuY2UgbnVtYmVyPC9mb290ZXI+Cgo8c2NyaXB0PgpsZXQgY3VycmVudEpvYklkID0gbnVsbDsKbGV0IHBvbGxJbnRlcnZhbCA9IG51bGw7CmNvbnN0IHByb2dyZXNzTWVzc2FnZXMgPSBbCiAgJ1JlYWRpbmcgd29ya2Jvb2tzIGFuZCBtYXRjaGluZyByZWNvcmRz4oCmJywKICAnUnVubmluZyBleGFjdCByZWZlcmVuY2UgbWF0Y2hpbmfigKYnLAogICdDaGVja2luZyBhbW91bnQgYW5kIGRhdGUgZmFsbGJhY2tz4oCmJywKICAnU2Nhbm5pbmcgZm9yIHR5cG9zIGFuZCBjcm9zcy1iYW5rIG1pc21hdGNoZXPigKYnLAogICdCdWlsZGluZyBhbm5vdGF0ZWQgb3V0cHV0IHdvcmtib29r4oCmJywKICAnQWRkaW5nIHRvdGFscyByb3dzIGFuZCByZXBvcnQgc2hlZXTigKYnLAogICdBbG1vc3QgZG9uZeKApicsCl07CmxldCBtc2dJZHggPSAwOwoKZnVuY3Rpb24gc2V0dXBEcm9wWm9uZSh6b25lSWQsIGlucHV0SWQsIG5hbWVJZCkgewogIGNvbnN0IHpvbmUgID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoem9uZUlkKTsKICBjb25zdCBpbnB1dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlucHV0SWQpOwogIGNvbnN0IG5hbWUgID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQobmFtZUlkKTsKCiAgaW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgaWYgKGlucHV0LmZpbGVzWzBdKSB7CiAgICAgIHpvbmUuY2xhc3NMaXN0LmFkZCgnaGFzLWZpbGUnKTsKICAgICAgbmFtZS50ZXh0Q29udGVudCA9IGlucHV0LmZpbGVzWzBdLm5hbWU7CiAgICB9CiAgICBjaGVja1JlYWR5KCk7CiAgfSk7CgogIHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignZHJhZ292ZXInLCBlID0+IHsgZS5wcmV2ZW50RGVmYXVsdCgpOyB6b25lLmNsYXNzTGlzdC5hZGQoJ2RyYWctb3ZlcicpOyB9KTsKICB6b25lLmFkZEV2ZW50TGlzdGVuZXIoJ2RyYWdsZWF2ZScsICgpID0+IHpvbmUuY2xhc3NMaXN0LnJlbW92ZSgnZHJhZy1vdmVyJykpOwogIHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignZHJvcCcsIGUgPT4gewogICAgZS5wcmV2ZW50RGVmYXVsdCgpOwogICAgem9uZS5jbGFzc0xpc3QucmVtb3ZlKCdkcmFnLW92ZXInKTsKICAgIGNvbnN0IGYgPSBlLmRhdGFUcmFuc2Zlci5maWxlc1swXTsKICAgIGlmIChmKSB7CiAgICAgIC8vIGNyZWF0ZSBhIERhdGFUcmFuc2ZlciB0byBhc3NpZ24gdG8gaW5wdXQKICAgICAgY29uc3QgZHQgPSBuZXcgRGF0YVRyYW5zZmVyKCk7CiAgICAgIGR0Lml0ZW1zLmFkZChmKTsKICAgICAgaW5wdXQuZmlsZXMgPSBkdC5maWxlczsKICAgICAgem9uZS5jbGFzc0xpc3QuYWRkKCdoYXMtZmlsZScpOwogICAgICBuYW1lLnRleHRDb250ZW50ID0gZi5uYW1lOwogICAgfQogICAgY2hlY2tSZWFkeSgpOwogIH0pOwp9CgpmdW5jdGlvbiBjaGVja1JlYWR5KCkgewogIGNvbnN0IGIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYmFuay1maWxlJykuZmlsZXNbMF07CiAgY29uc3QgbyA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZG9vLWZpbGUnKS5maWxlc1swXTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncnVuLWJ0bicpLmRpc2FibGVkID0gIShiICYmIG8pOwp9CgpmdW5jdGlvbiBzdGFydEpvYigpIHsKICBjb25zdCBiYW5rICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdiYW5rLWZpbGUnKS5maWxlc1swXTsKICBjb25zdCBvZG9vICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZG9vLWZpbGUnKS5maWxlc1swXTsKICBjb25zdCBtb250aCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtb250aC1sYWJlbCcpLnZhbHVlLnRyaW0oKTsKCiAgY29uc3QgZmQgPSBuZXcgRm9ybURhdGEoKTsKICBmZC5hcHBlbmQoJ2JhbmsnLCBiYW5rKTsKICBmZC5hcHBlbmQoJ29kb28nLCBvZG9vKTsKICBmZC5hcHBlbmQoJ21vbnRoJywgbW9udGgpOwoKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZm9ybS1zZWN0aW9uJykuc3R5bGUuZGlzcGxheSAgICAgPSAnbm9uZSc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Byb2dyZXNzLXNlY3Rpb24nKS5zdHlsZS5kaXNwbGF5ID0gJ2Jsb2NrJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncmVzdWx0LXNlY3Rpb24nKS5zdHlsZS5kaXNwbGF5ICAgPSAnbm9uZSc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Vycm9yLXNlY3Rpb24nKS5zdHlsZS5kaXNwbGF5ICAgID0gJ25vbmUnOwoKICBtc2dJZHggPSAwOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcm9ncmVzcy1sYWJlbCcpLnRleHRDb250ZW50ID0gcHJvZ3Jlc3NNZXNzYWdlc1swXTsKCiAgZmV0Y2goJy9ydW4nLCB7IG1ldGhvZDogJ1BPU1QnLCBib2R5OiBmZCB9KQogICAgLnRoZW4ociA9PiByLmpzb24oKSkKICAgIC50aGVuKGRhdGEgPT4gewogICAgICBpZiAoZGF0YS5lcnJvcikgdGhyb3cgbmV3IEVycm9yKGRhdGEuZXJyb3IpOwogICAgICBjdXJyZW50Sm9iSWQgPSBkYXRhLmpvYl9pZDsKICAgICAgcG9sbEludGVydmFsID0gc2V0SW50ZXJ2YWwocG9sbCwgMTIwMCk7CiAgICAgIC8vIGN5Y2xlIHRocm91Z2ggbWVzc2FnZXMKICAgICAgY29uc3QgbXNnVGltZXIgPSBzZXRJbnRlcnZhbCgoKSA9PiB7CiAgICAgICAgbXNnSWR4ID0gKG1zZ0lkeCArIDEpICUgcHJvZ3Jlc3NNZXNzYWdlcy5sZW5ndGg7CiAgICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Byb2dyZXNzLWxhYmVsJykudGV4dENvbnRlbnQgPSBwcm9ncmVzc01lc3NhZ2VzW21zZ0lkeF07CiAgICAgIH0sIDIyMDApOwogICAgICB3aW5kb3cuX21zZ1RpbWVyID0gbXNnVGltZXI7CiAgICB9KQogICAgLmNhdGNoKGVyciA9PiBzaG93RXJyb3IoZXJyLm1lc3NhZ2UpKTsKfQoKZnVuY3Rpb24gcG9sbCgpIHsKICBmZXRjaChgL3N0YXR1cy8ke2N1cnJlbnRKb2JJZH1gKQogICAgLnRoZW4ociA9PiByLmpzb24oKSkKICAgIC50aGVuKGRhdGEgPT4gewogICAgICBpZiAoZGF0YS5zdGF0dXMgPT09ICdydW5uaW5nJykgcmV0dXJuOwogICAgICBjbGVhckludGVydmFsKHBvbGxJbnRlcnZhbCk7CiAgICAgIGNsZWFySW50ZXJ2YWwod2luZG93Ll9tc2dUaW1lcik7CiAgICAgIGlmIChkYXRhLnN0YXR1cyA9PT0gJ2Vycm9yJykgewogICAgICAgIHNob3dFcnJvcihkYXRhLmVycm9yKTsKICAgICAgfSBlbHNlIHsKICAgICAgICBjb25zdCBiYXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncHJvZ3Jlc3MtYmFyJyk7CiAgICAgICAgYmFyLmNsYXNzTGlzdC5hZGQoJ2RvbmUnKTsKICAgICAgICBzZXRUaW1lb3V0KCgpID0+IHNob3dSZXN1bHRzKGRhdGEucmVzdWx0KSwgNTAwKTsKICAgICAgfQogICAgfSkKICAgIC5jYXRjaChlcnIgPT4geyBjbGVhckludGVydmFsKHBvbGxJbnRlcnZhbCk7IHNob3dFcnJvcihlcnIubWVzc2FnZSk7IH0pOwp9CgpmdW5jdGlvbiBzaG93UmVzdWx0cyhyZXN1bHQpIHsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncHJvZ3Jlc3Mtc2VjdGlvbicpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Jlc3VsdC1zZWN0aW9uJykuc3R5bGUuZGlzcGxheSAgID0gJ2Jsb2NrJzsKCiAgY29uc3Qgc3RhdHMgPSBbCiAgICB7IG51bTogcmVzdWx0LnRvdGFsLCAgICAgbGJsOiAnVG90YWwgRGVwb3NpdHMnLCAgIGNsczogJ3N0YXQtdG90YWwnIH0sCiAgICB7IG51bTogcmVzdWx0Lm1hdGNoZWQsICAgbGJsOiAnTWF0Y2hlZCcsICAgICAgICAgICBjbHM6ICdzdGF0LW1hdGNoZWQnIH0sCiAgICB7IG51bTogcmVzdWx0LnVubWF0Y2hlZCwgbGJsOiAnVW5tYXRjaGVkJywgICAgICAgICBjbHM6ICdzdGF0LXVubWF0Y2hlZCcgfSwKICAgIHsgbnVtOiByZXN1bHQuZmxhZ3MsICAgICBsYmw6ICdBbW91bnQgTWlzbWF0Y2hlcycsIGNsczogJ3N0YXQtZmxhZycgfSwKICAgIHsgbnVtOiByZXN1bHQub3JwaGFucywgICBsYmw6ICdPZG9vIE9ycGhhbnMnLCAgICAgIGNsczogJ3N0YXQtb3JwaGFuJyB9LAogIF07CiAgY29uc3QgZ3JpZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzdGF0cy1ncmlkJyk7CiAgZ3JpZC5pbm5lckhUTUwgPSBzdGF0cy5tYXAocyA9PgogICAgYDxkaXYgY2xhc3M9InN0YXQtYm94ICR7cy5jbHN9Ij4KICAgICAgIDxkaXYgY2xhc3M9InN0YXQtbnVtIj4ke3MubnVtfTwvZGl2PgogICAgICAgPGRpdiBjbGFzcz0ic3RhdC1sYmwiPiR7cy5sYmx9PC9kaXY+CiAgICAgPC9kaXY+YAogICkuam9pbignJyk7CgogIGNvbnN0IGxpc3QgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2hlZXQtcGFpcnMtbGlzdCcpOwogIGxldCBodG1sID0gKHJlc3VsdC5zaGVldF9wYWlycyB8fCBbXSkubWFwKChbYiwgb10pID0+CiAgICBgPGRpdiBjbGFzcz0icGFpci1yb3ciPgogICAgICAgPHNwYW4gY2xhc3M9InBhaXItYmFuayI+JHtifTwvc3Bhbj4KICAgICAgIDxzcGFuIGNsYXNzPSJwYWlyLWFycm93Ij7ihpI8L3NwYW4+CiAgICAgICA8c3BhbiBjbGFzcz0icGFpci1vZG9vIj4ke299PC9zcGFuPgogICAgIDwvZGl2PmAKICApLmpvaW4oJycpOwogIGlmIChyZXN1bHQubm9fcGFpciAmJiByZXN1bHQubm9fcGFpci5sZW5ndGgpIHsKICAgIGh0bWwgKz0gcmVzdWx0Lm5vX3BhaXIubWFwKGIgPT4KICAgICAgYDxkaXYgY2xhc3M9InBhaXItcm93Ij4KICAgICAgICAgPHNwYW4gY2xhc3M9InBhaXItYmFuayI+JHtifTwvc3Bhbj4KICAgICAgICAgPHNwYW4gY2xhc3M9InBhaXItYXJyb3ciPuKGkjwvc3Bhbj4KICAgICAgICAgPHNwYW4gY2xhc3M9InBhaXItbm9wYWlyIj5ObyBtYXRjaGluZyBPZG9vIHNoZWV0IGZvdW5kPC9zcGFuPgogICAgICAgPC9kaXY+YAogICAgKS5qb2luKCcnKTsKICB9CiAgbGlzdC5pbm5lckhUTUwgPSBodG1sOwoKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZG93bmxvYWQtYnRuJykuaHJlZiA9IGAvZG93bmxvYWQvJHtjdXJyZW50Sm9iSWR9YDsKfQoKZnVuY3Rpb24gc2hvd0Vycm9yKG1zZykgewogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcm9ncmVzcy1zZWN0aW9uJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZXJyb3Itc2VjdGlvbicpLnN0eWxlLmRpc3BsYXkgICAgPSAnYmxvY2snOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdlcnJvci1tc2cnKS50ZXh0Q29udGVudCA9IG1zZzsKfQoKZnVuY3Rpb24gcmVzZXRGb3JtKCkgewogIGN1cnJlbnRKb2JJZCA9IG51bGw7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Zvcm0tc2VjdGlvbicpLnN0eWxlLmRpc3BsYXkgICAgID0gJ2Jsb2NrJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncHJvZ3Jlc3Mtc2VjdGlvbicpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Jlc3VsdC1zZWN0aW9uJykuc3R5bGUuZGlzcGxheSAgID0gJ25vbmUnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdlcnJvci1zZWN0aW9uJykuc3R5bGUuZGlzcGxheSAgICA9ICdub25lJzsKICBbJ2JhbmsnLCAnb2RvbyddLmZvckVhY2goaWQgPT4gewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoYCR7aWR9LXpvbmVgKS5jbGFzc0xpc3QucmVtb3ZlKCdoYXMtZmlsZScpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoYCR7aWR9LWZpbGVgKS52YWx1ZSA9ICcnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoYCR7aWR9LW5hbWVgKS50ZXh0Q29udGVudCA9ICcnOwogIH0pOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdydW4tYnRuJykuZGlzYWJsZWQgPSB0cnVlOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwcm9ncmVzcy1iYXInKS5jbGFzc0xpc3QucmVtb3ZlKCdkb25lJyk7Cn0KCi8vIGluaXQKc2V0dXBEcm9wWm9uZSgnYmFuay16b25lJywgJ2JhbmstZmlsZScsICdiYW5rLW5hbWUnKTsKc2V0dXBEcm9wWm9uZSgnb2Rvby16b25lJywgJ29kb28tZmlsZScsICdvZG9vLW5hbWUnKTsKPC9zY3JpcHQ+CjwvYm9keT4KPC9odG1sPgo="
_HTML = base64.b64decode(_HTML_B64).decode()

@app.route("/")
def index():
    return Response(_HTML, mimetype="text/html")

@app.route("/run", methods=["POST"])
def run_route():
    bank_file   = request.files.get("bank")
    odoo_file   = request.files.get("odoo")
    month_label = request.form.get("month", "").strip()
    if not bank_file or not odoo_file:
        return jsonify(error="Both files are required."), 400
    job_id    = str(uuid.uuid4())
    bank_path = os.path.join(UPLOAD_DIR, f"{job_id}_bank.xlsx")
    odoo_path = os.path.join(UPLOAD_DIR, f"{job_id}_odoo.xlsx")
    out_path  = os.path.join(OUTPUT_DIR, f"{job_id}_result.xlsx")
    bank_file.save(bank_path)
    odoo_file.save(odoo_path)
    jobs[job_id] = dict(status="running")
    t = threading.Thread(target=_run_job,
                         args=(job_id, bank_path, odoo_path, out_path, month_label),
                         daemon=True)
    t.start()
    return jsonify(job_id=job_id)

def _run_job(job_id, bank_path, odoo_path, out_path, month_label):
    try:
        result = build_output(bank_path, odoo_path, out_path, month_label)
        jobs[job_id] = dict(status="done", result=result, out_path=out_path)
    except Exception as e:
        import traceback
        jobs[job_id] = dict(status="error", error=str(e), trace=traceback.format_exc())
    finally:
        for p in (bank_path, odoo_path):
            try: os.remove(p)
            except: pass

@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job: return jsonify(error="Unknown job"), 404
    if job["status"] == "error":   return jsonify(status="error", error=job["error"])
    if job["status"] == "running": return jsonify(status="running")
    return jsonify(status="done", result=job["result"])

@app.route("/download/<job_id>")
def download(job_id):
    job = jobs.get(job_id)
    if not job or job["status"] != "done": return "Not ready", 404
    return send_file(job["out_path"], as_attachment=True,
                     download_name="Bank_Reconciliation_Reviewed.xlsx")

if __name__ == "__main__":
    port = 5050
    url  = f"http://localhost:{port}"
    print(f"""
  ╔══════════════════════════════════════════╗
  ║   Bank Reconciliation Tool  — running    ║
  ║   Open your browser at:                  ║
  ║     http://localhost:5050                ║
  ║   Press Ctrl+C to stop                   ║
  ╚══════════════════════════════════════════╝
""")
    def _open(): time.sleep(1.2); webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
