"""
fs_audit.py - integrity audit of a financial-statements workbook

  1. double links      : the same ledger / schedule cell feeding two statement lines, the same cell referenced twice inside
                         one formula, the same link repeated twice inside one total, a subtotal summed together with its own rows
  2. subtotal sums     : total rows whose SUM / '+' skips rows, typed (hard-coded) totals, totals built differently from the
                         neighbouring period columns
  3. gap reasons table : every finding put in one table with location, amount, effect on the balance sheet gap and the evidence

Effect on the gap is always  (Total Equity & Liabilities - Total Assets).
  A line that is overstated by d changes the gap by:  asset -d | liability / equity +d | income +d | expense -d
  (income and expense reach equity through profit)
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict

import pandas as pd
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

import fs_builder as bld

TOTAL_RX = re.compile(r"^\W*(grand\s+total|sub\s*-?\s*total|total|net\s+(total|block|profit|carrying))\b", re.I)
SUM_RX = re.compile(r"^=\s*[\(\s]*(SUM|SUBTOTAL)\s*\(", re.I)
QUOTED_SHEET = re.compile(r"'[^']*'!")
MAX_ROWS, MAX_COLS = 3000, 60


def effect_on_gap(nature, delta):
    """delta = reported - correct (positive = overstated)"""
    return {"A": -delta, "L": delta, "I": delta, "E": -delta}.get(nature, 0.0) if delta is not None else 0.0


def _is_sumlike(raw):
    if not isinstance(raw, str) or not raw.startswith("="):
        return False
    t = QUOTED_SHEET.sub("X!", raw)
    if SUM_RX.match(t):
        return not re.search(r"[*/]", t) and t.count("-") == 0
    # references joined by + only (at least two)
    if re.fullmatch(r"=[\s\(\)\$A-Za-z0-9:+!_.]+", t) and t.count("+") >= 1:
        return True
    return False


# --------------------------------------------------------------------------------------------- which sheets are schedules
def schedule_sheets(bk, start, tb_sheets, limit=40, max_rows=MAX_ROWS):
    """sheets reachable from the statement sheets through formulas (the notes / schedules / workings), trial balances excluded"""
    names = set(bk.sheetnames)
    seen, queue = [], list(start)
    tbs = set(tb_sheets)
    rx = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_\.]+))!")
    while queue and len(seen) < limit:
        sh = queue.pop(0)
        if sh in seen or sh in tbs or sh not in names:
            continue
        if (bk.wf[sh].max_row or 0) > max_rows and sh not in start:
            continue
        seen.append(sh)
        for row in bk.wf[sh].iter_rows(max_row=min(bk.wf[sh].max_row or 1, max_rows), max_col=MAX_COLS, values_only=True):
            for v in row:
                if isinstance(v, str) and v.startswith("=") and "!" in v:
                    for m in rx.finditer(v):
                        n = (m.group(1) or m.group(2) or "").replace("''", "'")
                        if n in names and n not in seen and n not in tbs:
                            queue.append(n)
    return seen


# --------------------------------------------------------------------------------------------- sums inside one sheet
def audit_sheet(eng, bk, sh, tol=0.01, relevant=None, check_shape=False):
    statement_sheet = check_shape
    F = list(bk.wf[sh].iter_rows(min_row=1, max_row=min(bk.wf[sh].max_row or 1, MAX_ROWS), max_col=MAX_COLS, values_only=True))
    V = list(bk.wv[sh].iter_rows(min_row=1, max_row=min(bk.wv[sh].max_row or 1, MAX_ROWS), max_col=MAX_COLS, values_only=True))
    nr = len(F)
    out = []
    lab_cache = {}

    def raw(r, c):
        return F[r - 1][c - 1] if 1 <= r <= nr and c <= len(F[r - 1]) else None

    def val(r, c):
        v = V[r - 1][c - 1] if 1 <= r <= nr and c <= len(V[r - 1]) else None
        return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    def label(r):
        if r not in lab_cache:
            lab_cache[r] = eng.label_for_row(bk, sh, r) or f"row {r}"
        return lab_cache[r]

    def tot_label(r):
        return bool(TOTAL_RX.match(label(r)))

    def sumlike(r, c):
        return _is_sumlike(raw(r, c))

    cov_cache = {}

    def covered(r, c):
        k = (r, c)
        if k not in cov_cache:
            rows, pure = set(), True
            for csh, ca, _sg in bk.expand(sh, eng.addr(c, r)):
                cc, rr = eng.split_addr(ca)
                if csh == sh and cc == c:
                    rows.add(rr)
                else:
                    pure = False
            cov_cache[k] = (rows, pure)
        return cov_cache[k]

    def add(cell, check, detail, amount, formula, sev=None, affects_row=None):
        if relevant is not None and (sh, cell) not in relevant:
            return
        out.append({"Sheet": sh, "Cell": cell, "Item": label(affects_row) if affects_row else label(int(re.sub(r"[A-Z]+", "", cell))),
                    "Check": check, "Detail": detail, "Amount": amount, "Formula": formula if isinstance(formula, str) else "",
                    "Severity": sev or ("High" if abs(amount or 0) > tol else "Low")})

    totals_by_row = defaultdict(list)
    for r in range(1, nr + 1):
        for c in range(1, len(F[r - 1]) + 1):
            rw = raw(r, c)
            if not (isinstance(rw, str) and rw.startswith("=")):
                continue
            if sumlike(r, c):
                totals_by_row[r].append(c)
    for r, cols in totals_by_row.items():
        for c in cols:
            rw = raw(r, c)
            R, pure = covered(r, c)
            if not pure or len(R) < 2 or r in R:
                continue
            a, b = min(R), max(R)
            comps = {s: covered(s, c)[0] for s in R if s != r and sumlike(s, c)}
            eff = set(R)
            stack = list(comps)
            while stack:
                s = stack.pop()
                for x in comps.get(s, covered(s, c)[0]):
                    if x not in eff:
                        eff.add(x)
                        if sumlike(x, c):
                            comps.setdefault(x, covered(x, c)[0])
                            stack.append(x)
            cell = f"{bld.get_column_letter(c)}{r}"
            # ---- rows the total leaves out
            is_range = bool(SUM_RX.match(QUOTED_SHEET.sub("X!", rw)))
            if is_range:
                cand = [x for x in range(a, r) if x not in eff and (x <= b or r - b <= 3)]
                x = a - 1   # a range that starts one row late: only the row directly above, and only if it is a plain figure (not a calculation)
                rx_ = raw(x, c) if x >= 1 else None
                kind_rows = {isinstance(raw(y, c), str) and str(raw(y, c)).startswith("=") for y in R}
                if (x >= 1 and r - b <= 3 and val(x, c) is not None and not tot_label(x) and not sumlike(x, c)
                        and not (isinstance(rx_, str) and re.search(r"[-*/]", QUOTED_SHEET.sub("X!", rx_)))
                        and kind_rows == {isinstance(rx_, str) and rx_.startswith("=")}):
                    cand.append(x)
            else:
                # '+' chain: only a hole inside an otherwise contiguous run looks like a forgotten row
                holes = [x for x in range(a, b + 1) if x not in eff]
                cand = holes if len(R) >= 3 and len(holes) <= 2 else []
            for m in sorted(set(cand)):
                v = val(m, c)
                if v is None or m in R or m in eff or (sumlike(m, c) and covered(m, c)[0] and covered(m, c)[0] <= eff):
                    continue
                if abs(v) > tol:
                    add(cell, "Subtotal skips a row", f"{rw_text(F, r, c)} adds rows {a}-{b} but row {m} ({label(m)}) sits in the same block and is left out",
                        -v, raw(r, c), affects_row=r)
            # ---- a subtotal summed together with its own rows
            for s, cs in comps.items():
                if cs and (cs & (R - {s})) and val(s, c) is not None and abs(val(s, c)) > tol:
                    add(cell, "Subtotal added together with its own rows",
                        f"row {s} ({label(s)}) is itself a total of rows {min(cs)}-{max(cs)}, and those rows are also inside this total, so {val(s, c):,.2f} is counted twice",
                        val(s, c), raw(r, c), affects_row=r)
            # ---- the same cell referenced twice in one formula
            for (csh, ca, sg), n in bk._mult.get((sh, eng.addr(c, r)), {}).items():
                if n > 1:
                    v = eng.num(bk.val(csh, ca)) or 0.0
                    add(cell, "Cell referenced twice in one formula", f"{csh}!{ca} appears {n} times in the formula, so {v:,.2f} is counted {n} times",
                        v * (n - 1), raw(r, c), affects_row=r)
            # ---- the same link repeated twice inside one total
            seen_f = defaultdict(list)
            for x in ([] if statement_sheet else sorted(R)):
                if x in comps:
                    continue
                f = raw(x, c)
                if isinstance(f, str) and f.startswith("=") and ("!" in f or "," in f):
                    seen_f[re.sub(r"\s+", "", f.upper())].append(x)
            for f, rows_ in seen_f.items():
                if len(rows_) > 1:
                    v = val(rows_[0], c) or val(rows_[1], c) or 0.0
                    add(cell, "Same link used twice inside a total",
                        f"rows {', '.join(map(str, rows_))} ({'; '.join(label(x) for x in rows_[:3])}) have the identical formula {raw(rows_[0], c)}, so the source is added {len(rows_)} times",
                        v * (len(rows_) - 1), raw(rows_[0], c), affects_row=rows_[1])
    # ---- typed (hard-coded) totals
    for r in range(1, nr + 1):
        if not tot_label(r):
            continue
        for c in range(2, len(F[r - 1]) + 1):
            rw = raw(r, c)
            v = val(r, c)
            if v is None or (isinstance(rw, str) and rw.startswith("=")):
                continue
            rows_ = []
            x = r - 1
            while x >= 1 and val(x, c) is not None and not tot_label(x):
                rows_.append(x)
                x -= 1
            if len(rows_) >= 2:
                s = sum(val(x, c) or 0 for x in rows_)
                add(f"{bld.get_column_letter(c)}{r}", "Total is typed, not a formula",
                    f"{label(r)} is a typed number {v:,.2f}; the {len(rows_)} rows above add up to {s:,.2f} (difference {v - s:,.2f})", v - s, str(rw))
    # ---- total built differently from the neighbouring period columns
    for r, cols in (totals_by_row.items() if check_shape else []):
        if len(cols) < 3:
            continue
        shapes = {}
        for c in cols:
            try:
                shapes[c] = (eng._shape(raw(r, c), c), raw(r, c))
            except Exception:  # noqa: BLE001
                pass
        cnt = Counter(s for s, _ in shapes.values())
        if len(cnt) < 2:
            continue
        (dom, n_dom), = cnt.most_common(1)
        if n_dom < 2 or list(cnt.values()).count(n_dom) > 1:
            continue
        dcol = next(c for c, (s, _) in shapes.items() if s == dom)
        drows = eng._rows_in(shapes[dcol][1])
        for c, (s, rw) in shapes.items():
            if s == dom:
                continue
            mine = eng._rows_in(rw)
            diff = mine ^ drows
            impact = sum(abs(val(x, c) or 0) for x in diff if x != r)
            miss, extra = sorted(drows - mine), sorted(mine - drows)
            what = ([f"leaves out row(s) {', '.join(map(str, miss))}"] if miss else []) + ([f"adds row(s) {', '.join(map(str, extra))}"] if extra else [])
            add(f"{bld.get_column_letter(c)}{r}", "Total built differently from the other period columns",
                f"{rw} {' and '.join(what) or 'uses a different layout'} (the other {n_dom} column(s) use {shapes[dcol][1]})", impact if impact > tol else 0.0, rw,
                sev="High" if impact > tol else "Low")
    return out


def rw_text(F, r, c):
    v = F[r - 1][c - 1] if r <= len(F) and c <= len(F[r - 1]) else ""
    return str(v) if isinstance(v, str) else ""


# --------------------------------------------------------------------------------------------- statement lines and their dependencies
def _tcache(bk):
    return bk.__dict__.setdefault("_audit_trace_cache", {})


def trace_cached(eng, bk, tbs, sheet, a):
    k = (tuple(tbs), sheet, a)
    c = _tcache(bk)
    if k not in c:
        c[k] = eng.trace_line(bk, list(tbs), sheet, a)
    return c[k]


def pl_line_cells(bk, sheet, col):
    ws = bk.wv[sheet]
    rows = []
    for r in range(1, min(ws.max_row or 1, 80) + 1):
        lab = ""
        for c in range(1, 5):
            v = ws.cell(r, c).value
            if isinstance(v, str) and bld.clean(v):
                lab = bld.clean(v)
                break
        if lab and not re.match(r"^total", lab, re.I):
            rows.append((r, lab))
    out = {}
    for lid, name in bld.PL_TARGETS:
        cand = [(r, bld._label_sim(lab, name)) for r, lab in rows]
        cand = [x for x in cand if x[1] >= 0.8]
        if cand:
            r = max(cand, key=lambda x: x[1])[0]
            out[lid] = f"{bld.get_column_letter(col)}{r}"
    return out


def statement_lines(eng, bk, bs_sheet, bs_col, pl_sheet=None, pl_col=None):
    """[{key, label, nature, stmt, sheet, addr, value}] for every balance sheet and P&L line of the original"""
    st = eng.find_structure(bk, bs_sheet)
    lines = []
    for ln in bld.match_original_lines(eng, bk, st, bs_col):
        if ln["match"]:
            lines.append({"key": ln["match"], "label": ln["label"], "nature": bld.LINE[ln["match"]]["nature"], "stmt": "BS", "sheet": bs_sheet,
                          "addr": ln["addr"], "value": ln["value"]})
    if pl_sheet and pl_col:
        for lid, a in pl_line_cells(bk, pl_sheet, pl_col).items():
            lab = bld.LINE[lid]["label"]
            lines.append({"key": lid, "label": lab, "nature": bld.LINE[lid]["nature"], "stmt": "PL", "sheet": pl_sheet, "addr": a,
                          "value": eng.num(bk.val(pl_sheet, a))})
    return lines


def audit_links(eng, bk, lines, tbs, ledgers=None, tol=0.01, scale=1.0):
    """double links at statement level + dependency map (which statement lines does a cell feed)"""
    out = []
    # ---- same source cell feeding two statement lines
    direct = defaultdict(list)
    for ln in lines:
        for csh, ca, sg in bk.expand(ln["sheet"], ln["addr"]):
            if csh != ln["sheet"] or ln["stmt"] == "PL":
                direct[(csh, ca)].append(ln)
    for (csh, ca), ls in direct.items():
        keys = {(l["stmt"], l["key"], l["addr"]) for l in ls}
        if len(keys) > 1 and csh not in tbs:
            v = eng.num(bk.val(csh, ca)) or 0.0
            if abs(v) <= tol:
                continue
            first, second = ls[0], ls[-1]
            out.append({"Sheet": csh, "Cell": ca, "Item": eng.label_for_row(bk, csh, eng.split_addr(ca)[1]) or ca,
                        "Check": "Same cell feeds two statement lines",
                        "Detail": f"{csh}!{ca} ({v:,.2f}) is linked into {' and '.join(l['label'].strip() + ' [' + l['sheet'] + '!' + l['addr'] + ']' for l in ls[:3])}",
                        "Amount": v, "Formula": "", "Severity": "High", "_lines": [l["key"] for l in ls], "_nature": second["nature"]})
    flagged_sets = [set(f['_lines']) for f in out if '_lines' in f]
    # ---- ledgers read by two lines of the same statement
    line_rows, contrib = defaultdict(set), defaultdict(float)
    for ln in lines:
        if ln["key"] == "OE":
            continue
        try:
            tr = trace_cached(eng, bk, tbs, ln["sheet"], ln["addr"])
        except Exception:  # noqa: BLE001
            continue
        for (tsh, ta), sg in tr["tb"].items():
            rr = int(re.sub(r"[A-Z]+", "", ta))
            line_rows[(ln["stmt"], ln["key"])].add((tsh, rr))
            contrib[((ln["stmt"], ln["key"]), (tsh, rr))] += sg * (eng.num(bk.val(tsh, ta)) or 0.0)
    allrows = set().union(*line_rows.values()) if line_rows else set()
    broad = {k for k, v in line_rows.items() if allrows and len(v) > 0.4 * len(allrows)}
    per = defaultdict(list)
    for k, rs in line_rows.items():
        if k in broad:
            continue
        for x in rs:
            per[x].append(k)
    names = {}
    if ledgers is not None:
        for rec in ledgers.to_dict("records"):
            names[int(rec["Row"])] = rec
    by_key = {(l["stmt"], l["key"]): l for l in lines}
    for (tsh, rr), ks in per.items():
        for stmt in ("BS", "PL"):
            kk = [k for k in ks if k[0] == stmt]
            if len(kk) < 2:
                continue
            c = [abs(contrib[(k, (tsh, rr))]) for k in kk]
            amt = max(c) if c else 0.0
            if amt <= tol:
                continue
            rec = names.get(rr)
            if stmt == "BS" and rec is not None and bld.LINE.get(rec.get("Line"), {}).get("stmt") == "PL":
                continue   # tax / provision workings legitimately read income and expense ledgers
            if any(set(k[1] for k in kk) <= fs for fs in flagged_sets):
                continue   # already explained by the same cell feeding both lines
            led = rec["Ledger"] if rec else eng.label_for_row(bk, tsh, rr) or f"{tsh} row {rr}"
            second = by_key[kk[-1]]
            out.append({"Sheet": tsh, "Cell": f"row {rr}", "Item": led, "Check": "Ledger read by two lines of the same statement",
                        "Detail": f"ledger '{led}' ({tsh} row {rr}, {amt:,.2f}) is read by {' and '.join(by_key[k]['label'].strip() for k in kk[:3])}",
                        "Amount": amt, "Formula": "", "Severity": "High", "_lines": [k[1] for k in kk], "_nature": second["nature"]})
    # ---- dependency map: statement line -> cells it depends on (non-TB sheets)
    dep = {}
    for ln in lines:
        try:
            tr = trace_cached(eng, bk, tbs, ln["sheet"], ln["addr"])
            dep[(ln["stmt"], ln["key"])] = {c for c in tr["seen"] if c[0] not in tbs}
        except Exception:  # noqa: BLE001
            dep[(ln["stmt"], ln["key"])] = set()
    return out, dep


def run_audit(eng, bk, bs_sheet, bs_col, tbs, pl_sheet=None, pl_col=None, ledgers=None, tol=0.01, scale=1.0, extra_sheets=()):
    """-> DataFrame of findings with Affects / Effect on gap"""
    lines = statement_lines(eng, bk, bs_sheet, bs_col, pl_sheet, pl_col)
    # every cell that feeds a number of the statement columns (statement cells, totals, notes, workings) - the audit looks only at these
    relevant = set()
    st = eng.find_structure(bk, bs_sheet)
    roots = []
    if st:
        for r in range(min(st["a_hdr"], st["l_hdr"]), max(st["a_tot"], st["l_tot"]) + 1):
            v = bk.raw(bs_sheet, eng.addr(bs_col, r))
            if v is not None:
                roots.append((bs_sheet, eng.addr(bs_col, r)))
    if pl_sheet and pl_col:
        for r in range(1, min(bk.maxrow(pl_sheet), 80) + 1):
            if bk.raw(pl_sheet, eng.addr(pl_col, r)) is not None:
                roots.append((pl_sheet, eng.addr(pl_col, r)))
    for sh0, a0 in roots:
        try:
            tr = trace_cached(eng, bk, tbs, sh0, a0)
        except Exception:  # noqa: BLE001
            continue
        relevant.update(c for c in tr["seen"] if c[0] not in tbs)
    sheets = sorted({c[0] for c in relevant}, key=lambda x: (x not in (bs_sheet, pl_sheet), x))
    found = []
    for sh in sheets:
        try:
            found += audit_sheet(eng, bk, sh, tol, relevant, check_shape=sh in (bs_sheet, pl_sheet))
        except Exception as e:  # noqa: BLE001
            found.append({"Sheet": sh, "Cell": "", "Item": "", "Check": "Sheet could not be audited", "Detail": str(e)[:200], "Amount": 0.0,
                          "Formula": "", "Severity": "Low"})
    lk, dep = audit_links(eng, bk, lines, tbs, ledgers, tol, scale)
    found += lk
    by_key = {(l["stmt"], l["key"]): l for l in lines}
    rows = []
    for f in found:
        aff = []
        if "_lines" in f:
            aff = [by_key[k] for k in by_key if k[1] in f["_lines"]]
        else:
            cell = (f["Sheet"], f["Cell"])
            hit = [(len(cells), k) for k, cells in dep.items() if cell in cells]
            hit.sort()
            aff = [by_key[k] for _, k in hit[:3]]
        nat = f.get("_nature") or (aff[0]["nature"] if aff else None)
        amt = f.get("Amount") or 0.0
        lv = abs(aff[0]["value"]) if aff and aff[0].get("value") else None
        eff = effect_on_gap(nat, amt) if nat and not (lv and abs(amt) > 1000 * lv) else None
        rows.append({"Area": "Schedule / workings" if f["Sheet"] not in (bs_sheet, pl_sheet) else "Statement", "Check": f["Check"], "Severity": f["Severity"],
                     "Sheet": f["Sheet"], "Cell": f["Cell"], "Item": f["Item"], "Detail": f["Detail"],
                     "Amount": amt, "Affects": "; ".join(dict.fromkeys(a["label"].strip() for a in aff)),
                     "Effect on gap": eff, "Formula": f.get("Formula", "")})
    df = pd.DataFrame(rows, columns=["Area", "Check", "Severity", "Sheet", "Cell", "Item", "Detail", "Amount", "Affects", "Effect on gap", "Formula"])
    if len(df):
        df["_abs"] = df["Amount"].abs()
        df = df.sort_values(["Severity", "_abs"], ascending=[True, False]).drop(columns="_abs").reset_index(drop=True)
    return df, lines, dep


# --------------------------------------------------------------------------------------------- gap reasons table
def _evidence(d, unit):
    bits = []
    if d.get("only_orig"):
        bits.append("Original reads ledger(s) that the rebuild groups elsewhere: " + "; ".join(f"{x['ledger']} {x['net']:,.2f}" for x in d["only_orig"][:5]))
    if d.get("only_auto"):
        bits.append("Rebuild groups ledger(s) here that the original does not read: " + "; ".join(f"{x['ledger']} {x['net']:,.2f}" for x in d["only_auto"][:5]))
    sched = [x for x in d.get("schedule", []) if x["found"]]
    if sched:
        sched.sort(key=lambda x: -abs(x["diff"] or 0))
        bits.append("Schedule vs ledger: " + "; ".join(f"{x['ledger']} (ledger {x['amount']:,.2f}) vs {x['found']} = {x['amount'] + (x['diff'] or 0):,.2f} [{(x['diff'] or 0):+,.2f}]" for x in sched[:4]))
    if d.get("typed"):
        bits.append("Typed numbers in the chain: " + "; ".join(f"{t['cell']} = {t['value']:,.2f}" for t in d["typed"][:4]))
    return " | ".join(bits)


def gap_reason_tables(rc, findings: pd.DataFrame, unit="Lakhs"):
    """-> (bridge DataFrame, reasons DataFrame, summary dict)"""
    tol = rc["tol"]
    fl = findings if findings is not None else pd.DataFrame()
    bridge = []
    n = 0
    for r in rc["rows"]:
        if abs(r["var"]) <= tol and not r["err"]:
            continue
        n += 1
        d = rc["drill"].get(r["id"], {})
        eff = r["var"] if r["nature"] == "L" else -r["var"]
        why = []
        if r["err"]:
            why.append("formula error in " + ", ".join(r["err_cells"]))
        if r["id"] == "OE":
            why.append("Other equity is rolled forward from opening reserves + profit in the original; the rebuild takes reserves + P&L opening + current profit from the trial balance")
        elif d.get("sheets") and not d.get("reads_tb"):
            why.append("fed from schedule(s) " + ", ".join(d["sheets"][:4]) + " instead of the trial balance")
        elif d.get("sheets"):
            why.append("fed by schedule(s) " + ", ".join(d["sheets"][:4]) + f" and {len(d.get('reads_tb', []))} trial balance row(s)")
        hit = fl[fl["Affects"].str.contains(re.escape(r["label"].strip()), regex=True, na=False)] if len(fl) else fl
        bridge.append({"#": n, "Statement line": r["label"].strip(), "Original cell": ", ".join(r["cells"]), "Original": r["orig"], "Rebuilt from TB": r["auto"],
                       "Difference": r["var"], "Effect on gap": eff, "Why it differs": "; ".join(why),
                       "Evidence": _evidence(d, unit), "Audit findings on this line": int(len(hit))})
    bdf = pd.DataFrame(bridge)
    explained = float(bdf["Effect on gap"].sum()) if len(bdf) else 0.0
    ar = rc["arithmetic"]
    summary = {"gap": rc["gap"], "line_effects": explained, "arith_assets": ar.get("A"), "arith_el": ar.get("EL"),
               "unmatched_effect": rc.get("unmatched_effect", 0.0)}
    return bdf, summary


def reasons_table(rc, findings: pd.DataFrame, unit="Lakhs"):
    """one table: line-level reasons first, then each audit finding, with effect on the gap"""
    bdf, summary = gap_reason_tables(rc, findings, unit)
    rows = []
    for _, b in bdf.iterrows():
        rows.append({"Type": "Line differs from trial balance", "Where": b["Original cell"], "Item": b["Statement line"],
                     "Original": b["Original"], "Rebuilt": b["Rebuilt from TB"], "Amount": b["Difference"], "Effect on gap": b["Effect on gap"],
                     "Reason": b["Why it differs"], "Evidence": b["Evidence"]})
    ar = summary
    for k, nm in (("arith_assets", "Total Assets"), ("arith_el", "Total Equity & Liabilities")):
        v = ar.get(k)
        if v is not None and abs(v) > rc["tol"]:
            rows.append({"Type": "Total is not the sum of its lines", "Where": "", "Item": nm, "Original": None, "Rebuilt": None, "Amount": v,
                         "Effect on gap": (-v if k == "arith_assets" else v), "Reason": "the total in the original differs from the sum of the lines above it",
                         "Evidence": ""})
    if len(findings):
        for _, f in findings.iterrows():
            rows.append({"Type": f["Check"], "Where": f"{f['Sheet']}!{f['Cell']}", "Item": f["Item"], "Original": None, "Rebuilt": None,
                         "Amount": f["Amount"], "Effect on gap": f["Effect on gap"],
                         "Reason": f["Detail"] + (f"  [affects: {f['Affects']}]" if f["Affects"] else ""), "Evidence": f["Formula"]})
    return pd.DataFrame(rows), summary


DOUBLE_CHECKS = ("Same cell feeds two statement lines", "Ledger read by two lines of the same statement", "Cell referenced twice in one formula",
                 "Same link used twice inside a total", "Subtotal added together with its own rows")
SUM_CHECKS = ("Subtotal skips a row", "Total is typed, not a formula", "Total built differently from the other period columns", "Sheet could not be audited")


def audit_to_excel(rc, res, findings, reasons, bridge) -> bytes:
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        reasons.to_excel(xw, sheet_name="Gap reasons", index=False)
        bridge.to_excel(xw, sheet_name="Gap bridge (lines)", index=False)
        if len(findings):
            findings[findings["Check"].isin(DOUBLE_CHECKS)].to_excel(xw, sheet_name="Double links", index=False)
            findings[findings["Check"].isin(SUM_CHECKS)].to_excel(xw, sheet_name="Missing & wrong sums", index=False)
        pd.DataFrame([{"Line": r["label"].strip(), "Original": r["orig"], "Rebuilt from TB": r["auto"], "Difference": r["var"]} for r in rc["rows"]]
                     ).to_excel(xw, sheet_name="Line comparison", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(80, max(12, max(len(str(c.value or "")) for c in col[:80]) + 2))
    return buf.getvalue()
