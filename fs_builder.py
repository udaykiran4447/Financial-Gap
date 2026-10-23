"""
fs_builder.py  -  Build financial statements automatically from a trial balance.

Pipeline
  1. read_tb()      read a trial balance (Tally / ERP export, xlsx / csv) -> one row per ledger
  2. auto_group()   assign every ledger to a statement line (+ sub group) using accounting knowledge
  3. build()        post closing stock + manual adjustments, then compute Balance Sheet, P&L and notes
  4. to_excel()     write the statements in the same layout as the reference workbook (BS / PL / Note-* sheets,
                    live SUMIFS formulas into a 'TB Mapped' sheet, so a regrouped ledger flows through)
  5. reconcile()    compare the automated statements with an existing (non-tallying) workbook and point at
                    the lines / ledgers that cause the difference

Sign convention everywhere: net = debit - credit.  Statement values are shown in natural sign
(assets = Dr - Cr, liabilities / equity / income = Cr - Dr, expenses = Dr - Cr).
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import itertools
import re
from bisect import bisect_left
from difflib import SequenceMatcher

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

UNITS = {"Rupees": 1.0, "Thousands": 1e3, "Lakhs": 1e5, "Crores": 1e7, "Millions": 1e6}

# --------------------------------------------------------------------------------------------- chart of statements
# id, statement, label, note, nature (A asset, L liability/equity, I income, E expense), section
BS_LINES = [
    # id,        label,                                      note, nature, section
    ("PPE",      "(a) Property plant and Equipment",           2,   "A", "NCA"),
    ("CWIP",     "(b) Capital work  in Progress",              None, "A", "NCA"),
    ("ROU",      "(c) Right of Use Asset",                     3,   "A", "NCA"),
    ("INVPROP",  "(d) Investment Property",                    4,   "A", "NCA"),
    ("INVEST",   "  (i) Investments",                          5,   "A", "NCA"),
    ("LOANS_NC", "  (ii) Loans",                               6,   "A", "NCA"),
    ("OFA_NC",   "   (iii) Other Financial assets",            7,   "A", "NCA"),
    ("ONCA",     "(f) other Non current asset",                None, "A", "NCA"),
    ("DTA",      "(g) Deferred Tax Assets ( Net) ",            8,   "A", "NCA"),
    ("INVENT",   "(a) Inventories",                            9,   "A", "CA"),
    ("TR",       "(i) Trade  receivables",                     10,  "A", "CA"),
    ("CASH",     "(ii) Cash and cash equivalent",              11,  "A", "CA"),
    ("BANK",     "(iii) Bank Balance other than (ii) above",   12,  "A", "CA"),
    ("OFA_C",    "(iv)Other financial assets",                 13,  "A", "CA"),
    ("OCA",      "(c) Other current assets",                   14,  "A", "CA"),
    ("SC",       "(a) Equity Share Capital",                   15,  "L", "EQ"),
    ("OE",       "(b) Other Equity",                           16,  "L", "EQ"),
    ("BORR_NC",  "(i)   Borrowings",                           17,  "L", "NCL"),
    ("LEASE_NC", "(ii)  Lease Liabilities",                    18,  "L", "NCL"),
    ("OFL_NC",   "(iii)  Other Financial Liabilities",         None, "L", "NCL"),
    ("PROV_NC",  "(b) Provisions",                             19,  "L", "NCL"),
    ("BORR_C",   "i)    Borrowings",                           20,  "L", "CL"),
    ("LEASE_C",  "ii)  Lease Liabilities",                     21,  "L", "CL"),
    ("TP_MSME",  " a) Total outstanding due of Micro enterprises and small enterprises", 22, "L", "CL"),
    ("TP_OTH",   " b) Total Outstanding dues of Creditors other than Micro enterprises and small enterprises", 22, "L", "CL"),
    ("OFL_C",    "iv)   Other financial liabilities",          23,  "L", "CL"),
    ("OCL",      "(b) Other current liabilities",              24,  "L", "CL"),
    ("PROV_C",   "(c) Provisions",                             25,  "L", "CL"),
]
PL_LINES = [
    ("REV",   "Revenue from operations ",                      26, "I"),
    ("OI",    "Other income",                                  27, "I"),
    ("COGS",  "Cost of materials consumed",                    28, "E"),
    ("EMP",   "Employee benefits expense",                     29, "E"),
    ("FIN",   "Finance costs",                                 30, "E"),
    ("DEP",   "Depreciation and amortization expense",         31, "E"),
    ("OEXP",  "Other expenses",                                32, "E"),
    ("EXC",   "Exceptional items",                             None, "E"),
    ("CTAX",  "(1) Current tax",                               None, "E"),
    ("DTAX",  "(2) Deferred tax",                              None, "E"),
    ("OCI",   "(a) (i) Items that will not be reclassified to profit or loss", None, "I"),
]
SPECIAL = [
    ("SUSP",   "Inter-branch / suspense (should net to nil)", None, "S"),
    ("IGNORE", "Ignore (not part of statements)",             None, "S"),
]
LINE = {}
for _i, _l, _n, _nat, _sec in BS_LINES:
    LINE[_i] = {"id": _i, "label": _l.strip(), "note": _n, "nature": _nat, "stmt": "BS", "section": _sec}
for _i, _l, _n, _nat in PL_LINES:
    LINE[_i] = {"id": _i, "label": _l.strip(), "note": _n, "nature": _nat, "stmt": "PL", "section": "PL"}
for _i, _l, _n, _nat in SPECIAL:
    LINE[_i] = {"id": _i, "label": _l, "note": _n, "nature": _nat, "stmt": "X", "section": "X"}


def line_choices():
    """Display names used in the grouping editor:  'PPE  |  (a) Property plant and Equipment  [BS note 2]'"""
    out = []
    for lid, d in LINE.items():
        tag = {"BS": "Balance sheet", "PL": "P&L", "X": "Other"}[d["stmt"]]
        sec = {"NCA": "Non-current assets", "CA": "Current assets", "EQ": "Equity", "NCL": "Non-current liabilities",
               "CL": "Current liabilities", "PL": "P&L", "X": ""}[d["section"]]
        out.append(f"{lid} | {d['label']} ({sec or tag})")
    return out


def choice_to_id(txt):
    return str(txt).split("|")[0].strip()


def id_to_choice(lid):
    for c in line_choices():
        if choice_to_id(c) == lid:
            return c
    return line_choices()[0]


# --------------------------------------------------------------------------------------------- small helpers
def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower().replace("&", " and ")).strip()


def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.replace(",", "").replace("\u00a0", "").strip()
        if t.endswith(("Dr", "dr", "DR")):
            t = t[:-2].strip()
        if t.endswith(("Cr", "cr", "CR")):
            try:
                return -float(t[:-2].strip())
            except ValueError:
                return None
        try:
            return float(t)
        except ValueError:
            return None
    return None


RE_DATE = re.compile(r"(\d{1,2})[-./ ]([A-Za-z]{3,9}|\d{1,2})[-./ ,]+(\d{2,4})")
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def parse_dates(text):
    out = []
    for d, m, y in RE_DATE.findall(str(text)):
        try:
            mo = int(m) if m.isdigit() else MONTHS.get(m[:3].lower())
            yy = int(y)
            yy += 2000 if yy < 100 else 0
            if mo:
                out.append(dt.date(yy, mo, int(d)))
        except ValueError:
            continue
    return out


# --------------------------------------------------------------------------------------------- 1. read the trial balance
KNOWN_GROUPS = {
    "capital account", "reserves surplus", "reserves and surplus", "loans liability", "secured loans", "unsecured loans",
    "bank accounts", "bank od a c", "bank occ a c", "current liabilities", "duties taxes", "provisions", "sundry creditors",
    "fixed assets", "investments", "current assets", "stock in hand", "deposits asset", "loans advances asset", "sundry debtors",
    "cash in hand", "misc expenses asset", "branch divisions", "suspense a c", "sales accounts", "purchase accounts",
    "direct incomes", "indirect incomes", "direct expenses", "indirect expenses", "income direct", "income indirect",
    "expenses direct", "expenses indirect", "depreciation", "employee benefits", "other employee benefits", "finance expenses",
    "income tax", "vehicle maintenance", "inter corporate loans", "other current assets", "control accounts", "revenue accounts",
    "reserves", "capital", "liabilities", "assets", "income", "expenses", "current assets", "loans and advances",
}
KNOWN_GROUPS |= {re.sub(r"\s+", " ", k.replace(" and ", " ")).strip() for k in KNOWN_GROUPS} | {"duties and taxes", "loans and advances asset", "reserves and surplus"}
PARENT_NATURE = {
    "capital account": "L", "reserves surplus": "L", "duties and taxes": "L", "loans and advances asset": "A", "reserves and surplus": "L", "reserves": "L", "capital": "L",
    "loans liability": "L", "secured loans": "L", "unsecured loans": "L", "bank od a c": "L", "bank occ a c": "L",
    "current liabilities": "L", "duties taxes": "L", "duties and taxes": "L", "provisions": "L", "sundry creditors": "L", "liabilities": "L",
    "fixed assets": "A", "investments": "A", "current assets": "A", "stock in hand": "A", "deposits asset": "A",
    "loans advances asset": "A", "loans and advances": "A", "sundry debtors": "A", "cash in hand": "A", "bank accounts": "A",
    "misc expenses asset": "A", "inter corporate loans": "A", "other current assets": "A", "assets": "A",
    "branch divisions": "S", "suspense a c": "S", "control accounts": "A",
    "sales accounts": "I", "direct incomes": "I", "indirect incomes": "I", "income direct": "I", "income indirect": "I",
    "income": "I", "revenue accounts": "I",
    "purchase accounts": "E", "direct expenses": "E", "indirect expenses": "E", "expenses direct": "E",
    "expenses indirect": "E", "depreciation": "E", "employee benefits": "E", "other employee benefits": "E",
    "finance expenses": "E", "income tax": "E", "vehicle maintenance": "E", "expenses": "E",
}


def _cell_text(v):
    return clean(v) if isinstance(v, str) else ("" if v is None else str(v))


def _sheet_grid(data: bytes, name: str, sheet=None, max_rows=3000, max_cols=80):
    """returns (grid(list of lists of values), sheetnames, chosen sheet)"""
    if name.lower().endswith(".csv"):
        txt = data.decode("utf-8-sig", errors="replace")
        rows = list(csv.reader(io.StringIO(txt)))
        grid = [[(c if c != "" else None) for c in r] for r in rows[:max_rows]]
        return grid, ["csv"], "csv"
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    names = wb.sheetnames
    sh = sheet if sheet in names else None
    if sh is None:
        scored = []
        for n in names:
            nl = n.lower()
            sc = 0
            if re.search(r"(^|[^a-z])tb([^a-z]|$)|trial", nl):
                sc += 2
            if wb[n].sheet_state == "visible":
                sc += 1
            scored.append((sc, n))
        sh = max(scored, key=lambda x: x[0])[1] if scored else names[0]
    ws = wb[sh]
    grid = []
    for row in ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols, values_only=True):
        grid.append(list(row))
    wb.close()
    return grid, names, sh


def tb_sheet_candidates(data: bytes, name: str):
    """sheet names ranked by how much they look like a trial balance"""
    if name.lower().endswith(".csv"):
        return ["csv"]
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    res = []
    for n in wb.sheetnames:
        ws = wb[n]
        txt = []
        for row in ws.iter_rows(min_row=1, max_row=12, max_col=40, values_only=True):
            txt += [v.lower() for v in row if isinstance(v, str)]
        j = " | ".join(txt)
        ds = []
        for t in txt[:12]:
            ds += parse_dates(t)
        dscore = max(0.0, (max(ds).toordinal() - 726000) / 60.0) if ds else 0.0   # later period wins
        sc = dscore + (3 if "trial balance" in j else 0) + (2 if ("debit" in j and "credit" in j) else 0) \
            + (1 if re.search(r"(^|[^a-z])tb([^a-z]|$)", n.lower()) else 0) + (1 if ws.sheet_state == "visible" else 0)
        res.append((sc, n))
    wb.close()
    res.sort(key=lambda x: -x[0])
    return [n for _, n in res]


def detect_layout(grid):
    """find header row, Dr/Cr column pairs, label / code / group columns, title info"""
    hdr = None
    pairs = []
    for r, row in enumerate(grid[:40]):
        cur = []
        for c, v in enumerate(row):
            t = norm(v)
            if t in ("debit", "dr", "debit amount", "closing debit") and c + 1 < len(row) and norm(row[c + 1]) in ("credit", "cr", "credit amount", "closing credit"):
                cur.append(c)
        if cur:
            hdr, pairs = r, cur
            break
    if hdr is None:
        return None
    out_pairs = []
    for c in pairs:
        ctx = []
        for rr in range(max(0, hdr - 4), hdr):
            v = grid[rr][c] if c < len(grid[rr]) else None
            if isinstance(v, str) and clean(v):
                ctx.append(clean(v))
        # also include a header from the cell to the left in the same rows (merged-cell style titles)
        out_pairs.append({"dr": c, "cr": c + 1, "ctx": " / ".join(ctx)})
    first_pair = pairs[0]
    # label column
    label_col = None
    for r in range(0, hdr + 1):
        for c in range(0, first_pair):
            t = norm(grid[r][c] if c < len(grid[r]) else None)
            if t in ("particulars", "ledger", "ledger name", "account", "account name", "name", "description", "name of ledger"):
                label_col = c
    body = grid[hdr + 1: hdr + 400]
    if label_col is None:
        best, bc = 0, 0
        for c in range(0, first_pair):
            n = sum(1 for row in body if c < len(row) and isinstance(row[c], str) and not num(row[c]) and len(clean(row[c])) > 2)
            if n > best:
                best, bc = n, c
        label_col = bc
    # does the pair balance?  (rows named 'closing stock' are manual entries outside the trial balance and are left out of the test)
    for p in out_pairs:
        drs = crs = 0.0
        for row in grid[hdr + 1:]:
            if label_col < len(row) and re.search(r"closing stock|stock at end", norm(row[label_col])):
                continue
            d = num(row[p["dr"]]) if p["dr"] < len(row) else None
            k = num(row[p["cr"]]) if p["cr"] < len(row) else None
            drs += d or 0.0
            crs += k or 0.0
        p["dr_sum"], p["cr_sum"] = drs, crs
        p["balanced"] = bool(abs(drs - crs) <= 1e-6 * max(abs(drs), abs(crs), 1.0) + 0.5 and (drs or crs))
    code_col = group_col = None
    for r in range(0, hdr + 1):
        for c in range(0, first_pair):
            t = norm(grid[r][c] if c < len(grid[r]) else None)
            if c == label_col:
                continue
            if t in ("code", "ledger code", "account code", "a c code") and code_col is None:
                code_col = c
            if t in ("financial head", "group", "primary group", "group name", "head", "under") and group_col is None:
                group_col = c
    # title info
    top = [clean(v) for row in grid[:8] for v in row[:12] if isinstance(v, str) and clean(v)]
    company = ""
    for t in top:
        if re.search(r"(limited|ltd|llp|pvt|private|company|corporation|inc\b)", t, re.I) and len(t) < 90:
            company = t
            break
    dates = []
    for t in top:
        dates += parse_dates(t)
    start = end = None
    if dates:
        end = max(dates)
        start = min(dates) if len(dates) > 1 else None
    unit_hint = None
    for t in top + [p["ctx"] for p in out_pairs]:
        m = re.search(r"in\s+(lakh|lac|thousand|crore|million|rupee)", t, re.I)
        if m:
            unit_hint = {"lakh": "Lakhs", "lac": "Lakhs", "thousand": "Thousands", "crore": "Crores", "million": "Millions", "rupee": "Rupees"}[m.group(1).lower()]
    return {"hdr": hdr, "pairs": out_pairs, "label": label_col, "code": code_col, "group": group_col,
            "company": company, "start": start, "end": end, "unit_hint": unit_hint}


def default_pair(layout):
    """prefer a closing-balance pair that balances (Debit total = Credit total), then any balanced pair, then the first closing pair"""
    ps = layout["pairs"]
    for i, p in enumerate(ps):
        if "closing" in p["ctx"].lower() and p.get("balanced"):
            return i
    for i, p in enumerate(ps):
        if p.get("balanced"):
            return i
    for i, p in enumerate(ps):
        if "closing" in p["ctx"].lower():
            return i
    return 0


def pair_label(p):
    letters = f"{get_column_letter(p['dr'] + 1)} / {get_column_letter(p['cr'] + 1)}"
    tag = "balanced" if p.get("balanced") else "does not balance"
    return f"Columns {letters}  -  {p['ctx'] or 'no heading'}  [{tag}]"


def read_tb(data: bytes, name: str, sheet=None, pair=None, label_col=None, input_unit="Rupees"):
    """-> (DataFrame, info).  Columns: Row, Code, Ledger, Parent, FileGroup, Debit, Credit  (in input unit -> converted later)"""
    grid, names, sh = _sheet_grid(data, name, sheet)
    lay = detect_layout(grid)
    if lay is None:
        raise ValueError("Could not find 'Debit' and 'Credit' column headings in the sheet. Pick another sheet, or make sure the "
                         "trial balance has Debit and Credit columns side by side.")
    pi = default_pair(lay) if pair is None else pair
    p = lay["pairs"][pi]
    lc = lay["label"] if label_col is None else label_col
    rows = []
    for i in range(lay["hdr"] + 1, len(grid)):
        row = grid[i]
        lab = _cell_text(row[lc]).replace("_x000D_", "").strip() if lc < len(row) else ""
        dr = num(row[p["dr"]]) if p["dr"] < len(row) else None
        cr = num(row[p["cr"]]) if p["cr"] < len(row) else None
        code = row[lay["code"]] if lay["code"] is not None and lay["code"] < len(row) else None
        fg = _cell_text(row[lay["group"]]).replace("_x000D_", "").strip() if lay["group"] is not None and lay["group"] < len(row) else ""
        rows.append({"Row": i + 1, "Code": code, "Ledger": lab, "FileGroup": fg, "Debit": dr or 0.0, "Credit": cr or 0.0,
                     "_has_amt": (dr is not None) or (cr is not None)})
    df = pd.DataFrame(rows)
    # drop empty lines and the grand-total line
    df["_n"] = df["Ledger"].map(norm)
    tot_mask = df["_n"].isin({"grand total", "total", "totals"}) | ((df["_n"] == "") & df["_has_amt"] & (df["Debit"] != 0) & (df["Credit"] != 0)
                                                                    & ((df["Debit"] - df["Credit"]).abs() < 0.5 + 1e-9 * df["Debit"].abs()))
    df = df[~tot_mask & (df["_n"] != "")].copy()
    # heading rows (group names with zero or sub-total amounts)
    parents, is_head = [], []
    cur_parent = ""
    recs = df.to_dict("records")
    for k, rec in enumerate(recs):
        n = rec["_n"]
        head = False
        if n in KNOWN_GROUPS:
            net = rec["Debit"] - rec["Credit"]
            if abs(net) < 1e-9 and rec["Debit"] == rec["Credit"]:
                head = True
            else:
                s, j = 0.0, k + 1
                while j < len(recs) and recs[j]["_n"] not in KNOWN_GROUPS:
                    s += recs[j]["Debit"] - recs[j]["Credit"]
                    j += 1
                head = abs(s - net) < max(0.5, abs(net) * 1e-6) and j > k + 1
        if head:
            cur_parent = n
        parents.append(cur_parent)
        is_head.append(head)
    df["Parent"] = parents
    df["_head"] = is_head
    df = df[~df["_head"]].copy()
    df["Debit"] = df["Debit"].astype(float)
    df["Credit"] = df["Credit"].astype(float)
    df = df[["Row", "Code", "Ledger", "FileGroup", "Parent", "Debit", "Credit"]].reset_index(drop=True)
    info = {"sheet": sh, "sheets": names, "layout": lay, "pair_index": pi, "label_col": lc,
            "unit_guess": lay["unit_hint"] if lay["pairs"][pi]["ctx"] and lay["unit_hint"] and "lakh" in p["ctx"].lower() else "Rupees"}
    return df, info


# --------------------------------------------------------------------------------------------- 2. automatic grouping
# (regex on lowercased ledger text, line id, sub group, nature restriction)
# order matters: first match wins.  nature: A/L/I/E restricts the rule to that class of ledger (None = any)
RULES = [
    # ---- special
    (r"closing stock|stock at end|closing inventory", "INVENT", "Closing stock (manual)", None),
    (r"opening stock|stock at beginning|opening inventory", "COGS", "Opening stock", None),
    (r"inter ?(branch|unit|division)|branch|division|elimination|head office a c|ho a c", "SUSP", "Inter-branch", "S"),
    # ---- equity
    (r"share capital|equity share|paid up|capital a c$|^capital$|authori[sz]ed capital", "SC", "Equity share capital", "L"),
    (r"^profit and loss( a c)?$|^p ?and ?l( a c)?$|retained earning|surplus in (the )?statement|profit loss a c", "OE", "Profit & Loss A/c (opening)", None),
    (r"reserve|surplus|securities premium|share premium|general reserve|capital redemption|revaluation", "OE", "Reserves & Surplus", "L"),
    # ---- other comprehensive income
    (r"other comprehensive|\boci\b|actuarial (gain|loss)|re.?measurement", "OCI", "Actuarial gain/(loss) on post-employment benefits", None),
    # ---- tax
    (r"deferred tax (asset|liab)|^dta$|^dtl$", "DTA", "Deferred tax asset", "A"),
    (r"def+er+ed tax.*(p ?and ?l|expense|income)|deffer?red tax p ?l", "DTAX", "Deferred tax", None),
    (r"^income tax$|current tax|tax expense|provision for taxation.*(p ?l|expense)", "CTAX", "Current tax", "E"),
    (r"provision for (income )?tax|income tax provision|tax provision|provision for taxation", "PROV_C", "Provision for income tax", "L"),
    (r"advance tax|self assessment|tds receivable|tds recoverable|tcs receivable|income tax refund|input (tax )?credit|gst input|input (igst|cgst|sgst|gst)|gst receivable|gst recoverable|balance with (govt|government)|gst itc", "OCA", "Balance with Govt authorities", "A"),
    # ---- borrowings / leases
    (r"lease liabilit", "LEASE_NC", "Lease liabilities", "L"),
    (r"bank o ?d|overdraft|cash credit|\bcc a c|\bocc\b|working capital (loan|demand)|\bwcdl\b|\bod a c|\bod\b|seasonal credit", "BORR_C", "Secured - bank OD / CC", "L"),
    (r"current maturit", "BORR_C", "Current maturities of long-term borrowings", "L"),
    (r"unsecured loan|loan from (director|promoter|related|shareholder)|inter ?corporate (deposit|borrow)|director loan|loan from (others|friends)", "BORR_NC", "Unsecured", "L"),
    (r"secured loan|term loan|vehicle loan|car loan|loan from (bank|nbfc|financial)|\bnbfc\b|hdfc|icici|axis bank|sbi (loan|term)|loan a c", "BORR_NC", "Secured", "L"),
    # ---- fixed assets
    (r"provision (for|of) depreciation|accumulated depreciation|accumulated amorti[sz]ation|depreciation reserve", "PPE", "Accumulated depreciation", "A"),
    (r"right of use|rou asset|leased asset|lease asset|leasehold asset", "ROU", "Right of use assets", "A"),
    (r"capital work|\bcwip\b|assets under construction", "CWIP", "Capital work in progress", "A"),
    (r"investment propert|land development|investment in (land|propert)", "INVPROP", "Investment property", "A"),
    (r"investment in|investments? (in )?(shares|equity|mutual|bond)|mutual fund|equity shares of|subsidiar(y|ies) (shares|equity)", "INVEST", "Investment in subsidiaries / others", "A"),
    # ---- current assets
    (r"sundry debtor|trade receivable|debtors|accounts receivable|nats receivable|unbilled|receivable from customer", "TR", "Trade receivables", "A"),
    (r"expected credit loss|provision for (doubtful|bad|credit)|\becl\b|allowance for (doubtful|credit)", "TR", "Expected credit loss", None),
    (r"cash in hand|petty cash|^cash$|cash balance", "CASH", "Cash in hand", "A"),
    (r"fixed deposit|\bfdr?s?\b|term deposit|deposits? with bank|margin money|lien", "BANK", "Fixed deposits with banks", "A"),
    (r"bank guarantee|\bbg\b deposit", "OFA_NC", "Deposits against bank guarantees", "A"),
    (r"accrued interest|interest accrued|interest receivable", "OFA_NC", "Interest accrued on deposits", "A"),
    (r"\bemd\b|earnest money|security deposit|rental deposit|rent deposit|deposits? \(asset\)|other deposit|utility deposit|tender deposit|lease deposit|deposits?$", "OFA_C", "Security & other deposits", "A"),
    (r"inter ?corporate loan|loan to (subsidiar|group|related|associate)|loans? (and|&) advances to (related|subsidiar)|loan given", "LOANS_NC", "Loans to subsidiaries / related parties", "A"),
    (r"current account|bank account|\bbank\b|\bsbi\b|\bhdfc\b|\bicici\b|\baxis\b|\bkotak\b|\byes bank\b|\bidfc\b|\bcanara\b|\bpnb\b|\bbob\b|\bunion bank", "CASH", "Balances with banks", "A"),
    (r"skill project|project (receivable|debtor)|accrued income|unbilled project", "OFA_C", "Receivables from projects", "A"),
    (r"control a c|control account|difference in opening", "TR", "Control accounts (unreconciled receipts / adjustments)", "A"),
    (r"prepaid|advance to (supplier|vendor)|staff advance|employee advance|loan to employee|salary advance|advance (paid|recoverable)|other advance|imprest|project advance|work in progress|\bwip\b", "OCA", "Advances & prepaid expenses", "A"),
    (r"stock in trade|stock in hand|inventor|consumables? stock|uniform stock|raw material stock", "INVENT", "Inventories", "A"),
    # ---- liabilities
    (r"gratuity (payable|provision)|provision for gratuity|leave encashment|provision for (leave|employee|bonus)|compensated absence", "PROV_NC", "Provision for employee benefits", "L"),
    (r"sundry creditor|trade payable|creditors?$|creditor", "TP_OTH", "Sundry creditors", "L"),
    (r"micro|msme (payable|creditor)", "TP_MSME", "MSME creditors", "L"),
    (r"duties (and|&) tax|duties$|gst (payable|output)|output (igst|cgst|sgst|gst)|tds payable|tcs payable|statutory|pf payable|esi payable|ptax payable|professional tax payable|tax deducted|labou?r welfare fund payable|pf (and|&) esi payable|epf payable|esic payable", "OCL", "Statutory dues payable", "L"),
    (r"advance (from|received from) customer|customer advance|unearned|deferred (revenue|income)|advance received|received in advance|security deposit (received|payable)|retention (money )?payable", "OCL", "Advances from customers / deferred income", "L"),
    (r"salar(y|ies) payable|wages payable|staff (allowance|payable)|remuneration payable|bonus payable|payable to (employee|staff)", "OFL_C", "Employee related payables", "L"),
    (r"interest payable|interest accrued and due|interest accrued but not due", "OFL_C", "Interest payable", "L"),
    (r"audit fee payable|expense payable|expenses payable|outstanding expense|accrued expense|accounts payable|provision for expenses|payable|outstanding", "OFL_C", "Expenses payable", "L"),
    (r"provision for|^provisions?$", "PROV_C", "Other provisions", "L"),
    (r"loan|borrowing|debenture", "BORR_NC", "Other borrowings", "L"),
    # ---- fixed assets (name based; only for non expense ledgers)
    (r"land|building|plant|machiner|furniture|fixture|computer|printer|vehicle|motor (car|cycle)|\bcar\b|equipment|air ?condition|inverter|mobile|live ?stock|lease ?hold improvement|electrical (fitting|installation)|laptop|server|generator|cctv|tools|office equip|software|website|trademark|goodwill|intangible", "PPE", "Property, plant & equipment", "A"),
    # ---- income
    (r"sales|service income|housekeeping income|house keeping income|security service income|contract staffing income|staffing income|labou?r supply|revenue|fee income|commission (received|income)|job work income|consulting income|grant|turnover|unbilled", "REV", "Sale of services", "I"),
    (r"interest (income|received|on (fixed deposit|fd|intercorporate|inter corporate|loan|income tax refund|deposits?))|dividend income|discount received|misc(ellaneous)? income|other income|profit on sale|gain on|income from naps|rental income|scrap sale|commission received|insurance claim|bad debts recovered|sundry balance(s)? written back|liabilities? written back|rebate", "OI", "Other income", "I"),
    # ---- expenses
    (r"opening stock", "COGS", "Opening stock", "E"),
    (r"purchase|consumable|material|uniform|stores|packing material|raw material|outsourc|machinery on rental|equipment and tools|site expense|freight inward|carriage inward|hire charges|subcontract|sub contract|labou?r charges|machine rent", "COGS", "Purchases", "E"),
    (r"depreciation|amorti[sz]ation", "DEP", "Depreciation", "E"),
    (r"interest (on|paid|expense)|bank charges|bank commission|processing fee|finance (charge|cost|expense)|lease (interest|expense)|interset|loan processing|foreclosure|brokerage on loan|guarantee commission|lc charges|forex loss", "FIN", "Finance costs", "E"),
    (r"salar|wages|bonus|gratuity|\bpf\b|provident|\besi\b|\besic\b|\bepf\b|staff welfare|employee|labou?r welfare|director.*remuneration|remuneration|incentive|stipend|leave encashment|local conveyance|f ?and ?f settle|additional sites|security.*service expend|hk service|sg and hk|referral|referal|apprentice|ex gratia|medical (expense|reimb)", "EMP", "Employee benefits", "E"),
    (r"income tax|tax paid|\btds\b.*expense|provision for tax", "CTAX", "Current tax", "E"),
    (r".", "OEXP", "Other expenses", "E"),
]
RULES = [(re.compile(p), line, sub, nat) for p, line, sub, nat in RULES]
EXPENSE_WORDS = re.compile(r"expense|exps|expenditure|charges|maintenance|repair|rental|\brent\b|insurance|fees?\b|printing|stationery|postage|telephone|internet|travel|advertis|promotion|donation|audit|professional|consultan|electricity|conveyance|office|subscription|software exp|tender|licen[cs]e|courier|security charges|freight|loss on|written off|bad debt|penalt|rates|taxes?$|general exp|miscellaneous exp", re.I)

NATURE_OF_PL = {"I", "E"}


PARENT_DEFAULT = {"fixed assets": ("PPE", "Property, plant & equipment"), "investments": ("INVEST", "Investment in subsidiaries / others"),
                  "provisions": ("PROV_C", "Other provisions"), "control accounts": ("TR", "Control accounts"),
                  "sundry debtors": ("TR", "Trade receivables"), "sundry creditors": ("TP_OTH", "Sundry creditors"),
                  "duties and taxes": ("OCL", "Statutory dues payable"), "duties taxes": ("OCL", "Statutory dues payable"),
                  "inter corporate loans": ("LOANS_NC", "Loans to subsidiaries / related parties")}


def _nature_from(row):
    """A / L / I / E / S from the Tally parent group when known"""
    return PARENT_NATURE.get(row.get("Parent", ""), None)


def classify(ledger, parent="", filegroup="", debit=0.0, credit=0.0):
    """returns (line_id, sub_group, reason, confidence 0-1)"""
    text = norm(ledger)
    nat = PARENT_NATURE.get(parent)
    net = debit - credit
    exp_like = bool(EXPENSE_WORDS.search(ledger))
    # parent group decides balance sheet vs P&L; keyword decides the line
    allowed = None
    if nat in ("A", "L"):
        allowed = {nat, "S", None}
    elif nat in ("I", "E"):
        allowed = {"I", "E", None}
    elif nat == "S":
        allowed = {"S"}
    else:
        # no parent known: expense-looking names are P&L
        allowed = {"I", "E", None} if exp_like else None

    def run(txt, allow, conf, why):
        for rx, line, sub, rnat in RULES:
            if rnat is not None and allow is not None and rnat not in allow:
                continue
            if rx.search(txt):
                return line, sub, why, conf
        return None

    # a P&L parent whose ledger clearly is a balance-sheet item / vice versa: trust the strongest keyword
    r = run(text, allowed, 0.9 if nat else 0.8, f"keyword match" + (f" under '{parent}'" if nat else ""))
    if r and r[0] == "OEXP" and nat in ("A", "L"):
        r = None  # fallback rule must not capture a balance sheet ledger
    if r and nat in ("I", "E") and LINE[r[0]]["stmt"] == "BS" and exp_like:
        r = run(text, {"I", "E", None}, 0.7, "expense keyword")
    if r and nat in ("I", "E") and r[0] in ("REV", "OI") and net > 0 and exp_like:
        r = run(text, {"E"}, 0.6, "expense keyword with debit balance")
    if r is None and filegroup and norm(filegroup) not in ("group head", ""):
        r = run(norm(filegroup) + " " + text, allowed, 0.6, f"file group '{clean(filegroup)}'")
        if r and r[0] == "OEXP" and nat in ("A", "L"):
            r = None
    pd_line = PARENT_DEFAULT.get(parent)
    if r is None and pd_line:
        r = (pd_line[0], pd_line[1], f"default for ledgers under '{parent}'", 0.6)
    if r is None:
        if nat == "A":
            r = ("OCA", "Other current assets", f"default for assets under '{parent}'", 0.4)
        elif nat == "L":
            r = ("OCL", "Other current liabilities", f"default for liabilities under '{parent}'", 0.4)
        elif nat == "I":
            r = ("OI", "Other income", f"default for income under '{parent}'", 0.4)
        elif nat == "E":
            r = ("OEXP", "Other expenses", f"default for expenses under '{parent}'", 0.5)
        elif nat == "S":
            r = ("SUSP", "Inter-branch", "branch / division ledger", 0.8)
        else:
            r = ("OEXP", "Other expenses", "no rule matched - please review", 0.2) if net >= 0 else ("OI", "Other income", "no rule matched - please review", 0.2)
    line, sub, why, conf = r
    # sanity on balance: reduce confidence when the balance is on the unnatural side
    nat_line = LINE[line]["nature"]
    if abs(net) > 1e-9 and ((nat_line in ("A", "E") and net < 0) or (nat_line in ("L", "I") and net > 0)) and line not in ("OCI", "SUSP"):
        conf = min(conf, 0.55)
        why += " (balance on the unusual side)"
    return line, sub, why, round(conf, 2)


def pl_subgroup(line, sub, rec):
    """P&L notes list one row per head: use the file's own head (Financial Head) if it has one, else the ledger name"""
    if line in ("REV", "OI", "EMP", "OEXP", "FIN"):
        fg = clean(rec.get("FileGroup", ""))
        if fg and norm(fg) not in ("group head", "pl", "p and l"):
            return fg[:1].upper() + fg[1:]
        name = clean(rec.get("Ledger", ""))
        return name.title() if name.isupper() else name
    return sub


def auto_group(df: pd.DataFrame, memory: dict | None = None) -> pd.DataFrame:
    """adds Line, SubGroup, Reason, Confidence.  memory: {normalised ledger name: (line, sub)} from a saved mapping"""
    memory = memory or {}
    lines, subs, reasons, confs = [], [], [], []
    for rec in df.to_dict("records"):
        key = norm(rec["Ledger"])
        if key in memory:
            l, s = memory[key]
            lines.append(l)
            subs.append(s)
            reasons.append("saved mapping")
            confs.append(1.0)
            continue
        l, s, why, c = classify(rec["Ledger"], rec.get("Parent", ""), rec.get("FileGroup", ""), rec["Debit"], rec["Credit"])
        s = pl_subgroup(l, s, rec)
        lines.append(l)
        subs.append(s)
        reasons.append(why)
        confs.append(c)
    out = df.copy()
    out["Line"], out["SubGroup"], out["Reason"], out["Confidence"] = lines, subs, reasons, confs
    out["Include"] = True
    return out


def mapping_to_csv(df: pd.DataFrame) -> bytes:
    m = df[["Ledger", "Line", "SubGroup"]].copy()
    return m.to_csv(index=False).encode("utf-8")


def mapping_from_csv(data: bytes) -> dict:
    m = pd.read_csv(io.BytesIO(data))
    out = {}
    for _, r in m.iterrows():
        if r.get("Line") in LINE:
            out[norm(r["Ledger"])] = (r["Line"], str(r.get("SubGroup") or ""))
    return out


# --------------------------------------------------------------------------------------------- 3. build the statements
def default_adjustments():
    return pd.DataFrame({"Narration": pd.Series(dtype=str), "Line": pd.Series(dtype=str), "SubGroup": pd.Series(dtype=str),
                         "Debit": pd.Series(dtype=float), "Credit": pd.Series(dtype=float)})


def _canon_subs(post):
    """Excel SUMIFS is case-insensitive and treats * ? ~ and leading < > = specially, so sub group names are made safe and
    one spelling (first seen) is used for names that differ only by case."""
    seen, out = {}, []
    for line, sub in zip(post["Line"], post["SubGroup"]):
        t = re.sub(r"[*?~]", " ", str(sub or "")).replace('"', "'")
        t = re.sub(r"\s+", " ", t).strip()
        t = re.sub(r"^[<>=]+\s*", "", t)
        out.append(seen.setdefault((line, t.lower()), t))
    return out


def build(tb: pd.DataFrame, meta: dict, closing_stock: float = 0.0, adjustments: pd.DataFrame | None = None, unit: str = "Lakhs",
          input_unit: str = "Rupees", prev_bs: dict | None = None):
    """tb: grouped ledger frame (Include/Ledger/Debit/Credit/Line/SubGroup) in INPUT units.
    returns dict with posted ledger table 'post' (output units), 'bs', 'pl', 'notes', 'checks'."""
    f = UNITS[input_unit] / UNITS[unit]
    rows = []
    for rec in tb.to_dict("records"):
        if not rec.get("Include", True):
            continue
        rows.append({"Source": "Trial balance", "Ledger": rec["Ledger"], "Debit": rec["Debit"] * f, "Credit": rec["Credit"] * f,
                     "Line": rec["Line"], "SubGroup": rec["SubGroup"] or "", "TBRow": rec.get("Row")})
    if abs(closing_stock) > 1e-12:
        rows.append({"Source": "Manual", "Ledger": "Closing stock - to Inventories", "Debit": closing_stock, "Credit": 0.0, "Line": "INVENT",
                     "SubGroup": "Uniforms and housekeeping materials", "TBRow": None})
        rows.append({"Source": "Manual", "Ledger": "Closing stock - credit to cost of materials", "Debit": 0.0, "Credit": closing_stock,
                     "Line": "COGS", "SubGroup": "Closing stock", "TBRow": None})
    if adjustments is not None and len(adjustments):
        for rec in adjustments.to_dict("records"):
            if not rec.get("Line") or (not rec.get("Debit") and not rec.get("Credit")):
                continue
            rows.append({"Source": "Adjustment", "Ledger": rec.get("Narration") or "Adjustment", "Debit": float(rec.get("Debit") or 0),
                         "Credit": float(rec.get("Credit") or 0), "Line": choice_to_id(rec["Line"]), "SubGroup": rec.get("SubGroup") or "",
                         "TBRow": None})
    post = pd.DataFrame(rows, columns=["Source", "Ledger", "Debit", "Credit", "Line", "SubGroup", "TBRow"])
    post["SubGroup"] = _canon_subs(post)
    post["Net"] = post["Debit"] - post["Credit"]
    # move the manual closing stock row if the TB carries its own 'closing stock' ledger: handled by the caller (Include = False)
    net_by = post.groupby("Line")["Net"].sum().to_dict()

    def nat_val(lid):
        n = net_by.get(lid, 0.0)
        nat = LINE[lid]["nature"]
        return -n if nat in ("L", "I") else n

    # ---- P&L
    pl = {lid: nat_val(lid) for lid, *_ in PL_LINES}
    pl["INCOME"] = pl["REV"] + pl["OI"]
    pl["EXPENSES"] = sum(pl[k] for k in ("COGS", "EMP", "FIN", "DEP", "OEXP"))
    pl["PBT_EXC"] = pl["INCOME"] - pl["EXPENSES"]
    pl["PBT"] = pl["PBT_EXC"] - pl["EXC"]
    pl["NET"] = pl["PBT"] - pl["DTAX"] - pl["CTAX"]
    oci_tax = -pl["OCI"] * float(meta.get("oci_tax_rate", 0.0))
    pl["OCI_TAX"] = oci_tax
    pl["OCI_TOTAL"] = pl["OCI"] + oci_tax
    pl["TCI"] = pl["NET"] + pl["OCI_TOTAL"]
    # ---- equity roll-forward: profit and OCI are added to other equity
    oe_ledgers = nat_val("OE")
    net_by_oe = {s: -v for s, v in post[post["Line"] == "OE"].groupby("SubGroup")["Net"].sum().items()}
    # ---- Balance sheet
    bs = {}
    for lid, _l, _n, nat, _s in BS_LINES:
        bs[lid] = nat_val(lid)
    bs["OE"] = oe_ledgers + pl["TCI"]
    susp = net_by.get("SUSP", 0.0)
    susp_note = None
    if abs(susp) > 1e-6:
        susp_note = susp
        if susp > 0:
            bs["OCA"] += susp
        else:
            bs["OCL"] += -susp
    nca = sum(bs[k] for k, *_ , sec in [(l[0], l[4]) for l in BS_LINES] if False)  # placeholder to keep linters calm
    sec_tot = {}
    for lid, _l, _n, nat, sec in BS_LINES:
        sec_tot[sec] = sec_tot.get(sec, 0.0) + bs[lid]
    bs["TOT_NCA"], bs["TOT_CA"] = sec_tot["NCA"], sec_tot["CA"]
    bs["TOT_ASSETS"] = bs["TOT_NCA"] + bs["TOT_CA"]
    bs["TOT_EQ"], bs["TOT_NCL"], bs["TOT_CL"] = sec_tot["EQ"], sec_tot["NCL"], sec_tot["CL"]
    bs["TOT_EL"] = bs["TOT_EQ"] + bs["TOT_NCL"] + bs["TOT_CL"]
    # ---- EPS
    sc = bs["SC"]
    fv = float(meta.get("face_value") or 10.0)
    shares = float(meta.get("shares") or 0.0)
    if shares <= 0 and fv > 0:
        shares = sc * UNITS[unit] / fv
    pl["SHARES"] = shares
    pl["EPS"] = (pl["NET"] * UNITS[unit] / shares) if shares else 0.0
    checks = {
        "tb_debit": float(post[post["Source"] == "Trial balance"]["Debit"].sum()),
        "tb_credit": float(post[post["Source"] == "Trial balance"]["Credit"].sum()),
        "post_debit": float(post["Debit"].sum()), "post_credit": float(post["Credit"].sum()),
        "difference_bs": bs["TOT_ASSETS"] - bs["TOT_EL"],
        "suspense_net": susp_note,
    }
    checks["tb_diff"] = checks["tb_debit"] - checks["tb_credit"]
    checks["post_diff"] = checks["post_debit"] - checks["post_credit"]
    # ---- notes: sub group breakdown per line
    notes = {}
    for lid, d in LINE.items():
        if d["stmt"] == "X":
            continue
        sub = post[post["Line"] == lid].groupby("SubGroup", sort=False)["Net"].sum()
        sign = -1 if d["nature"] in ("L", "I") else 1
        items = [(s, sign * v) for s, v in sub.items()]
        notes[lid] = items
    if susp_note is not None:
        if susp_note > 0:
            notes["OCA"].append(("Inter-branch balances (net debit)", susp_note))
        else:
            notes["OCL"].append(("Inter-branch balances (net credit)", -susp_note))
    return {"post": post, "bs": bs, "pl": pl, "notes": notes, "checks": checks, "meta": meta, "unit": unit, "prev_bs": prev_bs,
            "oe_parts": net_by_oe, "net_by": net_by}


# --------------------------------------------------------------------------------------------- 4. Excel in the reference layout
TNR = "Times New Roman"
MED = Side(style="medium")
THICK = Side(style="thick")
NUMF = "#,##0.00"
ACC = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'


def _f(sz=12, b=False):
    return Font(name=TNR, size=sz, bold=b)


def fmt_date(d):
    return d.strftime("%d.%m.%Y") if isinstance(d, (dt.date, dt.datetime)) else str(d or "")


def to_excel(res: dict) -> bytes:
    meta, unit = res["meta"], res["unit"]
    post = res["post"]
    wb = Workbook()
    wsb = wb.active
    wsb.title = "BS"
    wsp = wb.create_sheet("PL")
    wsa = wb.create_sheet("Note-Assets")
    wsl = wb.create_sheet("Note-Liab")
    wsn = wb.create_sheet("PL Note")
    wst = wb.create_sheet("TB Mapped")
    wsc = wb.create_sheet("Checks")
    end = fmt_date(meta.get("end"))
    company = meta.get("company") or "Company"
    cin = meta.get("cin") or ""
    unit_txt = f"(All amount are in Indian Rupees in {unit})" if unit != "Rupees" else "(All amount are in Indian Rupees)"

    # ---------------- TB Mapped (source of every figure)
    hdr = ["Source", "Ledger", "Debit", "Credit", "Line", "Sub group", "Net (Dr-Cr)", "TB row"]
    for j, h in enumerate(hdr, 1):
        c = wst.cell(1, j, h)
        c.font, c.fill = _f(11, True), PatternFill("solid", fgColor="FAF6F1")
    for i, rec in enumerate(post.to_dict("records"), 2):
        wst.cell(i, 1, rec["Source"])
        wst.cell(i, 2, rec["Ledger"])
        wst.cell(i, 3, rec["Debit"]).number_format = NUMF
        wst.cell(i, 4, rec["Credit"]).number_format = NUMF
        wst.cell(i, 5, rec["Line"])
        wst.cell(i, 6, rec["SubGroup"])
        wst.cell(i, 7, f"=C{i}-D{i}").number_format = NUMF
        wst.cell(i, 8, rec["TBRow"])
    last = len(post) + 1
    wst.cell(last + 2, 2, "Total").font = _f(11, True)
    wst.cell(last + 2, 3, f"=SUM(C2:C{last})").number_format = NUMF
    wst.cell(last + 2, 4, f"=SUM(D2:D{last})").number_format = NUMF
    wst.cell(last + 2, 7, f"=SUM(G2:G{last})").number_format = NUMF
    for col, w in zip("ABCDEFGH", (14, 48, 16, 16, 12, 40, 16, 8)):
        wst.column_dimensions[col].width = w
    wst.freeze_panes = "A2"
    R = f"$2:${last}"
    NET, LNE, SUB = f"'TB Mapped'!$G$2:$G${last}", f"'TB Mapped'!$E$2:$E${last}", f"'TB Mapped'!$F$2:$F${last}"

    def sumifs(lid, sub=None, sign=1):
        if sub and str(sub).startswith("Inter-branch balances"):
            return f"={'-' if 'credit' in sub else ''}SUMIFS({NET},{LNE},\"SUSP\")"
        core = f"SUMIFS({NET},{LNE},\"{lid}\"" + (f",{SUB},\"{sub}\"" if sub is not None else "") + ")"
        return f"={'-' if sign < 0 else ''}{core}"

    # ---------------- notes
    note_cell = {}      # line id -> (sheet, cell of total)

    def write_notes(ws, lids, title_col, amt_col, code_col, start_row, header_text, pl_style=False):
        r = start_row
        by_note = {}
        for lid in lids:
            d = LINE[lid]
            by_note.setdefault(d["note"] if d["note"] else f"x{lid}", []).append(lid)
        for key, group in sorted(by_note.items(), key=lambda kv: (isinstance(kv[0], str), kv[0] if not isinstance(kv[0], str) else 0)):
            if all(not res["notes"].get(l) for l in group) and not any(l in ("OE",) for l in group):
                continue
            d0 = LINE[group[0]]
            ttl = d0["label"] if not isinstance(key, int) else f"Note No.{key} {re.sub(r'^[\\(\\)a-z0-9 ]+[\\)] *', '', d0['label']).strip() or d0['label']}"
            if isinstance(key, int):
                nm = {2: "Property, plant and equipment", 3: "Right of use assets", 4: "Investment property", 5: "Investments", 6: "Loans",
                      7: "Other financial assets", 8: "Deferred tax assets", 9: "Inventories", 10: "Trade receivables", 11: "Cash and cash equivalents",
                      12: "Bank balances other than cash and cash equivalents", 13: "Other financial assets", 14: "Other current assets",
                      15: "Equity share capital", 16: "Other equity", 17: "Long term borrowings", 18: "Lease liabilities", 19: "Provisions",
                      20: "Short-term borrowings", 21: "Lease liabilities", 22: "Trade payables", 23: "Other financial liabilities",
                      24: "Other current liabilities", 25: "Provisions", 26: "Revenue from operations", 27: "Other income",
                      28: "Cost of materials consumed", 29: "Employee benefits", 30: "Finance costs", 31: "Depreciation and amortization expense",
                      32: "Other expenses"}.get(key, d0["label"])
                ttl = f"Note No.{key} {nm}"
            ws.cell(r, title_col, ttl).font = _f(13, True)
            r += 1
            h1 = ws.cell(r, title_col, "Particulars")
            h2 = ws.cell(r, amt_col, header_text)
            for c in (h1, h2):
                c.font, c.alignment = _f(13, True), Alignment(horizontal="center", wrap_text=True)
                c.border = Border(top=MED, bottom=MED)
            r += 1
            first = r
            for lid in group:
                d = LINE[lid]
                items = res["notes"].get(lid, [])
                sign = -1 if d["nature"] in ("L", "I") else 1
                if lid == "OE":
                    parts = [("Reserves & Surplus", "Reserves & Surplus"), ("Profit & Loss A/c (opening)", "Profit & Loss A/c (opening)")]
                    seen = set()
                    for s, _v in items:
                        if s not in seen:
                            seen.add(s)
                            ws.cell(r, code_col, lid)
                            ws.cell(r, title_col, s or "Other reserves").font = _f(13)
                            ws.cell(r, amt_col, sumifs(lid, s, -1)).number_format = ACC
                            r += 1
                    ws.cell(r, title_col, "Profit / (loss) for the period").font = _f(13)
                    ws.cell(r, amt_col, "=PL!{}".format("{PLNET}")).number_format = ACC
                    note_cell["_OE_PROFIT"] = (ws, r, amt_col)
                    r += 1
                    ws.cell(r, title_col, "Other comprehensive income for the period").font = _f(13)
                    ws.cell(r, amt_col, "=PL!{}".format("{PLOCI}")).number_format = ACC
                    note_cell["_OE_OCI"] = (ws, r, amt_col)
                    r += 1
                    continue
                if len(group) > 1:
                    ws.cell(r, title_col, d["label"]).font = _f(13, True)
                    r += 1
                if not items:
                    ws.cell(r, code_col, lid)
                    ws.cell(r, title_col, d["label"]).font = _f(13)
                    ws.cell(r, amt_col, 0).number_format = ACC
                    r += 1
                for s, _v in items:
                    ws.cell(r, code_col, lid)
                    ws.cell(r, title_col, s or d["label"]).font = _f(13)
                    ws.cell(r, amt_col, sumifs(lid, s, sign)).number_format = ACC
                    ws.cell(r, amt_col).font = _f(12)
                    r += 1
            tot = ws.cell(r, title_col, "Total")
            tot.font, tot.alignment, tot.border = _f(13, True), Alignment(horizontal="center"), Border(top=MED, bottom=MED)
            tc = ws.cell(r, amt_col, f"=SUM({get_column_letter(amt_col)}{first}:{get_column_letter(amt_col)}{r - 1})")
            tc.font, tc.number_format, tc.border = _f(13, True), ACC, Border(top=MED, bottom=MED)
            for lid in group:
                note_cell[lid] = (ws, r, amt_col) if len(group) == 1 else note_cell.get(lid)
            if len(group) > 1:
                # sub totals for the multi-line note (trade payables): reference the individual rows directly
                pass
            r += 3
        return r

    asset_ids = [l[0] for l in BS_LINES if l[3] == "A"]
    liab_ids = [l[0] for l in BS_LINES if l[3] == "L"]
    write_notes(wsa, asset_ids, 2, 3, 1, 4, f"As at {end}")
    write_notes(wsl, liab_ids, 3, 4, 1, 3, f"As at {end}")
    write_notes(wsn, [p[0] for p in PL_LINES if p[0] in ("REV", "OI", "COGS", "EMP", "FIN", "DEP", "OEXP")], 3, 4, 2, 5, f"For the period ended {end}")
    wsn.cell(2, 3, company).font = _f(12, True)
    wsn.cell(4, 3, "Notes to Accounts").font = _f(12, True)
    for ws, widths in ((wsa, {"A": 11, "B": 68, "C": 22}), (wsl, {"A": 11, "B": 11, "C": 68, "D": 22}), (wsn, {"B": 11, "C": 68, "D": 22})):
        for k, v in widths.items():
            ws.column_dimensions[k].width = v
        ws.sheet_view.showGridLines = False
    wsl.cell(1, 3, unit_txt).font = _f(13)

    def ref(lid):
        """formula text that gives the value of a balance sheet line"""
        if lid in note_cell and note_cell[lid]:
            ws, r, c = note_cell[lid]
            return f"='{ws.title}'!{get_column_letter(c)}{r}"
        d = LINE[lid]
        return sumifs(lid, None, -1 if d["nature"] in ("L", "I") else 1)

    # ---------------- P&L sheet
    p = wsp
    p.sheet_view.showGridLines = False
    for k, v in {"A": 3, "B": 73, "C": 10, "D": 22}.items():
        p.column_dimensions[k].width = v
    p["B3"], p["B4"] = company.upper(), f"CIN:{cin}" if cin else ""
    p["B5"] = meta.get("address") or ""
    p["B6"] = "Statement of Standalone Profit and Loss "
    p["B7"] = f"All amount are  in {unit} unless otherwise stated"
    for a in ("B3", "B4", "B5", "B6"):
        p[a].font, p[a].alignment = _f(12, True), Alignment(horizontal="center")
    for col in "BCD":
        p[f"{col}3"].alignment = Alignment(horizontal="center")
    p["B7"].font, p["B7"].alignment = _f(12), Alignment(horizontal="right")
    ptxt = meta.get("period_label") or "For the period ended"
    for a, t in (("B8", "Particulars"), ("C8", "Note No"), ("D8", f"{ptxt} {end}")):
        p[a] = t
        p[a].font, p[a].alignment, p[a].border = _f(12, True), Alignment(horizontal="center", wrap_text=True), Border(top=MED, bottom=MED)
    pl_rows = {}

    def prow(r, label, lid=None, formula=None, bold=False, note=None, topb=None):
        p.cell(r, 2, label).font = _f(12, bold)
        p.cell(r, 2).alignment = Alignment(horizontal="left", wrap_text=True)
        if note:
            p.cell(r, 3, note).alignment = Alignment(horizontal="center")
            p.cell(r, 3).font = _f(12)
        if lid or formula:
            c = p.cell(r, 4, formula if formula else None)
            c.font, c.number_format = _f(12, bold), NUMF
            if topb:
                c.border = Border(top=topb, bottom=topb)
        pl_rows[label] = r

    def plref(lid):
        if lid in note_cell and note_cell[lid]:
            ws, r, c = note_cell[lid]
            return f"='{ws.title}'!{get_column_letter(c)}{r}"
        d = LINE[lid]
        return sumifs(lid, None, -1 if d["nature"] in ("L", "I") else 1)

    prow(9, "Income", bold=True)
    prow(10, "Revenue from operations ", "REV", plref("REV"), note=26)
    prow(11, "Other income", "OI", plref("OI"), note=27)
    prow(12, "Total Income", "x", "=SUM(D10:D11)", True, topb=THICK)
    prow(13, "Expenses", bold=True)
    prow(14, "Cost of materials consumed", "COGS", plref("COGS"), note=28)
    prow(15, "Employee benefits expense", "EMP", plref("EMP"), note=29)
    prow(16, "Finance costs", "FIN", plref("FIN"), note=30)
    prow(17, "Depreciation and amortization expense", "DEP", plref("DEP"), note=31)
    prow(18, "Other expenses", "OEXP", plref("OEXP"), note=32)
    prow(19, "Total Expenses", "x", "=SUM(D14:D18)", True, topb=THICK)
    prow(20, "Profit before tax and exceptional items", "x", "=D12-D19", True)
    prow(21, "Exceptional items", "EXC", sumifs("EXC", None, 1))
    prow(22, "Profit before tax", "x", "=D20-D21", True)
    prow(23, "Tax expense", bold=True)
    prow(24, "(1) Current tax", "CTAX", sumifs("CTAX", None, 1))
    prow(25, "(2) Deferred tax", "DTAX", sumifs("DTAX", None, 1))
    prow(26, "Net Profit for the Period", "x", "=D22-D25-D24", True, topb=THICK)
    prow(27, "Other comprehensive income (OCI)", bold=True)
    prow(28, "(a) (i) Items that will not be reclassified to profit or loss", "OCI", sumifs("OCI", None, -1))
    rate = float(meta.get("oci_tax_rate", 0.0))
    prow(29, "      (ii) Tax on items that will not be reclassified to profit or loss", "x", f"=-D28*{rate}")
    prow(30, "(b) (i) Items that will be reclassified to profit or loss")
    prow(31, "    (ii) Income tax relating to items that will be reclassified to profit or loss")
    prow(32, "Total Other Comprehensive income", "x", "=D28+D29", True)
    prow(33, "Total Comprehensive income", "x", "=D26+D32", True, topb=THICK)
    prow(34, "Earnings per share", bold=True)
    prow(35, "Basic earnings per share of Rs.{}/-each".format(int(meta.get("face_value") or 10)), "x", "=IF(D39=0,0,D26*{}/D39)".format(int(UNITS[unit])))
    prow(36, "Diluted earnings per share of Rs.{}/- each".format(int(meta.get("face_value") or 10)), "x", "=D35")
    prow(38, "Weighted average equity shares used in Computation of EPS", bold=True)
    prow(39, "Basic", "x", round(res["pl"]["SHARES"]))
    prow(40, "Diluted", "x", "=D39")
    p["D39"].number_format = p["D40"].number_format = "#,##0"
    p.cell(41, 2, "Summary of Material accounting Policy information").font = _f(12)
    p.cell(41, 3, 1)
    p.cell(42, 2, "The accompanying notes form an integral part of the financial statements").font = _f(12)
    # fix the OE note references to the P&L
    for key, cellref in (("_OE_PROFIT", "D26"), ("_OE_OCI", "D32")):
        if key in note_cell:
            ws, r, c = note_cell[key]
            ws.cell(r, c).value = f"=PL!{cellref}"
            ws.cell(r, c).number_format = ACC

    # ---------------- Balance sheet
    b = wsb
    b.sheet_view.showGridLines = False
    prev = res.get("prev_bs")
    for k, v in {"A": 4, "B": 4, "C": 6, "D": 66, "E": 12, "F": 20, "G": 20}.items():
        b.column_dimensions[k].width = v
    b["C2"], b["C3"] = company, f"CIN:{cin}" if cin else ""
    b["C4"] = "Standalone Statement of Assets and Liabilities"
    b["C5"] = unit_txt
    for a, bold in (("C2", True), ("C3", False), ("C4", True), ("C5", False)):
        b[a].font, b[a].alignment = _f(13, bold), Alignment(horizontal="left")
    for col in "CDEFG" if prev else "CDEF":
        b[f"{col}5"].border = Border(bottom=MED)
    hdrs = [("C6", "S.No"), ("D6", "Particulars"), ("E6", "Note"), ("F6", f"As at \n{end}")]
    if prev:
        hdrs.append(("G6", f"As at \n{fmt_date(meta.get('prev_end'))}"))
    for a, t in hdrs:
        b[a] = t
        b[a].font, b[a].alignment, b[a].border = _f(13, True), Alignment(horizontal="center" if a != "D6" else None, wrap_text=True), Border(top=MED, bottom=MED)
    b.row_dimensions[6].height = 35

    r = 7
    pos = {}

    def head(label, sno=None, bold=True):
        nonlocal r
        if sno:
            b.cell(r, 3, sno).font = _f(13, True)
        b.cell(r, 4, label).font = _f(13, bold)
        r += 1

    def line(lid):
        nonlocal r
        d = LINE[lid]
        b.cell(r, 4, d["label"]).font = _f(13)
        if d["note"]:
            b.cell(r, 5, d["note"]).alignment = Alignment(horizontal="center")
            b.cell(r, 5).font = _f(13)
        c = b.cell(r, 6, ref(lid))
        c.font, c.number_format = _f(12), ACC
        if prev:
            pc = b.cell(r, 7, round(prev.get(lid, 0.0), 6))
            pc.font, pc.number_format = _f(12), ACC
        pos[lid] = r
        r += 1

    def total(label, formula, key):
        nonlocal r
        b.cell(r, 4, label).font = _f(13, True)
        c = b.cell(r, 6, formula)
        c.font, c.number_format, c.border = _f(13, True), NUMF, Border(top=MED, bottom=MED)
        c.alignment = Alignment(wrap_text=True)
        if prev:
            pc = b.cell(r, 7, formula.replace("F", "G"))
            pc.font, pc.number_format, pc.border = _f(13, True), NUMF, Border(top=MED, bottom=MED)
        pos[key] = r
        r += 1

    head("ASSETS", "I")
    head(" Non-current assets", "(1)")
    n0 = r
    for lid in ("PPE", "CWIP", "ROU", "INVPROP"):
        line(lid)
    head("(e) Financial assets", bold=False)
    for lid in ("INVEST", "LOANS_NC", "OFA_NC", "ONCA", "DTA"):
        line(lid)
    total("Total Non current Assets ", f"=SUM(F{n0}:F{r - 1})", "TOT_NCA")
    r += 1
    head(" Current assets", "(2)")
    c0 = r
    line("INVENT")
    head("(b) Financial assets", bold=False)
    for lid in ("TR", "CASH", "BANK", "OFA_C", "OCA"):
        line(lid)
    total("Total Current Assets", f"=SUM(F{c0}:F{r - 1})", "TOT_CA")
    r += 1
    total("Total Assets", f"=F{pos['TOT_NCA']}+F{pos['TOT_CA']}", "TOT_ASSETS")
    r += 1
    head("EQUITY AND LIABILITIES", "II")
    head("EQUITY", bold=True)
    e0 = r
    line("SC")
    line("OE")
    total("Total Equity", f"=SUM(F{e0}:F{r - 1})", "TOT_EQ")
    r += 1
    head("LIABILITIES", bold=True)
    head(" Non-current liabilities", "(1)")
    head("(a) Financial  Liabilities", bold=False)
    l0 = r
    for lid in ("BORR_NC", "LEASE_NC", "OFL_NC", "PROV_NC"):
        line(lid)
    total("Total Non current Liabilities  ", f"=SUM(F{l0}:F{r - 1})", "TOT_NCL")
    r += 1
    head(" Current liabilities", "(2)")
    head("(a) Financial  Liabilities", bold=False)
    k0 = r
    line("BORR_C")
    line("LEASE_C")
    head("iii)   Trade payables", bold=False)
    for lid in ("TP_MSME", "TP_OTH", "OFL_C", "OCL", "PROV_C"):
        line(lid)
    total("Total Current Liabilites", f"=SUM(F{k0}:F{r - 1})", "TOT_CL")
    total("Total Equity and Liabilities", f"=F{pos['TOT_EQ']}+F{pos['TOT_NCL']}+F{pos['TOT_CL']}", "TOT_EL")
    b.cell(r, 4, "Difference (Total Equity and Liabilities - Total Assets)").font = _f(11)
    b.cell(r, 6, f"=F{pos['TOT_EL']}-F{pos['TOT_ASSETS']}").number_format = NUMF
    if prev:
        b.cell(r, 7, f"=G{pos['TOT_EL']}-G{pos['TOT_ASSETS']}").number_format = NUMF
    r += 2
    b.cell(r, 4, "Summary of Material accounting Policy information").font = _f(13)
    b.cell(r, 5, 1)
    r += 1
    b.cell(r, 4, "The accompanying notes form an integral part of the financial statements").font = _f(13)
    r += 2
    for t in ("As per our report attached of even date", "", f"For {meta.get('auditor') or '________________'}", "Chartered Accountants",
              "Firm Registration No:", "", "Partner", "M No:", "UDIN:", "", "Date:", f"Place:{meta.get('place') or ''}"):
        b.cell(r, 4, t).font = _f(13)
        r += 1

    # ---------------- Checks
    wsc["B2"], wsc["B2"].font = "Automated financial statements - checks", _f(13, True)
    rows = [("Trial balance debit total", f"=SUMIFS('TB Mapped'!C:C,'TB Mapped'!A:A,\"Trial balance\")"),
            ("Trial balance credit total", f"=SUMIFS('TB Mapped'!D:D,'TB Mapped'!A:A,\"Trial balance\")"),
            ("Posted debit total (incl. closing stock & adjustments)", f"='TB Mapped'!C{last + 2}"),
            ("Posted credit total", f"='TB Mapped'!D{last + 2}"),
            ("Total Assets", f"=BS!F{pos['TOT_ASSETS']}"),
            ("Total Equity and Liabilities", f"=BS!F{pos['TOT_EL']}"),
            ("Difference", f"=BS!F{pos['TOT_EL']}-BS!F{pos['TOT_ASSETS']}")]
    for i, (t, fm) in enumerate(rows, 4):
        wsc.cell(i, 2, t).font = _f(12)
        wsc.cell(i, 3, fm).number_format = NUMF
    wsc.column_dimensions["B"].width = 60
    wsc.column_dimensions["C"].width = 20
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------------------------- 5. reconcile with an existing workbook
def _words(t):
    return {w for w in norm(t).split() if len(w) > 2}


def _label_sim(a, b):
    na, nb = norm(re.sub(r"^\s*[\(\[]?[a-z0-9ivx]{1,4}[\)\]]\s*", "", str(a).strip().lower())), norm(re.sub(r"^\s*[\(\[]?[a-z0-9ivx]{1,4}[\)\]]\s*", "", str(b).strip().lower()))
    if not na or not nb:
        return 0.0
    r = SequenceMatcher(None, na, nb).ratio()
    wa, wb = set(na.split()), set(nb.split())
    j = len(wa & wb) / max(1, len(wa | wb))
    return max(r, j)


def _row_note(eng, bk, sh, row, col):
    note = None
    for c in range(1, col):
        v = bk.val(sh, eng.addr(c, row))
        if isinstance(v, (int, float)) and not isinstance(v, bool) and 0 < v < 100 and float(v).is_integer():
            note = int(v)
    return note


def match_original_lines(eng, bk, st, col):
    """map every line of the original balance sheet to one of our line ids"""
    lines = eng.get_lines(bk, st, col)
    out = []
    bs_ids = [(l[0], l[1], l[2], l[4]) for l in BS_LINES]
    for ln in lines:
        lab = ln["label"]
        if re.match(r"^\s*total\b", lab, re.I) or ln["kind"] == "subtotal":
            continue
        note = _row_note(eng, bk, st["sheet"], ln["row"], col)
        best, bsc = None, 0.0
        for lid, lbl, n, sec in bs_ids:
            if (ln["side"] == "A") != (LINE[lid]["nature"] == "A"):
                continue
            sc = _label_sim(lab, lbl)
            if note is not None and n is not None:
                sc += 0.5 if note == n else -0.35
            if sc > bsc:
                best, bsc = lid, sc
        out.append({**ln, "note": note, "match": best if bsc >= 0.6 else None, "score": round(bsc, 2)})
    return out


def pick_column(st, end_date):
    best = None
    for c in st["cols"]:
        ds = parse_dates(c["header"])
        if end_date and ds and ds[-1] == end_date:
            best = c
            break
    return best["col"] if best else (st["cols"][0]["col"] if st["cols"] else None)


def _scan_numbers(bk, sh, limit=60000):
    ws = bk.wv[sh]
    out = []
    n = 0
    for row in ws.iter_rows(values_only=False):
        for c in row:
            v = c.value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) > 1e-9:
                out.append((float(v), c.coordinate, c.row, c.column))
                n += 1
        if n > limit:
            break
    return out


def _context(bk, sh, row, col):
    ws = bk.wv[sh]
    lab = ""
    for c in range(1, min(col, 8)):
        v = ws.cell(row, c).value
        if isinstance(v, str) and clean(v) and len(clean(v)) > len(lab):
            lab = clean(v)
    head = ""
    for r in range(row - 1, max(0, row - 40), -1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and clean(v) and not re.match(r"^[\d.,\- ]+$", clean(v)):
            head = clean(v)
            break
    return " / ".join(x for x in (lab, head) if x)


def reconcile(eng, bk, res, ledgers: pd.DataFrame, sheet, col, tb_sheets, tol=0.01, scale=1.0):
    """res: result of build(); ledgers: grouped TB frame (with Row, Ledger, Line, Debit, Credit in INPUT units)"""
    st = eng.find_structure(bk, sheet)
    if st is None:
        return {"ok": False, "error": "balance sheet structure not found"}
    ws = bk.wv[sheet]
    colinfo = next((c for c in st["cols"] if c["col"] == col), None)
    A = ws.cell(st["a_tot"], col).value
    EL = ws.cell(st["l_tot"], col).value
    A = A if isinstance(A, (int, float)) else None
    EL = EL if isinstance(EL, (int, float)) else None
    orig_lines = match_original_lines(eng, bk, st, col)
    # per-line comparison
    by_id = {}
    for ln in orig_lines:
        if ln["match"]:
            by_id.setdefault(ln["match"], []).append(ln)
    rows = []
    for lid, lbl, n, nat, sec in BS_LINES:
        auto = res["bs"][lid]
        ol = by_id.get(lid, [])
        orig = sum((l["value"] or 0.0) for l in ol)
        errs = [l for l in ol if l["err"]]
        var = orig - auto
        rows.append({"id": lid, "label": LINE[lid]["label"], "note": n, "nature": nat, "auto": auto, "orig": orig, "var": var,
                     "cells": [l["addr"] for l in ol], "err": bool(errs), "err_cells": [l["addr"] for l in errs]})
    unmatched = [l for l in orig_lines if not l["match"] and (l["value"] or 0) != 0]
    # effect on (E&L - Assets) of the original relative to an always-balanced build
    gap = None if (A is None or EL is None) else EL - A
    eff = lambda r: (r["var"] if r["nature"] == "L" else -r["var"])
    explained = sum(eff(r) for r in rows if abs(r["var"]) > tol)
    # unmatched original lines' effect
    um_eff = 0.0
    for l in unmatched:
        um_eff += (l["value"] or 0.0) * (1 if l["side"] == "L" else -1)
    own_tot_a = sum((l["value"] or 0.0) for l in orig_lines if l["side"] == "A")
    own_tot_l = sum((l["value"] or 0.0) for l in orig_lines if l["side"] == "L")
    arithmetic = {"A": (A - own_tot_a) if A is not None else None, "EL": (EL - own_tot_l) if EL is not None else None}
    # ---- subset search: which line variances (effects) reproduce the gap
    sols = []
    flagged = [r for r in rows if abs(r["var"]) > tol and not r["err"]]
    if gap is not None and abs(gap) > tol:
        for k in (1, 2, 3):
            for combo in itertools.combinations(flagged, k):
                if abs(sum(eff(r) for r in combo) - gap) <= max(tol, 0.005):
                    sols.append([r["id"] for r in combo])
            if sols:
                break
    # ---- ledger level drill-down for each flagged line
    names = {}
    for rec in ledgers.to_dict("records"):
        names.setdefault(norm(rec["Ledger"]), []).append(rec)
    my_post = res["post"][res["post"]["Source"] == "Trial balance"]
    drill = {}
    for r in rows:
        if abs(r["var"]) <= tol and not r["err"]:
            continue
        info = {"reads_tb": [], "tb_total": 0.0, "only_auto": [], "only_orig": [], "sheets": [], "typed": [], "schedule": [], "schedule_agree": 0}
        leaves, seen_sheets, typed = {}, set(), []
        for a in r["cells"]:
            try:
                tr = eng.trace_line(bk, tb_sheets, sheet, a)
            except Exception:  # noqa: BLE001
                continue
            for cell, sg in tr["tb"].items():
                leaves[cell] = leaves.get(cell, 0) + sg
            for cell in tr["seen"]:
                if cell[0] not in tb_sheets and cell[0] != sheet:
                    seen_sheets.add(cell[0])
            typed += tr["hard"]
        info["sheets"] = sorted(seen_sheets)
        info["typed"] = [{"cell": f"{c[0]}!{c[1]}", "value": v} for c, v in typed[:8]]
        read_rows = set()
        for (tsh, ta), sg in leaves.items():
            rr = int(re.sub(r"[A-Z]+", "", ta))
            rw = bk.wv[tsh][rr]
            texts = [norm(c.value) for c in rw[:10] if isinstance(c.value, str)]
            cand = [x for t in texts for x in names.get(t, [])]
            if cand:
                read_rows.add(cand[0]["Row"] if len(cand) == 1 else min(cand, key=lambda x: abs(x["Row"] - rr))["Row"])
        mine = ledgers[(ledgers["Include"]) & (ledgers["Line"] == r["id"])]
        info["reads_tb"] = sorted(read_rows)
        if leaves and r["id"] != "OE":
            for rec in mine.to_dict("records"):
                if rec["Row"] not in read_rows:
                    info["only_auto"].append({"ledger": rec["Ledger"], "net": (rec["Debit"] - rec["Credit"]) * scale})
            for rec in ledgers.to_dict("records"):
                if rec["Row"] in read_rows and rec["Line"] != r["id"] and rec["Include"] and LINE[rec["Line"]]["stmt"] == "BS":
                    info["only_orig"].append({"ledger": rec["Ledger"], "net": (rec["Debit"] - rec["Credit"]) * scale, "auto_line": rec["Line"]})
        if r["id"] == "OE":
            info["reads_tb"] = []
        if info["sheets"] and r["id"] != "OE" and LINE[r["id"]]["stmt"] == "BS":
            # the line is (also) fed by schedules / notes: see which of our ledger totals the schedules reproduce
            end_txt = fmt_date((res.get("meta") or {}).get("end"))
            comp = [(rec["Ledger"], rec["Net"]) for rec in my_post[my_post["Line"] == r["id"]].to_dict("records") if abs(rec["Net"]) > tol]
            cache = {}
            agree = 0
            for sh in info["sheets"][:5]:
                cache[sh] = _scan_numbers(bk, sh)
            for lab, net in comp:
                amt = abs(net)
                found = False
                for sh, nums in cache.items():
                    if any(abs(abs(v) - amt) <= tol for v, *_ in nums):
                        found = True
                        break
                if found:
                    agree += 1
                    continue
                w = {x.rstrip("s") for x in _words(lab)}
                best = None
                for sh, nums in cache.items():
                    for v, a, rr, cc in nums:
                        if abs(v) < tol or abs(abs(v) - amt) > 0.6 * amt:
                            continue
                        ctx = _context(bk, sh, rr, cc)
                        ov = len(w & {x.rstrip("s") for x in _words(ctx)})
                        if not ov:
                            continue
                        key = (ov + (1 if end_txt and end_txt in ctx else 0), -abs(abs(v) - amt))
                        if best is None or key > best[0]:
                            best = (key, v, f"{sh}!{a}", ctx)
                info["schedule"].append({"ledger": lab, "amount": amt, "found": best[2] if best else "", "context": best[3] if best else "",
                                         "diff": (abs(best[1]) - amt) if best else None})
            info["schedule_agree"] = agree
        drill[r["id"]] = info
    return {"ok": True, "sheet": sheet, "col": col, "header": clean(colinfo["header"]) if colinfo else "", "orig_assets": A, "orig_el": EL,
            "gap": gap, "rows": rows, "unmatched": unmatched, "explained": explained, "unmatched_effect": um_eff,
            "arithmetic": arithmetic, "solutions": sols, "drill": drill, "tol": tol}


def _fm(v, dec=2):
    if v is None:
        return "-"
    s = f"{abs(v):,.{dec}f}"
    return f"-{s}" if v < 0 else s


def recon_narrative(rc: dict, unit: str = "Lakhs") -> str:
    """plain-language result of reconcile()"""
    if not rc.get("ok"):
        return rc.get("error", "")
    tol = rc["tol"]
    out = []
    gap = rc["gap"]
    if gap is None:
        out.append("The original balance sheet total could not be read (formula errors in the totals).")
    elif abs(gap) <= tol:
        out.append(f"The original balance sheet for **{rc['header']}** tallies (difference {_fm(gap)}).")
    else:
        side = "Total Equity & Liabilities is higher than Total Assets" if gap > 0 else "Total Assets is higher than Total Equity & Liabilities"
        out.append(f"In the original workbook ({rc['sheet']}, {rc['header']}) {side} by **{_fm(abs(gap))} {unit}**. "
                   f"The statements built from the trial balance tally, so every line that differs from them is a suspect.")
    flagged = [r for r in rc["rows"] if abs(r["var"]) > tol or r["err"]]
    eff = lambda r: (r["var"] if r["nature"] == "L" else -r["var"])
    if rc["solutions"]:
        sol = rc["solutions"][0]
        names = {r["id"]: r for r in rc["rows"]}
        parts = []
        for lid in sol:
            r = names[lid]
            if r["nature"] == "A":
                parts.append(f"**{r['label'].strip()}** (original {_fm(r['orig'])} vs trial balance {_fm(r['auto'])}: Total Assets "
                             f"{'under' if r['var'] < 0 else 'over'}stated by {_fm(abs(r['var']))})")
            else:
                parts.append(f"**{r['label'].strip()}** (original {_fm(r['orig'])} vs trial balance {_fm(r['auto'])}: Equity & Liabilities "
                             f"{'under' if r['var'] < 0 else 'over'}stated by {_fm(abs(r['var']))})")
        out.append("**The whole difference is explained by:** " + "; ".join(parts) + ".")
    elif gap is not None and abs(gap) > tol:
        out.append(f"The lines that differ from the trial balance explain {_fm(rc['explained'])} of the {_fm(gap)} difference.")
    ar = rc["arithmetic"]
    for k, nm in (("A", "Total Assets"), ("EL", "Total Equity & Liabilities")):
        if ar.get(k) is not None and abs(ar[k]) > tol:
            out.append(f"**{nm}** in the original is not the sum of the lines above it: it is off by {_fm(ar[k])} (a SUM range or a link may skip a line).")
    errs = [r for r in rc["rows"] if r["err"]]
    if errs:
        out.append("Lines showing formula errors in the original: " + ", ".join(f"{r['label'].strip()} ({', '.join(r['err_cells'])})" for r in errs) + ".")
    # reclassifications: variances that offset each other
    rest = [r for r in flagged if not r["err"] and not (rc["solutions"] and r["id"] in rc["solutions"][0])]
    if rest:
        out.append("**Other lines that differ but do not change the tally** (classification differences between the original and the automatic grouping): "
                   + "; ".join(f"{r['label'].strip()} {_fm(r['var'])}" for r in rest) + ".")
    return "\n\n".join(out)


def recon_to_excel(rc: dict, res: dict) -> bytes:
    rows = [{"Line": r["label"].strip(), "Note": r["note"], "Original": r["orig"], "From trial balance": r["auto"], "Difference (orig - TB)": r["var"],
             "Original cells": ", ".join(r["cells"]), "Status": "Error in original" if r["err"] else ("Differs" if abs(r["var"]) > rc["tol"] else "Agrees")}
            for r in rc["rows"]]
    det = []
    for lid, d in rc["drill"].items():
        lab = LINE[lid]["label"]
        for x in d["only_auto"]:
            det.append({"Line": lab, "Finding": "Ledger grouped here by automation but not read by the original line", "Item": x["ledger"], "Amount": x["net"]})
        for x in d["only_orig"]:
            det.append({"Line": lab, "Finding": f"Ledger read by the original line but grouped to {x['auto_line']} by automation", "Item": x["ledger"], "Amount": x["net"]})
        for x in d["schedule"]:
            if x["found"]:
                det.append({"Line": lab, "Finding": f"Closest schedule figure {x['found']} ({x['context']}) differs from ledger by", "Item": x["ledger"], "Amount": x["diff"]})
        for x in d["typed"]:
            det.append({"Line": lab, "Finding": "Typed (hard-coded) number feeding the original line", "Item": x["cell"], "Amount": x["value"]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"Analysis": recon_narrative(rc, res["unit"]).replace("**", "").split("\n\n")}).to_excel(xw, sheet_name="Summary", index=False)
        pd.DataFrame(rows).to_excel(xw, sheet_name="Line comparison", index=False)
        pd.DataFrame(det).to_excel(xw, sheet_name="Details", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(90, max(12, max(len(str(c.value or "")) for c in col[:60]) + 2))
    return buf.getvalue()


# --------------------------------------------------------------------------------------------- 6. one-upload workflow helpers
def sheet_unit_hint(bk, sheet):
    """'in Lakhs' style note near the top of a statement sheet"""
    ws = bk.wv[sheet]
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=12, values_only=True):
        for v in row:
            if isinstance(v, str):
                m = re.search(r"in\s+(lakh|lac|thousand|crore|million|rupee)", v, re.I)
                if m:
                    return {"lakh": "Lakhs", "lac": "Lakhs", "thousand": "Thousands", "crore": "Crores", "million": "Millions", "rupee": "Rupees"}[m.group(1).lower()]
    return None


def learn_grouping(eng, bk, g, sheet, col, tb_sheets, tb_sheet):
    """Use the original workbook's own formulas as the grouping: every ledger row that a balance sheet line reads
    is put on that line.  Ledgers read by two different lines, or by none, keep the automatic grouping."""
    st = eng.find_structure(bk, sheet)
    lines = match_original_lines(eng, bk, st, col)
    row_lines = {}
    for ln in lines:
        if not ln["match"] or ln["err"] or ln["match"] == "OE":
            continue   # Other Equity is a roll-forward of profit: it reads every P&L ledger, so it says nothing about grouping
        try:
            tr = eng.trace_line(bk, tb_sheets, sheet, ln["addr"])
        except Exception:  # noqa: BLE001
            continue
        for (tsh, ta), _sg in tr["tb"].items():
            if tsh == tb_sheet:
                row_lines.setdefault(int(re.sub(r"[A-Z]+", "", ta)), set()).add(ln["match"])
    out = g.copy()
    n = 0
    for i, r in out.iterrows():
        ls = row_lines.get(int(r["Row"]))
        if LINE[r["Line"]]["stmt"] != "BS":
            continue   # income / expense ledgers are never moved into a balance sheet line
        if ls and len(ls) == 1:
            lid = next(iter(ls))
            if r["Line"] != lid:
                out.at[i, "SubGroup"] = (r["FileGroup"] if r["FileGroup"] and norm(r["FileGroup"]) != "group head" else r["SubGroup"])
            out.at[i, "Line"], out.at[i, "Reason"], out.at[i, "Confidence"] = lid, "learned from the original workbook's formulas", 1.0
            n += 1
    return out, n


def manual_postings(current_tax=0.0, deferred_tax=0.0, other=None):
    """journal lines for the manual items that finish the statements.
    current_tax  : additional current tax expense   -> Dr current tax, Cr provision for income tax
    deferred_tax : + charge / - credit              -> Dr deferred tax, Cr deferred tax asset (reverse when negative)"""
    rows = []
    if abs(current_tax) > 1e-12:
        rows += [["Current tax for the period", id_to_choice("CTAX"), "Current tax", max(current_tax, 0), max(-current_tax, 0)],
                 ["Current tax - provision", id_to_choice("PROV_C"), "Provision for income tax", max(-current_tax, 0), max(current_tax, 0)]]
    if abs(deferred_tax) > 1e-12:
        rows += [["Deferred tax for the period", id_to_choice("DTAX"), "Deferred tax", max(deferred_tax, 0), max(-deferred_tax, 0)],
                 ["Deferred tax - deferred tax asset", id_to_choice("DTA"), "Deferred tax asset", max(-deferred_tax, 0), max(deferred_tax, 0)]]
    df = pd.DataFrame(rows, columns=["Narration", "Line", "SubGroup", "Debit", "Credit"])
    if other is not None and len(other):
        df = pd.concat([df, other], ignore_index=True)
    return df


# --------------------------------------------------------------------------------------------- 7. previous-year financials + current trial balance
PL_TARGETS = [("REV", "Revenue from operations"), ("OI", "Other income"), ("COGS", "Cost of materials consumed"),
              ("EMP", "Employee benefits expense"), ("FIN", "Finance costs"), ("DEP", "Depreciation and amortization expense"),
              ("OEXP", "Other expenses"), ("CTAX", "Current tax"), ("DTAX", "Deferred tax")]


def find_pl_sheets(bk):
    out = []
    for n in bk.sheetnames:
        ws = bk.wv[n]
        txt = " ".join(str(v).lower() for row in ws.iter_rows(min_row=1, max_row=10, max_col=8, values_only=True) for v in row if isinstance(v, str))
        sc = (2 if "profit and loss" in txt or "profit & loss" in txt else 0) + (1 if "statement of" in txt else 0) + (1 if norm(n) in ("pl", "p and l", "profit and loss") else 0)
        if "assets" in txt and "liabilit" in txt:
            sc -= 2
        if sc >= 2:
            out.append((sc, n))
    out.sort(key=lambda x: -x[0])
    return [n for _, n in out]


def pick_pl_column(bk, sheet, end):
    ws = bk.wv[sheet]
    rank = {"year": 4, "nine": 3, "half": 2, "quarter": 1}
    best, bs = None, -1
    for c in range(3, 20):
        h = " ".join(str(ws.cell(r, c).value) for r in range(5, 12) if isinstance(ws.cell(r, c).value, str))
        ds = parse_dates(h)
        if not ds or (end and ds[-1] != end):
            continue
        hl = h.lower()
        k = 4 if re.search(r"year\s+ended", hl) and "half" not in hl else 3 if "nine" in hl else 2 if "half" in hl else 1 if "quarter" in hl else 4
        if k > bs:
            best, bs = c, k
    return best


def learn_pl_grouping(eng, bk, g, sheet, col, tb_sheets, tb_sheet):
    """group income / expense ledgers the way the original P&L reads them"""
    ws = bk.wv[sheet]
    rows = []
    for r in range(1, min(ws.max_row, 80) + 1):
        lab = ""
        for c in range(1, 5):
            v = ws.cell(r, c).value
            if isinstance(v, str) and clean(v):
                lab = clean(v)
                break
        if lab and not re.match(r"^total", lab, re.I):
            rows.append((r, lab))
    row_lines = {}
    for lid, name in PL_TARGETS:
        cand = [(r, _label_sim(lab, name)) for r, lab in rows]
        cand = [x for x in cand if x[1] >= 0.8]
        if not cand:
            continue
        r = max(cand, key=lambda x: x[1])[0]
        addr = f"{get_column_letter(col)}{r}"
        try:
            tr = eng.trace_line(bk, tb_sheets, sheet, addr)
        except Exception:  # noqa: BLE001
            continue
        for (tsh, ta), _sg in tr["tb"].items():
            if tsh == tb_sheet:
                row_lines.setdefault(int(re.sub(r"[A-Z]+", "", ta)), set()).add(lid)
    out = g.copy()
    n = 0
    for i, r in out.iterrows():
        ls = row_lines.get(int(r["Row"]))
        if ls and len(ls) == 1 and LINE[r["Line"]]["stmt"] == "PL":
            lid = next(iter(ls))
            out.at[i, "SubGroup"] = pl_subgroup(lid, r["SubGroup"], r)
            out.at[i, "Line"], out.at[i, "Reason"], out.at[i, "Confidence"] = lid, "learned from the P&L of the original workbook", 1.0
            n += 1
    return out, n


def previous_mapping(prev_g: pd.DataFrame) -> dict:
    """{normalised ledger name: (line, sub group)} from a grouped previous-year trial balance"""
    m = {}
    for rec in prev_g.to_dict("records"):
        k = norm(rec["Ledger"])
        if k and k not in m:
            m[k] = (rec["Line"], rec["SubGroup"] or "")
    return m


def apply_previous(cur: pd.DataFrame, prev_g: pd.DataFrame, cutoff: float = 0.9):
    """group the current-year ledgers with the previous-year grouping; new ledgers use system knowledge.
    adds Origin: 'Previous year' / 'Similar to previous year' / 'New ledger'"""
    from difflib import get_close_matches
    m = previous_mapping(prev_g)
    keys = list(m)
    fresh = auto_group(cur)
    out = fresh.copy()
    origin = []
    for i, rec in out.iterrows():
        k = norm(rec["Ledger"])
        if k in m:
            out.at[i, "Line"], out.at[i, "SubGroup"] = m[k]
            out.at[i, "Reason"], out.at[i, "Confidence"] = "same ledger in previous-year financials", 1.0
            origin.append("Previous year")
            continue
        close = get_close_matches(k, keys, n=1, cutoff=cutoff) if len(k) > 6 else []
        if close:
            out.at[i, "Line"], out.at[i, "SubGroup"] = m[close[0]]
            out.at[i, "Reason"], out.at[i, "Confidence"] = f"similar to previous-year ledger '{close[0]}'", 0.85
            origin.append("Similar to previous year")
            continue
        out.at[i, "Reason"] = "new ledger - system knowledge: " + str(rec["Reason"])
        out.at[i, "Confidence"] = min(float(rec["Confidence"]), 0.69)
        origin.append("New ledger")
    out["Origin"] = origin
    seen_cur = {norm(x) for x in cur["Ledger"]}
    missing = prev_g[~prev_g["Ledger"].map(norm).isin(seen_cur) & ((prev_g["Debit"] - prev_g["Credit"]).abs() > 1e-9) & prev_g["Include"]]
    return out, missing


def previous_statement_values(eng, bk, sheet, col, unit_from, unit_to):
    """balance sheet line values of the previous year, in the unit of the new statements"""
    st = eng.find_structure(bk, sheet)
    vals = {}
    f = UNITS[unit_from] / UNITS[unit_to]
    for ln in match_original_lines(eng, bk, st, col):
        if ln["match"] and ln["value"] is not None and not ln["err"]:
            vals[ln["match"]] = vals.get(ln["match"], 0.0) + float(ln["value"]) * f
    return vals
