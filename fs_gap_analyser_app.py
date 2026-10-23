"""
Balance Sheet Gap Analyser  -  Streamlit front end for fs_gap_engine.py

Run standalone :  streamlit run fs_gap_analyser_app.py
Integrate later:  import fs_gap_analyser_app as fsgap ; fsgap.render(key_prefix="fsgap")   (two modes: gap finder / automated FS builder)
                  (all widget / session keys are namespaced with key_prefix)
"""
import datetime as dt
import hashlib
import traceback

import pandas as pd
import streamlit as st

import fs_builder as bld
import fs_gap_engine as eng

ACCENT = "#D85A30"
PAPER = "#FAF6F1"

CSS = f"""
<style>
.fsgap-hero {{background:{PAPER};border-left:6px solid {ACCENT};padding:14px 18px;border-radius:6px;margin-bottom:12px}}
.fsgap-hero h2 {{margin:0 0 4px 0;color:{ACCENT}}}
.fsgap-hero p {{margin:0;color:#444}}
.fsgap-pill {{display:inline-block;padding:2px 10px;border-radius:12px;font-size:0.85em;font-weight:600;color:#fff}}
.fsgap-ok {{background:#3B8E5A}} .fsgap-gap {{background:{ACCENT}}} .fsgap-err {{background:#8A8A8A}}
</style>
"""


@st.cache_resource(show_spinner="Reading workbook (formulas and values)...")
def _load(data_hash: str, data: bytes, name: str):
    return eng.load_book(data, name)


def _table(df, **kw):
    """st.dataframe that works on old (use_container_width) and new (width='stretch') Streamlit."""
    try:
        st.dataframe(df, hide_index=True, width="stretch", **kw)
    except TypeError:
        st.dataframe(df, hide_index=True, use_container_width=True, **kw)


def _editor(df, **kw):
    try:
        return st.data_editor(df, width="stretch", **kw)
    except TypeError:
        return st.data_editor(df, use_container_width=True, **kw)


def _k(prefix, name):
    return f"{prefix}_{name}"


def _render_gap(key_prefix: str = "fsgap"):
    st.markdown(
        '<div class="fsgap-hero"><h2>Balance Sheet Gap Analyser</h2>'
        "<p>Upload a financial-statements workbook that does not tally. The tool follows every balance sheet line back to the "
        "trial balance (SUMIF or direct links), and tells you where the difference comes from.</p></div>",
        unsafe_allow_html=True,
    )

    up = st.file_uploader("Financial statements workbook (.xlsx / .xlsm)", type=["xlsx", "xlsm"], key=_k(key_prefix, "file"))
    if not up:
        st.info("The workbook must contain formulas with saved values (open it in Excel and save once if it was produced by a script). "
                "Nothing in your file is changed.")
        return

    data = up.getvalue()
    h = hashlib.md5(data).hexdigest()
    try:
        bk = _load(h, data, up.name)
        cands = eng.candidate_sheets(bk)
        tb_default = eng.detect_tb_sheets(bk)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the workbook: {e}")
        with st.expander("Details"):
            st.code(traceback.format_exc())
        return

    if not cands:
        st.error("No balance sheet found. The tool looks for a sheet with 'ASSETS', 'EQUITY AND LIABILITIES' headings and total rows.")
        return

    c1, c2 = st.columns([2, 1])
    default_sheet = eng.pick_default_sheet(bk, cands)
    sheet = c1.selectbox("Balance sheet to check", cands, index=cands.index(default_sheet), key=_k(key_prefix, "sheet"),
                         help="Sheets that look like a balance sheet. The one with the latest period is pre-selected.")
    tol = c2.number_input("Tolerance (statement units)", min_value=0.0, value=0.01, step=0.01, format="%.4f", key=_k(key_prefix, "tol"),
                          help="Differences up to this size are treated as rounding.")
    with st.expander("Advanced: trial balance sheets"):
        tbs = st.multiselect("Sheets to treat as trial balance (tracing stops there)", [s for s in bk.sheetnames if s != sheet],
                             default=[s for s in tb_default if s != sheet], key=_k(key_prefix, "tbs"))

    if not st.button("Analyse", type="primary", key=_k(key_prefix, "run")):
        st.caption("Choose the sheet and click Analyse.")
        return

    with st.spinner("Tracing statement lines back to the trial balance..."):
        try:
            res = eng.analyze(bk, sheet=sheet, tb_sheets=tbs, tol=tol)
        except Exception as e:  # noqa: BLE001
            st.error(f"The analysis failed: {e}")
            with st.expander("Details"):
                st.code(traceback.format_exc())
            return
    if not res.get("ok"):
        st.error(res.get("error", "Analysis failed"))
        return

    cols = res["columns"]
    n_ok = sum(c["status"] == "Tallies" for c in cols)
    n_gap = sum(c["status"] == "Gap" for c in cols)
    n_err = sum(c["status"] == "Error" for c in cols)
    m1, m2, m3 = st.columns(3)
    m1.metric("Columns that tally", n_ok)
    m2.metric("Columns with a gap", n_gap)
    m3.metric("Columns with formula errors", n_err)

    df = pd.DataFrame([{
        "Column": c["col"], "Period": " ".join(str(c["period"]).split()),
        "Total Assets": c["assets"], "Total Equity & Liabilities": c["el"], "Difference": c["gap"], "Status": c["status"],
    } for c in cols])
    _table(df, column_config={k: st.column_config.NumberColumn(format="%.2f") for k in ("Total Assets", "Total Equity & Liabilities", "Difference")})

    t1, t2, t3 = st.tabs(["Analysis", "Trial balance coverage", "Download"])
    with t1:
        st.markdown(res["markdown"])
    with t2:
        any_rows = False
        for col, a in res["analyses"].items():
            if not a.get("tb"):
                continue
            any_rows = True
            st.subheader(f"Column {col} - {a['tb']}")
            st.caption(f"{a.get('tb_rows_linked', 0)} of {a.get('tb_rows_total', 0)} ledger rows are read by this column. "
                       f"Debit/credit columns used: {a.get('tb_cols')}")
            tie = a.get("tieout")
            if tie:
                st.markdown("**Ledger tie-out - ledgers no line reads vs. figures typed into schedules/notes**")
                rows_ = ([{"Result": "Agrees", "Ledgers": ", ".join(x.strip() for x in v["rows"]), "Ledger total": v["total"], "Schedule cells": v["cells"],
                           "Schedule - ledger": None} for v in tie["verified"]]
                         + [{"Result": "No schedule figure reproduces it", "Ledgers": u["label"].strip(), "Ledger total": abs(u["net"]), "Schedule cells": "",
                             "Schedule - ledger": None} for u in tie["unverified"]]
                         + [{"Result": "Closest schedule figure (difference equals the gap)", "Ledgers": ", ".join(x.strip() for x in c["rows"]),
                             "Ledger total": c["ledger_total"], "Schedule cells": c["cells"], "Schedule - ledger": c["difference"]} for c in tie["candidates"]])
                _table(pd.DataFrame(rows_))
                for c_ in tie["candidates"]:
                    if c_["inputs"]:
                        st.markdown("**Typed inputs behind " + c_["cells"] + "**")
                        _table(pd.DataFrame([{"Cell": i["cell"], "Value": i["value"], "Row label": i["label"]} for i in c_["inputs"]]))
            if a.get("tb_unlinked"):
                st.markdown("**Ledgers not read by any line**")
                _table(pd.DataFrame([{"Row": u["row"], "Code": u["code"], "Ledger": u["label"], "Debit": u["dr"], "Credit": u["cr"],
                                      "Net (Dr-Cr)": u["net"], "Neighbours feed": u["hint"], "Same amount typed in at": u.get("typed_in", "")}
                                     for u in a["tb_unlinked"]]))
            if a.get("tb_multi"):
                st.markdown("**Ledgers read by more than one line (same side)**")
                _table(pd.DataFrame([{"Row": m["row"], "Ledger": m["label"], "Lines": " | ".join(m["lines"]),
                                      "Amounts": " | ".join(f"{x:,.2f}" for x in m["contrib"])} for m in a["tb_multi"]]))
            if a.get("tb_partial"):
                st.markdown("**Ledgers only partly read**")
                _table(pd.DataFrame(a["tb_partial"]))
        if not any_rows:
            st.info("No column with a gap reached a trial-balance sheet, so there is no coverage table to show.")
    with t3:
        st.download_button("Download Excel report", eng.to_excel(res), file_name="balance_sheet_gap_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_k(key_prefix, "dl_x"))
        st.download_button("Download analysis (Markdown)", res["markdown"].encode("utf-8"), file_name="balance_sheet_gap_analysis.md",
                           mime="text/markdown", key=_k(key_prefix, "dl_m"))


# =====================================================================================================================
#  Mode 2 : build the financial statements from a trial balance, then match them with the workbook that does not tally
# =====================================================================================================================
@st.cache_data(show_spinner="Reading trial balance...")
def _tb_probe(data_hash: str, data: bytes, name: str):
    return bld.tb_sheet_candidates(data, name)


@st.cache_data(show_spinner="Reading trial balance...")
def _tb_layout(data_hash: str, data: bytes, name: str, sheet: str):
    grid, _names, sh = bld._sheet_grid(data, name, sheet)
    return bld.detect_layout(grid)


def _ss(prefix, name, default=None):
    return st.session_state.get(_k(prefix, name), default)


def _set(prefix, name, value):
    st.session_state[_k(prefix, name)] = value


def _fmt_money(v):
    return f"{v:,.2f}"


def _prev_totals(pb):
    t = {k: 0.0 for k in ("NCA", "CA", "EQ", "NCL", "CL")}
    for lid, _l, _n, _nat, sec in bld.BS_LINES:
        t[sec] += pb.get(lid, 0.0)
    return {"TOT_NCA": t["NCA"], "TOT_CA": t["CA"], "TOT_ASSETS": t["NCA"] + t["CA"], "TOT_EQ": t["EQ"], "TOT_NCL": t["NCL"], "TOT_CL": t["CL"],
            "TOT_EL": t["EQ"] + t["NCL"] + t["CL"]}


def _statement_frames(res):
    bs, pl = res["bs"], res["pl"]
    rows = []
    pb = res.get("prev_bs")
    pt = _prev_totals(pb) if pb else None
    labels_prev = {}

    def add(label, note=None, val=None, bold=False, prev=None):
        row = {"Particulars": label, "Note": "" if note is None else str(note), "Amount": val, "_b": bold}
        if pb is not None:
            row["Previous year"] = prev
        rows.append(row)

    add("ASSETS", bold=True)
    add("Non-current assets", bold=True)
    for lid in ("PPE", "CWIP", "ROU", "INVPROP", "INVEST", "LOANS_NC", "OFA_NC", "ONCA", "DTA"):
        add(bld.LINE[lid]["label"], bld.LINE[lid]["note"], bs[lid], prev=(pb or {}).get(lid))
    add("Total Non current Assets", None, bs["TOT_NCA"], True, prev=(pt or {}).get("TOT_NCA"))
    add("Current assets", bold=True)
    for lid in ("INVENT", "TR", "CASH", "BANK", "OFA_C", "OCA"):
        add(bld.LINE[lid]["label"], bld.LINE[lid]["note"], bs[lid], prev=(pb or {}).get(lid))
    add("Total Current Assets", None, bs["TOT_CA"], True, prev=(pt or {}).get("TOT_CA"))
    add("Total Assets", None, bs["TOT_ASSETS"], True, prev=(pt or {}).get("TOT_ASSETS"))
    add("EQUITY AND LIABILITIES", bold=True)
    for lid in ("SC", "OE"):
        add(bld.LINE[lid]["label"], bld.LINE[lid]["note"], bs[lid], prev=(pb or {}).get(lid))
    add("Total Equity", None, bs["TOT_EQ"], True, prev=(pt or {}).get("TOT_EQ"))
    add("Non-current liabilities", bold=True)
    for lid in ("BORR_NC", "LEASE_NC", "OFL_NC", "PROV_NC"):
        add(bld.LINE[lid]["label"], bld.LINE[lid]["note"], bs[lid], prev=(pb or {}).get(lid))
    add("Total Non current Liabilities", None, bs["TOT_NCL"], True, prev=(pt or {}).get("TOT_NCL"))
    add("Current liabilities", bold=True)
    for lid in ("BORR_C", "LEASE_C", "TP_MSME", "TP_OTH", "OFL_C", "OCL", "PROV_C"):
        add(bld.LINE[lid]["label"], bld.LINE[lid]["note"], bs[lid], prev=(pb or {}).get(lid))
    add("Total Current Liabilities", None, bs["TOT_CL"], True, prev=(pt or {}).get("TOT_CL"))
    add("Total Equity and Liabilities", None, bs["TOT_EL"], True, prev=(pt or {}).get("TOT_EL"))
    bsdf = pd.DataFrame(rows)
    pb_keep, pb = pb, None
    rows = []
    add("Revenue from operations", 26, pl["REV"])
    add("Other income", 27, pl["OI"])
    add("Total Income", None, pl["INCOME"], True)
    for lid, nt in (("COGS", 28), ("EMP", 29), ("FIN", 30), ("DEP", 31), ("OEXP", 32)):
        add(bld.LINE[lid]["label"], nt, pl[lid])
    add("Total Expenses", None, pl["EXPENSES"], True)
    add("Profit before tax", None, pl["PBT"], True)
    add("Current tax", None, pl["CTAX"])
    add("Deferred tax", None, pl["DTAX"])
    add("Net Profit for the Period", None, pl["NET"], True)
    add("Total Other Comprehensive income", None, pl["OCI_TOTAL"])
    add("Total Comprehensive income", None, pl["TCI"], True)
    add("Basic earnings per share", None, pl["EPS"])
    pldf = pd.DataFrame(rows)
    return bsdf, pldf


def _show_statement(df, title):
    st.markdown(f"**{title}**")
    d = df.copy()
    bold = d.pop("_b")
    _table(d, column_config={"Amount": st.column_config.NumberColumn(format="%.2f"), "Previous year": st.column_config.NumberColumn(format="%.2f")})


def _render_build(key_prefix="fsbuild"):
    P = key_prefix
    st.markdown(
        '<div class="fsgap-hero"><h2>Automated Financial Statements</h2>'
        "<p>Upload the original trial balance, let the tool group every ledger into statement lines (you can correct any group), add closing stock "
        "and adjustments, download the statements in your standard format, and match them with the workbook that does not tally.</p></div>",
        unsafe_allow_html=True,
    )
    t1, t2, t3, t4, t5 = st.tabs(["1 · Trial balance", "2 · Grouping", "3 · Manual inputs", "4 · Financial statements", "5 · Match & find mistakes"])

    # ------------------------------------------------------------------ 1. upload
    with t1:
        up = st.file_uploader("Original trial balance (.xlsx / .xlsm / .csv)", type=["xlsx", "xlsm", "csv"], key=_k(P, "tbfile"))
        if not up:
            st.info("Any Tally / ERP trial balance with Debit and Credit columns works. A multi-sheet workbook is fine - you pick the sheet. "
                    "Group heading rows and total rows are recognised and skipped.")
        else:
            data = up.getvalue()
            h = hashlib.md5(data).hexdigest()
            try:
                sheets = _tb_probe(h, data, up.name)
                c1, c2, c3 = st.columns([2, 2, 1])
                sheet = c1.selectbox("Trial balance sheet", sheets, key=_k(P, "tbsheet"))
                lay = _tb_layout(h, data, up.name, sheet)
                if lay is None:
                    st.error("No 'Debit' / 'Credit' column headings found in this sheet.")
                else:
                    labels = [bld.pair_label(p) for p in lay["pairs"]]
                    pi = c2.selectbox("Debit / Credit columns to use", range(len(labels)), format_func=lambda i: labels[i],
                                      index=bld.default_pair(lay), key=_k(P, "tbpair") + h + sheet,
                                      help="Pick the closing-balance pair of the combined (all branches) figures.")
                    units = list(bld.UNITS)
                    guess = lay["unit_hint"] if lay["pairs"][pi]["ctx"] and "lakh" in lay["pairs"][pi]["ctx"].lower() else "Rupees"
                    in_unit = c3.selectbox("Amounts in", units, index=units.index(guess), key=_k(P, "inunit") + h + sheet + str(pi))
                    sig = (h, sheet, pi, in_unit)
                    if _ss(P, "sig") != sig:
                        df, info = bld.read_tb(data, up.name, sheet=sheet, pair=pi, input_unit=in_unit)
                        g = bld.auto_group(df, _ss(P, "memory"))
                        cs = g[g["SubGroup"] == "Closing stock (manual)"]
                        cs_val = 0.0
                        if len(cs):
                            f = bld.UNITS[in_unit] / bld.UNITS["Lakhs"]
                            cs_val = float((cs["Debit"] - cs["Credit"]).sum() * f)
                            g.loc[cs.index, "Include"] = False
                        _set(P, "sig", sig)
                        _set(P, "gen", int(_ss(P, "gen", 0)) + 1)
                        _set(P, "grouped", g)
                        _set(P, "info", info)
                        _set(P, "in_unit", in_unit)
                        _set(P, "closing_default", cs_val)
                        _set(P, "closing", cs_val)
                        _set(P, "company", lay["company"].title() if lay["company"] else "")
                        _set(P, "end", lay["end"])
                        for k in ("result", "recon"):
                            _set(P, k, None)
                    g = _ss(P, "grouped")
                    if g is not None:
                        n_hi = int((g["Confidence"] >= 0.7).sum())
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Ledgers read", len(g))
                        m2.metric("Confidently grouped", n_hi)
                        m3.metric("To review (low confidence)", len(g) - n_hi)
                        gi = g[g["Include"]]
                        diff = float(gi["Debit"].sum() - gi["Credit"].sum())
                        m4.metric("Debit - Credit", f"{diff:,.2f}" + ("" if in_unit == "Rupees" else f" {in_unit}"))
                        if abs(diff) > 1:
                            st.warning("The trial balance does not balance (debits differ from credits). The statements can only balance if the trial balance does - "
                                       "check that the right columns were picked.")
                        st.caption(f"Detected: company **{_ss(P, 'company') or '-'}**, period end **{_ss(P, 'end') or '-'}**, unit **{in_unit}**. "
                                   "Continue in tab 2 to review the grouping.")
                        mem = st.file_uploader("Optional: saved grouping (CSV from an earlier run) to apply to matching ledger names", type=["csv"], key=_k(P, "memfile"))
                        if mem is not None and st.button("Apply saved grouping", key=_k(P, "applymem")):
                            m = bld.mapping_from_csv(mem.getvalue())
                            gg = _ss(P, "grouped").copy()
                            n = 0
                            for i, r in gg.iterrows():
                                k = bld.norm(r["Ledger"])
                                if k in m:
                                    gg.at[i, "Line"], gg.at[i, "SubGroup"], gg.at[i, "Reason"], gg.at[i, "Confidence"] = m[k][0], m[k][1], "saved mapping", 1.0
                                    n += 1
                            _set(P, "grouped", gg)
                            st.success(f"Applied to {n} ledgers.")
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not read the trial balance: {e}")
                with st.expander("Details"):
                    st.code(traceback.format_exc())

    g = _ss(P, "grouped")
    in_unit = _ss(P, "in_unit", "Rupees")

    # ------------------------------------------------------------------ 2. grouping
    with t2:
        if g is None:
            st.info("Upload a trial balance in tab 1 first.")
        else:
            st.markdown("Every ledger is placed in a **statement line** (balance sheet / P&L line of your standard format) and a **sub group** (the row shown in the note). "
                        "Change any of them in the table; the statements update from here.")
            f1, f2, f3, f4 = st.columns([2, 2, 1, 1])
            q = f1.text_input("Search ledger / sub group", key=_k(P, "q"))
            lf = f2.multiselect("Show only lines", [bld.id_to_choice(i) for i in bld.LINE], key=_k(P, "lf"))
            low = f3.checkbox("Low confidence only", key=_k(P, "low"))
            inc_only = f4.checkbox("Included only", key=_k(P, "inc"))
            view = g.copy()
            if q:
                m = view["Ledger"].str.contains(q, case=False, regex=False) | view["SubGroup"].str.contains(q, case=False, regex=False)
                view = view[m]
            if lf:
                view = view[view["Line"].isin([bld.choice_to_id(x) for x in lf])]
            if low:
                view = view[view["Confidence"] < 0.7]
            if inc_only:
                view = view[view["Include"]]
            ev = view[["Include", "Row", "Ledger", "Debit", "Credit", "Line", "SubGroup", "Confidence", "Reason"]].copy()
            ev["Line"] = ev["Line"].map(bld.id_to_choice)
            fkey = f"{len(view)}|{q}|{lf}|{low}|{inc_only}"
            edited = _editor(
                ev, hide_index=False, height=460, key=_k(P, "ed") + str(abs(hash(fkey))),
                disabled=["Row", "Ledger", "Debit", "Credit", "Confidence", "Reason"],
                column_config={
                    "Include": st.column_config.CheckboxColumn("Include", help="Untick to leave a ledger out of the statements"),
                    "Line": st.column_config.SelectboxColumn("Statement line", options=bld.line_choices(), width="large", required=True),
                    "SubGroup": st.column_config.TextColumn("Sub group (note row)", width="medium"),
                    "Debit": st.column_config.NumberColumn(format="%.2f"), "Credit": st.column_config.NumberColumn(format="%.2f"),
                    "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0, max_value=1, format="%.2f"),
                })
            changed = False
            for i in edited.index:
                new_line, new_sub, new_inc = bld.choice_to_id(edited.at[i, "Line"]), edited.at[i, "SubGroup"] or "", bool(edited.at[i, "Include"])
                if (g.at[i, "Line"], g.at[i, "SubGroup"], g.at[i, "Include"]) != (new_line, new_sub, new_inc):
                    if g.at[i, "Line"] != new_line or g.at[i, "SubGroup"] != new_sub:
                        g.at[i, "Reason"], g.at[i, "Confidence"] = "manually corrected", 1.0
                    g.at[i, "Line"], g.at[i, "SubGroup"], g.at[i, "Include"] = new_line, new_sub, new_inc
                    changed = True
            if changed:
                _set(P, "grouped", g)
                _set(P, "result", None)
            with st.expander("Bulk change: move all ledgers matching a text to one line"):
                b1, b2, b3, b4 = st.columns([2, 3, 2, 1])
                txt = b1.text_input("Ledger name contains", key=_k(P, "bt"))
                tgt = b2.selectbox("Move to statement line", [bld.id_to_choice(i) for i in bld.LINE], key=_k(P, "bl"))
                sub = b3.text_input("Sub group", key=_k(P, "bs"))
                n = int(g["Ledger"].str.contains(txt, case=False, regex=False).sum()) if txt else 0
                if b4.button(f"Apply ({n})", key=_k(P, "bb"), disabled=n == 0):
                    mm = g["Ledger"].str.contains(txt, case=False, regex=False)
                    g.loc[mm, "Line"] = bld.choice_to_id(tgt)
                    if sub:
                        g.loc[mm, "SubGroup"] = sub
                    g.loc[mm, ["Reason", "Confidence"]] = ["manually corrected", 1.0]
                    _set(P, "grouped", g)
                    st.rerun()
            st.download_button("Save this grouping (CSV) - reuse it next period", bld.mapping_to_csv(g), file_name="ledger_grouping.csv", mime="text/csv", key=_k(P, "dlmap"))
            # live summary by line / sub group
            f = bld.UNITS[in_unit] / bld.UNITS["Lakhs"]
            s = g[g["Include"]].copy()
            s["Net (Dr-Cr), Lakhs"] = (s["Debit"] - s["Credit"]) * f
            sm = s.groupby(["Line", "SubGroup"], as_index=False).agg(Ledgers=("Ledger", "count"), **{"Net (Dr-Cr), Lakhs": ("Net (Dr-Cr), Lakhs", "sum")})
            sm["Statement line"] = sm["Line"].map(lambda x: bld.LINE[x]["label"])
            order = {k: i for i, k in enumerate(bld.LINE)}
            sm = sm.sort_values(by="Line", key=lambda c: c.map(order))
            st.markdown("**Grouping summary**")
            _table(sm[["Statement line", "SubGroup", "Ledgers", "Net (Dr-Cr), Lakhs"]],
                   column_config={"Net (Dr-Cr), Lakhs": st.column_config.NumberColumn(format="%.2f")})
            susp = float(s[s["Line"] == "SUSP"]["Net (Dr-Cr), Lakhs"].sum())
            if abs(susp) > 0.005:
                st.warning(f"Inter-branch / suspense ledgers do not net to nil ({susp:,.2f} Lakhs). The difference is shown in other current assets / liabilities; "
                           "branch eliminations normally net to zero, so check the branch balances.")

    # ------------------------------------------------------------------ 3. manual inputs
    with t3:
        if g is None:
            st.info("Upload a trial balance in tab 1 first.")
        else:
            c1, c2, c3 = st.columns(3)
            gen = _ss(P, "gen", 0)
            company = c1.text_input("Company name", value=_ss(P, "company", ""), key=_k(P, "co") + str(gen))
            cin = c2.text_input("CIN", key=_k(P, "cin"))
            address = c3.text_input("Registered office (printed under the P&L title)", key=_k(P, "addr"))
            d1, d2, d3 = st.columns(3)
            default_end = _ss(P, "end") or dt.date.today()
            end = d1.date_input("Statement date (period end)", value=default_end, key=_k(P, "end_in") + str(gen))
            plabel = d2.selectbox("P&L column heading", ["For the Half Year Ended", "For the Quarter Ended", "For the Nine Months Ended", "For the Year Ended"], key=_k(P, "pl"))
            unit = d3.selectbox("Present amounts in", ["Lakhs", "Rupees", "Thousands", "Crores"], key=_k(P, "unit"))
            e1, e2, e3 = st.columns(3)
            auditor = e1.text_input("Auditor firm (signature block)", key=_k(P, "aud"))
            place = e2.text_input("Place", key=_k(P, "place"))
            fv = e3.number_input("Face value per share (Rs)", min_value=0.0, value=10.0, key=_k(P, "fv"))
            h1, h2 = st.columns(2)
            shares = h1.number_input("Weighted average shares (0 = derive from share capital)", min_value=0.0, value=0.0, step=1.0, key=_k(P, "sh"))
            oci_rate = h2.number_input("Tax rate on OCI (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.01, key=_k(P, "ocir"),
                                       help="Leave 0 if the OCI ledger in the trial balance is already net of tax.")
            st.markdown("##### Closing stock")
            cdef = _ss(P, "closing_default", 0.0)
            f_out = bld.UNITS["Lakhs"] / bld.UNITS[unit]
            cs = st.number_input(f"Closing stock value ({unit})", value=float(_ss(P, "closing", 0.0)) * f_out if cdef else 0.0, step=1.0, format="%.2f", key=_k(P, "cs") + unit + str(gen),
                                 help="Closing stock is not a ledger balance - it is posted as a debit to Inventories and a credit to Cost of materials consumed.")
            if cdef:
                st.caption(f"Pre-filled from a 'closing stock' row found below the trial balance total ({cdef:,.2f} Lakhs); that row is excluded from the ledgers to avoid counting it twice.")
            st.markdown("##### Adjustment entries (optional)")
            st.caption("Journal lines posted on top of the trial balance (e.g. unrecorded wages payable, expected credit loss). Debits and credits should be equal.")
            adj0 = _ss(P, "adj")
            if adj0 is None:
                adj0 = bld.default_adjustments()
            adj = _editor(adj0, num_rows="dynamic", key=_k(P, "adj_ed"),
                                 column_config={"Line": st.column_config.SelectboxColumn("Statement line", options=bld.line_choices(), width="large"),
                                                "Debit": st.column_config.NumberColumn(format="%.2f"), "Credit": st.column_config.NumberColumn(format="%.2f")})
            _set(P, "adj", adj)
            if len(adj):
                d, c = float(adj["Debit"].fillna(0).sum()), float(adj["Credit"].fillna(0).sum())
                if abs(d - c) > 0.005:
                    st.warning(f"Adjustments are not balanced: debit {d:,.2f} vs credit {c:,.2f}. The balance sheet will show this difference.")
            meta = {"company": company, "cin": cin, "address": address, "end": end, "period_label": plabel, "auditor": auditor, "place": place,
                    "face_value": fv, "shares": shares, "oci_tax_rate": oci_rate / 100.0}
            _set(P, "meta", meta)
            _set(P, "unit_out", unit)
            _set(P, "closing_val", cs)
            st.success("Inputs saved. Open tab 4 to see the statements.")

    # ------------------------------------------------------------------ 4. statements
    with t4:
        if g is None:
            st.info("Upload a trial balance in tab 1 first.")
        else:
            meta = _ss(P, "meta") or {"company": _ss(P, "company", ""), "end": _ss(P, "end"), "period_label": "For the Half Year Ended"}
            unit = _ss(P, "unit_out", "Lakhs")
            cs_val = _ss(P, "closing_val", _ss(P, "closing", 0.0))
            adj = _ss(P, "adj")
            res = bld.build(g, meta, closing_stock=cs_val, adjustments=adj, unit=unit, input_unit=in_unit)
            _set(P, "result", res)
            ck = res["checks"]
            diff = res["bs"]["TOT_ASSETS"] - res["bs"]["TOT_EL"]
            if abs(diff) < 0.01 * bld.UNITS["Lakhs"] / bld.UNITS[unit] + 1e-6:
                st.success(f"The automated balance sheet tallies: Total Assets = Total Equity & Liabilities = {res['bs']['TOT_ASSETS']:,.2f} {unit}.")
            else:
                st.error(f"The automated balance sheet differs by {diff:,.2f} {unit}. Causes: trial balance not balanced ({ck['tb_diff']:,.2f}), unbalanced adjustments, or ledgers left out.")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"Total Assets ({unit})", _fmt_money(res["bs"]["TOT_ASSETS"]))
            m2.metric(f"Revenue ({unit})", _fmt_money(res["pl"]["REV"]))
            m3.metric(f"Net profit ({unit})", _fmt_money(res["pl"]["NET"]))
            m4.metric(f"Other equity ({unit})", _fmt_money(res["bs"]["OE"]))
            bsdf, pldf = _statement_frames(res)
            cA, cB = st.columns(2)
            with cA:
                _show_statement(bsdf, "Statement of Assets and Liabilities")
            with cB:
                _show_statement(pldf, "Statement of Profit and Loss")
            with st.expander("Notes - sub group breakdown"):
                for lid, items in res["notes"].items():
                    if items:
                        nt = bld.LINE[lid]["note"]
                        st.markdown(f"**{('Note ' + str(nt) + ' - ') if nt else ''}{bld.LINE[lid]['label']}**")
                        _table(pd.DataFrame(items, columns=["Sub group", unit]), column_config={unit: st.column_config.NumberColumn(format="%.2f")})
            st.download_button("Download financial statements (Excel, same layout as your workbook)", bld.to_excel(res), file_name="automated_financial_statements.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_k(P, "dlfs"), type="primary")
            st.caption("The Excel file has BS, PL, Note-Assets, Note-Liab, PL Note, a 'TB Mapped' sheet (every ledger with its line / sub group) and Checks. "
                       "Figures are live SUMIFS formulas over 'TB Mapped': change a ledger's line there and the statements follow.")

    # ------------------------------------------------------------------ 5. match
    with t5:
        res = _ss(P, "result")
        if g is None or res is None:
            st.info("Build the statements first (tabs 1-4).")
        else:
            st.markdown("Upload the **original financial statements workbook that does not tally**. The tool compares each balance sheet line with the statements built from the "
                        "trial balance, shows which lines carry the difference, and traces them to the ledgers / schedules behind them.")
            src = st.radio("Original workbook", ["Upload a workbook", "It is the same workbook as the trial balance"], horizontal=True, key=_k(P, "osrc"),
                           disabled=not (_ss(P, "sig") and st.session_state.get(_k(P, "tbfile")) is not None and not str(st.session_state[_k(P, "tbfile")].name).lower().endswith(".csv")))
            data_o, name_o = None, None
            if src.startswith("Upload"):
                uo = st.file_uploader("Original financial statements (.xlsx / .xlsm)", type=["xlsx", "xlsm"], key=_k(P, "ofile"))
                if uo:
                    data_o, name_o = uo.getvalue(), uo.name
            else:
                tf = st.session_state[_k(P, "tbfile")]
                data_o, name_o = tf.getvalue(), tf.name
            if data_o:
                try:
                    ho = hashlib.md5(data_o).hexdigest()
                    bk = _load(ho, data_o, name_o)
                    cands = eng.candidate_sheets(bk)
                    if not cands:
                        st.error("No balance sheet found in that workbook.")
                    else:
                        c1, c2, c3 = st.columns([2, 2, 1])
                        sheet = c1.selectbox("Balance sheet", cands, index=cands.index(eng.pick_default_sheet(bk, cands)), key=_k(P, "osheet"))
                        stt = eng.find_structure(bk, sheet)
                        end = (_ss(P, "meta") or {}).get("end")
                        dc = bld.pick_column(stt, end)
                        colopts = [c["col"] for c in stt["cols"]]
                        col = c2.selectbox("Period column to match", colopts, index=colopts.index(dc), key=_k(P, "ocol") + sheet,
                                           format_func=lambda c: next(" ".join(x["header"].split()) or x["letter"] for x in stt["cols"] if x["col"] == c))
                        tol = c3.number_input("Tolerance", min_value=0.0, value=0.01, step=0.01, format="%.3f", key=_k(P, "otol"))
                        tbs_default = eng.detect_tb_sheets(bk)
                        tbs = st.multiselect("Trial balance sheets inside the original workbook (tracing stops there)", [s for s in bk.sheetnames if s != sheet],
                                             default=[s for s in tbs_default if s != sheet], key=_k(P, "otbs") + sheet)
                        if st.button("Match & find mistakes", type="primary", key=_k(P, "orun")):
                            with st.spinner("Comparing line by line and tracing the original formulas..."):
                                scale = bld.UNITS[in_unit] / bld.UNITS[res["unit"]]
                                rc = bld.reconcile(eng, bk, res, g, sheet, col, tbs, tol=tol, scale=scale)
                            _set(P, "recon", rc)
                except Exception as e:  # noqa: BLE001
                    st.error(f"The comparison failed: {e}")
                    with st.expander("Details"):
                        st.code(traceback.format_exc())
            rc = _ss(P, "recon")
            if rc and rc.get("ok"):
                unit = res["unit"]
                m1, m2, m3 = st.columns(3)
                m1.metric("Original: Total Assets", _fmt_money(rc["orig_assets"]) if rc["orig_assets"] is not None else "error")
                m2.metric("Original: Equity & Liabilities", _fmt_money(rc["orig_el"]) if rc["orig_el"] is not None else "error")
                m3.metric("Original difference", _fmt_money(rc["gap"]) if rc["gap"] is not None else "n/a")
                st.markdown(bld.recon_narrative(rc, unit))
                tol = rc["tol"]
                rows = [{"Line": r["label"].strip(), "Note": r["note"], "Original": r["orig"], "From trial balance": r["auto"], "Difference (orig - TB)": r["var"],
                         "Original cell": ", ".join(r["cells"]), "Status": "Error in original" if r["err"] else ("Differs" if abs(r["var"]) > tol else "Agrees")}
                        for r in rc["rows"] if r["cells"] or abs(r["auto"]) > tol]
                dfc = pd.DataFrame(rows)
                show_all = st.checkbox("Show lines that agree", value=False, key=_k(P, "showall"))
                _table(dfc if show_all else dfc[dfc["Status"] != "Agrees"],
                       column_config={k: st.column_config.NumberColumn(format="%.2f") for k in ("Original", "From trial balance", "Difference (orig - TB)")})
                if rc["unmatched"]:
                    st.warning("Original lines that could not be matched with a standard line: " + "; ".join(f"{u['label']} ({u['addr']}) {u['value']:,.2f}" for u in rc["unmatched"]))
                st.markdown("#### Where each difference comes from")
                for r in rc["rows"]:
                    d = rc["drill"].get(r["id"])
                    if not d:
                        continue
                    with st.expander(f"{r['label'].strip()}  -  original {r['orig']:,.2f} vs trial balance {r['auto']:,.2f}  (difference {r['var']:,.2f})", expanded=abs(r["var"]) > 1):
                        if r["err"]:
                            st.error("Formula error in " + ", ".join(r["err_cells"]) + " - the original total cannot be trusted until this is fixed.")
                        if r["id"] == "OE":
                            st.markdown("The original **Other Equity** is rolled forward on the Equity sheet (opening balance + profit + OCI) instead of being read from the trial balance. "
                                        "The trial-balance figure is *Reserves + opening Profit & Loss A/c + profit for the period + OCI*. "
                                        "A difference here usually means the opening balance carried in the roll-forward is not the same as the reserves in the trial balance.")
                        elif d["sheets"] and not d["only_auto"] and not d["only_orig"] and not d["reads_tb"]:
                            st.markdown("This line is **not read from the trial balance**; it comes from: " + ", ".join(f"`{s}`" for s in d["sheets"]) + ".")
                        elif d["sheets"]:
                            st.markdown("This line is fed by: " + ", ".join(f"`{s}`" for s in d["sheets"]) + (f" and {len(d['reads_tb'])} trial-balance row(s)." if d["reads_tb"] else "."))
                        if d["only_orig"]:
                            st.markdown("**Ledgers the original line reads, but automatic grouping puts elsewhere**")
                            _table(pd.DataFrame([{"Ledger": x["ledger"], "Amount (Dr-Cr)": x["net"], "Grouped by automation to": bld.LINE[x["auto_line"]]["label"]} for x in d["only_orig"]]),
                                   column_config={"Amount (Dr-Cr)": st.column_config.NumberColumn(format="%.2f")})
                        if d["only_auto"] and (d["reads_tb"] or not d["sheets"]):
                            st.markdown("**Ledgers grouped here by automation that the original line does not read**")
                            _table(pd.DataFrame([{"Ledger": x["ledger"], "Amount (Dr-Cr)": x["net"]} for x in d["only_auto"]]),
                                   column_config={"Amount (Dr-Cr)": st.column_config.NumberColumn(format="%.2f")})
                        sched = [x for x in d["schedule"] if x["found"]]
                        if sched:
                            sched.sort(key=lambda x: -abs(x["diff"] or 0))
                            st.markdown(f"**Ledger totals that the schedule does not reproduce** ({d['schedule_agree']} ledger(s) agree exactly). "
                                        "Closest schedule figure shown - a difference here is where the schedule and the books disagree:")
                            _table(pd.DataFrame([{"Ledger (trial balance)": x["ledger"], "Ledger amount": x["amount"], "Closest schedule figure": x["found"],
                                                  "Which figure": x["context"], "Schedule - ledger": x["diff"]} for x in sched[:12]]),
                                   column_config={"Ledger amount": st.column_config.NumberColumn(format="%.2f"), "Schedule - ledger": st.column_config.NumberColumn(format="%.2f")})
                        if d["typed"]:
                            st.markdown("**Typed (hard-coded) numbers feeding this line**")
                            _table(pd.DataFrame(d["typed"]), column_config={"value": st.column_config.NumberColumn("Value", format="%.2f")})
                st.download_button("Download reconciliation (Excel)", bld.recon_to_excel(rc, res), file_name="reconciliation_report.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_k(P, "dlrc"))


# =====================================================================================================================
#  Mode 3 : one upload - the company's financials workbook -> automated statements + gap
# =====================================================================================================================
def _render_quick(key_prefix="fsquick"):
    P = key_prefix
    st.markdown(
        '<div class="fsgap-hero"><h2>Automated statements from the company financials</h2>'
        "<p>Upload the company's financial-statements workbook (the one with the trial balance inside). The tool rebuilds the statements from the trial balance, "
        "compares them with your balance sheet and shows what causes the gap. Closing stock, taxes and other items can be entered by hand.</p></div>",
        unsafe_allow_html=True,
    )
    up = st.file_uploader("Company financial statements workbook (.xlsx / .xlsm)", type=["xlsx", "xlsm"], key=_k(P, "file"))
    if not up:
        st.info("The workbook needs a balance sheet sheet and a trial balance sheet (Debit / Credit columns), with formulas and saved values.")
        return
    data = up.getvalue()
    h = hashlib.md5(data).hexdigest()
    try:
        bk = _load(h, data, up.name)
        cands = eng.candidate_sheets(bk)
        if not cands:
            st.error("No balance sheet found in this workbook.")
            return
        c1, c2 = st.columns(2)
        sheet = c1.selectbox("Balance sheet", cands, index=cands.index(eng.pick_default_sheet(bk, cands)), key=_k(P, "sheet"))
        stt = eng.find_structure(bk, sheet)
        colopts = [c["col"] for c in stt["cols"]]
        dc = bld.pick_column(stt, None)
        col = c2.selectbox("Period to check", colopts, index=colopts.index(dc), key=_k(P, "col") + sheet,
                           format_func=lambda c: next(" ".join(x["header"].split()) or x["letter"] for x in stt["cols"] if x["col"] == c))
        hdr = next(x["header"] for x in stt["cols"] if x["col"] == col)
        ds = bld.parse_dates(hdr)
        end = ds[-1] if ds else None
        tb_cands = [s for s in bld.tb_sheet_candidates(data, up.name) if s != sheet]
        auto_tbs = [s for s in eng.detect_tb_sheets(bk) if s != sheet]
        # TB sheet whose title carries the same period end first
        d1, d2 = st.columns(2)
        pref = [s for s in tb_cands if s in auto_tbs] + [s for s in tb_cands if s not in auto_tbs]
        tbsheet = d1.selectbox("Trial balance sheet to rebuild from", pref, key=_k(P, "tbs") + sheet + str(col),
                               help="Pick the trial balance of the same period as the balance sheet column.")
        lay = _tb_layout(h, data, up.name, tbsheet)
        if lay is None:
            st.error("No Debit / Credit columns found on that sheet.")
            return
        labels = [bld.pair_label(p) for p in lay["pairs"]]
        pi = d2.selectbox("Debit / Credit columns", range(len(labels)), format_func=lambda i: labels[i], index=bld.default_pair(lay), key=_k(P, "pair") + tbsheet)
        units = list(bld.UNITS)
        e1, e2, e3 = st.columns(3)
        guess = lay["unit_hint"] if lay["pairs"][pi]["ctx"] and "lakh" in lay["pairs"][pi]["ctx"].lower() else "Rupees"
        in_unit = e1.selectbox("Trial balance amounts in", units, index=units.index(guess), key=_k(P, "inu") + tbsheet + str(pi))
        ou = bld.sheet_unit_hint(bk, sheet) or "Lakhs"
        out_unit = e2.selectbox("Statements are in", units, index=units.index(ou), key=_k(P, "outu") + sheet)
        basis = e3.radio("Grouping logic", ["Learn from this workbook's formulas", "Automatic (system knowledge)"], key=_k(P, "basis"),
                         help="'Learn' puts each ledger on the line your workbook already reads it into, so the comparison isolates what is wrong beyond grouping. "
                              "'Automatic' uses the tool's own rules, which also exposes grouping mistakes in the workbook.")

        df, info = bld.read_tb(data, up.name, sheet=tbsheet, pair=pi, input_unit=in_unit)
        g = bld.auto_group(df)
        cs_rows = g[g["SubGroup"] == "Closing stock (manual)"]
        tb_cs = 0.0
        if len(cs_rows):
            tb_cs = float((cs_rows["Debit"] - cs_rows["Credit"]).sum() * bld.UNITS[in_unit] / bld.UNITS[out_unit])
            g.loc[cs_rows.index, "Include"] = False
        tbs_trace = list(dict.fromkeys([tbsheet] + auto_tbs))
        n_learned = 0
        if basis.startswith("Learn"):
            g, n_learned = bld.learn_grouping(eng, bk, g, sheet, col, tbs_trace, tbsheet)
            st.caption(f"{n_learned} of {len(g)} ledgers were grouped exactly as the workbook's own formulas read them; the rest use the automatic grouping.")
        # original inventory value, to pre-fill closing stock when the TB has none
        orig_inv = None
        for ln in bld.match_original_lines(eng, bk, stt, col):
            if ln["match"] == "INVENT" and ln["value"] is not None:
                orig_inv = float(ln["value"])
        diff_tb = float(g[g["Include"]]["Debit"].sum() - g[g["Include"]]["Credit"].sum())
        if abs(diff_tb) > 1e-6 * max(1.0, float(g["Debit"].sum())):
            st.warning("The selected trial balance columns do not balance; check the Debit / Credit pair and the period.")

        st.markdown("##### Entries needed to finish the statements (manual)")
        m1, m2, m3 = st.columns(3)
        cs_default = tb_cs if abs(tb_cs) > 1e-9 else (orig_inv or 0.0)
        cs = m1.number_input(f"Closing stock ({out_unit})", value=float(round(cs_default, 4)), step=1.0, format="%.2f", key=_k(P, "cs") + tbsheet + str(pi) + str(col),
                             help="Posted as Dr Inventories / Cr Cost of materials consumed.")
        m1.caption("Pre-filled from the trial balance's closing-stock row." if abs(tb_cs) > 1e-9 else ("Pre-filled from Inventories in the original balance sheet." if orig_inv else "Enter the value."))
        ctx_in_tb = float(g[(g["Include"]) & (g["Line"] == "CTAX")][["Debit", "Credit"]].apply(lambda r: r["Debit"] - r["Credit"], axis=1).sum() * bld.UNITS[in_unit] / bld.UNITS[out_unit])
        dtx_in_tb = float(g[(g["Include"]) & (g["Line"] == "DTAX")][["Debit", "Credit"]].apply(lambda r: r["Debit"] - r["Credit"], axis=1).sum() * bld.UNITS[in_unit] / bld.UNITS[out_unit])
        ct = m2.number_input(f"Additional current tax ({out_unit})", value=0.0, step=1.0, format="%.2f", key=_k(P, "ct"),
                             help="Tax expense not already in the trial balance. Posted Dr Current tax / Cr Provision for income tax.")
        m2.caption(f"Current tax already in the trial balance: {ctx_in_tb:,.2f}")
        dt_ = m3.number_input(f"Additional deferred tax: charge + / credit - ({out_unit})", value=0.0, step=1.0, format="%.2f", key=_k(P, "dt"),
                              help="Posted Dr Deferred tax / Cr Deferred tax asset (reversed for a credit).")
        m3.caption(f"Deferred tax already in the trial balance: {dtx_in_tb:,.2f}")
        with st.expander("Other adjustment entries (optional)"):
            adj = _editor(bld.default_adjustments(), num_rows="dynamic", key=_k(P, "adj"),
                          column_config={"Line": st.column_config.SelectboxColumn("Statement line", options=bld.line_choices(), width="large"),
                                         "Debit": st.column_config.NumberColumn(format="%.2f"), "Credit": st.column_config.NumberColumn(format="%.2f")})
        with st.expander("Statement details (names, EPS)"):
            f1, f2, f3 = st.columns(3)
            company = f1.text_input("Company name", value=(lay["company"].title() if lay["company"] else ""), key=_k(P, "co") + tbsheet)
            cin = f2.text_input("CIN", key=_k(P, "cin"))
            fv = f3.number_input("Face value per share (Rs)", min_value=0.0, value=10.0, key=_k(P, "fv"))
            plabel = st.selectbox("P&L column heading", ["For the Half Year Ended", "For the Quarter Ended", "For the Nine Months Ended", "For the Year Ended"], key=_k(P, "pl"))
        tol = st.number_input("Tolerance", min_value=0.0, value=0.01, step=0.01, format="%.3f", key=_k(P, "tol"))
        if not st.button("Build automated statements and find the gap", type="primary", key=_k(P, "run")):
            return
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not prepare the analysis: {e}")
        with st.expander("Details"):
            st.code(traceback.format_exc())
        return

    with st.spinner("Building statements and comparing with the original..."):
        try:
            meta = {"company": company, "cin": cin, "end": end, "period_label": plabel, "face_value": fv, "oci_tax_rate": 0.0}
            adjustments = bld.manual_postings(ct, dt_, adj)
            res = bld.build(g, meta, closing_stock=cs, adjustments=adjustments, unit=out_unit, input_unit=in_unit)
            scale = bld.UNITS[in_unit] / bld.UNITS[out_unit]
            rc = bld.reconcile(eng, bk, res, g, sheet, col, tbs_trace, tol=tol, scale=scale)
        except Exception as e:  # noqa: BLE001
            st.error(f"The analysis failed: {e}")
            with st.expander("Details"):
                st.code(traceback.format_exc())
            return
    _show_recon(P, res, rc, g)


def _pl_health_warning(res, g):
    n_pl = int(g[(g["Include"]) & (g["Line"].map(lambda x: bld.LINE[x]["stmt"] == "PL"))].shape[0])
    pl = res["pl"]
    if n_pl == 0 or (abs(pl["REV"]) < 1e-9 and abs(pl["EXPENSES"]) < 1e-9):
        st.warning("No income or expense ledgers reached the Profit & Loss. Check the grouping (tab 2 in the trial-balance mode): "
                   "P&L ledgers must be on lines such as Revenue, Employee benefits, Other expenses.")


def _show_recon(P, res, rc, g):
    _pl_health_warning(res, g)
    unit = res["unit"]
    diff = res["bs"]["TOT_ASSETS"] - res["bs"]["TOT_EL"]
    if abs(diff) < 0.02:
        st.success(f"The automated balance sheet tallies at {res['bs']['TOT_ASSETS']:,.2f} {unit}.")
    else:
        st.error(f"The automated balance sheet differs by {diff:,.2f} {unit}: the trial balance used does not balance or an entry above is one-sided.")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Original: Total Assets", _fmt_money(rc["orig_assets"]) if rc["orig_assets"] is not None else "error")
    m2.metric("Original: Equity & Liabilities", _fmt_money(rc["orig_el"]) if rc["orig_el"] is not None else "error")
    m3.metric("Original difference", _fmt_money(rc["gap"]) if rc["gap"] is not None else "n/a")
    m4.metric("Automated Total Assets", _fmt_money(res["bs"]["TOT_ASSETS"]))
    st.markdown(bld.recon_narrative(rc, unit))
    tol = rc["tol"]
    rows = [{"Line": r["label"].strip(), "Note": r["note"], "Original": r["orig"], "Automated": r["auto"], "Difference (orig - auto)": r["var"],
             "Original cell": ", ".join(r["cells"]), "Status": "Error in original" if r["err"] else ("Differs" if abs(r["var"]) > tol else "Agrees")}
            for r in rc["rows"] if r["cells"] or abs(r["auto"]) > tol]
    dfc = pd.DataFrame(rows)
    show_all = st.checkbox("Show lines that agree", value=False, key=_k(P, "showall"))
    _table(dfc if show_all else dfc[dfc["Status"] != "Agrees"],
           column_config={k: st.column_config.NumberColumn(format="%.2f") for k in ("Original", "Automated", "Difference (orig - auto)")})
    if rc["unmatched"]:
        st.warning("Original lines that could not be matched with a standard line: " + "; ".join(f"{u['label']} ({u['addr']}) {u['value']:,.2f}" for u in rc["unmatched"]))
    st.markdown("#### Where each difference comes from")
    for r in rc["rows"]:
        d = rc["drill"].get(r["id"])
        if not d:
            continue
        with st.expander(f"{r['label'].strip()}  -  original {r['orig']:,.2f} vs automated {r['auto']:,.2f}  (difference {r['var']:,.2f})", expanded=abs(r["var"]) > 1):
            if r["err"]:
                st.error("Formula error in " + ", ".join(r["err_cells"]) + " - the original total cannot be trusted until this is fixed.")
            if r["id"] == "OE":
                st.markdown("The original **Other Equity** is rolled forward on a separate sheet (opening balance + profit + OCI) instead of being read from the trial balance. "
                            "A difference here usually means the opening balance in the roll-forward differs from the reserves in the trial balance, or the profit differs.")
            elif d["sheets"] and not d["only_auto"] and not d["only_orig"] and not d["reads_tb"]:
                st.markdown("This line is **not read from the trial balance**; it comes from: " + ", ".join(f"`{s}`" for s in d["sheets"]) + ".")
            elif d["sheets"]:
                st.markdown("This line is fed by: " + ", ".join(f"`{s}`" for s in d["sheets"]) + (f" and {len(d['reads_tb'])} trial-balance row(s)." if d["reads_tb"] else "."))
            if d["only_orig"]:
                st.markdown("**Ledgers the original line reads, but the automated grouping puts elsewhere**")
                _table(pd.DataFrame([{"Ledger": x["ledger"], "Amount (Dr-Cr)": x["net"], "Automated line": bld.LINE[x["auto_line"]]["label"]} for x in d["only_orig"]]),
                       column_config={"Amount (Dr-Cr)": st.column_config.NumberColumn(format="%.2f")})
            if d["only_auto"] and (d["reads_tb"] or not d["sheets"]):
                st.markdown("**Ledgers grouped here by the automated statements that the original line does not read**")
                _table(pd.DataFrame([{"Ledger": x["ledger"], "Amount (Dr-Cr)": x["net"]} for x in d["only_auto"]]),
                       column_config={"Amount (Dr-Cr)": st.column_config.NumberColumn(format="%.2f")})
            sched = [x for x in d["schedule"] if x["found"]]
            if sched:
                sched.sort(key=lambda x: -abs(x["diff"] or 0))
                st.markdown(f"**Ledger totals the schedule does not reproduce** ({d['schedule_agree']} ledger(s) agree exactly). Closest schedule figure shown:")
                _table(pd.DataFrame([{"Ledger (trial balance)": x["ledger"], "Ledger amount": x["amount"], "Closest schedule figure": x["found"],
                                      "Which figure": x["context"], "Schedule - ledger": x["diff"]} for x in sched[:12]]),
                       column_config={"Ledger amount": st.column_config.NumberColumn(format="%.2f"), "Schedule - ledger": st.column_config.NumberColumn(format="%.2f")})
            if d["typed"]:
                st.markdown("**Typed (hard-coded) numbers feeding this line**")
                _table(pd.DataFrame(d["typed"]), column_config={"value": st.column_config.NumberColumn("Value", format="%.2f")})
    bsdf, pldf = _statement_frames(res)
    with st.expander("Automated statements (preview)"):
        cA, cB = st.columns(2)
        with cA:
            _show_statement(bsdf, "Statement of Assets and Liabilities")
        with cB:
            _show_statement(pldf, "Statement of Profit and Loss")
    x1, x2 = st.columns(2)
    x1.download_button("Download automated financial statements (Excel)", bld.to_excel(res), file_name="automated_financial_statements.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_k(P, "dlfs"), type="primary")
    x2.download_button("Download gap report (Excel)", bld.recon_to_excel(rc, res), file_name="gap_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_k(P, "dlrc"))


# =====================================================================================================================
#  Mode 4 : previous-year financials (grouping method) + current-year trial balance -> current-year financials
# =====================================================================================================================
def _grouping_editor(P, g, in_unit, show_origin=True):
    """review / correct the grouping of the current trial balance (returns True when something changed)"""
    f1, f2, f3, f4 = st.columns([2, 2, 1, 1])
    q = f1.text_input("Search ledger / sub group", key=_k(P, "q"))
    only = f2.multiselect("Show only", ["New ledger", "Similar to previous year", "Previous year"], default=["New ledger", "Similar to previous year"], key=_k(P, "only"))
    low = f3.checkbox("Low confidence only", key=_k(P, "low"))
    inc_only = f4.checkbox("Included only", key=_k(P, "inc"))
    view = g.copy()
    if q:
        view = view[view["Ledger"].str.contains(q, case=False, regex=False) | view["SubGroup"].str.contains(q, case=False, regex=False)]
    if only:
        view = view[view["Origin"].isin(only)]
    if low:
        view = view[view["Confidence"] < 0.7]
    if inc_only:
        view = view[view["Include"]]
    cols = ["Include", "Row", "Ledger", "Debit", "Credit", "Origin", "Line", "SubGroup", "Confidence", "Reason"]
    ev = view[cols].copy()
    ev["Line"] = ev["Line"].map(bld.id_to_choice)
    fkey = f"{len(view)}|{q}|{only}|{low}|{inc_only}"
    st.caption(f"{len(view)} ledger(s) shown. Ledgers found in the previous-year financials keep last year's group; new ones are grouped by system knowledge and need a quick review.")
    edited = _editor(
        ev, hide_index=False, height=420, key=_k(P, "ed") + str(abs(hash(fkey))),
        disabled=["Row", "Ledger", "Debit", "Credit", "Origin", "Confidence", "Reason"],
        column_config={
            "Include": st.column_config.CheckboxColumn("Include"),
            "Line": st.column_config.SelectboxColumn("Statement line", options=bld.line_choices(), width="large", required=True),
            "SubGroup": st.column_config.TextColumn("Sub group (note row)", width="medium"),
            "Debit": st.column_config.NumberColumn(format="%.2f"), "Credit": st.column_config.NumberColumn(format="%.2f"),
            "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0, max_value=1, format="%.2f"),
        })
    changed = False
    for i in edited.index:
        nl, ns, ni = bld.choice_to_id(edited.at[i, "Line"]), edited.at[i, "SubGroup"] or "", bool(edited.at[i, "Include"])
        if (g.at[i, "Line"], g.at[i, "SubGroup"], g.at[i, "Include"]) != (nl, ns, ni):
            if g.at[i, "Line"] != nl or g.at[i, "SubGroup"] != ns:
                g.at[i, "Reason"], g.at[i, "Confidence"] = "manually corrected", 1.0
            g.at[i, "Line"], g.at[i, "SubGroup"], g.at[i, "Include"] = nl, ns, ni
            changed = True
    return changed


def _render_prev(key_prefix="fsprev"):
    P = key_prefix
    st.markdown(
        '<div class="fsgap-hero"><h2>Current-year financials from last year\'s grouping</h2>'
        "<p>Upload the previous-year financial statements and the current-year trial balance. Every ledger is grouped the way last year's statements "
        "grouped it; ledgers that did not exist last year are grouped by system knowledge and flagged for review. Last year's figures become the comparative column.</p></div>",
        unsafe_allow_html=True,
    )
    units = list(bld.UNITS)
    # ------------------------------------------------------------ A. previous-year financials
    st.markdown("##### A. Previous-year financial statements")
    up = st.file_uploader("Previous-year financials workbook (with its trial balance inside)", type=["xlsx", "xlsm"], key=_k(P, "pfile"))
    if not up:
        st.info("Upload last year's workbook first. It must contain the balance sheet and the trial balance it was prepared from.")
        return
    data = up.getvalue()
    h = hashlib.md5(data).hexdigest()
    try:
        bk = _load(h, data, up.name)
        cands = eng.candidate_sheets(bk)
        if not cands:
            st.error("No balance sheet found in the previous-year workbook.")
            return
        c1, c2, c3 = st.columns(3)
        sheet = c1.selectbox("Balance sheet", cands, index=cands.index(eng.pick_default_sheet(bk, cands)), key=_k(P, "psheet"))
        stt = eng.find_structure(bk, sheet)
        colopts = [c["col"] for c in stt["cols"]]
        col = c2.selectbox("Previous-year column (comparative)", colopts, index=colopts.index(bld.pick_column(stt, None)), key=_k(P, "pcol") + sheet,
                           format_func=lambda c: next(" ".join(x["header"].split()) or x["letter"] for x in stt["cols"] if x["col"] == c))
        hdr = next(x["header"] for x in stt["cols"] if x["col"] == col)
        pds = bld.parse_dates(hdr)
        prev_end = pds[-1] if pds else None
        tb_cands = [s for s in bld.tb_sheet_candidates(data, up.name) if s != sheet]
        auto_tbs = [s for s in eng.detect_tb_sheets(bk) if s != sheet]
        pref = [s for s in tb_cands if s in auto_tbs] + [s for s in tb_cands if s not in auto_tbs]
        ptb = c3.selectbox("Previous-year trial balance sheet", pref, key=_k(P, "ptb") + sheet + str(col))
        d1, d2, d3 = st.columns(3)
        play = _tb_layout(h, data, up.name, ptb)
        if play is None:
            st.error("No Debit / Credit columns found on the previous-year trial balance sheet.")
            return
        plabels = [bld.pair_label(p) for p in play["pairs"]]
        ppi = d1.selectbox("Previous TB Debit / Credit columns", range(len(plabels)), format_func=lambda i: plabels[i], index=bld.default_pair(play), key=_k(P, "ppair") + ptb)
        pguess = play["unit_hint"] if play["pairs"][ppi]["ctx"] and "lakh" in play["pairs"][ppi]["ctx"].lower() else "Rupees"
        p_in = d2.selectbox("Previous TB amounts in", units, index=units.index(pguess), key=_k(P, "pinu") + ptb + str(ppi))
        pl_sheets = bld.find_pl_sheets(bk)
        pl_sheet = d3.selectbox("Previous-year Profit & Loss sheet", ["(none)"] + pl_sheets, index=1 if pl_sheets else 0, key=_k(P, "ppl") + sheet)
        p_unit_stmt = bld.sheet_unit_hint(bk, sheet) or "Lakhs"
        psig = (h, sheet, col, ptb, ppi, p_in, pl_sheet)
        if _ss(P, "psig") != psig:
            with st.spinner("Learning the grouping from the previous-year financials..."):
                pdf, pinfo = bld.read_tb(data, up.name, sheet=ptb, pair=ppi, input_unit=p_in)
                pg = bld.auto_group(pdf)
                pg.loc[pg["SubGroup"] == "Closing stock (manual)", "Include"] = False
                tbs_trace = list(dict.fromkeys([ptb] + auto_tbs))
                pg, n_bs = bld.learn_grouping(eng, bk, pg, sheet, col, tbs_trace, ptb)
                n_pl = 0
                if pl_sheet != "(none)":
                    pcol = bld.pick_pl_column(bk, pl_sheet, prev_end)
                    if pcol:
                        pg, n_pl = bld.learn_pl_grouping(eng, bk, pg, pl_sheet, pcol, tbs_trace, ptb)
                prev_vals = bld.previous_statement_values(eng, bk, sheet, col, p_unit_stmt, p_unit_stmt)
            _set(P, "psig", psig)
            _set(P, "prev_g", pg)
            _set(P, "prev_vals", prev_vals)
            _set(P, "prev_unit", p_unit_stmt)
            _set(P, "prev_learn", (n_bs, n_pl, len(pg)))
            _set(P, "csig", None)
        n_bs, n_pl, n_all = _ss(P, "prev_learn")
        st.success(f"Previous-year grouping learned: {n_bs} balance sheet ledgers and {n_pl} P&L ledgers grouped exactly as last year's formulas read them "
                   f"(of {n_all} ledgers). The rest use system knowledge.")
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the previous-year financials: {e}")
        with st.expander("Details"):
            st.code(traceback.format_exc())
        return

    # ------------------------------------------------------------ B. current-year trial balance
    st.markdown("##### B. Current-year trial balance")
    cu = st.file_uploader("Current-year trial balance (.xlsx / .xlsm / .csv)", type=["xlsx", "xlsm", "csv"], key=_k(P, "cfile"))
    if not cu:
        return
    cdata = cu.getvalue()
    ch = hashlib.md5(cdata).hexdigest()
    try:
        sheets = _tb_probe(ch, cdata, cu.name)
        e1, e2, e3, e4 = st.columns([2, 2, 1, 1])
        csheet = e1.selectbox("Trial balance sheet", sheets, key=_k(P, "csheet"))
        clay = _tb_layout(ch, cdata, cu.name, csheet)
        if clay is None:
            st.error("No Debit / Credit columns found on that sheet.")
            return
        clabels = [bld.pair_label(p) for p in clay["pairs"]]
        cpi = e2.selectbox("Debit / Credit columns", range(len(clabels)), format_func=lambda i: clabels[i], index=bld.default_pair(clay), key=_k(P, "cpair") + ch + csheet)
        cguess = clay["unit_hint"] if clay["pairs"][cpi]["ctx"] and "lakh" in clay["pairs"][cpi]["ctx"].lower() else "Rupees"
        c_in = e3.selectbox("Amounts in", units, index=units.index(cguess), key=_k(P, "cinu") + ch + csheet + str(cpi))
        out_unit = e4.selectbox("Statements in", units, index=units.index(_ss(P, "prev_unit", "Lakhs")), key=_k(P, "outu"))
        csig = (_ss(P, "psig"), ch, csheet, cpi, c_in)
        if _ss(P, "csig") != csig:
            cdf, cinfo = bld.read_tb(cdata, cu.name, sheet=csheet, pair=cpi, input_unit=c_in)
            g, missing = bld.apply_previous(cdf, _ss(P, "prev_g"))
            cs_rows = g[g["SubGroup"] == "Closing stock (manual)"]
            cs_val = float((cs_rows["Debit"] - cs_rows["Credit"]).sum() * bld.UNITS[c_in] / bld.UNITS[out_unit]) if len(cs_rows) else 0.0
            if len(cs_rows):
                g.loc[cs_rows.index, "Include"] = False
            _set(P, "csig", csig)
            _set(P, "g", g)
            _set(P, "missing", missing)
            _set(P, "cs_tb", cs_val)
            _set(P, "gen", int(_ss(P, "gen", 0)) + 1)
            _set(P, "cend", clay["end"])
            _set(P, "ccompany", clay["company"])
        g = _ss(P, "g")
        gen = _ss(P, "gen", 0)
        vc = g["Origin"].value_counts()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Same as previous year", int(vc.get("Previous year", 0)))
        m2.metric("Similar name", int(vc.get("Similar to previous year", 0)))
        m3.metric("New ledgers (system knowledge)", int(vc.get("New ledger", 0)))
        m4.metric("In last year, not in this TB", len(_ss(P, "missing")))
        gi = g[g["Include"]]
        if abs(float(gi["Debit"].sum() - gi["Credit"].sum())) > 1e-6 * max(1.0, float(g["Debit"].sum())):
            st.warning("The selected trial balance columns do not balance; check the Debit / Credit pair.")
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the current-year trial balance: {e}")
        with st.expander("Details"):
            st.code(traceback.format_exc())
        return

    # ------------------------------------------------------------ C. review grouping
    st.markdown("##### C. Review the grouping (new ledgers first)")
    if _grouping_editor(P, g, c_in):
        _set(P, "g", g)
    ms = _ss(P, "missing")
    if len(ms):
        with st.expander(f"Ledgers with a balance last year that are not in this trial balance ({len(ms)})"):
            _table(pd.DataFrame({"Ledger": ms["Ledger"], "Previous Debit": ms["Debit"], "Previous Credit": ms["Credit"], "Last year's line": ms["Line"].map(lambda x: bld.LINE[x]["label"])}),
                   column_config={"Previous Debit": st.column_config.NumberColumn(format="%.2f"), "Previous Credit": st.column_config.NumberColumn(format="%.2f")})
    with st.expander("Bulk change: move all ledgers matching a text to one line"):
        b1, b2, b3, b4 = st.columns([2, 3, 2, 1])
        txt = b1.text_input("Ledger name contains", key=_k(P, "bt"))
        tgt = b2.selectbox("Move to statement line", [bld.id_to_choice(i) for i in bld.LINE], key=_k(P, "bl"))
        sub = b3.text_input("Sub group", key=_k(P, "bs"))
        n = int(g["Ledger"].str.contains(txt, case=False, regex=False).sum()) if txt else 0
        if b4.button(f"Apply ({n})", key=_k(P, "bb"), disabled=n == 0):
            mm = g["Ledger"].str.contains(txt, case=False, regex=False)
            g.loc[mm, "Line"] = bld.choice_to_id(tgt)
            if sub:
                g.loc[mm, "SubGroup"] = sub
            g.loc[mm, ["Reason", "Confidence"]] = ["manually corrected", 1.0]
            _set(P, "g", g)
            st.rerun()
    st.download_button("Save this grouping (CSV)", bld.mapping_to_csv(g), file_name="ledger_grouping.csv", mime="text/csv", key=_k(P, "dlmap"))

    # ------------------------------------------------------------ D. manual inputs
    st.markdown("##### D. Entries needed to finish the statements (manual)")
    pv = _ss(P, "prev_vals", {})
    f_in = bld.UNITS[c_in] / bld.UNITS[out_unit]
    m1, m2, m3 = st.columns(3)
    cs_tb = _ss(P, "cs_tb", 0.0)
    cs = m1.number_input(f"Closing stock ({out_unit})", value=float(round(cs_tb, 4)), step=1.0, format="%.2f", key=_k(P, "cs") + str(gen) + out_unit,
                         help="Posted as Dr Inventories / Cr Cost of materials consumed.")
    m1.caption("Pre-filled from the trial balance." if abs(cs_tb) > 1e-9 else f"Enter the value. Last year's inventories: {pv.get('INVENT', 0):,.2f}")
    ctx_in = float(sum((r["Debit"] - r["Credit"]) for _, r in g[(g["Include"]) & (g["Line"] == "CTAX")].iterrows()) * f_in)
    dtx_in = float(sum((r["Debit"] - r["Credit"]) for _, r in g[(g["Include"]) & (g["Line"] == "DTAX")].iterrows()) * f_in)
    ct = m2.number_input(f"Additional current tax ({out_unit})", value=0.0, step=1.0, format="%.2f", key=_k(P, "ct"),
                         help="Not already in the trial balance. Dr Current tax / Cr Provision for income tax.")
    m2.caption(f"Current tax already in the trial balance: {ctx_in:,.2f}")
    dt_ = m3.number_input(f"Additional deferred tax: charge + / credit - ({out_unit})", value=0.0, step=1.0, format="%.2f", key=_k(P, "dt"),
                          help="Dr Deferred tax / Cr Deferred tax asset (reversed for a credit).")
    m3.caption(f"Deferred tax already in the trial balance: {dtx_in:,.2f}")
    with st.expander("Other adjustment entries (optional)"):
        adj = _editor(bld.default_adjustments(), num_rows="dynamic", key=_k(P, "adj"),
                      column_config={"Line": st.column_config.SelectboxColumn("Statement line", options=bld.line_choices(), width="large"),
                                     "Debit": st.column_config.NumberColumn(format="%.2f"), "Credit": st.column_config.NumberColumn(format="%.2f")})
    with st.expander("Statement details (names, date, EPS)"):
        k1, k2, k3 = st.columns(3)
        company = k1.text_input("Company name", value=(_ss(P, "ccompany") or "").title(), key=_k(P, "co") + str(gen))
        cin = k2.text_input("CIN", key=_k(P, "cin"))
        end = k3.date_input("Statement date", value=_ss(P, "cend") or dt.date.today(), key=_k(P, "end") + str(gen))
        k4, k5 = st.columns(2)
        plabel = k4.selectbox("P&L column heading", ["For the Year Ended", "For the Half Year Ended", "For the Quarter Ended", "For the Nine Months Ended"], key=_k(P, "pl"))
        fv = k5.number_input("Face value per share (Rs)", min_value=0.0, value=10.0, key=_k(P, "fv"))
    prev_scaled = {k: v * bld.UNITS[_ss(P, "prev_unit", "Lakhs")] / bld.UNITS[out_unit] for k, v in pv.items()}
    meta = {"company": company, "cin": cin, "end": end, "prev_end": prev_end, "period_label": plabel, "face_value": fv, "oci_tax_rate": 0.0}
    adjustments = bld.manual_postings(ct, dt_, adj)
    res = bld.build(g, meta, closing_stock=cs, adjustments=adjustments, unit=out_unit, input_unit=c_in, prev_bs=prev_scaled)

    # ------------------------------------------------------------ E. result
    st.markdown("##### E. Current-year financial statements")
    _pl_health_warning(res, g)
    diff = res["bs"]["TOT_ASSETS"] - res["bs"]["TOT_EL"]
    if abs(diff) < 0.02:
        st.success(f"The balance sheet tallies: Total Assets = Total Equity & Liabilities = {res['bs']['TOT_ASSETS']:,.2f} {out_unit}.")
    else:
        st.error(f"The balance sheet differs by {diff:,.2f} {out_unit}: the trial balance does not balance or an entry above is one-sided.")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(f"Revenue ({out_unit})", _fmt_money(res["pl"]["REV"]))
    k2.metric(f"Net profit ({out_unit})", _fmt_money(res["pl"]["NET"]))
    k3.metric(f"Total Assets ({out_unit})", _fmt_money(res["bs"]["TOT_ASSETS"]))
    k4.metric(f"Other equity ({out_unit})", _fmt_money(res["bs"]["OE"]))
    bsdf, pldf = _statement_frames(res)
    cA, cB = st.columns(2)
    with cA:
        _show_statement(bsdf, "Statement of Assets and Liabilities")
    with cB:
        _show_statement(pldf, "Statement of Profit and Loss")
    # continuity with last year
    ob = sum(res["oe_parts"].values())
    cont = [{"Check": "Equity share capital: last year closing vs this trial balance", "Last year": prev_scaled.get("SC", 0.0), "This year (TB)": res["bs"]["SC"]},
            {"Check": "Other equity: last year closing vs opening reserves in this trial balance (before this year's profit)", "Last year": prev_scaled.get("OE", 0.0), "This year (TB)": ob}]
    for r_ in cont:
        r_["Difference"] = r_["This year (TB)"] - r_["Last year"]
    st.markdown("**Continuity with last year**")
    _table(pd.DataFrame(cont), column_config={k: st.column_config.NumberColumn(format="%.2f") for k in ("Last year", "This year (TB)", "Difference")})
    if any(abs(r_["Difference"]) > 0.01 for r_ in cont):
        st.caption("A difference means the trial balance opening does not agree with last year's published closing (dividend, prior-period adjustment, or last year's statements need correcting).")
    with st.expander("Year-on-year movement by line (look for unusual swings)"):
        mv = []
        for lid, lbl, n_, nat, sec in bld.BS_LINES:
            a_, b_ = res["bs"][lid], prev_scaled.get(lid, 0.0)
            if abs(a_) > 0.005 or abs(b_) > 0.005:
                mv.append({"Line": lbl.strip(), "Previous year": b_, "This year": a_, "Change": a_ - b_, "Change %": ((a_ - b_) / abs(b_) * 100) if abs(b_) > 1e-9 else None})
        _table(pd.DataFrame(mv), column_config={k: st.column_config.NumberColumn(format="%.2f") for k in ("Previous year", "This year", "Change", "Change %")})
    with st.expander("Notes - sub group breakdown"):
        for lid, items in res["notes"].items():
            if items:
                nt = bld.LINE[lid]["note"]
                st.markdown(f"**{('Note ' + str(nt) + ' - ') if nt else ''}{bld.LINE[lid]['label']}**")
                _table(pd.DataFrame(items, columns=["Sub group", out_unit]), column_config={out_unit: st.column_config.NumberColumn(format="%.2f")})
    st.download_button("Download current-year financial statements (Excel, with previous-year column)", bld.to_excel(res), file_name="current_year_financial_statements.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_k(P, "dlfs"), type="primary")


def render(key_prefix: str = "fsgap"):
    st.markdown(CSS, unsafe_allow_html=True)
    modes = ["Upload company financials: build automated statements and find the gap",
             "Find the gap in financial statements that do not tally",
             "Build automated financial statements from a separate trial balance",
             "Previous-year financials + current-year trial balance: prepare current-year financials"]
    mode = st.radio("What do you want to do?", modes, key=_k(key_prefix, "mode"))
    if mode == modes[0]:
        _render_quick(key_prefix + "_quick")
    elif mode == modes[1]:
        _render_gap(key_prefix)
    elif mode == modes[2]:
        _render_build(key_prefix + "_build")
    else:
        _render_prev(key_prefix + "_prev")


if __name__ == "__main__":
    st.set_page_config(page_title="Balance Sheet Gap Analyser", page_icon=None, layout="wide")
    render()
