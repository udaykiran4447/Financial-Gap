"""
Balance Sheet Gap Analyser  -  Streamlit front end for fs_gap_engine.py

Run standalone :  streamlit run fs_gap_analyser_app.py
Integrate later:  import fs_gap_analyser_app as fsgap ; fsgap.render(key_prefix="fsgap")
                  (all widget / session keys are namespaced with key_prefix)
"""
import hashlib
import traceback

import pandas as pd
import streamlit as st

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


def _k(prefix, name):
    return f"{prefix}_{name}"


def render(key_prefix: str = "fsgap"):
    st.markdown(CSS, unsafe_allow_html=True)
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


if __name__ == "__main__":
    st.set_page_config(page_title="Balance Sheet Gap Analyser", page_icon=None, layout="wide")
    render()
