"""
Paired Comparison Topline Generator – Streamlit app
Run:  streamlit run app.py
"""
import datetime as dt

import pandas as pd
import streamlit as st

import topline_engine as te

st.set_page_config(page_title="Topline Generator", page_icon="📊", layout="wide")
st.title("📊 Topline Generator – Paired comparison / product tests")
st.caption("Upload an Excel file with an INDEX sheet (data map) and a DATA sheet. "
           "Candidates are tested against a benchmark and the topline is exported to Excel "
           "with colour-coded significant gaps.")

# --------------------------------------------------------------------------- #
# 1. File
# --------------------------------------------------------------------------- #
up = st.file_uploader("Data file (.xlsx)", type=["xlsx", "xls"])
if not up:
    st.info("Waiting for a file…")
    st.stop()


@st.cache_data(show_spinner="Reading workbook…")
def _read(file_bytes):
    import io
    return te.load_workbook_sheets(io.BytesIO(file_bytes))


sheets = _read(up.getvalue())
names = list(sheets)
c1, c2 = st.columns(2)
idx_default = next((i for i, n in enumerate(names) if "INDEX" in n.upper()), 0)
dat_default = next((i for i, n in enumerate(names) if n.upper().startswith("DATA")), min(1, len(names) - 1))
idx_sheet = c1.selectbox("Index sheet (data map)", names, index=idx_default)
data_sheet = c2.selectbox("Data sheet", names, index=dat_default)

data, variables = te.build_dataset(sheets[idx_sheet], sheets[data_sheet])
cols = list(data.columns)
lab = lambda c: f"{c}" if variables[c].label == c else f"{c}  –  {variables[c].label}"
n_linked = sum(1 for c in cols if variables[c].codes)
st.success(f"{len(data)} rows × {len(cols)} variables · {n_linked} variables linked to the index with modality labels.")

guess = te.guess_roles(cols, variables)

# --------------------------------------------------------------------------- #
# 2. Variables
# --------------------------------------------------------------------------- #
st.header("1 · Variables")
c1, c2 = st.columns(2)
product_col = c1.selectbox("Product code column", cols, index=cols.index(guess["product"]), format_func=lab)
id_col = c2.selectbox("Consumer ID column", cols, index=cols.index(guess["id"]), format_func=lab)

q_cols = [c for c in cols if c not in (product_col, id_col)]
ms = lambda title, key: st.multiselect(title, q_cols, default=[c for c in guess[key] if c in q_cols],
                                       format_func=lab, key=f"ms_{key}")
liking = ms("Liking", "liking")
strength = ms("Strength / intensity", "strength")
fragrance = ms("Fragrance attributes", "fragrance")
other = ms("Other attributes", "other")
pref_cols = st.multiselect("Preference question (optional)", q_cols,
                           default=[c for c in guess["preference"] if c in q_cols], format_func=lab)

# --------------------------------------------------------------------------- #
# 3. Products
# --------------------------------------------------------------------------- #
st.header("2 · Benchmark & candidates")
pvar = variables[product_col]
codes = sorted(data[product_col].dropna().unique())
plab = lambda c: f"{pvar.code_label(c)}  (code {te._fmt_code(c)}, n={int((data[product_col] == c).sum())})"
c1, c2 = st.columns([1, 2])
benchmark = c1.selectbox("Benchmark", codes, format_func=plab)
candidates = c2.multiselect("Candidates", [c for c in codes if c != benchmark],
                            default=[c for c in codes if c != benchmark], format_func=plab)

# --------------------------------------------------------------------------- #
# 4. Statistics
# --------------------------------------------------------------------------- #
st.header("3 · Metrics & significance tests")
c1, c2, c3 = st.columns(3)
metrics = c1.multiselect("Metrics", te.ALL_METRICS, default=["Mean", "Std", "Top box", "Top 2 boxes", "Bottom 2 boxes"])
detail = c1.radio("Output detail", ["Metrics only", "Metrics + % per modality"], index=1) != "Metrics only"

TAILS = {"2-tailed t-test": "two", "1-tailed – superior (candidate > benchmark)": "greater",
         "1-tailed – inferior (candidate < benchmark)": "less"}
mean_tail = TAILS[c2.radio("Test on means", list(TAILS))]
equal_var = c2.checkbox("Independent samples: assume equal variances (Student)", value=True,
                        help="Only used when consumers rated a single product. Untick for Welch.")

prop_test = c3.radio("Test on proportions", ["Z-test", "McNemar"])
prop_tail_choice = c3.radio("Proportions tail", ["Same as means", "2-tailed"], horizontal=True)
prop_tail = mean_tail if prop_tail_choice == "Same as means" else "two"

levels = [99, 95, 90]
if st.checkbox("Also colour the 80% level", value=False):
    levels.append(80)

# legend preview
leg = "".join(
    f"<tr><td style='padding:2px 10px'>{lv}%</td>"
    f"<td style='background:#{te.COLORS[('pos', lv)]};width:90px'></td>"
    f"<td style='background:#{te.COLORS[('neg', lv)]};width:90px'></td></tr>" for lv in levels)
st.markdown(f"<table style='border-collapse:collapse;font-size:13px'><tr><th>Threshold</th>"
            f"<th>Positive</th><th>Negative</th></tr>{leg}</table>", unsafe_allow_html=True)

all_q = liking + strength + fragrance + other
with st.expander("Advanced – scale direction & missing codes"):
    reversed_cols = st.multiselect(
        "Scales where the LOWEST code is the best answer (top box = lowest codes)", all_q, format_func=lab,
        default=[c for c in all_q if "LONGLAST" in c.upper()])
    inverted_cols = st.multiselect(
        "Negative attributes – a higher score is worse (green/red flipped)", all_q, format_func=lab,
        default=[c for c in all_q if any(k in c.upper() for k in ("OVERWHELM", "CHEMICAL", "OLD FASHION"))])
    st.caption("Codes labelled NA / don't know / refused, or ≥ 98, are automatically excluded from means and box scores.")
    extra = st.text_input("Extra codes to exclude from means/boxes for all questions (comma separated)", "")
    extra_codes = {float(x) for x in extra.replace(";", ",").split(",") if x.strip().replace(".", "").isdigit()}

# --------------------------------------------------------------------------- #
# 5. Preference
# --------------------------------------------------------------------------- #
pref_settings = None
if pref_cols:
    st.header("4 · Preference question")
    pcol = pref_cols[0]
    if len(pref_cols) > 1:
        pcol = st.selectbox("Preference question to analyse", pref_cols, format_func=lab)
    pv = variables[pcol]
    pcodes = sorted(set(pv.codes) | set(data[pcol].dropna().unique()))
    fl = lambda c: f"{te._fmt_code(c)} - {pv.code_label(c)}"
    c1, c2, c3 = st.columns(3)
    a = c1.selectbox("Modality A (e.g. prefer candidate)", pcodes, index=0, format_func=fl)
    b = c2.selectbox("Modality B (e.g. prefer benchmark)", pcodes, index=min(1, len(pcodes) - 1), format_func=fl)
    np_opts = [None] + pcodes
    np_code = c3.selectbox("No preference modality", np_opts, index=min(3, len(np_opts) - 1) if len(pcodes) > 2 else 0,
                           format_func=lambda c: "— none —" if c is None else fl(c))
    ptest = st.radio("Significance test for preference", list(te.PREF_TESTS),
                     format_func=te.PREF_TESTS.get)
    by_prod = st.checkbox("Also split preference by product code column", value=False,
                          help="Useful only if the preference answer is stored on the row of a specific product/pair code.")
    pref_settings = dict(col=pcol, a=a, b=b, np=np_code, test=ptest, by_product=by_prod,
                         a_label=pv.code_label(a), b_label=pv.code_label(b))

# --------------------------------------------------------------------------- #
# 6. Run
# --------------------------------------------------------------------------- #
st.header("5 · Generate")
title = st.text_input("Report title", value=f"Topline – {up.name.rsplit('.', 1)[0]}")
pvals = st.checkbox("Add a sheet with confidence levels (1 − p)", value=True)

if not candidates:
    st.warning("Select at least one candidate.")
    st.stop()
if not all_q and not pref_settings:
    st.warning("Select at least one question.")
    st.stop()

if st.button("▶ Compute topline", type="primary"):
    s = te.Settings(
        product_col=product_col, id_col=id_col, benchmark=benchmark, candidates=candidates,
        sections={k: v for k, v in [("Liking", liking), ("Strength", strength),
                                    ("Fragrance attributes", fragrance), ("Other attributes", other)] if v},
        metrics=metrics, show_modalities=detail, mean_tail=mean_tail,
        prop_test="mcnemar" if prop_test == "McNemar" else "z", prop_tail=prop_tail,
        equal_var=equal_var, levels=tuple(levels), reversed_cols=reversed_cols,
        inverted_cols=inverted_cols, extra_missing={c: extra_codes for c in all_q})
    with st.spinner("Computing…"):
        rows, products, meta = te.compute_topline(data, variables, s)
        pref_res = None
        if pref_settings:
            pref_res = te.compute_preference(
                data, variables, pref_settings["col"], id_col, product_col, pref_settings["a"],
                pref_settings["b"], pref_settings["np"], pref_settings["test"], levels,
                by_product=pref_settings["by_product"])
        xls = te.write_excel(rows, products, meta, s, variables, pref_res, pref_settings,
                             title=title, include_pvalues=pvals)

    st.info(f"Design detected: **{'paired (same consumers rated several products)' if meta['paired'] else 'independent samples'}** "
            f"· {meta['n_ids']} consumers")
    for n in meta["notes"]:
        st.warning(n)

    # on-screen preview with colours
    plabel = lambda p: pvar.code_label(p) + (" (BM)" if p == benchmark else "")
    recs, styles = [], []
    for r in rows:
        rec = {"Section": r["section"], "Question": r["question"], "Metric": r["metric"]}
        sty = {"Section": "", "Question": "", "Metric": ""}
        for p in products:
            c = r["cells"][p]
            v = c.value
            if v is None or pd.isna(v):
                txt = ""
            elif c.kind == "pct":
                txt = f"{v:.0%}"
            elif c.kind == "n":
                txt = f"{int(v)}"
            else:
                txt = f"{v:.2f}"
            rec[plabel(p)] = txt
            fill, white = te._fill_for(c)
            sty[plabel(p)] = (f"background-color:#{fill.fgColor.rgb[-6:]};color:{'white' if white else 'black'};"
                              f"font-weight:bold") if fill else ""
        recs.append(rec)
        styles.append(sty)
    df = pd.DataFrame(recs)
    sdf = pd.DataFrame(styles)
    st.dataframe(df.style.apply(lambda _: sdf, axis=None), use_container_width=True, height=600, hide_index=True)

    if pref_res:
        st.subheader("Preference")
        pv = variables[pref_settings["col"]]
        tab = pd.DataFrame({g["name"]: {f"{te._fmt_code(c)} - {pv.code_label(c)}": f"{g['pct'].get(c, 0):.0%}"
                                        for c in g["codes"]} for g in pref_res})
        tab.loc["Base (n)"] = [g["n"] for g in pref_res]
        tab.loc["Confidence"] = [("" if pd.isna(g["p"]) else f"{1 - g['p']:.1%}") for g in pref_res]
        st.dataframe(tab, use_container_width=True)

    st.download_button("⬇ Download Excel topline", xls,
                       file_name=f"Topline_{dt.date.today():%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
