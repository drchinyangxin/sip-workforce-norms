"""
PKP & PPKP Distribution - SIMPLE VIEW (provisional, current data only)
----------------------------------------------------------------------
Shows how evenly the EXISTING PPKP/PKP match the recorded activity across districts.
It is NOT a staffing requirement.

Two independent ways of measuring workload are compared:
  Method A  Share of national activity  - every task counts equally (adjustable by task group), no minutes needed
  Method B  Minutes per task            - placeholder minutes, editable
A district is flagged strongly only when BOTH methods agree.

Run:  streamlit run app_simple.py   (Excel file next to the script, or upload in the sidebar)
"""
import glob
import io
import math
import os

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(layout="wide", page_title="PKP & PPKP Distribution - Simple View", page_icon="⚖️")

# key, Malay label, short label, header keyword, group, placeholder min/unit, % by PKP, unit note
TASKS = [
    ("kawalan", "Kawalan penyakit berjangkit", "Kawalan PB", "kawalan penyakit berjangkit", "Disease control", 10, 0, "count as recorded"),
    ("sampel", "Persampelan penyakit berjangkit", "Persampelan", "persampelan penyakit", "Disease control", 45, 0, "samples"),
    ("siasat_kes", "Siasatan kes & kontak", "Siasatan kes", "siasatan kes dan kontak", "Disease control", 30, 0, "cases/contacts"),
    ("advokasi", "Advokasi & pendidikan kesihatan", "Advokasi", "advokasi", "Governance & training", 90, 0, "sessions"),
    ("risikan", "Risikan undang-undang kesihatan awam", "Risikan UU", "risikan undang", "Legal", 120, 0, "intelligence ops"),
    ("notis", "Notis kompaun / kompaun", "Kompaun", "notis kompaun", "Legal", 30, 0, "notices"),
    ("kertas", "Kertas siasatan (kesalahan UU KA)", "Kertas siasatan", "kertas siasatan", "Legal", 240, 15, "investigation papers"),
    ("osh_audit", "Audit KKP di tempat kerja", "Audit KKP", "audit keselamatan", "Environmental / OSH", 240, 0, "audits"),
    ("osh_siasat", "Notifikasi & siasatan kemalangan/penyakit pekerjaan", "Siasat KKP", "notifikasi dan siasatan kemalangan", "Environmental / OSH", 180, 0, "investigations"),
    ("premis", "Pemeriksaan kesihatan persekitaran di premis", "Periksa premis", "pemeriksaan kesihatan persekitaran", "Environmental / OSH", 30, 0, "premises"),
    ("bilik", "Bilik gerakan wabak/bencana dibuka", "Bilik gerakan", "bilik gerakan", "Disease control", 600, 20, "activations"),
    ("survelans", "Laporan aktiviti survelans", "Survelans", "laporan aktiviti survelans", "Disease control", 30, 0, "reports"),
    ("latihan", "Latihan teknikal / amalan profesional PKP/PPKP", "Latihan", "latihan teknikal", "Governance & training", 240, 10, "sessions"),
    ("mayat", "Import/eksport mayat, tisu, bahan patogenik", "Mayat/patogen", "mayat", "Points of entry (PMA)", 60, 0, "consignments"),
    ("kapal", "Pemeriksaan & penilaian risiko kapal/vessel/penumpang", "Kapal/PMA", "kapal", "Points of entry (PMA)", 0.1, 0, "CHECK UNIT: looks like passenger counts"),
    ("makanan_import", "Pemeriksaan & persampelan produk makanan import", "Makanan import", "produk makanan import", "Points of entry (PMA)", 60, 0, "consignments"),
    ("penilaian", "Penilaian risiko kesihatan persekitaran, PB & bencana", "Penilaian risiko", "penilaian risiko kesihatan persekitaran", "Governance & training", 20, 10, "assessments"),
    ("makanan_premis", "Penilaian risiko premis makanan", "Risiko premis mkn", "penilaian risiko premis makanan", "Environmental / OSH", 30, 0, "premises"),
    ("air", "Pemantauan kualiti air minum", "Kualiti air", "pemantauan kualiti air", "Environmental / OSH", 45, 0, "samples/visits"),
    ("pewasapan", "Audit pewasapan", "Audit wasap", "audit pewasapan", "Legal", 180, 0, "audits"),
]
KEYS = [t[0] for t in TASKS]
LABEL = {t[0]: t[1] for t in TASKS}
SHORT = {t[0]: t[2] for t in TASKS}
GROUPS = ["Disease control", "Legal", "Environmental / OSH", "Points of entry (PMA)", "Governance & training"]
ALLOC = ["PKD", "PKB", "PKK"]


# ----------------------------------------------------------------------------- data
def _clean(x):
    s = " ".join(str(x).split())
    return "" if s.startswith("Unnamed") else s


@st.cache_data(show_spinner=False)
def load_data(file_bytes, path):
    src = io.BytesIO(file_bytes) if file_bytes else path
    raw = pd.read_excel(src, header=[0, 1])
    heads = [(_clean(a), _clean(b)) for a, b in raw.columns]

    def find(pred, what):
        hits = [i for i, (a, b) in enumerate(heads) if pred(a.lower(), b.lower())]
        if len(hits) != 1:
            raise ValueError(f"Column for '{what}' matched {len(hits)} columns (expected 1). Is this the same Excel template?")
        return hits[0]

    idx = {
        "Negeri": find(lambda a, b: a == "negeri", "Negeri"),
        "Daerah": find(lambda a, b: a == "daerah", "Daerah"),
        "Kategori": find(lambda a, b: a == "kategori", "Kategori"),
        "PKP_tetap": find(lambda a, b: a.startswith("bil pkp") and "kontrak" not in a, "PKP tetap"),
        "PKP_kontrak": find(lambda a, b: a.startswith("bil pkp") and "kontrak" in a, "PKP kontrak"),
        "PPKP_tetap": find(lambda a, b: a.startswith("bil ppkp") and "kontrak" not in a, "PPKP tetap"),
        "PPKP_kontrak": find(lambda a, b: a.startswith("bil ppkp") and "kontrak" in a, "PPKP kontrak"),
    }
    for key, _, _, kw, *_ in TASKS:
        idx[key] = find(lambda a, b, kw=kw: a.startswith("beban") and kw in b, kw)
    df = pd.DataFrame({k: raw.iloc[:, i] for k, i in idx.items()})
    for c in ["Negeri", "Daerah", "Kategori"]:
        df[c] = df[c].astype("string").str.strip()
    df = df[df["Daerah"].notna() & (df["Daerah"].str.lower() != "total") & (df["Kategori"] != "Total")].copy()
    for c in ["PKP_tetap", "PKP_kontrak", "PPKP_tetap", "PPKP_kontrak"] + KEYS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in ["Negeri", "Daerah", "Kategori"]:
        df[c] = df[c].astype(str)
    return df.reset_index(drop=True)


def find_default_file():
    here = os.path.dirname(os.path.abspath(__file__))
    for folder in (here, os.getcwd()):
        hits = sorted(glob.glob(os.path.join(folder, "Data*Perjawatan*.xlsx")))
        if hits:
            return hits[-1]
    return None


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


# ----------------------------------------------------------------------------- header + sidebar
st.title("⚖️ PKP & PPKP Distribution - Simple View")
st.caption("Provisional view from current data only | KKM")
st.warning("**Provisional.** This shows how evenly the *existing* officers match the *recorded* activity. "
           "It does **not** say how many officers are needed - that requires time-per-task data (see tab *What we still need*).")

st.sidebar.header("📁 Data")
upload = st.sidebar.file_uploader("Excel file (optional if next to the script)", type=["xlsx"])
path = find_default_file()
if upload is None and path is None:
    st.info("Place the Excel file next to the script or upload it in the sidebar.")
    st.stop()
try:
    full = load_data(upload.getvalue() if upload else None, None if upload else path)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

st.sidebar.header("⚙️ Settings")
use_contract = st.sidebar.checkbox("Count contract (2025) staff", True)
floor = st.sidebar.number_input("Minimum PPKP per district", 0, 20, 3, help="Each district keeps at least this many before the rest is shared by workload.")
pct_tol = st.sidebar.slider("Ignore gaps smaller than (% of fair share)", 0, 50, 15, help="Small gaps are treated as 'balanced'.") / 100
abs_tol = 2.0
show_method = st.sidebar.radio("Charts show fair share from", ["Average of both methods", "Method A only", "Method B only"])

st.sidebar.header("Method A - how much does each task group count?")
st.sidebar.caption("1 = normal, 0 = ignore, 3 = counts triple. Every task counts equally within the weights below.")
gw = {g: st.sidebar.slider(g, 0, 5, 1) for g in GROUPS}

full["PKP"] = full["PKP_tetap"] + (full["PKP_kontrak"] if use_contract else 0)
full["PPKP"] = full["PPKP_tetap"] + (full["PPKP_kontrak"] if use_contract else 0)
work = full[full["Kategori"].isin(ALLOC)].reset_index(drop=True)
n = len(work)
C = work[KEYS].to_numpy(float)
cur = work["PPKP"].to_numpy(float)
pool = float(cur.sum())

# ----------------------------------------------------------------------------- Method B minutes (advanced)
with st.expander("⚙️ Advanced: minutes per task (used by Method B) - placeholders, please replace with real values", expanded=False):
    base = pd.DataFrame([{"key": t[0], "Task": t[1], "Group": t[4], "Unit": t[7], "Minutes_per_unit": t[5], "PKP_pct": t[6]} for t in TASKS])
    edited = st.data_editor(base, hide_index=True, disabled=["key", "Task", "Group", "Unit"], key="min_table",
                            column_config={"key": None,
                                           "Minutes_per_unit": st.column_config.NumberColumn("Minutes per unit", min_value=0.0, format="%.2f"),
                                           "PKP_pct": st.column_config.NumberColumn("% done by PKP", min_value=0, max_value=100, format="%d")})
    st.caption("The vessel/passenger row is very small (0.1 min) because those counts look like passengers, not inspections. Confirm the unit before relying on it.")
minutes = edited["Minutes_per_unit"].to_numpy(float)
pkp_share = edited["PKP_pct"].to_numpy(float) / 100

# ----------------------------------------------------------------------------- model
floor_eff = min(floor, pool / n)
distributable = pool - floor_eff * n
vol = C.sum(axis=0)
share = np.divide(C, vol, out=np.zeros_like(C), where=vol > 0)           # district share of national volume, per task
w_task = np.array([gw[t[4]] for t in TASKS], float) * (vol > 0)
idxA = share @ w_task / w_task.sum() if w_task.sum() > 0 else np.full(n, 1 / n)
fairA = floor_eff + idxA * distributable
demB = C @ (minutes * (1 - pkp_share))
idxB = demB / demB.sum() if demB.sum() > 0 else np.full(n, 1 / n)
fairB = floor_eff + idxB * distributable


def classify(fair):
    tol = np.maximum(abs_tol, pct_tol * fair)
    return np.where(cur < fair - tol, -1, np.where(cur > fair + tol, 1, 0))


sA, sB = classify(fairA), classify(fairB)
VERDICTS = ["🔴 Short (both methods)", "🟢 Surplus (both methods)", "⚪ Methods disagree", "🟠 Leaning short", "🟡 Leaning surplus", "⚪ Balanced"]
verdict = np.select([(sA == -1) & (sB == -1), (sA == 1) & (sB == 1), sA * sB == -1, sA + sB < 0, sA + sB > 0], VERDICTS[:5], default=VERDICTS[5])
fair = {"Average of both methods": (fairA + fairB) / 2, "Method A only": fairA, "Method B only": fairB}[show_method]

res = pd.DataFrame({"Negeri": work["Negeri"], "Daerah": work["Daerah"], "Kategori": work["Kategori"], "Current": cur,
                    "Fair_A": fairA, "Fair_B": fairB, "Fair": fair, "Verdict": verdict})
res["Gap"] = res["Current"] - res["Fair"]
res["Label"] = res["Daerah"] + " (" + res["Negeri"] + ")"

tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "📋 Districts", "👔 PKP", "📝 What we still need"])

# ----------------------------------------------------------------------------- overview
with tab1:
    a, b, c, d, e = st.columns(5)
    a.metric("Districts", n)
    b.metric("PPKP (existing)", int(pool))
    c.metric("PKP (existing)", int(work["PKP"].sum()))
    d.metric("Short in both methods", int((verdict == VERDICTS[0]).sum()))
    e.metric("Surplus in both methods", int((verdict == VERDICTS[1]).sum()))
    st.caption("**How to read:** *Fair share* = the district's portion of the existing PPKP pool if officers followed recorded activity. "
               "*Gap* = current − fair share (negative = fewer officers than activity suggests). Total gap across the country is always zero, so a surplus here means 'relatively more', not 'too many'.")

    n_show = st.slider("Districts to show in each chart", 5, 30, 15)
    l, r = st.columns(2)
    with l:
        st.markdown("#### 🔴 Relatively short of officers")
        dd = res.sort_values("Gap").head(n_show)
        st.altair_chart(alt.Chart(dd).mark_bar(color="#d62728").encode(
            x=alt.X("Gap:Q", title="current − fair (posts)"), y=alt.Y("Label:N", sort="x", title=None),
            tooltip=[alt.Tooltip("Label:N", title="District"), alt.Tooltip("Current:Q", format=".0f"), alt.Tooltip("Fair:Q", title="Fair share", format=".1f"),
                     alt.Tooltip("Gap:Q", format="+.1f"), alt.Tooltip("Verdict:N")]), height=max(220, 24 * n_show))
    with r:
        st.markdown("#### 🟢 Relatively more officers")
        ss = res.sort_values("Gap", ascending=False).head(n_show)
        st.altair_chart(alt.Chart(ss).mark_bar(color="#2ca02c").encode(
            x=alt.X("Gap:Q", title="current − fair (posts)"), y=alt.Y("Label:N", sort="-x", title=None),
            tooltip=[alt.Tooltip("Label:N", title="District"), alt.Tooltip("Current:Q", format=".0f"), alt.Tooltip("Fair:Q", title="Fair share", format=".1f"),
                     alt.Tooltip("Gap:Q", format="+.1f"), alt.Tooltip("Verdict:N")]), height=max(220, 24 * n_show))

    st.markdown("#### By state")
    g = res.groupby("Negeri").agg(Current=("Current", "sum"), Fair=("Fair", "sum")).reset_index()
    g["Gap"] = g["Current"] - g["Fair"]
    g["Gap %"] = 100 * g["Gap"] / g["Fair"]
    sc1, sc2 = st.columns([3, 2])
    with sc1:
        st.altair_chart(alt.Chart(g).mark_bar().encode(
            x=alt.X("Negeri:N", sort="y", title=None), y=alt.Y("Gap:Q", title="current − fair (posts)"),
            color=alt.condition(alt.datum.Gap < 0, alt.value("#d62728"), alt.value("#2ca02c")),
            tooltip=[alt.Tooltip("Negeri:N", title="State"), alt.Tooltip("Current:Q", format=".0f"), alt.Tooltip("Fair:Q", title="Fair share", format=".0f"),
                     alt.Tooltip("Gap:Q", format="+.0f")]), height=340)
    with sc2:
        st.dataframe(g.sort_values("Gap"), hide_index=True, column_config={
            "Current": st.column_config.NumberColumn(format="%d"), "Fair": st.column_config.NumberColumn("Fair share", format="%.0f"),
            "Gap": st.column_config.NumberColumn(format="%+.0f"), "Gap %": st.column_config.NumberColumn(format="%+.0f%%")})

# ----------------------------------------------------------------------------- districts
with tab2:
    f1, f2 = st.columns(2)
    states = f1.multiselect("State", sorted(res["Negeri"].unique()))
    verd = f2.multiselect("Verdict", VERDICTS)
    view = res.copy()
    if states:
        view = view[view["Negeri"].isin(states)]
    if verd:
        view = view[view["Verdict"].isin(verd)]
    show = view[["Negeri", "Daerah", "Kategori", "Current", "Fair_A", "Fair_B", "Gap", "Verdict"]].sort_values("Gap")
    st.dataframe(show, hide_index=True, column_config={
        "Current": st.column_config.NumberColumn("Current PPKP", format="%d"),
        "Fair_A": st.column_config.NumberColumn("Fair A (share of activity)", format="%.1f"),
        "Fair_B": st.column_config.NumberColumn("Fair B (minutes)", format="%.1f"),
        "Gap": st.column_config.NumberColumn("Gap (current − fair)", format="%+.1f")})
    st.caption("**Verdict:** 🔴/🟢 = both methods agree the district is beyond the tolerance. 🟠/🟡 = only one method says so. ⚪ = balanced or methods disagree - treat as uncertain.")
    st.download_button("📥 Download table (CSV)", csv_bytes(show), "district_results.csv", "text/csv")

    st.markdown("---")
    st.markdown("#### Drill-down: why is this district flagged?")
    pick = st.selectbox("District", sorted(res["Label"]), index=int(np.argmin(res["Gap"].to_numpy())) if len(res) else 0)
    i = int(res.index[res["Label"] == pick][0])
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Current PPKP", int(cur[i]))
    k2.metric("Fair A", f"{fairA[i]:.1f}")
    k3.metric("Fair B", f"{fairB[i]:.1f}")
    k4.metric("Verdict", verdict[i].split(" (")[0])
    nat = C.sum(axis=0) / pool
    with np.errstate(divide="ignore", invalid="ignore"):
        idx_task = np.where((cur[i] > 0) & (nat > 0), 100 * (C[i] / cur[i]) / nat, np.nan)
    dr = pd.DataFrame({"Task": [SHORT[k] for k in KEYS], "Index": idx_task, "Count": C[i]})
    dr = dr[dr["Count"] > 0]
    st.caption("Units per PPKP in this district vs the national average (100). Above 100 = heavier per officer than the national pattern. Tasks the district does not record are hidden.")
    if len(dr):
        bars = alt.Chart(dr).mark_bar().encode(
            x=alt.X("Index:Q", title="index (national average = 100)"), y=alt.Y("Task:N", sort="-x", title=None),
            color=alt.condition(alt.datum.Index > 100, alt.value("#d62728"), alt.value("#4575b4")),
            tooltip=[alt.Tooltip("Task:N"), alt.Tooltip("Count:Q", format=",.0f"), alt.Tooltip("Index:Q", format=".0f")])
        rule = alt.Chart(pd.DataFrame({"x": [100]})).mark_rule(strokeDash=[4, 3]).encode(x="x:Q")
        st.altair_chart(bars + rule, height=max(240, 26 * len(dr)))

# ----------------------------------------------------------------------------- PKP
with tab3:
    st.markdown("PKP work is mostly supervision and governance and **cannot be split from PPKP work in the data**, so it is checked against a simple rule: "
                "**required PKP = the larger of (minimum per district) and (PPKP ÷ span of control, rounded up).**")
    p1, p2 = st.columns(2)
    span = p1.slider("Span of control (PPKP per PKP)", 10, 40, 25)
    min_pkp = p2.number_input("Minimum PKP per district", 0, 5, 1)
    req = np.maximum(min_pkp, np.ceil(cur / span))
    pk_cur = work["PKP"].to_numpy(float)
    pk = pd.DataFrame({"Negeri": work["Negeri"], "Daerah": work["Daerah"], "Kategori": work["Kategori"], "PPKP": cur, "PKP_current": pk_cur,
                       "PKP_required": req, "Gap": pk_cur - req, "Label": res["Label"]})
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("PKP now", int(pk_cur.sum()))
    m2.metric("PKP by rule", int(req.sum()))
    m3.metric("Shortfall", int(max(0, req.sum() - pk_cur.sum())))
    m4.metric("Districts with 0 PKP", int((pk_cur == 0).sum()))
    st.altair_chart(alt.Chart(pk.sort_values("Gap").head(20)).mark_bar(color="#d62728").encode(
        x=alt.X("Gap:Q", title="current − required PKP"), y=alt.Y("Label:N", sort="x", title=None),
        tooltip=[alt.Tooltip("Label:N", title="District"), alt.Tooltip("PPKP:Q", format=".0f"), alt.Tooltip("PKP_current:Q", title="PKP now", format=".0f"),
                 alt.Tooltip("PKP_required:Q", title="PKP required", format=".0f")]), height=480)
    st.dataframe(pk.drop(columns="Label").sort_values("Gap"), hide_index=True)
    st.caption("Some districts with 0 PKP may in practice be supervised from a neighbouring district - please confirm. The span value is a policy choice; try 20, 25 and 30 to see how much it matters.")
    st.download_button("📥 Download PKP table (CSV)", csv_bytes(pk.drop(columns="Label")), "pkp_rule.csv", "text/csv")

# ----------------------------------------------------------------------------- requests
with tab4:
    st.subheader("What we still need to make a defensible norm")
    st.markdown("""
**A. Quick confirmations (a few lines each)**
1. What exactly does one *count* mean for **Kawalan penyakit berjangkit** and **kapal / vessel / penumpang** (visits, premises, cases, passengers)? What does **"(5%)"** in the *Penilaian risiko* column mean?
2. Which year is this data for? Can we also get the previous two years in the same template?
3. Who covers the work of the 5 districts with no PTJ (Sebauh, Bukit Mabong, Telang Usan, Subis, Beluru)?
4. What does "fair" mean here: equal workload per officer, or equal service coverage?

**B. Information to provide (priority 1)**
5. **Applicability table:** for each district, tick which tasks apply (e.g. vessel/PMA work only where there is a port or airport).
6. **Time per task:** best estimate of minutes for one unit of each task (and a low/high), from a few experienced officers per task.
7. **Who does what:** for each task, PKP, PPKP or both (and roughly what share), plus PKP duties not in the 20 tasks.
8. **Available time:** leave/course days per year, and the share of working time spent on these 20 tasks (the rest is admin, meetings, other programmes).

**C. Information to provide (priority 2)**
9. **Demand indicators per district:** population, number of food premises / factories / water sources, area, port or airport, island or interior flag, travel time to the farthest site.
10. **Posts vs staff:** approved posts (waran), filled, vacant, contract; any limits on transfers.
11. **Minimum team:** the smallest PPKP/PKP team a district needs to run outbreak response and on-call.
12. **Targets:** SOP/KPI frequencies (for example inspections per premises per year, response times).
""")
