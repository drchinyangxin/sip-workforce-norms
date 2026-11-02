"""
PKP & PPKP Distribution Review  (v2)
------------------------------------
Stage 1  Descriptive picture            - no assumptions, raw task counts per PPKP
Stage 2  Workload model (PPKP)          - editable minutes/unit table + Monte-Carlo ranges
Stage 3  Robust priority list & moves   - state-first, capped, within-state transfer plan
Stage 4  PKP structural rule            - span-of-control rule (workload cannot be split PKP/PPKP)

Run:  streamlit run app.py
Place the Excel file ("Data Perjawatan dan Beban Tugas PKP PPKP di PKD ...xlsx") next to app.py,
or upload it from the sidebar.
"""
import glob
import io
import math
import os

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(layout="wide", page_title="PKP & PPKP Distribution Review", page_icon="⚖️")

# =============================================================================
# CONFIG: the 20 task columns.  base = PLACEHOLDER minutes per recorded unit.
# Replace with time-study / expert-panel values (edit in the app or upload a CSV).
# =============================================================================
TASKS = [
    # key, Malay label, short label, header keyword, group, base min/unit, % by PKP, unit note
    ("kawalan", "Kawalan penyakit berjangkit", "Kawalan PB", "kawalan penyakit berjangkit", "Disease control", 10, 0, "count as recorded"),
    ("sampel", "Persampelan penyakit berjangkit", "Persampelan", "persampelan penyakit", "Disease control", 45, 0, "samples"),
    ("siasat_kes", "Siasatan kes & kontak", "Siasatan kes", "siasatan kes dan kontak", "Disease control", 30, 0, "cases/contacts"),
    ("bilik", "Bilik gerakan wabak/bencana dibuka", "Bilik gerakan", "bilik gerakan", "Disease control", 600, 20, "activations"),
    ("survelans", "Laporan aktiviti survelans", "Survelans", "laporan aktiviti survelans", "Disease control", 30, 0, "reports"),
    ("risikan", "Risikan undang-undang kesihatan awam", "Risikan UU", "risikan undang", "Legal", 120, 0, "intelligence ops"),
    ("notis", "Notis kompaun / kompaun", "Kompaun", "notis kompaun", "Legal", 30, 0, "notices"),
    ("kertas", "Kertas siasatan (kesalahan UU KA)", "Kertas siasatan", "kertas siasatan", "Legal", 240, 15, "investigation papers"),
    ("pewasapan", "Audit pewasapan", "Audit wasap", "audit pewasapan", "Legal", 180, 0, "audits"),
    ("premis", "Pemeriksaan kesihatan persekitaran di premis", "Periksa premis", "pemeriksaan kesihatan persekitaran", "Environmental / OSH", 30, 0, "premises"),
    ("makanan_premis", "Penilaian risiko premis makanan", "Risiko premis mkn", "penilaian risiko premis makanan", "Environmental / OSH", 30, 0, "premises"),
    ("osh_audit", "Audit KKP di tempat kerja", "Audit KKP", "audit keselamatan", "Environmental / OSH", 240, 0, "audits"),
    ("osh_siasat", "Notifikasi & siasatan kemalangan/penyakit pekerjaan", "Siasat KKP", "notifikasi dan siasatan kemalangan", "Environmental / OSH", 180, 0, "investigations"),
    ("air", "Pemantauan kualiti air minum", "Kualiti air", "pemantauan kualiti air", "Environmental / OSH", 45, 0, "samples/visits"),
    ("kapal", "Pemeriksaan & penilaian risiko kapal/vessel/penumpang", "Kapal/PMA", "kapal", "Points of entry (PMA)", 0.1, 0, "CHECK UNIT: values look like passenger counts"),
    ("mayat", "Import/eksport mayat, tisu, bahan patogenik", "Mayat/patogen", "mayat", "Points of entry (PMA)", 60, 0, "consignments"),
    ("makanan_import", "Pemeriksaan & persampelan produk makanan import", "Makanan import", "produk makanan import", "Points of entry (PMA)", 60, 0, "consignments"),
    ("latihan", "Latihan teknikal / amalan profesional PKP/PPKP", "Latihan", "latihan teknikal", "Governance & training", 240, 10, "sessions"),
    ("advokasi", "Advokasi & pendidikan kesihatan", "Advokasi", "advokasi", "Governance & training", 90, 0, "sessions"),
    ("penilaian", "Penilaian risiko kesihatan persekitaran, PB & bencana", "Penilaian risiko", "penilaian risiko kesihatan persekitaran", "Governance & training", 20, 10, "assessments"),
]
KEYS = [t[0] for t in TASKS]
LABEL = {t[0]: t[1] for t in TASKS}
SHORT = {t[0]: t[2] for t in TASKS}
ALLOC_CATEGORIES = ["PKD", "PKB", "PKK"]
DEF, SUR, BAL = "🔴 Deficit", "🟢 Surplus", "⚪ Balanced / uncertain"


# =============================================================================
# 1. DATA LOADING (strict column matching: raises instead of silently returning 0)
# =============================================================================
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
            raise ValueError(f"Column for '{what}' matched {len(hits)} columns (expected 1).")
        return hits[0]

    idx = {
        "Negeri": find(lambda a, b: a == "negeri", "Negeri"),
        "Daerah": find(lambda a, b: a == "daerah", "Daerah"),
        "Status_PTJ": find(lambda a, b: a == "status ptj", "Status PTJ"),
        "Kategori": find(lambda a, b: a == "kategori", "Kategori"),
        "PKP_tetap": find(lambda a, b: a.startswith("bil pkp") and "kontrak" not in a, "PKP tetap"),
        "PKP_kontrak": find(lambda a, b: a.startswith("bil pkp") and "kontrak" in a, "PKP kontrak"),
        "PPKP_tetap": find(lambda a, b: a.startswith("bil ppkp") and "kontrak" not in a, "PPKP tetap"),
        "PPKP_kontrak": find(lambda a, b: a.startswith("bil ppkp") and "kontrak" in a, "PPKP kontrak"),
    }
    for key, _, _, kw, *_ in TASKS:
        idx[key] = find(lambda a, b, kw=kw: a.startswith("beban") and kw in b, kw)

    df = pd.DataFrame({name: raw.iloc[:, i] for name, i in idx.items()})
    for c in ["Negeri", "Daerah", "Status_PTJ", "Kategori"]:
        df[c] = df[c].astype("string").str.strip()
    df = df[df["Daerah"].notna() & (df["Daerah"].str.lower() != "total") & (df["Kategori"] != "Total")].copy()
    num = ["PKP_tetap", "PKP_kontrak", "PPKP_tetap", "PPKP_kontrak"] + KEYS
    for c in num:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["Negeri"] = df["Negeri"].astype(str)
    df["Daerah"] = df["Daerah"].astype(str)
    df["Kategori"] = df["Kategori"].astype(str)
    df["Status_PTJ"] = df["Status_PTJ"].fillna("").astype(str)
    return df.reset_index(drop=True)


def find_default_file():
    here = os.path.dirname(os.path.abspath(__file__))
    for folder in (here, os.getcwd()):
        hits = sorted(glob.glob(os.path.join(folder, "Data*Perjawatan*.xlsx")))
        if hits:
            return hits[-1]
    return None


# =============================================================================
# 2. CORE MODEL FUNCTIONS
# =============================================================================
def simulate(C, lo, base, hi, pkp_share, floor, pool, n_draws, seed=42):
    """Monte-Carlo fair-share of a fixed PPKP pool.
    Each task's minutes/unit is drawn independently from Triangular(low, base, high).
    NB: scaling all minutes by the same factor cannot change shares, so ranges
    only matter when tasks are uncertain by *different* amounts - hence independent draws."""
    n, T = C.shape
    rng = np.random.default_rng(seed)
    S = np.empty((n_draws, T))
    for j in range(T):
        if hi[j] - lo[j] < 1e-12:
            S[:, j] = base[j]
        else:
            S[:, j] = rng.triangular(lo[j], base[j], hi[j], n_draws)
    eff = S * (1 - pkp_share)
    dem = C @ eff.T                                  # n x N  (minutes)
    floor_eff = min(floor, pool / n) if n else 0
    distributable = pool - floor_eff * n
    tot = dem.sum(axis=0)
    tot[tot == 0] = np.nan
    fair = floor_eff + np.nan_to_num(dem / tot) * distributable
    base_dem = C @ (base * (1 - pkp_share))
    base_fair = floor_eff + (base_dem / base_dem.sum()) * distributable if base_dem.sum() > 0 else np.full(n, floor_eff)
    return fair, base_fair, base_dem


def plan_moves(res, cap_pct, floor):
    """Greedy within-state transfers: robust-surplus districts -> robust-deficit districts."""
    moves, resid = [], []
    for state, g in res.groupby("Negeri"):
        give, need = {}, {}
        for r in g.to_dict("records"):
            if r["Status"] == SUR:
                give[r["Daerah"]] = int(max(0, min(math.floor(r["Gap"]), math.floor(cap_pct / 100 * r["Current"]), r["Current"] - floor)))
            elif r["Status"] == DEF:
                need[r["Daerah"]] = int(max(0, math.floor(-r["Gap"])))
        give = {k: v for k, v in give.items() if v > 0}
        need = {k: v for k, v in need.items() if v > 0}
        while give and need:
            d = max(give, key=give.get)
            r_ = max(need, key=need.get)
            amt = min(give[d], need[r_])
            moves.append({"Negeri": state, "From": d, "To": r_, "Posts": amt})
            give[d] -= amt
            need[r_] -= amt
            give = {k: v for k, v in give.items() if v > 0}
            need = {k: v for k, v in need.items() if v > 0}
        resid.append({"Negeri": state, "Unmet need (posts)": sum(need.values()), "Unused surplus (posts)": sum(give.values())})
    return pd.DataFrame(moves, columns=["Negeri", "From", "To", "Posts"]), pd.DataFrame(resid)


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


# =============================================================================
# 3. SIDEBAR + LOAD
# =============================================================================
st.title("⚖️ PKP & PPKP Distribution Review")
st.caption("Staged, assumption-transparent review of manpower distribution across districts | KKM")

st.sidebar.header("📁 Data")
upload = st.sidebar.file_uploader("Excel file (optional if next to app.py)", type=["xlsx"])
path = find_default_file()
if upload is None and path is None:
    st.info("Place the Excel file next to `app.py` or upload it in the sidebar.")
    st.stop()
try:
    full = load_data(upload.getvalue() if upload else None, None if upload else path)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

st.sidebar.header("⚙️ Scope")
use_contract = st.sidebar.checkbox("Count contract (2025) staff", value=True,
                                   help="Contract staff are not permanent posts. Untick to see the picture with permanent staff only.")
include_tiada = st.sidebar.checkbox("Include districts with no PTJ (Kategori 'Tiada')", value=False,
                                    help="These have recorded workload but no staff. Default: excluded; their work is presumably absorbed by neighbouring districts - confirm with the stakeholder.")

full["PKP"] = full["PKP_tetap"] + (full["PKP_kontrak"] if use_contract else 0)
full["PPKP"] = full["PPKP_tetap"] + (full["PPKP_kontrak"] if use_contract else 0)
elig = full["Kategori"].isin(ALLOC_CATEGORIES) | (include_tiada & (full["Kategori"] == "Tiada"))
work = full[elig].reset_index(drop=True)
excluded = full[~elig]
n = len(work)
C_raw = work[KEYS].to_numpy(float)
cur = work["PPKP"].to_numpy(float)
pool = float(cur.sum())

c1, c2, c3, c4 = st.columns(4)
c1.metric("Districts in scope", n)
c2.metric("PPKP pool", int(pool))
c3.metric("PKP pool", int(work["PKP"].sum()))
c4.metric("Districts with 0 PKP", int((work["PKP"] == 0).sum()))
if len(excluded):
    st.caption(f"Excluded from allocation: {len(excluded)} districts ({', '.join(excluded['Daerah'].head(6))}{'…' if len(excluded) > 6 else ''}).")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1 · Descriptive picture", "2 · Workload model (PPKP)", "3 · Priority list & moves", "4 · PKP structural rule", "ℹ️ Notes & data gaps"])

# =============================================================================
# TAB 1 - DESCRIPTIVE (no weights, no minutes)
# =============================================================================
with tab1:
    st.subheader("What the data already shows - no assumptions needed")
    st.caption("Task counts per PPKP, compared across districts. Counts are as recorded; units differ between tasks, so tasks are never added together here.")

    ppkp_safe = np.where(cur > 0, cur, np.nan)
    R = C_raw / ppkp_safe[:, None]                                   # task units per PPKP
    mask = (C_raw > 0) & (cur[:, None] > 0)
    applicable_task = mask.sum(axis=0) >= 8                          # only tasks done in >= 8 districts
    pct = np.full(R.shape, np.nan)
    for j in range(len(KEYS)):
        if applicable_task[j]:
            s = pd.Series(R[mask[:, j], j])
            pct[mask[:, j], j] = s.rank(pct=True).to_numpy()
    high = np.nansum(pct >= 0.75, axis=1)
    appl = np.nansum(~np.isnan(pct), axis=1)

    left, right = st.columns([1, 1])
    with left:
        task_sel = st.selectbox("Task", KEYS, format_func=lambda k: LABEL[k], index=0)
        j = KEYS.index(task_sel)
        d = pd.DataFrame({"District": work["Daerah"] + " (" + work["Negeri"] + ")", "PerPPKP": R[:, j], "Count": C_raw[:, j], "PPKP": cur})
        d = d[d["Count"] > 0].sort_values("PerPPKP", ascending=False)
        nat = C_raw[:, j].sum() / pool if pool else 0
        st.caption(f"National average: {nat:,.1f} per PPKP · done in {len(d)} of {n} districts")
        top = d.head(25)
        bars = alt.Chart(top).mark_bar().encode(
            x=alt.X("PerPPKP:Q", title="units per PPKP"),
            y=alt.Y("District:N", sort="-x", title=None),
            tooltip=["District", alt.Tooltip("Count:Q", format=",.0f"), "PPKP", alt.Tooltip("PerPPKP:Q", format=",.1f")])
        rule = alt.Chart(pd.DataFrame({"x": [nat]})).mark_rule(color="red", strokeDash=[4, 3]).encode(x="x:Q")
        st.altair_chart(bars + rule, height=520)
    with right:
        st.markdown("**State heatmap** - units per PPKP, indexed to national average = 100")
        g = work.groupby("Negeri")[KEYS + ["PPKP"]].sum()
        g = g[g["PPKP"] > 0]
        nat_ratio = C_raw.sum(axis=0) / pool
        rows = []
        for k_i, k in enumerate(KEYS):
            if not applicable_task[k_i] or nat_ratio[k_i] == 0:
                continue
            for state, r_ in g.iterrows():
                rows.append({"State": state, "Task": SHORT[k], "Index": 100 * (r_[k] / r_["PPKP"]) / nat_ratio[k_i]})
        hm = pd.DataFrame(rows)
        base_c = alt.Chart(hm).encode(x=alt.X("Task:N", sort=[SHORT[k] for k in KEYS], title=None, axis=alt.Axis(labelAngle=-60)),
                                      y=alt.Y("State:N", title=None))
        heat = base_c.mark_rect().encode(color=alt.Color("Index:Q", scale=alt.Scale(domain=[0, 100, 250], range=["#4575b4", "#ffffbf", "#d73027"], clamp=True),
                                                         legend=alt.Legend(title="Index")),
                                         tooltip=["State", "Task", alt.Tooltip("Index:Q", format=".0f")])
        txt = base_c.mark_text(fontSize=9).encode(text=alt.Text("Index:Q", format=".0f"))
        st.altair_chart(heat + txt, height=520)

    st.markdown("**District burden profile** - in how many applicable tasks is the district in the *top quartile* of units per PPKP?")
    prof = pd.DataFrame({
        "Negeri": work["Negeri"], "Daerah": work["Daerah"], "Kategori": work["Kategori"],
        "PPKP": cur, "PKP": work["PKP"],
        "Applicable tasks": appl.astype(int), "Top-quartile tasks": high.astype(int)})
    prof["Share top-quartile (%)"] = np.where(prof["Applicable tasks"] > 0, 100 * prof["Top-quartile tasks"] / prof["Applicable tasks"].replace(0, np.nan), 0).round(0)
    prof = prof.sort_values("Share top-quartile (%)", ascending=False)
    st.dataframe(prof, hide_index=True,
                 column_config={"Share top-quartile (%)": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%")})
    st.caption("A district that is top-quartile on most tasks while having few PPKP is a candidate for review under any weighting. Tasks done in fewer than 8 districts are not ranked.")
    st.download_button("📥 Download profile (CSV)", csv_bytes(prof), "stage1_burden_profile.csv", "text/csv")

# =============================================================================
# TAB 2 - WORKLOAD MODEL
# =============================================================================
with tab2:
    st.subheader("PPKP fair-share model: minutes per unit × recorded counts")
    st.warning("**Indicative only.** The minutes below are *placeholders* until a time study / expert panel (e.g. Delphi) provides real values. "
               "Results are shown as ranges across the plausible minutes, not a single 'norm'.")

    with st.expander("⚙️ Parameters", expanded=False):
        p1, p2, p3 = st.columns(3)
        with p1:
            floor = st.number_input("Minimum PPKP per district (floor)", 0, 20, 3)
            cap_pct_win = st.slider("Cap outlier counts at percentile (100 = off)", 90, 100, 100,
                                    help="Caps each task's counts at this percentile of non-zero districts, to limit single-district outliers.")
        with p2:
            abs_tol = st.number_input("Tolerance: absolute (posts)", 0.0, 20.0, 2.0, 0.5)
            pct_tol = st.slider("Tolerance: % of fair share", 0, 50, 15) / 100
            conf = st.slider("Confidence to call deficit/surplus", 50, 100, 80, help="Share of simulations that must agree.") / 100
        with p3:
            n_draws = st.select_slider("Simulation draws", [200, 500, 1000, 2000], value=1000)
            days = st.number_input("Effective working days / officer / year", 100, 300, 200)
            hours = st.number_input("Hours / day", 4.0, 10.0, 7.0, 0.5)
            share_listed = st.slider("% of time on these 20 tasks", 20, 100, 60) / 100

    # ---- assumptions table -------------------------------------------------
    def default_assumptions():
        return pd.DataFrame([{"key": t[0], "Task": t[1], "Group": t[4], "Unit": t[7],
                              "Low": round(t[5] * 0.6, 3), "Base": t[5], "High": round(t[5] * 1.6, 3), "PKP_pct": t[6]} for t in TASKS])

    a_up = st.file_uploader("Load assumptions from CSV (columns: key, Low, Base, High, PKP_pct)", type=["csv"], key="assump_upload")
    adf = default_assumptions()
    ver = "default"
    if a_up is not None:
        try:
            up = pd.read_csv(a_up)
            adf = adf.drop(columns=["Low", "Base", "High", "PKP_pct"]).merge(up[["key", "Low", "Base", "High", "PKP_pct"]], on="key", how="left")
            adf[["Low", "Base", "High", "PKP_pct"]] = adf[["Low", "Base", "High", "PKP_pct"]].fillna(0)
            ver = a_up.name
        except Exception as e:
            st.error(f"Could not read assumptions CSV: {e}")

    st.markdown("**Assumptions (editable)** - minutes per recorded unit (low / base / high) and % of the task done by PKP rather than PPKP.")
    edited = st.data_editor(
        adf, key=f"assump_{ver}", hide_index=True, disabled=["key", "Task", "Group", "Unit"],
        column_config={
            "key": None,
            "Low": st.column_config.NumberColumn("Min/unit low", min_value=0.0, format="%.2f"),
            "Base": st.column_config.NumberColumn("Min/unit base", min_value=0.0, format="%.2f"),
            "High": st.column_config.NumberColumn("Min/unit high", min_value=0.0, format="%.2f"),
            "PKP_pct": st.column_config.NumberColumn("% done by PKP", min_value=0, max_value=100, format="%d"),
        })
    st.download_button("📥 Save assumptions (CSV)", csv_bytes(edited), "assumptions.csv", "text/csv")

    base_m = edited["Base"].to_numpy(float)
    lo_m = np.minimum(edited["Low"].to_numpy(float), base_m)
    hi_m = np.maximum(edited["High"].to_numpy(float), base_m)
    pkp_share = edited["PKP_pct"].to_numpy(float) / 100

    # ---- outlier cap -------------------------------------------------------
    C = C_raw.copy()
    if cap_pct_win < 100:
        for j in range(C.shape[1]):
            nz = C[C[:, j] > 0, j]
            if len(nz):
                C[:, j] = np.minimum(C[:, j], np.percentile(nz, cap_pct_win))

    # ---- simulate ----------------------------------------------------------
    fair, base_fair, base_dem = simulate(C, lo_m, base_m, hi_m, pkp_share, floor, pool, n_draws)
    tol = np.maximum(abs_tol, pct_tol * fair)
    p_def = (cur[:, None] < fair - tol).mean(axis=1)
    p_sur = (cur[:, None] > fair + tol).mean(axis=1)
    status = np.select([p_def >= conf, p_sur >= conf], [DEF, SUR], default=BAL)
    capacity = days * hours * 60 * share_listed
    res = pd.DataFrame({
        "Negeri": work["Negeri"], "Daerah": work["Daerah"], "Kategori": work["Kategori"],
        "Current": cur, "Fair_base": base_fair, "Fair_p10": np.percentile(fair, 10, axis=1),
        "Fair_med": np.median(fair, axis=1), "Fair_p90": np.percentile(fair, 90, axis=1),
        "P_deficit": 100 * p_def, "P_surplus": 100 * p_sur, "Status": status,
        "FTE_indicative": base_dem / capacity})
    res["Gap"] = res["Current"] - res["Fair_med"]
    res["Gap_lo"] = res["Current"] - res["Fair_p90"]     # most pessimistic (largest deficit)
    res["Gap_hi"] = res["Current"] - res["Fair_p10"]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Robust deficit districts", int((status == DEF).sum()))
    m2.metric("Robust surplus districts", int((status == SUR).sum()))
    m3.metric("Balanced / uncertain", int((status == BAL).sum()))
    m4.metric("Indicative FTE demand vs PPKP pool", f"{res['FTE_indicative'].sum():,.0f} vs {int(pool)}")
    st.caption("The fair share only redistributes the **existing** pool by relative workload, so it shows *maldistribution*, not an absolute shortage. "
               "The FTE figure depends entirely on the placeholder minutes and must not be quoted as a requirement.")

    show = res[["Negeri", "Daerah", "Kategori", "Current", "Fair_med", "Fair_p10", "Fair_p90", "Gap", "P_deficit", "P_surplus", "Status", "FTE_indicative"]].sort_values("Gap")
    st.dataframe(show, hide_index=True, column_config={
        "Current": st.column_config.NumberColumn("Current PPKP", format="%d"),
        "Fair_med": st.column_config.NumberColumn("Fair share (median)", format="%.1f"),
        "Fair_p10": st.column_config.NumberColumn("P10", format="%.1f"),
        "Fair_p90": st.column_config.NumberColumn("P90", format="%.1f"),
        "Gap": st.column_config.NumberColumn("Gap (current − fair)", format="%+.1f"),
        "P_deficit": st.column_config.ProgressColumn("% sims in deficit", min_value=0, max_value=100, format="%.0f%%"),
        "P_surplus": st.column_config.ProgressColumn("% sims in surplus", min_value=0, max_value=100, format="%.0f%%"),
        "FTE_indicative": st.column_config.NumberColumn("FTE (indicative)", format="%.1f")})
    st.download_button("📥 Download results (CSV)", csv_bytes(show), "stage2_ppkp_results.csv", "text/csv")

    st.markdown("**Largest deficits (median gap, bar = P10–P90 range)**")
    topd = res.sort_values("Gap").head(20).copy()
    topd["Label"] = topd["Daerah"] + " (" + topd["Negeri"] + ")"
    ch = alt.Chart(topd).encode(y=alt.Y("Label:N", sort=alt.EncodingSortField("Gap", order="ascending"), title=None))
    st.altair_chart(ch.mark_bar(color="#d62728", opacity=0.55).encode(x=alt.X("Gap:Q", title="current − fair share (posts)"), tooltip=["Label", alt.Tooltip("Gap:Q", format="+.1f")])
                    + ch.mark_rule(color="black").encode(x="Gap_lo:Q", x2="Gap_hi:Q"), height=480)

# =============================================================================
# TAB 3 - PRIORITY LIST, STATE-FIRST, MOVES
# =============================================================================
with tab3:
    st.subheader("Robust priority list - districts flagged in most simulations")
    st.caption(f"A district is listed only if at least {int(conf * 100)}% of simulations place it beyond the tolerance. These are the findings least sensitive to the placeholder minutes.")

    pl, pr = st.columns(2)
    d_cols = ["Negeri", "Daerah", "Current", "Fair_med", "Gap", "P_deficit"]
    with pl:
        st.markdown("**🔴 Robust deficit**")
        df_def = res[res["Status"] == DEF].sort_values("Gap")[d_cols]
        st.dataframe(df_def, hide_index=True, column_config={"Fair_med": st.column_config.NumberColumn("Fair", format="%.1f"), "Gap": st.column_config.NumberColumn(format="%+.1f"),
                                                             "P_deficit": st.column_config.NumberColumn("% sims", format="%.0f")})
    with pr:
        st.markdown("**🟢 Robust surplus**")
        df_sur = res[res["Status"] == SUR].sort_values("Gap", ascending=False)[["Negeri", "Daerah", "Current", "Fair_med", "Gap", "P_surplus"]]
        st.dataframe(df_sur, hide_index=True, column_config={"Fair_med": st.column_config.NumberColumn("Fair", format="%.1f"), "Gap": st.column_config.NumberColumn(format="%+.1f"),
                                                             "P_surplus": st.column_config.NumberColumn("% sims", format="%.0f")})

    st.markdown("---")
    st.subheader("State level first")
    states = sorted(work["Negeri"].unique())
    onehot = np.array([(work["Negeri"] == s).to_numpy(float) for s in states])
    fs = onehot @ fair
    state_tbl = pd.DataFrame({"Negeri": states, "Current": onehot @ cur, "Fair_med": np.median(fs, axis=1),
                              "Fair_p10": np.percentile(fs, 10, axis=1), "Fair_p90": np.percentile(fs, 90, axis=1)})
    state_tbl["Gap"] = state_tbl["Current"] - state_tbl["Fair_med"]
    state_tbl["Gap %"] = 100 * state_tbl["Gap"] / state_tbl["Fair_med"]
    sc1, sc2 = st.columns([1, 1])
    with sc1:
        st.dataframe(state_tbl.sort_values("Gap"), hide_index=True, column_config={
            "Current": st.column_config.NumberColumn(format="%d"), "Fair_med": st.column_config.NumberColumn("Fair (median)", format="%.1f"),
            "Fair_p10": st.column_config.NumberColumn("P10", format="%.1f"), "Fair_p90": st.column_config.NumberColumn("P90", format="%.1f"),
            "Gap": st.column_config.NumberColumn(format="%+.1f"), "Gap %": st.column_config.NumberColumn(format="%+.0f%%")})
    with sc2:
        st.altair_chart(alt.Chart(state_tbl).mark_bar().encode(
            x=alt.X("Negeri:N", sort="y", title=None), y=alt.Y("Gap:Q", title="current − fair (posts)"),
            color=alt.condition(alt.datum.Gap < 0, alt.value("#d62728"), alt.value("#2ca02c")),
            tooltip=["Negeri", alt.Tooltip("Gap:Q", format="+.1f")]), height=320)

    st.markdown("---")
    st.subheader("Suggested within-state transfers (capped)")
    mv1, mv2 = st.columns(2)
    cap_move = mv1.slider("Max share of a donor district's current PPKP that can be moved (%)", 0, 50, 20)
    mv2.caption("Only robust-surplus districts give and robust-deficit districts receive, within the same state. "
                "Whatever cannot be matched is shown as residual: it needs cross-state rebalancing or new posts.")
    moves, resid = plan_moves(res, cap_move, floor)
    if moves.empty:
        st.info("No within-state transfers suggested under the current settings.")
    else:
        st.dataframe(moves, hide_index=True)
        st.download_button("📥 Download suggested moves (CSV)", csv_bytes(moves), "suggested_moves.csv", "text/csv")
    st.markdown("**Residual after within-state transfers**")
    st.dataframe(resid[(resid.iloc[:, 1] > 0) | (resid.iloc[:, 2] > 0)], hide_index=True)

# =============================================================================
# TAB 4 - PKP STRUCTURAL RULE
# =============================================================================
with tab4:
    st.subheader("PKP: span-of-control rule (not workload-based)")
    st.caption("The data cannot separate which recorded tasks are done by PKP versus PPKP, so PKP is assessed against a transparent structural rule: "
               "required PKP = max(minimum per district, ⌈PPKP ÷ span⌉).")
    k1, k2, k3 = st.columns(3)
    span = k1.slider("Span of control (PPKP per PKP)", 10, 40, 25)
    min_pkp = k2.number_input("Minimum PKP per district", 0, 5, 1)
    basis = k3.radio("PPKP count used", ["Current PPKP", "Fair-share PPKP (Tab 2)"], horizontal=True)

    ppkp_basis = cur if basis == "Current PPKP" else res["Fair_med"].to_numpy()
    req = np.maximum(min_pkp, np.ceil(ppkp_basis / span))
    pkp_cur = work["PKP"].to_numpy(float)
    pkp_pool = pkp_cur.sum()
    pr_share = pkp_pool * req / req.sum() if req.sum() else np.zeros(n)
    pk = pd.DataFrame({"Negeri": work["Negeri"], "Daerah": work["Daerah"], "Kategori": work["Kategori"], "PPKP": ppkp_basis.round(1),
                       "PKP_current": pkp_cur, "PKP_required": req, "Gap_vs_rule": pkp_cur - req,
                       "Prorata_share": pr_share.round(2), "Gap_vs_prorata": (pkp_cur - pr_share).round(2)})

    q1, q2, q3, q4 = st.columns(4)
    q1.metric("PKP current", int(pkp_pool))
    q2.metric("PKP required by rule", int(req.sum()))
    q3.metric("National shortfall", int(max(0, req.sum() - pkp_pool)))
    q4.metric("Districts below rule", int((pkp_cur < req).sum()))
    st.caption("Two views: (a) **gap vs rule** shows the absolute shortfall; (b) **pro-rata** spreads the scarce PKP pool in proportion to each district's requirement - useful for prioritising where a PKP is placed first. "
               "Check whether districts with 0 PKP are in practice supervised from a neighbouring district.")
    st.dataframe(pk.sort_values("Gap_vs_rule"), hide_index=True, column_config={
        "PKP_current": st.column_config.NumberColumn("PKP current", format="%d"), "PKP_required": st.column_config.NumberColumn("PKP required", format="%d"),
        "Gap_vs_rule": st.column_config.NumberColumn("Gap vs rule", format="%+d"), "Prorata_share": st.column_config.NumberColumn("Pro-rata share", format="%.2f"),
        "Gap_vs_prorata": st.column_config.NumberColumn("Gap vs pro-rata", format="%+.2f")})
    st.download_button("📥 Download PKP table (CSV)", csv_bytes(pk), "stage4_pkp_rule.csv", "text/csv")

    st.markdown("**Sensitivity to span of control**")
    sens = pd.DataFrame([{"Span (PPKP per PKP)": s_, "PKP required": int(np.maximum(min_pkp, np.ceil(ppkp_basis / s_)).sum()),
                          "Shortfall vs current": int(max(0, np.maximum(min_pkp, np.ceil(ppkp_basis / s_)).sum() - pkp_pool))} for s_ in (15, 20, 25, 30, 35, 40)])
    st.dataframe(sens, hide_index=True)

# =============================================================================
# TAB 5 - NOTES
# =============================================================================
with tab5:
    st.subheader("How to read this tool")
    st.markdown("""
**Question answered:** given the *current* pool of PKP/PPKP, how evenly does it match the recorded workload across districts?
It does **not** state how many officers the service needs - that requires validated time-per-task data.

**Stage 1** uses only raw counts per PPKP. **Stage 2** converts counts to minutes using editable placeholder values and draws each task's minutes
from its low-base-high range 1,000 times; a district is flagged only if it is in deficit/surplus in most draws.
**Stage 3** aggregates to state level first and proposes capped within-state transfers. **Stage 4** treats PKP with a span-of-control rule.

**Data still needed to finalise**
1. Minutes (or days) per unit for each task - time study or Delphi panel.
2. Share of each task performed by PKP vs PPKP, and the PKP scope that is not in the 20 listed tasks.
3. Task applicability per district (e.g. ports/airports for vessel/passenger work, island or interior districts) as an explicit flag.
4. The unit of the vessel/passenger column (inspections vs passenger counts) and of "Kawalan penyakit berjangkit".
5. Who covers districts with no PTJ (Kategori 'Tiada') that still record workload.
6. Meaning of "(5%)" in the *Penilaian risiko kesihatan persekitaran, PB & bencana* column header.
7. Optional: rurality/travel-time class per district (e.g. interior Sarawak/Sabah). Staff residence does not matter for a posting-based norm; travel time to work sites does.

**Why contract staff are a toggle:** contract (2025) staff are not permanent posts; results should be shown both ways.
""")
