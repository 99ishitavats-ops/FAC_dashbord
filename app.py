"""
Ampyr Solar Europe — FAC & 1st Year Testing Dashboard
Light / pastel Streamlit dashboard built on the FAC Tracker (Gantt) workbook.

Run:  streamlit run app.py
"""

from __future__ import annotations

import datetime as dt
import io
import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from fac_parser import parse_workbook

DATA_DEFAULT = os.path.join(os.path.dirname(__file__), "data", "FAC_Tracker_Gantt.xlsx")

# ---- pastel palette ---------------------------------------------------------
PASTEL = {
    "Prerequisite": "#C4B5FD",
    "Compliance": "#6EE7B7",
    "Closure": "#FDE68A",
    "Financial": "#FED7AA",
    "Technical": "#A5F3FC",
    "Testing": "#FBCFE8",
    "Document": "#BFDBFE",
}
STATUS_COLORS = {
    "Overdue": "#FCA5A5",   # soft red
    "Due soon": "#FDE68A",  # soft amber
    "On track": "#A7F3D0",  # soft green
    "TBC": "#E5E7EB",       # grey
}
STATUS_TEXT = {
    "Overdue": "#991B1B", "Due soon": "#92400E",
    "On track": "#065F46", "TBC": "#374151",
}

st.set_page_config(page_title="FAC & 1st Year Tracking",
                   page_icon="🌤️", layout="wide")

# ---- global styling ---------------------------------------------------------
st.markdown("""
<style>
:root { --card-radius: 16px; }
.stApp { background: linear-gradient(160deg,#fbfcff 0%,#f4f7fb 45%,#fef6fb 100%); }
section[data-testid="stSidebar"] { background: #f7f5ff; }
h1,h2,h3 { color:#3f3d56; font-weight:700; }
.block-container { padding-top: 1.6rem; }
.kpi {
  background:#ffffff; border-radius:var(--card-radius); padding:18px 20px;
  box-shadow:0 3px 14px rgba(120,120,160,0.10); border:1px solid #eef0f7;
  height:100%;
}
.kpi .label { font-size:0.82rem; color:#7c7a94; font-weight:600;
  text-transform:uppercase; letter-spacing:.4px; }
.kpi .value { font-size:2.0rem; font-weight:800; color:#3f3d56; line-height:1.1; }
.kpi .sub { font-size:0.8rem; color:#9a98b0; margin-top:2px; }
.pill { padding:3px 12px; border-radius:999px; font-size:0.78rem; font-weight:700;
  display:inline-block; }
.legend-chip { display:inline-block; padding:3px 10px; margin:2px 4px 2px 0;
  border-radius:999px; font-size:0.75rem; color:#3f3d56; font-weight:600; }
[data-testid="stMetricValue"] { color:#3f3d56; }
</style>
""", unsafe_allow_html=True)


# ---- data loading -----------------------------------------------------------
@st.cache_data(show_spinner=False)
def load(file_bytes: bytes | None, today: dt.date):
    src = io.BytesIO(file_bytes) if file_bytes else DATA_DEFAULT
    return parse_workbook(src, today=today)


def kpi(col, label, value, sub=""):
    col.markdown(
        f'<div class="kpi"><div class="label">{label}</div>'
        f'<div class="value">{value}</div>'
        f'<div class="sub">{sub}</div></div>', unsafe_allow_html=True)


def status_pill(s):
    bg = STATUS_COLORS.get(s, "#E5E7EB")
    fg = STATUS_TEXT.get(s, "#374151")
    return f'<span class="pill" style="background:{bg};color:{fg}">{s}</span>'


# ---- sidebar ----------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🌤️ FAC Dashboard")
    st.caption("Ampyr Solar Europe · FAC & 1st Year Testing")

    up = st.file_uploader("Upload updated FAC Tracker (.xlsx)", type=["xlsx"])
    st.caption("The dashboard refreshes automatically with your latest file.")

    today = st.date_input("Reference date ('today')", value=dt.date(2026, 7, 21),
                          help="Milestone statuses are calculated relative to this date.")
    st.divider()

file_bytes = up.getvalue() if up else None
try:
    data = load(file_bytes, today)
except Exception as e:  # noqa
    st.error(f"Could not read the workbook: {e}")
    st.stop()

sites = data["sites"].copy()
ms = data["milestones"].copy()
tasks = data["tasks"].copy()

# ---- sidebar filters --------------------------------------------------------
with st.sidebar:
    st.markdown("#### Filters")
    countries = sorted([c for c in sites["Country"].dropna().unique() if c])
    contractors = sorted([c for c in sites["Contractor"].dropna().unique() if c])
    sel_country = st.multiselect("Country", countries, default=countries)
    sel_contractor = st.multiselect("O&M Contractor", contractors, default=contractors)
    if up:
        st.success("Using your uploaded file ✔")
    else:
        st.info("Showing the bundled tracker. Upload a file to update.")

mask = sites["Country"].isin(sel_country) & sites["Contractor"].isin(sel_contractor)
sites_f = sites[mask]
keep = set(sites_f["Site"])
ms_f = ms[ms["Site"].isin(keep)]
tasks_f = tasks[tasks["Site"].isin(keep)]

# ---- header -----------------------------------------------------------------
st.markdown("# FAC & 1st Year Testing Tracker")
st.caption(f"Reference date: **{today:%d %b %Y}** · "
           f"{len(sites_f)} sites · {len(tasks_f)} scheduled activities")

# ---- KPI row ----------------------------------------------------------------
fac_dates = sites_f["FAC Date"].dropna()
fyt = sites_f["FYT Status"]
c1, c2, c3, c4, c5 = st.columns(5)
kpi(c1, "Sites tracked", len(sites_f),
    f"{sites_f['Country'].nunique()} countries")
kpi(c2, "FAC overdue", int((sites_f["FAC Status"] == "Overdue").sum()),
    "past target date")
kpi(c3, "FAC due ≤ 90 days", int((sites_f["FAC Status"] == "Due soon").sum()),
    "approaching")
kpi(c4, "1st-yr testing overdue", int((fyt == "Overdue").sum()),
    "FYT past due")
next_fac = fac_dates[fac_dates >= today].min() if (fac_dates >= today).any() else None
kpi(c5, "Next FAC", f"{next_fac:%d %b %Y}" if next_fac else "—",
    "earliest upcoming")

st.write("")

# ---- tabs -------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    ["📋 Site Overview", "📅 FAC Status & Dates",
     "🧪 1st Year Performance", "🗂️ Activities & Gantt"])

# === TAB 1 : Site overview ===================================================
with tab1:
    left, right = st.columns([1.35, 1])
    with left:
        st.subheader("Portfolio")
        show = sites_f.copy()
        show["FAC Date"] = show["FAC Date"].apply(lambda d: f"{d:%d %b %Y}" if pd.notnull(d) else "TBC")
        show["FYT Date"] = show["FYT Date"].apply(lambda d: f"{d:%d %b %Y}" if pd.notnull(d) else "TBC")
        show = show[["Site", "Country", "Contractor", "FYT Date", "FYT Status",
                     "FAC Date", "FAC Status"]]

        def _color(val):
            if val in STATUS_COLORS:
                return f"background-color:{STATUS_COLORS[val]};color:{STATUS_TEXT[val]};font-weight:600;"
            return ""
        st.dataframe(show.style.applymap(_color, subset=["FYT Status", "FAC Status"]),
                     use_container_width=True, hide_index=True, height=560)
    with right:
        st.subheader("FAC status mix")
        vc = sites_f["FAC Status"].value_counts()
        fig = px.pie(values=vc.values, names=vc.index, hole=0.55,
                     color=vc.index, color_discrete_map=STATUS_COLORS)
        fig.update_traces(textinfo="value+percent", textfont_size=13)
        fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=260,
                          legend=dict(orientation="h", y=-0.1),
                          paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Sites by O&M contractor")
        cc = sites_f["Contractor"].value_counts()
        figc = px.bar(x=cc.values, y=cc.index, orientation="h",
                      color=cc.index,
                      color_discrete_sequence=list(PASTEL.values()))
        figc.update_layout(showlegend=False, margin=dict(t=10, b=10, l=10, r=10),
                           height=240, xaxis_title="", yaxis_title="",
                           paper_bgcolor="rgba(0,0,0,0)",
                           plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(figc, use_container_width=True)

# === TAB 2 : FAC status & dates =============================================
with tab2:
    st.subheader("Milestone timeline (FYT → FAC)")
    tl = ms_f.dropna(subset=["Date"]).copy()
    if not tl.empty:
        # build a start/end per site connecting FYT to FAC
        piv = tl.pivot_table(index="Site", columns="Milestone", values="Date",
                             aggfunc="first")
        piv = piv.reset_index()
        rows = []
        for _, r in piv.iterrows():
            fyt_d = r.get("First Year Testing")
            fac_d = r.get("FAC")
            if pd.notnull(fyt_d) and pd.notnull(fac_d):
                rows.append({"Site": r["Site"], "Start": fyt_d, "Finish": fac_d})
        if rows:
            gdf = pd.DataFrame(rows).sort_values("Start")
            figt = px.timeline(gdf, x_start="Start", x_end="Finish", y="Site",
                               color_discrete_sequence=["#BFD7FF"])
            figt.update_yaxes(autorange="reversed")
            figt.add_vline(x=pd.Timestamp(today), line_width=2,
                           line_dash="dash", line_color="#F9A8B4")
            # milestone dots
            figt.add_trace(go.Scatter(
                x=tl[tl.Milestone == "First Year Testing"]["Date"],
                y=tl[tl.Milestone == "First Year Testing"]["Site"],
                mode="markers", name="FYT",
                marker=dict(size=12, color="#F9A8D4", line=dict(width=1, color="#fff"))))
            figt.add_trace(go.Scatter(
                x=tl[tl.Milestone == "FAC"]["Date"],
                y=tl[tl.Milestone == "FAC"]["Site"],
                mode="markers", name="FAC",
                marker=dict(size=13, color="#93C5FD", symbol="diamond",
                            line=dict(width=1, color="#fff"))))
            figt.update_layout(height=460, margin=dict(t=20, b=10, l=10, r=10),
                               paper_bgcolor="rgba(0,0,0,0)",
                               plot_bgcolor="rgba(0,0,0,0)",
                               legend=dict(orientation="h", y=1.08))
            st.plotly_chart(figt, use_container_width=True)
            st.caption("Bar spans First Year Testing start → Final Acceptance. "
                       "Pink dashed line = reference date.")

    colA, colB = st.columns(2)
    with colA:
        st.subheader("Upcoming milestones")
        up_ms = ms_f[(ms_f["Date"].notna()) & (ms_f["Date"] >= today)].sort_values("Date")
        if up_ms.empty:
            st.info("No upcoming milestones after the reference date.")
        else:
            for _, r in up_ms.head(12).iterrows():
                days = (r["Date"] - today).days
                st.markdown(
                    f"**{r['Site']}** — {r['Milestone']}  "
                    f"{status_pill(r['Status'])}<br>"
                    f"<span style='color:#8a889f'>{r['Date']:%d %b %Y} · in {days} days</span>",
                    unsafe_allow_html=True)
    with colB:
        st.subheader("Overdue milestones")
        od = ms_f[ms_f["Status"] == "Overdue"].sort_values("Date")
        if od.empty:
            st.success("Nothing overdue 🎉")
        else:
            for _, r in od.iterrows():
                days = (today - r["Date"]).days
                st.markdown(
                    f"**{r['Site']}** — {r['Milestone']}  "
                    f"{status_pill('Overdue')}<br>"
                    f"<span style='color:#8a889f'>{r['Date']:%d %b %Y} · {days} days ago</span>",
                    unsafe_allow_html=True)

# === TAB 3 : 1st year performance ===========================================
with tab3:
    st.subheader("First Year Testing status")
    fy = sites_f.copy()
    vc = fy["FYT Status"].value_counts()
    c1, c2, c3 = st.columns([1, 1, 1.2])
    kpi(c1, "FYT scheduled", int((fy["FYT Date"].notna()).sum()), "with a date set")
    kpi(c2, "FYT overdue", int((fy["FYT Status"] == "Overdue").sum()), "need attention")
    kpi(c3, "FYT due ≤ 90 days", int((fy["FYT Status"] == "Due soon").sum()), "approaching")
    st.write("")

    l, r = st.columns([1, 1])
    with l:
        st.markdown("**FYT status breakdown**")
        figp = px.pie(values=vc.values, names=vc.index, hole=0.55,
                      color=vc.index, color_discrete_map=STATUS_COLORS)
        figp.update_traces(textinfo="value+percent")
        figp.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10),
                           legend=dict(orientation="h", y=-0.1),
                           paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(figp, use_container_width=True)
    with r:
        st.markdown("**Testing activities by category**")
        tcat = tasks_f[tasks_f["Phase"] == "First Year Testing"]["Category"].value_counts()
        tcat = tcat[tcat.index != ""]
        figb = px.bar(x=tcat.index, y=tcat.values, color=tcat.index,
                      color_discrete_map=PASTEL)
        figb.update_layout(showlegend=False, height=300,
                           margin=dict(t=10, b=10, l=10, r=10),
                           xaxis_title="", yaxis_title="activities",
                           paper_bgcolor="rgba(0,0,0,0)",
                           plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(figb, use_container_width=True)

    st.markdown("**First Year Testing schedule by site**")
    fyt_tbl = fy[["Site", "Country", "Contractor", "FYT Date", "FYT Status"]].copy()
    fyt_tbl["FYT Date"] = fyt_tbl["FYT Date"].apply(
        lambda d: f"{d:%d %b %Y}" if pd.notnull(d) else "TBC")

    def _c2(v):
        return (f"background-color:{STATUS_COLORS[v]};color:{STATUS_TEXT[v]};font-weight:600;"
                if v in STATUS_COLORS else "")
    st.dataframe(fyt_tbl.style.applymap(_c2, subset=["FYT Status"]),
                 use_container_width=True, hide_index=True)

# === TAB 4 : activities & gantt =============================================
with tab4:
    st.subheader("Activity Gantt")
    chips = "".join(
        f'<span class="legend-chip" style="background:{c}">{k}</span>'
        for k, c in PASTEL.items())
    st.markdown(chips, unsafe_allow_html=True)

    site_opts = list(sites_f["Site"])
    sel_site = st.selectbox("Select a site", site_opts) if site_opts else None
    if sel_site:
        gt = tasks_f[(tasks_f["Site"] == sel_site) &
                     tasks_f["Start"].notna() & tasks_f["End"].notna()].copy()
        if gt.empty:
            st.info("No scheduled Gantt bars for this site (dates To Be Confirmed).")
        else:
            gt["Label"] = gt["Sub-Task"].where(gt["Sub-Task"] != "", gt["Task"])
            gt = gt.iloc[::-1]
            figg = px.timeline(gt, x_start="Start", x_end="End", y="Label",
                               color="Category", color_discrete_map=PASTEL,
                               hover_data=["Phase", "Task", "Responsibility", "Accepted By"])
            figg.add_vline(x=pd.Timestamp(today), line_width=2,
                           line_dash="dash", line_color="#F9A8B4")
            figg.update_layout(height=max(400, 22 * len(gt)),
                               margin=dict(t=20, b=10, l=10, r=10),
                               legend=dict(orientation="h", y=1.04),
                               paper_bgcolor="rgba(0,0,0,0)",
                               plot_bgcolor="rgba(0,0,0,0)",
                               yaxis_title="")
            st.plotly_chart(figg, use_container_width=True)

    st.subheader("All activities")
    tbl = tasks_f.copy()
    for c in ["Start", "End"]:
        tbl[c] = tbl[c].apply(lambda d: f"{d:%d %b %Y}" if pd.notnull(d) else "")
    fcat = st.multiselect("Filter by category",
                          sorted([c for c in tbl["Category"].unique() if c]))
    if fcat:
        tbl = tbl[tbl["Category"].isin(fcat)]
    st.dataframe(tbl[["Site", "Phase", "Task", "Sub-Task", "Category",
                      "Responsibility", "Accepted By", "Start", "End"]],
                 use_container_width=True, hide_index=True, height=460)
    st.download_button("⬇️ Download filtered activities (CSV)",
                       tbl.to_csv(index=False).encode(),
                       file_name="fac_activities.csv", mime="text/csv")

st.divider()
st.caption("Statuses are derived from milestone dates vs the reference date "
           "(Overdue < today ≤ Due soon ≤ 90 days < On track). "
           "The source workbook does not record completion, so labels reflect timing only.")
