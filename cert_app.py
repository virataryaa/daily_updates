"""
Certified Stocks Dashboard — KC Arabica & LRC Robusta
Run: streamlit run cert_app.py
"""

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit as st
from pathlib import Path

st.set_page_config(page_title="Certified Stocks", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
  [data-testid="stAppViewContainer"],[data-testid="stMain"],.main{background:#fafafa!important}
  [data-testid="stHeader"]{background:transparent!important}
  .block-container{padding-top:2.5rem!important;padding-bottom:1.5rem;max-width:1440px}
  hr{border:none!important;border-top:1px solid #e8e8ed!important;margin:.4rem 0!important}
  [data-testid="stRadio"] label{font-size:.75rem!important}
  .stDataFrame{font-size:.75rem}
</style>""", unsafe_allow_html=True)

_DATA = Path(__file__).parent
NAVY = "#0a2463"; RED = "#8b1a00"; AMBER = "#e8a020"
GREEN = "#1a7a1a"; DRED = "#c0392b"

PIE_COLORS = ["#0a2463","#c0553a","#4a7fb5","#e8a020","#1a7a1a",
              "#7b2d8b","#e85d04","#2dc4b2","#9b2226","#6a994e",
              "#3d405b","#f4a261","#264653","#e9c46a","#2a9d8f","#e76f51"]

_D = dict(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(family="-apple-system,Helvetica Neue,sans-serif", color="#1d1d1f", size=10))

ORIGIN_NAMES = {
    "BRZ":"Brazil","BUR":"Burundi","COL":"Colombia","COS":"Costa Rica",
    "ELS":"El Salvador","HON":"Honduras","IND":"India","MEX":"Mexico",
    "NIC":"Nicaragua","PAN":"Papua New Guinea","PER":"Peru","RWA":"Rwanda",
    "TAN":"Tanzania","UGA":"Uganda","VEN":"Venezuela","GUA":"Guatemala","TOT":"Total"
}
PORT_NAMES = {
    "AN":"Antwerp","BA":"Barcelona","HA":"Hamburg","HO":"Houston",
    "MI":"Miami","NO":"New Orleans","NY":"New York","TOT":"Total"
}
ORIGINS_NO_TOT = [o for o in ORIGIN_NAMES if o != "TOT"]
PORTS_NO_TOT   = [p for p in PORT_NAMES  if p != "TOT"]
GRADE_PORTS    = ["AN","HA","HO","MI","NO","NY"]

ORIGIN_ARB = {
    "BRZ":"Brazil NY 3/4","COL":"Colombia Excelso","HON":"Honduras HG",
    "IND":"India Cherry","UGA":"Uganda Drugar","PER":"Peru MCM","GUA":"Guatemala SHB",
}


def lbl(t, color=NAVY):
    return (f"<div style='background:{color};padding:5px 13px;border-radius:5px;"
            f"margin:0 0 5px 0;text-align:center'><span style='font-size:.78rem;"
            f"font-weight:500;letter-spacing:.07em;text-transform:uppercase;"
            f"color:#dde4f0'>{t}</span></div>")


def kpi(label, val, delta=None, delta_color=None):
    d_html = ""
    if delta is not None:
        dc = delta_color or (GREEN if delta >= 0 else DRED)
        sign = "▲" if delta >= 0 else "▼"
        d_html = f"<span style='font-size:.75rem;color:{dc};margin-left:6px'>{sign} {abs(delta):,}</span>"
    return (f"<div style='display:inline-flex;flex-direction:column;background:#f0f2f8;"
            f"border-radius:8px;padding:7px 14px;margin:3px;min-width:110px'>"
            f"<span style='font-size:.58rem;color:#6e6e73;text-transform:uppercase;"
            f"letter-spacing:.1em'>{label}</span>"
            f"<span style='font-size:.95rem;font-weight:700;color:{NAVY}'>{val}{d_html}</span>"
            f"</div>")


@st.cache_data
def load_kc():
    df = pd.read_parquet(_DATA / "cert_kc.parquet")
    df["Date"] = pd.to_datetime(df["Date"])
    return df.sort_values("Date").reset_index(drop=True)


@st.cache_data
def load_lrc():
    df = pd.read_parquet(_DATA / "cert_lrc.parquet")
    df["Date"] = pd.to_datetime(df["Date"])
    return df.sort_values("Date").reset_index(drop=True)


@st.cache_data
def load_rc_grading() -> pd.DataFrame:
    path = _DATA / "grading" / "RC_Grading_Feed.xlsx"
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_excel(path)
    # PanelDate is an Excel serial number — convert to datetime
    df["PanelDate"] = pd.to_datetime(df["PanelDate"], unit="D", origin="1899-12-30").dt.normalize()
    df["NoLots"]    = pd.to_numeric(df["NoLots"],    errors="coerce").fillna(0).astype(int)
    df["Allowance"] = pd.to_numeric(df["Allowance"], errors="coerce").fillna(0).astype(int)
    df["Class"]     = df["Class"].astype(str).str.strip()
    return df


# ── Commodity selector ────────────────────────────────────────────────────────
comm = st.radio("", ["Arabica", "Robusta"], horizontal=True, label_visibility="collapsed")
st.markdown("<hr>", unsafe_allow_html=True)

# =============================================================================
# KC / ARABICA VIEW
# =============================================================================
if comm == "Arabica":
    kc    = load_kc()
    max_d = kc["Date"].max().date()
    min_d = kc["Date"].min().date()

    # ── SECTION 1: Two-date change matrix ─────────────────────────────────
    st.markdown(lbl("Certified Stocks Change — Two-Date Matrix"), unsafe_allow_html=True)
    cd1, cd2, _ = st.columns([1, 1, 4])
    with cd1:
        older_date  = st.date_input("Older Date",
                                    value=max_d - pd.Timedelta(days=7),
                                    min_value=min_d, max_value=max_d, key="older_d")
    with cd2:
        latest_date = st.date_input("Latest Date",
                                    value=max_d,
                                    min_value=min_d, max_value=max_d, key="latest_d")

    older_sub  = kc[kc["Date"] <= pd.Timestamp(older_date)]
    latest_sub = kc[kc["Date"] <= pd.Timestamp(latest_date)]

    if len(older_sub) > 0 and len(latest_sub) > 0:
        older_row  = older_sub.iloc[-1]
        latest_row = latest_sub.iloc[-1]
        st.caption(f"Comparing  {older_row['Date'].strftime('%d/%m/%Y')}  →  "
                   f"{latest_row['Date'].strftime('%d/%m/%Y')}")

        change_rows = []
        for o in ORIGINS_NO_TOT:
            row = {"Origin": ORIGIN_NAMES[o]}
            for p in PORTS_NO_TOT + ["TOT"]:
                col = f"KC-{o}-{p}"
                v1  = older_row.get(col,  np.nan)
                v2  = latest_row.get(col, np.nan)
                row[PORT_NAMES[p]] = int(v2 - v1) if pd.notna(v1) and pd.notna(v2) else 0
            change_rows.append(row)
        # Total row
        tot_chg = {"Origin": "TOTAL"}
        for p in PORTS_NO_TOT + ["TOT"]:
            col = f"KC-TOT-{p}"
            v1  = older_row.get(col,  np.nan)
            v2  = latest_row.get(col, np.nan)
            tot_chg[PORT_NAMES[p]] = int(v2 - v1) if pd.notna(v1) and pd.notna(v2) else 0
        change_rows.append(tot_chg)

        chg_df = pd.DataFrame(change_rows).set_index("Origin")

        def color_chg(v):
            if v > 0:   return f"background:rgba(26,122,26,{min(abs(v)/5000,.65):.2f});color:#0d4d0d"
            elif v < 0: return f"background:rgba(192,57,43,{min(abs(v)/5000,.65):.2f});color:#5a0d0d"
            return "color:#bbb"

        def _chg_row_style(row):
            if row.name == "TOTAL":
                return ["font-weight:700; border-top:2px solid #ccc"] * len(row)
            return [""] * len(row)

        st.dataframe(
            chg_df.style
                  .applymap(color_chg)
                  .apply(_chg_row_style, axis=1)
                  .format(lambda v: f"{v:+,}" if v != 0 else "—"),
            use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── SECTION 2: Latest absolute matrix ────────────────────────────────
    st.markdown(lbl("Latest Certified Stocks — Origin × Port (bags)"), unsafe_allow_html=True)

    kc_sorted  = kc.dropna(subset=["KC-TOT-TOT"]).sort_values("Date")
    abs_latest = kc_sorted.iloc[-1]
    _gtv       = abs_latest.get("KC-TOT-TOT", np.nan)
    grand_tot  = float(_gtv) if pd.notna(_gtv) else 0.0

    def _safe_int(series_row, col):
        v = series_row.get(col, np.nan)
        return int(v) if pd.notna(v) else 0

    abs_rows = []
    for o in ORIGINS_NO_TOT:
        row = {"Origin": ORIGIN_NAMES[o]}
        for p in PORTS_NO_TOT:
            row[PORT_NAMES[p]] = _safe_int(abs_latest, f"KC-{o}-{p}")
        o_tot = _safe_int(abs_latest, f"KC-{o}-TOT")
        row["Total"] = o_tot
        row["% Share"] = round(o_tot / grand_tot * 100, 1) if grand_tot else 0.0
        abs_rows.append(row)
    # Total row
    tot_abs = {"Origin": "TOTAL"}
    for p in PORTS_NO_TOT:
        tot_abs[PORT_NAMES[p]] = _safe_int(abs_latest, f"KC-TOT-{p}")
    tot_abs["Total"]   = int(grand_tot)
    tot_abs["% Share"] = 100.0
    abs_rows.append(tot_abs)
    # % per port row
    pct_row = {"Origin": "% per Port"}
    for p in PORTS_NO_TOT:
        pv = float(_safe_int(abs_latest, f"KC-TOT-{p}"))
        pct_row[PORT_NAMES[p]] = round(pv / grand_tot * 100, 1) if grand_tot else 0.0
    pct_row["Total"]   = 100.0
    pct_row["% Share"] = 0.0
    abs_rows.append(pct_row)

    abs_df = pd.DataFrame(abs_rows).set_index("Origin")

    def fmt_abs(v):
        if isinstance(v, float) and v != int(v):
            return f"{v:.1f}%"
        return f"{int(v):,}" if pd.notna(v) and v != 0 else "—"

    def _abs_row_style(row):
        if row.name in ("TOTAL", "% per Port"):
            return ["font-weight:700; border-top:2px solid #ccc"] * len(row)
        return [""] * len(row)

    st.dataframe(
        abs_df.style
              .format(fmt_abs)
              .apply(_abs_row_style, axis=1)
              .applymap(
                  lambda v: "background:#e8ecf8;font-weight:700" if isinstance(v, (int, float)) and v > 0 else "",
                  subset=["% Share"]),
        use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Date slider for time-series sections ──────────────────────────────
    date_range = st.slider("Date range", min_value=min_d, max_value=max_d,
                           value=(pd.Timestamp("2024-01-01").date(), max_d),
                           format="YYYY-MM-DD")
    dff = kc[(kc["Date"] >= pd.Timestamp(date_range[0])) &
             (kc["Date"] <= pd.Timestamp(date_range[1]))]

    # KPI strip
    tot_now  = int(kc["KC-TOT-TOT"].dropna().iloc[-1])
    tot_prev = int(kc["KC-TOT-TOT"].dropna().iloc[-2])
    hon_now  = int(kc["KC-HON-TOT"].dropna().iloc[-1])
    pend_now = int(kc["KC-TOT-PENDING"].dropna().iloc[-1]) if "KC-TOT-PENDING" in kc.columns else 0
    st.markdown(
        kpi("Latest Date", kc["Date"].max().strftime("%d/%m/%Y")) +
        kpi("KC Total Certs", f"{tot_now/1000:.1f}k bags", tot_now - tot_prev) +
        kpi("Honduras", f"{hon_now/1000:.1f}k", hon_now - int(kc["KC-HON-TOT"].dropna().iloc[-2])) +
        kpi("Honduras %", f"{hon_now/tot_now*100:.1f}%") +
        kpi("Total Pending", f"{pend_now/1000:.1f}k bags"),
        unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── SECTION 3: Total certs + Pie (fixed) ─────────────────────────────
    c1, c2 = st.columns([3, 1])
    with c1:
        st.markdown(lbl("KC Total Certified Stocks (bags)"), unsafe_allow_html=True)
        valid = dff[["Date","KC-TOT-TOT"]].dropna()
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=valid["Date"], y=valid["KC-TOT-TOT"],
            mode="lines", line=dict(color=NAVY, width=2),
            fill="tozeroy", fillcolor="rgba(10,36,99,0.07)",
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:,.0f} bags</b><extra></extra>"))
        fig.update_layout(height=240, margin=dict(t=8,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=9)),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), tickformat=","),
            **_D)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown(lbl("Origin Share (%)"), unsafe_allow_html=True)
        shares = {}
        for o in ORIGINS_NO_TOT:
            col = f"KC-{o}-TOT"
            if col in kc.columns:
                v = kc[col].dropna()
                if len(v) > 0 and float(v.iloc[-1]) > 0:
                    shares[ORIGIN_NAMES[o]] = float(v.iloc[-1])
        if shares:
            sorted_shares = dict(sorted(shares.items(), key=lambda x: x[1], reverse=True))
            fig_pie = go.Figure(go.Pie(
                labels=list(sorted_shares.keys()), values=list(sorted_shares.values()),
                hole=0.5, textinfo="label+percent",
                textfont=dict(size=7.5, color="#1d1d1f"),
                marker=dict(colors=PIE_COLORS[:len(sorted_shares)],
                            line=dict(color="white", width=1.5)),
                sort=True))
            fig_pie.update_layout(height=240, margin=dict(t=8,b=8,l=0,r=0),
                showlegend=False,
                paper_bgcolor="#fafafa", plot_bgcolor="#fafafa",
                font=dict(color="#1d1d1f", size=10))
            st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── SECTION 4: Multi-origin certs + Pending ───────────────────────────
    st.markdown(lbl("KC Certified Stocks by Origin"), unsafe_allow_html=True)

    avail_origins = [o for o in ORIGINS_NO_TOT
                     if f"KC-{o}-TOT" in kc.columns and kc[f"KC-{o}-TOT"].dropna().sum() > 0]
    sel_origins = st.multiselect(
        "Select origins:",
        options=avail_origins,
        default=[o for o in ["HON","BRZ","COL","GUA"] if o in avail_origins],
        format_func=lambda x: ORIGIN_NAMES[x])

    if sel_origins:
        fig_mo = go.Figure()
        pal = px.colors.qualitative.Set1
        for i, o in enumerate(sel_origins):
            col = f"KC-{o}-TOT"
            s   = dff[["Date", col]].dropna(subset=[col])
            fig_mo.add_trace(go.Scatter(x=s["Date"], y=s[col],
                name=ORIGIN_NAMES[o], mode="lines",
                line=dict(color=pal[i % len(pal)], width=1.8),
                hovertemplate=f"{ORIGIN_NAMES[o]}<br>%{{x|%d/%m/%Y}}<br>"
                              f"<b>%{{y:,.0f}} bags</b><extra></extra>"))
        fig_mo.update_layout(height=240,
            legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
            margin=dict(t=10,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=9)),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), tickformat=","),
            **_D)
        st.plotly_chart(fig_mo, use_container_width=True)

    # Pending time series
    if "KC-TOT-PENDING" in kc.columns:
        st.markdown(lbl("KC Total Pending Certification (bags)"), unsafe_allow_html=True)
        pend_ts = dff[["Date","KC-TOT-PENDING"]].dropna(subset=["KC-TOT-PENDING"])
        fig_pend = go.Figure()
        fig_pend.add_trace(go.Scatter(
            x=pend_ts["Date"], y=pend_ts["KC-TOT-PENDING"] / 1000,
            mode="lines", line=dict(color=AMBER, width=2),
            fill="tozeroy", fillcolor="rgba(232,160,32,0.1)",
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:.1f}k bags</b><extra></extra>"))
        fig_pend.update_layout(height=200, margin=dict(t=10,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=9)),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="k bags"),
            **_D)
        st.plotly_chart(fig_pend, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── SECTION 5: Origin Deep-Dive (collapsible) ─────────────────────────
    with st.expander("Origin Deep-Dive", expanded=False):
        od_origin = st.selectbox(
            "Select origin:",
            options=avail_origins,
            index=avail_origins.index("HON") if "HON" in avail_origins else 0,
            format_func=lambda x: ORIGIN_NAMES[x],
            key="od_sel")

        od_col  = f"KC-{od_origin}-TOT"
        od_data = dff[["Date","KC-TOT-TOT", od_col]].dropna(subset=[od_col]).copy()
        od_data["pct"]    = od_data[od_col] / od_data["KC-TOT-TOT"] * 100
        od_data["chg"]    = od_data[od_col].diff()
        od_data["roll20"] = od_data["chg"].rolling(20).sum() / 1000

        full_od = kc[[od_col]].dropna()[od_col]
        z_now   = (float(full_od.iloc[-1]) - float(full_od.mean())) / float(full_od.std()) \
                  if full_od.std() > 0 else 0.0
        z_color = DRED if abs(z_now) > 1.5 else NAVY

        st.markdown(
            f"<div style='font-size:.72rem;color:#6e6e73;margin-bottom:6px'>"
            f"Latest: <b>{float(full_od.iloc[-1])/1000:.1f}k bags</b> &nbsp;·&nbsp; "
            f"Mean (full hist): <b>{float(full_od.mean())/1000:.1f}k</b> &nbsp;·&nbsp; "
            f"Z-Score: <b style='color:{z_color}'>{z_now:+.2f}σ</b></div>",
            unsafe_allow_html=True)

        r1c1, r1c2 = st.columns(2)
        with r1c1:
            st.markdown(lbl(f"{ORIGIN_NAMES[od_origin]} vs Total (k bags)"), unsafe_allow_html=True)
            fig_od1 = make_subplots(specs=[[{"secondary_y": True}]])
            fig_od1.add_trace(go.Scatter(x=od_data["Date"], y=od_data[od_col]/1000,
                name=ORIGIN_NAMES[od_origin], mode="lines",
                line=dict(color=NAVY, width=2)), secondary_y=False)
            fig_od1.add_trace(go.Scatter(x=od_data["Date"], y=od_data["KC-TOT-TOT"]/1000,
                name="Total", mode="lines",
                line=dict(color="#aaa", width=1.5, dash="dot")), secondary_y=True)
            fig_od1.update_layout(height=260,
                legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
                margin=dict(t=10,b=8,l=4,r=4), **_D)
            fig_od1.update_yaxes(title_text=f"{ORIGIN_NAMES[od_origin]} (k)", secondary_y=False,
                tickfont=dict(size=9), showgrid=True, gridcolor="#f0f0f0")
            fig_od1.update_yaxes(title_text="Total (k)", secondary_y=True,
                tickfont=dict(size=9), showgrid=False)
            fig_od1.update_xaxes(showgrid=False, tickfont=dict(size=9))
            st.plotly_chart(fig_od1, use_container_width=True)

        with r1c2:
            st.markdown(lbl(f"{ORIGIN_NAMES[od_origin]} % of Total"), unsafe_allow_html=True)
            fig_od2 = go.Figure()
            fig_od2.add_trace(go.Scatter(x=od_data["Date"], y=od_data["pct"],
                mode="lines", line=dict(color=NAVY, width=2),
                fill="tozeroy", fillcolor="rgba(10,36,99,0.07)",
                hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:.1f}%</b><extra></extra>"))
            fig_od2.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="%"),
                **_D)
            st.plotly_chart(fig_od2, use_container_width=True)

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            st.markdown(lbl(f"{ORIGIN_NAMES[od_origin]} Rolling 20d Change (k bags)"), unsafe_allow_html=True)
            roll_s = od_data[["Date","roll20"]].dropna(subset=["roll20"])
            fig_od3 = go.Figure(go.Bar(
                x=roll_s["Date"], y=roll_s["roll20"],
                marker_color=[GREEN if v >= 0 else DRED for v in roll_s["roll20"]],
                hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:+.1f}k bags</b><extra></extra>"))
            fig_od3.add_hline(y=0, line_color="#cccccc", line_width=1)
            fig_od3.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="k bags"),
                **_D)
            st.plotly_chart(fig_od3, use_container_width=True)

        with r2c2:
            st.markdown(lbl(f"{ORIGIN_NAMES[od_origin]} Distribution"), unsafe_allow_html=True)
            fig_od4 = go.Figure()
            fig_od4.add_trace(go.Histogram(x=full_od, nbinsx=30,
                marker_color=NAVY, opacity=0.8,
                hovertemplate="<b>%{y} days</b> at ~%{x:,.0f} bags<extra></extra>"))
            fig_od4.add_vline(x=float(full_od.iloc[-1]),
                line_color=z_color, line_width=2, line_dash="dash",
                annotation_text=f"Now ({z_now:+.2f}σ)",
                annotation_position="top right",
                annotation_font=dict(color=z_color, size=8))
            fig_od4.add_vline(x=float(full_od.mean()),
                line_color="#888", line_width=1,
                annotation_text="Mean",
                annotation_position="top left",
                annotation_font=dict(size=8))
            fig_od4.update_layout(height=260, margin=dict(t=30,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9), title="bags"),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="days"),
                **_D)
            st.plotly_chart(fig_od4, use_container_width=True)

        # Arb differential for selected origin
        if od_origin in ORIGIN_ARB:
            arb_name = ORIGIN_ARB[od_origin]
            if arb_name in dff.columns:
                st.markdown(lbl(f"{ORIGIN_NAMES[od_origin]} — Origin Differential (cts/lb)"),
                            unsafe_allow_html=True)
                arb_s = dff[["Date", arb_name]].dropna(subset=[arb_name])
                fig_arb_od = go.Figure()
                fig_arb_od.add_trace(go.Scatter(x=arb_s["Date"], y=arb_s[arb_name],
                    mode="lines", line=dict(color=NAVY, width=2),
                    fill="tozeroy", fillcolor="rgba(10,36,99,0.07)",
                    hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:+.0f} cts/lb</b><extra></extra>"))
                fig_arb_od.add_hline(y=0, line_color="#cccccc", line_width=1)
                fig_arb_od.update_layout(height=200, margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="cts/lb"),
                    **_D)
                st.plotly_chart(fig_arb_od, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── SECTION 6: Arbitrage bar + time series ────────────────────────────
    arb_cols = [c for c in kc.columns
                if not c.startswith("KC-") and c not in ["Date","KC_Price"]]

    ca1, ca2 = st.columns([3, 2])
    with ca1:
        st.markdown(lbl("Origin Differentials vs KC Futures (cts/lb)"), unsafe_allow_html=True)
        if arb_cols:
            arb_latest, arb_prev = {}, {}
            for c in arb_cols:
                s = kc[["Date", c]].dropna(subset=[c]).sort_values("Date")
                if len(s) >= 2:
                    arb_latest[c] = float(s[c].iloc[-1])
                    arb_prev[c]   = float(s[c].iloc[-2])
            if arb_latest:
                names  = list(arb_latest.keys())
                vals   = list(arb_latest.values())
                chgs   = [vals[i] - arb_prev.get(names[i], vals[i]) for i in range(len(vals))]
                abbrev = {"Brazil NY 3/4":"BRZ NY","Brazil Santos":"BRZ STS",
                          "Brazil Rio 15/16":"BRZ R1","Brazil Rio 17/18":"BRZ R2",
                          "Ethiopia Djimmah":"ETH","Uganda Drugar":"UGA",
                          "India Cherry":"IND","Colombia Excelso":"COL",
                          "Honduras HG":"HON","Peru MCM":"PER","Guatemala SHB":"GUA"}
                short_names = [abbrev.get(n, n[:8]) for n in names]
                fig_arb = go.Figure(go.Bar(
                    x=short_names, y=vals,
                    marker_color=[GREEN if v >= 0 else DRED for v in vals],
                    text=[f"{v:+.0f}" for v in vals],
                    textposition="outside", textfont=dict(size=8.5),
                    customdata=[[chgs[i]] for i in range(len(vals))],
                    hovertemplate="%{x}<br><b>%{y:+.0f} cts/lb</b><br>DoD: %{customdata[0]:+.0f}<extra></extra>"))
                fig_arb.add_hline(y=0, line_color="#cccccc", line_width=1)
                fig_arb.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="cts/lb"),
                    **_D)
                st.plotly_chart(fig_arb, use_container_width=True)

    with ca2:
        with st.expander("Differentials — Full Time Series", expanded=False):
            if arb_cols:
                fig_ats = go.Figure()
                pal_ats = px.colors.qualitative.Set1
                for i, c in enumerate(arb_cols):
                    s = dff[["Date", c]].dropna(subset=[c])
                    if len(s):
                        fig_ats.add_trace(go.Scatter(x=s["Date"], y=s[c], name=c,
                            mode="lines",
                            line=dict(color=pal_ats[i % len(pal_ats)], width=1.5)))
                fig_ats.add_hline(y=0, line_color="#cccccc", line_width=1)
                fig_ats.update_layout(height=260,
                    legend=dict(orientation="h", y=1.02, x=0, font=dict(size=7.5)),
                    margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="cts/lb"),
                    **_D)
                st.plotly_chart(fig_ats, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── SECTION 7: Grading Flow Table ─────────────────────────────────────
    st.markdown(lbl("KC Grading Flow"), unsafe_allow_html=True)

    pass_cols_g = [f"KC-{p}-PASSGRAD" for p in GRADE_PORTS if f"KC-{p}-PASSGRAD" in kc.columns]
    fail_cols_g = [f"KC-{p}-FAILGRAD" for p in GRADE_PORTS if f"KC-{p}-FAILGRAD" in kc.columns]

    if pass_cols_g or fail_cols_g:
        grade_base = kc[["Date","KC-TOT-PENDING","KC-TOT-TOT"] +
                        pass_cols_g + fail_cols_g].copy().sort_values("Date")
        grade_base["Passed"] = grade_base[pass_cols_g].sum(axis=1) if pass_cols_g else 0
        grade_base["Failed"] = grade_base[fail_cols_g].sum(axis=1) if fail_cols_g else 0
        grade_base = grade_base[grade_base["Passed"] + grade_base["Failed"] > 0].copy()

        if len(grade_base) > 0:
            grade_base["Pct_Fail"]      = grade_base["Failed"] / (grade_base["Passed"] + grade_base["Failed"]) * 100
            grade_base["Certs_Chg"]     = grade_base["KC-TOT-TOT"].diff()
            grade_base["Pend_Chg"]      = grade_base["KC-TOT-PENDING"].diff()
            grade_base["Fresh_Pending"] = grade_base["Pend_Chg"] + grade_base["Passed"] + grade_base["Failed"]
            grade_base["Imp_Decerts"]   = grade_base["Passed"] - grade_base["Certs_Chg"]

            disp = grade_base.tail(40).sort_values("Date", ascending=False).copy()
            disp["Date_str"] = disp["Date"].dt.strftime("%d/%m/%Y")

            tbl = disp[["Date_str","Passed","Failed","Pct_Fail",
                         "KC-TOT-PENDING","KC-TOT-TOT","Certs_Chg",
                         "Fresh_Pending","Imp_Decerts"]].rename(columns={
                "Date_str":"Date","Pct_Fail":"% Fail","KC-TOT-PENDING":"Pending",
                "KC-TOT-TOT":"Certs","Certs_Chg":"Certs Δ",
                "Fresh_Pending":"Fresh Pend","Imp_Decerts":"Impl Decerts"
            }).reset_index(drop=True)

            def _fmt_signed(v):
                if pd.isna(v): return "—"
                return f"{v:+,.0f}"

            def _color_signed(v):
                if isinstance(v, (int, float)) and pd.notna(v):
                    if v > 0: return "color:#1a7a1a"
                    elif v < 0: return "color:#c0392b"
                return ""

            st.dataframe(
                tbl.style
                   .applymap(_color_signed, subset=["Certs Δ","Fresh Pend","Impl Decerts"])
                   .format({"Passed":"{:,.0f}","Failed":"{:,.0f}","% Fail":"{:.1f}%",
                             "Pending":"{:,.0f}","Certs":"{:,.0f}",
                             "Certs Δ":_fmt_signed,"Fresh Pend":_fmt_signed,"Impl Decerts":_fmt_signed}),
                use_container_width=True, height=420)

            # Charts
            g1, g2 = st.columns(2)
            with g1:
                st.markdown(lbl("Fresh Pending & Implied Decerts Over Time"), unsafe_allow_html=True)
                fp_s = grade_base[["Date","Fresh_Pending","Imp_Decerts"]].dropna(subset=["Fresh_Pending"])
                fig_fp = go.Figure()
                fig_fp.add_trace(go.Scatter(x=fp_s["Date"], y=fp_s["Fresh_Pending"]/1000,
                    name="Fresh Pending (k)", mode="lines+markers",
                    marker=dict(size=4), line=dict(color=AMBER, width=1.8)))
                fig_fp.add_trace(go.Scatter(x=fp_s["Date"], y=fp_s["Imp_Decerts"]/1000,
                    name="Impl Decerts (k)", mode="lines+markers",
                    marker=dict(size=4), line=dict(color=DRED, width=1.8)))
                fig_fp.add_hline(y=0, line_color="#cccccc", line_width=1)
                fig_fp.update_layout(height=260,
                    legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
                    margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="k bags"),
                    **_D)
                st.plotly_chart(fig_fp, use_container_width=True)

            with g2:
                st.markdown(lbl("Passed vs Failed Grading Queue"), unsafe_allow_html=True)
                pf_s = grade_base[["Date","Passed","Failed"]].dropna(subset=["Passed"])
                fig_pf = go.Figure()
                fig_pf.add_trace(go.Bar(x=pf_s["Date"], y=pf_s["Passed"]/1000,
                    name="Passed (k)", marker_color="rgba(26,122,26,0.75)"))
                fig_pf.add_trace(go.Bar(x=pf_s["Date"], y=pf_s["Failed"]/1000,
                    name="Failed (k)", marker_color="rgba(192,57,43,0.75)"))
                fig_pf.update_layout(height=260, barmode="group",
                    legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
                    margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="k bags"),
                    **_D)
                st.plotly_chart(fig_pf, use_container_width=True)

# =============================================================================
# LRC / ROBUSTA VIEW
# =============================================================================
else:
    tab_cs, tab_gr = st.tabs(["Certified Stocks", "Grading"])

    # =========================================================================
    # ROBUSTA TAB 1 — Certified Stocks (existing content)
    # =========================================================================
    with tab_cs:
        lrc   = load_lrc()
        min_d = lrc["Date"].min().date()
        max_d = lrc["Date"].max().date()

        date_range = st.slider("Date range", min_value=min_d, max_value=max_d,
                               value=(pd.Timestamp("2020-01-01").date(), max_d),
                               format="YYYY-MM-DD")
        dff = lrc[(lrc["Date"] >= pd.Timestamp(date_range[0])) &
                  (lrc["Date"] <= pd.Timestamp(date_range[1]))]
        st.markdown("<hr>", unsafe_allow_html=True)

        # KPI strip
        tot_now  = int(lrc["LRC-TOT-VG"].dropna().iloc[-1])
        tot_prev = int(lrc["LRC-TOT-VG"].dropna().iloc[-2])
        lrc_port_cols = {"ANT":"LRC-ANT-VG","LON":"LRC-LON-VG","FEL":"LRC-FEL-VG","BAR":"LRC-BAR-VG"}
        port_vals_now = {}
        for pn, col in lrc_port_cols.items():
            if col in lrc.columns:
                s = lrc[col].dropna()
                if len(s): port_vals_now[pn] = int(s.iloc[-1])
        ant_pct = port_vals_now.get("ANT", 0) / tot_now * 100 if tot_now else 0

        st.markdown(
            kpi("Latest Date", lrc["Date"].max().strftime("%d/%m/%Y")) +
            kpi("LRC Total Certs", f"{tot_now:,}", tot_now - tot_prev) +
            kpi("Antwerp", f"{port_vals_now.get('ANT',0):,}") +
            kpi("ANT %", f"{ant_pct:.1f}%") +
            kpi("London", f"{port_vals_now.get('LON',0):,}"),
            unsafe_allow_html=True)
        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Row 1: Total + Port share pie ─────────────────────────────────────
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(lbl("LRC Total Certified Stocks — Full History (bags)", RED), unsafe_allow_html=True)
            valid_lt = dff[["Date","LRC-TOT-VG"]].dropna()
            fig_lt = go.Figure()
            fig_lt.add_trace(go.Scatter(x=valid_lt["Date"], y=valid_lt["LRC-TOT-VG"],
                mode="lines", line=dict(color=RED, width=2),
                fill="tozeroy", fillcolor="rgba(139,26,0,0.07)",
                hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:,} bags</b><extra></extra>"))
            fig_lt.update_layout(height=240, margin=dict(t=8,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), tickformat=","),
                **_D)
            st.plotly_chart(fig_lt, use_container_width=True)

        with c2:
            st.markdown(lbl("Port Share (%)", RED), unsafe_allow_html=True)
            all_lrc_ports = [p for p in ["ANT","LON","FEL","BAR","AMS","BRE","HAM","LIV","NOR","ROT","TRI"]
                             if f"LRC-{p}-VG" in lrc.columns]
            lrc_shares = {}
            for p in all_lrc_ports:
                s = lrc[f"LRC-{p}-VG"].dropna()
                v = int(s.iloc[-1]) if len(s) else 0
                if v > 0: lrc_shares[p] = v
            if lrc_shares:
                fig_lpie = go.Figure(go.Pie(
                    labels=list(lrc_shares.keys()), values=list(lrc_shares.values()),
                    hole=0.5, textinfo="label+percent",
                    textfont=dict(size=8, color="#1d1d1f"),
                    marker=dict(colors=PIE_COLORS[:len(lrc_shares)],
                                line=dict(color="white", width=1.5)),
                    sort=True))
                fig_lpie.update_layout(height=240, margin=dict(t=8,b=8,l=0,r=0),
                    showlegend=False,
                    paper_bgcolor="#fafafa", plot_bgcolor="#fafafa",
                    font=dict(color="#1d1d1f", size=10))
                st.plotly_chart(fig_lpie, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Row 2: Port stocks bar + DoD change ──────────────────────────────
        c3, c4 = st.columns(2)
        with c3:
            st.markdown(lbl("Current Stocks by Port (bags)", RED), unsafe_allow_html=True)
            pb_names, pb_vals = [], []
            for p in all_lrc_ports:
                s = lrc[f"LRC-{p}-VG"].dropna()
                v = int(s.iloc[-1]) if len(s) else 0
                if v > 0:
                    pb_names.append(p); pb_vals.append(v)
            if pb_names:
                fig_pb = go.Figure(go.Bar(
                    x=pb_names, y=pb_vals,
                    marker_color=PIE_COLORS[:len(pb_names)],
                    text=[f"{v:,}" for v in pb_vals],
                    textposition="outside", textfont=dict(size=9)))
                fig_pb.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=10)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"),
                    **_D)
                st.plotly_chart(fig_pb, use_container_width=True)

        with c4:
            st.markdown(lbl("Day-over-Day Change by Port (bags)", RED), unsafe_allow_html=True)
            lrc_sorted2 = lrc.dropna(subset=["LRC-TOT-VG"]).sort_values("Date")
            chg_names, chg_vals = [], []
            if len(lrc_sorted2) >= 2:
                lat = lrc_sorted2.iloc[-1]; prv = lrc_sorted2.iloc[-2]
                for p in all_lrc_ports:
                    col = f"LRC-{p}-VG"
                    nv  = lat.get(col, np.nan); pv2 = prv.get(col, np.nan)
                    if pd.notna(nv) and pd.notna(pv2):
                        chg = int(nv - pv2)
                        if chg != 0:
                            chg_names.append(p); chg_vals.append(chg)
            if chg_names:
                fig_pc = go.Figure(go.Bar(
                    x=chg_names, y=chg_vals,
                    marker_color=[GREEN if v >= 0 else DRED for v in chg_vals],
                    text=[f"{v:+,}" for v in chg_vals],
                    textposition="outside", textfont=dict(size=9)))
                fig_pc.add_hline(y=0, line_color="#cccccc", line_width=1)
                fig_pc.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=10)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"),
                    **_D)
                st.plotly_chart(fig_pc, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Row 3: Port time series + Rolling 20d ────────────────────────────
        c5, c6 = st.columns(2)
        with c5:
            st.markdown(lbl("Port Stocks Over Time", RED), unsafe_allow_html=True)
            fig_ts = go.Figure()
            ts_port_colors = {"ANT": RED, "LON": NAVY, "FEL": "#4a7fb5", "BAR": "#e8c96a"}
            for pn, col in {"ANT":"LRC-ANT-VG","LON":"LRC-LON-VG",
                             "FEL":"LRC-FEL-VG","BAR":"LRC-BAR-VG"}.items():
                if col in dff.columns:
                    s = dff[["Date", col]].dropna(subset=[col])
                    if len(s):
                        fig_ts.add_trace(go.Scatter(x=s["Date"], y=s[col],
                            name=pn, mode="lines",
                            line=dict(color=ts_port_colors.get(pn,"#aaa"), width=1.8)))
            fig_ts.update_layout(height=260,
                legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
                margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"),
                **_D)
            st.plotly_chart(fig_ts, use_container_width=True)

        with c6:
            st.markdown(lbl("Rolling 20-Day Change — Total / ANT / UK", RED), unsafe_allow_html=True)
            roll_df = dff[["Date","LRC-TOT-VG","LRC-ANT-VG","LRC-LON-VG","LRC-FEL-VG"]].copy().sort_values("Date")
            roll_df["20d TOT"] = roll_df["LRC-TOT-VG"].diff().rolling(20).sum()
            roll_df["20d ANT"] = roll_df["LRC-ANT-VG"].diff().rolling(20).sum()
            roll_df["20d UK"]  = (roll_df["LRC-LON-VG"].fillna(0) +
                                   roll_df["LRC-FEL-VG"].fillna(0)).diff().rolling(20).sum()
            fig_roll = go.Figure()
            for col, color, label in [("20d TOT", RED, "20d Total"),
                                        ("20d ANT", NAVY, "20d Antwerp"),
                                        ("20d UK", "#4a7fb5", "20d UK (Lon+Fel)")]:
                s = roll_df[["Date", col]].dropna(subset=[col])
                fig_roll.add_trace(go.Scatter(x=s["Date"], y=s[col],
                    name=label, mode="lines", line=dict(color=color, width=1.8)))
            fig_roll.add_hline(y=0, line_color="#cccccc", line_width=1)
            fig_roll.update_layout(height=260,
                legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
                margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"),
                **_D)
            st.plotly_chart(fig_roll, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── SECTION 8: Monthly Change Table + Bar ────────────────────────────
        st.markdown(lbl("LRC Monthly Certified Stocks Change", RED), unsafe_allow_html=True)

        lrc_m = lrc[["Date","LRC-TOT-VG"]].dropna().copy()
        lrc_m["YM"] = lrc_m["Date"].dt.to_period("M")
        monthly = lrc_m.groupby("YM")["LRC-TOT-VG"].agg(
            Start="first", End="last", High="max", Low="min"
        ).reset_index()
        monthly["Δ Bags"] = monthly["End"] - monthly["Start"]
        monthly["Δ %"]    = monthly["Δ Bags"] / monthly["Start"] * 100
        monthly["Month"]  = monthly["YM"].astype(str)
        monthly = monthly.sort_values("Month", ascending=False)

        def _fmt_s(v):
            if pd.isna(v): return "—"
            return f"{v:+,.0f}"
        def _fmt_pct(v):
            if pd.isna(v): return "—"
            return f"{v:+.1f}%"
        def _col_signed(v):
            if isinstance(v, (int, float)) and pd.notna(v):
                if v > 0: return "color:#1a7a1a"
                elif v < 0: return "color:#c0392b"
            return ""

        m_disp = monthly[["Month","Start","End","High","Low","Δ Bags","Δ %"]].reset_index(drop=True)
        st.dataframe(
            m_disp.style
                  .map(_col_signed, subset=["Δ Bags","Δ %"])
                  .format({"Start":"{:,.0f}","End":"{:,.0f}","High":"{:,.0f}","Low":"{:,.0f}",
                           "Δ Bags":_fmt_s,"Δ %":_fmt_pct}),
            use_container_width=True, height=400)

        # Monthly bar chart (show last 36 months)
        m_bar = monthly.sort_values("Month").tail(36)
        fig_mb = go.Figure(go.Bar(
            x=m_bar["Month"], y=m_bar["Δ Bags"],
            marker_color=[GREEN if v >= 0 else DRED for v in m_bar["Δ Bags"]],
            text=[f"{v:+,.0f}" if abs(v) > 50 else "" for v in m_bar["Δ Bags"]],
            textposition="outside", textfont=dict(size=7.5),
            hovertemplate="%{x}<br><b>%{y:+,.0f} bags</b><extra></extra>"))
        fig_mb.add_hline(y=0, line_color="#cccccc", line_width=1)
        fig_mb.update_layout(height=280, margin=dict(t=10,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=8), tickangle=-45),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags Δ"),
            **_D)
        st.plotly_chart(fig_mb, use_container_width=True)

    # =========================================================================
    # ROBUSTA TAB 2 — Grading
    # =========================================================================
    with tab_gr:
        gr = load_rc_grading()
        if gr.empty:
            st.warning("Grading feed not found — expected at: RC/Grading/RC_Grading_Feed.xlsx")
        else:
            all_panel_dates = sorted(gr["PanelDate"].unique(), reverse=True)

            # ── Filters ──────────────────────────────────────────────────────
            gf1, gf2, gf3 = st.columns([2, 2, 3])
            with gf1:
                sel_panel = st.selectbox(
                    "Panel Date", all_panel_dates,
                    format_func=lambda d: pd.Timestamp(d).strftime("%d/%m/%Y"),
                    key="gr_panel_date",
                )
            with gf2:
                sel_ukcont = st.multiselect(
                    "Exchange", ["C", "UK", "US"], default=["C", "UK", "US"],
                    format_func=lambda x: {"C": "Continent", "UK": "UK", "US": "US"}.get(x, x),
                    key="gr_ukcont",
                )
            with gf3:
                all_origins = sorted(gr["Origin"].dropna().unique())
                sel_origins = st.multiselect(
                    "Origin", all_origins, default=all_origins, key="gr_origins",
                )

            grd = gr[
                (gr["PanelDate"] == sel_panel) &
                (gr["UKContUS"].isin(sel_ukcont)) &
                (gr["Origin"].isin(sel_origins))
            ]
            st.markdown("<hr>", unsafe_allow_html=True)

            # ── KPI strip ────────────────────────────────────────────────────
            total_lots = int(grd["NoLots"].sum())
            cont_lots  = int(grd[grd["UKContUS"] == "C"]["NoLots"].sum())
            uk_lots    = int(grd[grd["UKContUS"] == "UK"]["NoLots"].sum())
            us_lots    = int(grd[grd["UKContUS"] == "US"]["NoLots"].sum())
            n_origins  = grd["Origin"].nunique()
            n_ports    = grd["PortId"].nunique()
            st.markdown(
                kpi("Panel Date", pd.Timestamp(sel_panel).strftime("%d/%m/%Y")) +
                kpi("Total Lots", f"{total_lots:,}") +
                kpi("Continent", f"{cont_lots:,}") +
                kpi("UK", f"{uk_lots:,}") +
                kpi("US", f"{us_lots:,}") +
                kpi("Origins", str(n_origins)) +
                kpi("Ports", str(n_ports)),
                unsafe_allow_html=True)
            st.markdown("<hr>", unsafe_allow_html=True)

            # ── Row 1: Lots by Origin + Lots by Class ────────────────────────
            gc1, gc2 = st.columns(2)
            with gc1:
                st.markdown(lbl("Lots by Origin", RED), unsafe_allow_html=True)
                orig_df = (grd.groupby("Origin")["NoLots"].sum()
                              .sort_values(ascending=False).reset_index())
                fig_go = go.Figure(go.Bar(
                    x=orig_df["Origin"], y=orig_df["NoLots"],
                    marker_color=PIE_COLORS[:len(orig_df)],
                    text=orig_df["NoLots"].map(lambda v: f"{v:,}"),
                    textposition="outside", textfont=dict(size=9)))
                fig_go.update_layout(height=260, margin=dict(t=8,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="Lots"),
                    **_D)
                st.plotly_chart(fig_go, use_container_width=True)

            with gc2:
                st.markdown(lbl("Lots by Class", RED), unsafe_allow_html=True)
                cls_df = (grd.groupby("Class")["NoLots"].sum()
                             .reset_index().sort_values("Class"))
                fig_gc = go.Figure(go.Bar(
                    x=cls_df["Class"].astype(str), y=cls_df["NoLots"],
                    marker_color=PIE_COLORS[:len(cls_df)],
                    text=cls_df["NoLots"].map(lambda v: f"{v:,}"),
                    textposition="outside", textfont=dict(size=9)))
                fig_gc.update_layout(height=260, margin=dict(t=8,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=11),
                               title="Class  (1=par · 2=−30pts · 4=−90pts · P=premium)"),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="Lots"),
                    **_D)
                st.plotly_chart(fig_gc, use_container_width=True)

            st.markdown("<hr>", unsafe_allow_html=True)

            # ── Row 2: Lots by Port + Lots by Allowance ──────────────────────
            gc3, gc4 = st.columns(2)
            with gc3:
                st.markdown(lbl("Lots by Port", RED), unsafe_allow_html=True)
                port_df = (grd.groupby("PortId")["NoLots"].sum()
                              .sort_values(ascending=False).reset_index())
                fig_gp = go.Figure(go.Bar(
                    x=port_df["PortId"], y=port_df["NoLots"],
                    marker_color=PIE_COLORS[:len(port_df)],
                    text=port_df["NoLots"].map(lambda v: f"{v:,}"),
                    textposition="outside", textfont=dict(size=9)))
                fig_gp.update_layout(height=260, margin=dict(t=8,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=11)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="Lots"),
                    **_D)
                st.plotly_chart(fig_gp, use_container_width=True)

            with gc4:
                st.markdown(lbl("Lots by Allowance (pts)", RED), unsafe_allow_html=True)
                alw_df = (grd.groupby("Allowance")["NoLots"].sum()
                             .reset_index().sort_values("Allowance"))
                alw_df["label"] = alw_df["Allowance"].map(lambda v: f"{v:+d} pts")
                fig_ga = go.Figure(go.Bar(
                    x=alw_df["label"], y=alw_df["NoLots"],
                    marker_color=[GREEN if v >= 0 else DRED for v in alw_df["Allowance"]],
                    text=alw_df["NoLots"].map(lambda v: f"{v:,}"),
                    textposition="outside", textfont=dict(size=9)))
                fig_ga.update_layout(height=260, margin=dict(t=8,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=11)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="Lots"),
                    **_D)
                st.plotly_chart(fig_ga, use_container_width=True)

            st.markdown("<hr>", unsafe_allow_html=True)

            # ── Row 3: Origin × Class heatmap ────────────────────────────────
            st.markdown(lbl("Origin × Class — Lots Heatmap", RED), unsafe_allow_html=True)
            hm_piv = grd.pivot_table(
                index="Origin", columns="Class", values="NoLots",
                aggfunc="sum", fill_value=0,
            )
            hm_piv.columns = [str(c) for c in hm_piv.columns]
            fig_hm = go.Figure(go.Heatmap(
                z=hm_piv.values,
                x=list(hm_piv.columns),
                y=list(hm_piv.index),
                colorscale=[[0, "#f0f2f8"], [0.5, "#4a7fb5"], [1, NAVY]],
                text=hm_piv.values,
                texttemplate="%{text:,}",
                textfont=dict(size=11),
                hovertemplate="Origin: %{y}<br>Class: %{x}<br>Lots: %{z:,}<extra></extra>",
                showscale=True,
            ))
            fig_hm.update_layout(
                height=max(200, 50 * len(hm_piv.index)),
                margin=dict(t=8, b=8, l=4, r=4),
                xaxis=dict(tickfont=dict(size=11), title="Class"),
                yaxis=dict(tickfont=dict(size=10)),
                **_D,
            )
            st.plotly_chart(fig_hm, use_container_width=True)

            # ── Row 4: Historical time series (shown when >1 panel date) ─────
            if len(all_panel_dates) > 1:
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown(lbl("Total Lots Over Time (all panel dates)", RED), unsafe_allow_html=True)
                ts_gr = (gr.groupby("PanelDate")["NoLots"].sum()
                           .reset_index().sort_values("PanelDate"))
                fig_ts_gr = go.Figure()
                fig_ts_gr.add_trace(go.Scatter(
                    x=ts_gr["PanelDate"], y=ts_gr["NoLots"],
                    mode="lines+markers",
                    line=dict(color=RED, width=2),
                    marker=dict(size=5, color=RED),
                    fill="tozeroy", fillcolor="rgba(139,26,0,0.07)",
                    hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:,} lots</b><extra></extra>",
                ))
                fig_ts_gr.update_layout(height=240, margin=dict(t=8,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="Lots"),
                    **_D)
                st.plotly_chart(fig_ts_gr, use_container_width=True)

            # ── Raw data table ────────────────────────────────────────────────
            st.markdown("<hr>", unsafe_allow_html=True)
            with st.expander("Raw Data", expanded=False):
                disp_gr = grd.copy()
                disp_gr["PanelDate"] = disp_gr["PanelDate"].dt.strftime("%d/%m/%Y")
                st.dataframe(disp_gr.reset_index(drop=True), use_container_width=True, height=400)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("Certified Stocks Dashboard  ·  ETG Softs  ·  Source: LSEG (Refinitiv)  ·  "
           f"Data as of {pd.Timestamp('today').strftime('%d/%m/%Y')}")
