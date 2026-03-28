"""
Certified Stocks Dashboard — KC Arabica & LRC Robusta
======================================================
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
NAVY = "#0a2463"; RED = "#8b1a00"
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

def lbl(t, color=NAVY):
    return (f"<div style='background:{color};padding:5px 13px;border-radius:5px;"
            f"margin:0 0 5px 0;text-align:center'><span style='font-size:.78rem;"
            f"font-weight:500;letter-spacing:.07em;text-transform:uppercase;"
            f"color:#dde4f0'>{t}</span></div>")

def kpi(label, val, delta=None, delta_color=None):
    d_html = ""
    if delta is not None:
        dc = delta_color or ("#1a7a1a" if delta >= 0 else "#c0392b")
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

# ── Commodity selector ────────────────────────────────────────────────────────
comm = st.radio("", ["KC — Arabica (ICE NYSE)", "LRC — Robusta (ICE Liffe)"],
                horizontal=True, label_visibility="collapsed")
st.markdown("<hr>", unsafe_allow_html=True)

# =============================================================================
# KC VIEW
# =============================================================================
if comm.startswith("KC"):
    kc = load_kc()

    # Date slider
    min_d = kc["Date"].min().date()
    max_d = kc["Date"].max().date()
    date_range = st.slider("Date range", min_value=min_d, max_value=max_d,
                           value=(pd.Timestamp("2024-01-01").date(), max_d),
                           format="YYYY-MM-DD")
    dff = kc[(kc["Date"] >= pd.Timestamp(date_range[0])) &
             (kc["Date"] <= pd.Timestamp(date_range[1]))]
    st.markdown("<hr>", unsafe_allow_html=True)

    # KPI row
    tot_now  = int(kc["KC-TOT-TOT"].dropna().iloc[-1])
    tot_prev = int(kc["KC-TOT-TOT"].dropna().iloc[-2])
    tot_chg  = tot_now - tot_prev
    hon_now  = int(kc["KC-HON-TOT"].dropna().iloc[-1])
    hon_pct  = hon_now / tot_now * 100
    pend_now = int(kc["KC-TOT-PENDING"].dropna().iloc[-1]) if "KC-TOT-PENDING" in kc.columns else 0
    latest_d = kc["Date"].max().strftime("%d/%m/%Y")
    st.markdown(
        kpi("Latest Date", latest_d) +
        kpi("KC Total Certs", f"{tot_now/1000:.1f}k bags", tot_chg) +
        kpi("vs Prior Day", f"{tot_chg:+,} bags") +
        kpi("Honduras", f"{hon_now/1000:.1f}k", hon_now - int(kc["KC-HON-TOT"].dropna().iloc[-2])) +
        kpi("Honduras %", f"{hon_pct:.1f}%") +
        kpi("Total Pending", f"{pend_now/1000:.1f}k bags"),
        unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 1: Total certs + Origin share ────────────────────────────────────
    c1, c2 = st.columns([3, 1])
    with c1:
        st.markdown(lbl("KC Total Certified Stocks (bags)"), unsafe_allow_html=True)
        fig = go.Figure()
        valid = dff[["Date","KC-TOT-TOT"]].dropna()
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
        origins_excl_tot = [o for o in ORIGIN_NAMES if o != "TOT"]
        shares = {}
        for o in origins_excl_tot:
            col = f"KC-{o}-TOT"
            if col in kc.columns:
                v = kc[col].dropna()
                if len(v) > 0 and float(v.iloc[-1]) > 0:
                    shares[ORIGIN_NAMES[o]] = float(v.iloc[-1])
        if shares:
            colors = px.colors.qualitative.Set3[:len(shares)]
            fig_d = go.Figure(go.Pie(
                labels=list(shares.keys()), values=list(shares.values()),
                hole=0.5, textinfo="label+percent", textfont=dict(size=7.5),
                marker=dict(colors=colors), sort=True))
            fig_d.update_layout(height=240, margin=dict(t=8,b=8,l=0,r=0), showlegend=False, **_D)
            st.plotly_chart(fig_d, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 2: Origin×Port Heatmap ───────────────────────────────────────────
    st.markdown(lbl("Origin × Port — Day-over-Day Change (bags)"), unsafe_allow_html=True)

    origins_show = [o for o in ORIGIN_NAMES if o != "TOT"]
    ports_show   = [p for p in PORT_NAMES if p != "TOT"]
    ports_plus   = ports_show + ["TOT"]

    rows = []
    kc_sorted = kc.dropna(subset=["KC-TOT-TOT"]).sort_values("Date")
    latest_row = kc_sorted.iloc[-1]
    prev_row   = kc_sorted.iloc[-2]

    for o in origins_show:
        row = {"Origin": ORIGIN_NAMES[o]}
        for p in ports_plus:
            col = f"KC-{o}-{p}"
            if col in kc.columns:
                now_  = latest_row.get(col, np.nan)
                prev_ = prev_row.get(col, np.nan)
                row[PORT_NAMES[p]] = int(now_ - prev_) if pd.notna(now_) and pd.notna(prev_) else 0
            else:
                row[PORT_NAMES[p]] = 0
        rows.append(row)

    chg_df = pd.DataFrame(rows).set_index("Origin")

    def color_val(v):
        if v > 0:   return f"background:rgba(26,122,26,{min(abs(v)/4000,0.65):.2f});color:#0d4d0d"
        elif v < 0: return f"background:rgba(192,57,43,{min(abs(v)/4000,0.65):.2f});color:#5a0d0d"
        return "color:#bbb"

    styled = chg_df.style.applymap(color_val).format(lambda v: f"{v:+,}" if v != 0 else "—")
    st.dataframe(styled, use_container_width=True, height=460)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 3: Honduras + Arbitrage ──────────────────────────────────────────
    c3, c4 = st.columns(2)

    with c3:
        st.markdown(lbl("Honduras Certified Stocks — Level + Rolling 20d Flow"), unsafe_allow_html=True)
        hon_data = dff[["Date","KC-HON-TOT","KC-TOT-TOT"]].dropna(subset=["KC-HON-TOT"])
        hon_data = hon_data.copy()
        hon_data["Chg"]   = hon_data["KC-HON-TOT"].diff()
        hon_data["Roll20"] = hon_data["Chg"].rolling(20).sum() / 1000
        hon_data["Pct"]   = hon_data["KC-HON-TOT"] / hon_data["KC-TOT-TOT"] * 100

        # z-score vs full history
        full_hon = kc["KC-HON-TOT"].dropna()
        z_now = (float(full_hon.iloc[-1]) - float(full_hon.mean())) / float(full_hon.std())

        z_color = "#c0392b" if abs(z_now) > 1.5 else NAVY
        st.markdown(f"<div style='font-size:.72rem;color:#6e6e73;margin-bottom:4px'>"
                    f"Latest: <b>{hon_now/1000:.1f}k bags</b> &nbsp;·&nbsp; "
                    f"% of Total: <b>{hon_pct:.1f}%</b> &nbsp;·&nbsp; "
                    f"Z-Score (full hist): <b style='color:{z_color}'>{z_now:+.2f}</b></div>",
                    unsafe_allow_html=True)

        fig_h = make_subplots(specs=[[{"secondary_y": True}]])
        fig_h.add_trace(go.Scatter(
            x=hon_data["Date"], y=hon_data["KC-HON-TOT"]/1000,
            name="Honduras (k bags)", mode="lines",
            line=dict(color=NAVY, width=2),
            fill="tozeroy", fillcolor="rgba(10,36,99,0.07)"), secondary_y=False)
        fig_h.add_trace(go.Bar(
            x=hon_data["Date"], y=hon_data["Roll20"],
            name="Rolling 20d Flow (k)", marker_color="rgba(139,26,0,0.3)"), secondary_y=True)
        fig_h.update_layout(height=280, legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
            margin=dict(t=10,b=8,l=4,r=4), **_D)
        fig_h.update_yaxes(title_text="k bags", secondary_y=False,
            showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9))
        fig_h.update_yaxes(title_text="20d Flow (k)", secondary_y=True,
            showgrid=False, tickfont=dict(size=9))
        fig_h.update_xaxes(showgrid=False, tickfont=dict(size=9))
        st.plotly_chart(fig_h, use_container_width=True)

    with c4:
        st.markdown(lbl("Origin Differentials vs KC Futures (cts/lb)"), unsafe_allow_html=True)
        arb_cols = [c for c in kc.columns if c not in
                    [x for x in kc.columns if x.startswith("KC-")] and
                    c not in ["Date","KC_Price"]]

        if arb_cols:
            arb_latest = {}
            arb_prev   = {}
            arb_sorted = kc[["Date"] + arb_cols].dropna(subset=[arb_cols[0]]).sort_values("Date")
            if len(arb_sorted) >= 2:
                for c in arb_cols:
                    s = kc[["Date",c]].dropna(subset=[c]).sort_values("Date")
                    if len(s) >= 2:
                        arb_latest[c] = float(s[c].iloc[-1])
                        arb_prev[c]   = float(s[c].iloc[-2])

            if arb_latest:
                names = list(arb_latest.keys())
                vals  = list(arb_latest.values())
                prev_ = [arb_prev.get(n, v) for n, v in zip(names, vals)]
                chgs  = [v - p for v, p in zip(vals, prev_)]
                colors_arb = ["#1a7a1a" if v >= 0 else "#c0392b" for v in vals]

                # Abbreviate names
                abbrev = {"Brazil NY 3/4":"BRZ NY","Brazil Santos":"BRZ STS",
                          "Brazil Rio 15/16":"BRZ R1","Brazil Rio 17/18":"BRZ R2",
                          "Ethiopia Djimmah":"ETH","Uganda Drugar":"UGA",
                          "India Cherry":"IND","Colombia Excelso":"COL",
                          "Honduras HG":"HON","Peru MCM":"PER","Guatemala SHB":"GUA"}
                short_names = [abbrev.get(n, n[:8]) for n in names]

                fig_arb = go.Figure(go.Bar(
                    x=short_names, y=vals,
                    marker_color=colors_arb,
                    text=[f"{v:+.0f}" for v in vals],
                    textposition="outside", textfont=dict(size=8.5),
                    customdata=[[chgs[i]] for i in range(len(vals))],
                    hovertemplate="%{x}<br><b>%{y:+.0f} cts/lb</b><br>WoW chg: %{customdata[0]:+.0f}<extra></extra>",
                ))
                fig_arb.add_hline(y=0, line_color="#cccccc", line_width=1)
                fig_arb.update_layout(height=280, margin=dict(t=10,b=8,l=4,r=4),
                    xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                    yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9),
                               title="cts/lb", zeroline=False), **_D)
                st.plotly_chart(fig_arb, use_container_width=True)

    # Arbitrage time series (collapsible)
    with st.expander("Origin Differentials — Full Time Series", expanded=False):
        if arb_cols:
            arb_ts = dff[["Date"] + [c for c in arb_cols if c in dff.columns]].copy()
            fig_ats = go.Figure()
            palette = px.colors.qualitative.Set1
            for i, c in enumerate([x for x in arb_cols if x in arb_ts.columns]):
                s = arb_ts[["Date",c]].dropna(subset=[c])
                if len(s) > 0:
                    fig_ats.add_trace(go.Scatter(x=s["Date"], y=s[c],
                        name=c, mode="lines", line=dict(color=palette[i % len(palette)], width=1.5)))
            fig_ats.add_hline(y=0, line_color="#cccccc", line_width=1)
            fig_ats.update_layout(height=280, legend=dict(orientation="h", y=1.02, x=0, font=dict(size=7.5)),
                margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="cts/lb"), **_D)
            st.plotly_chart(fig_ats, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 4: Grading Queue + Rolling Cert Flow ──────────────────────────────
    c5, c6 = st.columns(2)

    with c5:
        st.markdown(lbl("Grading Queue by Port — Passed / Failed"), unsafe_allow_html=True)
        grade_ports = ["AN","HA","HO","MI","NO","NY"]
        pass_vals, fail_vals, port_labels = [], [], []
        for p in grade_ports:
            pc = f"KC-{p}-PASSGRAD"
            fc = f"KC-{p}-FAILGRAD"
            pv = int(kc[pc].dropna().iloc[-1]) if pc in kc.columns and len(kc[pc].dropna()) > 0 else 0
            fv = int(kc[fc].dropna().iloc[-1]) if fc in kc.columns and len(kc[fc].dropna()) > 0 else 0
            if pv > 0 or fv > 0:
                pass_vals.append(pv)
                fail_vals.append(fv)
                port_labels.append(PORT_NAMES.get(p, p))

        if port_labels:
            fig_g = go.Figure()
            fig_g.add_trace(go.Bar(x=port_labels, y=pass_vals, name="Passed",
                marker_color="#1a7a1a",
                text=[f"{v:,}" for v in pass_vals], textposition="outside", textfont=dict(size=8)))
            fig_g.add_trace(go.Bar(x=port_labels, y=fail_vals, name="Failed",
                marker_color="#c0392b",
                text=[f"{v:,}" for v in fail_vals], textposition="outside", textfont=dict(size=8)))
            pending_val = int(kc["KC-TOT-PENDING"].dropna().iloc[-1]) if "KC-TOT-PENDING" in kc.columns else 0
            fig_g.update_layout(height=280, barmode="group",
                legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
                margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"), **_D)
            st.plotly_chart(fig_g, use_container_width=True)
            st.caption(f"Total Pending: {pending_val:,} bags  ·  "
                       f"Total Passed: {sum(pass_vals):,}  ·  Total Failed: {sum(fail_vals):,}")

    with c6:
        st.markdown(lbl("Pending Certification — Time Series"), unsafe_allow_html=True)
        if "KC-TOT-PENDING" in kc.columns:
            pend_ts = dff[["Date","KC-TOT-PENDING"]].dropna(subset=["KC-TOT-PENDING"])
            fig_pend = go.Figure()
            fig_pend.add_trace(go.Scatter(
                x=pend_ts["Date"], y=pend_ts["KC-TOT-PENDING"] / 1000,
                mode="lines", line=dict(color="#e8a020", width=2),
                fill="tozeroy", fillcolor="rgba(232,160,32,0.1)",
                hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:.1f}k bags</b><extra></extra>"))
            fig_pend.update_layout(height=280, margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=9)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="k bags"), **_D)
            st.plotly_chart(fig_pend, use_container_width=True)

# =============================================================================
# LRC VIEW
# =============================================================================
else:
    lrc = load_lrc()

    min_d = lrc["Date"].min().date()
    max_d = lrc["Date"].max().date()
    date_range = st.slider("Date range", min_value=min_d, max_value=max_d,
                           value=(pd.Timestamp("2020-01-01").date(), max_d),
                           format="YYYY-MM-DD")
    dff = lrc[(lrc["Date"] >= pd.Timestamp(date_range[0])) &
              (lrc["Date"] <= pd.Timestamp(date_range[1]))]
    st.markdown("<hr>", unsafe_allow_html=True)

    # KPI row
    tot_now  = int(lrc["LRC-TOT-VG"].dropna().iloc[-1])
    tot_prev = int(lrc["LRC-TOT-VG"].dropna().iloc[-2])
    tot_chg  = tot_now - tot_prev
    latest_d = lrc["Date"].max().strftime("%d/%m/%Y")

    lrc_port_cols = {"ANT":"LRC-ANT-VG","LON":"LRC-LON-VG","FEL":"LRC-FEL-VG",
                     "BAR":"LRC-BAR-VG","AMS":"LRC-AMS-VG"}
    port_vals_now = {}
    for pname, col in lrc_port_cols.items():
        if col in lrc.columns:
            s = lrc[col].dropna()
            if len(s): port_vals_now[pname] = int(s.iloc[-1])

    ant_pct = port_vals_now.get("ANT",0) / tot_now * 100 if tot_now else 0
    lon_pct = port_vals_now.get("LON",0) / tot_now * 100 if tot_now else 0

    st.markdown(
        kpi("Latest Date", latest_d) +
        kpi("LRC Total Certs", f"{tot_now:,} bags", tot_chg) +
        kpi("vs Prior Day", f"{tot_chg:+,} bags") +
        kpi("Antwerp", f"{port_vals_now.get('ANT',0):,}", None) +
        kpi("ANT %", f"{ant_pct:.1f}%") +
        kpi("London", f"{port_vals_now.get('LON',0):,}", None) +
        kpi("LON %", f"{lon_pct:.1f}%"),
        unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 1: Total + Port Share ─────────────────────────────────────────────
    c1, c2 = st.columns([3, 1])

    with c1:
        st.markdown(lbl("LRC Total Certified Stocks — Full History (bags)", RED), unsafe_allow_html=True)
        fig_lt = go.Figure()
        valid_lt = dff[["Date","LRC-TOT-VG"]].dropna()
        fig_lt.add_trace(go.Scatter(x=valid_lt["Date"], y=valid_lt["LRC-TOT-VG"],
            mode="lines", line=dict(color=RED, width=2),
            fill="tozeroy", fillcolor="rgba(139,26,0,0.07)",
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{y:,} bags</b><extra></extra>"))
        fig_lt.update_layout(height=240, margin=dict(t=8,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=9)),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), tickformat=","), **_D)
        st.plotly_chart(fig_lt, use_container_width=True)

    with c2:
        st.markdown(lbl("Port Share (%)", RED), unsafe_allow_html=True)
        active_ports = {k: v for k, v in port_vals_now.items() if v > 0}
        if active_ports:
            fig_pd = go.Figure(go.Pie(
                labels=list(active_ports.keys()), values=list(active_ports.values()),
                hole=0.5, textinfo="label+percent", textfont=dict(size=9),
                marker=dict(colors=[NAVY,"#c0553a","#4a7fb5","#e8c96a","#9bbcd6"])))
            fig_pd.update_layout(height=240, margin=dict(t=8,b=8,l=0,r=0), showlegend=False, **_D)
            st.plotly_chart(fig_pd, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 2: Port Stocks + Daily Change ────────────────────────────────────
    c3, c4 = st.columns(2)

    with c3:
        st.markdown(lbl("Current Stocks by Port (bags)", RED), unsafe_allow_html=True)
        all_port_rics = {p: f"LRC-{p}-VG" for p in
                        ["ANT","LON","FEL","BAR","AMS","BRE","HAM","LIV","NOR","ROT","TRI"]}
        pb_names, pb_vals = [], []
        for pname, col in all_port_rics.items():
            if col in lrc.columns:
                s = lrc[col].dropna()
                v = int(s.iloc[-1]) if len(s) else 0
                if v > 0:
                    pb_names.append(pname)
                    pb_vals.append(v)

        if pb_names:
            clrs = [NAVY, "#c0553a", "#4a7fb5", "#e8c96a"] + ["#aaa"] * 10
            fig_pb = go.Figure(go.Bar(
                x=pb_names, y=pb_vals,
                marker_color=clrs[:len(pb_names)],
                text=[f"{v:,}" for v in pb_vals],
                textposition="outside", textfont=dict(size=9)))
            fig_pb.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=10)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"), **_D)
            st.plotly_chart(fig_pb, use_container_width=True)

    with c4:
        st.markdown(lbl("Day-over-Day Change by Port (bags)", RED), unsafe_allow_html=True)
        chg_names, chg_vals = [], []
        lrc_sorted = lrc.dropna(subset=["LRC-TOT-VG"]).sort_values("Date")
        if len(lrc_sorted) >= 2:
            lat = lrc_sorted.iloc[-1]; prv = lrc_sorted.iloc[-2]
            for pname, col in all_port_rics.items():
                if col in lrc.columns:
                    nv = lat.get(col, np.nan); pv = prv.get(col, np.nan)
                    if pd.notna(nv) and pd.notna(pv):
                        chg = int(nv - pv)
                        if chg != 0:
                            chg_names.append(pname); chg_vals.append(chg)

        if chg_names:
            cc = ["#1a7a1a" if v >= 0 else "#c0392b" for v in chg_vals]
            fig_pc = go.Figure(go.Bar(
                x=chg_names, y=chg_vals, marker_color=cc,
                text=[f"{v:+,}" for v in chg_vals],
                textposition="outside", textfont=dict(size=9)))
            fig_pc.add_hline(y=0, line_color="#cccccc", line_width=1)
            fig_pc.update_layout(height=260, margin=dict(t=10,b=8,l=4,r=4),
                xaxis=dict(showgrid=False, tickfont=dict(size=10)),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags",
                           zeroline=False), **_D)
            st.plotly_chart(fig_pc, use_container_width=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Row 3: Port time series + Rolling 20d ────────────────────────────────
    c5, c6 = st.columns(2)

    with c5:
        st.markdown(lbl("Port Stocks Over Time (bags)", RED), unsafe_allow_html=True)
        fig_ts = go.Figure()
        port_colors = {"ANT": NAVY, "LON": "#c0553a", "FEL": "#4a7fb5", "BAR": "#e8c96a"}
        for pname, col in {"ANT":"LRC-ANT-VG","LON":"LRC-LON-VG",
                            "FEL":"LRC-FEL-VG","BAR":"LRC-BAR-VG"}.items():
            if col in dff.columns:
                s = dff[["Date",col]].dropna(subset=[col])
                if len(s):
                    fig_ts.add_trace(go.Scatter(x=s["Date"], y=s[col],
                        name=pname, mode="lines",
                        line=dict(color=port_colors.get(pname,"#aaa"), width=1.8)))
        fig_ts.update_layout(height=260,
            legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
            margin=dict(t=10,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=9)),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"), **_D)
        st.plotly_chart(fig_ts, use_container_width=True)

    with c6:
        st.markdown(lbl("Rolling 20-Day Change — Total / ANT / UK (Lon+Fel)", RED), unsafe_allow_html=True)
        roll_df = dff[["Date","LRC-TOT-VG","LRC-ANT-VG","LRC-LON-VG","LRC-FEL-VG"]].copy()
        roll_df = roll_df.sort_values("Date")
        roll_df["20d TOT"] = roll_df["LRC-TOT-VG"].diff().rolling(20).sum()
        roll_df["20d ANT"] = roll_df["LRC-ANT-VG"].diff().rolling(20).sum()
        roll_df["20d UK"]  = (roll_df["LRC-LON-VG"].fillna(0) +
                               roll_df["LRC-FEL-VG"].fillna(0)).diff().rolling(20).sum()

        fig_roll = go.Figure()
        for col, color, label in [("20d TOT", RED, "20d Total"),
                                    ("20d ANT", NAVY, "20d Antwerp"),
                                    ("20d UK", "#4a7fb5", "20d UK (Lon+Fel)")]:
            s = roll_df[["Date",col]].dropna(subset=[col])
            fig_roll.add_trace(go.Scatter(x=s["Date"], y=s[col],
                name=label, mode="lines", line=dict(color=color, width=1.8)))
        fig_roll.add_hline(y=0, line_color="#cccccc", line_width=1)
        fig_roll.update_layout(height=260,
            legend=dict(orientation="h", y=1.02, x=0, font=dict(size=8)),
            margin=dict(t=10,b=8,l=4,r=4),
            xaxis=dict(showgrid=False, tickfont=dict(size=9)),
            yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=9), title="bags"), **_D)
        st.plotly_chart(fig_roll, use_container_width=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("Certified Stocks Dashboard  ·  ETG Softs  ·  Source: LSEG (Refinitiv)  ·  "
           f"Data as of {pd.Timestamp('today').strftime('%d/%m/%Y')}")
