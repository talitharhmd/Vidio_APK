"""
Vidio Plays Dashboard (Streamlit).
Semua angka dihitung dari satu tabel: fact_dashboard
(sheet "fact_dashboard" di Vidio_Dashboard_Data.xlsx, atau fact_dashboard.csv; hasil bagian 10b notebook).
"""
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from PIL import Image

HERE = Path(__file__).parent
EXCEL_PATH = HERE / "Vidio_Dashboard_Data.xlsx"
SHEET_NAME = "fact_dashboard"
CSV_PATH = HERE / "fact_dashboard.csv"
AUTHOR_LINE = "Created by Talitha RW | 25/11/2026"

RED, INK, SLATE, GOLD, SKY, PINK, SAND, GREY = "#FC0F3B", "#14161C", "#3D5A80", "#E0A100", "#7FAECB", "#FF8DA0", "#F0C14B", "#9AA0B0"
PLATFORM_COLOR = {"app-android": RED, "app-ios": PINK, "web-mobile": SLATE, "web-desktop": SKY, "tv-android": GOLD, "tv-tizen": SAND}
CT_LABEL = {"vod": "VOD", "livestreaming": "Live TV", "catchup": "Catch-up", "unknown": "Unknown"}
CT_COLOR = {"VOD": SLATE, "Live TV": RED, "Catch-up": GOLD, "Unknown": GREY}
DAYS = {1: "Sen", 2: "Sel", 3: "Rab", 4: "Kam", 5: "Jum", 6: "Sab", 7: "Min"}
DAYPARTS = ["Dini hari 00–04", "Pagi 05–10", "Siang 11–14", "Sore 15–18", "Malam 19–23"]
THRESHOLDS = [("1", "≥ 1 dtk"), ("10", "≥ 10 dtk"), ("30", "≥ 30 dtk"), ("60", "≥ 60 dtk"), ("180", "≥ 3 mnt"), ("600", "≥ 10 mnt")]
REQUIRED = {"date", "hour", "platform", "family", "content_type", "status_user", "sumber", "plays", "zero_plays", "play_ge_1s", "play_ge_10s",
            "play_ge_30s", "play_ge_60s", "play_ge_180s", "play_ge_600s", "watch_seconds", "watch_seconds_capped", "buffer_plays", "weekday"}

st.set_page_config(page_title="Vidio Plays Dashboard", page_icon=Image.open(HERE / "vidio_icon.png"), layout="wide")

st.markdown(f"""
<style>
.block-container {{padding-top: 1.1rem; padding-bottom: 1rem; max-width: 1500px;}}
header[data-testid="stHeader"] {{background: transparent;}}
[class*="st-key-card"] {{background: #FFFFFF; border-radius: 12px;}}
.ct {{font-weight: 700; font-size: 1.02rem; color: {INK}; margin: 0 0 2px 0;}}
.cs {{color: #6B7280; font-size: .8rem; margin-bottom: 6px;}}
.banner {{background: linear-gradient(120deg, {RED}, #A01025); color: #fff; border-radius: 12px; padding: 18px 24px; height: 100%;}}
.banner .k {{font-size: .72rem; letter-spacing: .08em; text-transform: uppercase; font-weight: 600; opacity: .92;}}
.banner .h {{font-size: 1.55rem; font-weight: 700; line-height: 1.2; margin: 6px 0;}}
.banner .s {{font-size: .95rem; opacity: .95;}}
.kpis {{display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px;}}
.kpi {{background: #fff; border: 1px solid #E3E6EE; border-radius: 10px; padding: 9px 13px;}}
.kpi .l {{color: #6B7280; font-size: .76rem;}}
.kpi .v {{font-size: 1.55rem; font-weight: 700; color: {INK}; line-height: 1.25;}}
.strip {{display: flex; gap: 3px; flex-wrap: wrap; margin: 4px 0 10px 0;}}
.strip span {{width: 17px; height: 17px; border-radius: 4px; display: inline-block;}}
table.t {{width: 100%; border-collapse: collapse; font-size: .84rem;}}
table.t th {{text-align: left; color: #6B7280; font-weight: 600; font-size: .72rem; letter-spacing: .04em; border-bottom: 1px solid #E3E6EE; padding: 4px 6px;}}
table.t td {{padding: 5px 6px; border-bottom: 1px solid #EEF0F5;}}
.badge {{display: inline-block; font-size: .74rem; font-weight: 600; padding: 3px 10px; border-radius: 999px; margin-right: 6px;}}
.good {{background: #E4F6EA; color: #0F7B3A;}} .warn {{background: #FFE8EC; color: #B3102B;}}
.note {{font-size: .8rem; color: #4B5563;}}
.take {{background: #F4F5F8; border-radius: 10px; padding: 10px 12px; margin-bottom: 8px;}}
.take b {{font-size: .92rem;}} .take div {{font-size: .82rem; color: #4B5563;}}
.tbl-wrap {{overflow-x: auto;}}
@media (max-width: 1150px) {{
  div[data-testid="stHorizontalBlock"] {{flex-wrap: wrap !important; row-gap: .6rem;}}
  div[data-testid="stColumn"] {{flex: 1 1 calc(50% - 1rem) !important; min-width: calc(50% - 1rem) !important; width: auto !important;}}
}}
@media (max-width: 700px) {{
  div[data-testid="stColumn"] {{flex: 1 1 100% !important; min-width: 100% !important;}}
  .banner .h {{font-size: 1.25rem;}}
  .kpi .v {{font-size: 1.3rem;}}
}}
.footer {{text-align: center; color: #6B7280; font-size: .82rem; padding: 14px 0 4px 0;}}
</style>
""", unsafe_allow_html=True)


# ---------- helpers ----------
def idn(n, d=0):
    return f"{n:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x):
    x = 0 if pd.isna(x) else x
    return (idn(x * 100, 0) if abs(x) >= .1 else idn(x * 100, 1)) + "%"


def compact(n):
    n = float(n)
    return f"{n / 1e6:.1f}M" if n >= 1e6 else f"{n / 1e3:,.0f}K" if n >= 1e3 else f"{n:,.0f}"


def mmss(sec):
    sec = 0 if pd.isna(sec) else int(round(sec))
    return f"{sec // 60}:{sec % 60:02d}"


def pcolor(p):
    return PLATFORM_COLOR.get(p, GREY)


def fcolor(f):
    f = str(f).lower()
    return RED if "app" in f else SLATE if "web" in f else GOLD if "tv" in f else GREY


def show(fig, h=300):
    fig.update_layout(height=h, margin=dict(l=4, r=12, t=34, b=4), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Inter, Helvetica, Arial, sans-serif", size=12, color=INK),
                      legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, font=dict(size=11)))
    fig.update_yaxes(automargin=True)
    cfg = {"displayModeBar": False}
    try:
        st.plotly_chart(fig, width="stretch", config=cfg)
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config=cfg)


_n = [0]


def card(title, sub=None):
    _n[0] += 1
    c = st.container(border=True, key=f"card{_n[0]}")
    c.markdown(f"<div class='ct'>{title}</div>" + (f"<div class='cs'>{sub}</div>" if sub else ""), unsafe_allow_html=True)
    return c


def agg(g, by):
    a = g.groupby(by, observed=True).agg(plays=("plays", "sum"), zero=("zero_plays", "sum"), eng=("play_ge_30s", "sum"),
                                         ws=("watch_seconds", "sum"), wsc=("watch_seconds_capped", "sum"), buf=("buffer_plays", "sum"))
    a["zero_rate"], a["eng_rate"] = a.zero / a.plays, a.eng / a.plays
    a["avg_sec"], a["buf_rate"], a["hours"] = a.wsc / a.plays, a.buf / a.plays, a.ws / 3600
    return a


def pick(label, options, key):
    try:
        sel = st.pills(label, options, selection_mode="multi", default=options, key=key)
    except AttributeError:
        sel = st.multiselect(label, options, default=options, key=key)
    return list(sel or [])


# ---------- data ----------
@st.cache_data(show_spinner=False)
def load(src):
    name = getattr(src, "name", str(src)).lower()
    df = pd.read_csv(src) if name.endswith(".csv") else pd.read_excel(src, sheet_name=SHEET_NAME)

    miss = REQUIRED - set(df.columns)
    if miss:
        raise ValueError(f"Kolom ini tidak ada di file data: {', '.join(sorted(miss))}")

    df["date"] = pd.to_datetime(df["date"])
    df["tipe"] = df["content_type"].map(CT_LABEL).fillna(df["content_type"].astype(str).str.title())
    df["status_user"] = df["status_user"].astype(str)

    return df


if EXCEL_PATH.exists():
    data = load(EXCEL_PATH)
elif CSV_PATH.exists():
    data = load(CSV_PATH)
else:
    st.warning("Vidio_Dashboard_Data.xlsx atau fact_dashboard.csv belum ada di folder app.")
    up = st.file_uploader(
        "Upload Vidio_Dashboard_Data.xlsx (sheet fact_dashboard) atau fact_dashboard.csv",
        type=["xlsx", "csv"]
    )

    if up is None:
        st.stop()

    data = load(up)

PLATFORMS = sorted(data["platform"].unique())
TIPES = [t for t in ["VOD", "Live TV", "Catch-up", "Unknown"] if t in set(data["tipe"])] + sorted(set(data["tipe"]) - set(CT_LABEL.values()))
USERS = sorted(data["status_user"].unique(), key=lambda s: ["login", "guest"].index(s) if s in ("login", "guest") else 9)
DMIN, DMAX = data["date"].min().date(), data["date"].max().date()
_dt = data.groupby("date")["plays"].sum()
PARTIAL = [d for d in (_dt.index.min(), _dt.index.max()) if _dt[d] < .6 * _dt.median()]   # hari pertama/terakhir yang terpotong zona waktu


def reset():
    st.session_state["f_date"] = (DMIN, DMAX)
    for k, opts in (("f_plat", PLATFORMS), ("f_tipe", TIPES), ("f_user", USERS)):
        st.session_state[k] = list(opts)


# ---------- header + filters ----------
import base64
_logo = base64.b64encode((HERE / "vidio_logo.png").read_bytes()).decode()
st.markdown(f"<div style='display:flex;align-items:center;gap:14px;margin-bottom:8px'><img src='data:image/png;base64,{_logo}' style='height:42px'>"
            "<span style='font-size:1.35rem;font-weight:700'>Plays Dashboard</span></div>", unsafe_allow_html=True)

with st.container(border=True):
    c1, c2, c3, c4 = st.columns([1.5, 1.5, 1.2, .7], vertical_alignment="bottom")
    dr = c1.date_input("Date range", (DMIN, DMAX), min_value=DMIN, max_value=DMAX, key="f_date")
    with c2:
        tipe_sel = pick("Content type", TIPES, "f_tipe")
    with c3:
        user_sel = pick("User", USERS, "f_user")
    c4.button("Reset", on_click=reset)
    plat_sel = pick("Platform", PLATFORMS, "f_plat")

if not isinstance(dr, (tuple, list)) or len(dr) < 2:
    st.info("Pilih tanggal akhir dulu.")
    st.stop()
f = data[(data["date"].dt.date >= dr[0]) & (data["date"].dt.date <= dr[1]) & data["platform"].isin(plat_sel)
         & data["tipe"].isin(tipe_sel) & data["status_user"].isin(user_sel)]
if f.empty or f["plays"].sum() == 0:
    st.warning("Tidak ada data untuk kombinasi filter ini.")
    st.stop()

tot = f[["plays", "zero_plays", "play_ge_30s", "watch_seconds", "watch_seconds_capped", "buffer_plays"]].sum()
d0, d1 = f["date"].min(), f["date"].max()
rng = f"{d0:%d %b} – {d1:%d %b %Y}"
st.markdown(f"<div class='note' style='text-align:right'>{idn(tot.plays)} plays match the filter</div>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["1 · Engagement", "2 · Waktu & Konten"])

# =====================================================================================
# PAGE 1
# =====================================================================================
with tab1:
    p = agg(f, "platform")
    lead_plays, lead_eng, lead_ws = p.plays.idxmax(), p.eng.idxmax(), p.ws.idxmax()
    share = lambda col, who: p.loc[who, col] / p[col].sum()
    if len({lead_plays, lead_eng, lead_ws}) == 1:
        headline = f"{lead_eng} memimpin plays, engaged plays, dan watch time"
    else:
        headline = "Plays terbanyak belum tentu watch time terbanyak"
    sub = (f"{lead_plays} paling banyak plays ({pct(share('plays', lead_plays))}), {lead_eng} paling banyak engaged plays "
           f"({pct(share('eng', lead_eng))}), {lead_ws} paling banyak watch time ({pct(share('ws', lead_ws))}).")
    b, k = st.columns([1.15, 1], gap="small")
    b.markdown(f"<div class='banner' style='min-height:152px'><div class='k'>Overview · {rng}</div><div class='h'>{headline}</div><div class='s'>{sub}</div></div>", unsafe_allow_html=True)
    kp = [("Total plays", compact(tot.plays)), ("Watch time (jam)", compact(tot.watch_seconds / 3600)),
          ("Engaged rate (≥30 dtk)", pct(tot.play_ge_30s / tot.plays)), ("Zero-second rate", pct(tot.zero_plays / tot.plays)),
          ("Avg watch / play (menit:detik)", mmss(tot.watch_seconds_capped / tot.plays)), ("Buffering rate", pct(tot.buffer_plays / tot.plays))]
    k.markdown("<div class='kpis'>" + "".join(f"<div class='kpi'><div class='l'>{l}</div><div class='v'>{v}</div></div>" for l, v in kp) + "</div>", unsafe_allow_html=True)

    st.write("")
    a, b2, c, d = st.columns([1.25, 1.05, 1.0, 1.15], gap="small")

    with a:
        with card("Plays vs Engaged vs Watch Time", "Share per platform. Kalau urutannya beda, plays saja tidak cukup untuk menilai."):
            q = p.assign(s_plays=p.plays / p.plays.sum(), s_eng=p.eng / p.eng.sum(), s_ws=p.ws / p.ws.sum()).sort_values("s_eng")
            fig = go.Figure()
            for col, name, colr in [("s_ws", "Watch time", INK), ("s_eng", "Engaged ≥30s", RED), ("s_plays", "Plays", GREY)]:
                fig.add_bar(y=q.index, x=q[col], name=name, orientation="h", marker_color=colr, text=[pct(v) for v in q[col]],
                            textposition="outside", cliponaxis=False)
            fig.update_layout(barmode="group", bargap=.35, xaxis=dict(visible=False, range=[0, q[["s_plays", "s_eng", "s_ws"]].max().max() * 1.18]),
                              legend=dict(traceorder="reversed"))
            show(fig, 560)

    with b2:
        with card("Plays yang Langsung Berhenti", "% plays per platform dengan durasi 0 detik"):
            z = p.sort_values("zero_rate")
            fig = go.Figure(go.Bar(y=z.index, x=z.zero_rate, orientation="h", marker_color=[pcolor(i) for i in z.index],
                                   text=[pct(v) for v in z.zero_rate], textposition="outside", cliponaxis=False))
            fig.update_layout(xaxis=dict(visible=False, range=[0, max(z.zero_rate.max() * 1.25, .01)]))
            show(fig, 230)
            e = f.groupby("sumber").agg(plays=("plays", "sum"), zero=("zero_plays", "sum"))
            if {"embed", "non-embed"} <= set(e.index):
                e["share"], e["zr"] = e.plays / e.plays.sum(), e.zero / e.plays
                st.markdown(f"**Embed: {pct(e.loc['embed', 'zr'])} langsung berhenti, vs {pct(e.loc['non-embed', 'zr'])} di non-embed**")
                fig = go.Figure()
                for col, name, colr in [("zr", "% berhenti di 0 dtk", GREY), ("share", "% dari semua plays", SLATE)]:
                    fig.add_bar(y=["Non-embed", "Embed"], x=[e.loc["non-embed", col], e.loc["embed", col]], name=name, orientation="h",
                                marker_color=colr, text=[pct(e.loc["non-embed", col]), pct(e.loc["embed", col])], textposition="outside", cliponaxis=False)
                fig.update_layout(barmode="group", xaxis=dict(visible=False, range=[0, 1.15]), legend=dict(traceorder="reversed"))
                show(fig, 150)

    with c:
        with card("Watch Time per Device", "Donut = share jam tonton per family. Bar = rata-rata durasi per play (capped 40 menit)."):
            fam = f.groupby("family")["watch_seconds"].sum() / 3600
            fig = go.Figure(go.Pie(labels=fam.index, values=fam.values, hole=.62, marker=dict(colors=[fcolor(x) for x in fam.index]),
                                   textinfo="percent", sort=False))
            fig.update_layout(annotations=[dict(text=f"<b>{compact(fam.sum())}</b><br>jam", showarrow=False, font=dict(size=15))],
                              legend=dict(orientation="h", y=-.05))
            show(fig, 250)
            av = p.sort_values("avg_sec")
            fig = go.Figure(go.Bar(y=av.index, x=av.avg_sec, orientation="h", marker_color=[pcolor(i) for i in av.index],
                                   text=[f"{idn(v)} dtk" for v in av.avg_sec], textposition="outside", cliponaxis=False))
            fig.update_layout(xaxis=dict(visible=False, range=[0, av.avg_sec.max() * 1.3]))
            st.markdown("**Rata-rata durasi per play**")
            show(fig, 210)

    with d:
        days = f.groupby(["date", "platform"])["play_ge_30s"].sum().reset_index()
        win = days.loc[days.groupby("date")["play_ge_30s"].idxmax()].set_index("date")["platform"]
        lead_day = win.value_counts().idxmax()
        with card(f"{lead_day} memimpin {int((win == lead_day).sum())} dari {len(win)} hari", "Leader engaged plays per hari. 1 kotak = 1 hari."):
            st.markdown("<div class='strip'>" + "".join(f"<span style='background:{pcolor(w)}' title='{i:%d %b}: {w}'></span>" for i, w in win.items()) + "</div>",
                        unsafe_allow_html=True)
            st.markdown("**Sensitivity: ganti definisi 'engaged'**")
            rows, winners = "", []
            for t, lab in THRESHOLDS:
                col = f"play_ge_{t}s"
                s_ = f.groupby("platform")[col].sum()
                w_ = s_.idxmax(); winners.append(w_)
                rows += f"<tr><td>{lab}</td><td>{w_}</td><td>{pct(s_.max() / s_.sum())}</td></tr>"
            st.markdown(f"<div class='tbl-wrap'><table class='t'><tr><th>THRESHOLD</th><th>LEADER</th><th>SHARE</th></tr>{rows}</table></div>", unsafe_allow_html=True)
            top_n = pd.Series(winners).value_counts()
            st.markdown(f"<div style='margin-top:10px'><span class='badge {'good' if top_n.iloc[0] == len(winners) else 'warn'}'>"
                        f"{top_n.index[0]} leader di {top_n.iloc[0]} dari {len(winners)} threshold</span></div>", unsafe_allow_html=True)
            raw_l, cap_l = p.ws.idxmax(), p.wsc.idxmax()
            st.markdown(f"<span class='badge {'good' if raw_l == cap_l else 'warn'}'>Outlier check</span>"
                        f"<span class='note'>Watch time {'tetap' if raw_l == cap_l else 'berubah'}: {raw_l} (raw), {cap_l} (capped 40 menit)</span>", unsafe_allow_html=True)

# =====================================================================================
# PAGE 2
# =====================================================================================
with tab2:
    ft = f[~f["date"].isin(PARTIAL)]
    if ft["date"].nunique() < 3:
        ft = f
    note_partial = " Hari pertama dan terakhir tidak dihitung (data terpotong zona waktu)." if len(ft) < len(f) else ""
    hr = ft.groupby("hour")["plays"].sum()
    pk = int(hr.idxmax())
    ct = ft.groupby("tipe").agg(plays=("plays", "sum"), ws=("watch_seconds", "sum"))
    live_p = ct.loc["Live TV", "plays"] / ct.plays.sum() if "Live TV" in ct.index else 0
    live_w = ct.loc["Live TV", "ws"] / ct.ws.sum() if "Live TV" in ct.index and ct.ws.sum() else 0

    g = ft.assign(daypart=pd.cut(ft["hour"], [-1, 4, 10, 14, 18, 23], labels=DAYPARTS))
    ndays = g.groupby("weekday")["date"].nunique()
    cell = g.groupby(["daypart", "weekday"], observed=False)["plays"].sum().unstack("weekday").reindex(columns=range(1, 8)).fillna(0)
    cell = cell.div(ndays.reindex(range(1, 8)).replace(0, np.nan), axis=1).fillna(0)
    top_cell = cell.stack().idxmax()
    daily = f.groupby(["date", "family"])["plays"].sum().reset_index()
    dtot = ft.groupby("date")["plays"].sum()
    wk = dtot[dtot.index.dayofweek >= 5].mean()
    wd = dtot[dtot.index.dayofweek < 5].mean()
    wk_lift = (wk / wd - 1) if (wd and not np.isnan(wk) and not np.isnan(wd)) else np.nan

    st.markdown(f"<div class='banner'><div class='k'>When & What · {rng}</div><div class='h'>Peak hour: {pk:02d}:00 WIB</div>"
                f"<div class='s'>Live TV cuma {pct(live_p)} dari plays tapi {pct(live_w)} dari watch time. Slot paling padat: {top_cell[0]} WIB, hari {DAYS[top_cell[1]]}.</div></div>",
                unsafe_allow_html=True)
    st.write("")
    r1a, r1b, r1c = st.columns([1.05, 1.2, .95], gap="small")

    with r1a:
        with card("Peak Time: Daypart × Hari", "Share dari rata-rata plays per hari." + note_partial):
            z = (cell / cell.values.sum()).values
            fig = go.Figure(go.Heatmap(z=z, x=[DAYS[i] for i in range(1, 8)], y=DAYPARTS, colorscale=[[0, "#FFF0F3"], [1, RED]], showscale=False,
                                       text=[[pct(v) for v in r] for r in z], texttemplate="%{text}", textfont=dict(size=10), hovertemplate="%{y}, %{x}: %{text}<extra></extra>"))
            fig.update_yaxes(autorange="reversed")
            show(fig, 290)
    with r1b:
        with card("Pola Jam per Device", "% plays per jam (WIB), dinormalisasi per family"):
            hf = ft.groupby(["family", "hour"])["plays"].sum().unstack("hour").reindex(columns=range(24)).fillna(0)
            fig = go.Figure()
            for fam_, row in hf.iterrows():
                if row.sum() == 0:
                    continue
                s_ = row / row.sum()
                fig.add_scatter(x=list(range(24)), y=s_, name=fam_, mode="lines", line=dict(color=fcolor(fam_), width=2.5))
                j = int(s_.idxmax())
                fig.add_scatter(x=[j], y=[s_.max()], mode="markers+text", text=[f"{j:02d}:00 · {pct(s_.max())}"], textposition="top center",
                                marker=dict(color=fcolor(fam_), size=8), showlegend=False)
            fig.update_layout(yaxis=dict(tickformat=".0%", gridcolor="#EEF0F5", rangemode="tozero"), xaxis=dict(dtick=4, ticksuffix=":00"))
            show(fig, 290)
    with r1c:
        with card("Live TV vs VOD", "Share plays vs share watch time"):
            order_ = [t for t in ["VOD", "Live TV", "Catch-up", "Unknown"] if t in ct.index]
            fig = go.Figure()
            for t in order_:
                fig.add_bar(y=["Watch time", "Plays"], x=[ct.loc[t, "ws"] / ct.ws.sum(), ct.loc[t, "plays"] / ct.plays.sum()], name=t, orientation="h",
                            marker_color=CT_COLOR.get(t, GREY), text=[pct(ct.loc[t, "ws"] / ct.ws.sum()), pct(ct.loc[t, "plays"] / ct.plays.sum())],
                            textposition="inside")
            fig.update_layout(barmode="stack", xaxis=dict(visible=False, range=[0, 1]), bargap=.45)
            show(fig, 290)

    r2a, r2b, r2c = st.columns([1.1, .9, 1.2], gap="small")
    with r2a:
        ttl = f"Weekend {wk_lift:+.0%} vs weekday" if not np.isnan(wk_lift) else "Daily plays trend"
        with card(ttl, "Plays per hari per family. Titik pertama dan terakhir bisa lebih rendah karena data terpotong." if PARTIAL else "Plays per hari per family"):
            fig = go.Figure()
            for fam_, gg in daily.groupby("family"):
                fig.add_scatter(x=gg["date"], y=gg["plays"], name=fam_, mode="lines+markers", line=dict(color=fcolor(fam_), width=2.5), marker=dict(size=4))
            fig.update_layout(yaxis=dict(gridcolor="#EEF0F5", rangemode="tozero"))
            show(fig, 300)
    with r2b:
        with card("Distribusi Durasi per Play", "% plays per bucket durasi"):
            pl = f["plays"].sum()
            ge = {k: f[f"play_ge_{k}s"].sum() for k in ("10", "30", "60", "180", "600")}
            bins = [("0 dtk", f["zero_plays"].sum()), ("1–9", pl - f["zero_plays"].sum() - ge["10"]), ("10–29", ge["10"] - ge["30"]),
                    ("30–59", ge["30"] - ge["60"]), ("1–3 mnt", ge["60"] - ge["180"]), ("3–10 mnt", ge["180"] - ge["600"]), (">10 mnt", ge["600"])]
            fig = go.Figure(go.Bar(x=[b_[0] for b_ in bins], y=[b_[1] / pl for b_ in bins], marker_color=[RED] + [SLATE] * 6,
                                   text=[pct(b_[1] / pl) for b_ in bins], textposition="outside", cliponaxis=False))
            fig.update_layout(yaxis=dict(visible=False, range=[0, max(b_[1] / pl for b_ in bins) * 1.2]), bargap=.25)
            show(fig, 300)
    with r2c:
        with card("Login vs Guest dan Buffering"):
            lg = f[f["status_user"].isin(["login", "guest"])]
            if not lg.empty:
                e_ = agg(lg, ["family", "status_user"]).reset_index()
                fig = go.Figure()
                for s_, colr in [("login", RED), ("guest", GREY)]:
                    gg = e_[e_["status_user"] == s_]
                    fig.add_bar(y=gg["family"], x=gg["eng_rate"], name=s_.title(), orientation="h", marker_color=colr, text=[pct(v) for v in gg["eng_rate"]],
                                textposition="outside", cliponaxis=False)
                fig.update_layout(barmode="group", xaxis=dict(visible=False, range=[0, 1.15]), legend=dict(traceorder="reversed"))
                st.markdown("**Engaged rate: login vs guest**")
                show(fig, 170)
            bf = p.sort_values("buf_rate").tail(4)
            fig = go.Figure(go.Bar(y=bf.index, x=bf.buf_rate, orientation="h", marker_color=[pcolor(i) for i in bf.index],
                                   text=[pct(v) for v in bf.buf_rate], textposition="outside", cliponaxis=False))
            fig.update_layout(xaxis=dict(visible=False, range=[0, max(bf.buf_rate.max() * 1.25, .01)]))
            st.markdown("**Buffering rate per platform (top 4)**")
            show(fig, 150)

    # takeaways
    st.write("")
    emb = f.groupby("sumber").agg(plays=("plays", "sum"), zero=("zero_plays", "sum"))
    t1 = ("Fokus ke jam sibuk", f"Peak di {pk:02d}:00 WIB" + (f", weekend {wk_lift:+.0%} dibanding weekday" if not np.isnan(wk_lift) else "") + ". Cocok untuk jadwal promo dan rilis konten.")
    if {"embed", "non-embed"} <= set(emb.index) and emb.loc["embed", "plays"] > 0:
        t2 = ("Cek autoplay di embed", f"{pct(emb.loc['embed', 'zero'] / emb.loc['embed', 'plays'])} play embed berhenti di 0 detik (non-embed {pct(emb.loc['non-embed', 'zero'] / emb.loc['non-embed', 'plays'])}). "
              "Penyebabnya belum terbukti, jadi test A/B dulu sebelum autoplay diubah.")
    else:
        t2 = ("Cek play 0 detik", f"{pct(tot.zero_plays / tot.plays)} plays berhenti di 0 detik. Cari tahu sumber terbesarnya sebelum dipakai sebagai angka penontonan.")
    if not lg.empty and lg["status_user"].nunique() == 2:
        pool = lg.groupby("status_user")[["play_ge_30s", "plays"]].sum()
        diff = (pool.loc["login", "play_ge_30s"] / pool.loc["login", "plays"] - pool.loc["guest", "play_ge_30s"] / pool.loc["guest", "plays"]) * 100
        t3 = ("Cek efek login", f"Engaged rate login vs guest: {diff:+.0f} poin (semua platform digabung). Bandingkan di platform yang sama dulu sebelum jadi target.")
    else:
        t3 = ("Cek efek login", "Pilih login dan guest sekaligus di filter untuk membandingkan engaged rate.")
    with card("Key Takeaways"):
        tc = st.columns(3)
        for col_, (t, body) in zip(tc, (t1, t2, t3)):
            col_.markdown(f"<div class='take'><b>{t}</b><div>{body}</div></div>", unsafe_allow_html=True)

st.markdown(f"<div class='footer'>{AUTHOR_LINE}</div>", unsafe_allow_html=True)
