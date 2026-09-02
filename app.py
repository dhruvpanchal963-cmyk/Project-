from pathlib import Path
import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Consistent readable Plotly defaults
import plotly.io as pio
pio.templates.default = "plotly_white"

st.set_page_config(
    page_title="FITX UX Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Theme / CSS
# -----------------------------
st.markdown(
    """
    <style>
    .stApp { background: #f6f8fb !important; color: #101828 !important; }
    html, body { color: #101828 !important; }
    .stApp .stMarkdown, .stApp .stText, .stApp label, .stApp p { color: #101828 !important; }
    .stApp [data-testid="stMarkdownContainer"] * { color: #101828 !important; }
    .stApp [data-testid="stWidgetLabel"] * { color: #101828 !important; }
    .stApp input, .stApp textarea { color: #101828 !important; background: #ffffff !important; }
    .stApp input::placeholder, .stApp textarea::placeholder { color: #667085 !important; opacity: 1 !important; }
    .stApp header { background: transparent !important; }
    [data-testid="stHeader"] { background: transparent !important; }
    [data-testid="stToolbar"] { color: #101828 !important; }
    .stSelectbox label, .stMultiSelect label, .stSlider label, .stRadio label, .stFileUploader label { color: #101828 !important; font-weight: 650 !important; }
    [data-baseweb="select"] * { color: #101828 !important; }
    [data-baseweb="popover"] * { color: #101828 !important; }
    .stButton button, .stDownloadButton button { color: #101828 !important; background: #ffffff !important; border: 1px solid #d0d5dd !important; }
    .stMetric label, .stMetric [data-testid="stMetricValue"], .stMetric [data-testid="stMetricDelta"] { color: #101828 !important; }
    [data-testid="stDataFrame"] { background: #ffffff !important; }
    [data-testid="stDataFrame"] * { color: #101828 !important; }
    [data-testid="stDataFrame"] [role="columnheader"] { color: #101828 !important; background: #f8fafc !important; }
    [data-testid="stDataFrame"] [role="gridcell"] { color: #101828 !important; }
    .stTabs [data-baseweb="tab-list"] button { color: #344054 !important; }
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] { color: #111827 !important; }
    .stAlert p, .stInfo p, .stSuccess p, .stWarning p, .stError p { color: #101828 !important; }
    .stCaption, [data-testid="stCaptionContainer"] { color: #475467 !important; }
    .block-container { padding-top: 1.35rem; padding-bottom: 2.5rem; max-width: 1500px; }
    [data-testid="stSidebar"] { background: #101828; }
    [data-testid="stSidebar"] * { color: #f8fafc; }
    .hero {
        padding: 1.25rem 1.4rem;
        border-radius: 18px;
        background: linear-gradient(120deg, #111827 0%, #1f2937 55%, #334155 100%);
        color: white;
        margin-bottom: 1.0rem;
        box-shadow: 0 10px 30px rgba(15, 23, 42, .10);
    }
    .hero h1 { margin: 0; font-size: 2rem; }
    .hero p { margin: .45rem 0 0; color: #cbd5e1; }
    .section-title { font-size: 1.25rem; font-weight: 750; margin: .35rem 0 .5rem; color: #111827; }
    .insight-card {
        padding: 1rem 1.05rem;
        border-radius: 14px;
        background: #ffffff;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 16px rgba(15, 23, 42, .04);
        height: 100%;
    }
    .insight-card h4 { margin: 0 0 .45rem; color: #111827; }
    .insight-card p { margin: 0; color: #475467; line-height: 1.5; }
    [data-testid="stMetric"] {
        background: white;
        border: 1px solid #e5e7eb;
        padding: 1rem;
        border-radius: 14px;
        box-shadow: 0 4px 16px rgba(15,23,42,.04);
    }
    div[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
    .small-note { color:#667085; font-size:.86rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "FITX_UX_UI_Analysis_Dashboard_Aligned.xlsx"

@st.cache_data(show_spinner=False)
def load_excel_from_bytes(raw: bytes):
    xls = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

@st.cache_data(show_spinner=False)
def load_excel_from_path(path: str):
    xls = pd.ExcelFile(path, engine="openpyxl")
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}


def num(v, default=0.0):
    try:
        if pd.isna(v):
            return default
        return float(v)
    except Exception:
        return default


def clean_numeric(df, cols):
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


def remove_empty(df):
    return df.dropna(how="all").copy()


def get_metric(profile, metric_name, default=0):
    hit = profile.loc[profile["Metric"].astype(str).str.strip().eq(metric_name), "Value"]
    return num(hit.iloc[0], default) if len(hit) else default


def fmt_int(v):
    return f"{int(round(num(v))):,}"


def fmt_pct(v, decimals=1):
    return f"{num(v) * 100:.{decimals}f}%"


def style_fig(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=60, b=20),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Arial, sans-serif", color="#101828", size=13),
        title_font=dict(size=18, color="#101828"),
        hoverlabel=dict(bgcolor="#101828", font_color="#FFFFFF", font_size=13),
        legend_title_text="",
    )
    fig.update_xaxes(showgrid=False, zeroline=False, tickfont=dict(color="#344054"), title_font=dict(color="#101828"))
    fig.update_yaxes(gridcolor="#e5e7eb", zeroline=False, tickfont=dict(color="#344054"), title_font=dict(color="#101828"))
    return fig


# -----------------------------
# Data loading
# -----------------------------
with st.sidebar:
    st.markdown("## FITX Analytics")
    st.caption("UX/UI behavior intelligence")
    st.divider()
    uploaded = st.file_uploader("Upload workbook (optional)", type=["xlsx"])

if uploaded is not None:
    sheets = load_excel_from_bytes(uploaded.getvalue())
    source_name = uploaded.name
else:
    local_path = DATA_FILE
    if local_path.exists():
        sheets = load_excel_from_path(str(local_path))
        source_name = local_path.name
    else:
        st.error(
            f"Data file not found. Add `{DATA_FILE.name}` beside app.py in your GitHub repository, "
            "or upload the workbook from the sidebar."
        )
        st.stop()

# Main data frames
profile = remove_empty(sheets.get("01_User_Profile", pd.DataFrame(columns=["Metric", "Value"])))
segments = remove_empty(sheets.get("02_User_Segments", pd.DataFrame()))
devices = remove_empty(sheets.get("03_Device_Analysis", pd.DataFrame()))
flows = remove_empty(sheets.get("06_Top_Navigation_Flows", pd.DataFrame()))
funnels = remove_empty(sheets.get("07_Funnels", pd.DataFrame()))
cta = remove_empty(sheets.get("08_CTA_Analysis", pd.DataFrame()))
buttons = remove_empty(sheets.get("09_Button_Analysis", pd.DataFrame()))
exits = remove_empty(sheets.get("10_Exit_Analysis", pd.DataFrame()))
exit_reasons = remove_empty(sheets.get("11_Exit_Reason_By_Page", pd.DataFrame()))
forms = remove_empty(sheets.get("12_Form_Analysis", pd.DataFrame()))
errors = remove_empty(sheets.get("13_Validation_Errors", pd.DataFrame()))
scroll = remove_empty(sheets.get("14_Scroll_Analysis", pd.DataFrame()))
corr_raw = remove_empty(sheets.get("15_Correlation", pd.DataFrame()))
insights = remove_empty(sheets.get("16_UX_Insights", pd.DataFrame()))
executive = remove_empty(sheets.get("17_Executive_Summary", pd.DataFrame()))

# Remove embedded summary rows from device table
devices_main = devices.copy()
if "Device" in devices_main.columns:
    devices_main = devices_main[devices_main["Device"].astype(str).str.lower().isin(["desktop", "mobile", "tablet"])]

devices_main = clean_numeric(devices_main, [
    "no of users", "Events", "Avg Pages/Session", "CTA Clicks", "Button Clicks",
    "Form Abandonments", "Validation Errors", "Repeat Users", "One-time Users",
    "Repeat-user %", "One-time-user %", "Screenshot Device Distribution"
])

flows = clean_numeric(flows, ["Rank", "Unique Users", "Unique Sessions", "Transitions", "% of All Transitions"])
funnels = clean_numeric(funnels, ["Stage", "Users Reaching Stage", "Drop-off Count", "Drop-off %", "Continuation %"])
cta = clean_numeric(cta, ["Rank", "Unique Users", "Click Count", "Unique Sessions", "% of All Users"])
buttons = clean_numeric(buttons, ["Rank", "Unique Users", "Click Count", "Unique Sessions", "% of All Users"])
exits = clean_numeric(exits, ["Exit Users", "Exit Count", "Unique Exit Sessions"])
forms = clean_numeric(forms, ["Forms Started", "Forms Submitted", "Forms Abandoned"])
errors_top = errors.copy()
if "Error Type" in errors_top.columns:
    # Keep only actual error rows before embedded sub-section.
    stop_idx = errors_top.index[errors_top["Error Type"].astype(str).eq("Errors by Page/Form")]
    if len(stop_idx):
        errors_top = errors_top.loc[errors_top.index < stop_idx[0]]
errors_top = clean_numeric(errors_top, ["Error Count", "Users Encountering Error", "Sessions With Error"])

# KPIs
users = get_metric(profile, "Total unique users")
sessions = get_metric(profile, "Total unique sessions")
events = get_metric(profile, "Total events")
sessions_per_user = get_metric(profile, "Average sessions per user")
events_per_session = get_metric(profile, "Average events per session")
repeat_pct = get_metric(profile, "Repeat user percentage")

with st.sidebar:
    page = st.radio(
        "Explore",
        [
            "Executive Overview", "Audience & Devices", "Navigation", "Conversion",
            "Forms & Errors", "Exit & Scroll", "Correlation", "UX Recommendations", "Raw Data"
        ],
        index=0,
    )
    st.divider()
    st.caption(f"Data source: {source_name}")
    st.caption(f"Sheets loaded: {len(sheets)}")

st.markdown(
    """
    <div class="hero">
        <h1>FITX UX Analytics Dashboard</h1>
        <p>Behavior, conversion, interaction friction and product-experience insights from the FITX analytics workbook.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# EXECUTIVE OVERVIEW
# -----------------------------
if page == "Executive Overview":
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Unique users", fmt_int(users))
    c2.metric("Sessions", fmt_int(sessions))
    c3.metric("Events", fmt_int(events))
    c4.metric("Sessions / user", f"{sessions_per_user:.2f}")
    c5.metric("Repeat-user rate", fmt_pct(repeat_pct))

    st.markdown('<div class="section-title">Executive signals</div>', unsafe_allow_html=True)
    a, b, c = st.columns(3)

    mobile_repeat = 0
    if len(devices_main) and "Repeat-user %" in devices_main:
        m = devices_main.loc[devices_main["Device"].astype(str).str.lower().eq("mobile"), "Repeat-user %"]
        mobile_repeat = num(m.iloc[0]) if len(m) else 0

    class_final = 0
    class_start = 0
    if len(funnels) and "Funnel" in funnels:
        ff = funnels[funnels["Funnel"].astype(str).str.contains("Classes", case=False, na=False)].sort_values("Stage")
        if len(ff):
            class_start = num(ff.iloc[0]["Users Reaching Stage"])
            class_final = num(ff.iloc[-1]["Users Reaching Stage"])

    with a:
        st.markdown(
            f'<div class="insight-card"><h4>📱 Retention opportunity</h4><p>Mobile has the strongest displayed repeat-user rate at <b>{mobile_repeat*100:.2f}%</b>. Investigate what works well in the mobile journey and transfer those patterns to other devices.</p></div>',
            unsafe_allow_html=True,
        )
    with b:
        conv = (class_final / class_start * 100) if class_start else 0
        st.markdown(
            f'<div class="insight-card"><h4>🎯 Booking funnel risk</h4><p>The Classes journey moves from <b>{fmt_int(class_start)}</b> homepage users to only <b>{fmt_int(class_final)}</b> confirmed bookings, an end-to-end continuation of roughly <b>{conv:.1f}%</b>.</p></div>',
            unsafe_allow_html=True,
        )
    with c:
        top_err = errors_top.sort_values("Error Count", ascending=False).iloc[0] if len(errors_top) else None
        err_text = f"{top_err['Error Type']} ({fmt_int(top_err['Error Count'])})" if top_err is not None else "N/A"
        st.markdown(
            f'<div class="insight-card"><h4>⚠️ Form friction</h4><p>The largest recorded validation category is <b>{err_text}</b>. Form feedback, required-field guidance and input formatting should be UX priorities.</p></div>',
            unsafe_allow_html=True,
        )

    left, right = st.columns([1.05, 1])
    with left:
        if len(cta):
            chart = cta.dropna(subset=["CTA", "Click Count"]).sort_values("Click Count", ascending=True)
            fig = px.bar(chart, x="Click Count", y="CTA", orientation="h", title="CTA engagement", text="Click Count")
            fig.update_traces(textposition="outside")
            st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        if len(devices_main):
            fig = px.bar(
                devices_main,
                x="Device",
                y=["Repeat Users", "One-time Users"],
                barmode="stack",
                title="User retention mix by device",
            )
            st.plotly_chart(style_fig(fig), use_container_width=True)

    left, right = st.columns([1, 1])
    with left:
        if len(funnels):
            names = funnels["Funnel"].dropna().astype(str).unique().tolist()
            selected = names[0] if names else None
            if selected:
                fdf = funnels[funnels["Funnel"].eq(selected)].sort_values("Stage")
                fig = go.Figure(go.Funnel(y=fdf["Stage Name"], x=fdf["Users Reaching Stage"], textinfo="value+percent initial"))
                fig.update_layout(title=selected)
                st.plotly_chart(style_fig(fig), use_container_width=True)
    with right:
        if len(forms):
            f = forms.dropna(subset=["Form"]).copy()
            if len(f):
                f["Abandonment Rate"] = np.where(f["Forms Started"] > 0, f["Forms Abandoned"] / f["Forms Started"] * 100, np.nan)
                f = f.sort_values("Abandonment Rate", ascending=True)
                fig = px.bar(f, x="Abandonment Rate", y="Form", orientation="h", title="Form abandonment rate", text_auto=".1f")
                fig.update_xaxes(title="Abandonment rate (%)")
                st.plotly_chart(style_fig(fig), use_container_width=True)

    if len(executive):
        st.markdown('<div class="section-title">Executive findings</div>', unsafe_allow_html=True)
        cols = [c for c in ["Key Finding", "Supporting Metric / Evidence", "UX Meaning", "Recommended Action"] if c in executive.columns]
        st.dataframe(executive[cols].dropna(how="all"), use_container_width=True, hide_index=True)

# -----------------------------
# AUDIENCE & DEVICES
# -----------------------------
elif page == "Audience & Devices":
    st.markdown('<div class="section-title">Audience segmentation</div>', unsafe_allow_html=True)
    if len(segments):
        seg = clean_numeric(segments, ["Users", "Percentage of Users", "Avg Sessions", "Avg Pages Viewed", "Avg Events", "Avg Clicks", "Avg Scroll Depth"])
        c1, c2 = st.columns([.9, 1.1])
        with c1:
            fig = px.pie(seg, names="Segment", values="Users", hole=.55, title="Engagement segment mix")
            st.plotly_chart(style_fig(fig), use_container_width=True)
        with c2:
            fig = px.bar(seg, x="Segment", y=["Avg Pages Viewed", "Avg Clicks"], barmode="group", title="Average engagement by segment")
            st.plotly_chart(style_fig(fig), use_container_width=True)

    st.markdown('<div class="section-title">Device performance</div>', unsafe_allow_html=True)
    if len(devices_main):
        metric = st.selectbox(
            "Device metric",
            ["Events", "CTA Clicks", "Button Clicks", "Form Abandonments", "Validation Errors", "Repeat-user %"],
        )
        fig = px.bar(devices_main, x="Device", y=metric, text_auto=".2s" if metric != "Repeat-user %" else ".3f", title=f"{metric} by device")
        if metric == "Repeat-user %":
            fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(style_fig(fig), use_container_width=True)
        st.dataframe(devices_main, use_container_width=True, hide_index=True)

# -----------------------------
# NAVIGATION
# -----------------------------
elif page == "Navigation":
    st.markdown('<div class="section-title">Top navigation flows</div>', unsafe_allow_html=True)
    if len(flows):
        top_n = st.slider("Number of flows", 5, min(25, len(flows)), min(15, len(flows)))
        f = flows.dropna(subset=["from_page", "to_page", "Transitions"]).sort_values("Transitions", ascending=False).head(top_n)

        # Sankey
        labels = list(pd.unique(pd.concat([f["from_page"], f["to_page"]], ignore_index=True).astype(str)))
        idx = {label: i for i, label in enumerate(labels)}
        fig = go.Figure(
            go.Sankey(
                node=dict(label=labels, pad=14, thickness=18),
                link=dict(
                    source=[idx[str(v)] for v in f["from_page"]],
                    target=[idx[str(v)] for v in f["to_page"]],
                    value=f["Transitions"].fillna(0).tolist(),
                    customdata=f["Unique Users"].fillna(0).tolist(),
                    hovertemplate="%{source.label} → %{target.label}<br>Transitions: %{value}<br>Unique users: %{customdata}<extra></extra>",
                ),
            )
        )
        fig.update_layout(title="Navigation flow Sankey")
        st.plotly_chart(style_fig(fig, 560), use_container_width=True)

        plot = f.copy()
        plot["Flow"] = plot["from_page"].astype(str) + " → " + plot["to_page"].astype(str)
        fig = px.bar(plot.sort_values("Transitions"), x="Transitions", y="Flow", orientation="h", title="Top page-to-page transitions")
        st.plotly_chart(style_fig(fig, 500), use_container_width=True)
        st.dataframe(f, use_container_width=True, hide_index=True)

# -----------------------------
# CONVERSION
# -----------------------------
elif page == "Conversion":
    st.markdown('<div class="section-title">Conversion funnels</div>', unsafe_allow_html=True)
    if len(funnels):
        names = funnels["Funnel"].dropna().astype(str).unique().tolist()
        selected = st.selectbox("Choose funnel", names)
        fdf = funnels[funnels["Funnel"].astype(str).eq(selected)].sort_values("Stage")
        a, b, c = st.columns(3)
        start = num(fdf.iloc[0]["Users Reaching Stage"]) if len(fdf) else 0
        final = num(fdf.iloc[-1]["Users Reaching Stage"]) if len(fdf) else 0
        a.metric("Starting users", fmt_int(start))
        b.metric("Final users", fmt_int(final))
        c.metric("End-to-end conversion", f"{(final/start*100 if start else 0):.2f}%")

        fig = go.Figure(go.Funnel(y=fdf["Stage Name"], x=fdf["Users Reaching Stage"], textinfo="value+percent initial+percent previous"))
        fig.update_layout(title=selected)
        st.plotly_chart(style_fig(fig, 470), use_container_width=True)

    st.markdown('<div class="section-title">Call-to-action performance</div>', unsafe_allow_html=True)
    if len(cta):
        fig = px.scatter(
            cta.dropna(subset=["CTA"]), x="Unique Users", y="Click Count", size="Unique Sessions",
            text="CTA", title="CTA reach vs click volume", hover_name="CTA"
        )
        fig.update_traces(textposition="top center")
        st.plotly_chart(style_fig(fig, 470), use_container_width=True)
        st.dataframe(cta.dropna(subset=["CTA"]), use_container_width=True, hide_index=True)

    if len(buttons):
        st.markdown('<div class="section-title">Button interactions</div>', unsafe_allow_html=True)
        top_buttons = buttons.dropna().sort_values("Click Count", ascending=False).head(15)
        fig = px.bar(top_buttons.sort_values("Click Count"), x="Click Count", y=top_buttons.columns[1], orientation="h", title="Most clicked controls")
        st.plotly_chart(style_fig(fig, 500), use_container_width=True)

# -----------------------------
# FORMS & ERRORS
# -----------------------------
elif page == "Forms & Errors":
    st.markdown('<div class="section-title">Form completion and abandonment</div>', unsafe_allow_html=True)
    if len(forms):
        f = forms.dropna(subset=["Form"]).copy()
        f["Completion Rate"] = np.where(f["Forms Started"] > 0, f["Forms Submitted"] / f["Forms Started"], np.nan)
        f["Abandonment Rate"] = np.where(f["Forms Started"] > 0, f["Forms Abandoned"] / f["Forms Started"], np.nan)
        fig = px.bar(f, x="Form", y=["Forms Submitted", "Forms Abandoned"], barmode="group", title="Submitted vs abandoned forms")
        st.plotly_chart(style_fig(fig), use_container_width=True)
        display = f.copy()
        display["Completion Rate"] = (display["Completion Rate"] * 100).round(1).astype(str) + "%"
        display["Abandonment Rate"] = (display["Abandonment Rate"] * 100).round(1).astype(str) + "%"
        st.dataframe(display, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Validation errors</div>', unsafe_allow_html=True)
    if len(errors_top):
        e = errors_top.dropna(subset=["Error Type", "Error Count"]).sort_values("Error Count", ascending=True)
        fig = px.bar(e, x="Error Count", y="Error Type", orientation="h", title="Most frequent validation errors", text="Error Count")
        st.plotly_chart(style_fig(fig), use_container_width=True)
        st.dataframe(errors_top, use_container_width=True, hide_index=True)

# -----------------------------
# EXIT & SCROLL
# -----------------------------
elif page == "Exit & Scroll":
    st.markdown('<div class="section-title">Exit behavior</div>', unsafe_allow_html=True)
    if len(exits):
        x = exits.dropna(subset=["Page", "Exit Count"]).sort_values("Exit Count", ascending=True)
        fig = px.bar(x, x="Exit Count", y="Page", orientation="h", title="Exit count by page")
        st.plotly_chart(style_fig(fig, 500), use_container_width=True)

    if len(exit_reasons):
        st.markdown('<div class="section-title">Exit reasons by page</div>', unsafe_allow_html=True)
        er = exit_reasons.copy()
        num_cols = [c for c in er.columns if c not in ["Page", "Exit Reason"]]
        for c in num_cols:
            er[c] = pd.to_numeric(er[c], errors="coerce")
        if "Page" in er.columns and "Exit Reason" in er.columns and num_cols:
            metric_col = num_cols[0]
            plot_er = er.dropna(subset=["Page", "Exit Reason", metric_col]).sort_values(metric_col, ascending=False).head(15)
            fig = px.bar(plot_er, x=metric_col, y="Exit Reason", color="Page", orientation="h", title=f"Top exit reasons by {metric_col}")
            st.plotly_chart(style_fig(fig, 520), use_container_width=True)
        st.dataframe(exit_reasons, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Scroll depth</div>', unsafe_allow_html=True)
    if len(scroll):
        s = scroll.copy()
        # First section: event counts by depth.
        if "Scroll Depth %" in s.columns and "Event Count" in s.columns:
            ss = s[["Scroll Depth %", "Event Count"]].copy()
            ss["Scroll Depth %"] = pd.to_numeric(ss["Scroll Depth %"], errors="coerce")
            ss["Event Count"] = pd.to_numeric(ss["Event Count"], errors="coerce")
            ss = ss.dropna()
            ss = ss[(ss["Scroll Depth %"] >= 0) & (ss["Scroll Depth %"] <= 1)]
            if len(ss):
                ss["Depth"] = (ss["Scroll Depth %"] * 100).round(0).astype(int).astype(str) + "%"
                fig = px.line(ss, x="Depth", y="Event Count", markers=True, title="Scroll-depth event volume")
                st.plotly_chart(style_fig(fig), use_container_width=True)

# -----------------------------
# CORRELATION
# -----------------------------
elif page == "Correlation":
    st.markdown('<div class="section-title">Behavior correlation matrix</div>', unsafe_allow_html=True)
    if len(corr_raw):
        expected = ["sessions", "pages_viewed", "events", "cta_clicks", "button_clicks", "avg_scroll_depth", "form_abandonments", "validation_errors", "pages_per_session"]
        rows = corr_raw[corr_raw["Variable"].astype(str).isin(expected)].copy() if "Variable" in corr_raw.columns else pd.DataFrame()
        available = [c for c in expected if c in rows.columns]
        if len(rows) and available:
            # The workbook can contain duplicate Variable labels (or duplicate
            # rows after Excel parsing). Pandas refuses to reindex a DataFrame
            # with duplicate axis labels, so normalize both axes first.
            available = list(dict.fromkeys(available))
            matrix = rows.set_index("Variable")[available].apply(pd.to_numeric, errors="coerce")
            matrix = matrix.groupby(level=0, sort=False).mean(numeric_only=True)
            matrix = matrix.loc[~matrix.index.duplicated(keep="first")]
            matrix = matrix.loc[:, ~matrix.columns.duplicated(keep="first")]
            ordered_index = [x for x in expected if x in matrix.index]
            ordered_columns = [x for x in expected if x in matrix.columns]
            matrix = matrix.reindex(index=ordered_index, columns=ordered_columns)
            fig = px.imshow(matrix, text_auto=".2f", zmin=-1, zmax=1, aspect="auto", title="Correlation heatmap")
            fig.update_layout(height=620)
            st.plotly_chart(fig, use_container_width=True)

            # Strongest off-diagonal relationships
            pairs = []
            for i in matrix.index:
                for j in matrix.columns:
                    if i != j and pd.notna(matrix.loc[i, j]):
                        pairs.append((i, j, float(matrix.loc[i, j]), abs(float(matrix.loc[i, j]))))
            if pairs:
                pairs_df = pd.DataFrame(pairs, columns=["Variable A", "Variable B", "Correlation", "Strength"])
                pairs_df["pair_key"] = pairs_df.apply(lambda r: "|".join(sorted([r["Variable A"], r["Variable B"]])), axis=1)
                pairs_df = pairs_df.sort_values("Strength", ascending=False).drop_duplicates("pair_key").drop(columns=["pair_key", "Strength"]).head(10)
                st.markdown('<div class="section-title">Strongest relationships</div>', unsafe_allow_html=True)
                st.dataframe(pairs_df, use_container_width=True, hide_index=True)
        st.info("Correlation values are displayed from the workbook's retained analysis. Treat correlation as association, not proof of causation.")

# -----------------------------
# UX RECOMMENDATIONS
# -----------------------------
elif page == "UX Recommendations":
    st.markdown('<div class="section-title">Prioritized UX recommendations</div>', unsafe_allow_html=True)
    if len(insights):
        data = insights.dropna(subset=["Finding"]).copy()
        priority_order = ["High", "Medium", "Low"]
        if "Priority" in data:
            selected = st.multiselect("Priority", priority_order, default=priority_order)
            data = data[data["Priority"].isin(selected)]
        if "Category" in data:
            cats = sorted(data["Category"].dropna().astype(str).unique().tolist())
            selected_cats = st.multiselect("Category", cats, default=cats)
            data = data[data["Category"].isin(selected_cats)]

        for _, row in data.iterrows():
            p = str(row.get("Priority", ""))
            icon = "🔴" if p == "High" else "🟠" if p == "Medium" else "🟢"
            with st.expander(f"{icon} {row.get('Category', 'Insight')} — {row.get('Metric', '')}", expanded=(p == "High")):
                st.markdown(f"**Finding:** {row.get('Finding', '')}")
                st.markdown(f"**UX interpretation:** {row.get('UX Interpretation', '')}")
                st.markdown(f"**Recommended action:** {row.get('Potential Recommendation', '')}")
                if pd.notna(row.get("Evidence", np.nan)):
                    st.caption(f"Evidence: {row.get('Evidence')}")

# -----------------------------
# RAW DATA
# -----------------------------
elif page == "Raw Data":
    st.markdown('<div class="section-title">Workbook explorer</div>', unsafe_allow_html=True)
    selected_sheet = st.selectbox("Sheet", list(sheets.keys()))
    raw = sheets[selected_sheet]
    st.caption(f"{raw.shape[0]:,} rows × {raw.shape[1]:,} columns")
    st.dataframe(raw, use_container_width=True, hide_index=True)

    # Automatic visual for any numeric raw sheet
    numeric_cols = raw.select_dtypes(include=np.number).columns.tolist()
    text_cols = raw.select_dtypes(include=["object", "string"]).columns.tolist()
    if numeric_cols:
        st.markdown('<div class="section-title">Quick chart for selected sheet</div>', unsafe_allow_html=True)
        metric_col = st.selectbox("Numeric metric", numeric_cols, key=f"raw_metric_{selected_sheet}")
        if text_cols:
            cat_col = st.selectbox("Category", text_cols, key=f"raw_cat_{selected_sheet}")
            chart_df = raw[[cat_col, metric_col]].copy().dropna().head(40)
            chart_df[metric_col] = pd.to_numeric(chart_df[metric_col], errors="coerce")
            chart_df = chart_df.dropna().sort_values(metric_col, ascending=True)
            if len(chart_df):
                fig = px.bar(chart_df, x=metric_col, y=cat_col, orientation="h", title=f"{metric_col} by {cat_col}", text=metric_col)
                fig.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(style_fig(fig, 500), use_container_width=True)
        else:
            chart_df = raw[[metric_col]].dropna().reset_index().rename(columns={"index": "Row"})
            fig = px.line(chart_df.head(100), x="Row", y=metric_col, markers=True, title=f"{metric_col} across rows")
            st.plotly_chart(style_fig(fig, 420), use_container_width=True)

    csv = raw.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download selected sheet as CSV",
        data=csv,
        file_name=f"{selected_sheet}.csv",
        mime="text/csv",
    )

st.divider()
st.caption("FITX UX Analytics • Streamlit dashboard • Source metrics are read directly from the supplied Excel workbook.")
