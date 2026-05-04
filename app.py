import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from collections import Counter
import re
import os
import io

# =====================================================================
# CONFIGURATION
# =====================================================================
FILE_PATH = "data.xlsx"

BRAND_BLUE      = "#1D2E6F"
BRAND_RED       = "#E15049"
BRAND_WHITE     = "#FFFFFF"
BRAND_LIGHT_BLUE = "#E8ECF7"
BRAND_LIGHT_RED  = "#FDECEA"

COLOR_SEQ = [
    BRAND_BLUE, BRAND_RED,
    "#4A6FA5", "#E8897A",
    "#2E4B8F", "#C94A42",
    "#7B9CC9", "#F0B0AA",
    "#1A5276", "#922B21",
]

COLOR_CAT = [
    "#1D2E6F",
    "#E15049",
    "#27AE60",
    "#F39C12",
    "#8E44AD",
    "#16A085",
    "#E67E22",
    "#3498DB",
    "#C0392B",
    "#F1C40F",
    "#34495E",
    "#7F8C8D",
]

# =====================================================================
# COLUMN NAME PATTERNS (substring match against Excel header)
# =====================================================================
COL_PATTERNS = {
    "source":                    "מקור הקובץ",
    "timestamp":                 "Timestamp",
    "respondent":                "פרטי ממלא",
    "marital":                   "מצב משפחתי",
    "num_children":              "מספר ילדים",
    "children_ages":             "גילאי הילדים",
    "employment":                "מצב תעסוקתי",
    "religious":                 "שיוך דתי",
    "reserve_days":              "כמות ימי מילואים",
    "fatigue":                   "עייפות ושחיקה",
    "worry":                     "דאגה מתמשכת",
    "loneliness":                "תחושת בדידות",
    "stress":                    "לחץ ומתח",
    "difficulty_active":         "בסבב מילואים פעיל",
    "difficulty_between":        "בתקופה שבין הסבבים",
    "difficulty_near":           "סמוך ליציאה לסבב",
    "difficulty_constant":       "נשארת קבועה",
    "help_source":               "למי את/ה פונה",
    "community":                 "הקהילה הקרובה מודעת",
    "economic_concern":          "המצב הכלכלי מעסיק",
    "income_drop":               "ירידה בפרנסה",
    "main_difficulty":           "הקושי המרכזי עבורך",
    "elaboration":               "מוזמנים לפרט",
    "relationship":              "מתאר/ת את הזוגיות",
    "relationship_difficulties": "קשיים בקשר הזוגי",
    "daily_challenge":           "האתגר המרכזי שאת/ה חווה",
    "parenting":                 "קשה לך להתמודד כהורה",
    "child_behavior":            "דפוסי התנהגות",
    "education_change":          "שינוי לרעה בהתנהגות הילד",
    "education_difficulties":    "קושי מול המסגרות",
    "support_areas":             "תחומים הבאים היית רוצה",
    "free_response":             "מה הכי מקשה עלייך",
    "suggestions":               "מענים ותוכניות נוספות",
}

LABELS = {
    "source":                   "רשות מקומית",
    "timestamp":                "תאריך",
    "respondent":               "פרטי ממלא",
    "marital":                  "מצב משפחתי",
    "num_children":             "מספר ילדים",
    "children_ages":            "גילאי ילדים",
    "employment":               "מצב תעסוקתי",
    "religious":                "שיוך דתי",
    "reserve_days":             "ימי מילואים",
    "fatigue":                  "עייפות ושחיקה",
    "worry":                    "דאגה מתמשכת",
    "loneliness":               "תחושת בדידות",
    "stress":                   "לחץ ומתח נפשי",
    "difficulty_active":        "קושי בסבב פעיל",
    "difficulty_between":       "קושי בין סבבים",
    "difficulty_near":          "קושי סמוך לסבב",
    "difficulty_constant":      "קושי קבוע לאורך הזמן",
    "help_source":              "מקורות עזרה",
    "community":                "מודעות קהילה",
    "economic_concern":         "עיסוק כלכלי",
    "income_drop":              "ירידה בפרנסה",
    "main_difficulty":          "קושי כלכלי מרכזי",
    "elaboration":              "פירוט קושי",
    "relationship":             "מצב הזוגיות",
    "relationship_difficulties":"קשיים זוגיים",
    "daily_challenge":          "אתגר תפקוד יומיומי",
    "parenting":                "קושי כהורה",
    "child_behavior":           "דפוסי התנהגות ילדים",
    "education_change":         "שינוי במסגרות חינוך",
    "education_difficulties":   "קשיים במסגרות חינוכיות",
    "support_areas":            "תחומי תמיכה רצויים",
    "free_response":            "התייחסות חופשית",
    "suggestions":              "הצעות ומענים",
}

RATING_COLS = [
    "fatigue", "worry", "loneliness", "stress",
    "difficulty_active", "difficulty_between", "difficulty_near", "difficulty_constant",
    "community", "economic_concern", "parenting",
]
MULTI_CHOICE_COLS  = ["help_source", "child_behavior", "education_difficulties", "support_areas"]
CATEGORICAL_COLS   = ["income_drop", "main_difficulty", "relationship", "relationship_difficulties", "education_change"]
FREE_TEXT_COLS     = ["elaboration", "daily_challenge", "free_response", "suggestions"]
FILTER_COLS        = ["respondent", "marital", "num_children", "children_ages", "employment", "religious", "reserve_days"]

HEBREW_STOPWORDS = {
    'של','את','על','עם','לא','זה','כי','אני','הוא','היא','הם','לי','יש','כן',
    'כל','אבל','גם','יותר','כך','מה','אם','בתקופה','הזו','מאוד','רק','אחד',
    'כבר','עוד','היה','הייתי','הייתה','להיות','יכול','יכולה','צריך','צריכה',
    'מהם','בהם','להם','שלי','שלנו','שלהם','אנחנו','אתם','הן','ב','ל','מ',
    'כ','ה','ו','ש','מי','איפה','מתי','איך','למה','ולא','שלא','אין','בלי',
    'אחרי','לפני','בזמן','הילדים','ממש','בגלל','כדי','שבו','שבה','אשר',
    'כאשר','עצמי','עצמנו','שכן','משום','היות','עקב','לפיכך','לכן','אולם',
    'אך','ברם','לעומת','בעוד','שתי','שני','מספר','כמה','הרבה','מעט','קצת',
    'הכי','ביותר','אלו','אלה','אותו','אותה','אותם','אותן','זאת','זו','זוהי',
    'הרי','הנה','אפילו','עדיין','תמיד','לעולם','לפעמים','היי','ידי','תוך',
    'כלל','הדבר','דבר','היום','בו','בה','עליו','עליה','אליו','אליה','לו','לה',
}

# =====================================================================
# PAGE CONFIG
# =====================================================================
st.set_page_config(
    page_title="דשבורד שאלוני מילואים",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =====================================================================
# CSS
# =====================================================================
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&display=swap');

    html, body, [class*="css"], .stApp {{
        font-family: 'Heebo', Arial, sans-serif !important;
        direction: rtl;
        font-size: 16px;
    }}

    .main .block-container {{
        direction: rtl;
        text-align: right;
        padding-top: 1rem;
        max-width: 1500px;
    }}

    h1, h2, h3, h4, h5, h6, p, div, span, label {{
        direction: rtl !important;
        text-align: right !important;
    }}

    [data-testid="stSidebar"] {{
        background-color: {BRAND_LIGHT_BLUE};
        direction: rtl;
    }}
    [data-testid="stSidebar"] * {{
        direction: rtl !important;
        text-align: right !important;
    }}
    [data-testid="stSidebar"] .stRadio > div {{
        flex-direction: column;
    }}

    .logo-header {{
        background: linear-gradient(135deg, {BRAND_BLUE}, #2E4B8F);
        padding: 1.6rem 2rem;
        border-radius: 14px;
        color: white;
        margin-bottom: 1.8rem;
        direction: rtl;
        box-shadow: 0 4px 16px rgba(29,46,111,0.18);
    }}
    .logo-header h1 {{
        color: white !important;
        margin: 0 0 0.3rem 0;
        font-size: 2rem;
        font-weight: 800;
    }}
    .logo-header p {{
        color: {BRAND_LIGHT_BLUE};
        margin: 0;
        font-size: 1.05rem;
    }}

    .metric-card {{
        background: white;
        border-radius: 14px;
        padding: 1.4rem 1rem;
        box-shadow: 0 2px 12px rgba(29,46,111,0.10);
        border-right: 5px solid {BRAND_BLUE};
        text-align: center;
        direction: rtl;
        height: 100%;
    }}
    .metric-card.red  {{ border-right-color: {BRAND_RED}; }}
    .metric-card.green {{ border-right-color: #27AE60; }}

    .metric-value {{
        font-size: 2.6rem;
        font-weight: 900;
        color: {BRAND_BLUE};
        margin: 0;
        line-height: 1.1;
    }}
    .metric-value.red   {{ color: {BRAND_RED}; }}
    .metric-value.green {{ color: #27AE60; }}

    .metric-label {{
        font-size: 1rem;
        color: #444;
        margin: 0.5rem 0 0 0;
        font-weight: 600;
    }}
    .metric-sub {{
        font-size: 0.85rem;
        color: #888;
        margin: 0.2rem 0 0 0;
    }}

    .section-header {{
        background: linear-gradient(135deg, {BRAND_BLUE}, #2E4B8F);
        color: white !important;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 2rem 0 1.2rem 0;
        font-weight: 700;
        font-size: 1.3rem;
        direction: rtl;
        text-align: right !important;
        box-shadow: 0 2px 10px rgba(29,46,111,0.15);
    }}

    .insight-box {{
        background: {BRAND_LIGHT_RED};
        border-right: 5px solid {BRAND_RED};
        padding: 1.1rem 1.3rem;
        border-radius: 10px;
        margin: 0.7rem 0;
        direction: rtl;
        font-size: 1.05rem;
        line-height: 1.5;
    }}
    .insight-box.blue {{
        background: {BRAND_LIGHT_BLUE};
        border-right-color: {BRAND_BLUE};
    }}
    .insight-box.green {{
        background: #EAFAF1;
        border-right-color: #27AE60;
    }}

    .quote-box {{
        background: {BRAND_LIGHT_BLUE};
        border-right: 5px solid {BRAND_BLUE};
        padding: 1rem 1.3rem;
        margin: 0.7rem 0;
        border-radius: 10px;
        font-size: 1.1rem;
        line-height: 1.7;
        direction: rtl;
        color: #222;
        font-style: italic;
    }}

    /* Plotly chart container — clean white frame for screenshot */
    [data-testid="stPlotlyChart"] {{
        background: white;
        border-radius: 12px;
        padding: 0.6rem;
        box-shadow: 0 2px 12px rgba(29,46,111,0.08);
        margin-bottom: 1rem;
    }}

    /* Fix multiselect RTL */
    .stMultiSelect [data-baseweb="tag"] {{
        direction: rtl;
    }}
    [data-testid="stWidgetLabel"] {{
        direction: rtl !important;
        text-align: right !important;
    }}

    /* Filter count badge */
    .filter-badge {{
        background: {BRAND_RED};
        color: white;
        border-radius: 20px;
        padding: 0.15rem 0.6rem;
        font-size: 0.78rem;
        font-weight: 600;
    }}
</style>
""", unsafe_allow_html=True)

# =====================================================================
# DATA LOADING
# =====================================================================
def _find_col(raw_cols: list, pattern: str) -> str | None:
    """Returns the first column name that contains the given substring."""
    for col in raw_cols:
        if pattern in str(col):
            return col
    return None


def _process_df(df: pd.DataFrame) -> pd.DataFrame:
    """Column rename + type coercion. Shared by upload and default loaders."""
    raw_cols = list(df.columns)
    rename_map, not_found = {}, []
    for key, pattern in COL_PATTERNS.items():
        matched = _find_col(raw_cols, pattern)
        if matched:
            rename_map[matched] = key
        else:
            not_found.append(f"«{pattern}»")
    df = df.rename(columns=rename_map)
    if not_found:
        st.warning(f"⚠️ עמודות שלא נמצאו בקובץ: {', '.join(not_found)}")

    if "timestamp" in df.columns:
        df["timestamp"]   = pd.to_datetime(df["timestamp"], errors="coerce", dayfirst=True)
        df["month_str"]   = df["timestamp"].dt.strftime("%Y-%m")
        df["month_label"] = df["timestamp"].dt.strftime("%m/%Y")

    for col in RATING_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


@st.cache_data(show_spinner="טוען את הנתונים...")
def load_data_from_bytes(file_bytes: bytes, file_name: str = "") -> pd.DataFrame | None:
    """Load data from raw bytes. Caches by file content hash, so a different
    upload invalidates automatically."""
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    except Exception as e:
        st.error(f"שגיאה בטעינת הקובץ: {e}")
        return None
    return _process_df(df)


def load_default() -> pd.DataFrame | None:
    """Load the bundled fallback Excel from disk (for local + cloud first-run)."""
    try:
        with open(FILE_PATH, "rb") as f:
            return load_data_from_bytes(f.read(), FILE_PATH)
    except FileNotFoundError:
        return None


# =====================================================================
# FILTERS
# =====================================================================
def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    out = df.copy()

    def _filter(col, selected):
        if selected and "הכל" not in selected and col in out.columns:
            return out[out[col].astype(str).isin(selected)]
        return out

    out = _filter("source", filters.get("source"))
    if filters.get("months") and "הכל" not in filters["months"] and "month_str" in out.columns:
        out = out[out["month_str"].isin(filters["months"])]
    for col in FILTER_COLS:
        out = _filter(col, filters.get(col))

    return out


# =====================================================================
# HELPERS
# =====================================================================
def pct_ge4(series: pd.Series) -> float:
    s = series.dropna()
    if len(s) == 0:
        return 0.0
    return (s >= 4).sum() / len(s) * 100


def parse_multi(series: pd.Series) -> list[str]:
    items = []
    for val in series.dropna().astype(str):
        items.extend([v.strip() for v in val.split(",") if v.strip() and v.strip() != "nan"])
    return items


def base_layout() -> dict:
    return dict(
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Heebo, Arial", size=16, color="#222"),
        title=dict(
            font=dict(family="Heebo, Arial", size=22, color=BRAND_BLUE),
            x=0.5, xanchor="center",
            pad=dict(t=10, b=20),
        ),
        margin=dict(t=80, b=70, l=50, r=50),
        height=480,
        hoverlabel=dict(font_size=14, font_family="Heebo"),
    )


def _truncate(s: str, n: int = 35) -> str:
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


# =====================================================================
# CHART RENDERERS
# =====================================================================

def render_rating_chart(df: pd.DataFrame, col_key: str, title: str | None = None):
    if col_key not in df.columns:
        return
    label = title or LABELS.get(col_key, col_key)
    data = df[col_key].dropna()
    if data.empty:
        st.info(f"אין נתונים: {label}")
        return

    total  = len(data)
    avg    = data.mean()
    pct4   = pct_ge4(data)
    median = data.median()

    st.markdown(f'<div class="section-header">📊 {label}</div>', unsafe_allow_html=True)

    # Metric row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <p class="metric-value">{total:,}</p>
            <p class="metric-label">סך משיבים</p>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <p class="metric-value">{avg:.2f}</p>
            <p class="metric-label">ממוצע (מתוך 5)</p>
        </div>""", unsafe_allow_html=True)
    with c3:
        color = "red" if pct4 > 60 else ("green" if pct4 < 30 else "")
        st.markdown(f"""
        <div class="metric-card {color}">
            <p class="metric-value {color}">{pct4:.1f}%</p>
            <p class="metric-label">ציינו 4 או 5</p>
            <p class="metric-sub">(רמה גבוהה)</p>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="metric-card">
            <p class="metric-value">{median:.0f}</p>
            <p class="metric-label">חציון</p>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

    # Distribution — single full-width bar chart
    dist = (
        data.value_counts()
            .reindex([1, 2, 3, 4, 5], fill_value=0)
            .reset_index()
    )
    dist.columns = ["ציון", "מספר"]
    dist["אחוז"] = (dist["מספר"] / total * 100).round(1)
    dist["ציון"] = dist["ציון"].astype(str)
    dist["תווית"] = dist.apply(lambda r: f"{int(r['מספר']):,}<br>({r['אחוז']:.1f}%)", axis=1)

    fig = px.bar(
        dist, x="ציון", y="מספר",
        text="תווית",
        color="ציון",
        color_discrete_sequence=COLOR_SEQ,
        title="התפלגות הציונים",
    )
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=15, family="Heebo", color="#222"),
        cliponaxis=False,
    )
    fig.update_layout(**base_layout())
    fig.update_layout(
        showlegend=False,
        xaxis=dict(title=dict(text="ציון (1=נמוך, 5=גבוה)", font=dict(size=16)),
                   tickfont=dict(size=16)),
        yaxis=dict(title=dict(text="מספר משיבים", font=dict(size=16)),
                   tickfont=dict(size=14)),
        bargap=0.28,
        height=480,
    )
    # Padding above bars so labels don't clip
    y_max = dist["מספר"].max()
    fig.update_yaxes(range=[0, y_max * 1.18])
    st.plotly_chart(fig, use_container_width=True)


def render_multichoice_chart(df: pd.DataFrame, col_key: str, title: str | None = None):
    if col_key not in df.columns:
        return
    label = title or LABELS.get(col_key, col_key)
    items = parse_multi(df[col_key])
    if not items:
        st.info(f"אין נתונים: {label}")
        return

    n_respondents = df[col_key].notna().sum()
    freq = pd.DataFrame(Counter(items).most_common(), columns=["תשובה", "בחירות"])
    freq["אחוז ממשיבים"] = (freq["בחירות"] / n_respondents * 100).round(1)

    st.markdown(f'<div class="section-header">📋 {label}</div>', unsafe_allow_html=True)

    freq_top = freq.head(12).copy()
    freq_top["תווית_ערך"] = freq_top.apply(
        lambda r: f"{int(r['בחירות']):,} ({r['אחוז ממשיבים']:.1f}%)", axis=1
    )
    freq_top["index_str"] = ["%02d" % (i + 1) for i in range(len(freq_top))]

    n_bars = len(freq_top)
    chart_h = max(540, min(820, 50 * n_bars + 200))

    fig = px.bar(
        freq_top,
        y="index_str", x="בחירות",
        orientation="h",
        color="תשובה",
        color_discrete_sequence=COLOR_CAT,
        text="תווית_ערך",
        title=f"תדירות בחירה — {label}",
    )
    fig.update_traces(
        textposition="inside",
        textfont=dict(size=14, family="Heebo", color="white", weight="bold"),
        insidetextanchor="end",
        cliponaxis=False,
        hovertemplate="<b>%{fullData.name}</b><br>בחירות: %{x:,}<extra></extra>",
    )
    fig.update_layout(**base_layout())
    fig.update_layout(
        height=chart_h,
        yaxis=dict(autorange="reversed", title="", showticklabels=False, showgrid=False),
        xaxis=dict(title=dict(text="מספר בחירות", font=dict(size=16)), tickfont=dict(size=14)),
        margin=dict(t=80, b=70, l=40, r=320),
        bargap=0.32,
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="middle", y=0.5,
            xanchor="left", x=1.02,
            font=dict(size=13, family="Heebo"),
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="#ddd",
            borderwidth=1,
            itemsizing="constant",
            traceorder="normal",
        ),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("📋 טבלת נתונים מלאה"):
        st.dataframe(freq[["תשובה", "בחירות", "אחוז ממשיבים"]],
                     use_container_width=True, hide_index=True)


def render_categorical_chart(df: pd.DataFrame, col_key: str, title: str | None = None):
    if col_key not in df.columns:
        return
    label = title or LABELS.get(col_key, col_key)
    data = df[col_key].dropna().astype(str)
    data = data[~data.isin(["nan", ""])]
    if data.empty:
        st.info(f"אין נתונים: {label}")
        return

    freq = data.value_counts().reset_index()
    freq.columns = ["תשובה", "מספר"]
    freq["אחוז"] = (freq["מספר"] / len(data) * 100).round(1)

    st.markdown(f'<div class="section-header">📌 {label}</div>', unsafe_allow_html=True)

    n_unique = freq.shape[0]

    if n_unique <= 5:
        # Donut chart — large, full-width, external labels
        fig = px.pie(
            freq, names="תשובה", values="מספר",
            color_discrete_sequence=COLOR_SEQ,
            title=label, hole=0.38,
        )
        fig.update_traces(
            textposition="outside",
            textinfo="label+percent",
            textfont=dict(size=16, family="Heebo", color="#222"),
            marker=dict(line=dict(color="white", width=2)),
            pull=[0.03] * n_unique,
        )
        fig.update_layout(**base_layout())
        fig.update_layout(
            height=540,
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5,
                        xanchor="left", x=1.02,
                        font=dict(size=14, family="Heebo")),
            margin=dict(t=80, b=70, l=40, r=240),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        # Horizontal bar — categorical colors, legend on side, no Y-axis text
        freq_top = freq.head(12).copy()
        freq_top["תווית_ערך"] = freq_top.apply(
            lambda r: f"{int(r['מספר']):,} ({r['אחוז']:.1f}%)", axis=1
        )
        freq_top["index_str"] = ["%02d" % (i + 1) for i in range(len(freq_top))]
        n_bars = len(freq_top)
        chart_h = max(540, min(820, 50 * n_bars + 200))

        fig = px.bar(
            freq_top,
            y="index_str", x="מספר",
            orientation="h",
            color="תשובה",
            color_discrete_sequence=COLOR_CAT,
            text="תווית_ערך",
            title=f"התפלגות התשובות — {label}",
        )
        fig.update_traces(
            textposition="inside",
            textfont=dict(size=14, family="Heebo", color="white", weight="bold"),
            insidetextanchor="end",
            cliponaxis=False,
            hovertemplate="<b>%{fullData.name}</b><br>מספר: %{x:,}<extra></extra>",
        )
        fig.update_layout(**base_layout())
        fig.update_layout(
            height=chart_h,
            yaxis=dict(autorange="reversed", title="", showticklabels=False, showgrid=False),
            xaxis=dict(title=dict(text="מספר משיבים", font=dict(size=16)), tickfont=dict(size=14)),
            margin=dict(t=80, b=70, l=40, r=320),
            bargap=0.32,
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="middle", y=0.5,
                xanchor="left", x=1.02,
                font=dict(size=13, family="Heebo"),
                bgcolor="rgba(255,255,255,0.85)",
                bordercolor="#ddd",
                borderwidth=1,
                itemsizing="constant",
                traceorder="normal",
            ),
        )
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("📋 טבלת נתונים מלאה"):
            st.dataframe(freq[["תשובה", "מספר", "אחוז"]],
                         use_container_width=True, hide_index=True)


def _extract_ngrams(texts: list, n: int) -> Counter:
    """Returns a Counter of n-gram phrases (Hebrew tokens, stopwords removed)."""
    counter: Counter = Counter()
    for t in texts:
        clean = re.sub(r"[^\u0590-\u05FF\s]", " ", str(t))
        words = [w for w in clean.split() if len(w) > 1 and w not in HEBREW_STOPWORDS]
        for i in range(len(words) - n + 1):
            counter[" ".join(words[i:i + n])] += 1
    return counter


def render_freetext_summary(df: pd.DataFrame, col_key: str, title: str | None = None):
    if col_key not in df.columns:
        return
    label = title or LABELS.get(col_key, col_key)
    series = df[col_key].dropna().astype(str)
    series = series[~series.isin(["nan", ""])]
    data = series.tolist()
    if not data:
        st.info(f"אין נתונים: {label}")
        return

    st.markdown(f'<div class="section-header">💬 {label}</div>', unsafe_allow_html=True)

    # KPIs
    n_resp = len(data)
    pct = n_resp / len(df) * 100 if len(df) else 0
    word_lens = [len(t.split()) for t in data]
    avg_len = float(np.mean(word_lens)) if word_lens else 0

    c1, c2, c3 = st.columns(3)
    c1.markdown(
        f'''<div class="metric-card">
        <p class="metric-value">{n_resp:,}</p>
        <p class="metric-label">תשובות התקבלו</p>
        </div>''', unsafe_allow_html=True
    )
    c2.markdown(
        f'''<div class="metric-card">
        <p class="metric-value">{pct:.0f}%</p>
        <p class="metric-label">מהמשיבים ענו</p>
        </div>''', unsafe_allow_html=True
    )
    c3.markdown(
        f'''<div class="metric-card">
        <p class="metric-value">{avg_len:.0f}</p>
        <p class="metric-label">מילים בממוצע</p>
        </div>''', unsafe_allow_html=True
    )
    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

    # Phrase extraction (bigrams + trigrams; trigrams weighted higher)
    big = _extract_ngrams(data, 2)
    tri = _extract_ngrams(data, 3)
    combined: Counter = Counter()
    for k, v in big.items():
        combined[k] += v
    for k, v in tri.items():
        combined[k] += int(round(v * 1.3))

    top = combined.most_common(15)
    if top:
        ph_df = pd.DataFrame(top, columns=["ביטוי", "מספר הופעות"])
        ph_df["מספר הופעות"] = ph_df["מספר הופעות"].astype(int)
        n_bars = len(ph_df)
        chart_h = max(520, min(900, 55 * n_bars + 160))

        fig = px.bar(
            ph_df, y="ביטוי", x="מספר הופעות",
            orientation="h",
            text="מספר הופעות",
            color_discrete_sequence=[BRAND_BLUE],
            title=f"ביטויים חוזרים — {label}",
        )
        fig.update_traces(
            textposition="outside",
            textfont=dict(size=15, family="Heebo", color=BRAND_BLUE),
            cliponaxis=False,
        )
        fig.update_layout(**base_layout())
        fig.update_layout(
            height=chart_h,
            yaxis=dict(autorange="reversed", title="",
                       tickfont=dict(size=15, family="Heebo"), automargin=True),
            xaxis=dict(title=dict(text="מספר הופעות", font=dict(size=16)),
                       tickfont=dict(size=14)),
            margin=dict(t=80, b=70, l=260, r=100),
            bargap=0.32,
        )
        x_max = ph_df["מספר הופעות"].max()
        fig.update_xaxes(range=[0, x_max * 1.18])
        st.plotly_chart(fig, use_container_width=True)

    # Representative quotes
    top_phrases = [p for p, _ in top[:10]]
    candidates = []
    for t in data:
        n_words = len(t.split())
        if 8 <= n_words <= 28:
            hits = sum(1 for ph in top_phrases if ph in t)
            if hits >= 1:
                candidates.append((hits, t))
    candidates.sort(key=lambda x: -x[0])

    chosen, seen_phrase = [], set()
    for _, t in candidates:
        first = next((ph for ph in top_phrases if ph in t), None)
        if first and first not in seen_phrase:
            chosen.append(t)
            seen_phrase.add(first)
        if len(chosen) >= 5:
            break

    if chosen:
        st.markdown(
            '<div class="section-header" style="background:linear-gradient(135deg,#4A6FA5,#2E4B8F);">'
            '📣 ציטוטים מייצגים</div>',
            unsafe_allow_html=True,
        )
        for q in chosen:
            q_safe = q.replace("<", "&lt;").replace(">", "&gt;")
            st.markdown(
                f'<div class="quote-box">"{q_safe}"</div>',
                unsafe_allow_html=True,
            )




# =====================================================================
# HERO AGGREGATE CHART RENDERERS
# =====================================================================

def render_rating_hero(df: pd.DataFrame, cols: list, hero_title: str = "תמונת מצב"):
    """Vertical bar chart — % scored 4-5 for each col in cols."""
    rows = []
    for c in cols:
        if c in df.columns:
            s = df[c].dropna()
            if len(s):
                rows.append({"מדד": LABELS.get(c, c), "אחוז": pct_ge4(s)})
    if not rows:
        return
    h_df = pd.DataFrame(rows)
    h_df["תווית"] = h_df["אחוז"].map(lambda v: f"{v:.0f}%")

    st.markdown(f'<div class="section-header">{hero_title}</div>', unsafe_allow_html=True)
    fig = px.bar(
        h_df, x="מדד", y="אחוז",
        color="מדד",
        color_discrete_sequence=COLOR_CAT,
        text="תווית",
        title="% שציינו רמה גבוהה (4 או 5)",
    )
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=18, family="Heebo", weight="bold"),
        cliponaxis=False,
    )
    fig.update_layout(**base_layout())
    y_top = max(h_df["אחוז"]) * 1.22 if len(h_df) else 100
    fig.update_layout(
        height=480,
        showlegend=False,
        xaxis=dict(title="", tickfont=dict(size=15, family="Heebo")),
        yaxis=dict(
            title=dict(text="% שציינו 4 או 5", font=dict(size=16)),
            range=[0, min(y_top, 105)],
            tickfont=dict(size=14),
            ticksuffix="%",
        ),
        bargap=0.4,
        margin=dict(t=80, b=70, l=70, r=50),
    )
    st.plotly_chart(fig, use_container_width=True)


def render_mixed_hero(df: pd.DataFrame, items: list, hero_title: str = "תמונת מצב"):
    """Vertical bar chart — each item is {label, value} where value is already a percentage."""
    if not items:
        return
    h_df = pd.DataFrame(items)
    h_df["תווית"] = h_df["value"].map(lambda v: f"{v:.0f}%")
    st.markdown(f'<div class="section-header">{hero_title}</div>', unsafe_allow_html=True)
    fig = px.bar(
        h_df, x="label", y="value",
        color="label",
        color_discrete_sequence=COLOR_CAT,
        text="תווית",
        title="מדדי מצב מרכזיים",
    )
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=18, family="Heebo", weight="bold"),
        cliponaxis=False,
    )
    fig.update_layout(**base_layout())
    y_top = max(h_df["value"]) * 1.22 if len(h_df) else 100
    fig.update_layout(
        height=480,
        showlegend=False,
        xaxis=dict(title="", tickfont=dict(size=15, family="Heebo")),
        yaxis=dict(
            title=dict(text="אחוז", font=dict(size=16)),
            range=[0, min(y_top, 105)],
            tickfont=dict(size=14),
            ticksuffix="%",
        ),
        bargap=0.4,
        margin=dict(t=80, b=70, l=70, r=50),
    )
    st.plotly_chart(fig, use_container_width=True)


def render_multichoice_hero(df: pd.DataFrame, col_key: str, hero_title: str = "תמונת מצב", top_n: int = 5):
    """Horizontal bar chart hero — top-N answers for a multi-choice column."""
    if col_key not in df.columns:
        return
    items = parse_multi(df[col_key])
    if not items:
        return
    n_resp = df[col_key].notna().sum()
    top = Counter(items).most_common(top_n)
    h_df = pd.DataFrame(top, columns=["תשובה", "בחירות"])
    h_df["אחוז"] = (h_df["בחירות"] / n_resp * 100).round(1)
    h_df["תווית"] = h_df.apply(lambda r: f"{int(r['בחירות']):,} ({r['אחוז']:.1f}%)", axis=1)
    h_df["index_str"] = ["%02d" % (i + 1) for i in range(len(h_df))]

    st.markdown(f'<div class="section-header">{hero_title}</div>', unsafe_allow_html=True)
    fig = px.bar(
        h_df, y="index_str", x="בחירות",
        orientation="h",
        color="תשובה",
        color_discrete_sequence=COLOR_CAT,
        text="תווית",
        title=f"Top {top_n} תשובות נפוצות",
    )
    fig.update_traces(
        textposition="inside",
        textfont=dict(size=14, family="Heebo", color="white", weight="bold"),
        insidetextanchor="end",
        cliponaxis=False,
        hovertemplate="<b>%{fullData.name}</b><br>בחירות: %{x:,}<extra></extra>",
    )
    fig.update_layout(**base_layout())
    fig.update_layout(
        height=440,
        bargap=0.32,
        yaxis=dict(autorange="reversed", title="", showticklabels=False, showgrid=False),
        xaxis=dict(title=dict(text="מספר בחירות", font=dict(size=16)), tickfont=dict(size=14)),
        margin=dict(t=80, b=70, l=40, r=320),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="middle", y=0.5,
            xanchor="left", x=1.02,
            font=dict(size=13, family="Heebo"),
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="#ddd",
            borderwidth=1,
            itemsizing="constant",
        ),
    )
    st.plotly_chart(fig, use_container_width=True)

# =====================================================================
# SIDEBAR
# =====================================================================
def render_sidebar(df: pd.DataFrame) -> tuple[dict, str]:
    with st.sidebar:
        st.markdown(f"""
        <div style="text-align:center; padding:0.8rem 0 0.5rem;">
            <h2 style="color:{BRAND_BLUE};margin:0;font-size:1.2rem;">🎖️ דשבורד מילואים</h2>
            <p style="color:#666;font-size:0.8rem;margin:0.2rem 0 0;">ניתוח שאלוני רווחה</p>
        </div>
        <hr style="margin:0.6rem 0">
        """, unsafe_allow_html=True)

        st.markdown("**📁 קובץ נתונים**")
        uploaded = st.file_uploader(
            "העלה קובץ Excel חדש",
            type=["xlsx", "xls"],
            help="הקובץ נטען זמנית לסשן הנוכחי. רענון הדף יחזיר לקובץ ברירת המחדל.",
            label_visibility="collapsed",
            key="data_uploader",
        )
        if uploaded is not None:
            st.session_state["uploaded_bytes"] = uploaded.getvalue()
            st.session_state["uploaded_name"]  = uploaded.name
        if "uploaded_name" in st.session_state:
            st.success(f"✅ {st.session_state['uploaded_name']}")
            if st.button("חזור לקובץ ברירת מחדל"):
                del st.session_state["uploaded_bytes"]
                del st.session_state["uploaded_name"]
                st.rerun()
        else:
            st.caption(f"📄 ברירת מחדל: {FILE_PATH}")
        st.markdown('<hr style="margin:0.6rem 0">', unsafe_allow_html=True)

        st.markdown("**ניווט**")
        page = st.radio(
            "page",
            options=[
                "📊 סיכום ראשי",
                "😓 רגשות ורווחה",
                "🎖️ השפעת המילואים",
                "💰 קהילה וכלכלה",
                "💑 זוגיות ומשפחה",
                "👶 ילדים וחינוך",
                "🤝 תמיכה ומענה",
                "✍️ מענה חופשי",
            ],
            label_visibility="collapsed",
        )

        st.markdown('<hr style="margin:0.6rem 0">', unsafe_allow_html=True)
        st.markdown("**סינון נתונים**")

        filters: dict = {}

        # Source
        if "source" in df.columns:
            opts = sorted(df["source"].dropna().astype(str).unique().tolist())
            filters["source"] = st.multiselect("רשות מקומית", ["הכל"] + opts, default=["הכל"])

        # Months
        if "month_str" in df.columns:
            months_raw = sorted(df["month_str"].dropna().unique().tolist())
            months_disp = []
            for m in months_raw:
                try:
                    months_disp.append(pd.Period(m, "M").strftime("%m/%Y"))
                except Exception:
                    months_disp.append(m)
            month_map = dict(zip(months_disp, months_raw))
            sel_disp = st.multiselect("חודש מענה", ["הכל"] + months_disp, default=["הכל"])
            if "הכל" in sel_disp:
                filters["months"] = ["הכל"]
            else:
                filters["months"] = [month_map[d] for d in sel_disp if d in month_map]

        st.markdown("**פרטי המשיב**")
        for col in FILTER_COLS:
            if col in df.columns:
                opts = sorted(df[col].dropna().astype(str).unique().tolist())
                opts = [o for o in opts if o not in ("nan", "")]
                if opts:
                    filters[col] = st.multiselect(LABELS.get(col, col), ["הכל"] + opts, default=["הכל"])

        st.markdown('<hr style="margin:0.6rem 0">', unsafe_allow_html=True)
        st.markdown(f"**סך משיבים בקובץ:** {len(df):,}")

    return filters, page


# =====================================================================
# PAGES
# =====================================================================

def render_summary_page(df: pd.DataFrame, df_raw: pd.DataFrame):
    n = len(df)
    st.markdown(f"""
    <div class="logo-header">
        <h1>📊 דשבורד שאלוני רווחה — משפחות מילואים</h1>
        <p>סיכום ראשי | מציג {n:,} משיבים מתוך {len(df_raw):,}</p>
    </div>
    """, unsafe_allow_html=True)

    # ── KPI cards ──
    st.markdown("### מדדי מפתח — ממוצעי דירוג (1–5)")
    kpi_cols = ["fatigue", "worry", "loneliness", "stress", "community", "economic_concern", "parenting"]
    kpi_icons = ["😓", "😰", "😶", "😤", "🏘️", "💸", "👨‍👩‍👧"]
    cols = st.columns(len(kpi_cols))
    for i, (ck, icon) in enumerate(zip(kpi_cols, kpi_icons)):
        if ck in df.columns:
            s = df[ck].dropna()
            avg = s.mean() if len(s) else 0
            p4  = pct_ge4(s)
            color = "red" if avg >= 3.5 else ("green" if avg < 2.5 else "")
            with cols[i]:
                st.markdown(f"""
                <div class="metric-card {color}">
                    <p class="metric-value {color}">{avg:.2f}</p>
                    <p class="metric-label">{icon} {LABELS[ck]}</p>
                    <p class="metric-sub">{p4:.0f}% ציינו ≥4</p>
                </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Insights ──
    st.markdown("### מסקנות מרכזיות")

    emotion_cols = ["fatigue", "worry", "loneliness", "stress"]
    avgs = {c: df[c].mean() for c in emotion_cols if c in df.columns and df[c].notna().any()}
    if avgs:
        top_e = max(avgs, key=avgs.get)
        st.markdown(f"""
        <div class="insight-box">
            <strong>🔴 אתגר רגשי מוביל:</strong> {LABELS[top_e]} —
            ממוצע {avgs[top_e]:.2f}/5 | {pct_ge4(df[top_e]):.0f}% ציינו 4–5
        </div>""", unsafe_allow_html=True)

    if "income_drop" in df.columns:
        s = df["income_drop"].dropna().astype(str)
        yes_pct = s.str.contains("כן|yes", case=False, na=False).sum() / len(s) * 100 if len(s) else 0
        color = "red" if yes_pct > 40 else "blue"
        st.markdown(f"""
        <div class="insight-box {color}">
            <strong>💸 ירידה בפרנסה:</strong> {yes_pct:.0f}% מהמשיבים דיווחו על ירידה בפרנסה או בהיקף התעסוקה
        </div>""", unsafe_allow_html=True)

    if "help_source" in df.columns:
        items = parse_multi(df["help_source"])
        if items:
            top_h = Counter(items).most_common(1)[0]
            st.markdown(f"""
            <div class="insight-box blue">
                <strong>🤝 מקור עזרה מוביל:</strong> {top_h[0]} ({top_h[1]:,} בחירות)
            </div>""", unsafe_allow_html=True)

    if "support_areas" in df.columns:
        items = parse_multi(df["support_areas"])
        if items:
            top_s = Counter(items).most_common(1)[0]
            st.markdown(f"""
            <div class="insight-box green">
                <strong>🙏 תחום תמיכה מבוקש ביותר:</strong> {top_s[0]} ({top_s[1]:,} בחירות)
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Overview charts (full-width, one per row) ──
    st.markdown("### סקירה כללית")

    # Emotions averages — full width
    emotion_avgs = [(LABELS[c], df[c].mean()) for c in emotion_cols if c in df.columns and df[c].notna().any()]
    if emotion_avgs:
        e_df = pd.DataFrame(emotion_avgs, columns=["רגש", "ממוצע"])
        fig = px.bar(
            e_df, y="רגש", x="ממוצע", orientation="h",
            color="ממוצע",
            color_continuous_scale=[[0, BRAND_LIGHT_BLUE], [0.5, BRAND_BLUE], [1, BRAND_RED]],
            range_color=[1, 5],
            title="ממוצע רגשות מרכזיים (1–5)",
            text=e_df["ממוצע"].round(2).astype(str),
        )
        fig.update_traces(
            textposition="outside",
            textfont=dict(size=15, family="Heebo", color="#222"),
            cliponaxis=False,
        )
        fig.update_layout(**base_layout())
        fig.update_layout(
            height=460,
            coloraxis_showscale=False,
            xaxis=dict(range=[0, 5.5],
                       title=dict(text="ממוצע (1–5)", font=dict(size=16)),
                       tickfont=dict(size=14)),
            yaxis=dict(autorange="reversed", title="",
                       tickfont=dict(size=15, family="Heebo"), automargin=True),
            margin=dict(t=80, b=70, l=200, r=80),
            bargap=0.32,
        )
        st.plotly_chart(fig, use_container_width=True)

    # Source distribution — full width
    if "source" in df.columns:
        sc = df["source"].value_counts().head(12).reset_index()
        sc.columns = ["רשות", "משיבים"]
        fig2 = px.pie(
            sc, names="רשות", values="משיבים",
            color_discrete_sequence=COLOR_SEQ,
            title="התפלגות משיבים לפי רשות מקומית",
            hole=0.35,
        )
        fig2.update_traces(
            textposition="outside",
            textinfo="label+percent",
            textfont=dict(size=14, family="Heebo"),
            marker=dict(line=dict(color="white", width=2)),
        )
        fig2.update_layout(**base_layout())
        fig2.update_layout(
            height=560,
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5,
                        xanchor="left", x=1.02,
                        font=dict(size=13, family="Heebo")),
            margin=dict(t=80, b=70, l=40, r=240),
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── Trends over time ──
    if "month_str" in df.columns and len(df["month_str"].dropna().unique()) > 1:
        st.markdown("### מגמות לאורך זמן")
        time_df = (
            df.groupby("month_str")[emotion_cols]
              .mean()
              .reset_index()
              .sort_values("month_str")
        )
        fig_t = go.Figure()
        tc = [BRAND_BLUE, BRAND_RED, "#4A6FA5", "#E8897A"]
        for i, col in enumerate(emotion_cols):
            if col in time_df.columns:
                fig_t.add_trace(go.Scatter(
                    x=time_df["month_str"], y=time_df[col],
                    name=LABELS[col],
                    line=dict(color=tc[i % len(tc)], width=3),
                    mode="lines+markers",
                    marker=dict(size=9),
                ))
        fig_t.update_layout(**base_layout())
        fig_t.update_layout(
            title="מגמת רגשות לפי חודש",
            height=480,
            yaxis=dict(range=[1, 5],
                       title=dict(text="ממוצע", font=dict(size=16)),
                       tickfont=dict(size=14)),
            xaxis=dict(title=dict(text="חודש", font=dict(size=16)),
                       tickfont=dict(size=14)),
            legend=dict(orientation="h", yanchor="bottom", y=1.04, x=0,
                        font=dict(size=13, family="Heebo")),
        )
        st.plotly_chart(fig_t, use_container_width=True)

    # ── Support areas summary ──
    if "support_areas" in df.columns:
        st.markdown("### תחומי תמיכה מבוקשים")
        items = parse_multi(df["support_areas"])
        if items:
            sup_df = pd.DataFrame(Counter(items).most_common(10), columns=["תחום", "בחירות"])
            sup_df["תווית_מקוצרת"] = sup_df["תחום"].map(_truncate)
            n_bars = len(sup_df)
            chart_h = max(480, min(800, 55 * n_bars + 160))
            fig_s = px.bar(
                sup_df, y="תווית_מקוצרת", x="בחירות",
                orientation="h",
                color_discrete_sequence=[BRAND_RED],
                title="10 תחומי תמיכה מבוקשים ביותר",
                text="בחירות",
                custom_data=["תחום"],
            )
            fig_s.update_traces(
                textposition="outside",
                textfont=dict(size=15, family="Heebo", color=BRAND_RED),
                cliponaxis=False,
                hovertemplate="<b>%{customdata[0]}</b><br>בחירות: %{x:,}<extra></extra>",
            )
            fig_s.update_layout(**base_layout())
            fig_s.update_layout(
                height=chart_h,
                yaxis=dict(autorange="reversed", title="",
                           tickfont=dict(size=15, family="Heebo"), automargin=True),
                xaxis=dict(title=dict(text="מספר בחירות", font=dict(size=16)),
                           tickfont=dict(size=14)),
                margin=dict(t=80, b=70, l=260, r=120),
                bargap=0.32,
            )
            x_max = sup_df["בחירות"].max()
            fig_s.update_xaxes(range=[0, x_max * 1.18])
            st.plotly_chart(fig_s, use_container_width=True)


def render_emotions_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>😓 רגשות ורווחה</h1></div>', unsafe_allow_html=True)
    render_rating_hero(df, ["fatigue", "worry", "loneliness", "stress"],
                       hero_title="📈 תמונת מצב — % שציינו רמה גבוהה בכל רגש")
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    for col in ["fatigue", "worry", "loneliness", "stress"]:
        render_rating_chart(df, col)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)


def render_reserves_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>🎖️ השפעת המילואים</h1></div>', unsafe_allow_html=True)
    diff_cols = ["difficulty_active", "difficulty_between", "difficulty_near", "difficulty_constant"]
    render_rating_hero(df, diff_cols,
                       hero_title="📈 תמונת מצב — % שציינו קושי גבוה לפי שלב")
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    for col in diff_cols:
        render_rating_chart(df, col)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)


def render_community_economy_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>💰 קהילה וכלכלה</h1></div>', unsafe_allow_html=True)
    hero_items = []
    if "community" in df.columns:
        hero_items.append({"label": "מודעות קהילה (גבוהה)", "value": pct_ge4(df["community"])})
    if "economic_concern" in df.columns:
        hero_items.append({"label": "עיסוק כלכלי (גבוה)", "value": pct_ge4(df["economic_concern"])})
    if "income_drop" in df.columns:
        s = df["income_drop"].dropna().astype(str)
        if len(s):
            yes_pct = s.str.contains("כן|חלקית", case=False, na=False).sum() / len(s) * 100
            hero_items.append({"label": "ירידה בפרנסה", "value": yes_pct})
    render_mixed_hero(df, hero_items, hero_title="📈 תמונת מצב — קהילה וכלכלה")
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    render_rating_chart(df, "community")
    render_rating_chart(df, "economic_concern")
    render_categorical_chart(df, "income_drop")
    render_categorical_chart(df, "main_difficulty")


def render_family_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>💑 זוגיות ומשפחה</h1></div>', unsafe_allow_html=True)
    family_items = []
    if "parenting" in df.columns:
        family_items.append({"label": "קושי כהורה (גבוה)", "value": pct_ge4(df["parenting"])})
    if "relationship_difficulties" in df.columns:
        s = df["relationship_difficulties"].dropna().astype(str)
        if len(s):
            neg_pct = s.str.contains("כן|קיים|יש", case=False, na=False).sum() / len(s) * 100
            family_items.append({"label": "קשיים זוגיים (מדווחים)", "value": neg_pct})
    render_mixed_hero(df, family_items, hero_title="📈 תמונת מצב — זוגיות ומשפחה")
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    render_categorical_chart(df, "relationship")
    render_categorical_chart(df, "relationship_difficulties")
    render_rating_chart(df, "parenting")


def render_children_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>👶 ילדים וחינוך</h1></div>', unsafe_allow_html=True)
    render_multichoice_hero(df, "education_difficulties",
                            hero_title="📈 תמונת מצב — 5 הקשיים הנפוצים במסגרות חינוכיות", top_n=5)
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    render_multichoice_chart(df, "child_behavior")
    render_categorical_chart(df, "education_change")
    render_multichoice_chart(df, "education_difficulties")


def render_support_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>🤝 תמיכה ומענה</h1></div>', unsafe_allow_html=True)
    render_multichoice_hero(df, "support_areas",
                            hero_title="📈 תמונת מצב — 5 תחומי תמיכה מבוקשים ביותר", top_n=5)
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    render_multichoice_chart(df, "help_source")
    render_multichoice_chart(df, "support_areas")


def render_freetext_page(df: pd.DataFrame):
    st.markdown('<div class="logo-header"><h1>✍️ מענה חופשי</h1></div>', unsafe_allow_html=True)
    for col in FREE_TEXT_COLS:
        render_freetext_summary(df, col)
        st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)


# =====================================================================
# MAIN
# =====================================================================
def main():
    if "uploaded_bytes" in st.session_state:
        df_raw = load_data_from_bytes(
            st.session_state["uploaded_bytes"],
            st.session_state["uploaded_name"],
        )
    else:
        df_raw = load_default()

    if df_raw is None:
        st.error(f"⚠️ לא נמצא קובץ נתונים")
        st.info("העלה קובץ Excel דרך הסיידבר →")
        st.code(f"תיקייה נוכחית: {os.getcwd()}\nשם קובץ ברירת מחדל: {FILE_PATH}")
        return

    filters, page = render_sidebar(df_raw)
    df = apply_filters(df_raw, filters)

    n_filtered = len(df)
    n_total    = len(df_raw)
    if n_filtered < n_total:
        st.info(f"🔍 מציג **{n_filtered:,}** מתוך **{n_total:,}** משיבים לפי הסינון הנוכחי")

    route = {
        "סיכום ראשי":      lambda: render_summary_page(df, df_raw),
        "רגשות ורווחה":    lambda: render_emotions_page(df),
        "השפעת המילואים":  lambda: render_reserves_page(df),
        "קהילה וכלכלה":   lambda: render_community_economy_page(df),
        "זוגיות ומשפחה":  lambda: render_family_page(df),
        "ילדים וחינוך":   lambda: render_children_page(df),
        "תמיכה ומענה":    lambda: render_support_page(df),
        "מענה חופשי":     lambda: render_freetext_page(df),
    }
    for key, fn in route.items():
        if page.endswith(key):
            fn()
            break


if __name__ == "__main__":
    main()
