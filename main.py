import streamlit as st
from docx import Document
import datetime
from io import BytesIO
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import base64
import os
import json
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import pytz
# =================================================================
# PAGE CONFIG
# =================================================================
st.set_page_config(
    page_title="ระบบบันทึกจับกุม",
    page_icon="Thai_Police.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =================================================================
# USERS
# =================================================================
USERS = {
    "admin":      {"password": "123",  "fullname": "Admin"},
    "somkid":     {"password": "1234", "fullname": "ร.ต.อ.สมคิด เชื้อเวียง"},
    "chaiporn":   {"password": "1234", "fullname": "ร.ต.อ.ชัยพร ชอบงาม"},
    "teerawat":   {"password": "1234", "fullname": "ส.ต.อ.ธีระวัฒน์ แก่นสาร"},
    "narongrit":  {"password": "1234", "fullname": "ส.ต.อ.ณรงค์ฤทธิ์ เหล่าดี"},
    "wichakorn":  {"password": "1234", "fullname": "ส.ต.อ.วิชชากร วงษ์โท"},
    "watchara":   {"password": "1234", "fullname": "จ.ส.ต.วัชระ จันสุตะ"},
    "wachirawit": {"password": "1234", "fullname": "ส.ต.ท.วชิรวิชญ์ นันทรักษ์"},
    "wisarut":    {"password": "1234", "fullname": "ส.ต.ท.วิศรุต จันทร์สิงห์"},
    "adisorn":    {"password": "1234", "fullname": "ส.ต.ต.อดิศร ศุภนิกร"},
}

# =================================================================
# SESSION STATE INIT
# =================================================================
for key, default in [
    ("settings", {"theme": "light", "logo": None}),
    ("form", {}),
    ("tab", 0),
    ("records", []),
    ("doc_records", []),
    ("doc_running_numbers", {}),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# =================================================================
# KTB DESIGN SYSTEM — CSS (แก้ไขครั้งเดียว ครอบคลุมทุกหน้า)
# =================================================================
KTB_BLUE       = "#00AEEF"   # Krungthai primary
KTB_BLUE_DARK  = "#0076B6"   # Hover / dark shade
KTB_NAVY       = "#003B6F"   # Deep navy for sidebar
KTB_GOLD       = "#F5A623"   # Accent / KTB secondary
KTB_TEXT       = "#1A2B45"   # Body text light mode
KTB_SURFACE    = "#F5F8FB"   # Card bg light
KTB_BORDER     = "#D6E4F0"   # Subtle border

st.markdown(f"""
<style>
/* ─── GOOGLE FONTS ─────────────────────────────────────── */
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700;800&display=swap');

/* ─── CSS VARIABLES (Light + Dark) ─────────────────────── */
:root {{
    --ktb-blue:       {KTB_BLUE};
    --ktb-blue-dark:  {KTB_BLUE_DARK};
    --ktb-navy:       {KTB_NAVY};
    --ktb-gold:       {KTB_GOLD};
    --ktb-text:       {KTB_TEXT};
    --ktb-surface:    {KTB_SURFACE};
    --ktb-border:     {KTB_BORDER};
    --ktb-radius:     12px;
    --ktb-radius-lg:  18px;
    --ktb-shadow:     0 2px 12px rgba(0,174,239,0.08);
    --ktb-shadow-md:  0 4px 24px rgba(0,59,111,0.12);
    /* dynamic */
    --bg-page:        #EBF4FB;
    --bg-card:        #FFFFFF;
    --text-primary:   {KTB_TEXT};
    --text-muted:     #5A7A99;
    --text-label:     #2E4D6B;
    --border-color:   {KTB_BORDER};
    --input-bg:       #FFFFFF;
    --sidebar-bg:     {KTB_NAVY};
}}

/* Dark mode overrides */
@media (prefers-color-scheme: dark) {{
    :root {{
        --bg-page:      #0D1B2A;
        --bg-card:      #112236;
        --text-primary: #E8F4FD;
        --text-muted:   #7FB3D3;
        --text-label:   #A8CCE8;
        --border-color: #1E3A54;
        --input-bg:     #152D45;
        --sidebar-bg:   #08131E;
    }}
}}

/* Streamlit dark-mode class */
[data-theme="dark"] {{
    --bg-page:      #0D1B2A !important;
    --bg-card:      #112236 !important;
    --text-primary: #E8F4FD !important;
    --text-muted:   #7FB3D3 !important;
    --text-label:   #A8CCE8 !important;
    --border-color: #1E3A54 !important;
    --input-bg:     #152D45 !important;
}}

/* ─── BASE ──────────────────────────────────────────────── */
html, body, [class*="css"], .stApp {{
    font-family: 'Sarabun', sans-serif !important;
    background: var(--bg-page) !important;
    color: var(--text-primary) !important;
}}

/* ─── HIDE STREAMLIT CHROME ─────────────────────────────── */
header[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDeployButton"] {{
    display: none !important;
}}

/* ─── BLOCK CONTAINER ───────────────────────────────────── */
.block-container {{
    max-width: 1400px !important;
    padding: 1.5rem 1.5rem 3rem !important;
}}

/* ─── SIDEBAR ────────────────────────────────────────────── */
[data-testid="stSidebar"] {{
    background: var(--sidebar-bg) !important;
    border-right: 1px solid rgba(0,174,239,0.15) !important;
    min-width: 270px !important;
    max-width: 270px !important;
}}

[data-testid="stSidebar"] * {{
    color: #E8F4FD !important;
    font-family: 'Sarabun', sans-serif !important;
}}

[data-testid="stSidebar"] .stButton > button {{
    background: rgba(0,174,239,0.12) !important;
    border: 1px solid rgba(0,174,239,0.3) !important;
    color: #E8F4FD !important;
    border-radius: var(--ktb-radius) !important;
    font-size: 14px !important;
    padding: 10px 14px !important;
    text-align: left !important;
    transition: background 0.2s, border-color 0.2s !important;
}}

[data-testid="stSidebar"] .stButton > button:hover {{
    background: rgba(0,174,239,0.25) !important;
    border-color: var(--ktb-blue) !important;
}}

/* ─── BUTTONS ────────────────────────────────────────────── */
div.stButton > button {{
    width: 100% !important;
    background: linear-gradient(135deg, var(--ktb-blue) 0%, var(--ktb-blue-dark) 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: var(--ktb-radius) !important;
    padding: 12px 20px !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    font-family: 'Sarabun', sans-serif !important;
    letter-spacing: 0.2px !important;
    transition: opacity 0.2s, transform 0.15s !important;
    box-shadow: 0 3px 10px rgba(0,118,182,0.25) !important;
}}

div.stButton > button:hover {{
    opacity: 0.92 !important;
    transform: translateY(-1px) !important;
}}

div.stButton > button:active {{
    transform: translateY(0) !important;
    opacity: 1 !important;
}}

/* Secondary / ghost buttons (nav tabs) */
.ktb-nav-btn > button {{
    background: var(--bg-card) !important;
    color: var(--text-muted) !important;
    border: 1px solid var(--border-color) !important;
    box-shadow: none !important;
}}

.ktb-nav-btn-active > button,
.ktb-nav-btn > button:hover {{
    background: var(--ktb-blue) !important;
    color: #fff !important;
    border-color: var(--ktb-blue) !important;
}}

/* ─── INPUTS ─────────────────────────────────────────────── */
.stTextInput input,
.stDateInput input,
.stTimeInput input,
.stTextArea textarea,
.stNumberInput input,
.stSelectbox div[data-baseweb="select"] > div {{
    background: var(--input-bg) !important;
    color: var(--text-primary) !important;
    border: 1.5px solid var(--border-color) !important;
    border-radius: 10px !important;
    font-family: 'Sarabun', sans-serif !important;
    font-size: 15px !important;
    transition: border-color 0.2s !important;
}}

.stTextInput input:focus,
.stTextArea textarea:focus,
.stNumberInput input:focus {{
    border-color: var(--ktb-blue) !important;
    box-shadow: 0 0 0 3px rgba(0,174,239,0.12) !important;
    outline: none !important;
}}

/* Labels */
.stTextInput label, .stDateInput label, .stTimeInput label,
.stSelectbox label, .stTextArea label, .stNumberInput label,
.stRadio label, .stCheckbox label {{
    color: var(--text-label) !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    margin-bottom: 4px !important;
}}

input::placeholder, textarea::placeholder {{
    color: var(--text-muted) !important;
    opacity: 0.7 !important;
}}

/* ─── CARDS ──────────────────────────────────────────────── */
.ktb-card {{
    background: var(--bg-card);
    border: 1px solid var(--border-color);
    border-radius: var(--ktb-radius-lg);
    padding: 1.25rem 1.5rem;
    box-shadow: var(--ktb-shadow);
    margin-bottom: 1rem;
}}

/* ─── KPI CARDS ──────────────────────────────────────────── */
.kpi-card {{
    background: var(--bg-card);
    border: 1px solid var(--border-color);
    border-left: 4px solid var(--ktb-blue);
    border-radius: var(--ktb-radius);
    padding: 1.1rem 1.25rem;
    box-shadow: var(--ktb-shadow);
    transition: box-shadow 0.2s;
}}

.kpi-card:hover {{
    box-shadow: var(--ktb-shadow-md);
}}

.kpi-number {{
    font-size: 36px;
    font-weight: 800;
    color: var(--ktb-blue) !important;
    line-height: 1.1;
    margin: 4px 0;
}}

.kpi-label {{
    font-size: 13px;
    color: var(--text-muted) !important;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}

.kpi-icon {{
    font-size: 22px;
    margin-bottom: 4px;
}}

/* ─── METRIC CARD ────────────────────────────────────────── */
.metric-card {{
    background: var(--bg-card);
    border: 1px solid var(--border-color);
    border-radius: var(--ktb-radius);
    padding: 1rem 1.1rem;
    text-align: center;
    transition: border-color 0.2s, box-shadow 0.2s;
    cursor: pointer;
}}

.metric-card:hover {{
    border-color: var(--ktb-blue);
    box-shadow: 0 0 0 3px rgba(0,174,239,0.1);
}}

.metric-icon {{ font-size: 24px; margin-bottom: 6px; }}
.metric-no   {{ font-size: 32px; font-weight: 800; color: var(--ktb-blue) !important; }}
.metric-text {{ font-size: 13px; color: var(--text-muted) !important; font-weight: 500; }}

/* ─── HEADER BAR ─────────────────────────────────────────── */
.main-header {{
    background: linear-gradient(135deg, var(--ktb-blue) 0%, var(--ktb-blue-dark) 60%, var(--ktb-navy) 100%);
    padding: 18px 24px;
    border-radius: var(--ktb-radius-lg);
    color: #fff !important;
    font-size: 20px;
    font-weight: 700;
    display: flex;
    align-items: center;
    gap: 12px;
    box-shadow: 0 4px 20px rgba(0,59,111,0.2);
    margin-bottom: 1.25rem;
}}

.main-header * {{ color: #fff !important; }}

/* ─── SECTION DIVIDERS ───────────────────────────────────── */
hr {{
    border: none;
    border-top: 1.5px solid var(--border-color) !important;
    margin: 1.25rem 0 !important;
}}

/* ─── PLOTLY CHARTS ──────────────────────────────────────── */
.js-plotly-plot {{
    border-radius: var(--ktb-radius-lg) !important;
    overflow: hidden !important;
    box-shadow: var(--ktb-shadow) !important;
}}

/* ─── ALERTS / INFO ──────────────────────────────────────── */
.stAlert {{
    border-radius: var(--ktb-radius) !important;
    font-family: 'Sarabun', sans-serif !important;
}}

/* ─── EXPANDER ───────────────────────────────────────────── */
.streamlit-expanderHeader {{
    background: var(--bg-card) !important;
    border-radius: var(--ktb-radius) !important;
    color: var(--text-primary) !important;
    font-weight: 600 !important;
}}

/* ─── DATAFRAME ──────────────────────────────────────────── */
[data-testid="stDataFrame"] {{
    background: var(--bg-card) !important;
    border-radius: var(--ktb-radius) !important;
    overflow: hidden !important;
}}

/* ─── TABS (Streamlit native) ────────────────────────────── */
button[data-baseweb="tab"] {{
    font-family: 'Sarabun', sans-serif !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    color: var(--text-muted) !important;
}}

button[data-baseweb="tab"][aria-selected="true"] {{
    color: var(--ktb-blue) !important;
    border-bottom-color: var(--ktb-blue) !important;
}}

/* ─── SPINNER ────────────────────────────────────────────── */
.stSpinner > div {{ border-top-color: var(--ktb-blue) !important; }}

/* ─── SCROLLBAR ──────────────────────────────────────────── */
::-webkit-scrollbar {{ width: 6px; height: 6px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{
    background: rgba(0,174,239,0.3);
    border-radius: 99px;
}}
::-webkit-scrollbar-thumb:hover {{ background: var(--ktb-blue); }}

/* ─── MOBILE RESPONSIVE ──────────────────────────────────── */
@media (max-width: 768px) {{
    .block-container {{
        padding: 0.75rem 0.75rem 2rem !important;
    }}

    [data-testid="stSidebar"] {{
        min-width: 240px !important;
        max-width: 240px !important;
    }}

    .kpi-number {{ font-size: 28px !important; }}
    .main-header {{ font-size: 16px !important; padding: 14px 16px !important; }}

    div.stButton > button {{
        font-size: 14px !important;
        padding: 10px 14px !important;
    }}

    .stTextInput input,
    .stTextArea textarea,
    .stNumberInput input {{
        font-size: 16px !important;  /* ป้องกัน iOS zoom */
    }}
}}

@media (max-width: 480px) {{
    .kpi-number {{ font-size: 24px !important; }}
    .metric-no  {{ font-size: 24px !important; }}
    .main-header {{ font-size: 15px !important; }}
}}

/* ─── LOGIN CARD ─────────────────────────────────────────── */
.login-wrapper {{
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 2rem 1rem;
}}

.login-card {{
    background: var(--bg-card);
    border: 1px solid var(--border-color);
    border-radius: 24px;
    padding: 2.5rem 2rem;
    max-width: 420px;
    width: 100%;
    box-shadow: var(--ktb-shadow-md);
}}

.logo-spin {{
    width: 90px;
    display: block;
    margin: 0 auto 1rem;
    animation: spinY 6s linear infinite;
}}

@keyframes spinY {{
    from {{ transform: rotateY(0deg); }}
    to   {{ transform: rotateY(360deg); }}
}}

/* ─── BADGE ──────────────────────────────────────────────── */
.ktb-badge {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 99px;
    font-size: 12px;
    font-weight: 600;
}}

.ktb-badge-blue {{
    background: rgba(0,174,239,0.12);
    color: var(--ktb-blue-dark) !important;
}}

.ktb-badge-gold {{
    background: rgba(245,166,35,0.15);
    color: #A0660A !important;
}}

/* ─── SECTION TITLE ──────────────────────────────────────── */
.section-title {{
    font-size: 16px;
    font-weight: 700;
    color: var(--ktb-navy);
    margin: 0 0 12px;
    display: flex;
    align-items: center;
    gap: 8px;
}}

[data-theme="dark"] .section-title {{ color: var(--ktb-blue) !important; }}

/* ─── NAV STRIP ──────────────────────────────────────────── */
.ktb-nav-strip {{
    background: var(--bg-card);
    border: 1px solid var(--border-color);
    border-radius: var(--ktb-radius);
    padding: 6px;
    display: flex;
    gap: 4px;
    margin-bottom: 1.25rem;
}}

/* ─── SIDEBAR PROFILE BOX ────────────────────────────────── */
.sidebar-profile {{
    background: rgba(0,174,239,0.1);
    border: 1px solid rgba(0,174,239,0.2);
    border-radius: var(--ktb-radius);
    padding: 12px 14px;
    margin: 0 0 1rem;
    font-size: 13px;
    line-height: 1.7;
}}
</style>
""", unsafe_allow_html=True)


# =================================================================
# HELPER FUNCTIONS
# =================================================================
def date_th(d):
    if not d:
        return "-"
    months = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
              "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    try:
        return f"{d.day} {months[d.month]} {d.year + 543}"
    except:
        return str(d)


def safe_time(t):
    if isinstance(t, (datetime.time, datetime.datetime)):
        return t.strftime("%H:%M")
    return str(t) if t else "-"


def to_arabic_number(text):
    if text is None:
        return ""
    m = {"๐":"0","๑":"1","๒":"2","๓":"3","๔":"4",
         "๕":"5","๖":"6","๗":"7","๘":"8","๙":"9"}
    return "".join(m.get(c, c) for c in str(text))


def format_thai_id(id_number):
    digits = "".join(filter(str.isdigit, str(id_number or "")))
    if len(digits) != 13:
        return id_number or ""
    return f"{digits[0]}-{digits[1:5]}-{digits[5:10]}-{digits[10:12]}-{digits[12]}"


def clean_geo(text):
    if not text:
        return ""
    text = str(text).strip()
    for w in ["จังหวัด","อำเภอ","ตำบล"]:
        text = text.replace(w, "")
    return text.strip()


def format_amphur(name):
    if not name:
        return ""
    for x in ["อำเภอ","อ.","เขต"]:
        name = str(name).replace(x, "")
    return name.strip()


def format_tambon(name):
    return str(name or "").replace("ตำบล","").strip()


def safe_index(value, options):
    return options.index(value) if value in options else 0


def remove_empty_signature_rows(doc):
    for table in doc.tables:
        rows_to_remove = []
        for row in table.rows:
            if len(row.cells) < 3:
                continue
            center = row.cells[1].text.strip()
            right  = row.cells[2].text.strip()
            if right == "" and "ผู้จับกุม" in center:
                rows_to_remove.append(row)
        for row in rows_to_remove:
            table._tbl.remove(row._tr)


def replace_text(doc, data):
    def smart_replace(paragraph):
        full_text = "".join(run.text for run in paragraph.runs)
        for key, value in data.items():
            if key in full_text:
                full_text = full_text.replace(key, str(value))
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for run in paragraph.runs[1:]:
                run.text = ""

    for p in doc.paragraphs:
        smart_replace(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    smart_replace(p)
    remove_empty_signature_rows(doc)


@st.cache_data
def load_geo():
    try:
        df = pd.read_csv("thai_districts.csv", encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv("thai_districts.csv", encoding="cp874")
    df["ProvinceThai"] = df["ProvinceThai"].apply(clean_geo)
    df["DistrictThai"] = df["DistrictThai"].apply(clean_geo).apply(format_amphur)
    df["TambonThai"]   = df["TambonThai"].apply(clean_geo).apply(format_tambon)
    return df


def get_base64(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


VEHICLE_MODELS = {
    "Honda":  ["Wave 110i","Wave 125i","Click 125i","PCX","ADV160"],
    "Yamaha": ["NMAX","Aerox","Grand Filano","Fino"],
    "Toyota": ["Vios","Yaris","Hilux Revo","Fortuner"],
    "Isuzu":  ["D-Max","MU-X"],
    "Nissan": ["Navara","Almera"],
    "Mazda":  ["Mazda2","Mazda3","BT-50"],
}

OFFICERS = [
    "ร.ต.อ.สมคิด เชื้อเวียง","ร.ต.อ.ชัยพร ชอบงาม","ร.ต.ท.พจน์ ปาสาจันทร์",
    "ร.ต.ต.ประจวบ ศรีบุระ","ส.ต.อ.ธีระวัฒน์ แก่นสาร","ส.ต.อ.ณรงค์ฤทธิ์ เหล่าดี",
    "ส.ต.อ.วิชชากร วงษ์โท","ด.ต.ยุทธพงษ์ ชาดแดง","จ.ส.ต.วัชระ จันสุตะ",
    "ส.ต.ท.วชิรวิชญ์ นันทรักษ์","ส.ต.ท.วิศรุต จันทร์สิงห์","ส.ต.ท.เอกพจน์ อินผล",
    "ส.ต.ต.อดิศร ศุภนิกร",
]

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
logo_b64    = st.session_state.settings.get("logo") or get_base64(os.path.join(BASE_DIR, "police_logo.png"))


# =================================================================
# CHART THEME (Plotly)
# =================================================================
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Sarabun, sans-serif", color="#5A7A99", size=12),
    margin=dict(l=20, r=20, t=44, b=20),
    title_font=dict(size=15, color="#003B6F"),
    colorway=["#00AEEF","#0076B6","#003B6F","#F5A623","#2EC4B6","#E84855"],
    xaxis=dict(gridcolor="rgba(0,174,239,0.08)", linecolor="rgba(0,0,0,0.08)"),
    yaxis=dict(gridcolor="rgba(0,174,239,0.08)", linecolor="rgba(0,0,0,0.08)"),
)


# =================================================================
# ─── LOGIN ──────────────────────────────────────────────────────
# =================================================================
if "password_correct" not in st.session_state:
    if logo_b64:
        st.markdown(
            f'<img class="logo-spin" src="data:image/png;base64,{logo_b64}">',
            unsafe_allow_html=True,
        )
    st.markdown("""
    <div style="text-align:center; margin-bottom:1.5rem;">
        <div style="font-size:28px; font-weight:800; color:var(--ktb-blue);">ยินดีต้อนรับ</div>
        <div style="font-size:15px; color:var(--text-muted); margin-top:4px;">
            ระบบบันทึกการจับกุม งานจราจร<br>
            <strong>สถานีตำรวจภูธรตระการพืชผล</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)

    _, c, _ = st.columns([1, 2, 1])
    with c:
        u = st.text_input("👤 ชื่อผู้ใช้งาน", key="lu")
        p = st.text_input("🔒 รหัสผ่าน", type="password", key="lp")
        st.write("")
        if st.button("เข้าสู่ระบบ", use_container_width=True):
            ud = USERS.get(u)
            if ud and ud["password"] == p:
                st.session_state["password_correct"]  = True
                st.session_state["user_full_name"]     = ud["fullname"]
                st.session_state.page = "dashboard"
                st.rerun()
            else:
                st.error("❌ ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
    st.stop()


# =================================================================
# DEFAULT PAGE
# =================================================================
if "page" not in st.session_state:
    st.session_state.page = "dashboard"

records     = st.session_state.get("records", [])
doc_records = st.session_state.get("doc_records", [])

# =================================================================
# ─── SIDEBAR ────────────────────────────────────────────────────
# =================================================================
with st.sidebar:
    if logo_b64:
        st.markdown(
            f'<div style="text-align:center;margin-bottom:8px;">'
            f'<img src="data:image/png;base64,{logo_b64}" width="80" '
            f'style="border-radius:16px;padding:8px;background:rgba(0,174,239,0.15);"></div>',
            unsafe_allow_html=True,
        )

    st.markdown("""
    <div style="text-align:center;font-size:17px;font-weight:800;
                color:#00AEEF;margin-bottom:14px;line-height:1.35;">
        ระบบบันทึกจับกุม<br>
        <span style="font-size:13px;font-weight:400;color:#7FB3D3;">งานจราจร</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(
        f'<div class="sidebar-profile">'
        f'👮 <strong>{st.session_state.get("user_full_name","เจ้าหน้าที่")}</strong><br>'
        f'🏢 สภ.ตระการพืชผล</div>',
        unsafe_allow_html=True,
    )

    nav_items = [
        ("🏠", "หน้าแรก",      "dashboard"),
        ("📋", "บันทึกจับกุม", "form"),
        ("🔎", "ค้นหาเอกสาร",  "search"),
        ("📂", "จัดการเอกสาร", "documents"),
        ("📅", "ตารางเวร",      "schedule"),
        ("⚙️", "ตั้งค่า",       "settings"),
    ]
    for icon, label, page_key in nav_items:
        if st.button(f"{icon}  {label}", use_container_width=True, key=f"nav_{page_key}"):
            st.session_state.page = page_key
            st.rerun()

    st.markdown("<hr style='margin:10px 0;border-color:rgba(0,174,239,0.2);'>", unsafe_allow_html=True)

    if st.button("🚪  ออกจากระบบ", use_container_width=True, key="logout_btn"):
        st.session_state.confirm_logout = True

if st.session_state.get("confirm_logout"):
    st.warning("⚠️ ต้องการออกจากระบบ? ข้อมูลที่ยังไม่บันทึกจะหายไป")
    cy, cn = st.columns(2)
    if cy.button("✅ ยืนยัน", key="confirm_yes"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
    if cn.button("❌ ยกเลิก", key="confirm_no"):
        st.session_state.confirm_logout = False
        st.rerun()


# =================================================================
# ─── DASHBOARD ──────────────────────────────────────────────────
# =================================================================
if st.session_state.page == "dashboard":
    total   = len(records)
    today   = len([r for r in records if str(r.get("report_date")) == str(datetime.date.today())])
    latest  = records[-1].get("name", "-") if records else "-"

    st.markdown('<div class="main-header">🚔&nbsp; POLICE REALTIME DASHBOARD — สภ.ตระการพืชผล</div>', unsafe_allow_html=True)

    # KPI Row
    kpi_data = [
        ("📁", total,   "คดีทั้งหมด"),
        ("📅", today,   "วันนี้"),
        ("🖨️", total,   "เอกสารทั้งหมด"),
        ("👤", latest,  "ล่าสุด"),
    ]
    cols = st.columns(4)
    for col, (icon, val, label) in zip(cols, kpi_data):
        col.markdown(
            f'<div class="kpi-card">'
            f'<div class="kpi-icon">{icon}</div>'
            f'<div class="kpi-number">{val}</div>'
            f'<div class="kpi-label">{label}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.write("")

    # ── Sample / real data ──
    if not records:
        charges = ["เมาแล้วขับ","ไม่มีใบขับขี่","ขับเร็ว","ฝ่าไฟแดง","ไม่สวมหมวก"]
        sample = [{
            "report_date": datetime.date.today() - datetime.timedelta(days=int(np.random.randint(0,30))),
            "charge": str(np.random.choice(charges)),
            "hour":   int(np.random.randint(0,24)),
        } for _ in range(60)]
        df = pd.DataFrame(sample)
    else:
        rows = []
        for r in records:
            try:
                d = pd.to_datetime(r.get("report_date", datetime.date.today()))
            except Exception:
                d = pd.to_datetime(datetime.date.today())
            rows.append({"report_date": d, "charge": r.get("charge","อื่นๆ"), "hour": int(np.random.randint(0,24))})
        df = pd.DataFrame(rows)

    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["month"]       = df["report_date"].dt.strftime("%m")

    charge_df = df.groupby("charge").size().reset_index(name="count")
    month_df  = df.groupby("month").size().reset_index(name="count")

    g1, g2 = st.columns([1.5, 1])

    with g1:
        fig = px.bar(month_df, x="month", y="count", text_auto=True, title="📈 สถิติคดีรายเดือน",
                     color_discrete_sequence=["#00AEEF"])
        fig.update_layout(**PLOTLY_LAYOUT, height=380)
        fig.update_traces(marker_line_width=0, marker_cornerradius=4)
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        fig2 = px.pie(charge_df, names="charge", values="count", title="🥧 สัดส่วนข้อหา",
                      color_discrete_sequence=["#00AEEF","#0076B6","#003B6F","#F5A623","#2EC4B6"])
        fig2.update_layout(**PLOTLY_LAYOUT, height=380)
        fig2.update_traces(textfont_size=13)
        st.plotly_chart(fig2, use_container_width=True)

    g3, g4 = st.columns(2)

    with g3:
        fig3 = px.bar(charge_df, x="charge", y="count", color="charge", title="🚨 สถิติแยกตามข้อหา",
                      color_discrete_sequence=["#00AEEF","#0076B6","#003B6F","#F5A623","#2EC4B6"])
        fig3.update_layout(**PLOTLY_LAYOUT, height=340, showlegend=False)
        fig3.update_traces(marker_line_width=0, marker_cornerradius=4)
        st.plotly_chart(fig3, use_container_width=True)

    with g4:
        heat = df.groupby("hour").size().reset_index(name="count")
        fig4 = px.density_heatmap(heat, x="hour", y="count", title="🔥 Heatmap เวลาเกิดเหตุ",
                                  color_continuous_scale=["#EBF4FB","#00AEEF","#003B6F"])
        fig4.update_layout(**PLOTLY_LAYOUT, height=340)
        st.plotly_chart(fig4, use_container_width=True)

    # Realtime line
    rt = pd.DataFrame({
        "เวลา": pd.date_range(start=datetime.datetime.now(), periods=20, freq="min"),
        "คดี":  np.random.randint(1, 10, 20),
    })
    fig5 = px.line(rt, x="เวลา", y="คดี", title="📡 Realtime Monitoring",
                   color_discrete_sequence=["#00AEEF"])
    fig5.update_layout(**PLOTLY_LAYOUT, height=300)
    fig5.update_traces(line_width=2.5)
    st.plotly_chart(fig5, use_container_width=True)

    # ── Document type cards ──
    st.markdown("### 📄 ประเภทเอกสาร")
    DOC_META = {
        "หนังสือราชการ": "📋","บันทึกข้อความ": "📝","คำสั่ง": "📌",
        "ประกาศ": "📢","รายงาน": "📊","หนังสือรับ": "📥","หนังสือส่ง": "📤",
    }
    type_counts = {t: 0 for t in DOC_META}
    for d in doc_records:
        t = d.get("doc_type","")
        if t in type_counts:
            type_counts[t] += 1

    dc = st.columns(4)
    for idx, (dtype, icon) in enumerate(DOC_META.items()):
        dc[idx % 4].markdown(
            f'<div class="metric-card">'
            f'<div class="metric-icon">{icon}</div>'
            f'<div class="metric-no">{type_counts[dtype]}</div>'
            f'<div class="metric-text">{dtype}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

# =================================================================
# ─── SETTINGS ───────────────────────────────────────────────────
# =================================================================
if st.session_state.page == "settings":
    st.markdown('<div class="main-header">⚙️&nbsp; ตั้งค่าระบบ</div>', unsafe_allow_html=True)
    s = st.session_state.settings

    tab1, tab2, tab3, tab4 = st.tabs(["🏢 หน่วยงาน","🎨 ธีม","💾 สำรองข้อมูล","⚙️ ขั้นสูง"])

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            s["station"]   = st.text_input("ชื่อสถานีตำรวจ",  s.get("station","สภ.ตระการพืชผล"))
            s["province"]  = st.text_input("จังหวัด",          s.get("province","อุบลราชธานี"))
            s["commander"] = st.text_input("ชื่อผู้กำกับ",      s.get("commander",""))
        with c2:
            s["deputy"]    = st.text_input("รองผู้กำกับ",       s.get("deputy",""))
            s["inspector"] = st.text_input("สารวัตร",           s.get("inspector",""))
        logo_file = st.file_uploader("อัปโหลดโลโก้", type=["png","jpg","jpeg"])
        if logo_file:
            s["logo"] = base64.b64encode(logo_file.read()).decode()
            st.success("อัปโหลดโลโก้แล้ว")

    with tab2:
        s["theme"] = st.selectbox("โหมดระบบ", ["light","dark"], index=0 if s.get("theme","light")=="light" else 1)

    with tab3:
        recs = st.session_state.get("records",[])
        st.info(f"ข้อมูลทั้งหมด {len(recs)} รายการ")
        st.download_button("📥 ดาวน์โหลด Backup",
                           data=json.dumps(recs, ensure_ascii=False, indent=2, default=str),
                           file_name="backup_records.json", mime="application/json")
        up = st.file_uploader("📤 นำเข้าข้อมูล", type=["json"], key="restore_file")
        if up:
            try:
                st.session_state.records = json.load(up)
                st.success("นำเข้าข้อมูลสำเร็จ")
            except Exception:
                st.error("ไฟล์ไม่ถูกต้อง")

    with tab4:
        f = st.session_state.get("form",{})
        auto_text = (
            f"ตามวันเวลาที่เกิดเหตุ เจ้าหน้าที่ตำรวจชุดจับกุมขณะปฏิบัติหน้าที่\n"
            f"พบ {f.get('name','')} ขับขี่{f.get('vehicle_type','')} ยี่ห้อ {f.get('vehicle_brand','')}\n"
            f"หมายเลขทะเบียน {f.get('vehicle_plate','')}\n"
            f"จากการตรวจสอบพบ {f.get('seized_item','')}\n"
            f"จึงแจ้งข้อกล่าวหา {f.get('charge','')}"
        )
        behavior = st.text_area("พฤติการณ์จับกุม", value=f.get("behavior_text", auto_text), height=250)
        st.session_state.form["behavior_text"] = behavior
        s["auto_save"]   = st.toggle("💾 บันทึกอัตโนมัติ", value=s.get("auto_save", True))
        s["thai_number"] = st.toggle("🔢 ใช้เลขไทย",       value=s.get("thai_number", False))
        st.write("")
        c3, c4 = st.columns(2)
        with c3:
            if st.button("🗑️ ล้างข้อมูลทั้งหมด", use_container_width=True):
                st.session_state.records = []
                st.success("ล้างข้อมูลแล้ว")
        with c4:
            if st.button("🚪 รีเซ็ตระบบ", use_container_width=True):
                for k in list(st.session_state.keys()):
                    del st.session_state[k]
                st.rerun()

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 บันทึกการตั้งค่า", use_container_width=True):
            st.session_state.settings = s
            st.success("บันทึกแล้ว")
    with c2:
        if st.button("⬅️ กลับหน้าหลัก", use_container_width=True):
            st.session_state.page = "dashboard"
            st.rerun()


# =================================================================
# ─── SEARCH ─────────────────────────────────────────────────────
# =================================================================
if st.session_state.page == "search":
    st.markdown('<div class="main-header">🔎&nbsp; ค้นหาเอกสาร</div>', unsafe_allow_html=True)
    recs = st.session_state.get("records",[])
    c1, c2 = st.columns([3,1])
    keyword = c1.text_input("ค้นหาชื่อ / เลขบัตร / ทะเบียนรถ")
    mode    = c2.selectbox("ประเภท", ["ทั้งหมด","ชื่อ","เลขบัตร","ทะเบียน"])

    result = []
    for r in recs:
        name  = str(r.get("name",""))
        cid   = str(r.get("sub_id",""))
        plate = str(r.get("vehicle_plate",""))
        ok = True
        if keyword:
            kw = keyword.lower()
            if mode == "ชื่อ":    ok = kw in name.lower()
            elif mode == "เลขบัตร": ok = kw in cid.lower()
            elif mode == "ทะเบียน": ok = kw in plate.lower()
            else: ok = kw in name.lower() or kw in cid.lower() or kw in plate.lower()
        if ok:
            result.append(r)

    st.info(f"พบ {len(result)} รายการ")
    for i, r in enumerate(result):
        with st.expander(f"📄 {r.get('name','-')} | {r.get('vehicle_plate','-')}"):
            st.write(f"เลขบัตร: {r.get('sub_id','-')}")
            st.write(f"วันที่: {r.get('report_date','-')}")
            st.write(f"ข้อหา: {r.get('charge','-')}")
            c1, c2 = st.columns(2)
            if c1.button("🗑️ ลบ", key=f"del_{i}"):
                st.session_state.records.remove(r)
                st.rerun()
            if c2.button("📋 โหลด", key=f"load_{i}"):
                st.session_state.form = r
                st.session_state.page = "form"
                st.session_state.tab  = 3
                st.rerun()

    st.write("")
    if st.button("⬅️ กลับหน้าหลัก", use_container_width=True):
        st.session_state.page = "dashboard"
        st.rerun()


# =================================================================
# ─── DOCUMENTS ──────────────────────────────────────────────────
# =================================================================
if st.session_state.page == "documents":
    if "doc_records" not in st.session_state:
        st.session_state.doc_records = []
    if "doc_running_numbers" not in st.session_state:
        st.session_state.doc_running_numbers = {}

    DOC_TYPES = {
        "หนังสือราชการ": {"prefix":"ที่",    "icon":"📋"},
        "บันทึกข้อความ": {"prefix":"บันทึก", "icon":"📝"},
        "คำสั่ง":         {"prefix":"คำสั่ง", "icon":"📌"},
        "ประกาศ":         {"prefix":"ประกาศ", "icon":"📢"},
        "รายงาน":         {"prefix":"รายงาน", "icon":"📊"},
        "หนังสือรับ":     {"prefix":"รับที่",  "icon":"📥"},
        "หนังสือส่ง":     {"prefix":"ส่งที่",  "icon":"📤"},
    }

    DOC_OFFICERS = OFFICERS.copy()

    def get_next_doc_number(doc_type, year_th):
        key = f"{doc_type}_{year_th}"
        n = st.session_state.doc_running_numbers.get(key, 0) + 1
        st.session_state.doc_running_numbers[key] = n
        return f"0018.13/{n:04d}/{year_th}"

    def get_preview_number(doc_type, year_th):
        key = f"{doc_type}_{year_th}"
        n = st.session_state.doc_running_numbers.get(key, 0) + 1
        return f"0018.13/{n:04d}/{year_th}"

    st.markdown('<div class="main-header">📂&nbsp; ระบบจัดการเอกสารราชการ</div>', unsafe_allow_html=True)

    sub_page = st.session_state.get("doc_sub_page", "create")

    # ── CREATE ──
    if sub_page == "create":

        if "doc_type_preselect" not in st.session_state or st.session_state.get("doc_type_show_picker", True):
            st.markdown("## 📂 เลือกประเภทเอกสาร")
            st.write("")
            cols = st.columns(4)
            for idx, (dtype, meta) in enumerate(DOC_TYPES.items()):
                col = cols[idx % 4]
                if col.button(
                    f"{meta['icon']}  {dtype}",
                    use_container_width=True,
                    key=f"pick_doc_{dtype}"
                ):
                    st.session_state["doc_type_preselect"] = dtype
                    st.session_state["doc_type_show_picker"] = False
                    st.rerun()

        else:
            preselect = st.session_state.get("doc_type_preselect", "หนังสือราชการ")
            if preselect not in DOC_TYPES:
                preselect = "หนังสือราชการ"
            meta    = DOC_TYPES[preselect]
            year_th = datetime.date.today().year + 543

            if st.button("⬅️ เปลี่ยนประเภทเอกสาร", key="btn_change_type"):
                st.session_state["doc_type_show_picker"] = True
                st.rerun()

            st.markdown(f"## {meta['icon']} {preselect}")

            ca, cb = st.columns([2,1])
            with ca:
                auto_no = st.toggle("ออกเลขอัตโนมัติ", value=True, key="tog_auto_no")
                if auto_no:
                    doc_no = get_preview_number(preselect, year_th)
                    st.info(f"เลขที่จะออก: **{doc_no}**")
                else:
                    doc_no = st.text_input("เลขที่เอกสาร", placeholder="0018.13/0001/2568", key="doc_manual_no")
            with cb:
                doc_date    = st.date_input("วันที่เอกสาร", datetime.date.today(), key="doc_date")
                doc_urgency = st.selectbox("ชั้นความเร็ว", ["ปกติ","ด่วน","ด่วนมาก","ด่วนที่สุด"], key="doc_urgency")

            st.write("---")
            cc, cd = st.columns(2)
            with cc:
                doc_to   = st.text_input("เรียน / ถึง", key="doc_to", placeholder="เช่น ผู้กำกับการ สภ.ตระการพืชผล")
                doc_from = st.text_input("จาก", value=st.session_state.get("user_full_name","เจ้าหน้าที่"), key="doc_from")
                doc_ref  = st.text_input("อ้างถึง (ถ้ามี)", key="doc_ref")
            with cd:
                doc_subject = st.text_input("เรื่อง", key="doc_subject")
                doc_attach  = st.text_input("สิ่งที่ส่งมาด้วย", key="doc_attach")

            doc_body = st.text_area("เนื้อหา", height=200, key="doc_body", placeholder="พิมพ์เนื้อหาที่นี่...")
            st.write("---")
            st.caption("✍️ ผู้ลงนาม")
            ce, cf, cg = st.columns(3)
            signer     = ce.selectbox("ผู้ลงนาม", DOC_OFFICERS, key="doc_signer")
            signer_pos = cf.text_input("ตำแหน่ง", key="doc_signer_pos", placeholder="เช่น ผกก.สภ.ตระการพืชผล")
            doc_dept   = cg.text_input("หน่วยงาน", value="สภ.ตระการพืชผล", key="doc_dept")
            st.write("---")

            b1, b2 = st.columns(2)
            with b1:
                if st.button("💾 บันทึกเอกสาร", use_container_width=True, key="btn_save_doc"):
                    if not doc_subject:
                        st.warning("⚠️ กรุณากรอกเรื่อง")
                    else:
                        final_no = get_next_doc_number(preselect, year_th) if auto_no else doc_no
                        st.session_state.doc_records.append({
                            "doc_number": final_no, "doc_type": preselect,
                            "doc_date": str(doc_date), "doc_subject": doc_subject,
                            "doc_to": doc_to, "doc_from": doc_from, "doc_ref": doc_ref,
                            "doc_attach": doc_attach, "doc_body": doc_body,
                            "doc_urgency": doc_urgency, "signer": signer,
                            "signer_pos": signer_pos, "doc_dept": doc_dept,
                            "created_by": st.session_state.get("user_full_name","-"),
                            "created_at": str(datetime.datetime.now()),
                            "status": "หนังสือรับ" if preselect == "หนังสือรับ" else "หนังสือส่ง",
                        })
                        st.success(f"✅ บันทึกแล้ว เลขที่ {final_no}")
            with b2:
                st.button("📄 Export Word", use_container_width=True, key="btn_exp_word", disabled=True)

 # ── SEARCH ──
    elif sub_page == "search":
        st.markdown("### 🔎 ค้นหาเอกสาร")
        drecs = st.session_state.get("doc_records",[])
        sc1, sc2, sc3 = st.columns([3,1,1])
        kw    = sc1.text_input("ค้นหา เลขที่ / เรื่อง / ถึง", key="srch_kw")
        ftype = sc2.selectbox("ประเภท", ["ทั้งหมด"] + list(DOC_TYPES.keys()), key="srch_type")
        fyear = sc3.selectbox("ปี พ.ศ.", ["ทั้งหมด", str(datetime.date.today().year+543)], key="srch_year")

        results = [
            d for d in drecs
            if (not kw or any(kw.lower() in str(d.get(f,"")).lower() for f in ["doc_number","doc_subject","doc_to"]))
            and (ftype == "ทั้งหมด" or d.get("doc_type") == ftype)
            and (fyear == "ทั้งหมด" or fyear in d.get("doc_number",""))
        ]
        st.info(f"พบ {len(results)} รายการ")
        for i, d in enumerate(results):
            m = DOC_TYPES.get(d.get("doc_type",""),{"icon":"📄"})
            with st.expander(f"{m['icon']} {d.get('doc_number','-')} | {d.get('doc_subject','-')}"):
                xa, xb = st.columns(2)
                xa.write(f"**ประเภท:** {d.get('doc_type','-')}")
                xa.write(f"**เรียน:** {d.get('doc_to','-')}")
                xb.write(f"**ชั้นความเร็ว:** {d.get('doc_urgency','-')}")
                xb.write(f"**ผู้ลงนาม:** {d.get('signer','-')}")
                if d.get("doc_body"):
                    st.text_area("เนื้อหา", value=d["doc_body"], height=80, key=f"sbody_{i}", disabled=True)
                if st.button("🗑️ ลบ", key=f"sdel_{i}"):
                    st.session_state.doc_records.remove(d)
                    st.rerun()

    # ── REGISTER ──
    elif sub_page == "register":
        st.markdown("### 📬 ทะเบียนรับ-ส่งเอกสาร")
        drecs    = st.session_state.get("doc_records",[])
        sent     = [d for d in drecs if d.get("status") == "หนังสือส่ง"]
        received = [d for d in drecs if d.get("status") == "หนังสือรับ"]

        def render_table(docs, label):
            if not docs:
                st.warning(f"ยังไม่มี{label}")
                return
            tbl = pd.DataFrame([{
                "เลขที่":   d.get("doc_number","-"),
                "วันที่":   d.get("doc_date","-"),
                "ประเภท":   d.get("doc_type","-"),
                "เรื่อง":   d.get("doc_subject","-"),
                "ถึง/จาก":  d.get("doc_to","-") or d.get("doc_from","-"),
                "ผู้บันทึก": d.get("created_by","-"),
            } for d in docs])
            st.dataframe(tbl, use_container_width=True, hide_index=True)
            csv = tbl.to_csv(index=False).encode("utf-8-sig")
            st.download_button(f"📥 Export CSV ({label})", csv,
                               file_name=f"register_{label}_{datetime.date.today()}.csv",
                               mime="text/csv", key=f"exp_{label}")

        rt1, rt2 = st.tabs([f"📤 ทะเบียนส่ง ({len(sent)})", f"📥 ทะเบียนรับ ({len(received)})"])
        with rt1: render_table(sent, "หนังสือส่ง")
        with rt2: render_table(received, "หนังสือรับ")

    st.write("---")
    if st.button("⬅️ กลับหน้าหลัก", use_container_width=True, key="btn_back_docs"):
        st.session_state.page = "dashboard"
        st.rerun()


# =================================================================
# ─── SCHEDULE ───────────────────────────────────────────────────
# =================================================================
if st.session_state.page == "schedule":
    import calendar

    st.markdown('<div class="main-header">📅&nbsp; ตารางเวรประจำเดือน</div>', unsafe_allow_html=True)

    if "schedule_data" not in st.session_state:
        st.session_state.schedule_data = {}

    SHIFT_TYPES = {
        "เวรกลางวัน (06:00-18:00)": "🌤️",
        "เวรกลางคืน (18:00-06:00)": "🌙",
        "วันหยุด": "🏖️",
        "ลา": "📝",
    }

    # ── เลือกเดือน/ปี ──
    col_m, col_y = st.columns([2, 2])
    now = datetime.date.today()

    month_names = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

    sel_month = col_m.selectbox("เดือน", list(range(1,13)),
                                format_func=lambda x: month_names[x],
                                index=now.month - 1, key="sch_month")
    sel_year  = col_y.number_input("ปี พ.ศ.", value=now.year + 543,
                                   min_value=2560, max_value=2580, key="sch_year")
    sel_year_ad = sel_year - 543

    # ── แจ้งเตือนเวรวันนี้ ──
    today_key = str(now)
    today_shifts = {
        o: v for o, v in st.session_state.schedule_data.items()
        if isinstance(v, dict) and v.get(today_key)
    }
    if today_shifts:
        st.markdown("### 🔔 เวรวันนี้")
        for officer, days in today_shifts.items():
            shift = days.get(today_key, "")
            icon  = SHIFT_TYPES.get(shift, "")
            st.info(f"{icon} **{officer}** — {shift}")
    else:
        st.info(f"🔔 วันนี้ ({date_th(now)}) ยังไม่มีข้อมูลเวร")

    st.write("---")

    # ── กำหนดเวรรายวัน ──
    st.markdown("### ✏️ กำหนดเวร")

    col_a, col_b, col_c, col_d = st.columns([2, 2, 2, 1])
    sch_officer = col_a.selectbox("เลือกเจ้าหน้าที่", OFFICERS, key="sch_officer")
    sch_date    = col_b.date_input("วันที่", value=now, key="sch_date")
    sch_shift   = col_c.selectbox("ประเภทเวร", list(SHIFT_TYPES.keys()), key="sch_shift")

    if col_d.button("✅ บันทึก", use_container_width=True, key="sch_save"):
        if sch_officer not in st.session_state.schedule_data:
            st.session_state.schedule_data[sch_officer] = {}
        st.session_state.schedule_data[sch_officer][str(sch_date)] = sch_shift
        st.success(f"บันทึกเวร {sch_officer} วันที่ {date_th(sch_date)} แล้ว")

    st.write("---")

    # ── ตารางรายเดือน ──
    st.markdown(f"### 📋 ตารางเวรเดือน{month_names[sel_month]} {sel_year}")

    _, num_days = calendar.monthrange(sel_year_ad, sel_month)
    days_in_month = [datetime.date(sel_year_ad, sel_month, d) for d in range(1, num_days + 1)]

    rows = []
    for officer in OFFICERS:
        row = {"เจ้าหน้าที่": officer}
        officer_data = st.session_state.schedule_data.get(officer, {})
        for d in days_in_month:
            shift = officer_data.get(str(d), "")
            icon  = SHIFT_TYPES.get(shift, "")
            row[str(d.day)] = f"{icon}" if shift else "-"
        rows.append(row)

    df_sch = pd.DataFrame(rows)
    st.dataframe(df_sch, use_container_width=True, hide_index=True)

    st.write("---")

    # ── Export ──
    st.markdown("### 📤 Export ตารางเวร")

    ex1, ex2 = st.columns(2)

    with ex1:
        csv_rows = []
        for officer in OFFICERS:
            officer_data = st.session_state.schedule_data.get(officer, {})
            for d in days_in_month:
                shift = officer_data.get(str(d), "")
                if shift:
                    csv_rows.append({
                        "เจ้าหน้าที่": officer,
                        "วันที่": date_th(d),
                        "ประเภทเวร": shift,
                    })
        if csv_rows:
            df_export = pd.DataFrame(csv_rows)
            csv_bytes  = df_export.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "📥 Export CSV",
                data=csv_bytes,
                file_name=f"ตารางเวร_{month_names[sel_month]}_{sel_year}.csv",
                mime="text/csv",
                key="sch_export_csv"
            )
        else:
            st.warning("ยังไม่มีข้อมูลในเดือนนี้")

    with ex2:
        if st.button("📄 Export Word", use_container_width=True, key="sch_export_word"):
            try:
                from docx import Document as DocxDoc
                from docx.shared import Pt
                wdoc = DocxDoc()
                wdoc.add_heading(f"ตารางเวรเดือน{month_names[sel_month]} {sel_year}", level=1)
                table = wdoc.add_table(rows=1, cols=num_days + 1)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                hdr[0].text = "เจ้าหน้าที่"
                for i, d in enumerate(days_in_month):
                    hdr[i+1].text = str(d.day)
                for officer in OFFICERS:
                    officer_data = st.session_state.schedule_data.get(officer, {})
                    row_cells = table.add_row().cells
                    row_cells[0].text = officer
                    for i, d in enumerate(days_in_month):
                        shift = officer_data.get(str(d), "")
                        row_cells[i+1].text = SHIFT_TYPES.get(shift, "") if shift else "-"
                buf = BytesIO()
                wdoc.save(buf)
                buf.seek(0)
                st.download_button(
                    "📥 ดาวน์โหลด Word",
                    data=buf,
                    file_name=f"ตารางเวร_{month_names[sel_month]}_{sel_year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="sch_dl_word"
                )
            except Exception as e:
                st.error(f"❌ {e}")

    st.write("---")
    if st.button("⬅️ กลับหน้าหลัก", use_container_width=True, key="sch_back"):
        st.session_state.page = "dashboard"
        st.rerun()


# =================================================================
# ─── FORM PAGE ──────────────────────────────────────────────────
# =================================================================
if st.session_state.page != "form":
    st.stop()

# ── NAV ──
st.markdown('<div class="main-header">📋&nbsp; บันทึกการจับกุม</div>', unsafe_allow_html=True)

tab_labels = ["📍 ข้อมูลบันทึก","👮 เจ้าหน้าที่","👤 ผู้ต้องหา","📝 รายละเอียด"]

n1, n2, n3, n4 = st.columns(4)

for i, (col, lbl) in enumerate(zip([n1,n2,n3,n4], tab_labels)):
    with col:
        if st.button(lbl, key=f"nav_tab_{i}", use_container_width=True):
            st.session_state.tab = i
            st.rerun()
st.write("---")
# ======================TAB 1: ข้อมูลบันทึก ====================
# ===========================================================
if st.session_state.tab == 0:
    st.write("")

    col1, col2 = st.columns(2)

    with col1:
        # ================= สถานที่ทำบันทึก (NEW) =================
        record_loc_options = [
            "สภ.ตระการพืชผล ภ.จว.อุบลราชธานี",
            "สภ.เมืองอุบลราชธานี ภ.จว.อุบลราชธานี",
            "สภ.วารินชำราบ ภ.จว.อุบลราชธานี",
            "อื่นๆ"
        ]

        default_loc = st.session_state.form.get(
            "record_loc",
            "สภ.ตระการพืชผล ภ.จว.อุบลราชธานี"
        )

        record_loc_sel = st.selectbox(
            "สถานที่ทำบันทึก",
            record_loc_options,
            index=record_loc_options.index(default_loc) if default_loc in record_loc_options else 0
        )

        # 👉 ถ้าเลือกอื่นๆ ให้เด้ง input
        if record_loc_sel == "อื่นๆ":
            record_loc = st.text_input(
                "ระบุสถานที่ทำบันทึก",
                st.session_state.form.get("record_loc_custom", "")
            )
            st.session_state.form["record_loc_custom"] = record_loc
        else:
            record_loc = record_loc_sel

        # ================= วันที่จับกุม =================
        incident_date = st.date_input(
            "วันที่จับกุม",
            value=st.session_state.form.get('incident_date', datetime.date.today())
        )

        # ================= เวลาจับกุม =================
        arrest_time = st.text_input(
            "เวลาจับกุม (เช่น 10:30)",
            value=str(st.session_state.form.get('arrest_time', "10:00"))
        )

    with col2:
        # ================= วันที่ลงบันทึก =================
        report_date = st.date_input(
            "วันที่ลงบันทึก",
            value=st.session_state.form.get('report_date', datetime.date.today())
        )

        # ================= เวลาลงบันทึก =================
        tz_thai = pytz.timezone("Asia/Bangkok")
        current_time = datetime.datetime.now(tz_thai).replace(second=0, microsecond=0).time()

        # ล็อคเวลาบันทึกไว้ครั้งแรก ไม่ให้เปลี่ยนแปลง
        if "report_time_locked" not in st.session_state:
            st.session_state["report_time_locked"] = current_time

        report_time = st.session_state["report_time_locked"]

        st.text_input(
            "เวลาลงบันทึก",
            value=report_time.strftime("%H:%M"),
            disabled=True,
            key="report_time_display"
        )
 # ================= SAVE SESSION =================
    st.session_state.form.update({
        "record_loc": record_loc,
        "incident_date": incident_date,
        "arrest_time": arrest_time,
        "report_date": report_date,
    })

    st.write("---")

    # ================= NEXT BUTTON =================
    if st.button("ถัดไป ➔", use_container_width=True, type="primary"):
        st.session_state.tab = 1
        st.rerun()

# --- TAB 2: ผู้บังคับบัญชาและชุดจับกุม ---
if st.session_state.tab == 1:
    st.write("")

    # ================= ผู้บังคับบัญชา =================
    commander_1 = st.text_input(
        "ผกก.",
        st.session_state.form.get(
            'commander_1',
            "พ.ต.อ.เกรียงศักดิ์ ปัชโชโต ผกก.สภ.ตระการพืชผล"
        )
    )

    commander_2 = st.text_input(
        "รอง ผกก.",
        st.session_state.form.get(
            'commander_2',
            "พ.ต.ท.ยุทธนา กิติชัยชนานนท์ รอง ผกก.ป.สภ.ตระการพืชผล"
        )
    )

    commander_3 = st.text_input(
        "สวป.",
        st.session_state.form.get(
            'commander_3',
            "พ.ต.ต.รัฐ โสมสุพรรณ สวป.สภ.ตระการพืชผล"
        )
    )

    st.write("---")

    # =====================================================
    # 🔥 เลือกชุดจับกุมแบบใหม่
    # =====================================================
    st.markdown("### 👮‍♂️ เลือกชุดจับกุม")

    mode = st.radio(
        "เลือกรูปแบบชุด",
        ["ชุดที่ 1", "ชุดที่ 2", "เลือกชุดเอง"],
        horizontal=True
    )

    sub_team = [
        "ร.ต.ท.พจน์ ปาสาจันทร์",
        "ร.ต.ต.ประจวบ ศรีบุระ",
        "ส.ต.อ.ธีระวัฒน์ แก่นสาร",
        "ส.ต.อ.ณรงค์ฤทธิ์ เหล่าดี",
        "ส.ต.อ.วิชชากร วงษ์โท",
        "ด.ต.ยุทธพงษ์ ชาดแดง",
        "จ.ส.ต.วัชระ จันสุตะ",
        "ส.ต.ท.วชิรวิชญ์ นันทรักษ์",
        "ส.ต.ท.วิศรุต จันทร์สิงห์",
        "ส.ต.ท.เอกพจน์ อินผล",
        "ส.ต.ต.อดิศร ศุภนิกร"
    ]

    clean_selected = []

    # ================= ชุดที่ 1 =================
    if mode == "ชุดที่ 1":

        clean_selected = [
            "ร.ต.อ.สมคิด เชื้อเวียง",
            "ร.ต.อ.ชัยพร ชอบงาม",
            *sub_team
        ]

    # ================= ชุดที่ 2 =================
    elif mode == "ชุดที่ 2":

        clean_selected = [
            "ร.ต.อ.ชัยพร ชอบงาม",
            "ร.ต.อ.สมคิด เชื้อเวียง",
            *sub_team
        ]

    # ================= เลือกเอง =================
    else:

        leader = st.selectbox(
            "เลือกหัวหน้าชุด",
            [
                "ร.ต.อ.สมคิด เชื้อเวียง",
                "ร.ต.อ.ชัยพร ชอบงาม"
            ]
        )

        st.markdown("#### 👥 เลือกลูกชุด")

        cols = st.columns(2)
        selected_sub = []

        # 🔥 ใช้ OFFICERS ทั้งหมด (ยกเว้นคนที่เป็นหัวหน้า)
        custom_team = [x for x in OFFICERS if x != leader]

        for i, officer in enumerate(custom_team):
            col = cols[i % 2]

            if col.checkbox(officer, key=f"chk_custom_{i}"):
                selected_sub.append(officer)

        clean_selected = [leader] + selected_sub
        
    # ================= บันทึกลง session =================
    st.session_state.form.update({
        "commander_1": commander_1,
        "commander_2": commander_2,
        "commander_3": commander_3,
        "selected": clean_selected
    })

    # ================= แสดงผล =================
    st.caption(f"เลือกแล้ว {len(clean_selected)} นาย")

    st.text_area(
        "รายชื่อชุดจับกุม",
        value="\n".join(clean_selected),
        height=260
    )

    # ================= ปุ่ม =================
    c1, c2 = st.columns(2)

    with c1:
        if st.button("⬅️ ย้อนกลับ", use_container_width=True):
            st.session_state.tab = 0
            st.rerun()

    with c2:
        if st.button("ถัดไป ➔", use_container_width=True):
            st.session_state.tab = 2
            st.rerun()

# ================= TAB 3 (Index 2): ข้อมูลผู้ต้องหา =================
if st.session_state.tab == 2:
    st.write("")
    # ================= BASIC INFO =================
    col_fname, col_lname, col_id, col_age = st.columns([2,2,2,1])

    first_name = col_fname.text_input(
        "ชื่อ", 
        value=st.session_state.form.get("first_name", ""),  # ← เพิ่ม
        key="input_fname_t3"
    )

    last_name = col_lname.text_input(
        "นามสกุล",
        value=st.session_state.form.get("last_name", ""),   # ← เพิ่ม
        key="input_lname_t3"
    )

    name = f"{first_name} {last_name}"

    sub_id = col_id.text_input(
        "เลขบัตรประชาชน",
        st.session_state.form.get('sub_id', ''),
        key="input_subid_t3"
    )

    age = col_age.number_input(
        "อายุ",
        0, 120,
        st.session_state.form.get('age', 0),
        key="input_age_t3"
    )

    st.write("---")
    st.caption("🏠 ข้อมูลที่อยู่")

    # ================= GEO DATA =================
    df_geo = load_geo()
    # 🔥 safe index กันพัง selectbox
    def safe_index(value, options):
        return options.index(value) if value in options else 0


    # ================= ADDRESS =================
    col_h, col_p = st.columns([1,2])

    house = col_h.text_input(
        "บ้านเลขที่/หมู่",
        st.session_state.form.get('house',''),
        key="input_house_t3"
    )

    provinces = sorted(df_geo['ProvinceThai'].dropna().unique())
    def_prov = st.session_state.form.get('province', provinces[0])

    province = col_p.selectbox(
        "จังหวัด",
        provinces,
        index=safe_index(def_prov, provinces),
        key="sel_prov_t3"
    )

    # ================= DISTRICT =================
    col_a, col_t = st.columns(2)

    districts = sorted(
        df_geo[df_geo['ProvinceThai'] == province]['DistrictThai'].dropna().unique()
    )

    def_amphur = st.session_state.form.get('amphur', '')

    amphur = col_a.selectbox(
        "อำเภอ/เขต",
        districts,
        index=safe_index(def_amphur, districts),
        key="sel_amphur_t3"
    )

    # ================= TAMBON =================
    sub_districts = sorted(
        df_geo[
            (df_geo['ProvinceThai'] == province) &
            (df_geo['DistrictThai'] == amphur)
        ]['TambonThai'].dropna().unique()
    )

    def_tumbon = st.session_state.form.get('tumbon', '')

    tumbon = col_t.selectbox(
        "ตำบล/แขวง",
        sub_districts,
        index=safe_index(def_tumbon, sub_districts),
        key="sel_tumbon_t3"
    )

    # ================= VEHICLE =================
    st.write("---")
    st.caption("🚗 ข้อมูลยานพาหนะ")

    # 🔥 โหลดค่าล่าสุด (ต้องอยู่ก่อน selectbox)
    if "last_vehicle_brand" in st.session_state:
        if not st.session_state.form.get("vehicle_brand"):
            st.session_state.form["vehicle_brand"] = st.session_state["last_vehicle_brand"]

    if "last_vehicle_model" in st.session_state:
        if not st.session_state.form.get("vehicle_model"):
            st.session_state.form["vehicle_model"] = st.session_state["last_vehicle_model"]

    # ✅ ต้องมีบรรทัดนี้ก่อนใช้ col1 !!!
    col1, col2, col3 = st.columns(3)
            
    # ----- ประเภท -----
    v_types = ["รถยนต์นั่งส่วนบุคคล","รถจักรยานยนต์","รถกระบะ","รถตู้","อื่นๆ"]
    def_type = st.session_state.form.get('vehicle_type','รถยนต์นั่งส่วนบุคคล')

    vehicle_type_sel = col1.selectbox(
        "ประเภทรถ",
        v_types,
        index=v_types.index(def_type) if def_type in v_types else 0,
        key="sel_vtype_t3"
    )

    vehicle_type_custom = ""
    if vehicle_type_sel == "อื่นๆ":
        vehicle_type_custom = col1.text_input(
            "ระบุประเภทรถ",
            st.session_state.form.get("vehicle_type_custom",""),
            key="input_vtype_custom_t3"
        )

    vehicle_type = vehicle_type_custom if vehicle_type_sel == "อื่นๆ" else vehicle_type_sel

    # ----- ยี่ห้อ -----
    v_brands = ["Toyota","Honda","Isuzu","Nissan","Mazda","Yamaha","อื่นๆ"]
    def_brand = st.session_state.form.get('vehicle_brand','Toyota')

    vehicle_brand_sel = col2.selectbox(
        "ยี่ห้อรถ",
        v_brands,
        index=v_brands.index(def_brand) if def_brand in v_brands else 0,
        key="sel_vbrand_t3"
    )

    vehicle_brand_custom = ""
    if vehicle_brand_sel == "อื่นๆ":
        vehicle_brand_custom = col2.text_input(
            "ระบุยี่ห้อรถ",
            st.session_state.form.get("vehicle_brand_custom",""),
            key="input_vbrand_custom_t3"
        )

    vehicle_brand = vehicle_brand_custom if vehicle_brand_sel == "อื่นๆ" else vehicle_brand_sel

        # ================= MODEL (SMART) =================
    st.write("---")
    st.caption("🚘 รุ่นรถ")

    # 🔥 เอารุ่นตามยี่ห้อ
    brand_for_model = vehicle_brand

    model_list = VEHICLE_MODELS.get(brand_for_model, [])

    # 👉 ถ้าไม่มีในระบบ ให้ fallback
    if not model_list:
        model_list = ["ไม่ทราบรุ่น", "อื่นๆ"]
    else:
        model_list = model_list + ["อื่นๆ"]

    # 👉 default ค่าเดิม
    def_model = st.session_state.form.get('vehicle_model', model_list[0])

    vehicle_model_sel = st.selectbox(
        "รุ่นรถ",
        model_list,
        index=model_list.index(def_model) if def_model in model_list else 0,
        key="sel_vmodel_t3"
    )

    vehicle_model_custom = ""

    if vehicle_model_sel == "อื่นๆ":
        vehicle_model_custom = st.text_input(
            "ระบุรุ่นรถ",
            st.session_state.form.get("vehicle_model_custom", ""),
            key="input_vmodel_custom_t3"
        )

    vehicle_model = vehicle_model_custom if vehicle_model_sel == "อื่นๆ" else vehicle_model_sel

    # ----- สี -----
    v_colors = ["ขาว","ดำ","เทา","แดง","น้ำเงิน","อื่นๆ"]
    def_color = st.session_state.form.get('vehicle_color','ขาว')

    vehicle_color_sel = col3.selectbox(
        "สีรถ",
        v_colors,
        index=v_colors.index(def_color) if def_color in v_colors else 0,
        key="sel_vcolor_t3"
    )

    vehicle_color_custom = ""
    if vehicle_color_sel == "อื่นๆ":
        vehicle_color_custom = col3.text_input(
            "ระบุสีรถ",
            st.session_state.form.get("vehicle_color_custom",""),
            key="input_vcolor_custom_t3"
        )

    vehicle_color = vehicle_color_custom if vehicle_color_sel == "อื่นๆ" else vehicle_color_sel

    # ================= ทะเบียนรถ =================
    st.write("---")
    st.caption("🔢 หมายเลขทะเบียน")

    plate_option = st.selectbox(
        "สถานะแผ่นป้ายทะเบียน",
        ["กรุณาเลือก", "มีแผ่นป้ายทะเบียน", "ไม่ติดแผ่นป้ายทะเบียน"],
        index=0,
        key="sel_plate_option_t3"
    )

    plate_letter = ""
    plate_number = ""
    plate_province = ""
    vehicle_plate = ""

    if plate_option == "มีแผ่นป้ายทะเบียน":

        c1, c2, c3 = st.columns([1,2,2])

        plate_letter = c1.text_input("ตัวอักษร", "", key="input_pletter_t3")
        plate_number = c2.text_input("ตัวเลข", "", key="input_pnumber_t3")

        def_plate = st.session_state.form.get('plate_province', provinces[0])

        plate_province = c3.selectbox(
            "จังหวัด",
            provinces,
            index=provinces.index(def_plate) if def_plate in provinces else 0,
            key="sel_pprov_t3"
        )

        vehicle_plate = f"{plate_letter} {plate_number} {plate_province}"
        st.success(f"ทะเบียน: {vehicle_plate}")

    elif plate_option == "ไม่ติดแผ่นป้ายทะเบียน":
        vehicle_plate = "ไม่ติดแผ่นป้ายทะเบียน"
        st.warning("🚫 ไม่มีแผ่นป้ายทะเบียน")

    else:
        vehicle_plate = ""
        st.info("ℹ️ กรุณาเลือกสถานะทะเบียน")

    # ================= LICENSE =================
    st.write("---")
    st.caption("🪪 ใบขับขี่")

    license_options = ["มี", "ไม่มี", "มี (แต่สิ้นอายุ)"]

    def_license = st.session_state.form.get("license_status", "มี")

    license_status = st.selectbox(
        "ใบขับขี่",
        license_options,
        index=license_options.index(def_license) if def_license in license_options else 0,
        key="sel_license_t3"
    )

    # ================= CASE =================
    st.write("---")
    st.caption("⚖️ ข้อหา")

    charge_list = [
        "เป็นผู้ขับขี่ (รถยนต์) ในขณะเมาสุรา",
        "เป็นผู้ขับขี่ (จักรยานยนต์) ในขณะเมาสุรา",
        "อื่นๆ"
    ]

    charge_sel = st.selectbox("ข้อหา", charge_list, key="sel_charge_t3")

    charge_custom = ""
    if charge_sel == "อื่นๆ":
        charge_custom = st.text_input("กรอกข้อหา", "", key="input_charge_custom_t3")

    final_charge = charge_custom if charge_sel == "อื่นๆ" else charge_sel

    # ================= DEVICE =================
    device_brand_list = ["ไลออน อัลโคลมิเตอร์", "อื่นๆ"]

    device_brand_sel = st.selectbox("ยี่ห้อเครื่องวัด", device_brand_list, key="sel_device_brand_t3")

    device_brand_custom = ""
    if device_brand_sel == "อื่นๆ":
        device_brand_custom = st.text_input("กรอกยี่ห้อ", "", key="input_device_brand_custom_t3")

    final_device_brand = device_brand_custom if device_brand_sel == "อื่นๆ" else device_brand_sel

    device_serial_list = ["T29542", "อื่นๆ"]

    device_serial_sel = st.selectbox("หมายเลขเครื่อง", device_serial_list, key="sel_device_serial_t3")

    device_serial_custom = ""
    if device_serial_sel == "อื่นๆ":
        device_serial_custom = st.text_input("กรอกหมายเลขเครื่อง", "", key="input_device_serial_custom_t3")

    final_device_serial = device_serial_custom if device_serial_sel == "อื่นๆ" else device_serial_sel

    alcohol = st.number_input("แอลกอฮอล์ (มก.%)", value=0, key="num_alc_t3")

    # ================= UPDATE =================
    st.session_state.form.update({
        "name": name,
        "sub_id": sub_id,
        "age": age,
        "house": house,
        "province": province,
        "amphur": amphur,
        "tumbon": tumbon,
        "vehicle_plate": vehicle_plate,
        "plate_letter": plate_letter,
        "plate_number": plate_number,
        "plate_province": plate_province,

        "vehicle_type": vehicle_type,
        "vehicle_brand": vehicle_brand,
        "vehicle_model": vehicle_model,
        "vehicle_color": vehicle_color,

        "plate_option": plate_option,

        "license_status": license_status,

        "charge": final_charge,

        "device_brand": final_device_brand,
        "device_serial": final_device_serial,

        "alcohol": alcohol
    })
    st.session_state["last_vehicle_brand"] = vehicle_brand
    st.session_state["last_vehicle_model"] = vehicle_model
    c1, c2 = st.columns(2)

    if c1.button("⬅️ ย้อนกลับ", key="btn_back_t3"):
        st.session_state.tab = 1
        st.rerun()

    if c2.button("ถัดไป ➡️", key="btn_next_t3"):
        st.session_state.tab = 3
        st.rerun()

# ================= TAB 4 =================
if st.session_state.tab == 3:
    st.write("")
    f = st.session_state.form
    
    # ================= FUNCTION: Thai Number =================
    def to_arabic_number(text):
        if text is None:
            return ""

        thai_to_arabic = {
            "๐":"0","๑":"1","๒":"2","๓":"3","๔":"4",
            "๕":"5","๖":"6","๗":"7","๘":"8","๙":"9"
        }

        text = str(text)

        return "".join(thai_to_arabic.get(ch, ch) for ch in text)
    
    # ================= จุดเกิดเหตุ =================
    st.caption("📌 จุดเกิดเหตุ")

    spot_options = [
        "บริเวณสี่แยกตู้ยามตระการร่วมใจ (เขตเทศบาลตระการพืชผล) หมู่ 4",
        "บริเวณสำนักงานเขตพื้นที่การศึกษาประถมศึกษาอุบลราชธานีเขต2 (เขตเทศบาลตำบลตระการพืชผล) หมู่4",
        "บริเวณหน้าสำนักงานที่ดินอำเภอตระการพืชผล (เขตเทศบาลตระการพืชผล) หมู่ 2",
        "บริเวณสี่แยกธนาคารกรุงเทพตระการพืชผล(เขตเทศบาลตำบลตระการพืชผล) หมู่ 2",
        "อื่นๆ"
    ]

    current_spot = f.get("spot_location", spot_options[0])

    selected_spot = st.selectbox(
        "เลือกจุดเกิดเหตุ",
        spot_options,
        index=spot_options.index(current_spot) if current_spot in spot_options else 0,
        key="sel_spot_location"
    )

    if selected_spot == "อื่นๆ":
        spot_custom = st.text_input(
            "ระบุจุดเกิดเหตุ",
            value=f.get("spot_location_custom", "")
        )
        selected_spot = spot_custom

    st.session_state.form["spot_location"] = selected_spot

    st.caption("📍 รายละเอียดสถานที่เกิดเหตุ")

    # ================= จุดเกิดเหตุ =================
    loc_options = [
        "ตั้งจุดตรวจบริเวณสี่แยกตู้ยามตระการร่วมใจ หมู่ 4",
        "ตั้งจุดตรวจบริเวณสำนักงานเขตพื้นที่การศึกษาประถมศึกษาอุบลราชธานีเขต2 หมู่4",
        "ตั้งจุดตรวจบริเวณหน้าสำนักงานที่ดินอำเภอตระการพืชผล หมู่ 2",
        "อำนวยความสะดวกการจราจรบริเวณสี่แยกตู้ยามตระการร่วมใจ หมู่ 4",
        "อำนวยความสะดวกการจราจรบริเวณสี่แยกธนาคารกรุงเทพตระการพืชผลหมู่ 2",
        "อื่นๆ"
    ]

    current_loc_detail = f.get('record_loc_detail', loc_options[0])
    default_loc_idx = loc_options.index(current_loc_detail) if current_loc_detail in loc_options else 0

    selected_loc = st.selectbox(
        "เลือกสถานที่เกิดเหตุ",
        loc_options,
        index=default_loc_idx
    )

    if selected_loc == "อื่นๆ":
        final_loc_detail = st.text_input(
            "ระบุสถานที่เกิดเหตุอื่นๆ",
            f.get('record_loc_detail_custom', '')
        )
    else:
        final_loc_detail = selected_loc

    # ================= จังหวัด/อำเภอ/ตำบล =================
    df_geo = load_geo()

    col_p, col_a, col_t = st.columns(3)

    provinces = sorted(df_geo['ProvinceThai'].dropna().unique())

    incident_province = col_p.selectbox(
        "จังหวัดที่เกิดเหตุ",
        provinces,
        index=provinces.index(f.get('incident_province')) if f.get('incident_province') in provinces else 0,
        key="incident_province"
    )

    districts = sorted(
        df_geo[df_geo['ProvinceThai'] == incident_province]['DistrictThai'].dropna().unique()
    )

    incident_amphur = col_a.selectbox(
        "อำเภอที่เกิดเหตุ",
        districts if districts else ["-"],
        index=districts.index(f.get('incident_amphur')) if f.get('incident_amphur') in districts else 0,
        key="incident_amphur"
    )

    sub_districts = sorted(
        df_geo[
            (df_geo['ProvinceThai'] == incident_province) &
            (df_geo['DistrictThai'] == incident_amphur)
        ]['TambonThai'].dropna().unique()
    )

    incident_tumbon = col_t.selectbox(
        "ตำบลที่เกิดเหตุ",
        sub_districts if sub_districts else ["-"],
        index=sub_districts.index(f.get('incident_tumbon')) if f.get('incident_tumbon') in sub_districts else 0,
        key="incident_tumbon"
    )

    st.write("---")

    # ================= จุดเชิญตัว =================
    invite_options = ["ตู้ยามตระการร่วมใจ", "สภ.ตระการพืชผล", "อื่นๆ"]
    current_invite = f.get('invite_loc', invite_options[0])
    default_invite_idx = invite_options.index(current_invite) if current_invite in invite_options else 0

    selected_invite = st.selectbox(
        "📍 จุดเชิญตัวผู้ต้องหา",
        invite_options,
        index=default_invite_idx
    )

    if selected_invite == "อื่นๆ":
        invite_loc = st.text_input("ระบุจุดเชิญตัวอื่นๆ", f.get('invite_loc_custom', ''))
    else:
        invite_loc = selected_invite

    confession_option = st.radio(
        "คำให้การผู้ต้องหา:",
        ["รับสารภาพตลอดข้อกล่าวหา", "ให้การปฏิเสธทุกข้อกล่าวหา"],
        index=0 if f.get('confession_status') == "รับสารภาพ" else 1,
        horizontal=True
    )

    col_d_manual, col_t_manual = st.columns(2)

    arrest_date_val = col_d_manual.date_input(
        "เมื่อวันที่ (แจ้งสิทธิ์)",
        value=f.get("arrest_date_manual", datetime.date.today())
    )

    arrest_time_val = col_t_manual.time_input(
        "เวลาประมาณ (แจ้งสิทธิ์)",
        value=f.get("arrest_time_manual", datetime.datetime.now().time())
    )

    st.caption("👮 ผู้เกี่ยวข้องในคดี")
    col_acc, col_wit, col_reader, col_proc = st.columns([2, 2, 2, 1.5])

    accuser_name = col_acc.selectbox("ผู้กล่าวหา", OFFICERS, index=0)
    witness_name = col_wit.selectbox("พยาน", OFFICERS, index=0)
    selected_officers = st.session_state.form.get("selected", [])

    # 👉 กันกรณีว่าง / หรือถูกเลือกหมด
    available_readers = [
        o for o in OFFICERS
        if o not in selected_officers
    ]

    # 🔥 fallback สำคัญ (กัน No options)
    if not available_readers:
        available_readers = OFFICERS.copy()
        st.warning("⚠️ ไม่มีรายชื่อว่าง ระบบจะให้เลือกทั้งหมดแทน")

    reader_name = col_reader.selectbox(
        "พยาน/ผู้บันทึกอ่าน",
        available_readers,
        key="reader_select_t4"
    )

    prosecutor_options = ["082-921-211", "อื่นๆ"]
    current_phone = f.get("prosecutor_phone", "082-921-211")
    phone_index = prosecutor_options.index(current_phone) if current_phone in prosecutor_options else 0

    selected_phone = col_proc.selectbox("📞 เบอร์อัยการ", prosecutor_options, index=phone_index)

    if selected_phone == "อื่นๆ":
        prosecutor_phone = st.text_input("กรอกเบอร์อัยการ", f.get("prosecutor_phone_custom", ""))
    else:
        prosecutor_phone = selected_phone

    st.write("---")

    final_detail = st.text_area(
        "ตรวจสอบ/แก้ไข พฤติการณ์จับกุม",
        value=f.get("detail", ""),
        height=250
    )

    # ================= UPDATE SESSION =================
    st.session_state.form.update({
        "accuser_name": accuser_name,
        "witness_name": witness_name,
        "reader_name": reader_name,
        "prosecutor_phone": prosecutor_phone,
        "detail": final_detail,
        "confession_status": confession_option,
        "arrest_date_manual": arrest_date_val,
        "arrest_time_manual": arrest_time_val,
        "record_loc_detail": final_loc_detail,
        "invite_loc": invite_loc,
        "incident_province": incident_province,
        "incident_amphur": incident_amphur,
        "incident_tumbon": incident_tumbon,
        "spot_location_custom": f.get("spot_location_custom", "")
    })

    # ================= RESET =================
    if st.session_state.get("confirm_reset", False):
        st.warning("⚠️ ต้องการล้างข้อมูลทั้งหมดใช่หรือไม่?")

        c1, c2 = st.columns(2)

        if c1.button("✅ ยืนยัน"):
            st.session_state.form = {}
            st.session_state.tab = 0
            st.session_state.confirm_reset = False
            st.rerun()

        if c2.button("❌ ยกเลิก"):
            st.session_state.confirm_reset = False
            st.rerun()

# ================= EXPORT =================
if st.session_state.tab == 3:
    st.write("---")

    left, right = st.columns(2)

    with left:
        if st.button("🔄 เริ่มต้นบันทึกใหม่", key="reset_btn", use_container_width=True):
            st.session_state.confirm_reset = True

    with right:
        if st.button("💾 สร้างเอกสารWord", key="export_btn", use_container_width=True):
            st.session_state.export_now = True

    if st.session_state.get("export_now", False):
        f = st.session_state.form
        required_fields = {
            "ชื่อผู้ต้องหา": f.get("name", "").strip(),
            "ข้อหา": f.get("charge", "").strip(),
            "ทะเบียนรถ": f.get("vehicle_plate", "").strip(),
        }
        missing = [k for k, v in required_fields.items() if not v]
        if missing:
            st.error(f"❌ กรุณากรอกข้อมูลให้ครบ: {', '.join(missing)}")
            st.session_state.export_now = False
        else:
            with st.spinner("⏳ กำลังสร้างเอกสาร..."):
                try:
                    selected = f.get("selected", [])
                    reader = f.get("reader_name", "")
                    selected_clean = [x for x in selected if x != reader]
                    sign_data = {}

                    for i in range(12):
                        if i < len(selected_clean):
                            sign_data[f"{{{{s{i+1}}}}}"] = "(ลงชื่อ) ................................................................."
                            sign_data[f"{{{{n{i+1}}}}}"] = f"({selected_clean[i]})"
                        else:
                            sign_data[f"{{{{s{i+1}}}}}"] = ""
                            sign_data[f"{{{{n{i+1}}}}}"] = ""

                    if reader:
                        sign_data["{{s13}}"] = "(ลงชื่อ) ................................................................."
                        sign_data["{{n13}}"] = f"({reader})"
                    else:
                        sign_data["{{s13}}"] = ""

                    vehicle_plate = f"{f.get('plate_letter','')}{f.get('plate_number','')} {f.get('plate_province','')}".strip()

                    incident_loc = " ".join(filter(None, [
                        f.get("spot_location", ""),
                        f"ต.{format_tambon(f.get('incident_tumbon',''))}" if f.get("incident_tumbon") else "",
                        f"อ.{format_amphur(f.get('incident_amphur',''))}" if f.get("incident_amphur") else "",
                        f"จ.{f.get('incident_province','')}" if f.get("incident_province") else ""
                    ])).strip()

                    def to_arabic_number(text):
                        if text is None:
                            return ""
                        thai_to_arabic = {
                            "๐":"0","๑":"1","๒":"2","๓":"3","๔":"4",
                            "๕":"5","๖":"6","๗":"7","๘":"8","๙":"9"
                        }
                        return "".join(thai_to_arabic.get(ch, ch) for ch in str(text))

                    data = {
                        "{{record_loc}}": f.get("record_loc", ""),
                        "{{location}}": f.get("record_loc_detail", ""),
                        "{{date}}": date_th(f.get("incident_date")),
                        "{{time}}": safe_time(f.get("arrest_time")),
                        "{{officer_name}}": f.get("accuser_name", ""),
                        "{{incident_date}}": date_th(f.get("incident_date")),
                        "{{arrest_time}}": safe_time(f.get("arrest_time")),
                        "{{report_date}}": date_th(f.get("report_date")),
                        "{{report_time}}": safe_time(f.get("report_time")),
                        "{{commander_1}}": f.get("commander_1", ""),
                        "{{commander_2}}": f.get("commander_2", ""),
                        "{{commander_3}}": f.get("commander_3", ""),
                        "{{officers_list}}": ", ".join(f.get("selected", [])),
                        "{{sub_name}}": f.get("name", ""),
                        "{{reader_name}}": f.get("reader_name", ""),
                        "{{age}}": str(f.get("age", "")),
                        "{{alcohol}}": to_arabic_number(f.get("alcohol", "")),
                        "{{sub_id}}": format_thai_id(f.get("sub_id", "")),
                        "{{sub_full_addr}}": f"{f.get('house','')} ต.{format_tambon(f.get('tumbon',''))} อ.{format_amphur(f.get('amphur',''))} จ.{f.get('province','')}",
                        "{{charge}}": f.get("charge", ""),
                        "{{incident_loc}}": incident_loc,
                        "{{incident_loc_detail}}": f.get("record_loc_detail", ""),
                        "{{tumbon}}": clean_geo(f.get("incident_tumbon", "")),
                        "{{amphur}}": clean_geo(f.get("incident_amphur", "")),
                        "{{province}}": clean_geo(f.get("incident_province", "")),
                        "{{incident_full_location}}": " ".join(filter(None, [
                            f"ต.{clean_geo(f.get('incident_tumbon',''))}" if f.get("incident_tumbon") else "",
                            f"อ.{format_amphur(f.get('incident_amphur',''))}" if f.get("incident_amphur") else "",
                            f"จ.{clean_geo(f.get('incident_province',''))}" if f.get("incident_province") else ""
                        ])).strip(),
                        "{{vehicle_type}}": f.get("vehicle_type", ""),
                        "{{vehicle_brand}}": f.get("vehicle_brand", ""),
                        "{{vehicle_model}}": f.get("vehicle_model", "-"),
                        "{{vehicle_color}}": f.get("vehicle_color", ""),
                        "{{vehicle_plate}}": vehicle_plate,
                        "{{license_status}}": f.get("license_status", ""),
                        "{{spot_location}}": " ".join(filter(None, [
                            f.get("spot_location", ""),
                            f"ต.{format_tambon(f.get('incident_tumbon',''))}" if f.get("incident_tumbon") else "",
                            f"อ.{format_amphur(f.get('incident_amphur',''))}" if f.get("incident_amphur") else "",
                            f"จ.{f.get('incident_province','')}" if f.get("incident_province") else ""
                        ])).strip(),
                        "{{invite_loc}}": f.get("invite_loc", ""),
                        "{{device_brand}}": f.get("device_brand", ""),
                        "{{device_serial}}": f.get("device_serial", ""),
                        "{{confession}}": f.get("confession_status", ""),
                        "{{statement}}": " ".join([
                            str(f.get("confession_status", "")),
                            str(f.get("detail", ""))
                        ]).strip(),
                        "{{send_date}}": date_th(f.get("report_date")),
                        "{{prosecutor_phone}}": f.get("prosecutor_phone", ""),
                        "{{accuser_name}}": f.get("accuser_name", ""),
                        "{{witness_name}}": f.get("witness_name", ""),
                    }

                    data.update(sign_data)
                    for k, v in data.items():
                        data[k] = to_arabic_number(v)

                    if not os.path.exists("template.docx"):
                        st.error("❌ ไม่พบไฟล์ template.docx")
                        st.session_state.export_now = False
                        st.stop()

                    doc = Document("template.docx")
                    replace_text(doc, data)

                    for p in doc.paragraphs:
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing = 1

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    p.paragraph_format.space_before = Pt(0)
                                    p.paragraph_format.space_after = Pt(0)
                                    p.paragraph_format.line_spacing = 1

                    buffer = BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    st.download_button(
                        "📥 ดาวน์โหลดWord",
                        buffer,
                        file_name=f"{f.get('name','arrest_report')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.session_state.export_now = False

                except Exception as e:
                    st.error(f"❌ {e}")
                    st.session_state.export_now = False
