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
import time
# =================================================================
# LOGO weid
# =================================================================
st.set_page_config(
    page_title="ระบบบันทึกจับกุม",
    page_icon="Thai_Police.png",
    layout="wide",
    initial_sidebar_state="expanded"
)
# ================= LOGIN USERS =================
USERS = {
    "admin": {
        "password": "123",
        "fullname": "Admin"
    },

    "somkid": {
        "password": "1234",
        "fullname": "ร.ต.อ.สมคิด เชื้อเวียง"
    },

    "chaiporn": {
        "password": "1234",
        "fullname": "ร.ต.อ.ชัยพร ชอบงาม"
    },

    "teerawat": {
        "password": "1234",
        "fullname": "ส.ต.อ.ธีระวัฒน์ แก่นสาร"
    },

    "narongrit": {
        "password": "1234",
        "fullname": "ส.ต.อ.ณรงค์ฤทธิ์ เหล่าดี"
    },

    "wichakorn": {
        "password": "1234",
        "fullname": "ส.ต.อ.วิชชากร วงษ์โท"
    },

    "watchara": {
        "password": "1234",
        "fullname": "จ.ส.ต.วัชระ จันสุตะ"
    },

    "wachirawit": {
        "password": "1234",
        "fullname": "ส.ต.ท.วชิรวิชญ์ นันทรักษ์"
    },

    "wisarut": {
        "password": "1234",
        "fullname": "ส.ต.ท.วิศรุต จันทร์สิงห์"
    },

    "adisorn": {
        "password": "1234",
        "fullname": "ส.ต.ต.อดิศร ศุภนิกร"
    }
}
# =================================================================
# 1. INITIALIZE SESSION STATE
# =================================================================
if "settings" not in st.session_state:
    st.session_state.settings = {
        "username": st.session_state.get("user_full_name", "เจ้าหน้าที่ตำรวจ"),
        "bg_color": "#eef4ff",
        "logo": None
    }

if "form" not in st.session_state:
    st.session_state.form = {}

if "tab" not in st.session_state:
    st.session_state.tab = 0

if "records" not in st.session_state:
    st.session_state.records = []

# =========================================================
# THEME SYSTEM
# =========================================================
theme = st.session_state.get("settings", {}).get("theme", "Light")

if theme == "Dark":
    app_bg = "#0b1f4e"
    text_color = "#ffffff"
    card_bg = "#111827"
    input_bg = "#111827"

elif theme == "Blue Pro":
    app_bg = "linear-gradient(180deg,#0b1f4e,#123a8f)"
    text_color = "#ffffff"
    card_bg = "#111827"
    input_bg = "#111827"

else:
    app_bg = "linear-gradient(180deg,#f5f9ff,#eef4ff)"
    text_color = "#111827"
    card_bg = "#ffffff"
    input_bg = "#ffffff"
# =================================================================
# 2. HELPER FUNCTIONS (FINAL STABLE PRO)
# =================================================================
def date_th(d):
    if not d:
        return "-"
    months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
              "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    try:
        return f"{d.day} {months[d.month]} {d.year + 543}"
    except:
        return str(d)


def safe_time(t):
    if isinstance(t, (datetime.time, datetime.datetime)):
        return t.strftime("%H:%M")
    return str(t) if t else "-"

# ❌ ไม่ใช้
def format_signature_line(line, name_text=""):
    return line

def remove_empty_signature_rows(doc):
    """
    🔥 ลบแถวลายเซ็นที่ไม่มีชื่อจริง (เช็คทั้งแถว)
    """
    for table in doc.tables:
        rows_to_remove = []

        for row in table.rows:
            if len(row.cells) < 3:
                continue

            left = row.cells[0].text.strip()
            center = row.cells[1].text.strip()
            right = row.cells[2].text.strip()

            # 🔥 เงื่อนไขใหม่ (แม่นกว่าเดิม)
            if (
                right == "" and
                "ผู้จับกุม" in center
            ):
                rows_to_remove.append(row)

        for row in rows_to_remove:
            table._tbl.remove(row._tr)

def center_signature_lines(doc):
    """
    ✔ ล็อค "(ลงชื่อ)" ให้อยู่กลางคอลัมน์
    """
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) < 3:
                continue

            for para in row.cells[0].paragraphs:
                if "ลงชื่อ" in para.text:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def tighten_signature_spacing(doc):
    """
    ✔ spacing แน่น สวย ไม่ลอย
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                    para.paragraph_format.line_spacing = 1


def split_rank_name(fullname):
    """
    ✔ แยกยศ + ชื่อ
    """
    ranks = [
        "พ.ต.อ.", "พ.ต.ท.", "พ.ต.ต.",
        "ร.ต.อ.", "ร.ต.ท.", "ร.ต.ต.",
        "ด.ต.", "จ.ส.ต.",
        "ส.ต.อ.", "ส.ต.ท.", "ส.ต.ต."
    ]

    fullname = str(fullname).strip()

    for rank in ranks:
        if fullname.startswith(rank):
            name = fullname[len(rank):].strip()
            return rank, name

    return "", fullname
def replace_text(doc, data):

    def smart_replace(paragraph):
        full_text = "".join(run.text for run in paragraph.runs)

        for key, value in data.items():
            if key in full_text:
                full_text = full_text.replace(key, str(value))

        # 🔥 เขียนทับทั้งย่อหน้า (กันพัง)
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

def format_thai_id(id_number):
    if not id_number:
        return ""

    digits = "".join(filter(str.isdigit, str(id_number)))

    if len(digits) != 13:
        return id_number  # ถ้าไม่ครบ 13 หลัก ไม่แตะ

    return f"{digits[0]}-{digits[1:5]}-{digits[5:10]}-{digits[10:12]}-{digits[12]}"
# ✅ 🔥 วางตรงนี้เลย
def clean_geo(text):
    """
    ฟังก์ชันทำความสะอาดชื่อจังหวัด/อำเภอ/ตำบล
    - ลบช่องว่างเกิน
    - ลบคำที่ไม่จำเป็น เช่น 'จังหวัด', 'อำเภอ', 'ตำบล'
    """
    if not text:
        return ""
    text = str(text).strip()
    for word in ["จังหวัด", "อำเภอ", "ตำบล"]:
        text = text.replace(word, "")
    return text.strip()

def format_amphur(name):
    if not name:
        return ""
    name = str(name)
    for x in ["อำเภอ", "อ.", "เขต"]:
        name = name.replace(x, "")
    return name.strip()

def format_tambon(name):
    if not name:
        return ""
    return str(name).replace("ตำบล", "").strip()

@st.cache_data
def load_geo():
    try:
        df = pd.read_csv("thai_districts.csv", encoding="utf-8-sig")
    except:
        df = pd.read_csv("thai_districts.csv", encoding="cp874")

    # 🔥 CLEAN ทั้งระบบครั้งเดียว
    df["ProvinceThai"] = df["ProvinceThai"].apply(clean_geo)
    df["DistrictThai"] = df["DistrictThai"].apply(clean_geo)
    df["TambonThai"] = df["TambonThai"].apply(clean_geo)

    df["DistrictThai"] = df["DistrictThai"].apply(format_amphur)
    df["TambonThai"] = df["TambonThai"].apply(format_tambon)

    return df

# ================= VEHICLE MODEL MAP (PRO) =================
VEHICLE_MODELS = {
    "Honda": ["Wave 110i", "Wave 125i", "Click 125i", "PCX", "ADV160"],
    "Yamaha": ["NMAX", "Aerox", "Grand Filano", "Fino"],
    "Toyota": ["Vios", "Yaris", "Hilux Revo", "Fortuner"],
    "Isuzu": ["D-Max", "MU-X"],
    "Nissan": ["Navara", "Almera"],
    "Mazda": ["Mazda2", "Mazda3", "BT-50"]
}

# =========================================================
# 3. LOGO BASE64 (SAFE)
# =========================================================
def get_base64(path):
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
logo_path = os.path.join(BASE_DIR, "police_logo.png")
logo_base64 = st.session_state.settings.get("logo") or get_base64(logo_path)

# =========================================================
# 4. CSS (ONCE ONLY - FIXED + DARK MODE)
# =========================================================
st.markdown(f"""
<style>
header[data-testid="stHeader"]{{
    background:transparent !important;
    box-shadow:none !important;
}}

[data-testid="stToolbar"],
[data-testid="stDeployButton"]{{
    display:none !important;
}}

[data-testid="collapsedControl"]{{
    position:fixed !important;
    top:15px !important;
    left:15px !important;
    z-index:9999 !important;
    display:flex !important;
    align-items:center;
    justify-content:center;
    width:42px !important;
    height:42px !important;
    background:linear-gradient(135deg,#123a8f,#0b2159) !important;
    color:white !important;
    border-radius:12px !important;
    box-shadow:0 6px 14px rgba(0,0,0,0.25) !important;
    cursor:pointer !important;
}}

@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;700;800&display=swap');

html, body, [class*="css"]{{
    font-family:'Sarabun', sans-serif;
}}

/* ================= THEME COLORS ================= */
.stApp {{
    background:{app_bg} !important;
}}

h1,h2,h3,h4,h5,h6 {{
    color:{text_color} !important;
}}

p, span, div {{
    color:{text_color};
}}

label {{
    color:{text_color} !important;
    font-weight:600 !important;
}}

.stMarkdown, .stMarkdown p {{
    color:{text_color} !important;
}}

.stTextInput input,
.stDateInput input,
.stTimeInput input,
.stTextArea textarea,
.stNumberInput input {{
    color:{text_color} !important;
    background:{input_bg} !important;
    border:1px solid rgba(128,128,128,0.3) !important;
}}

.stSelectbox div[data-baseweb="select"] > div {{
    color:{text_color} !important;
    background:{input_bg} !important;
}}

.stTextInput label,
.stDateInput label,
.stTimeInput label,
.stSelectbox label,
.stTextArea label,
.stNumberInput label {{
    color:{text_color} !important;
    font-weight:700 !important;
}}

input::placeholder, textarea::placeholder {{
    color:{"#9ca3af" if theme != "Light" else "#6b7280"} !important;
    opacity:1 !important;
}}

button[data-baseweb="tab"] {{
    color:{text_color} !important;
}}

.streamlit-expanderHeader {{
    color:{text_color} !important;
}}

[data-testid="stDataFrame"] {{
    background:{card_bg} !important;
}}

/* ================= BLOCK ================= */
.block-container {{
    max-width:1450px !important;
    padding-top:1rem;
    padding-bottom:2rem;
    position:relative;
    z-index:1;
}}

/* ================= SIDEBAR ================= */
[data-testid="stSidebar"] {{
    background:#0b1f4e !important;
    color:white !important;
    border-right:1px solid rgba(255,255,255,0.08);
    min-width:320px !important;
    max-width:320px !important;
}}

[data-testid="stSidebar"] * {{
    color:#ffffff !important;
}}

/* ================= LOGIN ================= */
.logo-img {{
    width:160px;
    display:block;
    margin:auto;
    animation:spinlogo 4s linear infinite;
    transform-style:preserve-3d;
}}

@keyframes spinlogo {{
    from{{transform:rotateY(0deg);}}
    to{{transform:rotateY(360deg);}}
}}

.title-main {{
    text-align:center;
    font-size:40px;
    font-weight:800;
    color:{"#0b1f4e" if theme == "Light" else "#ffffff"};
    margin-top:10px;
}}

.title-sub {{
    text-align:center;
    font-size:18px;
    color:{"#555" if theme == "Light" else "#cccccc"};
    margin-bottom:20px;
}}

/* ================= BUTTON ================= */
div.stButton > button {{
    width:100% !important;
    border:none;
    border-radius:14px;
    padding:12px 14px;
    font-size:16px;
    font-weight:700;
    background:linear-gradient(135deg,#123a8f,#0b2159);
    color:white !important;
    box-shadow:0 8px 18px rgba(0,0,0,0.12);
}}

div.stButton > button:hover {{
    transform:translateY(-2px);
}}

/* ================= HEADER BAR ================= */
.main-header {{
    background:linear-gradient(135deg,#0b1f4e,#123a8f);
    padding:14px 20px;
    border-radius:16px;
    display:flex;
    align-items:center;
    justify-content:center;
    gap:15px;
    color:white !important;
    font-weight:800;
    font-size:22px;
    box-shadow:0 10px 25px rgba(0,0,0,0.15);
}}

.main-header * {{
    color:white !important;
}}

.header-line {{
    height:2px;
    background:#d9e6ff;
    margin-top:10px;
    margin-bottom:20px;
}}

/* ================= KPI CARD ================= */
.kpi-card {{
    background:linear-gradient(135deg,#111827,#1e3a8a);
    padding:22px;
    border-radius:22px;
    color:white;
    box-shadow:0 10px 30px rgba(0,0,0,0.25);
}}

.kpi-card * {{
    color:white !important;
}}

.kpi-number {{
    font-size:42px;
    font-weight:800;
    color:white !important;
}}

.kpi-label {{
    opacity:0.8;
    color:white !important;
}}

/* ================= METRIC CARD ================= */
.metric-card {{
    background:linear-gradient(135deg,#0f172a,#1e3a8a);
    border-radius:22px;
    padding:22px;
    text-align:center;
    box-shadow:0 10px 25px rgba(0,0,0,0.18);
    transition:0.25s;
    color:white;
    position:relative;
    overflow:hidden;
}}

.metric-card * {{
    color:white !important;
}}

.metric-card:hover {{
    transform:translateY(-4px);
    box-shadow:0 12px 25px rgba(0,0,0,0.22);
}}

.metric-icon {{ font-size:26px; margin-bottom:8px; }}
.metric-no {{ font-size:42px; font-weight:800; color:#ffffff !important; }}
.metric-text {{ font-size:15px; color:rgba(255,255,255,0.85) !important; font-weight:500; }}

/* ================= PLOTLY ================= */
.js-plotly-plot {{
    border-radius:18px !important;
    overflow:hidden !important;
    box-shadow:0 10px 30px rgba(0,0,0,0.25);
    background:#111827 !important;
    padding:10px;
}}

hr {{
    border:none;
    border-top:1px solid {"#d8e1f2" if theme == "Light" else "rgba(255,255,255,0.1)"};
}}

</style>
""", unsafe_allow_html=True)

# ================= LOGIN UI =================
if "password_correct" not in st.session_state:

    st.markdown('<div class="login-card">', unsafe_allow_html=True)

    if logo_base64:
        st.markdown(
            f'<img class="logo-img" src="data:image/png;base64,{logo_base64}">',
            unsafe_allow_html=True
        )

    st.markdown("""
    <div class="title-main">ยินดีต้อนรับ</div>
    <div class="title-sub">
        ระบบบันทึกการจับกุม งานจราจร<br>
        สภ.ตระการพืชผล
    </div>
    """, unsafe_allow_html=True)

    # ================= INPUT CENTER =================
    # ================= INPUT CENTER =================
    c1, c2, c3 = st.columns([1.4,2,1])

    with c2:
        u = st.text_input(
            "👤 Username",
            key="login_user"
        )

        p = st.text_input(
            "🔒 Password",
            type="password",
            key="login_pass"
        )
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "เข้าสู่ระบบ",
            use_container_width=True,
            key="login_button_2"
        ):
            user_data = USERS.get(u)
            if user_data and user_data["password"] == p:
                st.session_state["password_correct"] = True
                st.session_state["user_full_name"] = user_data["fullname"]
                st.session_state.page = "dashboard"
                st.rerun()
            else:
                st.error("❌ ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")

    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


# =========================================================
# DEFAULT PAGE
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "dashboard"

# ================= DASHBOARD DATA (FIX ERROR) =================
records = st.session_state.get("records", [])

total = len(records)

today = len([
    r for r in records
    if str(r.get("report_date")) == str(datetime.date.today())
])

printed = total
latest = records[-1].get("name", "-") if len(records) > 0 else "-"


# =========================================================
# DASHBOARD PAGE PRO
# =========================================================
if st.session_state.get("page") == "dashboard":

    st.markdown("""
    <style>

    .dashboard-title{
        font-size:34px;
        font-weight:800;
        color:white;
    }

    .kpi-card{
        background:linear-gradient(135deg,#111827,#1e3a8a);
        padding:22px;
        border-radius:22px;
        color:white;
        box-shadow:0 10px 30px rgba(0,0,0,0.25);
    }

    .kpi-number{
        font-size:42px;
        font-weight:800;
    }

    .kpi-label{
        opacity:0.8;
    }

    </style>
    """, unsafe_allow_html=True)

    records = st.session_state.get("records", [])
    doc_records = st.session_state.get("doc_records", [])

    total = len(records)

    today = len([
        r for r in records
        if str(r.get("report_date")) == str(datetime.date.today())
    ])

    printed = total

    latest = "-"
    if records:
        latest = records[-1].get("name", "-")

    # =====================================================
    # HEADER
    # =====================================================

    st.markdown("""
    <div class="main-header">
        🚔 POLICE REALTIME DASHBOARD
    </div>
    """, unsafe_allow_html=True)

    st.write("")

    # =====================================================
    # KPI CARDS
    # =====================================================

    c1,c2,c3,c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{total}</div>
            <div class="kpi-label">คดีทั้งหมด</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{today}</div>
            <div class="kpi-label">วันนี้</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{printed}</div>
            <div class="kpi-label">เอกสารทั้งหมด</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-number">{latest}</div>
            <div class="kpi-label">ล่าสุด</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")

    # =====================================================
    # SAMPLE DATA
    # =====================================================

    if len(records) == 0:

        sample = []

        charges = [
            "เมาแล้วขับ",
            "ไม่มีใบขับขี่",
            "ขับเร็ว",
            "ฝ่าไฟแดง",
            "ไม่สวมหมวก"
        ]

        for i in range(60):

            sample.append({
                "report_date": datetime.date.today() - datetime.timedelta(days=np.random.randint(0,30)),
                "charge": np.random.choice(charges),
                "hour": np.random.randint(0,24)
            })

        df = pd.DataFrame(sample)

    else:

        temp = []

        for r in records:

            dt = r.get("report_date", datetime.date.today())

            try:
                d = pd.to_datetime(dt)
            except:
                d = pd.to_datetime(datetime.date.today())

            temp.append({
                "report_date": d,
                "charge": r.get("charge","อื่นๆ"),
                "hour": np.random.randint(0,24)
            })

        df = pd.DataFrame(temp)

    # =====================================================
    # CREATE CHARTS
    # =====================================================

    # ---------- MONTH ----------
    df["report_date"] = pd.to_datetime(
        df["report_date"],
        errors="coerce"
    )

    df["month"] = df["report_date"].dt.strftime("%m")

    month_df = (
        df.groupby("month")
        .size()
        .reset_index(name="count")
    )

    fig_month = px.bar(
        month_df,
        x="month",
        y="count",
        text_auto=True,
        title="📈 สถิติคดีรายเดือน"
    )

    # ---------- CHARGE ----------
    charge_df = (
        df.groupby("charge")
        .size()
        .reset_index(name="count")
    )

    fig_charge = px.bar(
        charge_df,
        x="charge",
        y="count",
        color="charge",
        title="🚨 สถิติแยกตามข้อหา"
    )

    # ---------- PIE ----------
    fig_pie = px.pie(
        charge_df,
        names="charge",
        values="count",
        title="🥧 สัดส่วนข้อหา"
    )

    # ---------- HEATMAP ----------
    heat = (
        df.groupby("hour")
        .size()
        .reset_index(name="count")
    )

    fig_heat = px.density_heatmap(
        heat,
        x="hour",
        y="count",
        title="🔥 Heatmap เวลาเกิดเหตุ"
    )

    # ---------- REALTIME ----------
    realtime = pd.DataFrame({
        "เวลา": pd.date_range(
            start=datetime.datetime.now(),
            periods=20,
            freq="min"
        ),
        "คดี": np.random.randint(1,10,20)
    })

    fig_real = px.line(
        realtime,
        x="เวลา",
        y="คดี",
        title="📡 Realtime Monitoring"
    )

    # ---------- DOCUMENT ----------
    fig_doc = None

    if len(doc_records) > 0:

        doc_df = pd.DataFrame(doc_records)

        if "doc_type" in doc_df.columns:

            dfg = (
                doc_df.groupby("doc_type")
                .size()
                .reset_index(name="count")
            )

            fig_doc = px.bar(
                dfg,
                x="doc_type",
                y="count",
                color="doc_type",
                title="📂 เอกสารราชการ"
            )

    # =====================================================
    # DASHBOARD GRID LAYOUT
    # =====================================================

    # ---------- ROW 1 ----------
    g1, g2 = st.columns([1.4,1])

    with g1:

        fig_month.update_layout(
            height=420,
            template="plotly_dark" if theme != "Light" else "plotly_white",
            paper_bgcolor="#111827",
            plot_bgcolor="#111827",
            font_color="white"
        )

        st.plotly_chart(
            fig_month,
            use_container_width=True
        )

    with g2:

        fig_pie.update_layout(
            height=420,
            template="plotly_dark" if theme != "Light" else "plotly_white",
            paper_bgcolor="#111827",
            font_color="white"
        )

        st.plotly_chart(
            fig_pie,
            use_container_width=True
        )

    # ---------- ROW 2 ----------
    g3, g4 = st.columns(2)

    with g3:

        fig_charge.update_layout(
            height=420,
            template="plotly_dark" if theme != "Light" else "plotly_white",
            paper_bgcolor="#111827",
            plot_bgcolor="#111827",
            font_color="white"
        )

        st.plotly_chart(
            fig_charge,
            use_container_width=True
        )

    with g4:

        fig_heat.update_layout(
            height=420,
            template="plotly_dark" if theme != "Light" else "plotly_white",
            paper_bgcolor="#111827",
            plot_bgcolor="#111827",
            font_color="white"
        )

        st.plotly_chart(
            fig_heat,
            use_container_width=True
        )

    # ---------- ROW 3 ----------

    fig_real.update_layout(
        height=350,
        template="plotly_dark" if theme != "Light" else "plotly_white",
        paper_bgcolor="#111827",
        plot_bgcolor="#111827",
        font_color="white"
    )

    st.plotly_chart(
        fig_real,
        use_container_width=True
    )

    # ---------- DOCUMENT ----------

    if fig_doc is not None:

        fig_doc.update_layout(
            height=350,
            template="plotly_dark" if theme != "Light" else "plotly_white",
            paper_bgcolor="#111827",
            plot_bgcolor="#111827",
            font_color="white"
        )

        st.plotly_chart(
            fig_doc,
            use_container_width=True
        )

    #=================================================
    # AUTO REFRESH
    # =====================================================

    st.caption("ระบบ Realtime Refresh ทุกครั้งที่ Reload")

# =====================ระบบจัดการเอกสาร=================
    st.markdown("### 📄 ประเภทเอกสาร")

    doc_types_meta = {
        "หนังสือราชการ":  "📋",
        "บันทึกข้อความ":  "📝",
        "คำสั่ง":          "📌",
        "ประกาศ":          "📢",
        "รายงาน":          "📊",
        "หนังสือรับ":      "📥",
        "หนังสือส่ง":      "📤",
    }

    doc_records = st.session_state.get("doc_records", [])

    # นับแยกตามประเภท
    type_counts = {t: 0 for t in doc_types_meta}
    for d in doc_records:
        t = d.get("doc_type", "")
        if t in type_counts:
            type_counts[t] += 1

    # แถวการ์ด 4 ใบ
    dc1, dc2, dc3, dc4 = st.columns(4)
    doc_cols_row1 = [dc1, dc2, dc3, dc4]

    dc5, dc6, dc7, dc8 = st.columns(4)
    doc_cols_row2 = [dc5, dc6, dc7, dc8]

    all_doc_cols = doc_cols_row1 + doc_cols_row2
    doc_type_list = list(doc_types_meta.items())

    for idx, (dtype, icon) in enumerate(doc_type_list):
        col = all_doc_cols[idx]
        count = type_counts.get(dtype, 0)
        col.markdown(f"""
        <div class='metric-card' style='cursor:pointer;'>
            <div class='metric-icon'>{icon}</div>
            <div class='metric-no'>{count}</div>
            <div class='metric-text'>{dtype}</div>
        </div>
        """, unsafe_allow_html=True)

    # ปุ่มใหญ่ "จัดการเอกสารทั้งหมด"
    st.write("")
    if st.button("📂 จัดการเอกสารทั้งหมด", use_container_width=True, key="btn_goto_docs"):
        st.session_state.doc_type_preselect = None
        st.session_state.page = "documents"
        st.rerun()
    
# =========================================================
# SETTINGS PAGE (PRO VERSION)
# =========================================================
if st.session_state.get("page") == "settings":
    st.markdown("## ⚙️ ตั้งค่าระบบระดับโปร")

    s = st.session_state.settings

    tab1, tab2, tab3, tab4 = st.tabs([
        "🏢 หน่วยงาน",
        "🎨 ธีมระบบ",
        "💾 สำรองข้อมูล",
        "⚙️ ขั้นสูง"
    ])

    # ---------------- TAB 1 ----------------
    with tab1:

        c1, c2 = st.columns(2)

        with c1:
            s["username"] = st.text_input(
                "ชื่อผู้ใช้งาน",
                s.get("username", "เจ้าหน้าที่ตำรวจ")
            )

            s["station"] = st.text_input(
                "ชื่อสถานีตำรวจ",
                s.get("station", "สภ.ตระการพืชผล")
            )

            s["province"] = st.text_input(
                "จังหวัด",
                s.get("province", "อุบลราชธานี")
            )

        with c2:
            s["commander"] = st.text_input(
                "ชื่อผู้กำกับ",
                s.get("commander", "")
            )

            s["deputy"] = st.text_input(
                "รองผู้กำกับ",
                s.get("deputy", "")
            )

            s["inspector"] = st.text_input(
                "สารวัตร",
                s.get("inspector", "")
            )

        st.write("---")

        logo_file = st.file_uploader(
            "อัปโหลดโลโก้",
            type=["png", "jpg", "jpeg"]
        )

        if logo_file:
            logo_bytes = logo_file.read()
            s["logo"] = base64.b64encode(logo_bytes).decode()
            st.image(logo_file, width=140)

    # ---------------- TAB 2 ----------------
    with tab2:

        c1, c2 = st.columns(2)

        with c1:
            s["bg_color"] = st.color_picker(
                "สีพื้นหลัง",
                s.get("bg_color", "#eef4ff")
            )

            s["button_color"] = st.color_picker(
                "สีปุ่ม",
                s.get("button_color", "#123a8f")
            )

        with c2:
            s["header_color"] = st.color_picker(
                "สีหัวข้อ",
                s.get("header_color", "#0b1f4e")
            )

            theme_list = ["Light", "Dark", "Blue Pro"]

            current_theme = s.get("theme", "Light")
            if current_theme not in theme_list:
                current_theme = "Light"

            s["theme"] = st.selectbox(
                "โหมดระบบ",
                theme_list,
                index=theme_list.index(current_theme)
            )

    # ---------------- TAB 3 ----------------
    with tab3:

        records = st.session_state.get("records", [])

        st.info(f"ข้อมูลทั้งหมด {len(records)} รายการ")

        json_data = json.dumps(
            records,
            ensure_ascii=False,
            indent=2,
            default=str
        )

        st.download_button(
            "📥 ดาวน์โหลด Backup",
            data=json_data,
            file_name="backup_records.json",
            mime="application/json"
        )

        up = st.file_uploader(
            "📤 นำเข้าข้อมูล",
            type=["json"],
            key="restore_file"
        )

        if up:
            try:
                restore = json.load(up)
                st.session_state.records = restore
                st.success("นำเข้าข้อมูลสำเร็จ")
            except:
                st.error("ไฟล์ไม่ถูกต้อง")

    # ---------------- TAB 4 ----------------
    with tab4:

        st.subheader("📝 ตรวจสอบ / แก้ไข พฤติการณ์จับกุม")

        f = st.session_state.get("form", {})

        auto_text = f"""ตามวันเวลาที่เกิดเหตุ เจ้าหน้าที่ตำรวจชุดจับกุมขณะปฏิบัติหน้าที่ตั้งจุดตรวจบริเวณ {f.get('arrest_place','')}
พบ {f.get('name','')} ขับขี่{f.get('vehicle_type','')} ยี่ห้อ {f.get('vehicle_brand','')}
รุ่น {f.get('vehicle_model','')} สี {f.get('vehicle_color','')}
หมายเลขทะเบียน {f.get('vehicle_plate','')}

จากการตรวจสอบพบของกลาง {f.get('seized_item','')}

จึงแจ้งข้อกล่าวหาว่า {f.get('charge','')}
และนำตัวส่งพนักงานสอบสวนดำเนินคดีตามกฎหมายต่อไป"""

        behavior = st.text_area(
            "ข้อความพฤติการณ์จับกุม",
            value=f.get("behavior_text", auto_text),
            height=300
        )

        st.session_state.form["behavior_text"] = behavior

        st.write("---")

        s["auto_save"] = st.toggle(
            "💾 บันทึกอัตโนมัติ",
            value=s.get("auto_save", True)
        )

        s["thai_number"] = st.toggle(
            "🔢 ใช้เลขไทย",
            value=s.get("thai_number", False)
        )

        st.write("---")

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
            st.rerun()

    with c2:
        if st.button("⬅️ กลับหน้าหลัก", use_container_width=True):
            st.session_state.page = "dashboard"
            st.rerun()
# =========================================================
# SEARCH PAGE
# =========================================================
if st.session_state.get("page") == "search":

    st.markdown("## 🔎 ค้นหาเอกสาร")

    records = st.session_state.get("records", [])

    c1, c2 = st.columns([3, 1])

    keyword = c1.text_input(
        "ค้นหาชื่อ / เลขบัตร / ทะเบียนรถ"
    )

    mode = c2.selectbox(
        "ประเภท",
        ["ทั้งหมด", "ชื่อ", "เลขบัตร", "ทะเบียน"]
    )

    result = []

    for r in records:

        name = str(r.get("name", ""))
        cid = str(r.get("sub_id", ""))
        plate = str(r.get("vehicle_plate", ""))

        ok = True

        if keyword:
            kw = keyword.lower()

            if mode == "ชื่อ":
                ok = kw in name.lower()

            elif mode == "เลขบัตร":
                ok = kw in cid.lower()

            elif mode == "ทะเบียน":
                ok = kw in plate.lower()

            else:
                ok = (
                    kw in name.lower()
                    or kw in cid.lower()
                    or kw in plate.lower()
                )

        if ok:
            result.append(r)

    st.info(f"พบ {len(result)} รายการ")

    if result:

        for i, r in enumerate(result):

            with st.expander(
                f"📄 {r.get('name','-')} | {r.get('vehicle_plate','-')}"
            ):

                st.write(f"เลขบัตร: {r.get('sub_id','-')}")
                st.write(f"วันที่: {r.get('report_date','-')}")
                st.write(f"ข้อหา: {r.get('charge','-')}")

                c1, c2 = st.columns(2)

                with c1:
                    if st.button("🗑️ ลบ", key=f"del_{i}"):

                        st.session_state.records.remove(r)
                        st.success("ลบแล้ว")
                        st.rerun()

                with c2:
                    if st.button("📋 โหลดข้อมูลนี้", key=f"load_{i}"):

                        st.session_state.form = r
                        st.session_state.page = "form"
                        st.session_state.tab = 3
                        st.rerun()

    else:
        st.warning("ไม่พบข้อมูล")

    st.write("")

    if st.button("⬅️ กลับหน้าหลัก", use_container_width=True):
        st.session_state.page = "dashboard"
        st.rerun()

# =================================================================
# SIDEBAR
# =================================================================
with st.sidebar:
    # ================= LOGO =================
    if logo_base64:
        st.markdown(
            f"""
            <div style="text-align:center; margin-bottom:10px;">
                <img src="data:image/png;base64,{logo_base64}"
                     width="120"
                     style="
                        border-radius:20px;
                        padding:10px;
                        background:rgba(255,255,255,0.08);
                     ">
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown(
        """
        <div style="
            text-align:center;
            font-size:25px;
            font-weight:800;
            margin-bottom:20px;
            color:white;
        ">
             ระบบบันทึกจับกุม
            งานจราจร
        </div>
        """,
        unsafe_allow_html=True
    )

    st.info(
        f"""
👮 ผู้ใช้งาน:
{st.session_state.get('user_full_name','เจ้าหน้าที่')}

🏢 สังกัด:
สภ.ตระการพืชผล
"""
    )

    if st.button("🏠 หน้าแรก", use_container_width=True):
        st.session_state.page = "dashboard"
        st.rerun()

    if st.button("📋 บันทึกจับกุม", use_container_width=True):
        st.session_state.page = "form"
        st.rerun()

    if st.button("🔎 ค้นหาเอกสาร", use_container_width=True):
        st.session_state.page = "search"
        st.rerun()
    # ================= จัดการเอกสาร DROPDOWN =================
    with st.expander("📂 จัดการเอกสาร", expanded=False):

        doc_types = {
            "📋 หนังสือราชการ": "หนังสือราชการ",
            "📝 บันทึกข้อความ": "บันทึกข้อความ",
            "📌 คำสั่ง": "คำสั่ง",
            "📢 ประกาศ": "ประกาศ",
            "📊 รายงาน": "รายงาน",
            "📥 หนังสือรับ": "หนังสือรับ",
            "📤 หนังสือส่ง": "หนังสือส่ง"
        }

        for label, dtype in doc_types.items():

            if st.button(label, use_container_width=True, key=f"doc_{dtype}"):

                st.session_state.doc_type_preselect = dtype
                st.session_state.page = "documents"
                st.session_state.doc_sub_page = "create"

                st.rerun()
                
    if st.button("⚙️ ตั้งค่า", use_container_width=True):
        st.session_state.page = "settings"
        st.rerun()

    if st.button("🚪 ออกจากระบบ", use_container_width=True):
        st.session_state.confirm_logout = True

if st.session_state.get("confirm_logout"):
    st.warning("⚠️ ต้องการออกจากระบบใช่หรือไม่? ข้อมูลที่ยังไม่บันทึกจะหายไป")
    col_y, col_n = st.columns(2)
    if col_y.button("✅ ยืนยัน"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    if col_n.button("❌ ยกเลิก"):
        st.session_state.confirm_logout = False
        st.rerun()
#===================ระบบจัดการเอกสาร=======================

if st.session_state.get("page") == "documents":

    # ── init ──────────────────────────────────────────────
    if "doc_records" not in st.session_state:
        st.session_state.doc_records = []

    if "doc_running_numbers" not in st.session_state:
        st.session_state.doc_running_numbers = {}

    DOC_TYPES = {
        "หนังสือราชการ": {"prefix": "ที่",    "icon": "📋"},
        "บันทึกข้อความ": {"prefix": "บันทึก", "icon": "📝"},
        "คำสั่ง":         {"prefix": "คำสั่ง", "icon": "📌"},
        "ประกาศ":         {"prefix": "ประกาศ", "icon": "📢"},
        "รายงาน":         {"prefix": "รายงาน", "icon": "📊"},
        "หนังสือรับ":     {"prefix": "รับที่",  "icon": "📥"},
        "หนังสือส่ง":     {"prefix": "ส่งที่",  "icon": "📤"},
    }

    OFFICERS = [
        "ร.ต.อ.สมคิด เชื้อเวียง","ร.ต.อ.ชัยพร ชอบงาม","ร.ต.ท.พจน์ ปาสาจันทร์",
        "ร.ต.ต.ประจวบ ศรีบุระ","ส.ต.อ.ธีระวัฒน์ แก่นสาร","ส.ต.อ.ณรงค์ฤทธิ์ เหล่าดี",
        "ส.ต.อ.วิชชากร วงษ์โท","ด.ต.ยุทธพงษ์ ชาดแดง","จ.ส.ต.วัชระ จันสุตะ",
        "ส.ต.ท.วชิรวิชญ์ นันทรักษ์","ส.ต.ท.วิศรุต จันทร์สิงห์","ส.ต.ท.เอกพจน์ อินผล",
        "ส.ต.ต.อดิศร ศุภนิกร"
    ]

    def get_next_doc_number(doc_type, year_th):
        key = f"{doc_type}_{year_th}"
        n = st.session_state.doc_running_numbers.get(key, 0) + 1
        st.session_state.doc_running_numbers[key] = n
        return f"0018.13/{n:04d}/{year_th}"

    def get_preview_number(doc_type, year_th):
        key = f"{doc_type}_{year_th}"
        n = st.session_state.doc_running_numbers.get(key, 0) + 1
        return f"0018.13/{n:04d}/{year_th}"

    # ── header ────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        📂 ระบบจัดการเอกสารราชการ
    </div>
    <div class="header-line"></div>
    """, unsafe_allow_html=True)

    # ── sub navigation ────────────────────────────────────
    sub_page = st.session_state.get("doc_sub_page", "create")

    sn1, sn2, sn3 = st.columns(3)
    if sn1.button("➕ สร้างเอกสาร",    use_container_width=True, key="snav1"):
        st.session_state.doc_sub_page = "create";  st.rerun()
    if sn2.button("🔎 ค้นหาเอกสาร",   use_container_width=True, key="snav2"):
        st.session_state.doc_sub_page = "search";  st.rerun()
    if sn3.button("📬 ทะเบียนรับ-ส่ง", use_container_width=True, key="snav3"):
        st.session_state.doc_sub_page = "register"; st.rerun()

    st.write("---")

    # ═══════════════════════════════════════════════════════
    # SUB: สร้างเอกสาร
    # ═══════════════════════════════════════════════════════
    if sub_page == "create":

        # ================= รับค่าจาก Sidebar =================
        preselect = st.session_state.get("doc_type_preselect", "หนังสือราชการ")

        if preselect not in DOC_TYPES:
            preselect = "หนังสือราชการ"

        dtype = preselect
        meta = DOC_TYPES[dtype]
        year_th = datetime.date.today().year + 543

        st.markdown(f"## {meta['icon']} {dtype}")

        ca, cb = st.columns([2, 1])

        with ca:
            auto_no = st.toggle("ออกเลขอัตโนมัติ", value=True, key="tog_auto_no")
            if auto_no:
                doc_no = get_preview_number(dtype, year_th)
                st.info(f"เลขที่จะออก: **{doc_no}**")
            else:
                doc_no = st.text_input("เลขที่เอกสาร (กรอกเอง)",
                                       placeholder="0018.13/0001/2568",
                                       key="doc_manual_no")

        with cb:
            doc_date    = st.date_input("วันที่เอกสาร", datetime.date.today(), key="doc_date")
            doc_urgency = st.selectbox("ชั้นความเร็ว",
                                       ["ปกติ","ด่วน","ด่วนมาก","ด่วนที่สุด"],
                                       key="doc_urgency")

        st.write("---")

        cc, cd = st.columns(2)
        with cc:
            doc_to      = st.text_input("เรียน / ถึง", key="doc_to",
                                        placeholder="เช่น ผู้กำกับการ สภ.ตระการพืชผล")
            doc_from    = st.text_input("จาก",
                                        value=st.session_state.get("user_full_name","เจ้าหน้าที่"),
                                        key="doc_from")
            doc_ref     = st.text_input("อ้างถึง (ถ้ามี)", key="doc_ref")
        with cd:
            doc_subject = st.text_input("เรื่อง", key="doc_subject")
            doc_attach  = st.text_input("สิ่งที่ส่งมาด้วย", key="doc_attach")

        doc_body = st.text_area("เนื้อหา", height=220, key="doc_body",
                                placeholder="พิมพ์เนื้อหาที่นี่...")

        st.write("---")
        st.caption("✍️ ผู้ลงนาม")
        ce, cf, cg = st.columns(3)
        signer     = ce.selectbox("ผู้ลงนาม", OFFICERS, key="doc_signer")
        signer_pos = cf.text_input("ตำแหน่ง", key="doc_signer_pos",
                                   placeholder="เช่น ผกก.สภ.ตระการพืชผล")
        doc_dept   = cg.text_input("หน่วยงาน", value="สภ.ตระการพืชผล", key="doc_dept")

        st.write("---")

        b1, b2, b3 = st.columns(3)

        with b1:
            if st.button("💾 บันทึกเอกสาร", use_container_width=True, key="btn_save_doc"):
                if not doc_subject:
                    st.warning("⚠️ กรุณากรอกเรื่อง")
                else:
                    final_no = (get_next_doc_number(dtype, year_th)
                                if auto_no else doc_no)
                    new_doc = {
                        "doc_number":  final_no,
                        "doc_type":    dtype,
                        "doc_date":    str(doc_date),
                        "doc_subject": doc_subject,
                        "doc_to":      doc_to,
                        "doc_from":    doc_from,
                        "doc_ref":     doc_ref,
                        "doc_attach":  doc_attach,
                        "doc_body":    doc_body,
                        "doc_urgency": doc_urgency,
                        "signer":      signer,
                        "signer_pos":  signer_pos,
                        "doc_dept":    doc_dept,
                        "created_by":  st.session_state.get("user_full_name","-"),
                        "created_at":  str(datetime.datetime.now()),
                        "status":      "หนังสือรับ" if dtype == "หนังสือรับ" else "หนังสือส่ง"
                    }
                    st.session_state.doc_records.append(new_doc)
                    st.success(f"✅ บันทึกแล้ว เลขที่ {final_no}")

        with b2:
            if st.button("📄 Export Word", use_container_width=True, key="btn_exp_word"):
                st.info("⏳ รอเชื่อม template (เพิ่มในขั้นตอนถัดไป)")

        with b3:
            if st.button("🖨️ Export PDF", use_container_width=True, key="btn_exp_pdf"):
                st.info("⏳ รอเชื่อม template PDF")

    # ═══════════════════════════════════════════════════════
    # SUB: ค้นหาเอกสาร
    # ═══════════════════════════════════════════════════════
    elif sub_page == "search":

        st.markdown("### 🔎 ค้นหาเอกสาร")

        doc_records = st.session_state.get("doc_records", [])

        sc1, sc2, sc3 = st.columns([3, 1, 1])
        kw          = sc1.text_input("ค้นหา เลขที่ / เรื่อง / ถึง", key="srch_kw")
        ftype       = sc2.selectbox("ประเภท", ["ทั้งหมด"] + list(DOC_TYPES.keys()), key="srch_type")
        fyear       = sc3.selectbox("ปี พ.ศ.",
                                    ["ทั้งหมด", str(datetime.date.today().year + 543)],
                                    key="srch_year")

        results = [
            d for d in doc_records
            if (not kw or kw.lower() in d.get("doc_number","").lower()
                      or kw.lower() in d.get("doc_subject","").lower()
                      or kw.lower() in d.get("doc_to","").lower())
            and (ftype == "ทั้งหมด" or d.get("doc_type") == ftype)
            and (fyear == "ทั้งหมด" or fyear in d.get("doc_number",""))
        ]

        st.info(f"พบ {len(results)} รายการ")

        for i, d in enumerate(results):
            m = DOC_TYPES.get(d.get("doc_type",""), {"icon":"📄"})
            with st.expander(f"{m['icon']} {d.get('doc_number','-')} | {d.get('doc_subject','-')}"):
                xa, xb = st.columns(2)
                xa.write(f"**ประเภท:** {d.get('doc_type','-')}")
                xa.write(f"**เรียน:** {d.get('doc_to','-')}")
                xa.write(f"**จาก:** {d.get('doc_from','-')}")
                xb.write(f"**ชั้นความเร็ว:** {d.get('doc_urgency','-')}")
                xb.write(f"**ผู้ลงนาม:** {d.get('signer','-')}")
                xb.write(f"**บันทึกโดย:** {d.get('created_by','-')}")
                if d.get("doc_body"):
                    st.text_area("เนื้อหา", value=d["doc_body"], height=80,
                                 key=f"sbody_{i}", disabled=True)
                if st.button("🗑️ ลบ", key=f"sdel_{i}"):
                    st.session_state.doc_records.remove(d)
                    st.rerun()

    # ═══════════════════════════════════════════════════════
    # SUB: ทะเบียนรับ-ส่ง
    # ═══════════════════════════════════════════════════════
    elif sub_page == "register":

        st.markdown("### 📬 ทะเบียนรับ-ส่งเอกสาร")

        doc_records = st.session_state.get("doc_records", [])
        sent     = [d for d in doc_records if d.get("status") == "หนังสือส่ง"]
        received = [d for d in doc_records if d.get("status") == "หนังสือรับ"]

        rt1, rt2 = st.tabs([
            f"📤 ทะเบียนส่ง ({len(sent)})",
            f"📥 ทะเบียนรับ ({len(received)})"
        ])

        def render_table(docs, label):
            if not docs:
                st.warning(f"ยังไม่มี{label}")
                return
            rows = [{
                "เลขที่":    d.get("doc_number","-"),
                "วันที่":    d.get("doc_date","-"),
                "ประเภท":   d.get("doc_type","-"),
                "เรื่อง":    d.get("doc_subject","-"),
                "ถึง/จาก":  d.get("doc_to","-") or d.get("doc_from","-"),
                "ผู้บันทึก": d.get("created_by","-"),
            } for d in docs]
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)
            csv = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(f"📥 Export CSV ({label})", csv,
                               file_name=f"register_{label}_{datetime.date.today()}.csv",
                               mime="text/csv", key=f"exp_{label}")

        with rt1: render_table(sent, "หนังสือส่ง")
        with rt2: render_table(received, "หนังสือรับ")

    # ── ปุ่มกลับ ──────────────────────────────────────────
    st.write("---")
    if st.button("⬅️ กลับหน้าหลัก", use_container_width=True, key="btn_back_docs"):
        st.session_state.page = "dashboard"
        st.rerun()

# =========================================================
# FORM PAGE ONLY
# =========================================================
if st.session_state.get("page") != "form":
    st.stop()

# =================================================================
# 5. SIDEBAR & HEADER
# =================================================================
OFFICERS = [
    "ร.ต.อ.สมคิด เชื้อเวียง","ร.ต.อ.ชัยพร ชอบงาม","ร.ต.ท.พจน์ ปาสาจันทร์",
    "ร.ต.ต.ประจวบ ศรีบุระ","ส.ต.อ.ธีระวัฒน์ แก่นสาร","ส.ต.อ.ณรงค์ฤทธิ์ เหล่าดี",
    "ส.ต.อ.วิชชากร วงษ์โท","ด.ต.ยุทธพงษ์ ชาดแดง","จ.ส.ต.วัชระ จันสุตะ",
    "ส.ต.ท.วชิรวิชญ์ นันทรักษ์","ส.ต.ท.วิศรุต จันทร์สิงห์","ส.ต.ท.เอกพจน์ อินผล",
    "ส.ต.ต.อดิศร ศุภนิกร"
]

# =================================================================
# 6. NAVBAR ใหม่ (Krungthai Clean Minimal Pro)
# =================================================================

if theme == "Dark":
    app_bg = "#0b1f4e"
    text_color = "#ffffff"
    card_bg = "#111827"
    input_bg = "#111827"

elif theme == "Blue Pro":
    app_bg = "linear-gradient(180deg,#0b1f4e,#123a8f)"
    text_color = "#ffffff"
    card_bg = "#111827"
    input_bg = "#111827"

else:
    app_bg = "linear-gradient(180deg,#f5f9ff,#eef4ff)"
    text_color = "#111827"
    card_bg = "#ffffff"
    input_bg = "#ffffff"

st.markdown(f"""
<style>

/* ================= พื้นหลัง ================= */
.stApp {{
    background:{app_bg} !important;
}}

/* ================= BODY ================= */
.block-container {{
    max-width: 1100px !important;   /* ปรับให้พอดีกับจอ 1366px */
    padding-top: 0.5rem;            /* ลด padding ด้านบน */
    padding-bottom: 1rem;
    margin: auto;                   /* จัดให้อยู่กลาง */
}}
/* ================= ปุ่ม ================= */
.ktb-nav div.stButton > button {{
    width:100%;
    height:44px;
    border:none !important;
    border-radius:12px !important;

    background:#f3f8ff !important;
    color:#0a4fa8 !important;

    font-size:14px !important;
    font-weight:700 !important;

    transition:0.18s;
    box-shadow:none !important;
}}

.ktb-nav div.stButton > button:hover {{
    background:#0a56c2 !important;
    color:white !important;
    transform:translateY(-1px);
}}

/* ================= user box ================= */
.ktb-user {{
    height:44px;
    border-radius:12px;
    background:linear-gradient(135deg,#0a56c2,#0088ff);
    color:white;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:14px;
    font-weight:700;
    padding:0 14px;
    white-space:nowrap;
    overflow:hidden;
}}

/* ================= responsive ================= */
@media (max-width:768px) {{

.ktb-nav div.stButton > button {{
    font-size:12px !important;
    height:40px;
}}

.ktb-user {{
    font-size:12px;
    height:40px;
}}

}}

</style>
""", unsafe_allow_html=True)

# default tab
if "tab" not in st.session_state:
    st.session_state.tab = 0
    
# =========================================================
# NAVBAR
# =========================================================
st.markdown('<div class="ktb-nav">', unsafe_allow_html=True)

c1,c2,c3,c4,c5 = st.columns([1.15,1,1.1,1.1,1.25])

with c1:
    if st.button("🏠 หน้าแรก", use_container_width=True):
        st.session_state.page = "dashboard"
        st.rerun()

with c2:
    if st.button("📍 บันทึก", use_container_width=True):
        st.session_state.tab = 0
        st.rerun()

with c3:
    if st.button("👮 เจ้าหน้าที่", use_container_width=True):
        st.session_state.tab = 1
        st.rerun()

with c4:
    if st.button("👤 ข้อมูลผู้ต้องหา", use_container_width=True):
        st.session_state.tab = 2
        st.rerun()

with c5:
    if st.button("📝 รายละเอียด", use_container_width=True):
        st.session_state.tab = 3
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# CONTENT BOX
# =========================================================
st.markdown('<div class="content-box">', unsafe_allow_html=True)

if st.session_state.tab == 0:
    st.markdown(
        "<h2 style='color:#f9fafb;'>📍 ข้อมูลบันทึก</h2>",
        unsafe_allow_html=True
    )

elif st.session_state.tab == 1:
    st.markdown(
        "<h2 style='color:#f9fafb;'>👮 ผู้บังคับบัญชา</h2>",
        unsafe_allow_html=True
    )

elif st.session_state.tab == 2:
    st.markdown(
        "<h2 style='color:#f9fafb;'>👤 ข้อมูลผู้ต้องหา</h2>",
        unsafe_allow_html=True
    )

elif st.session_state.tab == 3:
    st.markdown(
        "<h2 style='color:#f9fafb;'>📝 รายละเอียดคดี</h2>",
        unsafe_allow_html=True
    )

st.markdown("</div>", unsafe_allow_html=True)

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
        report_time = st.time_input(
            "เวลาลงบันทึก",
            value=st.session_state.form.get(
                'report_time',
                datetime.datetime.now().time()
            ),
            step=60
        )
    # ================= SAVE SESSION =================
    st.session_state.form.update({
        "record_loc": record_loc,
        "incident_date": incident_date,
        "arrest_time": arrest_time,
        "report_date": report_date,
        "report_time": report_time
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

    # ================= SMART CARD =================
    def read_smart_card():
        """
        ฟังก์ชันอ่านข้อมูลจากเครื่องอ่านบัตรประชาชน
        ตอนนี้ยังเป็น placeholder (คืนค่า None ถ้าไม่มีการเชื่อมต่อจริง)
        """
        try:
            # TODO: เพิ่มโค้ดเชื่อมต่อเครื่องอ่านบัตรจริง
            return {
                "name": "นายสมชาย ใจดี",
                "sub_id": "1234567890123",
                "age": 35
            }
        except:
            return None

    # ปุ่มกดดึงข้อมูล
    if st.button("💳 ดึงข้อมูลจากเครื่องอ่านบัตรประชาชน", key="btn_smart_card_t3"):
        card_data = read_smart_card()
        if card_data:
            st.session_state.form.update({
                "name": card_data['name'],
                "sub_id": card_data['sub_id'],
                "age": card_data['age']
            })
            st.success(f"ดึงข้อมูลสำเร็จ: {card_data['name']}")
            st.rerun()
        else:
            st.error("ไม่พบเครื่องอ่านบัตร หรือไม่ได้เสียบบัตร")

    # ================= BASIC INFO =================
    col_fname, col_lname, col_id, col_age = st.columns([2,2,2,1])

    first_name = col_fname.text_input(
        "ชื่อ",
        "",
        key="input_fname_t3"
    )

    last_name = col_lname.text_input(
        "นามสกุล",
        "",
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
    try:
        df_geo = pd.read_csv("thai_districts.csv", encoding="utf-8-sig")
    except:
        df_geo = pd.read_csv("thai_districts.csv", encoding="cp874")

 # 🔥 CLEAN GEO DATA (แก้ “อำเภอ/เขต/ตำบล”)

    df_geo["ProvinceThai"] = df_geo["ProvinceThai"].apply(clean_geo)
    df_geo["DistrictThai"] = df_geo["DistrictThai"].apply(clean_geo)
    df_geo["TambonThai"] = df_geo["TambonThai"].apply(clean_geo)
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
