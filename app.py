import io
import re
import glob
import html
import streamlit as st
import PyPDF2
from docx import Document
from openai import OpenAI
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping

# ตั้งค่าหน้าเว็บให้เป็นโหมดสว่าง
st.set_page_config(
    page_title="ระบบสร้างข้อสอบ/แบบฝึกหัด", 
    page_icon="📝", 
    layout="wide"
)

# ---------- CSS ปรับสีให้ชัดเจน อ่านง่าย ----------
st.markdown("""
<style>
    /* พื้นหลังหลักสีขาว ตัวหนังสือสีดำ */
    .stApp { background-color: #ffffff; color: #0f172a; }
    h1, h2, h3, h4, h5, h6, p, label, li, th, td, span, a, div { color: #0f172a !important; }

    /* ---------- Sidebar สีกรมท่าเข้ม ---------- */
    section[data-testid="stSidebar"] { background-color: #0f172a !important; }
    
    /* ข้อความใน Sidebar ให้เป็นขาว */
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div { color: #ffffff !important; }

    /* กล่องเลือก / ช่องกรอก ใน Sidebar: พื้นขาว ตัวอักษรเข้ม */
    section[data-testid="stSidebar"] [data-baseweb="select"],
    section[data-testid="stSidebar"] input[type="text"],
    section[data-testid="stSidebar"] input[type="number"],
    section[data-testid="stSidebar"] input[type="range"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
    }
    
    /* ไอคอนและลูกศรในช่องเลือก */
    section[data-testid="stSidebar"] [data-baseweb="select"] svg { fill: #0f172a !important; }
    
    /* รายการตัวเลือกที่เด้งลงมา (Dropdown) */
    [data-baseweb="popover"] ul,
    [data-baseweb="popover"] li { background-color: #ffffff !important; }
    [data-baseweb="popover"] li,
    [data-baseweb="popover"] li * { color: #0f172a !important; }
    [data-baseweb="popover"] li:hover { background-color: #e2e8f0 !important; }

    /* ---------- ปรับสี Slider ให้เห็นชัด ---------- */
    /* เส้นฐานของ Slider */
    section[data-testid="stSidebar"] [data-testid="stSlider"] > div > div > div > div {
        background-color: #334155 !important;
    }
    /* เส้นที่ลากไปแล้ว (Fill) */
    section[data-testid="stSidebar"] [data-testid="stSlider"] > div > div > div > div > div {
        background-color: #00ffff !important; /* สีฟ้าสดใส */
    }
    /* จุดจับ (Handle) */
    section[data-testid="stSidebar"] [data-testid="stSlider"] > div > div > div > div > div > div {
        background-color: #ffffff !important; /* สีขาว */
        border: 2px solid #00ffff !important; /* ขอบสีฟ้า */
    }
    /* เลขจำนวน */
    section[data-testid="stSidebar"] [data-testid="stSlider"] > div > div > div > div > div > div > span {
        color: #ffffff !important;
        font-weight: bold;
    }

    /* ---------- ช่องอัปโหลดไฟล์ ---------- */
    [data-testid="stFileUploader"] section {
        background-color: #f8fafc !important;
        border: 2px dashed #64748b !important;
    }
    [data-testid="stFileUploader"] section *,
    [data-testid="stFileUploaderFile"] * { color: #0f172a !important; }
    [data-testid="stFileUploader"] section button {
        background-color: #ffffff !important;
        border: 1px solid #0f172a !important;
    }
    [data-testid="stFileUploader"] section button * { color: #0f172a !important; }
    [data-testid="stFileUploader"] svg { fill: #0f172a !important; }

    /* ---------- ปุ่ม: พื้นกรมท่า ตัวอักษรขาว ---------- */
    .stButton > button, .stDownloadButton > button {
        background-color: #0f172a !important;
        border: none !important;
        padding: 0.6rem 1.5rem;
        border-radius: 8px;
        font-weight: bold;
    }
    /* บังคับให้ตัวหนังสือและไอคอนปุ่มเป็นสีขาว */
    .stButton > button *, .stDownloadButton > button *,
    .stButton > button p, .stDownloadButton > button p,
    .stButton > button span, .stDownloadButton > button span {
        color: #ffffff !important;
        font-size: 1.05rem;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #1d4ed8 !important;
    }

    /* ---------- กล่องแจ้งเตือน ---------- */
    [data-testid="stAlert"] * { color: #0f172a !important; }
</style>
""", unsafe_allow_html=True)

st.title("📝 ระบบสร้างข้อสอบ/แบบฝึกหัด")

# ---------- Secrets ----------
API_KEY = st.secrets.get("OPENAI_API_KEY", "")
BASE_URL = st.secrets.get("OPENAI_BASE_URL", "https://ai.psu.blue/v1")

MAX_CHARS = 30000

NON_CHAT_KEYWORDS = (
    "embed", "whisper", "tts", "dall-e", "dalle", "image", "moderation",
    "rerank", "transcribe", "audio", "speech", "stable-diffusion", "sdxl", "flux",
)

THINK_TAGS = ("think", "thinking", "reasoning")


@st.cache_data(ttl=600, show_spinner=False)
def list_models(api_key, base_url):
    try:
        c = OpenAI(api_key=api_key, base_url=base_url, timeout=20)
        return sorted(m.id for m in c.models.list())
    except Exception:
        return []


def is_chat_model(model_id: str) -> bool:
    mid = model_id.lower()
    return not any(k in mid for k in NON_CHAT_KEYWORDS)


# ---------- Sidebar ----------
st.sidebar.header("⚙️ การตั้งค่า")

if not API_KEY:
    st.sidebar.error("❌ ยังไม่ได้ตั้ง OPENAI_API_KEY ใน Secrets")

model_name = ""
if API_KEY:
    all_models = list_models(API_KEY, BASE_URL)
    chat_models = [m for m in all_models if is_chat_model(m)] or all_models
    
    if chat_models:
        model_name = st.sidebar.selectbox("เลือกโมเดล", chat_models)
    else:
        st.sidebar.warning("ไม่พบรายการโมเดล")
        model_name = st.sidebar.text_input("พิมพ์ชื่อโมเดลเอง")
else:
    model_name = st.sidebar.text_input("พิมพ์ชื่อโมเดลเอง")

question_type = st.sidebar.selectbox(
    "รูปแบบข้อสอบ",
    ["Multiple Choice (ปรนัย)", "Short Answer (อัตนัย)", "True/False (ถูกผิด)"],
)
difficulty = st.sidebar.selectbox("ระดับความยาก", ["ง่าย", "ปานกลาง", "ยาก"])
num_questions = st.sidebar.slider("จำนวนข้อสอบ", 1, 20, 5)

with st.sidebar.expander("🔧 ตั้งค่าขั้นสูง"):
    temperature = st.slider("Temperature", 0.0, 1.5, 0.7, 0.1)
    auto_tokens = st.checkbox("คำนวณ max tokens อัตโนมัติ", value=True)
    manual_tokens = st.number_input("Max tokens (กรณีกำหนดเอง)", 500, 32000, 4000, 500)
    use_stream = st.checkbox("แสดงผลแบบ Streaming", value=True)

uploaded_file = st.file_uploader(
    "📂 อัปโหลดไฟล์เอกสาร (PDF / Word / TXT)", type=["pdf", "docx", "txt"]
)


# ---------- Functions ----------
def extract_text(f):
    if f is None:
        return ""
    data = f.getvalue()
    name = f.name.lower()
    try:
        if name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(io.BytesIO(data))
            return "\n".join((p.extract_text() or "") for p in reader.pages)
        if name.endswith(".docx"):
            doc = Document(io.BytesIO(data))
            parts = [p.text for p in doc.paragraphs]
            for table in doc.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(parts)
        return data.decode("utf-8", errors="ignore")
    except Exception as e:
        return f"Error: {e}"


def clean_output(text: str) -> str:
    """ตัดส่วนความคิด (reasoning) ของโมเดลบางตัวออก"""
    for name in THINK_TAGS:
        open_t = "<" + name + ">"
        close_t = "<" + "/" + name + ">"

        pattern = re.escape(open_t) + r".*?" + re.escape(close_t)
        text = re.sub(pattern, "", text, flags=re.DOTALL | re.IGNORECASE)

        idx = text.lower().rfind(close_t)
        if idx != -1:
            text = text[idx + len(close_t):]

        oi = text.lower().find(open_t)
        if oi != -1:
            text = text[:oi]
    return text.strip()


# ---------- ไฟล์ดาวน์โหลด ----------
def make_docx(markdown_text: str) -> bytes:
    doc = Document()
    doc.add_heading("ข้อสอบ/แบบฝึกหัด", level=1)
    for line in markdown_text.splitlines():
        stripped = line.strip()
        if not stripped or set(stripped) <= set("-|: "):
            continue
        if stripped.startswith("### "):
            doc.add_heading(stripped[4:].replace("**", ""), level=3)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:].replace("**", ""), level=2)
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:].replace("**", ""), level=1)
        else:
            doc.add_paragraph(stripped.replace("**", ""))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


_FONT_CANDIDATES = [
    ("Garuda", "Garuda.ttf", "Garuda-Bold.ttf"),
    ("Loma", "Loma.ttf", "Loma-Bold.ttf"),
    ("Norasi", "Norasi.ttf", "Norasi-Bold.ttf"),
    ("Waree", "Waree.ttf", "Waree-Bold.ttf"),
    ("Kinnari", "Kinnari.ttf", "Kinnari-Bold.ttf"),
]
_FONT_DIRS = [
    "/usr/share/fonts/truetype/tlwg",
    "/usr/share/fonts/truetype",
    "/usr/share/fonts",
    ".",
    "fonts",
]


def _find_file(filename):
    for d in _FONT_DIRS:
        hits = glob.glob(f"{d}/**/{filename}", recursive=True)
        if hits:
            return hits[0]
    return None


@st.cache_resource(show_spinner=False)
def register_thai_font():
    for name, reg, bold in _FONT_CANDIDATES:
        reg_path = _find_file(reg)
        if not reg_path:
            continue
        try:
            pdfmetrics.registerFont(TTFont(name, reg_path))
            bold_name = name
            bold_path = _find_file(bold)
            if bold_path:
                bold_name = name + "-Bold"
                pdfmetrics.registerFont(TTFont(bold_name, bold_path))
            addMapping(name, 0, 0, name)
            addMapping(name, 1, 0, bold_name)
            addMapping(name, 0, 1, name)
            addMapping(name, 1, 1, bold_name)
            return name, bold_name
        except Exception:
            continue
    return None


def make_pdf(markdown_text: str):
    fonts = register_thai_font()
    if not fonts:
        return None
    base, bold = fonts

    body = ParagraphStyle("body", fontName=base, fontSize=12, leading=20, spaceAfter=4)
    h1 = ParagraphStyle("h1", parent=body, fontName=bold, fontSize=18, leading=26, spaceAfter=8)
    h2 = ParagraphStyle("h2", parent=body, fontName=bold, fontSize=15, leading=22, spaceBefore=6, spaceAfter=6)
    h3 = ParagraphStyle("h3", parent=body, fontName=bold, fontSize=13, leading=20, spaceBefore=4, spaceAfter=4)

    def to_markup(s):
        s = html.escape(s, quote=False)
        s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
        s = s.replace("**", "")
        return s

    story = [Paragraph("ข้อสอบ/แบบฝึกหัด", h1), Spacer(1, 6)]
    for line in markdown_text.splitlines():
        s = line.strip()
        if not s:
            story.append(Spacer(1, 4))
            continue
        if set(s) <= set("-|: "):
            if s.startswith("---"):
                story.append(Spacer(1, 6))
            continue

        style = body
        if s.startswith("### "):
            style, s = h3, s[4:]
        elif s.startswith("## "):
            style, s = h2, s[3:]
        elif s.startswith("# "):
            style, s = h1, s[2:]
        elif s.startswith(("- ", "* ")):
            s = "• " + s[2:]
        elif s.startswith("|"):
            s = "   ".join(c.strip() for c in s.strip("|").split("|"))

        try:
            story.append(Paragraph(to_markup(s), style))
        except Exception:
            story.append(Paragraph(html.escape(s.replace("**", ""), quote=False), body))

    buf = io.BytesIO()
    SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
        title="ข้อสอบ/แบบฝึกหัด",
    ).build(story)
    return buf.getvalue()


# ---------- เรียกโมเดล ----------
def _err_text(e) -> str:
    return str(e).lower()


def _call_api(client, model, messages, params, stream, placeholder):
    if stream:
        resp = client.chat.completions.create(
            model=model, messages=messages, stream=True, **params
        )
        full = ""
        for chunk in resp:
            if not getattr(chunk, "choices", None):
                continue
            delta = chunk.choices[0].delta
            piece = getattr(delta, "content", None)
            if piece:
                full += piece
                if placeholder is not None:
                    placeholder.markdown(clean_output(full) + " ▌")
        return full
    resp = client.chat.completions.create(model=model, messages=messages, **params)
    return resp.choices[0].message.content or ""


def generate_exam(prompt_text, placeholder=None):
    if not model_name:
        return "⚠️ กรุณาระบุชื่อโมเดล"
    if not API_KEY:
        return "⚠️ กรุณาตั้งค่า OPENAI_API_KEY ใน Secrets"

    tokens = (
        min(16000, max(1500, num_questions * 450 + 800))
        if auto_tokens else int(manual_tokens)
    )
    system_prompt = "คุณคือครูผู้เชี่ยวชาญด้านการออกข้อสอบ"

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL, timeout=180, max_retries=1)

    system_modes = ["system", "merged"]
    param_variants = [
        {"temperature": temperature, "max_tokens": tokens},
        {"temperature": temperature, "max_completion_tokens": tokens},
        {"max_completion_tokens": tokens},
        {"max_tokens": tokens},
        {},
    ]
    stream_modes = [True, False] if use_stream else [False]

    last_error = None
    for sys_mode in system_modes:
        if sys_mode == "system":
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt_text},
            ]
        else:
            messages = [{"role": "user", "content": f"{system_prompt}\n\n{prompt_text}"}]

        for stream in stream_modes:
            for params in param_variants:
                try:
                    out = _call_api(client, model_name, messages, params, stream, placeholder)
                    out = clean_output(out)
                    if out:
                        return out
                    last_error = "โมเดลไม่ได้ส่งข้อความกลับมา"
                except Exception as e:
                    last_error = e
                    msg = _err_text(e)
                    if any(k in msg for k in (
                        "401", "invalid api key", "incorrect api key", "authentication"
                    )):
                        return "⛔ API Key ไม่ถูกต้องหรือไม่มีสิทธิ์ใช้งาน"
                    if any(k in msg for k in (
                        "429", "rate limit", "quota", "insufficient"
                    )):
                        return "⛔ ใช้งานเกินกำหนด (Credit/Rate Limit) กรุณาลองใหม่ภายหลัง หรือหลัง 00:00 น."
                    if any(k in msg for k in (
                        "404", "model_not_found", "does not exist"
                    )):
                        return f"⚠️ ไม่พบโมเดล `{model_name}` หรือไม่มีสิทธิ์ใช้งาน"
                    continue

    return f"⚠️ Error: {last_error}"


# ---------- Main ----------
if "exam_result" not in st.session_state:
    st.session_state.exam_result = ""

if uploaded_file is not None:
    text_content = extract_text(uploaded_file)

    if text_content.startswith("Error:"):
        st.error(f"อ่านไฟล์ไม่สำเร็จ: {text_content}")
        st.stop()

    if not text_content.strip():
        st.warning("ไม่พบข้อความในไฟล์ (หาก PDF เป็นภาพสแกน ต้องทำ OCR ก่อน)")
        st.stop()

    if len(text_content) > MAX_CHARS:
        st.info(f"เอกสารยาวมาก ใช้เฉพาะ {MAX_CHARS:,} ตัวอักษรแรก")
        text_content = text_content[:MAX_CHARS]

    st.caption(f"อ่านเนื้อหาได้ {len(text_content):,} ตัวอักษร")

    if st.button("🚀 เริ่มสร้างข้อสอบ"):
        prompt = f"""
คุณคือครูผู้เชี่ยวชาญที่ใจดี
จากเนื้อหาเอกสารต่อไปนี้:
---
{text_content}
---
กรุณาสร้างข้อสอบ {num_questions} ข้อ ในรูปแบบ {question_type} ระดับความยาก {difficulty}
ข้อกำหนด:
1. ข้อสอบต้องอิงจากเนื้อหาในเอกสารเท่านั้น
2. หากเป็นข้อสอบปรนัย ให้เขียนตัวเลือกเรียงลงมาด้านล่าง ดังนี้:
   ก. [ข้อความคำตอบ]
   ข. [ข้อความคำตอบ]
   ค. [ข้อความคำตอบ]
   ง. [ข้อความคำตอบ]
3. ด้านล่างสุดต้องมี "เฉลย" แยกออกมาชัดเจน
4. ใช้ภาษาไทยทางการ
5. จัดรูปแบบ Markdown ให้สวยงาม
6. ตอบเฉพาะข้อสอบและเฉลย ไม่ต้องอธิบายขั้นตอนการคิด
"""
        st.divider()
        st.markdown("## 📄 ผลข้อสอบและแบบฝึกหัด:")
        placeholder = st.empty()
        with st.spinner("🧠 กำลังประมวลผลข้อมูล..."):
            result = generate_exam(prompt, placeholder)
        placeholder.markdown(result)
        st.session_state.exam_result = result

    result = st.session_state.exam_result
    if result and not result.startswith(("⚠️", "⛔")):
        pdf_bytes = make_pdf(result)
        d1, d2 = st.columns(2)
        with d1:
            if pdf_bytes:
                st.download_button(
                    "⬇️ ดาวน์โหลด PDF", pdf_bytes,
                    file_name="exam.pdf", mime="application/pdf",
                )
            else:
                st.warning("สร้าง PDF ไม่ได้: ไม่พบฟอนต์ไทย (เพิ่มไฟล์ packages.txt ที่มี fonts-thai-tlwg แล้ว Reboot)")
        with d2:
            st.download_button(
                "⬇️ ดาวน์โหลด DOCX", make_docx(result),
                file_name="exam.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    elif result:
        st.warning(result)
else:
    st.info("กรุณาอัปโหลดไฟล์เอกสารเพื่อเริ่มต้น")
