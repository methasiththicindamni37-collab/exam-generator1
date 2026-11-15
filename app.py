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
import time

# --- Configuration ---
st.set_page_config(
    page_title="AI Exam Generator Pro", 
    page_icon="🎓", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Global Styles (Modern SaaS Look) ---
st.markdown("""
<style>
    /* --- Base Variables --- */
    :root {
        --primary-gradient: linear-gradient(135deg, #4F46E5 0%, #9333EA 100%);
        --secondary-gradient: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        --bg-light: #F8FAFC;
        --card-bg-light: #FFFFFF;
        --text-main: #1E293B;
        --text-muted: #64748B;
        --border-color: #E2E8F0;
        --shadow-sm: 0 1px 3px rgba(0,0,0,0.1);
        --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
        --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
        --radius-md: 12px;
        --radius-lg: 18px;
    }

    [data-theme="dark"] {
        --bg-light: #0B1120;
        --card-bg-light: #151F32;
        --text-main: #F1F5F9;
        --text-muted: #94A3B8;
        --border-color: #334155;
        --shadow-sm: 0 1px 3px rgba(0,0,0,0.3);
        --shadow-md: 0 4px 6px rgba(0,0,0,0.4);
    }

    /* --- Typography & Base --- */
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    
    body {
        font-family: 'Prompt', sans-serif !important;
        background-color: var(--bg-light) !important;
        color: var(--text-main) !important;
        transition: background-color 0.3s ease;
    }
    
    h1, h2, h3, h4, h5, h6 {
        font-weight: 700 !important;
        letter-spacing: -0.025em;
    }

    /* --- Main App Container --- */
    .main {
        padding: 2rem;
    }
    
    .stApp {
        background-color: var(--bg-light) !important;
    }

    /* --- Sidebar Design --- */
    section[data-testid="stSidebar"] {
        background: var(--secondary-gradient) !important;
        border-right: 1px solid rgba(255,255,255,0.1) !important;
        transition: all 0.3s ease;
    }
    
    section[data-testid="stSidebar"] .stDeployButton { display: none; }
    
    section[data-testid="stSidebar"] > div {
        padding-top: 2rem;
    }

    /* Sidebar Headers */
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] h4,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div {
        color: #F8FAFC !important;
    }
    
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: #CBD5E1 !important;
        font-size: 0.9rem;
    }

    /* Sidebar Inputs */
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] [data-baseweb="select"],
    section[data-testid="stSidebar"] details {
        background-color: rgba(255, 255, 255, 0.1) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        color: #FFFFFF !important;
        border-radius: var(--radius-md) !important;
        backdrop-filter: blur(10px);
    }
    
    section[data-testid="stSidebar"] input:focus,
    section[data-testid="stSidebar"] [data-baseweb="select"]:focus-within {
        border-color: #9333EA !important;
        box-shadow: 0 0 0 2px rgba(147, 51, 234, 0.3) !important;
    }

    /* Sidebar Labels */
    section[data-testid="stSidebar"] label {
        margin-bottom: 0.5rem !important;
        font-weight: 500;
    }

    /* Sidebar Toggle/Expander */
    section[data-testid="stSidebar"] details {
        background-color: rgba(0, 0, 0, 0.2) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: var(--radius-md) !important;
        overflow: hidden;
        transition: all 0.3s ease;
    }
    
    section[data-testid="stSidebar"] details[open] {
        background-color: rgba(255, 255, 255, 0.05) !important;
    }
    
    section[data-testid="stSidebar"] details summary {
        padding: 0.8rem 1rem;
        cursor: pointer;
        list-style: none;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    
    section[data-testid="stSidebar"] details summary::after {
        content: "⌄";
        font-size: 1.2rem;
        transition: transform 0.3s;
    }
    section[data-testid="stSidebar"] details[open] summary::after {
        transform: rotate(180deg);
    }
    
    section[data-testid="stSidebar"] details > div {
        background-color: transparent !important;
        padding: 0 1rem 1rem 1rem !important;
        border: none !important;
    }

    /* Sidebar Sliders */
    section[data-testid="stSlider"] > div > div > div > div {
        background-color: rgba(255, 255, 255, 0.2) !important;
    }
    section[data-testid="stSlider"] > div > div > div > div > div {
        background-color: #9333EA !important;
    }
    section[data-testid="stSlider"] > div > div > div > div > div > div {
        background-color: #FFFFFF !important;
        border: 2px solid #9333EA !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSlider"] span {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border: 1px solid #9333EA !important;
    }

    /* --- Main Content Components --- */
    
    /* Upload Card */
    .upload-card {
        background: var(--card-bg-light);
        border: 2px dashed var(--border-color);
        border-radius: var(--radius-lg);
        padding: 3rem;
        text-align: center;
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }
    
    .upload-card:hover {
        border-color: #9333EA;
        box-shadow: var(--shadow-md);
    }
    
    [data-testid="stFileUploader"] section {
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
    }
    
    [data-testid="stFileUploader"] section button {
        background: var(--primary-gradient) !important;
        color: white !important;
        border: none !important;
        padding: 0.8rem 2rem !important;
        border-radius: 50px !important;
        font-weight: 600 !important;
        font-family: 'Prompt', sans-serif !important;
        box-shadow: var(--shadow-md);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    [data-testid="stFileUploader"] section button:hover {
        transform: translateY(-2px);
        box-shadow: var(--shadow-lg);
    }

    /* File List */
    [data-testid="stFileUploaderFile"] {
        background: var(--card-bg-light) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: var(--radius-md) !important;
        margin-top: 1rem !important;
    }

    /* Generate Button */
    .stButton > button {
        background: var(--primary-gradient) !important;
        color: white !important;
        border: none !important;
        border-radius: 50px !important;
        padding: 0.8rem 2.5rem !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        font-family: 'Prompt', sans-serif !important;
        box-shadow: var(--shadow-md);
        transition: all 0.3s ease !important;
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 0.5rem;
    }
    
    .stButton > button:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 20px -5px rgba(79, 70, 229, 0.4) !important;
    }
    
    .stButton > button:active {
        transform: translateY(-1px);
    }

    /* Download Buttons */
    .download-col {
        display: flex;
        gap: 1rem;
        margin-top: 1.5rem;
    }
    
    .stDownloadButton > button {
        background-color: var(--card-bg-light) !important;
        color: var(--text-main) !important;
        border: 2px solid var(--border-color) !important;
        border-radius: var(--radius-md) !important;
        font-weight: 600 !important;
        flex: 1;
        transition: all 0.2s ease;
    }
    
    .stDownloadButton > button:hover {
        background-color: #F1F5F9 !important;
        border-color: #94A3B8 !important;
        transform: translateY(-2px);
    }

    /* Result Card */
    .result-card {
        background: var(--card-bg-light);
        border-radius: var(--radius-lg);
        padding: 2rem;
        box-shadow: var(--shadow-md);
        border: 1px solid var(--border-color);
        margin-top: 2rem;
        position: relative;
        overflow: hidden;
    }
    
    .result-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 4px;
        background: var(--primary-gradient);
    }

    /* Markdown Output */
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        color: #4F46E5 !important;
        border-bottom: 1px solid var(--border-color);
        padding-bottom: 0.5rem;
    }
    
    .stMarkdown ul {
        list-style-type: disc !important;
        padding-left: 1.5rem !important;
    }
    
    .stMarkdown li {
        margin-bottom: 0.5rem !important;
    }

    /* Theme Toggle Switch */
    .theme-switch-wrapper {
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 1000;
        background: var(--card-bg-light);
        border-radius: 50px;
        padding: 5px;
        box-shadow: var(--shadow-md);
        border: 1px solid var(--border-color);
        display: flex;
        align-items: center;
        cursor: pointer;
        transition: all 0.3s ease;
    }
    
    .theme-switch-wrapper:hover {
        transform: scale(1.05);
    }
    
    .theme-icon {
        font-size: 1.2rem;
        padding: 5px 10px;
        transition: color 0.3s;
    }

    /* Mobile Responsive */
    @media (max-width: 768px) {
        .main { padding: 0.5rem !important; }
        .download-col { flex-direction: column; }
        section[data-testid="stSidebar"] { min-width: 260px; }
    }
</style>
""", unsafe_allow_html=True)

# --- Theme Management ---
if "theme" not in st.session_state:
    st.session_state.theme = "light"

def toggle_theme():
    new_theme = "dark" if st.session_state.theme == "light" else "light"
    st.session_state.theme = new_theme
    st.rerun()

st.markdown(f"""
<div class="theme-switch-wrapper" onclick="document.documentElement.setAttribute('data-theme', '{st.session_state.theme}');">
    <span class="theme-icon">{ '🌙' if st.session_state.theme == 'light' else '☀️' }</span>
</div>
<script>
    document.documentElement.setAttribute('data-theme', '{st.session_state.theme}');
</script>
""", unsafe_allow_html=True)


# --- Core Functions ---
MAX_CHARS = 30000
NON_CHAT_KEYWORDS = ("embed", "whisper", "tts", "dall-e", "image", "moderation", "rerank")
THINK_TAGS = ("think", "thinking", "reasoning")

@st.cache_data(ttl=600)
def list_models(api_key, base_url):
    try:
        c = OpenAI(api_key=api_key, base_url=base_url, timeout=20)
        return sorted(m.id for m in c.models.list())
    except Exception:
        return []

def is_chat_model(model_id: str) -> bool:
    return not any(k in model_id.lower() for k in NON_CHAT_KEYWORDS)

def extract_text(f):
    if f is None: return ""
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
    for name in THINK_TAGS:
        open_t, close_t = f"<{name}>", f"</{name}>"
        pattern = re.escape(open_t) + r".*?" + re.escape(close_t)
        text = re.sub(pattern, "", text, flags=re.DOTALL | re.IGNORECASE)
        idx = text.lower().rfind(close_t)
        if idx != -1: text = text[idx + len(close_t):]
        oi = text.lower().find(open_t)
        if oi != -1: text = text[:oi]
    return text.strip()

def make_docx(md: str) -> bytes:
    doc = Document()
    doc.add_heading("ข้อสอบ/แบบฝึกหัด", level=1)
    for line in md.splitlines():
        s = line.strip()
        if not s or set(s) <= set("-|: "): continue
        if s.startswith("# "): doc.add_heading(s[2:].replace("**", ""), level=1)
        elif s.startswith("## "): doc.add_heading(s[3:].replace("**", ""), level=2)
        elif s.startswith("### "): doc.add_heading(s[4:].replace("**", ""), level=3)
        else: doc.add_paragraph(s.replace("**", ""))
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

# Font Registration for PDF
_FONT_CANDIDATES = [("Garuda", "Garuda.ttf", "Garuda-Bold.ttf"), ("Loma", "Loma.ttf", "Loma-Bold.ttf"), ("Norasi", "Norasi.ttf", "Norasi-Bold.ttf"), ("Waree", "Waree.ttf", "Waree-Bold.ttf")]
_FONT_DIRS = ["/usr/share/fonts/truetype/tlwg", "/usr/share/fonts", "."]

def _find_file(fn):
    for d in _FONT_DIRS:
        hits = glob.glob(f"{d}/**/{fn}", recursive=True)
        if hits: return hits[0]
    return None

@st.cache_resource(show_spinner=False)
def register_thai_font():
    for name, reg, bold in _FONT_CANDIDATES:
        reg_path = _find_file(reg)
        if not reg_path: continue
        try:
            pdfmetrics.registerFont(TTFont(name, reg_path))
            bold_name = name + "-Bold"
            bold_path = _find_file(bold)
            if bold_path: pdfmetrics.registerFont(TTFont(bold_name, bold_path))
            addMapping(name, 0, 0, name); addMapping(name, 1, 0, bold_name)
            addMapping(name, 0, 1, name); addMapping(name, 1, 1, bold_name)
            return name, bold_name
        except Exception: continue
    return None

def make_pdf(md: str):
    fonts = register_thai_font()
    if not fonts: return None
    base, bold = fonts
    style = ParagraphStyle("body", fontName=base, fontSize=12, leading=20)
    story = [Paragraph("ข้อสอบ/แบบฝึกหัด", ParagraphStyle("h1", parent=style, fontName=bold, fontSize=18)), Spacer(1, 10)]
    for line in md.splitlines():
        s = line.strip()
        if not s: story.append(Spacer(1, 5))
        elif s.startswith("# "): story.append(Paragraph(s[2:], ParagraphStyle("h1", parent=style, fontName=bold, fontSize=16)))
        elif s.startswith("## "): story.append(Paragraph(s[3:], ParagraphStyle("h2", parent=style, fontName=bold, fontSize=14)))
        elif s.startswith("### "): story.append(Paragraph(s[4:], ParagraphStyle("h3", parent=style, fontName=bold, fontSize=12)))
        else: story.append(Paragraph(s.replace("**", ""), style))
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm).build(story)
    return buf.getvalue()

# --- App Logic ---
st.sidebar.header("⚙️ ตั้งค่าระบบ")

API_KEY = st.secrets.get("OPENAI_API_KEY", "")
BASE_URL = st.secrets.get("OPENAI_BASE_URL", "https://ai.psu.blue/v1")

if not API_KEY:
    st.sidebar.error("❌ ยังไม่ได้ตั้ง OPENAI_API_KEY ใน Secrets")
else:
    all_models = list_models(API_KEY, BASE_URL)
    chat_models = [m for m in all_models if is_chat_model(m)] or all_models
    
    selected_model = st.sidebar.selectbox("เลือกโมเดล AI", chat_models, index=0) if chat_models else ""

with st.sidebar.expander("🔧 ตั้งค่าขั้นสูง"):
    q_type = st.selectbox("รูปแบบข้อสอบ", ["Multiple Choice (ปรนัย)", "Short Answer (อัตนัย)", "True/False (ถูกผิด)"])
    difficulty = st.selectbox("ระดับความยาก", ["ง่าย", "ปานกลาง", "ยาก"])
    num_q = st.slider("จำนวนข้อสอบ", 1, 20, 5)
    temp = st.slider("Temperature", 0.0, 1.5, 0.7, 0.1)
    
    col_auto, col_man = st.columns([1, 1])
    auto_tok = col_auto.checkbox("คำนวณอัตโนมัติ", value=True)
    max_tok = col_man.number_input("Max Tokens", 500, 32000, 4000) if not auto_tok else 4000
    
    stream_mode = st.checkbox("แสดงผลแบบ Live (Streaming)", value=True)

uploaded_file = st.file_uploader("📂 ลากไฟล์มาวางที่นี่ (PDF / Word / TXT)", type=["pdf", "docx", "txt"])

# --- Main Content ---
if uploaded_file:
    raw_text = extract_text(uploaded_file)
    if raw_text.startswith("Error"):
        st.error(raw_text)
    elif not raw_text.strip():
        st.warning("ไม่พบข้อความในไฟล์")
    else:
        chars_used = len(raw_text)
        st.info(f"📄 อ่านเนื้อหาสำเร็จ: {chars_used:,} ตัวอักษร")
        
        if chars_used > MAX_CHARS:
            raw_text = raw_text[:MAX_CHARS]
            st.info(f"ตัดเนื้อหาเหลือเพียง {MAX_CHARS:,} ตัวอักษรแรก")

        if st.button("🚀 เริ่มสร้างข้อสอบด้วย AI", type="primary"):
            prompt = f"""
คุณคือครูผู้เชี่ยวชาญที่ใจดี
จากเนื้อหาเอกสารต่อไปนี้:
---
{raw_text}
---
กรุณาสร้างข้อสอบ {num_q} ข้อ ในรูปแบบ {q_type} ระดับความยาก {difficulty}
ข้อกำหนด:
1. ข้อสอบต้องอิงจากเนื้อหาในเอกสารเท่านั้น
2. หากเป็นข้อสอบปรนัย ให้เขียนตัวเลือกเรียงลงมาด้านล่าง ดังนี้:
   ก. [ข้อความคำตอบ]
   ข. [ข้อความคำตอบ]
   ค. [ข้อความคำตอบ]
   ง. [ข้อความคำตอบ]
3. ด้านล่างสุดของแต่ละข้อต้องมี "เฉลย" แยกออกมาชัดเจน เช่น **เฉลย: ก**
4. ใช้ภาษาไทยทางการ
5. จัดรูปแบบ Markdown ให้สวยงาม
6. ตอบเฉพาะข้อสอบและเฉลย ไม่ต้องอธิบายขั้นตอนการคิด
"""
            
            with st.spinner("🧠 AI กำลังประมวลผลและแต่งคำถาม..."):
                client = OpenAI(api_key=API_KEY, base_url=BASE_URL, timeout=180, max_retries=1)
                
                tokens = min(16000, max(1500, num_q * 450 + 800)) if auto_tok else int(max_tok)
                system_prompt = "คุณคือครูผู้เชี่ยวชาญด้านการออกข้อสอบ"
                messages = [{"role": "system", "content": system_prompt}, {"role": "user", "content": prompt}]
                
                result = ""
                try:
                    params = {"temperature": temp, "max_tokens": tokens}
                    resp = client.chat.completions.create(model=selected_model, messages=messages, stream=stream_mode, **params)
                    
                    placeholder = st.empty()
                    for chunk in resp:
                        if chunk.choices and chunk.choices[0].delta.content:
                            result += chunk.choices[0].delta.content
                            placeholder.markdown(clean_output(result) + " ▌")
                    placeholder.markdown(clean_output(result))
                    
                    st.session_state.exam_result = clean_output(result)
                    st.toast("✅ สร้างข้อสอบสำเร็จ!", icon="🎉")
                    
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {str(e)}")
                    st.warning("ลองปิดโหมด Streaming หรือเปลี่ยนโมเดลดูนะครับ")
else:
    st.markdown("""
    <div style="text-align: center; padding: 4rem; background: white; border-radius: 18px; border: 2px dashed #E2E8F0;">
        <h1 style="color: #4F46E5; font-size: 2.5rem; margin-bottom: 0.5rem;">ระบบสร้างข้อสอบอัจฉริยะ</h1>
        <p style="color: #64748B; font-size: 1.2rem;">อัปโหลดเอกสารแล้วกดปุ่มเพื่อเริ่มสร้างข้อสอบทันที<br>รองรับ PDF, Word, และ Text files</p>
    </div>
    """, unsafe_allow_html=True)

# --- Render Results ---
if "exam_result" in st.session_state and st.session_state.exam_result:
    res = st.session_state.exam_result
    if not res.startswith(("⚠️", "⛔")):
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            pdf_bytes = make_pdf(res)
            st.download_button("⬇️ ดาวน์โหลด PDF", pdf_bytes, "exam.pdf", mime="application/pdf", use_container_width=True)
        with col2:
            st.download_button("⬇️ ดาวน์โหลด DOCX", make_docx(res), "exam.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            
        st.divider()
        st.markdown(f'<div class="result-card"><h2>📝 ผลข้อสอบและแบบฝึกหัด</h2>{res}</div>', unsafe_allow_html=True)
    else:
        st.warning(res)
