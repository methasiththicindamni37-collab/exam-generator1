import io
import streamlit as st
import PyPDF2
from docx import Document
from openai import OpenAI

st.set_page_config(page_title="AI Exam Generator", page_icon="🤖", layout="wide")

# ---------- UI ----------
st.markdown("""
<style>
    .stApp { background-color: #ffffff; }
    h1, h2, h3, h4, h5, h6, p, label, li, th, td { color: #000000 !important; }
    section[data-testid="stSidebar"] { background-color: #0f172a !important; }
    section[data-testid="stSidebar"] * { color: #ffffff !important; }
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] [data-baseweb="select"] * { color: #000000 !important; }
    .stButton>button {
        background-color: #0f172a !important; color: #ffffff !important;
        font-weight: bold; font-size: 1.1rem; padding: 0.6rem 1.5rem; border: none;
    }
    .stButton>button:hover { background-color: #1e293b !important; }
</style>
""", unsafe_allow_html=True)

st.title("🤖 AI Exam Generator")
st.markdown("### 📌 ระบบสร้างข้อสอบอัจฉริยะ")

# ---------- Secrets ----------
API_KEY = st.secrets.get("OPENAI_API_KEY", "")
BASE_URL = st.secrets.get("OPENAI_BASE_URL", "https://ai.psu.blue/v1")
GOOGLE_KEY = st.secrets.get("GOOGLE_API_KEY", "")


@st.cache_data(ttl=600, show_spinner=False)
def list_models(api_key, base_url):
    try:
        c = OpenAI(api_key=api_key, base_url=base_url)
        return sorted(m.id for m in c.models.list())
    except Exception:
        return []


# ---------- Sidebar ----------
st.sidebar.header("⚙️ การตั้งค่า")
provider = st.sidebar.radio("เลือกผู้ให้บริการ", ["PSU dotBlue (OpenAI-compatible)", "Google Gemini"])

model_name = ""
if provider.startswith("PSU"):
    if not API_KEY:
        st.sidebar.error("ยังไม่ได้ตั้ง OPENAI_API_KEY ใน Secrets")
    models = list_models(API_KEY, BASE_URL) if API_KEY else []
    if models:
        model_name = st.sidebar.selectbox("โมเดล", models)
    else:
        st.sidebar.warning("ดึงรายชื่อโมเดลไม่ได้ กรุณาพิมพ์ชื่อโมเดลเอง")
        model_name = st.sidebar.text_input("ชื่อโมเดล (ตามหน้า ai.psu.blue)")
else:
    if not GOOGLE_KEY:
        st.sidebar.error("ยังไม่ได้ตั้ง GOOGLE_API_KEY ใน Secrets")
    model_name = st.sidebar.text_input("ชื่อโมเดล Gemini", value="gemini-2.5-flash")

question_type = st.sidebar.selectbox(
    "รูปแบบข้อสอบ",
    ["Multiple Choice (ปรนัย)", "Short Answer (อัตนัย)", "True/False (ถูกผิด)"],
)
difficulty = st.sidebar.selectbox("ระดับความยาก", ["ง่าย", "ปานกลาง", "ยาก"])
num_questions = st.sidebar.slider("จำนวนข้อสอบ", 1, 20, 5)

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์เอกสาร (PDF / Word / TXT)", type=["pdf", "docx", "txt"])

MAX_CHARS = 30000


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
            return "\n".join(p.text for p in Document(io.BytesIO(data)).paragraphs)
        return data.decode("utf-8", errors="ignore")
    except Exception as e:
        return f"Error: {e}"


def generate_exam(prompt_text):
    if not model_name:
        return "⚠️ กรุณาระบุชื่อโมเดล"

    if provider.startswith("PSU"):
        if not API_KEY:
            return "⚠️ ยังไม่ได้ตั้ง OPENAI_API_KEY ใน Secrets"
        try:
            client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
            resp = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": "คุณคือครูผู้เชี่ยวชาญ"},
                    {"role": "user", "content": prompt_text},
                ],
                temperature=0.7,
                max_tokens=2000,
            )
            
            # --- จุดสำคัญ: แก้ Error 'str' has no attribute 'choices' ---
            # ถ้า API ส่งข้อความกลับมาเป็น String ธรรมดา (เช่น Error) ให้แสดงเลย
            if isinstance(resp, str):
                return resp
            
            # ถ้าเป็น Object ให้ดึงข้อความ
            return resp.choices[0].message.content
            
        except Exception as e:
            return f"เกิดข้อผิดพลาด dotBlue: {e}"

    if not GOOGLE_KEY:
        return "⚠️ ยังไม่ได้ตั้ง GOOGLE_API_KEY ใน Secrets"
    try:
        import google.generativeai as genai
        genai.configure(api_key=GOOGLE_KEY)
        model = genai.GenerativeModel(model_name)
        resp = model.generate_content(prompt_text)
        return resp.text
    except Exception as e:
        return f"เกิดข้อผิดพลาด Google: {e}"


# ---------- Main ----------
if uploaded_file is not None:
    text_content = extract_text(uploaded_file)
    if len(text_content) > MAX_CHARS:
        st.info(f"เอกสารยาวมาก ใช้เฉพาะ {MAX_CHARS:,} ตัวอักษรแรก")
        text_content = text_content[:MAX_CHARS]

    if st.button("🚀 เริ่มสร้างข้อสอบ"):
        st.divider()
        prompt = f"""
คุณคือครูผู้เชี่ยวชาญที่ใจดี
จากเนื้อหาเอกสารต่อไปนี้:
---
{text_content}
---
กรุณาสร้างข้อสอบ {num_questions} ข้อ ในรูปแบบ {question_type} ระดับความยาก {difficulty}
ข้อกำหนด:
1. ข้อสอบต้องอิงจากเนื้อหาในเอกสารเท่านั้น
2. หากเป็นข้อสอบปรนัย ต้องมีตัวเลือก ก. ข. ค. ง.
3. ด้านล่างสุดต้องมี "เฉลย" แยกออกมาชัดเจน
4. ใช้ภาษาไทยทางการ
5. จัดรูปแบบ Markdown ให้สวยงาม
"""
        with st.spinner("🧠 กำลังประมวลผลข้อมูล..."):
            result = generate_exam(prompt)
        st.markdown("## 📄 ผลข้อสอบและแบบฝึกหัด:")
        st.markdown(result)
