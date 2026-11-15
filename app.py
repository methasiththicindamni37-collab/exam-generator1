import streamlit as st
import google.generativeai as genai
import openai
import PyPDF2
from docx import Document
import io

# --- ⚠️ ส่วนสำคัญ: ใส่คีย์ API ตรงนี้ ---
# ผมใส่คีย์ที่คุณให้มาไว้ให้แล้ว เปิดมาใช้งานได้เลยครับ
GOOGLE_API_KEY_DEFAULT = ""
OPENAI_API_KEY_DEFAULT = "sk-user-7bf48f1c3e009ffb082e675c89e1e3121c464c6e884f8a262e8e87ee788c9987"
# ---------------------------------------

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="AI Exam Generator", page_icon="🤖", layout="wide")

# --- 🎨 ปรับแต่ง UI (โทนสีขาวดำ ตัดกันชัด อ่านง่าย) ---
st.markdown("""
<style>
    :root {
        --main-bg: #ffffff;
        --sidebar-bg: #0f172a;
        --text-main: #000000;
        --text-sub: #334155;
        --input-bg: #ffffff;
        --input-border: #000000;
        --btn-bg: #0f172a;
        --btn-text: #ffffff;
    }
    body { background-color: var(--main-bg); color: var(--text-main); }
    .stApp { background-color: var(--main-bg); }
    h1, h2, h3, h4, h5, h6, p, label, div, span, li, th, td { color: var(--text-main) !important; font-weight: bold; }
    
    .stSidebar { background-color: var(--sidebar-bg) !important; border-right: 2px solid #334155; }
    .stSidebar .stSidebarContent { background-color: var(--sidebar-bg); }
    .stSidebar p, .stSidebar div, .stSidebar label { color: #ffffff !important; }
    [data-testid="stAlert"] { background-color: var(--sidebar-bg) !important; color: var(--btn-text) !important; border: 1px solid #475569 !important; }
    
    .stTextInput>div>div>input, .stSelectbox>div>div, .stNumberInput>div>div>input {
        background-color: var(--input-bg) !important; color: var(--text-main) !important;
        border: 2px solid var(--input-border) !important; box-shadow: none !important; font-weight: bold;
    }
    .stTextInput>div>div>input::placeholder { color: #64748b !important; font-weight: normal !important; }
    
    .stButton>button {
        background-color: var(--btn-bg) !important; color: var(--btn-text) !important;
        border: none !important; font-weight: bold !important; font-size: 1.1rem !important;
        padding: 0.6rem 1.5rem !important; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
    }
    .stButton>button:hover { background-color: #1e293b !important; transform: scale(1.02); }
    
    div[data-testid="stFileUploaderDropzone"] {
        background-color: var(--input-bg) !important; border: 2px dashed var(--input-border) !important;
    }
    div[data-testid="stFileUploaderDropzone"] p { color: var(--text-sub) !important; }
    button[kind="primary"] { background-color: var(--btn-bg) !important; color: var(--btn-text) !important; }
    
    ::-webkit-scrollbar { width: 10px; }
    ::-webkit-scrollbar-track { background: #f1f5f9; }
    ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

st.title("🤖 AI Exam Generator")
st.markdown("### 📌 ระบบสร้างข้อสอบอัจฉริยะ")

# ส่วนตั้งค่า (Sidebar)
st.sidebar.header("⚙️ การตั้งค่า")

# 1. เลือกผู้ให้บริการ
ai_provider = st.sidebar.radio("เลือกผู้ให้บริการ (API)", ["OpenAI", "Google Gemini"])

# 2. แสดงสถานะคีย์ API (แบบ Hardcode)
if ai_provider == "Google Gemini":
    st.sidebar.warning("⚠️ กรุณาใส่คีย์ Google Gemini ในไฟล์โค้ด app.py เพื่อใช้งาน")
else:
    st.sidebar.success("✅ OpenAI API: พร้อมใช้งานแล้ว")

col1, col2 = st.sidebar.columns(2)
with col1:
    question_type = st.selectbox("รูปแบบข้อสอบ", ["Multiple Choice (ปรนัย)", "Short Answer (อัตนัย)", "True/False (ถูกผิด)"])
with col2:
    difficulty = st.selectbox("ระดับความยาก", ["ง่าย", "ปานกลาง", "ยาก"])

num_questions = st.slider("จำนวนข้อสอบ", 1, 20, 5, step=1)

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์เอกสาร (PDF / Word)", type=["pdf", "docx", "txt"])

def extract_text(uploaded_file):
    if uploaded_file is None: return ""
    try:
        if uploaded_file.type == "application/pdf":
            return "\n".join([page.extract_text() for page in PyPDF2.PdfReader(uploaded_file).pages])
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            return "\n".join([para.text for para in Document(io.BytesIO(uploaded_file.read())).paragraphs])
        else:
            return uploaded_file.read().decode("utf-8")
    except Exception as e:
        return f"Error: {str(e)}"

def generate_exam(prompt_text):
    if ai_provider == "Google Gemini":
        if not GOOGLE_API_KEY_DEFAULT:
            return "⚠️ กรุณาใส่คีย์ API ในไฟล์โค้ด"
        try:
            genai.configure(api_key=GOOGLE_API_KEY_DEFAULT)
            model = genai.GenerativeModel("gemini-1.5-flash")
            safety_settings = [
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
            ]
            response = model.generate_content(prompt_text, safety_settings=safety_settings)
            if response.prompt_feedback.block_reason:
                return f"⚠️ คำตอบถูกบล็อก: {response.prompt_feedback.block_reason}"
            return response.text
        except Exception as e:
            return f"เกิดข้อผิดพลาด Google: {str(e)}"

    elif ai_provider == "OpenAI":
        openai.api_key = OPENAI_API_KEY_DEFAULT
        try:
            response = openai.chat.completions.create(
                model="gpt-3.5-turbo", 
                messages=[
                    {"role": "system", "content": "คุณคือครูผู้เชี่ยวชาญ"},
                    {"role": "user", "content": prompt_text}
                ],
                temperature=0.7,
                max_tokens=2000
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"เกิดข้อผิดพลาด OpenAI: {str(e)}"

if uploaded_file is not None:
    text_content = extract_text(uploaded_file)
    
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
        
        if result:
            st.markdown("## 📄 ผลข้อสอบและแบบฝึกหัด:")
            st.markdown(result)
        else:
            st.warning("ยังไม่สามารถสร้างข้อสอบได้")