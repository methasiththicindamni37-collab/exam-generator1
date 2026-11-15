import io
import re
import streamlit as st
import PyPDF2
from docx import Document
from openai import OpenAI

st.set_page_config(page_title="AI Exam Generator", page_icon="🤖", layout="wide")

# ---------- UI (CSS) ----------
st.markdown("""
<style>
    .stApp { background-color: #ffffff; }
    h1, h2, h3, h4, h5, h6, p, label, li, th, td { color: #000000 !important; }
    section[data-testid="stSidebar"] { background-color: #0f172a !important; }
    section[data-testid="stSidebar"] * { color: #ffffff !important; }
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] [data-baseweb="select"] * { color: #000000 !important; }
    .stButton>button, .stDownloadButton>button {
        background-color: #0f172a !important; color: #ffffff !important;
        font-weight: bold; font-size: 1.1rem; padding: 0.6rem 1.5rem; border: none;
    }
    .stButton>button:hover, .stDownloadButton>button:hover { background-color: #1e293b !important; }
</style>
""", unsafe_allow_html=True)

st.title("🤖 AI Exam Generator")
st.markdown("### 📌 ระบบสร้างข้อสอบอัจฉริยะ")

# ---------- Secrets ----------
API_KEY = st.secrets.get("OPENAI_API_KEY", "")
BASE_URL = st.secrets.get("OPENAI_BASE_URL", "https://ai.psu.blue/v1")

MAX_CHARS = 30000

# คำที่บ่งบอกว่าไม่ใช่โมเดลแชต
NON_CHAT_KEYWORDS = (
    "embed", "whisper", "tts", "dall-e", "dalle", "image", "moderation",
    "rerank", "transcribe", "audio", "speech", "stable-diffusion", "sdxl", "flux",
)

# ชื่อแท็กความคิด (reasoning) ที่โมเดลบางตัวส่งออกมา
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
st.sidebar.subheader("🔑 สถานะ API Key")

if API_KEY:
    st.sidebar.success("✅ พบ API Key")
else:
    st.sidebar.error("❌ ยังไม่ได้ตั้ง OPENAI_API_KEY ใน Secrets")

model_name = ""
if API_KEY:
    all_models = list_models(API_KEY, BASE_URL)
    if all_models:
        show_all = st.sidebar.checkbox("แสดงโมเดลทั้งหมด (รวมที่ไม่ใช่แชต)", value=False)
        models = all_models if show_all else [m for m in all_models if is_chat_model(m)]
        if not models:
            models = all_models
        model_name = st.sidebar.selectbox("เลือกโมเดล", models)
    else:
        st.sidebar.warning("ไม่พบรายการโมเดล (อาจเป็นปัญหาการเชื่อมต่อ)")
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

        # 1) ตัดบล็อกที่มีทั้งเปิดและปิด
        pattern = re.escape(open_t) + r".*?" + re.escape(close_t)
        text = re.sub(pattern, "", text, flags=re.DOTALL | re.IGNORECASE)

        # 2) มีแต่แท็กปิด (บางโมเดลไม่ส่งแท็กเปิด) -> เก็บเฉพาะส่วนหลังแท็กปิด
        idx = text.lower().rfind(close_t)
        if idx != -1:
            text = text[idx + len(close_t):]

        # 3) มีแท็กเปิดแต่ยังไม่ปิด (ระหว่างสตรีม) -> ซ่อนส่วนที่ค้างไว้ก่อน
        oi = text.lower().find(open_t)
        if oi != -1:
            text = text[:oi]
    return text.strip()


def make_docx(markdown_text: str) -> bytes:
    doc = Document()
    doc.add_heading("ข้อสอบ", level=1)
    for line in markdown_text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("### "):
            doc.add_heading(stripped[4:], level=3)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:], level=2)
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:], level=1)
        else:
            doc.add_paragraph(stripped.replace("**", ""))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _err_text(e) -> str:
    return str(e).lower()


def _call_api(client, model, messages, params, stream, placeholder):
    """เรียก API หนึ่งครั้งด้วยพารามิเตอร์ที่กำหนด"""
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

    # ลำดับการลองแบบถอยหลัง (fallback) เพื่อให้ใช้ได้กับหลายโมเดล
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
                    # ข้อผิดพลาดที่ลองใหม่ก็ไม่ช่วย
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
                    # นอกนั้นลองรูปแบบถัดไป
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
2. หากเป็นข้อสอบปรนัย ต้องมีตัวเลือก ก. ข. ค. ง.
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

    # แสดงผลล่าสุด + ปุ่มดาวน์โหลด (กรณีหน้า rerun)
    result = st.session_state.exam_result
    if result and not result.startswith(("⚠️", "⛔")):
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ ดาวน์โหลด (.md)", result.encode("utf-8"),
                file_name="exam.md", mime="text/markdown",
            )
        with d2:
            st.download_button(
                "⬇️ ดาวน์โหลด (.docx)", make_docx(result),
                file_name="exam.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    elif result:
        st.warning(result)
else:
    st.info("กรุณาอัปโหลดไฟล์เอกสารเพื่อเริ่มต้น")
