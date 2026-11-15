import io
import re
import glob
import html
import hashlib
import logging

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


# =========================================================
# Configuration
# =========================================================
st.set_page_config(
    page_title="AI Exam Generator Pro",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

logging.basicConfig(level=logging.ERROR)
logger = logging.getLogger(__name__)

MAX_CHARS = 30000
NON_CHAT_KEYWORDS = (
    "embed", "whisper", "tts", "dall-e", "image", "moderation", "rerank"
)
THINK_TAGS = ("think", "thinking", "reasoning")


# =========================================================
# UI Styles
# =========================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');

    :root {
        --primary: #4F46E5;
        --primary-dark: #4338CA;
        --primary-light: #EEF2FF;
        --bg: #F8FAFC;
        --card: #FFFFFF;
        --text: #1E293B;
        --heading: #0F172A;
        --muted: #64748B;
        --border: #E2E8F0;
        --sidebar: #0F172A;
    }

    html, body, [class*="css"], .stApp {
        font-family: 'Prompt', sans-serif !important;
    }

    .stApp {
        background: var(--bg);
        color: var(--text);
    }

    /* Text selection: avoid browser-default bright blue highlight */
    ::selection {
        background: #C7D2FE;
        color: #1E1B4B;
    }

    h1, h2, h3, h4, h5, h6 {
        color: var(--heading);
        font-family: 'Prompt', sans-serif !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em;
    }

    p, li, label, .stMarkdown, [data-testid="stCaptionContainer"] {
        color: var(--text);
        font-family: 'Prompt', sans-serif !important;
    }

    .main .block-container {
        max-width: 1400px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    /* Sidebar: style text selectively. Do not force every div/span to white. */
