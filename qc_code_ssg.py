# -*- coding: utf-8 -*-
"""
QC Code SSG — Vertex AI Gemini 2.5
OLD FORMAT RESTORED + EDITORIAL SAFETY FIXES
"""

# ===================== CORE =====================
import re
import os
import requests
import tempfile
import streamlit as st
from bs4 import BeautifulSoup

# ===================== NLP =====================
import spacy

# ===================== SPELL & FREQUENCY =====================
from spellchecker import SpellChecker
from wordfreq import zipf_frequency
from symspellpy import SymSpell, Verbosity
from importlib.resources import files as pkg_files

# ===================== DOCX =====================
from docx import Document

# ===================== GRAMMAR =====================
import language_tool_python

# ===================== VERTEX AI =====================
import vertexai
from vertexai.generative_models import GenerativeModel


# =================================================
# STREAMLIT CONFIG (OLD STYLE)
# =================================================
st.set_page_config(page_title="Article QC Tool (Gemini 2.5)", layout="wide")
st.title("🧪 Article QC Tool (Gemini 2.5 – Vertex AI)")
st.caption("Spelling · Grammar · Editorial Safety · AI Review")


# =================================================
# 🔑 VERTEX AI AUTH (SAFE FOR STREAMLIT CLOUD)
# =================================================
if "GCP_SERVICE_ACCOUNT_JSON" not in os.environ:
    st.error("❌ GCP_SERVICE_ACCOUNT_JSON not set")
    st.stop()

if "GOOGLE_APPLICATION_CREDENTIALS" not in os.environ:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as f:
        f.write(os.environ["GCP_SERVICE_ACCOUNT_JSON"].encode("utf-8"))
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = f.name

PROJECT_ID = "prod-project-jnm-smart-cms"
REGION = "us-central1"


@st.cache_resource
def init_vertex():
    vertexai.init(project=PROJECT_ID, location=REGION)
    return True


try:
    init_vertex()
    gemini_model = GenerativeModel("publishers/google/models/gemini-2.5-pro")
    st.success("✅ Gemini 2.5 Pro loaded")
except Exception:
    gemini_model = GenerativeModel("publishers/google/models/gemini-2.5-flash")
    st.warning("⚠️ Falling back to Gemini 2.5 Flash")


# =================================================
# INPUT EXTRACTION
# =================================================
def clean_docx(file_path):
    doc = Document(file_path)
    content, seen = [], set()

    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt or txt in seen or len(txt) < 15:
            continue
        content.append(("paragraph", txt))
        seen.add(txt)

    return content


def clean_article(url):
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    article = (
        soup.find("article")
        or soup.find("div", class_=lambda c: c and "article" in c)
        or soup
    )

    content, seen = [], set()

    title = soup.find("h1")
    if title:
        content.append(("heading", title.get_text(strip=True)))

    for el in article.find_all(["p", "li"], recursive=True):
        txt = el.get_text(" ", strip=True)

        if not txt or len(txt) < 15:
            continue

        if any(j in txt.lower() for j in [
            "also read",
            "click here",
            "disclaimer:",
            "follow us",
            "to read more articles"
        ]):
            continue

        if txt in seen:
            continue

        content.append(("paragraph", txt))
        seen.add(txt)

    return content


# =================================================
# LOAD MODELS
# =================================================
@st.cache_resource
def load_nlp():
    return spacy.load("en_core_web_sm")


@st.cache_resource
def load_languagetool():
    try:
        return language_tool_python.LanguageToolPublicAPI("en-US")
    except Exception:
        return None


@st.cache_resource
def load_symspell():
    sym = SymSpell(max_dictionary_edit_distance=2, prefix_length=7)
    dictionary_path = str(
        pkg_files("symspellpy").joinpath("frequency_dictionary_en_82_765.txt")
    )
    sym.load_dictionary(dictionary_path, 0, 1)
    return sym


nlp = load_nlp()
lt_tool = load_languagetool()
sym_spell = load_symspell()
spell = SpellChecker()


# =================================================
# SAFETY HELPERS
# =================================================
def extract_named_entities(text):
    return {ent.text for ent in nlp(text).ents}


def extract_numbers(text):
    return re.findall(r"\d[\d.,]*", text)


def semantic_safe(a, b, t=0.92):
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio() >= t


# =================================================
# SPELLING + GRAMMAR
# =================================================
STOPWORDS = {
    "a","an","the","of","to","in","on","at","for","from","by",
    "and","or","but","so","are","is","was","were","be","been","being"
}


def correct_spelling_minimal(text):
    tokens = text.split()
    out = []

    entities = extract_named_entities(text)
    entity_tokens = {w for e in entities for w in e.split()}

    for tok in tokens:
        core = re.sub(r"[^\w]", "", tok)
        lower = core.lower()

        if not core or core in entity_tokens or lower in STOPWORDS:
            out.append(tok)
            continue

        suggestions = sym_spell.lookup(lower, Verbosity.CLOSEST, 2)
        best = suggestions[0].term if suggestions else spell.correction(lower)

        if best and best != lower and zipf_frequency(best, "en") > zipf_frequency(lower, "en"):
            best = best.capitalize() if core[0].isupper() else best
            out.append(tok.replace(core, best))
        else:
            out.append(tok)

    return " ".join(out)


def correct_grammar_languagetool(text):
    if not lt_tool:
        return text

    corrected = language_tool_python.utils.correct(
        text, lt_tool.check(text)
    )

    if extract_numbers(text) != extract_numbers(corrected):
        return text
    if not extract_named_entities(text).issubset(
        extract_named_entities(corrected)
    ):
        return text

    return corrected


# =================================================
# GEMINI QC — OLD TABLE FORMAT
# =================================================
def gemini_grammar_review(article_data):
    init_vertex()

    paragraphs = [
        text[:900]
        for ctype, text in article_data
        if ctype == "paragraph"
    ]

    prompt = f"""
You are a professional proofreader.

Rules:
- Only fix spelling and grammar
- Do NOT infer facts
- Do NOT change numbers
- Do NOT rename people or cases
- Do NOT normalize legal citations

Return output strictly as a table:
| Original | Corrected | Reason |

TEXT:
{chr(10).join(paragraphs)}
"""

    return gemini_model.generate_content(prompt).text


# =================================================
# PIPELINE
# =================================================
def run_pipeline(content):
    final = []
    for ctype, text in content:
        if ctype != "paragraph":
            final.append((ctype, text))
            continue

        step1 = correct_spelling_minimal(text)
        step2 = correct_grammar_languagetool(step1)
        final.append((ctype, step2))

    return final


# =================================================
# STREAMLIT UI
# =================================================
st.sidebar.header("Input")
source = st.sidebar.radio("Source", ["URL", "DOCX"])

article_content = None

if source == "URL":
    url = st.sidebar.text_input("Article URL")
    if st.sidebar.button("Analyze") and url:
        article_content = clean_article(url)
else:
    uploaded = st.sidebar.file_uploader("Upload DOCX", type=["docx"])
    if uploaded:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
            f.write(uploaded.read())
            article_content = clean_docx(f.name)

if article_content:
    qc_content = run_pipeline(article_content)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 Final Article")
        for _, t in qc_content:
            st.write(t)

    with col2:
        st.subheader("🤖 Gemini QC Review")
        st.markdown(gemini_grammar_review(qc_content))
