# -*- coding: utf-8 -*-
"""
QC Code SSG — Vertex AI Gemini 2.5
OLD FORMAT RESTORED + EDITORIAL SAFETY FIXES
"""

# ===================== CORE =====================
import re
import os
import json
import base64
import requests
import tempfile
import streamlit as st
import html
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from difflib import SequenceMatcher


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
from language_tool_python.exceptions import RateLimitError

# ===================== VERTEX AI =====================
import vertexai
from vertexai.generative_models import GenerativeModel


# =================================================
# STREAMLIT CONFIG
# =================================================
st.set_page_config(page_title="Article QC Tool (Gemini 2.5)", layout="wide")
st.title("🧪 Article QC Tool (Gemini 2.5 – Vertex AI)")
st.caption("Spelling · Grammar · Editorial Safety · AI Review")


# =================================================
# 🔑 VERTEX AI AUTH (BASE64 SAFE)
# =================================================
PROJECT_ID = "prod-project-jnm-smart-cms"
REGION = "us-central1"
CRED_PATH = "/tmp/gcp_service_account.json"


def load_gcp_credentials():
    if "GCP_SERVICE_ACCOUNT_JSON_B64" not in st.secrets:
        st.error("❌ GCP_SERVICE_ACCOUNT_JSON_B64 not set in Streamlit secrets")
        st.stop()

    try:
        decoded = base64.b64decode(
            st.secrets["GCP_SERVICE_ACCOUNT_JSON_B64"]
        ).decode("utf-8")
        creds_dict = json.loads(decoded)
    except Exception as e:
        st.error("❌ Invalid Base64 GCP credential")
        st.exception(e)
        st.stop()

    with open(CRED_PATH, "w") as f:
        json.dump(creds_dict, f)

    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = CRED_PATH
    return service_account.Credentials.from_service_account_info(creds_dict)


@st.cache_resource
def init_vertex_and_model():
    creds = load_gcp_credentials()

    vertexai.init(
        project=PROJECT_ID,
        location=REGION,
        credentials=creds,
    )

    try:
        st.success("✅ Gemini 2.5 Pro loaded")
        return GenerativeModel("publishers/google/models/gemini-2.5-pro")
    except Exception:
        st.warning("⚠️ Falling back to Gemini 2.5 Flash")
        return GenerativeModel("publishers/google/models/gemini-2.5-flash")


# =================================================
# INPUT EXTRACTION (INLINE-SAFE)
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

    content = []

    # ---------- 1️⃣ TRY STRUCTURED ARTICLE BODY (BEST SOURCE) ----------
    ld_json_blocks = soup.find_all("script", type="application/ld+json")

for block in ld_json_blocks:
    try:
        data = json.loads(block.string)

        if isinstance(data, dict) and "articleBody" in data:
            body = html.unescape(data["articleBody"])

            # ⛔ Guard against publisher-stripped punctuation
            punctuation_density = sum(body.count(p) for p in [".", ",", "?", "!"])
            if punctuation_density < len(body) * 0.005:
                break  # FALL BACK to HTML extraction

            paragraphs = [
                p.strip()
                for p in re.split(r"\n{2,}", body)
                if len(p.strip()) > 15
            ]

            if "headline" in data:
                content.append(("heading", data["headline"].strip()))

            for p in paragraphs:
                content.append(("paragraph", p))

            return content

    except Exception:
        pass


    # ---------- 2️⃣ FALLBACK: RAW HTML EXTRACTION ----------
    article = (
        soup.find("article")
        or soup.find("div", class_=lambda c: c and "article" in c.lower())
        or soup
    )

    seen = set()

    title = soup.find("h1")
    if title:
        content.append(("heading", title.get_text(strip=True)))

    for el in article.find_all(["p", "li"], recursive=True):
        txt = el.get_text(separator=" ", strip=True)

        if not txt or len(txt) < 15:
            continue

        lower = txt.lower()
        if any(x in lower for x in [
            "also read",
            "click here",
            "disclaimer",
            "follow us",
            "to read more articles"
        ]):
            continue

        if txt in seen:
            continue

        seen.add(txt)
        content.append(("paragraph", txt))

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
# SPELLING + GRAMMAR
# =================================================
STOPWORDS = {
    "a","an","the","of","to","in","on","at","for","from","by",
    "and","or","but","so","are","is","was","were","be","been","being"
}


def correct_spelling_minimal(text):
    tokens = text.split()
    out = []

    entities = {ent.text for ent in nlp(text).ents}
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

    try:
        return language_tool_python.utils.correct(text, lt_tool.check(text))
    except RateLimitError:
        return text
    except Exception:
        return text


# =================================================
# 🔍 VERBATIM DIFF VALIDATOR + CONFIDENCE SCORE
# =================================================
def filter_gemini_rows(raw_table, article_text):
    lines = raw_table.splitlines()
    output = []

    header_added = False

    for line in lines:
        if line.strip().startswith("| Original"):
            output.append("| Original | Corrected | Reason |")
            output.append("|---|---|---|")
            header_added = True
            continue

        if "|" not in line:
            continue

        cols = [c.strip() for c in line.split("|") if c.strip()]
        if len(cols) != 3:
            continue

        original, corrected, reason = cols

        # Verbatim safety: original must exist exactly in article text
        if original in article_text:
            output.append(f"| {original} | {corrected} | {reason} |")

    return "\n".join(output) if header_added else ""

# =================================================
# GEMINI QC — OLD TABLE FORMAT (SAFE)
# =================================================
def gemini_grammar_review(article_data):
    model = init_vertex_and_model()

    MAX_PARA_CHARS = 1800
    paragraphs = [
        text if len(text) <= MAX_PARA_CHARS else text[:MAX_PARA_CHARS]
        for ctype, text in article_data
        if ctype == "paragraph"
    ]

    BASE_PROMPT = """
You are a professional proofreader.

Rules (STRICT):
- Review each paragraph independently
- Only fix spelling and grammar
- Do NOT change numbers
- British English is the ONLY accepted standard
- Convert American English spellings to British English where applicable
- NEVER change proper nouns, political parties, or person names
- NEVER rename quoted speakers
- NEVER modify social media platform names or product/platform identifiers
  (e.g., X, Twitter, Facebook, Instagram)
- NEVER modify single-letter proper nouns (e.g., X)
- If unsure, do NOT hallucinate
- Do NOT normalize legal, political, or platform references

CRITICAL CONSTRAINTS:
- You may ONLY use text that appears verbatim in the TEXT section
- NEVER invent new examples, phrases, or sentences
- The "Original" column MUST be an exact, character-for-character substring
  of the provided TEXT
- If no correction is required, DO NOT create a table row
- If you cannot find an exact match in the TEXT, do NOT include it

ABSOLUTE RULE:
- Treat the TEXT as a raw byte string
- Do NOT normalize whitespace, punctuation, or casing
- Periods, commas, apostrophes, and abbreviations must be preserved exactly

ABBREVIATION SAFETY:
- Single-letter abbreviations followed by a period (e.g., "S.", "X.") are VALID

INLINE CONTENT SAFETY:
- Hyperlinks or anchor text may exist
- Treat input as already-rendered plain text
- Do NOT infer missing spaces caused by HTML

PLATFORM NAME SAFETY:
- The platform "X" must NEVER be interpreted as "A"

Return output strictly as a table:
| Original | Corrected | Reason |
"""

    responses = []

    # 1️⃣ Call Gemini per paragraph
    for para in paragraphs:
        prompt = BASE_PROMPT + "\n\nTEXT:\n" + para
        try:
            resp = model.generate_content(prompt).text
            responses.append(resp)
        except Exception:
            continue

    # 2️⃣ Collect ONLY table rows (no duplicate headers)
    rows = []

    for resp in responses:
        for line in resp.splitlines():
            line = line.strip()
            if not line:
                continue

            # Skip headers & separators
            if line.startswith("| Original"):
                continue
            if line.startswith("|---"):
                continue

            # Keep only valid table rows
            if line.count("|") >= 3:
                rows.append(line)

    if not rows:
        return ""

    # 3️⃣ Build ONE clean table
    table = [
        "| Original | Corrected | Reason |",
        "|---|---|---|",
        *rows
    ]

    return "\n".join(table)

# ============================
# Invalid rows
# ============================
def filter_invalid_rows(gemini_md, article_text):
    lines = gemini_md.splitlines()
    out = []

    for line in lines:
        if "|" not in line or line.strip().startswith("| Original"):
            out.append(line)
            continue

        cols = [c.strip() for c in line.split("|")]
        if len(cols) < 3:
            continue

        original = cols[1]

        # 1️⃣ Must exist verbatim
        if original not in article_text:
            continue

        # 2️⃣ Must NOT span sentence boundaries
        # Reject if original crosses punctuation joins
        boundary_pattern = re.escape(original).replace("\\ ", r"\s+")
        if not re.search(rf"(?<![.!?]){boundary_pattern}(?![.!?])", article_text):
            continue

        # 3️⃣ Reject multi-sentence hallucinations
        if any(p in original for p in [". ", "? ", "! "]):
            continue

        out.append(line)

    return "\n".join(out)

# =================================================
# FACT CHECK — SECOND PASS (ADDED, ISOLATED)
# =================================================
def gemini_fact_check(article_data):
    model = init_vertex_and_model()

    paragraphs = [
    text for ctype, text in article_data
    if ctype == "paragraph" and not is_structural_line(text)
    ]

    fact_prompt = f"""
You are a strict factual verifier.

Rules (STRICT):
- ONLY check factual correctness
- Identify statements that are factually incorrect or misleading
- Do NOT check grammar, spelling, punctuation, or style
- Do NOT rewrite sentences
- Do NOT infer intent
- If a statement is unverifiable, mark it as "Unverifiable"
- If a statement is factually correct, DO NOT include it

Return output strictly as a table:
| Statement | Issue | Correct Fact |

TEXT:
{chr(10).join(paragraphs)}
"""

    try:
        return model.generate_content(fact_prompt).text
    except Exception as e:
        return f"⚠️ Fact check unavailable:\n\n{e}"

# =================================================
# PIPELINE
# =================================================
def run_pipeline(content):
    final = []
    for ctype, text in content:
        if ctype != "paragraph":
            final.append((ctype, text))
            continue

        # ❗ DO NOT mutate article text before Gemini
        final.append((ctype, text))

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

        raw = gemini_grammar_review(qc_content)
        article_text = "\n".join(t for _, t in qc_content)

        clean = filter_invalid_rows(raw, article_text)
        st.markdown(clean)

    st.divider()

    if st.button("🔍 Run Fact Check (Second Pass)"):
        st.subheader("📌 Fact Check Results")
        st.markdown(gemini_fact_check(qc_content))



