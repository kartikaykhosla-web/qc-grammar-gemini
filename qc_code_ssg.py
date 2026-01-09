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
    seen = set()

    # ---------- TITLE ----------
    title = soup.find("h1")
    if title:
        content.append(("heading", title.get_text(strip=True)))

    # ---------- TRUE FOOTER STOP MARKERS ----------
    HARD_STOP_MARKERS = [
        "our aim is to provide",
        "explore the world",
        "copyright",
        "all rights reserved",
        "for any feedback",
        "compliant_gro@",
    ]

    # ---------- SOFT SKIP MARKERS (DO NOT BREAK) ----------
    SOFT_SKIP_MARKERS = [
        "don't miss",
        "dont miss",
        "also read",
        "click here",
        "follow us",
    ]

    for el in soup.find_all(["p", "li"], recursive=True):
        txt = el.get_text(separator=" ", strip=True)
        # 🔥 FIX: remove space BEFORE punctuation introduced by HTML
        txt = re.sub(r"\s+([,.;:!?])", r"\1", txt)

        if not txt or len(txt) < 20:
            continue

        lower = txt.lower()

        # 🛑 HARD STOP → footer / legal / site boilerplate
        if any(marker in lower for marker in HARD_STOP_MARKERS):
            break

        # ⛔ SKIP widget headers / navigation
        if any(marker in lower for marker in SOFT_SKIP_MARKERS):
            continue

        # ⛔ SKIP numbered widget lists (1…, 2…, 3…)
        if re.match(r"^\d+\s+", txt):
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
    init_vertex_and_model()  # ensures vertexai.init() is called
    model = GenerativeModel("publishers/google/models/gemini-2.5-flash")


    MAX_PARA_CHARS = 1800
    paragraphs = [
        text if len(text) <= MAX_PARA_CHARS else text[:MAX_PARA_CHARS]
        for ctype, text in article_data
        if ctype == "paragraph" and len(text.split()) >= 6
    ]
    paragraphs = paragraphs[:30]  # HARD SAFETY CAP

    BASE_PROMPT = """
You are a professional proofreader and a content QC professional.

Rules (STRICT):
- Review each paragraph independently
- Only fix spelling, grammar, and language-standard issues
- Do NOT change numbers or numerical values
- British English is the ONLY accepted standard
- Convert American English spellings AND formats to British English where applicable
- British date format must be used (e.g., "1 January", not "January 1")
- Date-order corrections are allowed ONLY when the numeric value remains unchanged
- Understand context before suggesting corrections
- If an issue exists once, it MUST be reported every time
- List *every* spelling correction independently, even if it seems minor

PROHIBITIONS:
- NEVER change proper nouns, political parties, or person names
- NEVER rename quoted speakers
- NEVER modify social media platform names or product/platform identifiers
  (e.g., X, Twitter, Facebook, Instagram)
- NEVER modify single-letter proper nouns (e.g., X)
- Do NOT hallucinate
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
- Do NOT normalize whitespace beyond the correction itself
- Periods, commas, apostrophes, abbreviations, and numerals must be preserved exactly

ABBREVIATION SAFETY:
- Single-letter abbreviations followed by a period (e.g., "S.", "X.") are VALID

INLINE CONTENT SAFETY:
- Hyperlinks or anchor text may exist
- Treat input as already-rendered plain text
- Do NOT infer missing spaces caused by HTML

PLATFORM NAME SAFETY:
- The platform "X" must NEVER be interpreted as "A"

EXHAUSTIVENESS REQUIREMENT (MANDATORY):
- You MUST scan the entire TEXT from start to end
- You MUST identify ALL applicable issues before responding
- Do NOT prioritise or skip issues due to importance
- Do NOT stop early once some issues are found
- The output must be COMPLETE and repeatable for the same TEXT

Return output strictly as a table:
| Original | Corrected | Reason |
"""

    responses = []

    # ✅ FIX: Gemini called PER PARAGRAPH (nothing else changed)
    for para in paragraphs:
        prompt = BASE_PROMPT + "\n\nTEXT:\n" + para
        try:
            out = model.generate_content(
                prompt,
                generation_config={
                    "temperature": 0,
                    "top_p": 1
                }
            ).text
            responses.append(out)
        except Exception:
            continue

    if not responses:
        return ""

    raw = "\n".join(responses)

    # 🔧 Existing extraction + dedupe (UNCHANGED)
    matches = re.findall(
        r"\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|",
        raw
    )

    rows = []
    seen = set()

    def canon(s: str) -> str:
        s = s.lower().strip()
        s = re.sub(r"\s+", " ", s)
        return s.strip(".,;:!?")

    for original, corrected, reason in matches:
        if original.strip().lower() == "original":
            continue

        key = (canon(original), canon(corrected), canon(reason))
        if key in seen:
            continue

        seen.add(key)
        rows.append(
            f"| {original.strip()} | {corrected.strip()} | {reason.strip()} |"
        )

    if not rows:
        return ""

    return "\n".join([
        "| Original | Corrected | Reason |",
        "|---|---|---|",
        *rows
    ])


# ============================
# Invalid rows
# ============================
def filter_invalid_rows(gemini_md, article_text):
    lines = gemini_md.splitlines()
    out = []
    seen = set()

    def normalise(text):
        return re.sub(r"[^a-z0-9]", "", text.lower())

    for line in lines:
        if "|" not in line or line.strip().startswith("| Original"):
            out.append(line)
            continue

        cols = [c.strip() for c in line.split("|") if c.strip()]
        if len(cols) != 3:
            continue

        original, corrected, reason = cols

        # must exist verbatim
        if original not in article_text:
            continue

        # 🔥 FIX #1: normalised edit signature (stable dedupe)
        key = (
            normalise(original),
            normalise(corrected),
        )

        if key in seen:
            continue

        seen.add(key)
        out.append(f"| {original} | {corrected} | {reason} |")

    return "\n".join(out)

# ==============================================
# Split spelling Grammar Function
# ==============================================

def split_spelling_grammar(table_md: str):
    spelling_rows = []
    grammar_rows = []

    rows = re.findall(
        r"\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|",
        table_md
    )

    for original, corrected, reason in rows:
        if original.lower() == "original":
            continue

        # 🔑 MINIMAL FIX
        is_spelling = (
            len(original.split()) == 1
            and len(corrected.split()) == 1
            and original != corrected
        )


        row = f"| {original} | {corrected} | {reason} |"

        if is_spelling:
            spelling_rows.append(row)
        else:
            grammar_rows.append(row)

    def build_table(rows):
        if not rows:
            return ""
        return "\n".join(
            ["| Original | Corrected | Reason |",
             "|---|---|---|"] + rows
        )

    return build_table(spelling_rows), build_table(grammar_rows)

# =============================
# Is structural Inline
# =============================
def is_structural_line(text: str) -> bool:
    t = text.strip().lower()

    # Headings / labels
    if len(t.split()) <= 3:
        return True

    # Section headers
    if any(t.startswith(x) for x in [
        "day ",
        "note:",
        "about ",
        "directed by",
        "cast",
    ]):
        return True

    # Lists / bullets
    if re.match(r"^\d+[\).]", t):
        return True

    return False

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
You are an internal factual consistency auditor.

SCOPE (STRICT):
- Treat the TEXT as a closed, self-contained document
- Do NOT use external knowledge, memory, news, timelines, or assumptions
- Do NOT rely on real-world verification
- You may ONLY evaluate statements using information present in the TEXT

SPAN ANCHORING (MANDATORY):
- Only evaluate statements that appear verbatim in the TEXT
- Quote the EXACT sentence fragment under "Statement"
- Do NOT paraphrase, rewrite, or infer

EVALUATION RULES:
- Identify ONLY internal contradictions, misleading implications,
  or factual inconsistencies within the TEXT
- If a statement cannot be verified using the TEXT alone,
  mark the Issue as "Unverifiable from article"
- If no correcting statement exists elsewhere in the TEXT,
  write "Not stated in article" in Correct Fact
- NEVER invent facts, dates, announcements, or corrections

DO NOT:
- Check grammar, spelling, or style
- Rewrite sentences
- Introduce external facts
- Create hypothetical corrections

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

    # ---------- FINAL ARTICLE ----------
    st.subheader("📄 Final Article")
    for _, t in qc_content:
        st.write(t)

    st.divider()

    # ---------- GEMINI QC ----------
    st.subheader("🤖 Gemini QC Review")

    article_text = "\n".join(
        t for c, t in article_content if c == "paragraph"
    )

    # Grammar + Spelling
    raw = gemini_grammar_review(qc_content)
    clean = filter_invalid_rows(raw, article_text)

    spelling_table, grammar_table = split_spelling_grammar(clean)

    st.markdown("### ✍️ Spelling Issues")
    if spelling_table:
        st.markdown(spelling_table)
    else:
        st.success("✅ No spelling issues found")

    st.markdown("### 🧠 Grammar Issues")
    if grammar_table:
        st.markdown(grammar_table)
    else:
        st.success("✅ No grammar issues found")

    # ---------- FACT CHECK ----------
    st.markdown("### 📌 Fact Check")

    fact_result = gemini_fact_check(qc_content)

    if not fact_result or "| Statement |" not in fact_result:
        st.success("✅ No factual issues found")
    else:
        st.markdown(fact_result)
