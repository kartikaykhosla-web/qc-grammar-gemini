# -*- coding: utf-8 -*-
"""
QC Code SSG — Vertex AI Gemini 2.5
OLD FORMAT RESTORED + EDITORIAL SAFETY FIXES
"""
FACT_CACHE = {}

# ===================== CORE =====================
import re
import os
import json
import base64
import requests
import hashlib
import tempfile
import streamlit as st
import html
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
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


# ===================== GEN AI CLIENT =====================
from google import genai
from google.genai import types as genai_types


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
MODEL_PRO = "gemini-2.5-pro"
MODEL_FLASH = "gemini-2.5-flash"
CLOUD_PLATFORM_SCOPE = "https://www.googleapis.com/auth/cloud-platform"


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
    return service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=[CLOUD_PLATFORM_SCOPE],
    )

@st.cache_resource
def init_vertex_and_model():
    creds = load_gcp_credentials()

    client = genai.Client(
        vertexai=True,
        project=PROJECT_ID,
        location=REGION,
        credentials=creds,
    )

    # Try pro, fallback to flash
    try:
        client.models.generate_content(
            model=MODEL_PRO,
            contents="Warmup",
            config=genai_types.GenerateContentConfig(
                temperature=0,
                topP=1,
                maxOutputTokens=8,
            ),
        )
        st.success("✅ Gemini 2.5 Pro loaded")
        default_model = MODEL_PRO
    except Exception:
        client.models.generate_content(
            model=MODEL_FLASH,
            contents="Warmup",
            config=genai_types.GenerateContentConfig(
                temperature=0,
                topP=1,
                maxOutputTokens=8,
            ),
        )
        st.warning("⚠️ Falling back to Gemini 2.5 Flash")
        default_model = MODEL_FLASH

    return client, default_model

def build_generate_config(generation_config=None):
    cfg = generation_config or {}
    return genai_types.GenerateContentConfig(
        temperature=cfg.get("temperature"),
        topP=cfg.get("top_p"),
        topK=cfg.get("top_k"),
        candidateCount=cfg.get("candidate_count"),
        maxOutputTokens=cfg.get("max_output_tokens"),
    )

def generate_text(prompt, generation_config=None, model_name=None):
    client, default_model = init_vertex_and_model()
    response = client.models.generate_content(
        model=model_name or default_model,
        contents=prompt,
        config=build_generate_config(generation_config),
    )
    return response.text or ""

def generate_stream_text(prompt, generation_config=None, model_name=None):
    client, default_model = init_vertex_and_model()
    stream = client.models.generate_content_stream(
        model=model_name or default_model,
        contents=prompt,
        config=build_generate_config(generation_config),
    )
    chunks = []
    for chunk in stream:
        if getattr(chunk, "text", None):
            chunks.append(chunk.text)
    return "".join(chunks)


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
        return language_tool_python.LanguageToolPublicAPI("en-GB")
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

GRAMMAR_MAX_CHARS = 20000
FACT_MAX_CHARS = 20000
FACT_MAX_ITEMS = 40

# =================================================
# EDITORIAL QC RULES (NEW)
# =================================================
US_UK_SPELLINGS = {
    "organize": "organise",
    "organizes": "organises",
    "organized": "organised",
    "organizing": "organising",
    "organization": "organisation",
    "analyze": "analyse",
    "analyzes": "analyses",
    "analyzed": "analysed",
    "analyzing": "analysing",
    "center": "centre",
    "color": "colour",
    "colors": "colours",
    "favor": "favour",
    "favored": "favoured",
    "favorite": "favourite",
    "behavior": "behaviour",
    "defense": "defence",
    "license": "licence",
    "offense": "offence",
    "traveling": "travelling",
    "traveled": "travelled",
    "canceled": "cancelled",
    "counseling": "counselling",
    "program": "programme",
    "gray": "grey",
}

HYPERBOLE_WORDS = [
    "best",
    "awesome",
    "amazing",
    "incredible",
    "mind-blowing",
    "unbelievable",
    "stunning",
    "ultimate",
    "perfect",
    "must-see",
    "must watch",
]

CLICHE_PHRASES = [
    "on the other hand",
    "moreover",
    "in this regard",
    "do the needful",
    "at the end of the day",
    "in a nutshell",
]

URL_RE = re.compile(r"\b(?:https?://|www\.)\S+\b", re.IGNORECASE)
ACRONYM_DOT_RE = re.compile(r"\b(?:[A-Z]\.\s?){2,}\b")
HONORIFIC_RE = re.compile(r"\b(Mr|Mrs|Ms)\.?\b")
HONOURS_PREFIX_RE = re.compile(
    r"\b(Bharat Ratna|Padma (?:Shri|Bhushan|Vibhushan))\s+([A-Z][\w'-]+)\b"
)
AGE_NO_HYPHEN_RE = re.compile(
    r"\b(\d{1,3})\s*(?:year|yr|yrs)\s*old\b", re.IGNORECASE
)
AGE_BAD_HYPHEN_RE = re.compile(r"\b(\d{1,3})-year\s+old\b", re.IGNORECASE)
RANGE_RE = re.compile(r"\b[0-9]\s*[–-]\s*[0-9]\b")
UNIT_RE = re.compile(
    r"\b(\d+(?:\.\d+)?)\s*(kg|kgs|KG|KGS|km|kms|KM|KMS|kmph|KMPH|KMPh|km/h|KM/H)\b"
)
LIST_ENUM_RE = re.compile(r"^\s*\(?\d+[.)]")
LEGAL_CONTEXT_RE = re.compile(
    r"\b(section|sec\.?|article|art\.?|rule|order|clause|schedule|"
    r"sub-?section|sub-?clause|act|ipc|crpc|bnss|bns|cpc)\b",
    re.IGNORECASE
)
LEGAL_REF_RE = re.compile(r"\b\d+\s*\(\d+\)\b")
DATE_TIME_RE = re.compile(
    r"\b(\d{1,2}:\d{2}\s*(?:am|pm)?)\b|\b\d{1,2}\s+[A-Za-z]{3,9}\b",
    re.IGNORECASE
)
DGMENT_RE = re.compile(r"\b[A-Za-z]*dgment(s)?\b")
QUOTE_RE = re.compile(r"[\"“”]")
EXPERT_HINT_RE = re.compile(
    r"\b(Dr|Doctor|Professor|expert|dermatologist|nutritionist|dietitian|doctor|"
    r"physician|surgeon|psychologist|cardiologist|gynecologist|gynaecologist|"
    r"paediatrician|pediatrician|dentist|trichologist)\b",
    re.IGNORECASE
)

HYPERBOLE_RE = re.compile(
    r"\b(" + "|".join(re.escape(w) for w in HYPERBOLE_WORDS) + r")\b",
    re.IGNORECASE
)
CLICHE_RE = re.compile(
    r"\b(" + "|".join(re.escape(p) for p in CLICHE_PHRASES) + r")\b",
    re.IGNORECASE
)

UNIT_NORMALIZATION = {
    "kg": "kg",
    "kgs": "kg",
    "km": "km",
    "kms": "km",
    "kmph": "kmph",
    "km/h": "km/h",
}


def _preserve_case(source, replacement):
    if source.isupper():
        return replacement.upper()
    if source[:1].isupper():
        return replacement.capitalize()
    return replacement


def _excerpt(text, start, end, width=40):
    left = max(0, start - width)
    right = min(len(text), end + width)
    snippet = text[left:right].replace("\n", " ").strip()
    return snippet


def _escape_md(text):
    return text.replace("|", "\\|").replace("\n", " ").strip()


def _render_issues_table(issues):
    if not issues:
        return ""
    lines = [
        "| Rule | Location | Excerpt | Suggestion |",
        "| --- | --- | --- | --- |",
    ]
    for issue in issues:
        rule = _escape_md(issue.get("rule", ""))
        location = _escape_md(issue.get("location", ""))
        excerpt = _escape_md(issue.get("excerpt", ""))[:160]
        suggestion = _escape_md(issue.get("suggestion", ""))
        lines.append(f"| {rule} | {location} | {excerpt} | {suggestion} |")
    return "\n".join(lines)


def _batch_texts(texts, max_chars):
    batches = []
    current = []
    size = 0

    for text in texts:
        t = text.strip()
        if not t:
            continue
        add_len = len(t) + 2
        if current and size + add_len > max_chars:
            batches.append("\n\n".join(current))
            current = [t]
            size = len(t)
        else:
            current.append(t)
            size += add_len

    if current:
        batches.append("\n\n".join(current))

    return batches


def _batch_statements(statements, max_chars, max_items):
    batches = []
    current = []
    size = 0

    for stmt in statements:
        line = f"- {stmt}"
        add_len = len(line) + 1
        if current and (size + add_len > max_chars or len(current) >= max_items):
            batches.append(current)
            current = [stmt]
            size = add_len
        else:
            current.append(stmt)
            size += add_len

    if current:
        batches.append(current)

    return batches


def editorial_qc_checks(content, web_story=False, health_beauty=False):
    issues = []
    seen_issues = set()
    paragraph_index = 0
    found_expert_quote = False

    def add_issue(rule, location, excerpt, suggestion):
        key = (
            rule,
            location,
            re.sub(r"\s+", " ", excerpt.strip().lower())
        )
        if key in seen_issues:
            return
        seen_issues.add(key)
        issues.append({
            "rule": rule,
            "location": location,
            "excerpt": excerpt,
            "suggestion": suggestion,
        })

    for ctype, text in content:
        if not text:
            continue

        if ctype == "paragraph":
            paragraph_index += 1
            location = f"Paragraph {paragraph_index}"
        else:
            location = "Heading"

        doc = nlp(text)
        entity_spans = [(ent.start_char, ent.end_char) for ent in doc.ents]

        def in_entity(start, end):
            return any(start >= s and end <= e for s, e in entity_spans)

        # British English spellings (skip official names/entities)
        for us, uk in US_UK_SPELLINGS.items():
            for match in re.finditer(rf"\b{re.escape(us)}\b", text, re.IGNORECASE):
                if in_entity(match.start(), match.end()):
                    continue
                suggestion = _preserve_case(match.group(), uk)
                add_issue(
                    "British English spelling",
                    location,
                    _excerpt(text, match.start(), match.end()),
                    f"Use '{suggestion}'"
                )

        # British English: -dgment -> -dgement (dynamic, not hardcoded)
        for match in DGMENT_RE.finditer(text):
            if in_entity(match.start(), match.end()):
                continue
            word = match.group(0)
            if "dgement" in word.lower():
                continue
            suggestion = re.sub(r"dgment(s)?$", r"dgement\1", word, flags=re.IGNORECASE)
            suggestion = _preserve_case(word, suggestion)
            add_issue(
                "British English spelling",
                location,
                _excerpt(text, match.start(), match.end()),
                f"Use '{suggestion}'"
            )

        # No ampersands unless part of an official name/entity
        for match in re.finditer(r"&", text):
            if in_entity(match.start(), match.end()):
                continue
            add_issue(
                "No ampersands",
                location,
                _excerpt(text, match.start(), match.end()),
                "Replace '&' with 'and'"
            )

        # Numerals 0-9 in body text (headlines and ranges allowed)
        if ctype == "paragraph":
            range_spans = [(m.start(), m.end()) for m in RANGE_RE.finditer(text)]
            for match in re.finditer(r"\b[0-9]\b", text):
                # Skip list numbering at the start of a line
                if match.start() <= 3 and LIST_ENUM_RE.match(text):
                    continue
                # Skip legal references and section/article numbering
                left = max(0, match.start() - 20)
                right = min(len(text), match.end() + 20)
                context = text[left:right]
                if LEGAL_CONTEXT_RE.search(context) or LEGAL_REF_RE.search(context):
                    continue
                # Skip dates and times
                if DATE_TIME_RE.search(context):
                    continue
                if any(match.start() >= s and match.end() <= e for s, e in range_spans):
                    continue
                add_issue(
                    "Number style (0–9)",
                    location,
                    _excerpt(text, match.start(), match.end()),
                    "Spell out numbers 0–9 in body text"
                )

        # Measurement units: lowercase, singular
        for match in UNIT_RE.finditer(text):
            number = match.group(1)
            unit_raw = match.group(2)
            unit_key = unit_raw.lower()
            if unit_key not in UNIT_NORMALIZATION:
                continue
            corrected_unit = UNIT_NORMALIZATION[unit_key]
            if unit_raw != corrected_unit:
                add_issue(
                    "Unit formatting",
                    location,
                    _excerpt(text, match.start(), match.end()),
                    f"Use '{number} {corrected_unit}'"
                )

        # Acronyms should not include periods/spaces
        for match in ACRONYM_DOT_RE.finditer(text):
            cleaned = re.sub(r"[.\s]", "", match.group())
            add_issue(
                "Acronym punctuation",
                location,
                _excerpt(text, match.start(), match.end()),
                f"Use '{cleaned}'"
            )

        # Hyperlinks: none in first two paragraphs; none at all for web stories
        if URL_RE.search(text):
            if web_story or (ctype == "paragraph" and paragraph_index <= 2):
                add_issue(
                    "Hyperlink placement",
                    location,
                    _excerpt(text, 0, min(len(text), 80)),
                    "Remove hyperlinks (web story)" if web_story
                    else "Avoid external hyperlinks in the first two paragraphs"
                )

        # Hyperbole and cliches
        for match in HYPERBOLE_RE.finditer(text):
            add_issue(
                "Avoid superlatives",
                location,
                _excerpt(text, match.start(), match.end()),
                "Remove or substantiate the claim"
            )
        for match in CLICHE_RE.finditer(text):
            add_issue(
                "Avoid cliches",
                location,
                _excerpt(text, match.start(), match.end()),
                "Rephrase in a fresh, direct way"
            )

        # Honorifics
        for match in HONORIFIC_RE.finditer(text):
            add_issue(
                "Honorifics",
                location,
                _excerpt(text, match.start(), match.end()),
                "Remove Mr/Mrs/Ms"
            )

        # Honours as suffixes
        for match in HONOURS_PREFIX_RE.finditer(text):
            award = match.group(1)
            name = match.group(2)
            add_issue(
                "Honours as suffix",
                location,
                _excerpt(text, match.start(), match.end()),
                f"Use '{award} awardee {name}'"
            )

        # Age formatting
        for match in AGE_BAD_HYPHEN_RE.finditer(text):
            add_issue(
                "Age formatting",
                location,
                _excerpt(text, match.start(), match.end()),
                "Use '38-year-old' or 'Name, 38,'"
            )
        for match in AGE_NO_HYPHEN_RE.finditer(text):
            if re.search(rf"\b{match.group(1)}-year-old\b", text, re.IGNORECASE):
                continue
            add_issue(
                "Age formatting",
                location,
                _excerpt(text, match.start(), match.end()),
                "Use '38-year-old' or 'Name, 38,'"
            )

        # Expert quote detection (health/beauty)
        if health_beauty and QUOTE_RE.search(text) and EXPERT_HINT_RE.search(text):
            found_expert_quote = True

    if health_beauty and not found_expert_quote:
        add_issue(
            "Expert quote required",
            "Overall",
            "No expert quote detected",
            "Add a quote from a relevant expert"
        )

    return issues


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
        if original not in article_text:
            continue

        if is_noop_correction(original, corrected, reason):
            continue

        output.append(f"| {original} | {corrected} | {reason} |")

    return "\n".join(output) if header_added else ""

# =================================================
# GEMINI QC — OLD TABLE FORMAT (SAFE)
# =================================================
def gemini_grammar_review(article_data, max_chars=GRAMMAR_MAX_CHARS):
    init_vertex_and_model()  # ensures vertexai.init() is called

    raw_paragraphs = [
        text
        for ctype, text in article_data
        if ctype == "paragraph" and len(text.split()) >= 6
    ]

    if not raw_paragraphs:
        return ""

    joined = "\n\n".join(raw_paragraphs)
    if len(joined) <= max_chars:
        paragraphs = [joined]
    else:
        paragraphs = _batch_texts(raw_paragraphs, max_chars)

    BASE_PROMPT = """
You are a professional proofreader and a content QC professional.

You are a professional proofreader and a content QC professional.

For maximum editorial efficiency, strictly apply the following stylistic directives:

STYLE & FORMATTING DIRECTIVES (MANDATORY):
- Use the MM DD, YYYY format for all dates.
- Eliminate periods in names, academic degrees, and titles
  (e.g., APJ Abdul Kalam, MSc, Dr — NOT A.P.J. Abdul Kalam, M.Sc., Dr.).
- Spell out the full name at first mention followed by the abbreviation;
  use the abbreviation for all subsequent mentions
  (e.g., State Bank of India, then SBI).
- Write numbers zero through nine alphabetically.
  Use numerals for 10 and above,
  EXCEPT for currency, percentages, time, dates, and ranges.
- Use the full name at first instance.
  Subsequently:
  - Use the surname only when accompanied by a title (Dr, Prof, Lt, Col).
  - If no title exists, use the first name
    (surname only for widely recognised senior figures).
- Avoid transitional words such as "meanwhile" and "moreover".
- Do NOT use em-dashes.
- Always apply the Oxford comma in lists of three or more items.
- Minimise repetitive phrasing or redundancy.
  Avoid repeated possessive pronouns in lists.
- Enclose titles of movies, songs, plays, books, and usernames in single quotes
  (e.g., 'DDLJ').
- Use double quotes exclusively for direct statements and verbatim quotations.
- Use lowercase "am" and "pm" for all time references.
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

    def call_gemini(prompt):
        return generate_text(
            prompt,
            generation_config={
                "temperature": 0,
                "top_p": 1,
                "top_k": 1,
                "candidate_count": 1
            },
            model_name=MODEL_FLASH,
        )

    # Parallel batch calls for speed (same logic/output)
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = []
        for para in paragraphs:
            prompt = BASE_PROMPT + "\n\nTEXT:\n" + para
            futures.append(executor.submit(call_gemini, prompt))

        for future in as_completed(futures):
            try:
                responses.append(future.result())
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


# =================================================
# GEMINI EDITORIAL QC — GUIDELINES
# =================================================
def gemini_editorial_review(article_data, web_story=False, health_beauty=False):
    init_vertex_and_model()

    paragraphs = [
        text[:900]
        for ctype, text in article_data
        if ctype == "paragraph"
    ]

    joined_paragraphs = "\n\n".join(paragraphs)

    context_notes = []
    if web_story:
        context_notes.append("- This is a web story: no hyperlinks are allowed.")
    if health_beauty:
        context_notes.append(
            "- This is a health/beauty story: include at least one expert quote."
        )
    extra_context = "\n".join(context_notes)

    prompt = f"""
You are an editorial QC reviewer. Identify only clear violations and keep fixes concise.

Rules:
1) British English spellings only (e.g., organise). Keep official names as-is.
2) No ampersands unless part of an official title.
3) Single quotes for titles; double quotes only for direct speech or press releases.
4) Use Oxford comma where needed for clarity.
5) Units: lowercase and singular (kg, km, kmph/km/h). No "kgs"/"kms".
6) Acronyms without periods or spaces (US, not U.S.).
7) No external hyperlinks in the first two paragraphs. Web stories: no hyperlinks at all.
8) Avoid superlatives/hyperbole and cliches. Avoid redundancy.
9) Seasons: use singular terms; "rains" only for the entire season.
10) Honorifics: no Mr/Mrs/Ms. Dr/Prof acceptable for doctors/professors.
11) Honours as suffixes (e.g., "Bharat Ratna awardee [Name]").
12) Age format: "[Name], 38," or "38-year-old [Name]".
13) Use man/woman for 18+, boy/girl for 17 and under; minor/juvenile only in legal context.
14) Health/beauty stories must include an expert quote.
15) Provide due credit for data and avoid plagiarism.
16) AI tools may assist but must not be the primary writing source.

Context:
{extra_context if extra_context else "- No extra context."}

Return output strictly as a markdown table:
| Issue | Location | Suggestion |

If there are no issues, return exactly one row:
| No issues found | - | - |

TEXT:
{joined_paragraphs}
"""

    return generate_text(
        prompt,
        generation_config={
            "temperature": 0,
            "top_p": 1,
            "top_k": 1,
            "candidate_count": 1
        },
        model_name=MODEL_FLASH,
    )


# =================================================
# CACHED GEMINI WRAPPERS (STABLE OUTPUTS)
# =================================================
@st.cache_data(show_spinner=False)
def cached_gemini_grammar_review(article_data, max_chars=GRAMMAR_MAX_CHARS):
    return gemini_grammar_review(article_data, max_chars)


@st.cache_data(show_spinner=False)
def cached_gemini_editorial_review(article_data, web_story=False, health_beauty=False):
    return gemini_editorial_review(article_data, web_story, health_beauty)


@st.cache_data(show_spinner=False)
def cached_gemini_fact_check(article_data, max_chars=FACT_MAX_CHARS, max_items=FACT_MAX_ITEMS):
    return gemini_fact_check(article_data, max_chars, max_items)


# ============================
# No-op correction filters
# ============================
def normalize_for_equality(text: str) -> str:
    text = html.unescape((text or "").strip())
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def is_noop_reason(reason: str) -> bool:
    normalised = normalize_for_equality(reason).lower()
    return normalised in {
        "",
        "no correction needed",
        "no corrections needed",
        "no change needed",
        "no changes needed",
        "already correct",
        "correct as is",
        "no issue",
        "no issues",
        "no error",
        "no errors",
    }


def is_noop_correction(original: str, corrected: str, reason: str = "") -> bool:
    if is_noop_reason(reason):
        return True
    return normalize_for_equality(original) == normalize_for_equality(corrected)


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

        if is_noop_correction(original, corrected, reason):
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

        if is_noop_correction(original, corrected, reason):
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

    def sort_key(row):
        cols = [c.strip() for c in row.split("|") if c.strip()]
        if len(cols) != 3:
            return ("", "", "")
        return tuple(re.sub(r"\W+", "", c.lower()) for c in cols)

    spelling_rows.sort(key=sort_key)
    grammar_rows.sort(key=sort_key)

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
# FACT CHECK — STATEMENT EXTRACTION (DETERMINISTIC)
# =================================================
def extract_fact_statements(article_data):
    """
    Deterministically extract candidate factual statements.
    SAME input → SAME statements → EVERY iteration.
    """
    statements = []
    seen = set()

    for ctype, text in article_data:
        if ctype != "paragraph":
            continue

        doc = nlp(text)
        for sent in doc.sents:
            s = sent.text.strip()

            # Basic factual heuristic (NO hard stops, NO assumptions)
            if len(s.split()) < 6:
                continue

            if not re.search(
                r"\b(is|was|are|were|has|have|had|will|announced|launched|reported|said|claims)\b",
                s.lower()
            ):
                continue

            # Canonical signature → stability
            key = re.sub(r"\s+", " ", s.lower())
            if key in seen:
                continue

            seen.add(key)
            statements.append(s)

    return statements

# =============
# Def Chunked
# =============

def chunked(lst, size):
    for i in range(0, len(lst), size):
        yield lst[i:i + size]

# =================================================
# FACT CHECK — SECOND PASS (FAST, STREAMING, STABLE)
# =================================================
def gemini_fact_check(article_data, max_chars=FACT_MAX_CHARS, max_items=FACT_MAX_ITEMS):
    init_vertex_and_model()

    # 1️⃣ Deterministic statement universe
    statements = extract_fact_statements(article_data)
    if not statements:
        return ""

    # Full article text (verbatim, unchanged)
    full_text = "\n".join(
        text for ctype, text in article_data if ctype == "paragraph"
    )

    batches = _batch_statements(statements, max_chars, max_items)

    rows = []
    seen = set()

    def call_batch(batch):
        batch_block = "\n".join(f"- {stmt}" for stmt in batch)

        fact_prompt = f"""
You are a factual accuracy reviewer.

SCOPE:
- Use general world knowledge to assess factual accuracy
- Only evaluate statements that appear verbatim in the TEXT
- Quote the EXACT sentence fragment under "Statement"
- Do NOT paraphrase, rewrite, or infer

EVALUATION RULES:
- If a statement is likely false, mark Issue as "Likely false" and provide the correct fact
- If a statement is uncertain or time-sensitive, mark Issue as "Needs verification"
- If a statement is likely true, omit it (do NOT create a row)
- NEVER invent facts; if unsure, use "Needs verification"

DO NOT:
- Check grammar, spelling, or style
- Rewrite sentences

Return output strictly as a table:
| Statement | Issue | Correct Fact |

TEXT:
{full_text}

        STATEMENTS:
{batch_block}
"""

        try:
            return generate_stream_text(
                fact_prompt,
                generation_config={
                    "temperature": 0,
                    "top_p": 1,
                    "top_k": 1,
                    "candidate_count": 1
                },
                model_name=MODEL_FLASH,
            )
        except Exception:
            return ""

    # 2️⃣ Batched + streaming Gemini calls (parallel)
    batch_results = []
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [executor.submit(call_batch, batch) for batch in batches]
        for future in as_completed(futures):
            try:
                batch_results.append(future.result())
            except Exception:
                continue

    for out in batch_results:
        if not out:
            continue

        # 3️⃣ Extract table rows (UNCHANGED)
        matches = re.findall(
            r"\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|",
            out
        )

        for s, issue, correction in matches:
            if s.lower() == "statement":
                continue

            sig = (
                re.sub(r"\W+", "", s.lower()),
                re.sub(r"\W+", "", issue.lower())
            )

            if sig in seen:
                continue

            seen.add(sig)
            rows.append((s.strip(), issue.strip(), correction.strip()))

    if not rows:
        return ""

    def sort_key(row):
        return (
            re.sub(r"\W+", "", row[0].lower()),
            re.sub(r"\W+", "", row[1].lower()),
            re.sub(r"\W+", "", row[2].lower()),
        )

    rows.sort(key=sort_key)

    rendered = [
        f"| {s} | {issue} | {correction} |"
        for s, issue, correction in rows
    ]

    return "\n".join([
        "| Statement | Issue | Correct Fact |",
        "|---|---|---|",
        *rendered
    ])


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

st.sidebar.header("QC Options")
web_story = st.sidebar.checkbox("Web story (no hyperlinks)", value=False)
health_beauty = st.sidebar.checkbox(
    "Health/Beauty story (expert quote required)", value=False
)
run_gemini_editorial = st.sidebar.checkbox(
    "Run Gemini editorial QC", value=True
)
if st.sidebar.button("Clear cached AI outputs"):
    st.cache_data.clear()
    for key in (
        "analysis_results",
        "analysis_key",
        "analysis_start",
        "article_content",
        "input_key",
    ):
        st.session_state.pop(key, None)

analyze_clicked = st.sidebar.button("Analyze")

article_content = None
current_key = None

if source == "URL":
    url = st.sidebar.text_input("Article URL")
    if url:
        current_key = f"url:{url.strip()}"
    if analyze_clicked and url:
        article_content = clean_article(url)
        st.session_state["article_content"] = article_content
        st.session_state["input_key"] = current_key
else:
    uploaded = st.sidebar.file_uploader("Upload DOCX", type=["docx"])
    if uploaded:
        file_bytes = uploaded.getvalue()
        current_key = "docx:" + hashlib.sha256(file_bytes).hexdigest()
        if analyze_clicked:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
                f.write(file_bytes)
                article_content = clean_docx(f.name)
            st.session_state["article_content"] = article_content
            st.session_state["input_key"] = current_key

if article_content is None:
    if current_key and st.session_state.get("input_key") == current_key:
        article_content = st.session_state.get("article_content")

analysis_key = None
if current_key:
    analysis_key = hashlib.sha256(
        f"{current_key}|{web_story}|{health_beauty}|{run_gemini_editorial}".encode("utf-8")
    ).hexdigest()

analysis_ready = False
if article_content and analysis_key:
    if analyze_clicked:
        if st.session_state.get("analysis_key") != analysis_key:
            st.session_state["analysis_key"] = analysis_key
            st.session_state["analysis_results"] = {}
            st.session_state["analysis_start"] = time.perf_counter()
        else:
            st.info("Using cached results. Clear cached AI outputs to refresh.")
    if st.session_state.get("analysis_key") == analysis_key:
        analysis_ready = True
    else:
        st.info("Input or options changed. Click Analyze to refresh.")

if analysis_ready:
    qc_content = run_pipeline(article_content)
    results = st.session_state.setdefault("analysis_results", {})
    if "analysis_start" not in results:
        results["analysis_start"] = st.session_state.get("analysis_start", time.perf_counter())

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
    # Grammar + Spelling placeholders
    st.markdown("### ✍️ Spelling Issues")
    spelling_placeholder = st.empty()
    st.markdown("### 🧠 Grammar Issues")
    grammar_placeholder = st.empty()

    editorial_placeholder = None
    if run_gemini_editorial:
        st.markdown("### 🧠 Gemini Editorial Review")
        editorial_placeholder = st.empty()

    # ---------- FACT CHECK ----------
    st.markdown("### 📌 Fact Check")
    fact_placeholder = st.empty()

    def render_grammar(raw_text):
        clean = filter_invalid_rows(raw_text, article_text)
        spelling_table, grammar_table = split_spelling_grammar(clean)
        if spelling_table:
            spelling_placeholder.markdown(spelling_table)
        else:
            spelling_placeholder.success("✅ No spelling issues found")
        if grammar_table:
            grammar_placeholder.markdown(grammar_table)
        else:
            grammar_placeholder.success("✅ No grammar issues found")

    def render_fact(fact_text):
        if not fact_text or "| Statement |" not in fact_text:
            fact_placeholder.success("✅ No factual issues found")
        else:
            fact_placeholder.markdown(fact_text)

    # Show cached results if present
    if "grammar_raw" in results:
        render_grammar(results["grammar_raw"])
    else:
        spelling_placeholder.info("Running Gemini grammar/spelling...")
        grammar_placeholder.info("Running Gemini grammar/spelling...")

    if run_gemini_editorial:
        if "gemini_editorial" in results:
            editorial_placeholder.markdown(results["gemini_editorial"])
        else:
            editorial_placeholder.info("Running Gemini editorial QC...")

    if "fact_result" in results:
        render_fact(results["fact_result"])
    else:
        fact_placeholder.info("Running Gemini fact check...")

    # Run missing AI tasks in parallel
    tasks = {}
    with ThreadPoolExecutor(max_workers=3) as executor:
        if "grammar_raw" not in results:
            tasks[executor.submit(
                cached_gemini_grammar_review,
                qc_content,
                GRAMMAR_MAX_CHARS
            )] = "grammar"
        if run_gemini_editorial and "gemini_editorial" not in results:
            tasks[executor.submit(
                cached_gemini_editorial_review,
                qc_content,
                web_story,
                health_beauty
            )] = "editorial"
        if "fact_result" not in results:
            tasks[executor.submit(
                cached_gemini_fact_check,
                qc_content,
                FACT_MAX_CHARS,
                FACT_MAX_ITEMS
            )] = "fact"

        for future in as_completed(tasks):
            key = tasks[future]
            try:
                result = future.result()
            except Exception as exc:
                result = f"Error: {exc}"

            if key == "grammar":
                results["grammar_raw"] = result
                render_grammar(result)
            elif key == "editorial":
                results["gemini_editorial"] = result
                if editorial_placeholder:
                    editorial_placeholder.markdown(result)
            elif key == "fact":
                results["fact_result"] = result
                render_fact(result)

    required_keys = {"grammar_raw", "fact_result"}
    if run_gemini_editorial:
        required_keys.add("gemini_editorial")

    if "elapsed" not in results and required_keys.issubset(results.keys()):
        results["elapsed"] = time.perf_counter() - results["analysis_start"]

    elapsed = results.get("elapsed")
    if elapsed is not None:
        minutes = int(elapsed // 60)
        seconds = int(round(elapsed % 60))
        st.caption(f"⏱️ Time taken: {minutes}m {seconds:02d}s")
        st.sidebar.metric("Time taken", f"{minutes}m {seconds:02d}s")
