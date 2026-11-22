import io
import json
import random
import re

import pandas as pd
import streamlit as st
from groq import Groq
from pypdf import PdfReader
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pptx import Presentation
from openpyxl import load_workbook
from sklearn.feature_extraction.text import TfidfVectorizer

# ======================= অ্যাপের নাম (এখানে বদলালেই সব জায়গায় বদলাবে) =======================
APP_NAME = "Shikkha.ai"
TAGLINE = "Teach smarter. Learn smarter. In Bangla."
VERSION = "v6"

st.set_page_config(page_title=APP_NAME, page_icon="📚", layout="centered")
client = Groq(api_key=st.secrets["GROQ_API_KEY"])   # key থাকে .streamlit/secrets.toml এ, কোডে না

@st.cache_data
def pick_model():
    try:
        ids = [m.id for m in client.models.list().data]
    except Exception:
        return "openai/gpt-oss-120b"
    chat = [m for m in ids if not any(x in m for x in ["whisper", "guard", "tts"])]
    if "openai/gpt-oss-120b" in chat:
        return "openai/gpt-oss-120b"
    best = [m for m in chat if any(x in m for x in ["llama", "gpt-oss", "qwen"])]
    return (best or chat or ["openai/gpt-oss-120b"])[0]
MODEL = pick_model()

def brand_html(small=""):
    """'Shikkha.ai' -> 'Shikkha' সাধারণ রঙে, '.ai' রঙিন"""
    name = APP_NAME
    if "." in name:
        a, b = name.rsplit(".", 1)
        inner = f"{a}<span>.{b}</span>"
    elif " " in name:
        a, b = name.rsplit(" ", 1)
        inner = f"{a} <span>{b}</span>"
    else:
        inner = f"<span>{name}</span>"
    return f'<div class="brand">{inner}{f"<small>{small}</small>" if small else ""}</div>'

# ======================= ডিজাইন (CSS) =======================
# রং কম, সাদা জায়গা বেশি। আলো/অন্ধকার (dark mode) দুই থিমেই কাজ করে।
css = """
@import url('https://fonts.googleapis.com/css2?family=Anek+Bangla:wght@500;600;700&family=Noto+Sans+Bengali:wght@400;500;600&family=Inter:wght@400;500;600&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');

:root {--ac: #6366f1; --ac2: #8b5cf6; --soft: rgba(99,102,241,.10); --line: rgba(128,128,128,.22);}
/* লেখার font: English এ Inter, বাংলায় Noto Sans Bengali (পরিষ্কার, পড়তে আরাম) */
.stApp, .stApp p, .stApp li, .stApp label, .stApp input, .stApp textarea, .stApp button,
[data-testid="stMarkdownContainer"], [data-baseweb="select"] div, [data-testid="stSidebar"] {
  font-family: 'Inter', 'Noto Sans Bengali', sans-serif;}
/* শিরোনামের font: Plus Jakarta Sans + Anek Bangla (modern, একটু মোটা) */
h1, h2, h3, h4, .brand, .hello h1, .paper-head h3, .step b, .feat b, .score b {
  font-family: 'Plus Jakarta Sans', 'Anek Bangla', sans-serif !important;}
.block-container {padding-top: 3.2rem; padding-bottom: 7rem;}
.stAppDeployButton, [data-testid="stDecoration"] {display: none;}
@keyframes rise {from {opacity: 0; transform: translateY(6px);} to {opacity: 1; transform: none;}}
.hello, .feat, .step, .paper-head, .score {animation: rise .35s ease-out both;}

/* উপরের সরু header */
.brand {font-size: 22px; font-weight: 800; letter-spacing: -.5px; white-space: nowrap;}
.brand span {background: linear-gradient(90deg, var(--ac), var(--ac2)); -webkit-background-clip: text;
  background-clip: text; color: transparent;}
.brand small {font-family: 'Inter', 'Noto Sans Bengali', sans-serif; font-size: 12.5px; font-weight: 500;
  opacity: .55; margin-left: 8px; letter-spacing: 0;}
[data-testid="stButtonGroup"] button {border-radius: 10px !important;}
[data-testid="stButtonGroup"] button p {font-weight: 500;}

/* খালি অবস্থার শুভেচ্ছা */
.hello {text-align: center; margin: 6vh 0 22px;}
.hello h1 {font-size: 36px; font-weight: 700; letter-spacing: 0; word-spacing: .18em; margin: 0; padding: 0; line-height: 1.25;}
.hello p {opacity: .62; font-size: 16px; margin: 10px 0 0;}
.hello .tag {display: inline-block; font-size: 12.5px; margin-bottom: 14px; padding: 5px 13px;
  border-radius: 999px; background: var(--soft); color: var(--ac); font-weight: 600;}

/* হোমের ৩টা কার্ড */
.feat {padding: 4px 2px 2px;}
.feat i {font-style: normal; display: inline-flex; width: 38px; height: 38px; border-radius: 11px; align-items: center;
  justify-content: center; font-size: 19px; background: var(--soft); margin-bottom: 10px;}
.feat b {display: block; font-size: 17px; font-weight: 700; margin-bottom: 3px;}
.feat small {display: none;}
.feat p {font-size: 13.5px; opacity: .65; margin: 0 0 4px; line-height: 1.55; min-height: 42px;}
.st-key-feats [data-testid="stVerticalBlockBorderWrapper"] {transition: border-color .15s, transform .15s, box-shadow .15s;}
.st-key-feats [data-testid="stVerticalBlockBorderWrapper"]:hover {border-color: var(--ac); transform: translateY(-2px);
  box-shadow: 0 8px 24px rgba(99,102,241,.12);}
.howto {display: flex; justify-content: center; gap: 22px; flex-wrap: wrap; font-size: 13.5px; opacity: .7; margin: 18px 0 8px;}
.howto span b {display: inline-flex; width: 20px; height: 20px; border-radius: 50%; align-items: center; justify-content: center;
  background: var(--ac); color: #fff; font-size: 11.5px; margin-right: 6px;}

/* ধাপ (১, ২, ৩...) */
.step {display: flex; gap: 12px; align-items: flex-start; margin: 22px 0 8px;}
.step > span {flex: none; width: 28px; height: 28px; border-radius: 9px; display: inline-flex; align-items: center;
  justify-content: center; font-weight: 700; font-size: 14px; color: #fff;
  background: linear-gradient(135deg, var(--ac), var(--ac2));}
.step b {display: block; font-size: 16.5px; font-weight: 700; line-height: 1.3;}
.step small {display: block; font-size: 13px; opacity: .6; margin-top: 2px; line-height: 1.5;}
.tip {font-size: 13.5px; padding: 10px 14px; border-radius: 12px; background: var(--soft); margin: 6px 0 10px; line-height: 1.6;}
.tip b {color: var(--ac);}

/* বোতাম: হালকা border, বেশি গোল */
.stButton button[kind="secondary"], .stDownloadButton button, [data-testid="stPopover"] button {border-radius: 12px;
  border: 1px solid var(--line); background: transparent; font-weight: 500; padding: 8px 14px;
  transition: background .15s, border-color .15s, transform .1s;}
.stButton button[kind="secondary"]:hover, .stDownloadButton button:hover {border-color: var(--ac);
  background: rgba(99,102,241,.06); color: inherit;}
.stButton button[kind="primary"], [data-testid="stFormSubmitButton"] button {border-radius: 12px; font-weight: 600; border: none;
  background: linear-gradient(90deg, var(--ac), var(--ac2)); color: #fff; box-shadow: 0 4px 14px rgba(99,102,241,.25);
  transition: transform .1s, box-shadow .15s;}
.stButton button[kind="primary"]:hover, [data-testid="stFormSubmitButton"] button:hover {transform: translateY(-1px);
  box-shadow: 0 6px 18px rgba(99,102,241,.35); color: #fff;}
.stButton button:active {transform: scale(.98);}
.st-key-suggest button {justify-content: flex-start; text-align: left; padding: 14px 16px !important;}

/* chat */
[data-testid="stChatMessage"] {background: transparent; padding: 6px 0; gap: 14px;}
[data-testid="stChatMessage"]:has([aria-label="Chat message from user"]) {flex-direction: row-reverse;}
[data-testid="stChatMessage"]:has([aria-label="Chat message from user"]) [data-testid^="stChatMessageAvatar"] {display: none;}
[aria-label="Chat message from user"] {flex: 0 1 auto !important; max-width: 85%;
  background: rgba(128,128,128,.12); border-radius: 18px; padding: 8px 16px;}
[data-testid^="stChatMessageAvatar"] {border: none; background: var(--soft); color: var(--ac);}
[data-testid="stChatMessageContent"] p, [data-testid="stChatMessageContent"] li {font-size: 16px; line-height: 1.8;}
[data-testid="stChatMessageContent"] h1, [data-testid="stChatMessageContent"] h2 {font-size: 20px; font-weight: 700;}
[data-testid="stChatMessageContent"] h3 {font-size: 17px; font-weight: 700;}
[data-testid="stChatInput"] textarea {font-size: 16px;}
[data-testid="stChatInput"] > div {border-radius: 18px;}

/* উৎস আর ফাইলের ছোট pill */
.src {margin: 6px 0 0; display: flex; flex-wrap: wrap; gap: 6px;}
.src span {font-size: 12.5px; padding: 3px 10px; border-radius: 999px; background: var(--soft); color: var(--ac); font-weight: 500;}
.st-key-pills button {padding: 2px 12px !important; min-height: 30px !important; border-radius: 999px !important;
  background: rgba(128,128,128,.08) !important;}
.st-key-pills button p {font-size: 13px !important;}
.meta {font-size: 12.5px; opacity: .55; margin: 2px 0 4px;}

/* ফর্ম, কার্ড */
[data-testid="stForm"] {border: 1px solid var(--line); border-radius: 16px; padding: 22px;}
[data-testid="stForm"] [data-testid="stRadio"] {padding: 10px 0 14px; border-bottom: 1px solid rgba(128,128,128,.15);}
[data-testid="stVerticalBlockBorderWrapper"] {border-radius: 16px;}
[data-testid="stExpander"] details {border-radius: 12px;}
[data-testid="stFileUploaderDropzone"] {border-radius: 14px; border: 1.5px dashed rgba(99,102,241,.45); background: rgba(99,102,241,.04);}
.score {text-align: center; padding: 18px 0 6px;}
.score b {font-size: 46px; font-weight: 800; letter-spacing: -1px;}
.score p {opacity: .65; margin: 0;}

/* প্রশ্নপত্র */
.paper-head {text-align: center; padding: 10px 0 12px; border-bottom: 1px solid var(--line); margin: 18px 0 12px;}
.paper-head h3 {font-size: 21px; font-weight: 700; margin: 0; padding: 0;}
.paper-head p {font-size: 13.5px; opacity: .65; margin: 4px 0 0;}
.sec-head {font-weight: 700; font-size: 15.5px; margin: 16px 0 2px;}
.sec-note {font-size: 13px; opacity: .6; margin: 0 0 6px;}
.total {display: inline-block; font-size: 13.5px; margin: 2px 0 8px; padding: 5px 12px; border-radius: 999px;
  background: var(--soft);}
.total b {color: var(--ac); font-size: 15px;}

/* sidebar */
[data-testid="stSidebar"] {border-right: 1px solid rgba(128,128,128,.15);}
.side-tag {font-size: 12px; opacity: .6; white-space: nowrap; margin: -6px 0 14px; letter-spacing: -.1px;}
.side-foot {font-size: 12.5px; opacity: .6; line-height: 1.7; margin-top: 8px;}
.side-foot a {color: inherit;}

@media (max-width: 640px) {
  .block-container {padding-top: 3.6rem;}
  .hello {margin: 0 0 12px;}
  .hello h1 {font-size: 26px;}
  .hello p {font-size: 14px;}
  .brand small {display: none;}
  .st-key-topbar .brand {line-height: 1.5; margin-bottom: 8px;}
  /* ফোনে ৩টা কার্ড পাশাপাশি ছোট করে */
  .st-key-feats [data-testid="stHorizontalBlock"] {flex-wrap: nowrap; gap: 8px;}
  .st-key-feats [data-testid="stColumn"] {min-width: 0 !important; flex: 1 1 33% !important; width: 33% !important;}
  .st-key-feats [data-testid="stVerticalBlockBorderWrapper"] {padding: 10px 8px !important;}
  .feat {text-align: center;}
  .feat p {display: none;}
  .feat small {display: block; font-size: 12px; opacity: .65; margin: 0 0 6px; line-height: 1.4;}
  .st-key-homeup, .hello .tag, .hello.home p {display: none;}
  .hello h1 {font-size: 24px !important;}
  .howto {margin: 8px 0 0; font-size: 12.5px; gap: 6px 12px;}
  .howto span b {width: 17px; height: 17px; font-size: 10px;}
  .hello h1 {font-size: 22px !important; word-spacing: .12em;}
  .feat i {width: 32px; height: 32px; font-size: 16px; margin-bottom: 4px;}
  .feat b {font-size: 15px;}
  .st-key-feats button {padding: 4px 6px !important; min-height: 32px;}
  .st-key-feats button p {font-size: 12px !important;}
  .howto {gap: 10px 16px;}
  .step {margin-top: 18px;}
  /* ফোনেও পরামর্শের বোতাম ২টা করে পাশাপাশি */
  .st-key-suggest [data-testid="stHorizontalBlock"] {flex-wrap: nowrap; gap: 8px;}
  .st-key-suggest [data-testid="stColumn"] {min-width: 0 !important; flex: 1 1 50% !important; width: 50% !important;}
  .st-key-suggest button {padding: 10px 10px !important;}
  .st-key-suggest button p {font-size: 14px;}
  .st-key-topbar [data-testid="stButtonGroup"] button p {font-size: 13.5px;}
}
"""
tag = "style"
st.markdown(f"<{tag}>{css}</{tag}>", unsafe_allow_html=True)

# ======================= অবস্থা (session state) =======================
TABS = ["চ্যাট", "অনুশীলন", "প্রশ্নপত্র"]
for key, value in {"chat": [], "quiz": None, "quiz_done": False, "quiz_id": 0, "papers": {},
                   "files": {}, "tab": "চ্যাট", "last_tab": "চ্যাট", "topic_stats": {}, "recorded": set(),
                   "viva": {"current": None, "done": []}, "past": None, "up_id": 0}.items():
    if key not in st.session_state:
        st.session_state[key] = value

TYPES = ["pdf", "docx", "pptx", "xlsx", "txt"]

# ======================= Sidebar (সেটিং ও পরিচয়) =======================
with st.sidebar:
    st.markdown(brand_html(), unsafe_allow_html=True)
    st.markdown(f'<div class="side-tag">{TAGLINE}</div>', unsafe_allow_html=True)
    if st.button("＋ নতুন chat", width="stretch"):
        st.session_state.chat = []
        st.rerun()
    st.write("")
    lang = st.radio("উত্তরের ভাষা", ["বাংলা", "English"], horizontal=True)
    TOP_K = st.slider("প্রতি প্রশ্নে কয়টা অংশ খুঁজবে", 3, 8, 5,
                      help="RAG: ফাইলের সবচেয়ে মিল থাকা এতগুলো টুকরো দেখে উত্তর লেখা হয়।")
    download_slot = st.empty()
    st.markdown(f'<div class="side-foot">Model: {MODEL}<br>Search: TF-IDF RAG<br><br>'
                f'Made by <a href="https://github.com/sanzida-siddike" target="_blank">Sanzida Siddike</a>'
                f' · {VERSION}</div>', unsafe_allow_html=True)

# ================================================================
#  ধাপ ১: ফাইল পড়া  ->  (জায়গার নাম, লেখা) এর তালিকা
# ================================================================
def bn(n):
    """ইংরেজি সংখ্যাকে বাংলা সংখ্যা বানায়: 12 -> ১২"""
    return str(n).translate(str.maketrans("0123456789", "০১২৩৪৫৬৭৮৯"))

def group_lines(lines, prefix, size=900):
    """ছোট ছোট লাইন/অনুচ্ছেদ একসাথে জুড়ে ~900 অক্ষরের ভাগ বানায়, নাম দেয় 'অনুচ্ছেদ ৩-৭'"""
    units, buf = [], []

    def flush():
        a, b = buf[0][0], buf[-1][0]
        units.append((f"{prefix} {bn(a)}" + (f"-{bn(b)}" if b != a else ""), "\n".join(t for _, t in buf)))

    for no, line in lines:
        if buf and sum(len(t) for _, t in buf) + len(line) > size:
            flush()
            buf = []
        buf.append((no, line))
    if buf:
        flush()
    return units

def read_file(data, name):
    """ফাইল থেকে লেখা বের করে: (units, সংখ্যা, কী গোনা হলো)"""
    name = name.lower()
    f = io.BytesIO(data)
    if name.endswith(".pdf"):
        pages = PdfReader(f).pages
        units = [(f"পাতা {bn(i)}", p.extract_text() or "") for i, p in enumerate(pages, 1)]
        return units, len(pages), "পাতা"
    if name.endswith(".docx"):
        doc = Document(f)
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                paras.append(" | ".join(c.text for c in row.cells))
        return group_lines(list(enumerate(paras, 1)), "অনুচ্ছেদ"), len(paras), "অংশ"
    if name.endswith(".pptx"):
        prs = Presentation(f)
        units = []
        for i, slide in enumerate(prs.slides, 1):
            parts = [s.text_frame.text for s in slide.shapes
                     if s.has_text_frame and s.text_frame.text.strip()]
            units.append((f"Slide {bn(i)}", "\n".join(parts)))
        return units, len(prs.slides), "Slide"
    if name.endswith(".xlsx"):
        wb = load_workbook(f, read_only=True, data_only=True)
        units = []
        for ws in wb.worksheets:
            rows = []
            for r, row in enumerate(ws.iter_rows(values_only=True), 1):
                cells = [str(c) for c in row if c is not None and str(c).strip()]
                if cells:
                    rows.append((r, " | ".join(cells)))
            if rows:
                units += group_lines(rows, f"Sheet '{ws.title}' · সারি")
        return units, len(wb.worksheets), "অংশ"
    try:
        txt = data.decode("utf-8")
    except UnicodeDecodeError:
        txt = data.decode("utf-16", errors="ignore")
    lines = [(i, l) for i, l in enumerate(txt.splitlines(), 1) if l.strip()]
    return group_lines(lines, "লাইন"), len(txt.splitlines()), "অংশ"

# ================================================================
#  ধাপ ২: Chunking  ->  বড় লেখাকে ছোট ছোট টুকরো (chunk) করা
# ================================================================
CHUNK_SIZE = 900   # একটা টুকরোতে মোটামুটি কত অক্ষর
OVERLAP = 150      # পাশের টুকরোর সাথে কতটুকু মিল থাকবে

def split_text(text):
    """বাক্য (। . ? !) ধরে লেখাকে ~CHUNK_SIZE অক্ষরের টুকরো করে, একটু overlap রেখে"""
    text = re.sub(r"[ \t]+", " ", text).strip()
    if len(text) <= CHUNK_SIZE:
        return [text] if text else []
    sentences = re.split(r"(?<=[।.!?\n])\s+", text)
    chunks, cur = [], ""
    for s in sentences:
        while len(s) > CHUNK_SIZE:
            chunks.append((cur + " " + s[:CHUNK_SIZE]).strip())
            s, cur = s[CHUNK_SIZE - OVERLAP:], ""
        if len(cur) + len(s) > CHUNK_SIZE and cur:
            chunks.append(cur.strip())
            cur = cur[-OVERLAP:]
        cur += " " + s
    if cur.strip():
        chunks.append(cur.strip())
    return chunks

def make_chunks(units, kind):
    """[(জায়গা, লেখা)] -> [{"src": "পাতা ১২", "text": "...", "n": 12, "kind": "পাতা"}]
    n = কত নম্বর পাতা/Slide/অংশ, যাতে পরে 'পাতা ৫ থেকে ২০' বাছাই করা যায়"""
    out = []
    for n, (label, text) in enumerate(units, 1):
        for piece in split_text(text):
            if len(piece) > 20:
                out.append({"src": label, "text": piece, "n": n, "kind": kind})
    return out

@st.cache_data(show_spinner=False, max_entries=8)
def load_document(data, name):
    """ফাইল পড়ে chunk বানায়। একই ফাইলের জন্য বারবার কাজ করে না (cache)।"""
    units, count, kind = read_file(data, name)
    full = "\n".join(t for _, t in units)
    return make_chunks(units, kind), full, count, kind

# ================================================================
#  ধাপ ৩: Index ও Retrieval (শব্দ + অক্ষর ভিত্তিক TF-IDF)
# ================================================================
def normalize(text):
    text = text.lower().replace("‌", "").replace("‍", "")
    return re.sub(r"[^ঀ-৿a-z0-9]+", " ", text)

def tokenize(text):
    return re.findall(r"[ঀ-৿]+|[a-z0-9]+", normalize(text))

@st.cache_resource(show_spinner=False, max_entries=8)
def build_index(texts):
    word_vec = TfidfVectorizer(tokenizer=tokenize, token_pattern=None, lowercase=False, sublinear_tf=True)
    char_vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), preprocessor=normalize,
                               sublinear_tf=True)
    return word_vec, word_vec.fit_transform(texts), char_vec, char_vec.fit_transform(texts)

def retrieve(query, chunks, k):
    """প্রশ্নের সাথে মিল অনুযায়ী সেরা k টা টুকরো ফেরত দেয় (score সহ)"""
    word_vec, W, char_vec, C = build_index(tuple(c["text"] for c in chunks))
    word_score = (W @ word_vec.transform([query]).T).toarray().ravel()
    char_score = (C @ char_vec.transform([query]).T).toarray().ravel()
    score = 0.5 * word_score + 0.5 * char_score
    best = score.argsort()[::-1][:k]
    return [(int(i), float(score[i])) for i in best]

def overview(chunks, budget=12000, shuffle=False):
    """সারাংশ/Quiz এর জন্য: পুরো ফাইলের শুরু-মাঝ-শেষ থেকে সমান ভাবে টুকরো নেয়"""
    n = max(1, min(len(chunks), budget // CHUNK_SIZE))
    if shuffle:
        idx = sorted(random.sample(range(len(chunks)), n))
    else:
        idx = sorted({round(i * (len(chunks) - 1) / max(1, n - 1)) for i in range(n)})
    return idx

def context_text(chunks, idx):
    return "\n\n".join(f"[{j}] ({chunks[i]['src']})\n{chunks[i]['text']}" for j, i in enumerate(idx, 1))

def src_of(chunks, idx, part):
    return chunks[idx[part - 1]]["src"] if isinstance(part, int) and 1 <= part <= len(idx) else ""

# ================================================================
#  ধাপ ৪: Generation  ->  খুঁজে পাওয়া টুকরো দিয়ে AI উত্তর লেখে
# ================================================================
def lang_rule():
    return "Answer in simple, clear English." if lang == "English" else "সহজ বাংলায় উত্তর দাও।"

def write_rule():
    return ("Write everything in English." if lang == "English"
            else "বাংলায় লেখো, technical শব্দ English এ রাখতে পারো।")

def llm(prompt, json_mode=False):
    extra = {"response_format": {"type": "json_object"}} if json_mode else {}
    r = client.chat.completions.create(model=MODEL, messages=[{"role": "user", "content": prompt}], **extra)
    return r.choices[0].message.content or ""

def parse_json(raw):
    m = re.search(r"\{.*\}", raw, re.S)
    return json.loads(m.group(0) if m else raw)

def search_words(q):
    """প্রশ্নটাকে বাংলা + English দুই ভাষার keyword এ বদলায়"""
    prev = [m["text"] for m in st.session_state.chat if m["role"] == "user"][-1:]
    try:
        kw = llm("Give search keywords to find the answer to this question inside a document. "
                 "Write the important words in BOTH Bangla and English (translate them), add 2-3 synonyms. "
                 "Only one line of keywords, nothing else."
                 + (f"\nPrevious question (for context): {prev[0]}" if prev else "")
                 + f"\nQuestion: {q}")
        return f"{q} {kw.strip()}"[:600]
    except Exception:
        return q

def ask(q, chunks, mode):
    """mode = 'search' হলে RAG (খুঁজে উত্তর), 'overview' হলে পুরো ফাইল থেকে নমুনা"""
    if mode == "search":
        query = search_words(q)
        idx = [i for i, _ in retrieve(query, chunks, TOP_K)]
    else:
        query, idx = "", overview(chunks)
    prompt = (f"তুমি একজন দয়ালু শিক্ষক। নিচে একটা ফাইলের কিছু অংশ দেওয়া আছে, প্রতিটার আগে নম্বর [1], [2]... "
              f"আর জায়গা (পাতা/Slide) লেখা। শুধু এই অংশগুলো থেকে heading আর bullet point দিয়ে সুন্দর করে উত্তর দাও। "
              f"যে তথ্য যে অংশ থেকে নিচ্ছ, তার পরে [1] এর মতো নম্বর দাও। "
              f"উত্তর এই অংশগুলোতে না থাকলে সোজা বলো যে ফাইলে এই তথ্য পাওয়া যায়নি, বানিয়ে কিছু বলবে না। "
              f"{lang_rule()}\n\nফাইলের অংশ:\n{context_text(chunks, idx)}\n\nপ্রশ্ন: {q}")
    ans = llm(prompt)
    used = {int(x) for x in re.findall(r"\[(\d+)\]", ans)}
    picked = [idx[j - 1] for j in sorted(used) if 1 <= j <= len(idx)] or idx
    sources = [{"src": chunks[i]["src"], "text": chunks[i]["text"]} for i in picked]
    return ans, sources, query

# ---------------- অনুশীলন: MCQ ----------------
def focus_idx(chunks, topics, per_topic=3):
    """দুর্বল বিষয়গুলোর সাথে মিল থাকা টুকরো RAG দিয়ে খুঁজে আনে"""
    idx = []
    for t in topics:
        idx += [i for i, _ in retrieve(t, chunks, per_topic) if i not in idx]
    return sorted(idx[:12])

def make_quiz(chunks, n, hard=False, focus=None):
    """MCQ বানায়। প্রতিটা প্রশ্নের সাথে ছোট একটা 'topic' নাম থাকে, দুর্বল বিষয় খুঁজতে কাজে লাগে।"""
    idx = focus_idx(chunks, focus) if focus else overview(chunks, budget=9000, shuffle=True)
    level = ("Make them hard, like a Bangladesh university admission test (ভর্তি পরীক্ষা): tricky options, "
             "applied thinking. " if hard else "")
    extra = f"Focus ONLY on these topics: {', '.join(focus)}. " if focus else ""
    prompt = (f"Make {n} multiple choice questions ONLY from these parts of a file. {level}{extra}"
              f"Write the questions and options in {'English' if lang == 'English' else 'Bangla (বাংলা)'}. "
              'Reply with JSON only, exactly like this: {"questions": [{"q": "question", '
              '"options": ["A", "B", "C", "D"], "answer": 0, "why": "short explanation", "part": 1, '
              '"topic": "2-4 word topic name"}]} '
              "where answer is the index (0-3) of the correct option, part is the [number] of the part "
              "the question comes from, and topic is a short name of the concept it tests (reuse the same "
              "topic name for questions on the same concept).\n\nParts:\n" + context_text(chunks, idx))
    data = parse_json(llm(prompt, json_mode=True))
    good = []
    for item in data.get("questions", []):
        opts = item.get("options", [])
        ans = item.get("answer")
        if item.get("q") and len(opts) >= 2 and isinstance(ans, int) and 0 <= ans < len(opts):
            good.append({"q": item["q"], "options": [str(o) for o in opts], "answer": ans,
                         "why": item.get("why", ""), "src": src_of(chunks, idx, item.get("part")),
                         "topic": str(item.get("topic") or "অন্যান্য").strip()})
    return good

# ---------------- অনুশীলন: Viva ----------------
def viva_question(chunks, asked):
    """ফাইলের এলোমেলো একটা অংশ থেকে একটা মৌখিক প্রশ্ন"""
    i = random.randrange(len(chunks))
    prompt = ("You are a friendly but strict viva (oral exam) examiner. From this part of the "
              "student's notes, ask ONE short viva question (conceptual, the kind teachers ask face to face). "
              f"Do not repeat these: {asked[-5:]}. "
              f"Write it in {'English' if lang == 'English' else 'Bangla (বাংলা)'}. "
              'Reply JSON only: {"q": "question"}'
              f"\n\nNotes part:\n{chunks[i]['text']}")
    return {"q": parse_json(llm(prompt, json_mode=True)).get("q", ""), "text": chunks[i]["text"],
            "src": chunks[i]["src"]}

def viva_grade(item, answer):
    """ছাত্রের উত্তর ১০ এর মধ্যে নম্বর দেয়, ছোট মন্তব্য আর সঠিক উত্তর সহ"""
    prompt = ("You are a viva examiner. Grade the student's answer out of 10 using ONLY these notes as the "
              "reference. Be fair: partial credit for partly correct answers. "
              f"Write feedback in {'English' if lang == 'English' else 'Bangla (বাংলা)'}. "
              'Reply JSON only: {"score": 7, "feedback": "1-2 sentences: what was good, what was missing", '
              '"model_answer": "a short ideal answer"}'
              f"\n\nNotes:\n{item['text']}\n\nQuestion: {item['q']}\nStudent answer: {answer}")
    data = parse_json(llm(prompt, json_mode=True))
    try:
        score = max(0, min(10, int(round(float(data.get("score", 0))))))
    except (TypeError, ValueError):
        score = 0
    return {"score": score, "feedback": data.get("feedback", ""), "model": data.get("model_answer", "")}

# ---------------- অনুশীলন: আগের প্রশ্ন বিশ্লেষণ ----------------
def analyze_past(chunks):
    """আগের বছরের প্রশ্নপত্র থেকে কোন বিষয় বারবার আসে আর এবার কী আসতে পারে"""
    idx = overview(chunks, budget=14000)
    prompt = ("These are previous years' exam question papers. Find which topics come again "
              "and again. Count how many questions (across all papers) are on each topic. Then suggest the most "
              "likely questions for the next exam, based on repetition and importance. "
              f"Write in {'English' if lang == 'English' else 'Bangla (বাংলা), technical words can stay in English'}. "
              'Reply JSON only: {"topics": [{"topic": "name", "count": 5, "note": "e.g. came in 2019, 2021, 2023"}], '
              '"likely": [{"q": "likely question", "why": "short reason"}]} '
              "Sort topics by count, most first. Give 6-10 topics and 5-8 likely questions."
              f"\n\nQuestion papers:\n{context_text(chunks, idx)}")
    data = parse_json(llm(prompt, json_mode=True))
    topics = [t for t in data.get("topics", []) if t.get("topic")]
    for t in topics:
        try:
            t["count"] = int(t.get("count", 1))
        except (TypeError, ValueError):
            t["count"] = 1
    return {"topics": sorted(topics, key=lambda t: -t["count"]),
            "likely": [x for x in data.get("likely", []) if x.get("q")]}

def past_text(res):
    out = ["বারবার আসা বিষয়", ""]
    out += [f"{bn(i)}. {t['topic']} ({bn(t['count'])} বার) {t.get('note', '')}" for i, t in enumerate(res["topics"], 1)]
    out += ["", "এবার আসতে পারে", ""]
    out += [f"{bn(i)}. {x['q']}  ({x.get('why', '')})" for i, x in enumerate(res["likely"], 1)]
    return "\n".join(out)

# ================================================================
#  প্রশ্নপত্র (শিক্ষকের জন্য)
# ================================================================
def make_creative(chunks, n, level):
    """বোর্ডের সৃজনশীল প্রশ্ন: উদ্দীপক + ক(১) খ(২) গ(৩) ঘ(৪), নমুনা উত্তর সহ"""
    idx = overview(chunks, budget=9000, shuffle=True)
    prompt = (f"তুমি বাংলাদেশের {level} বোর্ড পরীক্ষার অভিজ্ঞ প্রশ্নকর্তা। নিচের অংশগুলো থেকেই {n}টা সৃজনশীল প্রশ্ন বানাও। "
              "প্রতিটায় থাকবে: একটা নতুন বাস্তব উদ্দীপক (৩-৫ লাইনের ছোট ঘটনা/তথ্য, সরাসরি বই থেকে কপি নয়), তারপর "
              "ক) জ্ঞানমূলক (১ নম্বর), খ) অনুধাবনমূলক (২ নম্বর), গ) প্রয়োগমূলক (৩ নম্বর, উদ্দীপক ব্যবহার করে), "
              "ঘ) উচ্চতর দক্ষতামূলক (৪ নম্বর, বিশ্লেষণ/মূল্যায়ন)। প্রতিটা অংশের নম্বর অনুযায়ী দৈর্ঘ্যের নমুনা উত্তর দাও। "
              f"{'Write everything in English.' if lang == 'English' else 'সব বাংলায় লেখো।'} "
              'শুধু JSON দাও, এভাবে: {"items": [{"stem": "উদ্দীপক", "part": 1, "parts": [{"label": "ক", "q": "প্রশ্ন", '
              '"marks": 1, "answer": "নমুনা উত্তর"}]}]} যেখানে part হলো যে [নম্বর] অংশ থেকে প্রশ্ন এসেছে।'
              f"\n\nফাইলের অংশ:\n{context_text(chunks, idx)}")
    items = []
    for it in parse_json(llm(prompt, json_mode=True)).get("items", []):
        parts = [p for p in it.get("parts", []) if p.get("q")]
        if it.get("stem") and parts:
            items.append({"stem": it["stem"], "parts": parts, "src": src_of(chunks, idx, it.get("part"))})
    return {"kind": "creative", "title": f"{level} সৃজনশীল প্রশ্ন", "meta": "", "items": items}

# প্রতিটা ধরনের প্রশ্ন দেখতে কেমন হবে, AI কে এভাবে বুঝিয়ে দিই
Q_TYPES = {
    "অতি সংক্ষিপ্ত": "এক কথায় বা এক লাইনে উত্তর: সংজ্ঞা, নাম, তথ্য",
    "সংক্ষিপ্ত": "২-৫ লাইনের বা আধা পাতার উত্তর: ছোট ব্যাখ্যা, পার্থক্য, বৈশিষ্ট্য",
    "রচনামূলক": "বিস্তারিত আলোচনা, বিশ্লেষণ, উদাহরণ, চিত্রের বর্ণনা বা ধাপ",
    "MCQ": "৪টা option সহ বহুনির্বাচনী প্রশ্ন, একটাই সঠিক",
    "সমস্যা সমাধান": "অংক, হিসাব, কোড বা বাস্তব সমস্যা সমাধান করতে হবে",
    "Methodology": "research paper এর পদ্ধতি নিয়ে: research design, data, sample, tools, কেন এই পদ্ধতি",
    "Findings": "research paper এর ফলাফল নিয়ে: মূল findings, সংখ্যা/তুলনা, ফলাফলের অর্থ",
    "Critique": "research paper এর সমালোচনা: limitation, validity, bias, কী ভালো হতে পারত",
    "Contribution": "research paper এর নতুনত্ব আর অবদান, future work",
}

def row(name, count, marks, kind, answer=None):
    return {"বিভাগ": name, "প্রশ্ন": count, "উত্তর দেবে": answer or count, "নম্বর": marks, "ধরন": kind}

# প্রতিষ্ঠান -> পরীক্ষা -> ধাঁচ। সব শুধু শুরুর ধাঁচ, অ্যাপে টেবিলে বদলানো যায়।
PRESETS = {
    "diploma": {
        "BTEB": {
            "সেমিস্টার Final": [row("ক বিভাগ", 10, 2, "সংক্ষিপ্ত", 8), row("খ বিভাগ", 6, 5, "রচনামূলক", 4)],
            "ক্লাস টেস্ট": [row("ক বিভাগ", 5, 2, "সংক্ষিপ্ত"), row("খ বিভাগ", 2, 5, "রচনামূলক")],
            "MCQ টেস্ট": [row("MCQ", 20, 1, "MCQ")],
        },
        "নিজের মতো": {"ফাঁকা থেকে শুরু": [row("ক বিভাগ", 5, 2, "সংক্ষিপ্ত")]},
    },
    "uni": {
        "জাতীয় (NU)": {
            "অনার্স Final": [row("ক বিভাগ", 12, 1, "অতি সংক্ষিপ্ত", 10), row("খ বিভাগ", 8, 4, "সংক্ষিপ্ত", 5),
                             row("গ বিভাগ", 8, 10, "রচনামূলক", 5)],
            "ইনকোর্স": [row("ক বিভাগ", 5, 2, "সংক্ষিপ্ত"), row("খ বিভাগ", 3, 5, "রচনামূলক", 2)],
        },
        "পাবলিক": {
            "Final": [row("Section A", 6, 4, "সংক্ষিপ্ত", 5), row("Section B", 7, 10, "রচনামূলক", 5)],
            "Midterm / Incourse": [row("Section A", 5, 2, "সংক্ষিপ্ত"), row("Section B", 3, 10, "রচনামূলক", 2)],
        },
        "প্রাইভেট": {
            "Mid": [row("Part A", 5, 2, "সংক্ষিপ্ত"), row("Part B", 3, 5, "রচনামূলক")],
            "Final": [row("Part A", 5, 5, "সংক্ষিপ্ত", 4), row("Part B", 3, 8, "রচনামূলক", 2),
                      row("Part C", 1, 4, "সমস্যা সমাধান")],
            "Quiz / CT": [row("MCQ", 10, 1, "MCQ"), row("Short", 1, 5, "সংক্ষিপ্ত")],
        },
        "Research paper": {
            "পুরো বিশ্লেষণ": [row("Methodology", 2, 5, "Methodology"), row("Findings", 2, 5, "Findings"),
                              row("Critique", 1, 10, "Critique"), row("Contribution", 1, 5, "Contribution")],
            "ছোট ক্লাস টেস্ট": [row("Methodology", 1, 5, "Methodology"), row("Findings", 1, 5, "Findings"),
                               row("Critique", 1, 5, "Critique")],
        },
        "নিজের মতো": {"ফাঁকা থেকে শুরু": [row("ক বিভাগ", 5, 2, "সংক্ষিপ্ত")]},
    },
}

ROLES = {"diploma": "বাংলাদেশ কারিগরি শিক্ষা বোর্ড (BTEB) এর ডিপ্লোমা-ইন-ইঞ্জিনিয়ারিং পরীক্ষার অভিজ্ঞ প্রশ্নকর্তা",
         "uni": "একজন বাংলাদেশি বিশ্ববিদ্যালয়ের অভিজ্ঞ কোর্স শিক্ষক ও প্রশ্নকর্তা"}

def to_int(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default

def clean_rows(rows):
    """টেবিলের সারি থেকে ফাঁকা/ভুল সারি বাদ দেয়"""
    out = []
    for r in rows:
        count = to_int(r.get("প্রশ্ন"))
        if count <= 0:
            continue
        marks = r.get("নম্বর")
        try:
            marks = float(marks)
            marks = int(marks) if marks == int(marks) else marks
        except (TypeError, ValueError):
            marks = 1
        ans = to_int(r.get("উত্তর দেবে"), count)
        kind = r.get("ধরন") if r.get("ধরন") in Q_TYPES else "সংক্ষিপ্ত"
        out.append({"name": str(r.get("বিভাগ") or f"বিভাগ {bn(len(out) + 1)}").strip(), "count": min(count, 30),
                    "answer": max(1, min(ans, count)), "marks": marks, "kind": kind})
    return out

def total_marks(rows):
    return sum(r["answer"] * r["marks"] for r in rows)

def section_note(r):
    m = f"{bn(r['answer'])}×{bn(r['marks'])}={bn(r['answer'] * r['marks'])}"
    if r["answer"] < r["count"]:
        return f"যেকোনো {bn(r['answer'])}টি প্রশ্নের উত্তর দাও ({m})"
    return f"সবগুলো প্রশ্নের উত্তর দাও ({m})"

def make_sections(chunks, level, rows, title, meta, bloom=False):
    """টেবিলের ধাঁচ অনুযায়ী বিভাগ ভিত্তিক প্রশ্নপত্র (ডিপ্লোমা আর বিশ্ববিদ্যালয়)"""
    idx = overview(chunks, budget=11000, shuffle=True)
    plan = "\n".join(f"{i}. {r['name']}: ঠিক {r['count']}টা প্রশ্ন, প্রতিটা {r['marks']} নম্বর। "
                     f"ধরন: {r['kind']} ({Q_TYPES[r['kind']]})" for i, r in enumerate(rows, 1))
    tag_rule = ('প্রতিটা প্রশ্নের "tag" এ Bloom level আর CO লিখবে, যেমন "Understand · CO1" বা "Apply · CO2" '
                "(CO1 থেকে CO4 এর মধ্যে, বিষয় অনুযায়ী ভাগ করে)। " if bloom else 'tag ফাঁকা রাখবে। ')
    prompt = (f"তুমি {ROLES[level]}। নিচের ফাইলের অংশগুলো থেকেই এই নিয়মে প্রশ্নপত্র বানাও:\n{plan}\n"
              "প্রশ্নগুলো যেন পুরো অংশ জুড়ে বিভিন্ন বিষয় থেকে আসে, একই জিনিস দুইবার না। "
              "প্রতিটা প্রশ্নের নম্বর অনুযায়ী দৈর্ঘ্যের নমুনা উত্তর দাও (বড় উত্তর point আকারে)। "
              'MCQ ধরনের প্রশ্নে "options" এ ৪টা option দেবে আর "answer" এ সঠিক option টা লিখবে; অন্য ধরনে options ফাঁকা। '
              f"{tag_rule}{write_rule()} "
              'শুধু JSON দাও, এভাবে: {"sections": [{"items": [{"q": "প্রশ্ন", "options": [], "answer": "উত্তর", '
              '"part": 1, "tag": ""}]}]} sections এর ক্রম উপরের নিয়মের মতো, part হলো যে [নম্বর] অংশ থেকে প্রশ্ন এসেছে।'
              f"\n\nফাইলের অংশ:\n{context_text(chunks, idx)}")
    data = parse_json(llm(prompt, json_mode=True))
    out = []
    for r, sec in zip(rows, data.get("sections", [])):
        items = []
        for x in (sec.get("items") or [])[:r["count"]]:
            if x.get("q"):
                items.append({"q": x["q"], "marks": r["marks"], "answer": str(x.get("answer", "")),
                              "options": [str(o) for o in (x.get("options") or [])][:6],
                              "tag": x.get("tag", "") if bloom else "",
                              "src": src_of(chunks, idx, x.get("part"))})
        out.append({"head": r["name"], "note": section_note(r), "items": items})
    if not any(s["items"] for s in out):
        return None
    return {"kind": "sections", "title": title, "meta": meta, "sections": out}

def detect_pattern(text):
    """আগের প্রশ্নপত্র পড়ে তার বিভাগ, প্রশ্ন সংখ্যা আর নম্বর বের করে টেবিলের সারি বানায়"""
    prompt = ("This is an old exam question paper from Bangladesh. Find its pattern: the sections (groups), "
              "how many questions are in each section, how many the student must answer, and marks per question. "
              f"For each section also pick the closest type from this list: {list(Q_TYPES)}. "
              'Reply JSON only: {"title": "course or exam name if written", "time": "time if written", '
              '"sections": [{"name": "ক বিভাগ", "count": 8, "answer": 5, "marks": 4, "type": "সংক্ষিপ্ত"}]}'
              f"\n\nQuestion paper:\n{text[:9000]}")
    data = parse_json(llm(prompt, json_mode=True))
    rows = [row(s.get("name") or f"বিভাগ {i}", to_int(s.get("count"), 1), s.get("marks") or 1,
                s.get("type") if s.get("type") in Q_TYPES else "সংক্ষিপ্ত", to_int(s.get("answer"), 0) or None)
            for i, s in enumerate(data.get("sections", []), 1)]
    return rows, str(data.get("title") or ""), str(data.get("time") or "")

# ---------------- Download: Word (.docx) ----------------
def add_run(par, text, bold=False, size=11, italic=False):
    r = par.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.size = Pt(size)
    rpr = r._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        fonts.set(qn(a), "Nirmala UI")       # Windows এ বাংলা সুন্দর দেখায়
    sz = OxmlElement("w:szCs")
    sz.set(qn("w:val"), str(size * 2))
    rpr.append(sz)
    if bold:
        rpr.append(OxmlElement("w:bCs"))
    return r

def paper_docx(paper, with_answers):
    """প্রশ্নপত্র Word ফাইলে, সরাসরি print বা PDF করা যায়"""
    doc = Document()
    for s in doc.sections:
        s.left_margin = s.right_margin = Pt(54)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p, paper["title"], bold=True, size=15)
    if paper.get("meta"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, paper["meta"], size=10)
    if paper["kind"] == "creative":
        for i, it in enumerate(paper["items"], 1):
            p = doc.add_paragraph()
            add_run(p, f"{bn(i)}. ", bold=True)
            add_run(p, it["stem"])
            for x in it["parts"]:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Pt(18)
                add_run(p, f"{x.get('label', '')}) {x['q']}  ({bn(x.get('marks', ''))})")
                if with_answers:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Pt(36)
                    add_run(p, f"উত্তর: {x.get('answer', '')}", italic=True, size=10)
    else:
        for sec in paper["sections"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            add_run(p, sec["head"], bold=True, size=12)
            add_run(p, f"   {sec['note']}", size=10, italic=True)
            for i, x in enumerate(sec["items"], 1):
                p = doc.add_paragraph()
                tagtxt = f"  [{x['tag']}]" if x.get("tag") else ""
                add_run(p, f"{bn(i)}. {x['q']}  ({bn(x['marks'])}){tagtxt}")
                for k, o in enumerate(x.get("options", [])):
                    q = doc.add_paragraph()
                    q.paragraph_format.left_indent = Pt(18)
                    add_run(q, f"{'কখগঘঙচ'[k]}) {o}", size=10.5)
                if with_answers and x.get("answer"):
                    q = doc.add_paragraph()
                    q.paragraph_format.left_indent = Pt(18)
                    add_run(q, f"উত্তর: {x['answer']}", italic=True, size=10)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p, f"Made with {APP_NAME}", size=8, italic=True)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ======================= স্ক্রিনের ছোট অংশগুলো (UI) =======================
def show_sources(sources, query=""):
    """উত্তরের নিচে ছোট pill এ উৎস, আর চাইলে খুলে মূল লেখা দেখা যায়"""
    if not sources:
        return
    names = list(dict.fromkeys(s["src"] for s in sources))
    st.markdown('<div class="src">' + "".join(f"<span>{s}</span>" for s in names) + "</div>",
                unsafe_allow_html=True)
    with st.expander("উৎস দেখুন"):
        if query:
            st.caption(f"খোঁজা হয়েছে: {query}")
        for s in sources:
            st.markdown(f"**{s['src']}**")
            st.caption(s["text"][:500] + ("..." if len(s["text"]) > 500 else ""))

def build_chunks(files):
    """সব ফাইলের টুকরো এক তালিকায়; একাধিক ফাইল হলে উৎসে ফাইলের নামও থাকে"""
    chunks, words = [], 0
    for name, data in files.items():
        try:
            ch, full, _, _ = load_document(data, name)
        except Exception:
            st.toast(f"'{name}' পড়া গেল না")
            continue
        short = name if len(name) <= 28 else name[:25] + "..."
        for c in ch:
            chunks.append({**c, "file": name,
                           "src": f"{short} · {c['src']}" if len(files) > 1 else c["src"]})
        words += len(full.split())
    return chunks, words

def add_uploads():
    """অনুশীলন/প্রশ্নপত্র ট্যাবের uploader থেকে ফাইল জমা রাখে, তারপর uploader খালি করে"""
    for f in st.session_state.get(f"up_{st.session_state.up_id}") or []:
        st.session_state.files[f.name] = f.getvalue()
    st.session_state.up_id += 1

def uploader():
    st.file_uploader("ফাইল যোগ করুন", type=TYPES, accept_multiple_files=True,
                     key=f"up_{st.session_state.up_id}", on_change=add_uploads, label_visibility="collapsed")

def file_pills(words, chunks, show_add=False):
    """যোগ করা ফাইল ছোট pill এ, চাপলে সরে যায়"""
    files = st.session_state.files
    row_ = st.container(horizontal=True, gap="small", key="pills")
    for name in list(files):
        if row_.button(f"{name}  ✕", key=f"rm_{name}", help="চাপলে ফাইলটা সরে যাবে"):
            del st.session_state.files[name]
            st.rerun()
    if show_add:
        with row_.popover("＋ ফাইল"):
            uploader()
    st.markdown(f'<div class="meta">{bn(len(files))}টা ফাইল · {bn(f"{words:,}")} শব্দ · '
                f'{bn(len(chunks))}টা টুকরো</div>', unsafe_allow_html=True)

def pick_source(chunks, key):
    """কোন ফাইল আর কোন পাতা/Slide থেকে কাজ হবে, তা বাছাই করা"""
    names = list(dict.fromkeys(c["file"] for c in chunks))
    with st.expander("কোন ফাইল, কোন পাতা থেকে", icon=":material/tune:"):
        pick = st.selectbox("ফাইল", ["সব ফাইল"] + names, key=f"{key}_file") if len(names) > 1 else names[0]
        sel = chunks if pick == "সব ফাইল" else [c for c in chunks if c["file"] == pick]
        if pick != "সব ফাইল" and sel:
            top, word = max(c["n"] for c in sel), sel[0]["kind"]
            if top > 1:
                a, b = st.slider(f"কোন {word} থেকে কোন {word} পর্যন্ত", 1, top, (1, top), key=f"{key}_rng_{pick}")
                sel = [c for c in sel if a <= c["n"] <= b]
        st.caption(f"{bn(len(sel))}টা টুকরো থেকে কাজ হবে")
    return sel or chunks

def need_files(title, sub):
    st.markdown(f'<div class="hello"><h1>{title}</h1><p>{sub}</p></div>', unsafe_allow_html=True)
    uploader()

def step(n, title, hint=""):
    """ধাপের শিরোনাম: রঙিন নম্বর + লেখা, যাতে ব্যবহারকারী বোঝে এখন কী করতে হবে"""
    st.markdown(f'<div class="step"><span>{bn(n)}</span><div><b>{title}</b>'
                + (f"<small>{hint}</small>" if hint else "") + "</div></div>", unsafe_allow_html=True)

def tip(text):
    st.markdown(f'<div class="tip">{text}</div>', unsafe_allow_html=True)

def open_tab(name):
    st.session_state.tab = name
    st.session_state.last_tab = name

def go_practice():
    open_tab("অনুশীলন")

def keep_tab():
    """একই ট্যাবে আবার চাপলে ট্যাব যেন খালি না হয়ে যায়"""
    if st.session_state.tab is None:
        st.session_state.tab = st.session_state.last_tab
    st.session_state.last_tab = st.session_state.tab

TAB_LABEL = {"চ্যাট": ":material/forum: চ্যাট", "অনুশীলন": ":material/school: অনুশীলন",
             "প্রশ্নপত্র": ":material/edit_document: প্রশ্নপত্র"}

# ======================= উপরের সরু header =======================
top = st.container(key="topbar")
h1, h2 = top.columns([2, 3], vertical_alignment="center")
h1.markdown(brand_html("Learn · Teach"), unsafe_allow_html=True)
tab = h2.segmented_control("tab", TABS, key="tab", on_change=keep_tab, label_visibility="collapsed",
                           format_func=TAB_LABEL.get) or "চ্যাট"

# ======================= চ্যাট: নিচের বক্স =======================
typed = None
if tab == "চ্যাট":
    msg = st.chat_input("প্রশ্ন লিখুন, বা + চেপে ফাইল যোগ করুন", accept_file="multiple", file_type=TYPES)
    if msg:
        for f in msg.files:
            st.session_state.files[f.name] = f.getvalue()
        typed = msg.text or None

files = st.session_state.files
if files:
    with st.spinner("ফাইল পড়া হচ্ছে..."):
        chunks, words = build_chunks(files)
else:
    chunks, words = [], 0

if st.session_state.get("file_names") != tuple(files):   # ফাইল বদলালে Quiz/প্রশ্নপত্র নতুন করে
    st.session_state.file_names = tuple(files)
    st.session_state.quiz = None
    st.session_state.quiz_done = False
    st.session_state.papers = {}
    st.session_state.past = None
    st.session_state.viva = {"current": None, "done": []}

# ======================= ১) চ্যাট =======================
if tab == "চ্যাট":
    quick = None
    hello = st.empty()
    if not st.session_state.chat:
        with hello.container():
            if not files:
                st.markdown(f'<div class="hello home"><div class="tag">✨ {TAGLINE}</div><h1>আজ কী শিখবেন, কী শেখাবেন?</h1>'
                            '<p>বই, নোট, slide বা research paper দিন। বাকিটা Shikkha.ai করবে।</p></div>',
                            unsafe_allow_html=True)
                feats = st.container(key="feats")
                cols = feats.columns(3)
                for col, (icon, name, desc, short) in zip(cols, [
                        ("💬", "চ্যাট", "ফাইল নিয়ে যেকোনো প্রশ্ন করুন, পাতা নম্বর সহ উত্তর পান", "ফাইল থেকে প্রশ্ন-উত্তর"),
                        ("🎯", "অনুশীলন", "শিক্ষার্থীদের জন্য Quiz, Viva আর দুর্বল বিষয় খোঁজা", "MCQ · Viva অনুশীলন"),
                        ("📝", "প্রশ্নপত্র", "শিক্ষকদের জন্য নিজের ধাঁচে প্রশ্ন আর উত্তর, Word এ", "নিজের ধাঁচে প্রশ্ন বানান")]):
                    with col.container(border=True):
                        st.markdown(f'<div class="feat"><i>{icon}</i><b>{name}</b><small>{short}</small><p>{desc}</p></div>',
                                    unsafe_allow_html=True)
                        st.button("খুলুন", key=f"feat_{name}", icon=":material/arrow_forward:", width="stretch", on_click=open_tab, args=(name,))
                st.markdown('<div class="howto"><span><b>১</b>ফাইল দিন</span><span><b>২</b>উপরে ট্যাব বাছুন</span>'
                            '<span><b>৩</b>প্রশ্ন করুন বা বানান</span></div>', unsafe_allow_html=True)
                with st.container(key="homeup"):     # ফোনে লুকানো থাকে, সেখানে নিচের + দিয়েই ফাইল দেওয়া যায়
                    uploader()
            else:
                st.markdown('<div class="hello"><h1>কী জানতে চান?</h1>'
                            '<p>যেকোনো প্রশ্ন করুন, বা নিচের একটা দিয়ে শুরু করুন</p></div>',
                            unsafe_allow_html=True)
                sg = st.container(key="suggest")
                s1, s2 = sg.columns(2)
                if s1.button("সংক্ষেপে বলো", key="sg1", icon=":material/notes:", width="stretch"):
                    quick = "এই ফাইল(গুলো) সংক্ষেপে বুঝিয়ে বলো"
                if s2.button("মূল ৫টা পয়েন্ট", key="sg2", icon=":material/format_list_bulleted:", width="stretch"):
                    quick = "এখান থেকে সবচেয়ে গুরুত্বপূর্ণ ৫টা point দাও"
                s3, s4 = sg.columns(2)
                if s3.button("সহজ করে বোঝাও", key="sg3", icon=":material/lightbulb:", width="stretch"):
                    quick = "সবচেয়ে কঠিন বিষয়টা উদাহরণ দিয়ে সহজ করে বুঝিয়ে দাও"
                s4.button("Quiz দাও", key="sg4", icon=":material/quiz:", width="stretch", on_click=go_practice)

    q = typed or quick
    if q:
        hello.empty()
    for m in st.session_state.chat:
        with st.chat_message(m["role"], avatar=":material/person:" if m["role"] == "user" else ":material/auto_awesome:"):
            st.markdown(m["text"])
            show_sources(m.get("sources"), m.get("query", ""))
    if q:
        st.chat_message("user", avatar=":material/person:").markdown(q)
        with st.chat_message("assistant", avatar=":material/auto_awesome:"):
            if not chunks:
                ans, sources, query = "আগে নিচের + চেপে একটা ফাইল যোগ করুন, তারপর প্রশ্ন করুন।", [], ""
            else:
                with st.spinner("ফাইল খুঁজে উত্তর লেখা হচ্ছে..."):
                    try:
                        ans, sources, query = ask(q, chunks, "search" if typed else "overview")
                    except Exception:
                        ans, sources, query = "এখন উত্তর আনা গেল না। একটু পরে আবার চেষ্টা করুন।", [], ""
            st.markdown(ans)
            show_sources(sources, query)
        st.session_state.chat += [{"role": "user", "text": q},
                                  {"role": "assistant", "text": ans, "sources": sources, "query": query}]
    if files:
        file_pills(words, chunks)

# ======================= ২) অনুশীলন (শিক্ষার্থীর জন্য) =======================
def record_result(quiz, quiz_id):
    """প্রতিটা MCQ এর ফল topic অনুযায়ী জমা রাখি (একই quiz দুইবার গোনা হবে না)"""
    if quiz_id in st.session_state.recorded:
        return
    st.session_state.recorded.add(quiz_id)
    for i, item in enumerate(quiz):
        t = item["topic"]
        right, total = st.session_state.topic_stats.get(t, (0, 0))
        ok = st.session_state.get(f"quiz_{quiz_id}_{i}") == item["options"][item["answer"]]
        st.session_state.topic_stats[t] = (right + ok, total + 1)

def weak_topics():
    """৬০% এর কম সঠিক হলে দুর্বল বিষয়, সবচেয়ে খারাপটা আগে"""
    rows = [(t, r, n) for t, (r, n) in st.session_state.topic_stats.items() if n and r / n < 0.6]
    return sorted(rows, key=lambda x: x[1] / x[2])

def new_quiz(quiz):
    st.session_state.quiz = quiz
    st.session_state.quiz_id += 1
    st.session_state.quiz_done = False

if tab == "অনুশীলন":
    if not chunks:
        need_files("অনুশীলন করুন", "Quiz, Viva আর আগের প্রশ্ন বিশ্লেষণ। আগে নোট বা বই যোগ করুন")
    else:
        file_pills(words, chunks, show_add=True)
        step(1, "কী অনুশীলন করবেন?", "যেটা বাছবেন, নিচে তার নিয়ম দেখাবে")
        ptype = st.segmented_control("ধরন", ["MCQ Quiz", "Viva", "আগের প্রশ্ন বিশ্লেষণ"], key="ptype",
                                     default="MCQ Quiz", label_visibility="collapsed",
                                     format_func=lambda x: {"MCQ Quiz": ":material/quiz: MCQ Quiz",
                                                            "Viva": ":material/record_voice_over: Viva",
                                                            "আগের প্রশ্ন বিশ্লেষণ": ":material/insights: আগের প্রশ্ন"}[x]
                                     ) or "MCQ Quiz"
        tip({"MCQ Quiz": "<b>MCQ Quiz:</b> উত্তর জমা দিলে সাথে সাথে নম্বর আর ব্যাখ্যা পাবেন। বারবার ভুল হওয়া "
                         "বিষয়গুলো <b>দুর্বল বিষয়</b> হিসেবে জমা হবে, সেখান থেকে আবার Quiz নিতে পারবেন।",
             "Viva": "<b>Viva:</b> শিক্ষক যেভাবে মুখে প্রশ্ন করেন, অ্যাপ সেভাবে একটা করে প্রশ্ন করবে। আপনি লিখে উত্তর দিন, "
                     "১০ এর মধ্যে নম্বর, পরামর্শ আর আদর্শ উত্তর পাবেন।",
             "আগের প্রশ্ন বিশ্লেষণ": "<b>আগের প্রশ্ন:</b> কয়েক বছরের পুরনো প্রশ্নপত্র ফাইল হিসেবে যোগ করুন "
                                     "(উপরের ＋ ফাইল)। অ্যাপ বলবে কোন বিষয় বারবার আসে আর এবার কী আসতে পারে।"}[ptype])
        step(2, "কোন অংশ থেকে?", "পুরো ফাইল, নাকি নির্দিষ্ট পাতা বা Slide")
        src = pick_source(chunks, "prac")
        step(3, "শুরু করুন")

        # ---------- MCQ + দুর্বল বিষয় ----------
        if ptype == "MCQ Quiz":
            qc1, qc2, qc3 = st.columns([1, 1.3, 1.6], vertical_alignment="bottom")
            n = qc1.selectbox("প্রশ্ন", [5, 10, 15], index=0)
            lvl = qc2.selectbox("মান", ["সাধারণ", "ভর্তি পরীক্ষার মতো কঠিন"])
            if qc3.button("নতুন Quiz", type="primary", width="stretch", icon=":material/quiz:"):
                with st.spinner("Quiz তৈরি হচ্ছে..."):
                    try:
                        new_quiz(make_quiz(src, n, hard=lvl != "সাধারণ"))
                    except Exception:
                        new_quiz([])
                if st.session_state.quiz:
                    st.toast("Quiz তৈরি! নিচে উত্তর দিন", icon="📝")
                else:
                    st.error("Quiz বানানো গেল না, আবার চেষ্টা করুন।")

            weak = weak_topics()
            if st.session_state.topic_stats:
                with st.container(border=True, key="weak"):
                    st.markdown("**দুর্বল বিষয়**" if weak else "**দুর্বল বিষয়**: এখনো নেই, দারুণ!")
                    for t, r, total in weak[:5]:
                        st.progress(r / total, text=f"{t} · {bn(r)}/{bn(total)} সঠিক")
                    if weak and st.button("এই বিষয়গুলো আবার অনুশীলন করুন", icon=":material/refresh:"):
                        with st.spinner("দুর্বল বিষয় থেকে Quiz তৈরি হচ্ছে..."):
                            try:
                                new_quiz(make_quiz(chunks, 5, focus=[t for t, _, _ in weak[:4]]))
                            except Exception:
                                new_quiz([])

            quiz, qid = st.session_state.quiz, st.session_state.quiz_id
            if quiz:
                with st.form(f"quiz_form_{qid}", border=False):
                    for i, item in enumerate(quiz):
                        st.radio(f"**{i+1}. {item['q']}**", item["options"], index=None, key=f"quiz_{qid}_{i}")
                    submitted = st.form_submit_button("উত্তর জমা দিন", type="primary")
                if submitted:
                    st.session_state.quiz_done = True
                    record_result(quiz, qid)
                    st.rerun()

                if st.session_state.quiz_done:
                    score = sum(st.session_state.get(f"quiz_{qid}_{i}") == it["options"][it["answer"]]
                                for i, it in enumerate(quiz))
                    pct = round(score * 100 / len(quiz))
                    msg_txt = "অসাধারণ!" if pct >= 80 else "ভালো হয়েছে!" if pct >= 50 else "আরেকটু পড়ে আবার চেষ্টা করুন"
                    st.markdown(f'<div class="score"><b>{score}/{len(quiz)}</b><p>{pct}% · {msg_txt}</p></div>',
                                unsafe_allow_html=True)
                    for i, item in enumerate(quiz):
                        correct = item["options"][item["answer"]]
                        mine = st.session_state.get(f"quiz_{qid}_{i}")
                        if mine == correct:
                            st.success(f"{i+1}. সঠিক: {correct}", icon=":material/check:")
                        elif mine is None:
                            st.warning(f"{i+1}. উত্তর দেননি। সঠিক: {correct}", icon=":material/remove:")
                        else:
                            st.error(f"{i+1}. আপনার উত্তর: {mine} · সঠিক: {correct}", icon=":material/close:")
                        st.caption(" · ".join(x for x in [item["topic"], item["why"], item.get("src", "")] if x))
                    if pct >= 80:
                        st.balloons()

        # ---------- Viva ----------
        elif ptype == "Viva":
            viva = st.session_state.viva
            if not viva["current"]:
                if st.button("Viva শুরু করুন" if not viva["done"] else "পরের প্রশ্ন", type="primary",
                             icon=":material/record_voice_over:"):
                    with st.spinner("প্রশ্ন ভাবা হচ্ছে..."):
                        try:
                            viva["current"] = viva_question(src, [d["q"] for d in viva["done"]])
                        except Exception:
                            st.error("প্রশ্ন আনা গেল না, আবার চেষ্টা করুন।")
                    if viva["current"]:
                        st.rerun()
            else:
                with st.container(border=True):
                    st.markdown(f"**প্রশ্ন {bn(len(viva['done']) + 1)}:** {viva['current']['q']}")
                    with st.form(f"viva_{len(viva['done'])}", border=False):
                        answer = st.text_area("আপনার উত্তর", height=120, placeholder="যেভাবে মুখে বলতেন, সেভাবে লিখুন")
                        sent = st.form_submit_button("উত্তর দিন", type="primary")
                    if sent and answer.strip():
                        with st.spinner("নম্বর দেওয়া হচ্ছে..."):
                            try:
                                g = viva_grade(viva["current"], answer)
                            except Exception:
                                g = {"score": 0, "feedback": "এখন যাচাই করা গেল না।", "model": ""}
                        viva["done"].append({**viva["current"], "answer": answer, **g})
                        viva["current"] = None
                        st.rerun()
            if viva["done"]:
                avg = sum(d["score"] for d in viva["done"]) / len(viva["done"])
                st.markdown(f'<div class="score"><b>{bn(f"{avg:.1f}")}/১০</b>'
                            f'<p>{bn(len(viva["done"]))}টা প্রশ্নের গড়</p></div>', unsafe_allow_html=True)
                for k, d in reversed(list(enumerate(viva["done"], 1))):
                    with st.container(border=True):
                        st.markdown(f"**{bn(k)}. {d['q']}**  ·  **{bn(d['score'])}/১০**")
                        st.caption(f"আপনার উত্তর: {d['answer']}")
                        st.markdown(d["feedback"])
                        if d["model"]:
                            with st.expander("আদর্শ উত্তর"):
                                st.markdown(d["model"])
                                st.caption(f"উৎস: {d['src']}")
                if st.button("Viva আবার শুরু", icon=":material/restart_alt:"):
                    st.session_state.viva = {"current": None, "done": []}
                    st.rerun()

        # ---------- আগের প্রশ্ন বিশ্লেষণ ----------
        else:
            if st.button("প্রশ্ন বিশ্লেষণ করো", type="primary", icon=":material/insights:"):
                with st.spinner("সব প্রশ্নপত্র পড়ে বিশ্লেষণ করা হচ্ছে..."):
                    try:
                        st.session_state.past = analyze_past(src)
                    except Exception:
                        st.session_state.past = None
                if not st.session_state.past:
                    st.error("বিশ্লেষণ করা গেল না, আবার চেষ্টা করুন।")
            res = st.session_state.past
            if res:
                st.markdown("#### বারবার আসা বিষয়")
                topc = max([t["count"] for t in res["topics"]] or [1])
                for t in res["topics"]:
                    st.progress(min(1.0, t["count"] / topc), text=f"{t['topic']} · {bn(t['count'])} বার"
                                + (f" · {t['note']}" if t.get("note") else ""))
                st.markdown("#### এবার আসতে পারে")
                for i, x in enumerate(res["likely"], 1):
                    st.markdown(f"{bn(i)}. {x['q']}")
                    if x.get("why"):
                        st.caption(x["why"])
                st.download_button("বিশ্লেষণ download", past_text(res), file_name="past_analysis.txt",
                                   icon=":material/download:")

# ======================= ৩) প্রশ্নপত্র (শিক্ষকের জন্য) =======================
def set_preset(level, reset_exam=False):
    insts = PRESETS[level]
    inst = st.session_state.get(f"inst_{level}") or list(insts)[0]
    st.session_state[f"inst_{level}"] = inst
    exams = insts[inst]
    if reset_exam or st.session_state.get(f"exam_{level}") not in exams:
        st.session_state[f"exam_{level}"] = list(exams)[0]
    st.session_state[f"pat_{level}"] = [dict(r) for r in exams[st.session_state[f"exam_{level}"]]]
    st.session_state[f"patv_{level}"] = st.session_state.get(f"patv_{level}", 0) + 1

def exam_changed(level):
    if st.session_state.get(f"exam_{level}"):     # pill আবার চাপলে খালি হয়, তখন টেবিল বদলাই না
        set_preset(level)

def detect_old_paper(level):
    """বোতাম চাপলে: পুরনো প্রশ্নপত্র পড়ে টেবিল, কোর্সের নাম আর সময় ভরে দেয়"""
    old = st.session_state.get(f"old_{level}")
    try:
        _, text, _, _ = load_document(old.getvalue(), old.name)
        found, title, time_ = detect_pattern(text)
    except Exception:
        found, title, time_ = [], "", ""
    if not found:
        st.session_state[f"detect_fail_{level}"] = True
        return
    st.session_state[f"inst_{level}"] = "নিজের মতো"
    st.session_state[f"exam_{level}"] = list(PRESETS[level]["নিজের মতো"])[0]
    st.session_state[f"pat_{level}"] = found
    st.session_state[f"patv_{level}"] += 1
    if title:
        st.session_state[f"course_{level}"] = title
    if time_:
        st.session_state[f"time_{level}"] = time_
    st.toast("ধাঁচ পাওয়া গেছে, টেবিল ভরে দিলাম", icon="✅")

def pattern_editor(level):
    """প্রশ্নপত্রের ধাঁচ: প্রতিষ্ঠান -> পরীক্ষা বাছাই, তারপর টেবিলে নিজের মতো বদলানো"""
    insts = PRESETS[level]
    if f"pat_{level}" not in st.session_state:
        set_preset(level, reset_exam=True)
    st.segmented_control("প্রতিষ্ঠান", list(insts), key=f"inst_{level}", on_change=set_preset,
                         args=(level, True))
    inst = st.session_state.get(f"inst_{level}") or list(insts)[0]
    if len(insts[inst]) > 1:
        st.pills("পরীক্ষা", list(insts[inst]), key=f"exam_{level}", on_change=exam_changed, args=(level,))
    tip("<b>টিপস:</b> প্রতিটা প্রতিষ্ঠানের নিয়ম আলাদা। নিচের টেবিলের যেকোনো ঘরে চাপ দিয়ে বিভাগ, প্রশ্ন সংখ্যা, "
        "কয়টার উত্তর দেবে, নম্বর আর ধরন বদলান। নতুন বিভাগ লাগলে টেবিলের শেষ ফাঁকা সারিতে লিখুন।")
    df = st.data_editor(
        pd.DataFrame(st.session_state[f"pat_{level}"]), key=f"ed_{level}_{st.session_state[f'patv_{level}']}",
        num_rows="dynamic", hide_index=True, width="stretch",
        column_config={
            "বিভাগ": st.column_config.TextColumn("বিভাগ", width="small"),
            "প্রশ্ন": st.column_config.NumberColumn("প্রশ্ন", min_value=1, max_value=30, step=1,
                                                    help="প্রশ্নপত্রে কয়টা প্রশ্ন থাকবে"),
            "উত্তর দেবে": st.column_config.NumberColumn("উত্তর দেবে", min_value=1, max_value=30, step=1,
                                                         help="ছাত্রকে কয়টার উত্তর দিতে হবে (যেমন ৮টার মধ্যে ৫টা)"),
            "নম্বর": st.column_config.NumberColumn("নম্বর", min_value=0.5, max_value=100, step=0.5, format="%g",
                                                    help="প্রতিটা প্রশ্নের নম্বর"),
            "ধরন": st.column_config.SelectboxColumn("ধরন", options=list(Q_TYPES), required=True),
        })
    rows = clean_rows(df.to_dict("records"))
    n_q = sum(r["count"] for r in rows)
    st.markdown(f'<div class="total">পূর্ণমান <b>{bn(total_marks(rows))}</b> · {bn(len(rows))}টা বিভাগ · '
                f'{bn(n_q)}টা প্রশ্ন</div>', unsafe_allow_html=True)
    if n_q > 40:
        st.warning("একবারে ৪০টার বেশি প্রশ্ন হলে সময় বেশি লাগবে, ভুলও হতে পারে। দুই ভাগে বানানো ভালো।")

    with st.expander("আগের প্রশ্নপত্র দিলে ধাঁচ নিজে থেকে ভরে যাবে", icon=":material/auto_fix_high:"):
        st.caption("আগের বছরের একটা প্রশ্নপত্র দিন। অ্যাপ তার বিভাগ, প্রশ্ন সংখ্যা আর নম্বর বের করে টেবিলটা ভরে দেবে।")
        old = st.file_uploader("পুরনো প্রশ্নপত্র", type=TYPES, key=f"old_{level}", label_visibility="collapsed")
        if old:
            st.button("ধাঁচ বের করো", key=f"detect_{level}", icon=":material/table_chart:",
                      on_click=detect_old_paper, args=(level,))
        if st.session_state.pop(f"detect_fail_{level}", False):
            st.error("ধাঁচ বোঝা গেল না। লেখা আছে এমন PDF/Word দিন (ছবি থেকে scan করা হলে কাজ নাও করতে পারে)।")
    return rows

def show_paper(paper):
    """প্রশ্নপত্র স্ক্রিনে দেখানো, উত্তর expander এ লুকানো"""
    st.markdown(f'<div class="paper-head"><h3>{paper["title"]}</h3>'
                + (f'<p>{paper["meta"]}</p>' if paper.get("meta") else "") + "</div>", unsafe_allow_html=True)
    if paper["kind"] == "creative":
        for i, it in enumerate(paper["items"], 1):
            with st.container(border=True):
                st.markdown(f"**{bn(i)}. উদ্দীপক**")
                st.markdown(it["stem"])
                for p in it["parts"]:
                    st.markdown(f"**{p.get('label', '')})** {p['q']} *({bn(p.get('marks', ''))} নম্বর)*")
                with st.expander("নমুনা উত্তর"):
                    for p in it["parts"]:
                        st.markdown(f"**{p.get('label', '')})** {p.get('answer', '')}")
                if it["src"]:
                    st.caption(f"উৎস: {it['src']}")
    else:
        for sec in paper["sections"]:
            st.markdown(f'<div class="sec-head">{sec["head"]}</div><div class="sec-note">{sec["note"]}</div>',
                        unsafe_allow_html=True)
            for i, x in enumerate(sec["items"], 1):
                with st.container(border=True):
                    tagtxt = f" · `{x['tag']}`" if x.get("tag") else ""
                    st.markdown(f"{bn(i)}. {x['q']} *({bn(x['marks'])} নম্বর)*{tagtxt}")
                    if x.get("options"):
                        st.markdown("  \n".join(f"{'কখগঘঙচ'[k]}) {o}" for k, o in enumerate(x["options"])))
                    with st.expander("নমুনা উত্তর"):
                        st.markdown(x["answer"] or "-")
                        if x["src"]:
                            st.caption(f"উৎস: {x['src']}")
    d1, d2 = st.columns(2)
    d1.download_button("শুধু প্রশ্ন (Word)", paper_docx(paper, False), file_name="question_paper.docx",
                       width="stretch", icon=":material/description:")
    d2.download_button("উত্তর সহ (Word)", paper_docx(paper, True), file_name="question_answer.docx",
                       width="stretch", icon=":material/fact_check:")

if tab == "প্রশ্নপত্র":
    if not chunks:
        need_files("প্রশ্নপত্র বানান",
                   "শিক্ষকদের জন্য: বই, slide বা research paper দিন, নিজের ধাঁচে প্রশ্ন আর উত্তর পান")
    else:
        file_pills(words, chunks, show_add=True)
        step(1, "কোন পরীক্ষার প্রশ্ন?", "স্কুল-কলেজ, ডিপ্লোমা না বিশ্ববিদ্যালয়")
        qp = st.segmented_control("প্রশ্নপত্র", ["SSC/HSC সৃজনশীল", "ডিপ্লোমা", "বিশ্ববিদ্যালয়"], key="qp",
                                  default="বিশ্ববিদ্যালয়", label_visibility="collapsed") or "বিশ্ববিদ্যালয়"
        step(2, "কোন অংশ থেকে?", "পুরো ফাইল, নাকি নির্দিষ্ট পাতা বা Slide (যেমন ৫ থেকে ২০)")
        src = pick_source(chunks, "qp")

        if qp == "SSC/HSC সৃজনশীল":
            step(3, "কয়টা প্রশ্ন?", "প্রতিটায় উদ্দীপক + ক(১) খ(২) গ(৩) ঘ(৪) = ১০ নম্বর")
            pc1, pc2 = st.columns(2)
            level = pc1.selectbox("শ্রেণি", ["SSC", "HSC"])
            n_items = pc2.selectbox("প্রশ্ন", [1, 2, 3, 4, 5], index=1)
            make = lambda: make_creative(src, n_items, level)
            n_last = 4
        else:
            lv = "diploma" if qp == "ডিপ্লোমা" else "uni"
            step(3, "প্রশ্নের ধাঁচ ঠিক করুন", "প্রতিষ্ঠান আর পরীক্ষা বাছুন, তারপর দরকার হলে টেবিলে বদলান")
            rows = pattern_editor(lv)
            step(4, "শিরোনাম আর অপশন", "এগুলো প্রশ্নপত্রের উপরে ছাপা হবে (না দিলেও চলবে)")
            m1, m2 = st.columns([2, 1])
            course = m1.text_input("কোর্স / বিষয়ের নাম", key=f"course_{lv}",
                                   placeholder="যেমন: CSE 311 Database Systems")
            time_ = m2.text_input("সময়", key=f"time_{lv}", placeholder="যেমন: ৩ ঘণ্টা")
            bloom = st.toggle("প্রতিটা প্রশ্নে Bloom level আর CO লেবেল দাও (OBE সিলেবাসের জন্য)", key=f"bloom_{lv}",
                              value=False)
            title = course.strip() or ("ডিপ্লোমা প্রশ্নপত্র" if lv == "diploma" else "বিশ্ববিদ্যালয় প্রশ্নপত্র")
            meta = " · ".join(x for x in [f"সময়: {time_.strip()}" if time_.strip() else "",
                                          f"পূর্ণমান: {bn(total_marks(rows))}"] if x)
            make = lambda: make_sections(src, lv, rows, title, meta, bloom) if rows else None
            n_last = 5

        step(n_last, "বানান", "প্রশ্নের সাথে নমুনা উত্তরও তৈরি হবে, পরে Word এ download করতে পারবেন")
        if st.button("প্রশ্নপত্র বানাও", type="primary", width="stretch", icon=":material/auto_awesome:"):
            with st.spinner("প্রশ্ন ও নমুনা উত্তর লেখা হচ্ছে... (বড় প্রশ্নপত্রে ১ মিনিট লাগতে পারে)"):
                try:
                    st.session_state.papers[qp] = make()
                except Exception:
                    st.session_state.papers[qp] = None
            if st.session_state.papers.get(qp):
                st.toast("প্রশ্নপত্র তৈরি! নিচে দেখুন", icon="🎉")
            else:
                st.error("প্রশ্ন বানানো গেল না। টেবিলে অন্তত একটা সারি আছে কিনা দেখে আবার চেষ্টা করুন।")

        if st.session_state.papers.get(qp):
            show_paper(st.session_state.papers[qp])

# ======================= Chat download (sidebar এ) =======================
if st.session_state.chat:
    lines = []
    for m in st.session_state.chat:
        if m["role"] == "user":
            lines.append(f"প্রশ্ন: {m['text']}")
        else:
            s = ", ".join(dict.fromkeys(x["src"] for x in m.get("sources", [])))
            lines.append(f"উত্তর: {m['text']}" + (f"\n(উৎস: {s})" if s else ""))
    download_slot.download_button("Chat download", "\n\n".join(lines), file_name="chat.txt", width="stretch")
