import streamlit as st
from groq import Groq
from pypdf import PdfReader

st.set_page_config(page_title="বাংলা AI Assistant", page_icon="📚", layout="wide")
client = Groq(api_key=st.secrets["GROQ_API_KEY"])

@st.cache_data
def pick_model():
    ids = [m.id for m in client.models.list().data]
    chat = [m for m in ids if not any(x in m for x in ["whisper", "guard", "tts"])]
    best = [m for m in chat if any(x in m for x in ["llama", "gpt-oss", "qwen"])]
    return (best or chat)[0]
MODEL = pick_model()

css = """
.stApp {background: linear-gradient(135deg, #eef2ff 0%, #fdf4ff 50%, #ecfeff 100%);}
@keyframes flow {0% {background-position: 0% 50%;} 50% {background-position: 100% 50%;} 100% {background-position: 0% 50%;}}
.hero {background: linear-gradient(270deg, #4f46e5, #9333ea, #db2777, #0ea5e9);
  background-size: 600% 600%; animation: flow 10s ease infinite;
  padding: 35px; border-radius: 25px; text-align: center; color: white;
  box-shadow: 0 10px 30px rgba(79,70,229,0.35); margin-bottom: 25px;}
.hero h1 {color: white; font-size: 44px; margin: 0;}
.hero p {font-size: 18px; opacity: 0.95; margin-top: 8px;}
.stat {padding: 20px; border-radius: 18px; color: white; text-align: center;
  box-shadow: 0 6px 18px rgba(0,0,0,0.12); transition: transform .2s;}
.stat:hover {transform: translateY(-5px);}
.stat h2 {color: white; margin: 0; font-size: 30px;}
.s1 {background: linear-gradient(135deg, #6366f1, #8b5cf6);}
.s2 {background: linear-gradient(135deg, #ec4899, #f43f5e);}
.s3 {background: linear-gradient(135deg, #06b6d4, #10b981);}
.tip {background: white; padding: 20px; border-radius: 18px; text-align: center;
  border: 2px dashed #a78bfa; font-size: 18px;}
[data-testid="stChatMessage"] {background: white; border-radius: 18px;
  box-shadow: 0 4px 14px rgba(0,0,0,0.06); padding: 12px; margin-bottom: 10px;}
.stButton>button {width: 100%; border-radius: 12px; border: none; color: white;
  background: linear-gradient(90deg, #6366f1, #db2777); font-weight: 600;}
.stButton>button:hover {opacity: 0.85; color: white;}
[data-testid="stSidebar"] {background: linear-gradient(180deg, #312e81, #6b21a8);}
[data-testid="stSidebar"] * {color: white !important;}
.footer {text-align: center; color: #6b7280; margin-top: 40px; font-size: 14px;}
"""
tag = "style"
st.markdown(f"<{tag}>{css}</{tag}>", unsafe_allow_html=True)

st.markdown('<div class="hero"><h1>📚 বাংলা AI Learning Assistant</h1>'
            '<p>PDF দিন, প্রশ্ন করুন, সহজ বাংলায় উত্তর পান ✨</p></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 🎓 কীভাবে ব্যবহার করবেন")
    st.markdown("**১.** 📄 PDF upload করুন\n\n**২.** ⚡ নিচের বোতাম চাপুন বা প্রশ্ন লিখুন\n\n**৩.** 🤖 বাংলায় উত্তর পান")
    st.markdown("---")
    if st.button("🗑️ Chat মুছে ফেলুন"):
        st.session_state.chat = []
    st.markdown("---")
    st.caption(f"🤖 Model: {MODEL}")
    st.caption("Made with ❤️ by Sanzida Siddike")

if "chat" not in st.session_state:
    st.session_state.chat = []

def ask(q, text):
    r = client.chat.completions.create(model=MODEL, messages=[{"role": "user", "content":
        f"তুমি একজন দয়ালু শিক্ষক। শুধু এই notes থেকে সহজ বাংলায়, heading আর bullet point দিয়ে সুন্দর করে উত্তর দাও:\n{text}\n\nপ্রশ্ন: {q}"}])
    return r.choices[0].message.content

pdf = st.file_uploader("📄 আপনার PDF এখানে দিন", type="pdf")

if pdf:
    pages = PdfReader(pdf).pages
    full = "".join(p.extract_text() or "" for p in pages)
    text = full[:15000]

    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="stat s1"><h2>📑 {len(pages)}</h2>মোট পাতা</div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat s2"><h2>📝 {len(full.split()):,}</h2>মোট শব্দ</div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat s3"><h2>💬 {len(st.session_state.chat)//2}</h2>প্রশ্ন করা হয়েছে</div>', unsafe_allow_html=True)
    st.write("")

    st.markdown("#### ⚡ এক click এ জানুন")
    b1, b2, b3, b4 = st.columns(4)
    quick = None
    if b1.button("📋 সংক্ষেপে বলো"): quick = "এই PDF টা সংক্ষেপে বুঝিয়ে বলো"
    if b2.button("🎯 মূল Point"): quick = "এখান থেকে সবচেয়ে গুরুত্বপূর্ণ ৫টা point দাও"
    if b3.button("❓ MCQ বানাও"): quick = "এই notes থেকে ৫টা MCQ বানাও, শেষে উত্তর সহ"
    if b4.button("💡 সহজ ব্যাখ্যা"): quick = "সবচেয়ে কঠিন বিষয়টা উদাহরণ দিয়ে সহজ করে বুঝিয়ে দাও"

    for role, msg in st.session_state.chat:
        st.chat_message(role, avatar="🙋‍♀️" if role == "user" else "🤖").write(msg)

    q = st.chat_input("✍️ আপনার প্রশ্ন লিখুন...") or quick
    if q:
        st.chat_message("user", avatar="🙋‍♀️").write(q)
        with st.spinner("🤔 AI ভাবছে..."):
            ans = ask(q, text)
        st.chat_message("assistant", avatar="🤖").write(ans)
        st.session_state.chat += [("user", q), ("assistant", ans)]

    if st.session_state.chat:
        log = "\n\n".join(f"{'প্রশ্ন' if r == 'user' else 'উত্তর'}: {m}" for r, m in st.session_state.chat)
        st.download_button("⬇️ পুরো Chat download করুন", log, file_name="chat.txt")
else:
    st.markdown('<div class="tip">👆 শুরু করতে উপরে একটা PDF upload করুন<br><br>'
                '📘 Notes · 📗 Lecture · 📙 Book, যেকোনো কিছু!</div>', unsafe_allow_html=True)

st.markdown('<div class="footer">✨ বাংলা AI Learning Assistant · Built with Python, Streamlit & Groq</div>',
            unsafe_allow_html=True)