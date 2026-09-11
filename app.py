import os
import re
import io
import random
import streamlit as st
from google import genai
from gtts import gTTS

# ============================================================
# ARICE AI — Web Version
# QR → Web App → AI Chatbot
# ============================================================

st.set_page_config(
    page_title="ARICE AI",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------- CSS -------------------------------
st.markdown("""
<style>
.stApp {
    background: radial-gradient(circle at 50% 15%, #062236 0%, #020914 42%, #01060d 100%);
    color: #e8f4f8;
}
.block-container { max-width: 1100px; padding-top: 2rem; }
.arice-title {
    text-align:center; color:#00ffd5; font-size:3rem; font-weight:800;
    letter-spacing:2px; margin-bottom:0;
}
.arice-subtitle {
    text-align:center; color:#6f93a8; font-size:1rem; margin-bottom:2rem;
}
.chat-user {
    background:#0d3b4f; border:1px solid #123049; border-radius:18px;
    padding:12px 16px; margin:10px 0 10px auto; max-width:80%;
}
.chat-ai {
    background:#0a1f33; border:1px solid #123049; border-radius:18px;
    padding:14px 16px; margin:10px 0; max-width:88%;
}
.arice-badge {
    color:#00ffd5; font-weight:700; font-size:.82rem; margin-bottom:5px;
}
div[data-testid="stChatInput"] textarea {
    background:#08192b !important;
    color:#e8f4f8 !important;
}
</style>
""", unsafe_allow_html=True)

# ---------------------- API client ---------------------------
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    st.error("GEMINI_API_KEY is not configured on the server.")
    st.info("Add GEMINI_API_KEY in your hosting platform's Secrets / Environment Variables.")
    st.stop()

client = genai.Client(api_key=api_key)

# ---------------------- Session state ------------------------
if "messages" not in st.session_state:
    st.session_state.messages = []

if "language" not in st.session_state:
    st.session_state.language = "English"

if "last_reply" not in st.session_state:
    st.session_state.last_reply = ""

# ---------------------- Helpers ------------------------------
def detect_language(text):
    if any('\u0A80' <= ch <= '\u0AFF' for ch in text):
        return "Gujarati"
    if any('\u0900' <= ch <= '\u097F' for ch in text):
        return "Hindi"

    gu_words = {"shu","su","kem","kevi","kevu","chhe","che","hatu","hati",
                "maru","mari","mane","tame","aapde","kyare","batavo","bataavo"}
    hi_words = {"kya","kaise","kaun","hai","hain","tha","thi","the","kab",
                "kyu","kyon","mujhe","aap","tum","batao","bataiye","mera","meri"}

    words = re.findall(r"[a-zA-Z]+", text.lower())
    gs = sum(w in gu_words for w in words)
    hs = sum(w in hi_words for w in words)

    if gs > hs and gs:
        return "Gujarati"
    if hs > gs and hs:
        return "Hindi"
    return "English"

def clean_speech(text):
    text = re.sub(r"```.*?```", "", text, flags=re.S)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = re.sub(r"[*_#`]", "", text)
    return text.strip()

def make_voice(text, language):
    lang = {"Gujarati":"gu", "Hindi":"hi", "English":"en"}.get(language, "en")
    audio = io.BytesIO()
    gTTS(text=clean_speech(text) or "...", lang=lang, slow=False).write_to_fp(audio)
    audio.seek(0)
    return audio

def ask_arice(query):
    detected = detect_language(query)
    instructions = {
        "Gujarati": "Answer completely in natural Gujarati. Do not translate into Hindi or English.",
        "Hindi": "Answer completely in natural Hindi. Do not translate into Gujarati or English.",
        "English": "Answer completely in clear, natural English.",
    }

    prompt = f"""You are ARICE, a friendly, professional multilingual personal AI assistant.

User question:
{query}

Detected language: {detected}
{instructions[detected]}

Rules:
- Answer naturally and conversationally.
- Be concise, clear and useful.
- You may use light Markdown.
- If the user writes Romanized Gujarati or Hindi, answer in the same language style when appropriate.
"""
    last_error = None
    for model_name in ("gemini-3.6-flash", "gemini-2.5-flash", "gemini-2.0-flash"):
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            return response.text or "I could not generate a response.", detected
        except Exception as e:
            last_error = e
    raise RuntimeError(f"Gemini request failed: {last_error}")

# -------------------------- Sidebar --------------------------
with st.sidebar:
    st.markdown("## ⚡ ARICE AI")
    st.caption("Your Intelligent Personal Assistant")
    st.divider()

    language = st.selectbox(
        "🌐 Language",
        ["English", "Gujarati", "Hindi"],
        index=["English","Gujarati","Hindi"].index(st.session_state.language)
    )
    st.session_state.language = language

    if st.button("＋ New Chat", use_container_width=True):
        st.session_state.messages = []
        st.session_state.last_reply = ""
        st.rerun()

    st.divider()
    st.caption("ARICE AI Web Edition")
    st.caption("Voice input uses your browser microphone.")
    st.caption("Voice output is generated in the selected language.")

# --------------------------- Header ---------------------------
st.markdown('<div class="arice-title">⚡ ARICE AI</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="arice-subtitle">Your Intelligent Personal AI Assistant</div>',
    unsafe_allow_html=True
)

# ----------------------- Welcome screen ----------------------
if not st.session_state.messages:
    st.markdown(
        '<div style="text-align:center;font-size:5rem;margin-top:1rem;">◉</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div style="text-align:center;color:#6f93a8;">Hello! How can I help you today?</div>',
        unsafe_allow_html=True
    )

# ------------------------ Chat history -----------------------
for msg in st.session_state.messages:
    if msg["role"] == "user":
        st.markdown(
            f'<div class="chat-user"><b>You</b><br>{msg["content"]}</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<div class="chat-ai"><div class="arice-badge">🤖 ARICE</div>{msg["content"]}</div>',
            unsafe_allow_html=True
        )

# ------------------------- Voice input -----------------------
audio_value = st.audio_input("🎙️ Speak to ARICE")

if audio_value is not None:
    st.info("Voice recording received. Browser audio transcription is not enabled in this first web build; use the text box below for the most reliable input.")

# ------------------------- Text input ------------------------
query = st.chat_input("Ask ARICE anything…")

if query:
    st.session_state.messages.append({"role":"user", "content":query})

    with st.spinner("ARICE is thinking…"):
        try:
            reply, detected = ask_arice(query)
            st.session_state.messages.append({"role":"assistant", "content":reply})
            st.session_state.last_reply = reply

            st.rerun()
        except Exception as e:
            st.session_state.messages.append({
                "role":"assistant",
                "content":f"⚠️ Sorry, I could not reach Gemini. {e}"
            })
            st.rerun()

# -------------------------- Voice reply ----------------------
if st.session_state.last_reply:
    with st.expander("🔊 Listen to last ARICE answer"):
        try:
            audio = make_voice(st.session_state.last_reply, detect_language(st.session_state.last_reply))
            st.audio(audio, format="audio/mp3")
        except Exception as e:
            st.caption(f"Voice generation unavailable: {e}")
