# ==========================================================================
# ARICE AI — Premium Futuristic Personal AI Assistant
# --------------------------------------------------------------------------
# Rebuilt UI on top of the original ARICE architecture. See the chat
# response for a summary of what changed. Section map:
#   1. Imports
#   2. Configuration / Theme constants
#   3. Gemini AI client
#   4. Global state
#   5. Voice / TTS
#   6. Speech recognition
#   7. Language detection
#   8. Chat history (local JSON)
#   9. Memory (local JSON)
#  10. Gemini request logic
#  11. Command routing / voice + typed input
#  12. UI helpers (scrollable frames, markdown-lite renderer, chat bubbles)
#  13. Animated AI core canvas
#  14. Main window construction (sidebar, chat area, input bar)
#  15. Tools / language / memory / settings popovers
#  16. Theme application
#  17. Event bindings
#  18. Application startup
# ==========================================================================

# ---------------------------------------------------------------------
# 1. IMPORTS
# ---------------------------------------------------------------------
import tkinter as tk
import threading
import datetime
import webbrowser
import speech_recognition as sr
import win32com.client
import winsound
import pyautogui
from gtts import gTTS
import pygame
import os
import time
import json
import uuid
import re
import math
import random
from pathlib import Path

from google import genai

# ---------------------------------------------------------------------
# GEMINI API KEY — kept exactly as before (environment variable only)
# ---------------------------------------------------------------------
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY environment variable is not set.")

client = genai.Client(api_key=GEMINI_API_KEY)

# ---------------------------------------------------------------------
# 2. CONFIGURATION / THEME CONSTANTS
# ---------------------------------------------------------------------
FONT_FAMILY = "Segoe UI"
MONO_FAMILY = "Consolas"

THEMES = {
    "dark": {
        "bg_primary": "#020914",
        "bg_secondary": "#050f1c",
        "sidebar_bg": "#04101c",
        "sidebar_card": "#0a1f33",
        "main_bg": "#020914",
        "card_bg": "#0a1f33",
        "bubble_user_bg": "#0d3b4f",
        "bubble_user_fg": "#eaffff",
        "bubble_ai_bg": "#0a1f33",
        "bubble_ai_fg": "#e8f4f8",
        "input_bg": "#08192b",
        "accent": "#00ffd5",
        "accent_dim": "#0d5f6e",
        "text_primary": "#e8f4f8",
        "text_secondary": "#6f93a8",
        "danger": "#ff6b6b",
        "border": "#123049",
    },
    "light": {
        "bg_primary": "#f4f7fa",
        "bg_secondary": "#ffffff",
        "sidebar_bg": "#eef2f6",
        "sidebar_card": "#ffffff",
        "main_bg": "#f4f7fa",
        "card_bg": "#ffffff",
        "bubble_user_bg": "#d6f5ef",
        "bubble_user_fg": "#053b3a",
        "bubble_ai_bg": "#ffffff",
        "bubble_ai_fg": "#1a2733",
        "input_bg": "#ffffff",
        "accent": "#00a389",
        "accent_dim": "#8fd8cb",
        "text_primary": "#1a2733",
        "text_secondary": "#5c7386",
        "danger": "#d64545",
        "border": "#d7e1e9",
    },
}

current_theme = "dark"
THEME = dict(THEMES[current_theme])

# ---------------------------------------------------------------------
# 4. GLOBAL STATE
# ---------------------------------------------------------------------
last_reply = ""
current_voice_index = 0
thinking_active = False
thinking_frame = None
thinking_label = None
ai_core_canvases = []          # every AICoreCanvas instance currently alive
used_followups = set()

BASE_DIR = Path(__file__).resolve().parent
HISTORY_DIR = BASE_DIR / "arice_history"
HISTORY_DIR.mkdir(exist_ok=True)
MEMORY_FILE = BASE_DIR / "arice_memory.json"

current_chat_id = None
current_chat_title = "New Chat"
chat_messages = []

current_language = "English"

# Voice input state. None means "use Windows default microphone".
MIC_DEVICE_INDEX = 3
voice_listening_active = False
voice_lock = threading.Lock()

language_codes = {
    "Gujarati": "gu-IN",
    "Hindi": "hi-IN",
    "English": "en-IN"
}

language_prompts = {
    "Gujarati": "ગુજરાતીમાં જવાબ આપો. સરળ અને સ્વાભાવિક ગુજરાતી ભાષામાં સમજાવો.",
    "Hindi": "हिंदी में जवाब दें। सरल और प्राकृतिक हिंदी में समझाएं।",
    "English": "Answer in clear and natural English."
}

# widgets that get wired up during UI construction (declared here so
# every function can reference them; they are assigned real values in
# the "MAIN WINDOW CONSTRUCTION" section near the bottom of the file).
root = None
sidebar = None
chat_list_container = None
chat_list_canvas = None
main_area = None
content_area = None
welcome_frame = None
messages_outer = None
messages_canvas = None
messages_inner = None
input_entry = None
status_label = None
title_label = None
ai_status_dot = None
ai_status_text = None
mic_btn = None
send_btn = None
tools_btn = None
history_window = None
history_listbox = None  # kept only so any stray references never explode


# ---------------------------------------------------------------------
# TEXT CLEANUP (used for speech only — display keeps Markdown)
# ---------------------------------------------------------------------
def clean_text_for_speech(text):
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'__(.*?)__', r'\1', text)
    text = re.sub(r'_(.*?)_', r'\1', text)
    text = re.sub(r'#+\s*', '', text)
    text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
    text = re.sub(r'`([^`]*)`', r'\1', text)
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    text = re.sub(r'^\s*[-•]\s*', '', text, flags=re.MULTILINE)
    return text.strip()


# ---------------------------------------------------------------------
# 5. VOICE / TTS
# ---------------------------------------------------------------------
speech_lock = threading.Lock()


def say(text, language=None):
    """Speak `text` aloud. Always runs the heavy work off the caller's
    thread is the caller's responsibility — this function itself blocks
    until playback finishes, so call it from a background thread."""
    global last_reply
    last_reply = text
    speech_text = clean_text_for_speech(text)

    try:
        if language is None:
            language = current_language

        language_map = {"Gujarati": "gu", "Hindi": "hi", "English": "en"}
        lang = language_map.get(language, "en")

        filename = str(BASE_DIR / "arice_voice.mp3")

        try:
            if pygame.mixer.get_init():
                pygame.mixer.music.stop()
                pygame.mixer.quit()
        except Exception:
            pass

        try:
            if os.path.exists(filename):
                os.remove(filename)
        except PermissionError:
            time.sleep(0.3)
            try:
                os.remove(filename)
            except Exception:
                pass

        root.after(0, lambda: set_ai_status("SPEAKING...", online=True))

        tts = gTTS(text=speech_text or "...", lang=lang, slow=False)
        tts.save(filename)

        pygame.mixer.init()
        pygame.mixer.music.load(filename)
        pygame.mixer.music.play()

        while pygame.mixer.music.get_busy():
            time.sleep(0.1)

        pygame.mixer.music.stop()
        pygame.mixer.quit()

        try:
            os.remove(filename)
        except Exception:
            pass

    except Exception as e:
        print("🔊 Voice Error:", e)
    finally:
        root.after(0, lambda: set_ai_status("ARICE ONLINE", online=True))


def stop_speaking():
    play_click()
    try:
        if pygame.mixer.get_init():
            pygame.mixer.music.stop()
            pygame.mixer.quit()
        set_status("🔇 Voice Stopped")
        set_ai_status("ARICE ONLINE", online=True)
        log("🔇 Arice stopped speaking.")
    except Exception as e:
        print("Stop Voice Error:", e)


def play_click():
    try:
        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
    except Exception:
        pass


# ---------------------------------------------------------------------
# 6. SPEECH RECOGNITION
# ---------------------------------------------------------------------
def list_microphones():
    """Return available microphone names for troubleshooting/selection."""
    try:
        return list(sr.Microphone.list_microphone_names())
    except Exception as e:
        print("Microphone list error:", e)
        return []


def takeCommand():
    """Capture speech in a worker thread and return recognized text."""
    global voice_listening_active

    with voice_lock:
        if voice_listening_active:
            set_status("🎧 Already listening...")
            return ""
        voice_listening_active = True

    try:
        r = sr.Recognizer()
        r.dynamic_energy_threshold = False
        r.pause_threshold = 1.5
        r.non_speaking_duration = 0.7

        set_status("🎧 Listening… speak now")
        set_ai_status("LISTENING...", online=True)
        set_mic_listening(True)

        # ---------------- MICROPHONE ----------------
        try:
            mic_kwargs = {}

            if MIC_DEVICE_INDEX is not None:
                mic_kwargs["device_index"] = MIC_DEVICE_INDEX

            print("MIC LIST:")
            for i, name in enumerate(sr.Microphone.list_microphone_names()):
                print(i, name)

            with sr.Microphone(**mic_kwargs) as source:
                r.adjust_for_ambient_noise(source, duration=0.3)
                r.energy_threshold = 50
                audio = r.listen(
                    source,
                    timeout=20,
                    phrase_time_limit=15
                )

        except sr.WaitTimeoutError:
            set_status("⌛ No speech detected — tap the mic and try again")
            return ""

        except OSError as e:
            set_status("❌ Microphone device unavailable")
            print("Microphone OSError:", e)
            return ""

        except AttributeError as e:
            set_status("❌ PyAudio / microphone driver missing")
            print("Microphone dependency error:", e)
            return ""

        except Exception as e:
            set_status("❌ Microphone error")
            print("Microphone error:", repr(e))
            return ""

        # ---------------- RECOGNITION ----------------
        set_status("🧠 Recognizing...")
        set_ai_status("THINKING...", online=True)

        try:
            command = r.recognize_google(audio, language="en-IN")

        except sr.UnknownValueError:
            set_status("🤷 I couldn't understand that")
            return ""

        except sr.RequestError as e:
            set_status("🌐 Speech recognition network error")
            print("Speech recognition RequestError:", e)
            return ""

        except Exception as e:
            set_status("❌ Speech recognition failed")
            print("SPEECH ERROR:", repr(e))
            return ""

        command = command.strip()

        if command:
            set_status("✅ Voice captured")
            return command

        set_status("⌛ No words detected")
        return ""

    finally:
        voice_listening_active = False
        set_mic_listening(False)
        set_ai_status("ARICE ONLINE", online=True)


# Keep the original function name for compatibility.


# ---------------------------------------------------------------------
# 7. LANGUAGE DETECTION
# ---------------------------------------------------------------------
def detect_question_language(text):
    if any('\u0A80' <= ch <= '\u0AFF' for ch in text):
        return "Gujarati"
    if any('\u0900' <= ch <= '\u097F' for ch in text):
        return "Hindi"

    hindi_words = [
        "kya", "kaise", "kaun", "kon", "hai", "hain", "tha", "thi", "the",
        "kab", "kyu", "kyon", "mujhe", "aap", "tum", "batao", "bataiye",
        "ke", "ki", "ka", "mein", "mera", "meri"
    ]
    gujarati_words = [
        "shu", "su", "kem", "kevi", "kevu", "kon", "chhe", "che", "hatu",
        "hati", "hata", "maru", "mari", "mane", "tame", "aapde",
        "kyaare", "kyare", "bataavo", "batavo"
    ]

    words = text.lower().split()
    hindi_score = sum(1 for w in words if w in hindi_words)
    gujarati_score = sum(1 for w in words if w in gujarati_words)

    if hindi_score > gujarati_score and hindi_score > 0:
        return "Hindi"
    if gujarati_score > hindi_score and gujarati_score > 0:
        return "Gujarati"
    return "English"


# ---------------------------------------------------------------------
# 8. CHAT HISTORY (local JSON) — architecture preserved from original
# ---------------------------------------------------------------------
def _history_file(chat_id):
    return HISTORY_DIR / f"{chat_id}.json"


def elide_title(text, maxlen=30):
    text = " ".join(text.strip().split())
    return text if len(text) <= maxlen else text[:maxlen - 1].rstrip() + "…"


def create_new_chat():
    global current_chat_id, current_chat_title, chat_messages

    current_chat_id = uuid.uuid4().hex
    current_chat_title = "New Chat"
    chat_messages = []

    save_current_chat()
    refresh_history_list()
    clear_chat_box()
    show_welcome_view()
    set_status("🆕 New chat started")
    log("🆕 New chat started.")


def save_current_chat():
    if not current_chat_id:
        return
    data = {
        "id": current_chat_id,
        "title": current_chat_title,
        "updated": datetime.datetime.now().isoformat(),
        "messages": chat_messages
    }
    try:
        _history_file(current_chat_id).write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        print("History save error:", e)


def load_chat(chat_id):
    global current_chat_id, current_chat_title, chat_messages

    try:
        data = json.loads(_history_file(chat_id).read_text(encoding="utf-8"))
        current_chat_id = data.get("id", chat_id)
        current_chat_title = data.get("title", "Chat")
        chat_messages = data.get("messages", [])

        clear_chat_box()

        if chat_messages:
            show_chat_view()
            for item in chat_messages:
                add_message_bubble(item.get("role", ""), item.get("text", ""),
                                    animate_scroll=False)
        else:
            show_welcome_view()

        set_status("📂 Chat history loaded")
        refresh_history_list()
        scroll_messages_to_bottom()

    except Exception as e:
        print("History load error:", e)
        set_status("❌ Could not load history")


def clear_chat_box():
    try:
        for w in messages_inner.winfo_children():
            w.destroy()
        messages_canvas.update_idletasks()
        messages_canvas.configure(scrollregion=messages_canvas.bbox("all"))
    except Exception:
        pass


def record_message(role, text):
    global current_chat_title

    if not current_chat_id:
        create_new_chat()

    chat_messages.append({
        "role": role,
        "text": text,
        "time": datetime.datetime.now().isoformat()
    })

    if role == "user" and current_chat_title == "New Chat":
        current_chat_title = elide_title(text, 40) or "New Chat"

    save_current_chat()
    root.after(0, refresh_history_list)
    root.after(0, lambda: add_message_bubble(role, text))


def delete_chat(chat_id):
    try:
        _history_file(chat_id).unlink(missing_ok=True)
    except Exception as e:
        print("History delete error:", e)

    if current_chat_id == chat_id:
        create_new_chat()
    else:
        refresh_history_list()


def refresh_history_list():
    """Rebuilds the sidebar's TODAY / YESTERDAY / PREVIOUS CHATS list."""
    if chat_list_container is None:
        return

    try:
        for w in chat_list_container.winfo_children():
            w.destroy()

        files = sorted(HISTORY_DIR.glob("*.json"),
                        key=lambda p: p.stat().st_mtime, reverse=True)

        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        groups = {"TODAY": [], "YESTERDAY": [], "PREVIOUS CHATS": []}

        for f in files:
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
            except Exception:
                continue
            try:
                updated = datetime.datetime.fromisoformat(
                    data.get("updated", "")).date()
            except Exception:
                updated = today

            title = data.get("title") or "New Chat"
            entry = (data.get("id", f.stem), title)

            if updated == today:
                groups["TODAY"].append(entry)
            elif updated == yesterday:
                groups["YESTERDAY"].append(entry)
            else:
                groups["PREVIOUS CHATS"].append(entry)

        any_chat = False
        for label_text, entries in groups.items():
            if not entries:
                continue
            any_chat = True
            tk.Label(
                chat_list_container, text=label_text,
                font=(FONT_FAMILY, 9, "bold"),
                fg=THEME["text_secondary"], bg=THEME["sidebar_bg"], anchor="w"
            ).pack(fill="x", padx=14, pady=(12, 3))

            for chat_id, chat_title in entries:
                _build_history_row(chat_list_container, chat_id, chat_title)

        if not any_chat:
            tk.Label(
                chat_list_container, text="No chats yet.\nStart a new one!",
                font=(FONT_FAMILY, 9), fg=THEME["text_secondary"],
                bg=THEME["sidebar_bg"], justify="left"
            ).pack(padx=14, pady=12, anchor="w")

        chat_list_container.update_idletasks()
        chat_list_canvas.configure(scrollregion=chat_list_canvas.bbox("all"))
    except Exception as e:
        print("Sidebar refresh error:", e)


def _build_history_row(parent, chat_id, chat_title):
    is_active = (chat_id == current_chat_id)
    row = tk.Frame(parent, bg=THEME["sidebar_card"] if is_active else THEME["sidebar_bg"])
    row.pack(fill="x", padx=8, pady=1)

    icon = "🟢" if is_active else "💬"
    lbl = tk.Label(
        row, text=f"{icon}  {elide_title(chat_title, 26)}",
        font=(FONT_FAMILY, 10, "bold" if is_active else "normal"),
        fg=THEME["accent"] if is_active else THEME["text_primary"],
        bg=row["bg"], anchor="w", padx=8, pady=7, cursor="hand2"
    )
    lbl.pack(side="left", fill="x", expand=True)
    lbl.bind("<Button-1>", lambda e, cid=chat_id: load_chat(cid))

    del_lbl = tk.Label(
        row, text="✕", font=(FONT_FAMILY, 9), fg=THEME["text_secondary"],
        bg=row["bg"], cursor="hand2", padx=8
    )
    del_lbl.pack(side="right")
    del_lbl.bind("<Button-1>", lambda e, cid=chat_id: delete_chat(cid))


def open_history_window():
    """Superseded by the always-visible sidebar. Kept for compatibility
    with anything that still calls it."""
    refresh_history_list()


# ---------------------------------------------------------------------
# 9. MEMORY (local JSON, lightweight, optional & user-managed)
# ---------------------------------------------------------------------
DEFAULT_MEMORY = {"preferences": [], "projects": [], "language": [], "interests": []}
MEMORY_CATEGORY_ICONS = {
    "preferences": "👤 Preferences",
    "projects": "💻 Projects",
    "language": "🌐 Language",
    "interests": "🎯 Interests",
}


def load_memory():
    try:
        if MEMORY_FILE.exists():
            data = json.loads(MEMORY_FILE.read_text(encoding="utf-8"))
            for k in DEFAULT_MEMORY:
                data.setdefault(k, [])
            return data
    except Exception as e:
        print("Memory load error:", e)
    return dict(DEFAULT_MEMORY)


def save_memory(data):
    try:
        MEMORY_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                encoding="utf-8")
    except Exception as e:
        print("Memory save error:", e)


def memory_add(category, text):
    if not text.strip():
        return
    data = load_memory()
    data.setdefault(category, []).append(text.strip())
    save_memory(data)


def memory_remove(category, index):
    data = load_memory()
    items = data.get(category, [])
    if 0 <= index < len(items):
        items.pop(index)
    save_memory(data)


# ---------------------------------------------------------------------
# 10. GEMINI REQUEST LOGIC
# ---------------------------------------------------------------------
def ask_gemini(query=None):
    play_click()

    if not query:
        try:
            typed = input_entry.get().strip()
        except Exception:
            typed = ""
        query = typed or "Tell me something interesting"
        if typed:
            root.after(0, lambda: input_entry.delete(0, tk.END))

    detected_language = detect_question_language(query)

    if not current_chat_id:
        create_new_chat()

    record_message("user", query)
    root.after(0, show_thinking)

    set_status(f"🧠 Thinking - {detected_language}")
    set_ai_status("THINKING...", online=True)

    try:
        language_instructions = {
            "Gujarati": "Answer completely in natural Gujarati. Do not translate into Hindi or English.",
            "Hindi": "Answer completely in natural Hindi. Do not translate into Gujarati or English.",
            "English": "Answer completely in clear, natural English.",
        }

        prompt = f"""You are Arice, a friendly, professional multilingual personal AI assistant.

The user's question is:
{query}

The detected language is: {detected_language}
{language_instructions[detected_language]}

Formatting rules:
- You MAY use light Markdown: **bold**, `inline code`, bullet points with "-",
  numbered lists, short headings with "#", and fenced code blocks using
  ```language ... ``` for any code.
- Keep explanations concise, clear and well structured.
- Only use a code block when actually showing code or commands.
- Reply naturally and conversationally, like a helpful expert assistant.
- Always answer in the SAME language the user asked in (English, Hindi or
  Gujarati, including romanized Hindi/Gujarati typed in English letters).
"""

        models_to_try = ["gemini-3.6-flash", "gemini-3.5-flash-lite"]
        response = None
        last_error = None

        for model_name in models_to_try:
            for attempt in range(3):
                try:
                    response = client.models.generate_content(
                        model=model_name, contents=prompt
                    )
                    break
                except Exception as e:
                    last_error = e
                    error_text = str(e)
                    is_retryable = any(code in error_text for code in
                                        ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"))
                    if not is_retryable:
                        raise
                    if attempt < 2:
                        delay = 2 * (2 ** attempt)  # 2s, 4s
                        set_status(f"⏳ Gemini busy - retry {attempt + 1}/2")
                        time.sleep(delay)
            if response is not None:
                break

        if response is None:
            raise RuntimeError(f"Gemini unavailable after retries: {last_error}")

        reply = response.text or "I could not generate a response."

        root.after(0, hide_thinking)
        record_message("assistant", reply)
        root.after(0, lambda: add_followups(query, reply))

        # Speak every normal Gemini answer automatically.
        # Without this call ARICE displays the answer but stays silent.
        threading.Thread(
            target=lambda: say(reply, detected_language),
            daemon=True,
            name="AriceTTS"
        ).start()

        set_status(f"✅ Ready - {detected_language}")
        set_ai_status("GENERATING...", online=True)

    except Exception as e:
        root.after(0, hide_thinking)
        error_message = str(e)
        log("❌ GEMINI ERROR:")
        log(error_message)
        set_status("❌ Gemini Error")
        set_ai_status("OFFLINE / ERROR", online=False)
        root.after(0, lambda: add_message_bubble(
            "assistant", "⚠️ Sorry, I ran into a problem reaching Gemini. Please try again."))
        print("\n========== GEMINI ERROR ==========")
        print(error_message)
        print("==================================\n")


# ---------------------------------------------------------------------
# 11. COMMAND ROUTING / VOICE + TYPED INPUT
# ---------------------------------------------------------------------
def open_website(name):
    play_click()
    urls = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com",
        "github": "https://www.github.com",
        "stackoverflow": "https://stackoverflow.com",
        "wikipedia": "https://www.wikipedia.org"
    }
    if name in urls:
        say(f"Opening {name}")
        webbrowser.open(urls[name])
        log(f"🔗 Opened {name}")
    else:
        say("Website not configured.")


def tell_time():
    play_click()
    now = datetime.datetime.now().strftime("%I:%M %p")
    say(f"The time is {now}")
    log(f"🕒 Time: {now}")


def select_language(language):
    global current_language

    try:
        if pygame.mixer.get_init():
            pygame.mixer.music.stop()
            pygame.mixer.quit()
    except Exception:
        pass

    current_language = language

    if language == "Gujarati":
        message = "હવે હું ગુજરાતી ભાષામાં જવાબ આપીશ."
    elif language == "Hindi":
        message = "अब मैं हिंदी में जवाब दूंगा."
    else:
        message = "I will now respond in English."

    log(f"🌐 Language changed to: {language}")
    set_status(f"🌐 {language}")
    threading.Thread(target=lambda: say(message), daemon=True).start()


def repeat_last():
    play_click()
    if last_reply:
        threading.Thread(target=lambda: say(last_reply), daemon=True).start()
    else:
        threading.Thread(target=lambda: say("No response to repeat yet."), daemon=True).start()


def take_screenshot():
    play_click()
    filename = f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    try:
        pyautogui.screenshot(filename)
        say(f"Screenshot saved as {filename}")
        log(f"📸 Screenshot taken: {filename}")
    except Exception as e:
        log(f"❌ Screenshot failed: {e}")


def toggle_voice():
    global current_voice_index
    current_voice_index = 1 - current_voice_index
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        voice_name = speaker.GetVoices().Item(current_voice_index).GetDescription()
    except Exception:
        voice_name = "default"
    say(f"Voice switched to {voice_name}")
    log(f"🗣 Voice changed to: {voice_name}")


def start_voice_thread():
    play_click()
    if voice_listening_active:
        set_status("🎧 Already listening…")
        return
    threading.Thread(target=run_voice, daemon=True, name="AriceVoiceInput").start()


def run_voice():
    query = takeCommand()
    if query:
        # Return to Tk's main thread before touching widgets. Then use the
        # exact same send pipeline as typed input, so voice and keyboard
        # requests behave identically.
        root.after(0, lambda q=query: send_typed_message(q))


def send_typed_message(preset_text=None):
    play_click()
    query = preset_text if preset_text is not None else input_entry.get().strip()
    if query:
        input_entry.delete(0, tk.END)
        threading.Thread(target=lambda: handle_query(query), daemon=True).start()


def handle_enter(event=None):
    send_typed_message()
    return "break"


def handle_query(query):
    q = query.lower()
    if "youtube" in q:
        open_website("youtube")
    elif "google" in q:
        open_website("google")
    elif "github" in q:
        open_website("github")
    elif "stackoverflow" in q:
        open_website("stackoverflow")
    elif "wikipedia" in q:
        open_website("wikipedia")
    elif "time" in q and "sometime" not in q:
        tell_time()
    elif "screenshot" in q:
        take_screenshot()
    elif "change voice" in q or "switch voice" in q:
        toggle_voice()
    elif any(x in q for x in ["exit", "quit", "close arice"]):
        say("Goodbye.")
        root.after(500, root.quit)
    else:
        ask_gemini(query)


def log(text):
    """Lightweight debug/status logger — console + status bar only.
    Actual conversation turns are rendered as chat bubbles via
    record_message(), not through this function."""
    print(text)
    set_status(str(text)[:70])


def set_status(text):
    if root is None:
        return
    root.after(0, lambda: status_label.config(text=text))


def set_ai_status(text, online=True):
    def update():
        ai_status_text.config(text=text)
        ai_status_dot.config(fg=THEME["accent"] if online else THEME["danger"])
        state_map = {
            "LISTENING...": "listening",
            "THINKING...": "thinking",
            "GENERATING...": "generating",
            "SPEAKING...": "speaking",
        }
        state = state_map.get(text, "idle" if online else "offline")
        for canvas in list(ai_core_canvases):
            try:
                canvas.set_state(state)
            except tk.TclError:
                pass
    root.after(0, update)


def set_mic_listening(is_listening):
    def update():
        try:
            mic_btn.config(
                text="🔴" if is_listening else "🎙",
                fg=THEME["danger"] if is_listening else THEME["accent"]
            )
        except Exception:
            pass
    root.after(0, update)


# ---------------------------------------------------------------------
# 12. UI HELPERS — scrollable frames, markdown-lite renderer, bubbles
# ---------------------------------------------------------------------
def make_scrollable(parent, bg):
    """Returns (outer_frame, canvas, inner_frame). inner_frame is the
    widget you pack children into; the canvas scrolls it vertically."""
    outer = tk.Frame(parent, bg=bg)
    canvas = tk.Canvas(outer, bg=bg, highlightthickness=0)
    scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=bg)

    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_wheel(_):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _unbind_wheel(_):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _bind_wheel)
    canvas.bind("<Leave>", _unbind_wheel)

    return outer, canvas, inner


_INLINE_PATTERN = re.compile(r'\*\*(.+?)\*\*|`([^`]+)`')


def _insert_inline(text_widget, s):
    pos = 0
    for m in _INLINE_PATTERN.finditer(s):
        if m.start() > pos:
            text_widget.insert("end", s[pos:m.start()])
        if m.group(1) is not None:
            text_widget.insert("end", m.group(1), ("bold",))
        else:
            text_widget.insert("end", m.group(2), ("inlinecode",))
        pos = m.end()
    text_widget.insert("end", s[pos:])


def _configure_text_tags(txt, fg, accent):
    txt.tag_configure("bold", font=(FONT_FAMILY, 11, "bold"))
    txt.tag_configure("heading", font=(FONT_FAMILY, 14, "bold"), foreground=accent,
                       spacing1=4, spacing3=4)
    txt.tag_configure("inlinecode", font=(MONO_FAMILY, 10), foreground=accent,
                       background=THEME["input_bg"])


def render_text_block(parent, content, fg, bg, width_chars=62):
    """Render markdown-lite text without creating oversized empty bubbles."""
    txt = tk.Text(
        parent, wrap="word", width=width_chars, height=1,
        bg=bg, fg=fg, font=(FONT_FAMILY, 11),
        bd=0, highlightthickness=0, padx=2, pady=1, cursor="arrow"
    )
    _configure_text_tags(txt, fg, THEME["accent"])

    lines = content.strip("\n").split("\n")
    for line in lines:
        stripped = line.rstrip()
        if not stripped.strip():
            txt.insert("end", "\n")
            continue

        m_head = re.match(r'^(#{1,4})\s+(.*)', stripped)
        m_bullet = re.match(r'^[-*]\s+(.*)', stripped)
        m_num = re.match(r'^(\d+)\.\s+(.*)', stripped)

        if m_head:
            txt.insert("end", m_head.group(2) + "\n", ("heading",))
        elif m_bullet:
            txt.insert("end", "•  ")
            _insert_inline(txt, m_bullet.group(1))
            txt.insert("end", "\n")
        elif m_num:
            txt.insert("end", f"{m_num.group(1)}.  ")
            _insert_inline(txt, m_num.group(2))
            txt.insert("end", "\n")
        else:
            _insert_inline(txt, stripped)
            txt.insert("end", "\n")

    txt.configure(state="disabled")
    txt.update_idletasks()

    try:
        display_lines = int(txt.count("1.0", "end-1c", "displaylines")[0])
    except Exception:
        display_lines = len(lines)

    # Hard cap prevents a malformed/wide Tk Text calculation from producing
    # the huge blank boxes visible in the previous UI.
    txt.configure(height=max(1, min(display_lines, 24)))
    txt.pack(fill="x", padx=2, pady=0)
    return txt


def render_code_block(parent, lang, code):
    holder = tk.Frame(parent, bg=THEME["input_bg"], highlightthickness=1,
                       highlightbackground=THEME["border"])
    holder.pack(fill="x", padx=2, pady=6)

    header = tk.Frame(holder, bg="#04101c" if current_theme == "dark" else "#eef2f6")
    header.pack(fill="x")

    tk.Label(
        header, text=(lang or "code").upper(), font=(MONO_FAMILY, 9, "bold"),
        fg=THEME["accent"], bg=header["bg"], padx=10, pady=5
    ).pack(side="left")

    copy_btn = tk.Button(
        header, text="📋 Copy", font=(FONT_FAMILY, 9), bg=header["bg"],
        fg=THEME["text_secondary"], bd=0, activebackground=header["bg"],
        cursor="hand2", padx=8
    )
    copy_btn.pack(side="right")

    def do_copy():
        try:
            root.clipboard_clear()
            root.clipboard_append(code)
            root.update()
        except Exception:
            pass
        copy_btn.config(text="✅ Copied!")
        copy_btn.after(1500, lambda: copy_btn.config(text="📋 Copy"))

    copy_btn.config(command=do_copy)

    code_txt = tk.Text(
        holder, wrap="char", width=60, bg=THEME["input_bg"],
        fg="#d7fff3" if current_theme == "dark" else "#0a3d33",
        font=(MONO_FAMILY, 10), bd=0, highlightthickness=0,
        padx=10, pady=8, cursor="arrow"
    )
    code_txt.insert("end", code.strip("\n"))
    code_txt.configure(state="disabled")
    code_txt.update_idletasks()
    try:
        lines_count = int(code_txt.count("1.0", "end", "displaylines")[0])
    except Exception:
        lines_count = int(code_txt.index("end-1c").split(".")[0])
    code_txt.configure(height=max(1, min(lines_count, 30)))
    code_txt.pack(fill="x")
    return holder


_CODE_FENCE_PATTERN = re.compile(r'```(\w+)?\n?(.*?)```', re.DOTALL)


def render_message_content(parent, content):
    segments = []
    last_end = 0
    for m in _CODE_FENCE_PATTERN.finditer(content):
        if m.start() > last_end:
            segments.append(("text", content[last_end:m.start()]))
        segments.append(("code", m.group(1) or "code", m.group(2)))
        last_end = m.end()
    if last_end < len(content):
        segments.append(("text", content[last_end:]))
    if not segments:
        segments = [("text", content)]

    for seg in segments:
        if seg[0] == "text":
            if seg[1].strip():
                render_text_block(parent, seg[1], fg=parent.text_fg, bg=parent.bubble_bg)
        else:
            render_code_block(parent, seg[1], seg[2])


def add_message_bubble(role, text, animate_scroll=True):
    if welcome_frame.winfo_ismapped():
        show_chat_view()

    is_user = (role == "user")
    row = tk.Frame(messages_inner, bg=THEME["main_bg"])
    row.pack(fill="x", padx=28, pady=(8, 3))

    align_wrap = tk.Frame(row, bg=THEME["main_bg"])
    align_wrap.pack(anchor="e" if is_user else "w")

    bubble_bg = THEME["bubble_user_bg"] if is_user else THEME["bubble_ai_bg"]
    bubble_fg = THEME["bubble_user_fg"] if is_user else THEME["bubble_ai_fg"]

    # User messages stay compact like ChatGPT instead of creating a full-width
    # empty Text widget.
    if is_user:
        bubble = tk.Frame(
            align_wrap, bg=bubble_bg, highlightthickness=1,
            highlightbackground=THEME["border"]
        )
        bubble.pack(anchor="e", pady=(0, 0))

        user_label = tk.Label(
            bubble, text=text.strip(), font=(FONT_FAMILY, 11),
            fg=bubble_fg, bg=bubble_bg, justify="left", anchor="w",
            wraplength=620, padx=14, pady=9
        )
        user_label.pack()
    else:
        tk.Label(
            align_wrap, text="🤖  ARICE",
            font=(FONT_FAMILY, 9, "bold"),
            fg=THEME["text_secondary"], bg=THEME["main_bg"]
        ).pack(anchor="w", padx=4, pady=(0, 2))

        bubble = tk.Frame(
            align_wrap, bg=bubble_bg, highlightthickness=1,
            highlightbackground=THEME["border"]
        )
        bubble.pack(anchor="w")

        bubble.text_fg = bubble_fg
        bubble.bubble_bg = bubble_bg

        inner = tk.Frame(bubble, bg=bubble_bg)
        inner.pack(padx=12, pady=9)
        inner.text_fg = bubble_fg
        inner.bubble_bg = bubble_bg

        render_message_content(inner, text)

    if animate_scroll:
        scroll_messages_to_bottom()


def scroll_messages_to_bottom():
    def _do():
        messages_canvas.update_idletasks()
        messages_canvas.configure(scrollregion=messages_canvas.bbox("all"))
        messages_canvas.yview_moveto(1.0)
    root.after(50, _do)


# ---- thinking indicator -------------------------------------------------
def show_thinking():
    global thinking_active, thinking_frame, thinking_label

    if welcome_frame.winfo_ismapped():
        show_chat_view()

    thinking_active = True
    row = tk.Frame(messages_inner, bg=THEME["main_bg"])
    row.pack(fill="x", padx=16, pady=6, anchor="w")

    wrap = tk.Frame(row, bg=THEME["main_bg"])
    wrap.pack(anchor="w")

    tk.Label(wrap, text="🤖 ARICE", font=(FONT_FAMILY, 9, "bold"),
             fg=THEME["text_secondary"], bg=THEME["main_bg"]).pack(anchor="w", padx=6)

    bubble = tk.Frame(wrap, bg=THEME["bubble_ai_bg"], highlightthickness=1,
                       highlightbackground=THEME["border"])
    bubble.pack(anchor="w", pady=(2, 0))

    thinking_label = tk.Label(
        bubble, text="ARICE is thinking.", font=(FONT_FAMILY, 11, "italic"),
        fg=THEME["text_secondary"], bg=THEME["bubble_ai_bg"], padx=14, pady=10
    )
    thinking_label.pack()
    thinking_frame = row

    scroll_messages_to_bottom()
    _animate_thinking(0)


def _animate_thinking(step):
    if not thinking_active or thinking_frame is None:
        return
    try:
        if not thinking_frame.winfo_exists():
            return
        dots = "." * ((step % 3) + 1)
        thinking_label.config(text=f"ARICE is thinking{dots}")
        root.after(450, lambda: _animate_thinking(step + 1))
    except tk.TclError:
        pass


def hide_thinking():
    global thinking_active, thinking_frame, thinking_label
    thinking_active = False
    if thinking_frame is not None:
        try:
            thinking_frame.destroy()
        except Exception:
            pass
    thinking_frame = None
    thinking_label = None


# ---- suggestion chips ----------------------------------------------------
def create_suggestion_chip(parent, text, command):
    chip = tk.Label(
        parent, text=text, font=(FONT_FAMILY, 10), fg=THEME["accent"],
        bg=THEME["card_bg"], padx=12, pady=7, cursor="hand2",
        highlightthickness=1, highlightbackground=THEME["accent_dim"]
    )
    chip.pack(side="left", padx=5, pady=5)
    chip.bind("<Button-1>", lambda e: command())
    return chip


FOLLOWUP_POOL = [
    "What can you do?",
    "Explain that more simply",
    "Give me a real example",
    "What else should I know?",
    "Tell me the time",
]

TOPIC_FOLLOWUPS = {
    "python": ["How do I install Python?", "Show me a Python example", "What is Python used for?"],
    "ai": ["How does AI learning work?", "What are common AI use cases?"],
    "artificial intelligence": ["How does AI learning work?", "What are common AI use cases?"],
    "machine learning": ["How is ML different from AI?", "Give me a beginner ML project idea"],
}


def generate_followups(query, reply):
    ql = query.lower()
    candidates = []
    for key, opts in TOPIC_FOLLOWUPS.items():
        if key in ql:
            candidates.extend(opts)
    if not candidates:
        candidates = list(FOLLOWUP_POOL)
    random.shuffle(candidates)
    return candidates[:3]


def add_followups(query, reply):
    if len(reply.strip()) < 4:
        return
    suggestions = generate_followups(query, reply)
    if not suggestions:
        return

    row = tk.Frame(messages_inner, bg=THEME["main_bg"])
    row.pack(fill="x", padx=16, pady=(0, 10), anchor="w")

    tk.Label(row, text="You may also ask:", font=(FONT_FAMILY, 9),
             fg=THEME["text_secondary"], bg=THEME["main_bg"]).pack(anchor="w", padx=6)

    chip_row = tk.Frame(row, bg=THEME["main_bg"])
    chip_row.pack(anchor="w")

    for s in suggestions:
        create_suggestion_chip(chip_row, s, lambda t=s: send_typed_message(t))

    scroll_messages_to_bottom()


# ---------------------------------------------------------------------
# 13. ANIMATED AI CORE CANVAS
# ---------------------------------------------------------------------
class AICoreCanvas(tk.Canvas):
    STATE_SPEED = {
        "idle": 0.03, "listening": 0.10, "thinking": 0.13,
        "generating": 0.13, "speaking": 0.09, "offline": 0.0,
    }
    STATE_INTERVAL = {
        "listening": 45, "thinking": 40, "generating": 40, "speaking": 45,
    }

    def __init__(self, parent, size=140, **kwargs):
        bg = kwargs.pop("bg", THEME["main_bg"])
        super().__init__(parent, width=size, height=size, bg=bg,
                          highlightthickness=0, **kwargs)
        self.size = size
        self.state = "idle"
        self.tick = 0
        self._running = True
        ai_core_canvases.append(self)
        self.bind("<Destroy>", self._on_destroy)
        self._animate()

    def _on_destroy(self, _event):
        self._running = False
        if self in ai_core_canvases:
            ai_core_canvases.remove(self)

    def set_state(self, state):
        self.state = state

    def _animate(self):
        if not self._running:
            return
        try:
            self._draw()
        except tk.TclError:
            return
        self.tick += 1
        interval = self.STATE_INTERVAL.get(self.state, 70)
        self.after(interval, self._animate)

    def _draw(self):
        self.delete("all")
        s = self.size
        cx = cy = s / 2
        t = self.tick
        speed = self.STATE_SPEED.get(self.state, 0.04)
        accent = THEME["accent"]
        dim = THEME["accent_dim"]

        for i, base_r in enumerate((s * 0.44, s * 0.33, s * 0.23)):
            pulse = math.sin(t * speed + i * 1.3) * s * 0.02
            r = max(4, base_r + pulse)
            self.create_oval(cx - r, cy - r, cx + r, cy + r, outline=dim, width=1)

        node_count = 6
        for i in range(node_count):
            angle = (2 * math.pi / node_count) * i + t * speed * 0.4
            nr = s * 0.47
            nx = cx + math.cos(angle) * nr
            ny = cy + math.sin(angle) * nr
            self.create_line(cx, cy, nx, ny, fill=dim)
            self.create_oval(nx - 2.5, ny - 2.5, nx + 2.5, ny + 2.5, fill=accent, outline="")

        core_color = accent if self.state != "offline" else THEME["danger"]
        core_pulse = math.sin(t * speed * 2.2) * s * 0.015 if self.state != "offline" else 0
        core_r = s * 0.10 + core_pulse
        self.create_oval(cx - core_r, cy - core_r, cx + core_r, cy + core_r,
                          fill=core_color, outline="")

        if self.state == "listening":
            bars = 5
            for i in range(bars):
                bh = 6 + ((i * 7 + t) % 16)
                bx = cx - (bars * 7) / 2 + i * 7
                self.create_line(bx, cy + s * 0.36, bx, cy + s * 0.36 - bh,
                                  fill=accent, width=3)


# ---------------------------------------------------------------------
# 14. MAIN WINDOW CONSTRUCTION
# ---------------------------------------------------------------------
def build_ui():
    global root, sidebar, chat_list_container, chat_list_canvas
    global main_area, content_area, welcome_frame
    global messages_outer, messages_canvas, messages_inner
    global input_entry, status_label, title_label
    global ai_status_dot, ai_status_text, mic_btn, send_btn, tools_btn

    root = tk.Tk()
    root.title("ARICE AI — Intelligent Personal Assistant")
    root.geometry("1400x860")
    root.minsize(1100, 700)
    root.configure(bg=THEME["bg_primary"])

    root.columnconfigure(0, weight=0)
    root.columnconfigure(1, weight=1)
    root.rowconfigure(0, weight=1)

    # ------------------------------------------------------------ SIDEBAR
    sidebar = tk.Frame(root, bg=THEME["sidebar_bg"], width=270)
    sidebar.grid(row=0, column=0, sticky="ns")
    sidebar.grid_propagate(False)

    brand = tk.Frame(sidebar, bg=THEME["sidebar_bg"])
    brand.pack(fill="x", padx=16, pady=(20, 6))
    tk.Label(brand, text="⚡ ARICE AI", font=(FONT_FAMILY, 17, "bold"),
             fg=THEME["accent"], bg=THEME["sidebar_bg"]).pack(anchor="w")
    tk.Label(brand, text="Your Intelligent Assistant", font=(FONT_FAMILY, 9),
             fg=THEME["text_secondary"], bg=THEME["sidebar_bg"]).pack(anchor="w")

    new_chat_btn = tk.Button(
        sidebar, text="＋  New Chat", font=(FONT_FAMILY, 11, "bold"),
        bg=THEME["accent"], fg="#01201c", bd=0, pady=9, cursor="hand2",
        activebackground=THEME["accent"], command=lambda: create_new_chat()
    )
    new_chat_btn.pack(fill="x", padx=16, pady=(10, 14))

    sb_outer, chat_list_canvas, chat_list_container = make_scrollable(sidebar, THEME["sidebar_bg"])
    sb_outer.pack(fill="both", expand=True)

    # ------------------------------------------------------------ MAIN AREA
    main_area = tk.Frame(root, bg=THEME["main_bg"])
    main_area.grid(row=0, column=1, sticky="nsew")
    main_area.rowconfigure(2, weight=1)
    main_area.columnconfigure(0, weight=1)

    # top bar
    top_bar = tk.Frame(main_area, bg=THEME["main_bg"], height=56)
    top_bar.grid(row=0, column=0, sticky="ew", padx=20, pady=(14, 0))

    title_label = tk.Label(top_bar, text="ARICE AI", font=(FONT_FAMILY, 16, "bold"),
                            fg=THEME["accent"], bg=THEME["main_bg"])
    title_label.pack(side="left")

    status_group = tk.Frame(top_bar, bg=THEME["main_bg"])
    status_group.pack(side="right")
    ai_status_dot = tk.Label(status_group, text="●", font=(FONT_FAMILY, 14, "bold"),
                              fg=THEME["accent"], bg=THEME["main_bg"])
    ai_status_dot.pack(side="left", padx=(0, 6))
    ai_status_text = tk.Label(status_group, text="ARICE ONLINE", font=(FONT_FAMILY, 11, "bold"),
                               fg=THEME["accent"], bg=THEME["main_bg"])
    ai_status_text.pack(side="left")

    tk.Frame(main_area, bg=THEME["border"], height=1).grid(row=1, column=0, sticky="ew", padx=20, pady=(10, 0))

    # content area (welcome OR chat)
    content_area = tk.Frame(main_area, bg=THEME["main_bg"])
    content_area.grid(row=2, column=0, sticky="nsew")
    content_area.rowconfigure(0, weight=1)
    content_area.columnconfigure(0, weight=1)

    # ---- welcome screen -------------------------------------------------
    welcome_frame = tk.Frame(content_area, bg=THEME["main_bg"])
    welcome_frame.grid(row=0, column=0, sticky="nsew")

    welcome_center = tk.Frame(welcome_frame, bg=THEME["main_bg"])
    welcome_center.place(relx=0.5, rely=0.42, anchor="center")

    AICoreCanvas(welcome_center, size=170, bg=THEME["main_bg"]).pack(pady=(0, 14))

    tk.Label(welcome_center, text="ARICE AI", font=(FONT_FAMILY, 26, "bold"),
             fg=THEME["accent"], bg=THEME["main_bg"]).pack()
    tk.Label(welcome_center, text="Your Intelligent Personal Assistant",
             font=(FONT_FAMILY, 12), fg=THEME["text_secondary"],
             bg=THEME["main_bg"]).pack(pady=(2, 4))
    tk.Label(welcome_center, text='"Hello! How can I help you today?"',
             font=(FONT_FAMILY, 11, "italic"), fg=THEME["text_primary"],
             bg=THEME["main_bg"]).pack(pady=(4, 18))

    suggestion_row = tk.Frame(welcome_center, bg=THEME["main_bg"])
    suggestion_row.pack()
    for s in ["What can you do?", "Help me with Python", "Explain AI", "Tell me the time"]:
        create_suggestion_chip(suggestion_row, s, lambda t=s: send_typed_message(t))

    # ---- chat message list ----------------------------------------------
    messages_outer, messages_canvas, messages_inner = make_scrollable(content_area, THEME["main_bg"])
    messages_outer.grid(row=0, column=0, sticky="nsew")
    messages_outer.grid_remove()  # start hidden; welcome screen shows first

    # ---- input bar --------------------------------------------------------
    input_bar_wrap = tk.Frame(main_area, bg=THEME["main_bg"])
    input_bar_wrap.grid(row=3, column=0, sticky="ew", padx=20, pady=16)
    input_bar_wrap.columnconfigure(0, weight=1)

    input_bar = tk.Frame(input_bar_wrap, bg=THEME["input_bg"], highlightthickness=1,
                          highlightbackground=THEME["border"])
    input_bar.grid(row=0, column=0, sticky="ew")
    input_bar.columnconfigure(1, weight=1)

    tools_btn = tk.Button(
        input_bar, text="＋", font=(FONT_FAMILY, 15, "bold"), bg=THEME["input_bg"],
        fg=THEME["accent"], bd=0, activebackground=THEME["input_bg"], cursor="hand2",
        command=lambda: show_tools_menu()
    )
    tools_btn.grid(row=0, column=0, padx=(10, 4), pady=8)

    input_entry = tk.Entry(
        input_bar, font=(FONT_FAMILY, 13), bg=THEME["input_bg"],
        fg=THEME["text_primary"], insertbackground=THEME["accent"], bd=0,
        relief="flat"
    )
    input_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=12)
    input_entry.insert(0, "")
    input_entry.bind("<Return>", handle_enter)

    mic_btn = tk.Button(
        input_bar, text="🎙", font=(FONT_FAMILY, 13), bg=THEME["input_bg"],
        fg=THEME["accent"], bd=0, activebackground=THEME["input_bg"], cursor="hand2",
        command=lambda: start_voice_thread()
    )
    mic_btn.grid(row=0, column=2, padx=4, pady=8)

    send_btn = tk.Button(
        input_bar, text="↑", font=(FONT_FAMILY, 14, "bold"), bg=THEME["accent"],
        fg="#01201c", bd=0, activebackground=THEME["accent"], cursor="hand2",
        width=3, command=lambda: send_typed_message()
    )
    send_btn.grid(row=0, column=3, padx=(4, 8), pady=8)

    status_label = tk.Label(input_bar_wrap, text="🎤 Ready", font=(FONT_FAMILY, 9),
                             fg=THEME["text_secondary"], bg=THEME["main_bg"])
    status_label.grid(row=1, column=0, sticky="w", pady=(6, 0))

    root.protocol("WM_DELETE_WINDOW", root.quit)


def show_welcome_view():
    messages_outer.grid_remove()
    welcome_frame.grid()


def show_chat_view():
    welcome_frame.grid_remove()
    messages_outer.grid()


# ---------------------------------------------------------------------
# 15. TOOLS / LANGUAGE / MEMORY / SETTINGS POPOVERS
# ---------------------------------------------------------------------
def _themed_menu():
    return tk.Menu(
        root, tearoff=0, font=(FONT_FAMILY, 11),
        bg=THEME["card_bg"], fg=THEME["text_primary"],
        activebackground=THEME["accent"], activeforeground="#01201c",
        bd=0
    )


def test_microphone():
    """Quick microphone test using the same capture path as voice chat."""
    play_click()
    set_status("🎧 Testing microphone…")
    threading.Thread(target=_microphone_test_worker, daemon=True,
                     name="AriceMicTest").start()


def _microphone_test_worker():
    query = takeCommand()
    if query:
        set_status(f"✅ Mic working — heard: {elide_title(query, 48)}")
    else:
        set_status("❌ Mic test failed — check microphone permissions/device")


def show_tools_menu():
    play_click()
    menu = _themed_menu()
    menu.add_command(label="🕐 Date & Time", command=tell_time)
    menu.add_command(label="🔁 Repeat Last", command=repeat_last)
    menu.add_command(label="🎨 Mode / Theme", command=toggle_theme)
    menu.add_command(label="📸 Screenshot", command=take_screenshot)
    menu.add_command(label="🌐 Language", command=show_language_menu)
    menu.add_command(label="🎙 Test Microphone", command=test_microphone)
    menu.add_command(label="🔊 Voice", command=toggle_voice)
    menu.add_command(label="⏹ Stop Speaking", command=stop_speaking)
    menu.add_command(label="🧠 Memory", command=open_memory_window)
    menu.add_command(label="⚙ Settings", command=open_settings_window)

    x = tools_btn.winfo_rootx()
    y = tools_btn.winfo_rooty()
    menu.post(x, y - 250)


def show_language_menu():
    menu = _themed_menu()
    menu.add_command(label="🇬🇺 Gujarati", command=lambda: select_language("Gujarati"))
    menu.add_command(label="🇮🇳 Hindi", command=lambda: select_language("Hindi"))
    menu.add_command(label="🇬🇧 English", command=lambda: select_language("English"))
    x = tools_btn.winfo_rootx() + 40
    y = tools_btn.winfo_rooty() - 90
    menu.post(x, y)


def open_memory_window():
    play_click()
    win = tk.Toplevel(root)
    win.title("ARICE Memory")
    win.geometry("440x560")
    win.configure(bg=THEME["main_bg"])

    tk.Label(win, text="🧠 ARICE MEMORY", font=(FONT_FAMILY, 18, "bold"),
             fg=THEME["accent"], bg=THEME["main_bg"]).pack(pady=(18, 4))

    status_row = tk.Frame(win, bg=THEME["main_bg"])
    status_row.pack(pady=(0, 14))
    tk.Label(status_row, text="Memory Status", font=(FONT_FAMILY, 10),
              fg=THEME["text_secondary"], bg=THEME["main_bg"]).pack(side="left", padx=(0, 6))
    tk.Label(status_row, text="● Active", font=(FONT_FAMILY, 10, "bold"),
              fg=THEME["accent"], bg=THEME["main_bg"]).pack(side="left")

    data = load_memory()

    cat_var = tk.StringVar(value="preferences")
    cat_row = tk.Frame(win, bg=THEME["main_bg"])
    cat_row.pack(pady=6)

    listbox_holder = tk.Frame(win, bg=THEME["main_bg"])
    listbox_holder.pack(fill="both", expand=True, padx=20, pady=8)

    listbox = tk.Listbox(
        listbox_holder, font=(FONT_FAMILY, 11), bg=THEME["card_bg"],
        fg=THEME["text_primary"], selectbackground=THEME["accent"],
        selectforeground="#01201c", bd=0, highlightthickness=0
    )
    listbox.pack(fill="both", expand=True)

    def refresh_list():
        listbox.delete(0, tk.END)
        for item in data.get(cat_var.get(), []):
            listbox.insert(tk.END, item)

    def select_category(cat):
        cat_var.set(cat)
        for b, c in cat_buttons:
            b.config(bg=THEME["accent"] if c == cat else THEME["card_bg"],
                     fg="#01201c" if c == cat else THEME["text_primary"])
        refresh_list()

    cat_buttons = []
    for cat, label in MEMORY_CATEGORY_ICONS.items():
        b = tk.Button(cat_row, text=label, font=(FONT_FAMILY, 9), bd=0,
                       bg=THEME["card_bg"], fg=THEME["text_primary"], cursor="hand2",
                       command=lambda c=cat: select_category(c))
        b.pack(side="left", padx=4)
        cat_buttons.append((b, cat))

    entry_row = tk.Frame(win, bg=THEME["main_bg"])
    entry_row.pack(fill="x", padx=20, pady=(0, 8))
    new_entry = tk.Entry(entry_row, font=(FONT_FAMILY, 11), bg=THEME["input_bg"],
                          fg=THEME["text_primary"], insertbackground=THEME["text_primary"], bd=2)
    new_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

    def add_item():
        text = new_entry.get().strip()
        if text:
            data.setdefault(cat_var.get(), []).append(text)
            save_memory(data)
            new_entry.delete(0, tk.END)
            refresh_list()

    tk.Button(entry_row, text="Add", font=(FONT_FAMILY, 10, "bold"), bg=THEME["accent"],
              fg="#01201c", bd=0, cursor="hand2", command=add_item).pack(side="left")

    def remove_selected():
        sel = listbox.curselection()
        if sel:
            items = data.get(cat_var.get(), [])
            del items[sel[0]]
            save_memory(data)
            refresh_list()

    tk.Button(win, text="🗑 Remove Selected", font=(FONT_FAMILY, 10), bg=THEME["danger"],
              fg="white", bd=0, cursor="hand2", command=remove_selected).pack(pady=(0, 16))

    select_category("preferences")


def open_settings_window():
    play_click()
    win = tk.Toplevel(root)
    win.title("ARICE Settings")
    win.geometry("360x260")
    win.configure(bg=THEME["main_bg"])

    tk.Label(win, text="⚙ Settings", font=(FONT_FAMILY, 16, "bold"),
             fg=THEME["accent"], bg=THEME["main_bg"]).pack(pady=(18, 10))

    tk.Button(win, text="🌓 Toggle Dark / Light Theme", font=(FONT_FAMILY, 11),
              bg=THEME["card_bg"], fg=THEME["text_primary"], bd=0, pady=8,
              cursor="hand2", command=toggle_theme).pack(fill="x", padx=24, pady=6)

    tk.Button(win, text="🔊 Switch Voice", font=(FONT_FAMILY, 11),
              bg=THEME["card_bg"], fg=THEME["text_primary"], bd=0, pady=8,
              cursor="hand2", command=toggle_voice).pack(fill="x", padx=24, pady=6)

    mic_info = tk.Label(
        win, text="Microphone: Windows default device",
        font=(FONT_FAMILY, 9), fg=THEME["text_secondary"], bg=THEME["main_bg"]
    )
    mic_info.pack(pady=(10, 2))

    tk.Button(
        win, text="🎙 Test Microphone", font=(FONT_FAMILY, 10),
        bg=THEME["card_bg"], fg=THEME["text_primary"], bd=0, pady=6,
        cursor="hand2", command=test_microphone
    ).pack(fill="x", padx=24, pady=5)

    tk.Label(win, text="ARICE AI — Personal Desktop Assistant\nBuilt with Tkinter + Gemini",
             font=(FONT_FAMILY, 9), fg=THEME["text_secondary"], bg=THEME["main_bg"],
             justify="center").pack(pady=(12, 8))


# ---------------------------------------------------------------------
# 16. THEME APPLICATION
# ---------------------------------------------------------------------
def toggle_theme():
    global current_theme
    play_click()
    current_theme = "light" if current_theme == "dark" else "dark"
    THEME.clear()
    THEME.update(THEMES[current_theme])
    apply_theme()


def apply_theme():
    root.configure(bg=THEME["bg_primary"])
    sidebar.configure(bg=THEME["sidebar_bg"])
    main_area.configure(bg=THEME["main_bg"])
    content_area.configure(bg=THEME["main_bg"])
    welcome_frame.configure(bg=THEME["main_bg"])
    messages_canvas.configure(bg=THEME["main_bg"])
    messages_inner.configure(bg=THEME["main_bg"])
    title_label.configure(bg=THEME["main_bg"], fg=THEME["accent"])
    status_label.configure(bg=THEME["main_bg"], fg=THEME["text_secondary"])
    ai_status_dot.configure(bg=THEME["main_bg"])
    ai_status_text.configure(bg=THEME["main_bg"], fg=THEME["accent"])

    for canvas in list(ai_core_canvases):
        try:
            canvas.configure(bg=THEME["main_bg"])
        except tk.TclError:
            pass

    # New chat bubbles/messages will pick up the fresh THEME automatically.
    # Existing bubbles keep their original colors (Tkinter has no live
    # global restyle for arbitrary widget trees) — reloading the current
    # chat refreshes them instantly:
    if current_chat_id:
        load_chat(current_chat_id)

    refresh_history_list()
    set_status(f"🎨 {current_theme.title()} theme applied")


# kept for naming compatibility with the original file
def apply_dark_theme():
    global current_theme
    current_theme = "dark"
    THEME.clear()
    THEME.update(THEMES["dark"])
    apply_theme()


def apply_light_theme():
    global current_theme
    current_theme = "light"
    THEME.clear()
    THEME.update(THEMES["light"])
    apply_theme()


# ---------------------------------------------------------------------
# 18. APPLICATION STARTUP
# ---------------------------------------------------------------------
if __name__ == "__main__":
    build_ui()
    create_new_chat()
    set_ai_status("ARICE ONLINE", online=True)

    threading.Thread(
        target=lambda: say(
            "Hello! I am Arice. You can speak, type, or repeat my answers."
        ),
        daemon=True
    ).start()

    log("👋 Hello! I am Arice. Ready to assist you.")

    root.mainloop()
