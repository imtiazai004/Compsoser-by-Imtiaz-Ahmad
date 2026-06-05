import streamlit as st
import streamlit.components.v1 as components
import json
import os
import requests
import base64
import mimetypes
import time
import tempfile
from docx import Document
import io

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except Exception:
    EASYOCR_AVAILABLE = False

# ── Developer API Keys ────────────────────────────────────────────────────────
# Works with: Streamlit Cloud, HF Spaces, local secrets.toml, env variable
def _load_dev_keys():
    # 1. st.secrets as list (Streamlit Cloud / local secrets.toml)
    try:
        raw = st.secrets["developer_keys"]
        if isinstance(raw, (list, tuple)):
            return [k for k in raw if k]
        return json.loads(str(raw))
    except Exception:
        pass
    # 2. Environment variable fallback (HF Spaces)
    try:
        raw = os.environ.get("DEVELOPER_KEYS", "")
        if raw:
            return json.loads(raw)
    except Exception:
        pass
    return []

DEVELOPER_KEYS = _load_dev_keys()

GEMINI_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-1.5-flash-8b",
    "gemini-1.5-pro",
]

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


# ── Config ────────────────────────────────────────────────────────────────────

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE) as f:
                return json.load(f)
    except Exception:
        pass
    return {"user_key": ""}

def save_config(data):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def get_all_keys():
    keys = [k.strip() for k in DEVELOPER_KEYS if k.strip()]
    user_key = load_config().get("user_key", "").strip()
    if user_key and user_key not in keys:
        keys.append(user_key)
    return keys


# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Composer — AI Document Scanner",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── PWA manifest ─────────────────────────────────────────────────────────────
st.markdown("""
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/static/icon-192.png">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Composer">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="#6366f1">
""", unsafe_allow_html=True)

st.markdown("""
<style>
/* ── Global ── */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 2rem 2rem 2rem !important; max-width: 1300px; }
* { font-family: 'Segoe UI', system-ui, sans-serif; }
body, .stApp { background: #f5f6fa !important; }

/* ── Header spacer ── */
.header-spacer { margin-bottom: 1rem; }

/* ── Steps ── */
.steps-row {
    display: flex;
    gap: 0;
    margin-bottom: 1.4rem;
    background: #ffffff;
    border-radius: 14px;
    padding: 14px 24px;
    align-items: center;
    box-shadow: 0 1px 8px rgba(0,0,0,0.07);
    border: 1px solid #e8eaf6;
}
.step { display: flex; align-items: center; gap: 10px; flex: 1; }
.step-num {
    width: 30px; height: 30px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.85rem;
    background: #f1f3f9; color: #b0b8d8;
    border: 2px solid #dde3f5;
    transition: all 0.3s;
    flex-shrink: 0;
}
.step-num.active  { background: #6366f1; color: #fff; border-color: #6366f1; box-shadow: 0 0 12px #6366f140; }
.step-num.done    { background: #dcfce7; color: #16a34a; border-color: #86efac; }
.step-label { color: #94a3b8; font-size: 0.85rem; }
.step-label.active { color: #1e1b4b; font-weight: 600; }
.step-divider { width: 40px; height: 2px; background: #e8eaf6; margin: 0 8px; flex-shrink: 0; }
.step-divider.done { background: #86efac; }

/* ── Cards ── */
.card {
    background: #ffffff;
    border: 1px solid #e8eaf6;
    border-radius: 16px;
    padding: 20px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.05);
}

/* ── Upload zone ── */
.upload-hint {
    border: 2px dashed #c7d2fe;
    border-radius: 12px;
    padding: 32px 20px;
    text-align: center;
    color: #818cf8;
    background: #f8f9ff;
    margin-bottom: 12px;
}
.upload-hint .icon { font-size: 2.5rem; margin-bottom: 8px; }
.upload-hint p { margin: 0; font-size: 0.9rem; color: #6366f1; }

/* ── Result card ── */
.result-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 10px;
}
.result-title { color: #1e1b4b; font-size: 1rem; font-weight: 600; margin: 0; }
.stat-chip {
    background: #f1f3f9;
    border: 1px solid #e0e7ff;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.78rem;
    color: #6366f1;
    display: inline-block;
    margin-right: 6px;
}

/* ── Buttons ── */
.stButton > button {
    border-radius: 10px !important;
    font-weight: 700 !important;
    transition: all 0.2s !important;
    position: relative !important;
}
.stButton > button:hover { transform: translateY(-1px); }

/* ── Rainbow primary button ── */
.stButton > button[kind="primary"] {
    background: #000 !important;
    color: #fff !important;
    border: none !important;
    padding: 10px 24px !important;
    font-size: 0.95rem !important;
    z-index: 0 !important;
    overflow: visible !important;
    border-radius: 12px !important;
}
.stButton > button[kind="primary"]::before,
.stButton > button[kind="primary"]::after {
    content: '' !important;
    position: absolute !important;
    left: -2px !important;
    top: -2px !important;
    border-radius: 14px !important;
    background: linear-gradient(45deg,
        #fb0094, #0000ff, #00ff00, #ffff00,
        #ff0000, #fb0094, #0000ff, #00ff00,
        #ffff00, #ff0000) !important;
    background-size: 400% !important;
    width: calc(100% + 4px) !important;
    height: calc(100% + 4px) !important;
    z-index: -1 !important;
    animation: rainbow 20s linear infinite !important;
}
.stButton > button[kind="primary"]::after {
    filter: blur(12px) !important;
    opacity: 0.6 !important;
}
@keyframes rainbow {
    0%   { background-position: 0 0; }
    50%  { background-position: 400% 0; }
    100% { background-position: 0 0; }
}

/* ── Language pills ── */
.stRadio > div { gap: 8px !important; }
.stRadio label {
    background: #f1f3f9 !important;
    border-radius: 20px !important;
    padding: 4px 14px !important;
    border: 1px solid #e0e7ff !important;
    color: #4f46e5 !important;
}

/* ── Spinner ── */
.stSpinner > div { border-top-color: #6366f1 !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e8eaf6 !important; }

/* ── Text area ── */
.stTextArea textarea {
    background: #fafbff !important;
    border: 1px solid #e0e7ff !important;
    border-radius: 10px !important;
    color: #1e293b !important;
    font-size: 0.9rem !important;
}
</style>
""", unsafe_allow_html=True)


# ── Session state ─────────────────────────────────────────────────────────────

for key, default in [
    ("exhausted_keys", set()),
    ("show_key_dialog", False),
    ("use_easyocr_fallback", False),
    ("extracted_text", ""),
    ("last_engine", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ── Dialogs ───────────────────────────────────────────────────────────────────

@st.dialog("⚙️ API Key Settings")
def settings_dialog():
    cfg = load_config()
    current = cfg.get("user_key", "").strip()
    if current:
        st.success(f"Personal key active: `{current[:8]}...{current[-4:]}`")
    else:
        n = len([k for k in DEVELOPER_KEYS if k.strip()])
        st.info(f"Using {n} built-in developer key(s). Add yours for extra quota.")

    new_key = st.text_input("Your Gemini API Key:", type="password",
                            placeholder="AIza... or AQ.Ab8...",
                            label_visibility="collapsed")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Save Key", type="primary", use_container_width=True):
            if new_key.strip():
                cfg["user_key"] = new_key.strip()
                save_config(cfg)
                st.session_state.exhausted_keys = set()
                st.session_state.use_easyocr_fallback = False
                st.rerun()
            else:
                st.warning("Please enter a key.")
    with c2:
        if st.button("🗑 Remove Key", use_container_width=True):
            cfg["user_key"] = ""
            save_config(cfg)
            st.session_state.exhausted_keys = set()
            st.rerun()


# ── OCR engines ───────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner="Loading EasyOCR model...")
def load_reader(langs):
    return easyocr.Reader(langs, gpu=False)


def try_key(key, image_bytes, filename):
    mime_type, _ = mimetypes.guess_type(filename)
    if mime_type not in ("image/jpeg", "image/png", "image/gif", "image/webp"):
        mime_type = "image/jpeg"
    encoded = base64.b64encode(image_bytes).decode("utf-8")

    for model in GEMINI_MODELS:
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"{model}:generateContent?key={key}"
        )
        payload = {"contents": [{"parts": [
            {"inline_data": {"mime_type": mime_type, "data": encoded}},
            {"text": (
                "Extract all text from this image exactly as it appears. "
                "Preserve line breaks and formatting. "
                "Return only the extracted text, no commentary."
            )},
        ]}]}
        try:
            resp = requests.post(url, json=payload, timeout=30)
            data = resp.json()
        except Exception:
            continue

        if "error" in data:
            code = data["error"].get("code", 0)
            status = data["error"].get("status", "")
            if code in (400, 401, 403) and any(
                x in status for x in ("API_KEY_INVALID", "PERMISSION_DENIED", "UNAUTHENTICATED")
            ):
                return None, "INVALID_KEY"
            continue

        try:
            return data["candidates"][0]["content"]["parts"][0]["text"], None
        except (KeyError, IndexError):
            continue

    return None, "QUOTA"


def run_gemini(image_bytes, filename):
    all_keys = get_all_keys()
    active = [k for k in all_keys if k not in st.session_state.exhausted_keys]
    if not active:
        return None, "ALL_KEYS_EXHAUSTED"
    for key in active:
        text, _ = try_key(key, image_bytes, filename)
        if text:
            return text, None
        st.session_state.exhausted_keys.add(key)
    return None, "ALL_KEYS_EXHAUSTED"


def run_easyocr(image_bytes, lang_choice):
    if not EASYOCR_AVAILABLE:
        return None, "EasyOCR unavailable — restart your PC to fix the DLL error."
    langs = (["en"] if lang_choice == "English"
             else ["ur"] if lang_choice == "Urdu"
             else ["en", "ur"])
    reader = load_reader(tuple(langs))
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(image_bytes)
        tmp_path = tmp.name
    try:
        results = reader.readtext(tmp_path)
        return "\n".join(r[1] for r in results), None
    finally:
        os.unlink(tmp_path)


def to_docx_bytes(text):
    doc = Document()
    doc.add_paragraph(text)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def copy_button(text):
    safe = json.dumps(text)
    components.html(f"""
    <button onclick="
        navigator.clipboard.writeText({safe}).then(() => {{
            this.innerText = '✅ Copied!';
            this.style.background = '#dcfce7';
            this.style.color = '#16a34a';
            this.style.borderColor = '#86efac';
            setTimeout(() => {{
                this.innerText = '📋 Copy Text';
                this.style.background = '#eef2ff';
                this.style.color = '#4f46e5';
                this.style.borderColor = '#c7d2fe';
            }}, 2000);
        }});
    " style="
        background:#eef2ff; color:#4f46e5; border:1px solid #c7d2fe;
        border-radius:8px; padding:7px 18px; font-size:13px;
        cursor:pointer; font-weight:600; transition:all 0.2s;
        font-family: Segoe UI, system-ui, sans-serif;
    ">📋 Copy Text</button>
    """, height=42)


def pwa_install_button():
    components.html("""
    <style>
    *{box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif;margin:0;}
    body{background:transparent;overflow:hidden;}
    #wrap{display:none;flex-direction:column;gap:8px;}
    #btn{
        background:linear-gradient(135deg,#6366f1,#8b5cf6);
        color:white;border:none;border-radius:10px;
        padding:10px 16px;font-weight:700;font-size:13px;
        cursor:pointer;width:100%;text-align:left;
        display:flex;align-items:center;gap:10px;
        box-shadow:0 2px 12px rgba(99,102,241,0.35);
        transition:all 0.2s;
    }
    #btn:hover{transform:translateY(-1px);box-shadow:0 4px 18px rgba(99,102,241,0.45);}
    #guide{
        background:#f0f4ff;border:1px solid #c7d2fe;
        border-radius:10px;padding:12px 14px;
        font-size:12px;color:#374151;line-height:2;display:none;
    }
    </style>
    <div id="wrap">
        <button id="btn" onclick="doInstall()">
            <span style="font-size:1.4rem;">📲</span>
            <span><b>Install App</b><br><span style="font-weight:400;opacity:0.85;font-size:11px;">Add to home screen</span></span>
        </button>
        <div id="guide">
            🖥️ <b>PC/Mac:</b> Click <b>⊕</b> in address bar<br>
            📱 <b>iPhone:</b> Safari → Share → <b>Add to Home Screen</b><br>
            📱 <b>Android:</b> Chrome Menu → <b>Add to Home Screen</b>
        </div>
    </div>
    <script>
    let installPrompt = null;
    const isStandalone = window.matchMedia('(display-mode:standalone)').matches || navigator.standalone;

    if (!isStandalone) {
        document.getElementById('wrap').style.display = 'flex';

        const capture = (w) => {
            try {
                w.addEventListener('beforeinstallprompt', e => {
                    e.preventDefault();
                    installPrompt = e;
                });
                w.addEventListener('appinstalled', () => {
                    document.getElementById('wrap').style.display = 'none';
                });
            } catch(e){}
        };
        capture(window);
        try { capture(window.top); } catch(e){}
        try { capture(window.parent); } catch(e){}

        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.register('/sw.js', {scope:'/'}).catch(()=>{});
            try { window.top.navigator.serviceWorker.register('/sw.js',{scope:'/'}).catch(()=>{}); } catch(e){}
        }
    }

    function doInstall() {
        if (installPrompt) {
            installPrompt.prompt();
            installPrompt.userChoice.then(() => { installPrompt = null; });
        } else {
            const g = document.getElementById('guide');
            g.style.display = g.style.display === 'block' ? 'none' : 'block';
        }
    }
    </script>
    """, height=100)


def show_result_panel(text, engine_label):
    words = len(text.split())
    chars = len(text)

    st.markdown(f"""
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
        <span style="color:#1e1b4b;font-weight:700;font-size:1rem;">📋 Extracted Text</span>
        <span>
            <span class="stat-chip">📝 {words} words</span>
            <span class="stat-chip">🔤 {chars} chars</span>
            <span class="stat-chip" style="color:#16a34a;border-color:#86efac;background:#f0fdf4;">✅ {engine_label}</span>
        </span>
    </div>
    """, unsafe_allow_html=True)

    st.text_area("", value=text, height=360, label_visibility="collapsed", key="result_textarea")

    col_copy, col_word, col_txt = st.columns([1, 1.2, 1.2])
    with col_copy:
        copy_button(text)
    with col_word:
        st.download_button("💾 Save as Word",
            data=to_docx_bytes(text), file_name="extracted_text.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    with col_txt:
        st.download_button("📄 Save as TXT",
            data=text, file_name="extracted_text.txt",
            mime="text/plain", use_container_width=True)


def show_quota_panel():
    st.markdown("""
<div style="background:#fffbeb;border:1px solid #fde68a;
     border-radius:16px;padding:28px 32px;margin-top:8px;
     box-shadow:0 2px 12px rgba(234,179,8,0.1);">

<h3 style="color:#b45309;margin:0 0 8px 0;">⚠️ Gemini API Quota Exhausted</h3>
<p style="color:#78716c;margin:0 0 22px 0;font-size:14.5px;line-height:1.6;">
All available Gemini API keys have reached their daily limit.<br>
<b style="color:#0891b2;">✦ Quota resets automatically every 24 hours.</b>
</p>

<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">

<div style="background:#eff6ff;border-radius:12px;padding:18px 20px;border:1px solid #bfdbfe;">
<h4 style="color:#1d4ed8;margin:0 0 10px 0;">🔑 Option 1 — Add Your Free Key</h4>
<ol style="color:#374151;font-size:13.5px;margin:0;padding-left:18px;line-height:2;">
<li>Visit <a href="https://aistudio.google.com/apikey" target="_blank"
    style="color:#2563eb;font-weight:600;">aistudio.google.com/apikey</a></li>
<li>Sign in with Google</li>
<li>Click <b style="color:#1e40af;">"Create API key"</b></li>
<li>Copy &amp; paste it using the button below</li>
</ol>
</div>

<div style="background:#f0fdf4;border-radius:12px;padding:18px 20px;border:1px solid #bbf7d0;">
<h4 style="color:#15803d;margin:0 0 10px 0;">🔍 Option 2 — EasyOCR (Free, Offline)</h4>
<p style="color:#374151;font-size:13.5px;margin:0;line-height:1.8;">
No API key needed.<br>
Works completely offline.<br>
<span style="color:#b45309;">⚠ Lower accuracy for Urdu text.</span>
</p>
</div>

</div></div>
""", unsafe_allow_html=True)

    st.markdown("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("➕ Add My API Key", type="primary", use_container_width=True):
            settings_dialog()
    with c2:
        label = "🔍 Use EasyOCR Instead" if EASYOCR_AVAILABLE else "🔍 EasyOCR (Restart PC first)"
        if st.button(label, use_container_width=True, disabled=not EASYOCR_AVAILABLE):
            st.session_state.use_easyocr_fallback = True
            st.rerun()


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    pwa_install_button()
    st.markdown("---")
    if st.button("🔑 Manage API Key", use_container_width=True):
        settings_dialog()
    st.markdown("---")
    cfg = load_config()
    if cfg.get("user_key", "").strip():
        st.success("Personal key: Active ✓")
    else:
        n = len([k for k in DEVELOPER_KEYS if k.strip()])
        st.info(f"{n} built-in key(s) active")
    if st.session_state.exhausted_keys:
        st.warning(f"{len(st.session_state.exhausted_keys)} key(s) exhausted")
        if st.button("🔄 Reset Keys", use_container_width=True):
            st.session_state.exhausted_keys = set()
            st.session_state.use_easyocr_fallback = False
            st.rerun()


# ── Header ────────────────────────────────────────────────────────────────────

components.html("""
<!DOCTYPE html>
<html>
<head>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', system-ui, sans-serif; }
html, body { background: transparent; overflow: hidden; height: 100%; }

.app-header {
    background: linear-gradient(120deg, #ffffff 0%, #eef2ff 100%);
    padding: 20px 32px;
    border-radius: 0 0 20px 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid #e0e7ff;
    box-shadow: 0 2px 14px rgba(99,102,241,0.1);
}

.sparkles-wrapper {
    position: relative;
    display: inline-block;
    padding: 0 6px 2px 0;
}

.sparkles-title {
    font-size: 1.65rem;
    font-weight: 800;
    color: #1e1b4b;
    letter-spacing: 0.3px;
    line-height: 1.2;
}

.app-subtitle {
    color: #6366f1;
    font-size: 0.83rem;
    margin-top: 5px;
    font-weight: 500;
    opacity: 0.9;
}

.sparkle-star {
    position: absolute;
    pointer-events: none;
    z-index: 20;
    opacity: 0;
    animation: sparkle-life 0.8s ease-in-out forwards;
}

@keyframes sparkle-life {
    0%   { opacity: 0; transform: scale(0) rotate(75deg); }
    40%  { opacity: 1; transform: scale(1) rotate(120deg); }
    100% { opacity: 0; transform: scale(0) rotate(150deg); }
}
</style>
</head>
<body>
<div class="app-header">
    <div>
        <div class="sparkles-wrapper" id="sparkles-wrapper">
            <span class="sparkles-title">📄 Composer</span>
        </div>
        <p class="app-subtitle">✨ AI-powered document &amp; image text extractor — by Imtiaz Ahmad</p>
    </div>
    <div style="font-size:2.2rem;opacity:0.25;line-height:1;user-select:none;">🔍</div>
</div>

<script>
const wrapper = document.getElementById('sparkles-wrapper');
const colors = ['#9E7AFF', '#FE8BBB', '#818cf8', '#f472b6', '#c084fc'];
const STAR_PATH = 'M9.82531 0.843845C10.0553 0.215178 10.9446 0.215178 11.1746 0.843845L11.8618 2.72026C12.4006 4.19229 12.3916 6.39157 13.5 7.5C14.6084 8.60843 16.8077 8.59935 18.2797 9.13822L20.1561 9.82534C20.7858 10.0553 20.7858 10.9447 20.1561 11.1747L18.2797 11.8618C16.8077 12.4007 14.6084 12.3916 13.5 13.5C12.3916 14.6084 12.4006 16.8077 11.8618 18.2798L11.1746 20.1562C10.9446 20.7858 10.0553 20.7858 9.82531 20.1562L9.13819 18.2798C8.59932 16.8077 8.60843 14.6084 7.5 13.5C6.39157 12.3916 4.19225 12.4007 2.72023 11.8618L0.843814 11.1747C0.215148 10.9447 0.215148 10.0553 0.843814 9.82534L2.72023 9.13822C4.19225 8.59935 6.39157 8.60843 7.5 7.5C8.60843 6.39157 8.59932 4.19229 9.13819 2.72026L9.82531 0.843845Z';

function spawnSparkle() {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    const size = 9 + Math.random() * 13;
    svg.setAttribute('width', size);
    svg.setAttribute('height', size);
    svg.setAttribute('viewBox', '0 0 21 21');
    svg.classList.add('sparkle-star');

    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', STAR_PATH);
    path.setAttribute('fill', colors[Math.floor(Math.random() * colors.length)]);
    svg.appendChild(path);

    svg.style.left  = (Math.random() * 108 - 4) + '%';
    svg.style.top   = (Math.random() * 110 - 10) + '%';
    svg.style.animationDuration = (0.5 + Math.random() * 0.9) + 's';
    svg.style.animationDelay    = (Math.random() * 0.15) + 's';

    wrapper.appendChild(svg);
    svg.addEventListener('animationend', () => { try { svg.remove(); } catch(e){} });
}

setInterval(spawnSparkle, 160);
</script>
</body>
</html>
""", height=86, scrolling=False)


# ── Main layout ───────────────────────────────────────────────────────────────

col_left, col_right = st.columns([1, 1.55], gap="large")

with col_left:
    # ── Upload ──────────────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("**📁 Upload Image**")

    uploaded = st.file_uploader(
        "upload",
        type=["png", "jpg", "jpeg", "bmp", "tiff", "tif"],
        label_visibility="collapsed",
    )

    if uploaded:
        st.image(uploaded, use_container_width=True, caption=f"📎 {uploaded.name}")
    else:
        st.markdown("""
        <div class="upload-hint">
            <div class="icon">📄</div>
            <p><b>Click above to upload</b></p>
            <p style="color:#3a3a6a;font-size:0.8rem;margin-top:6px;">
            PNG · JPG · JPEG · BMP · TIFF
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Options ─────────────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("**🌐 Language**")
    lang = st.radio("lang", ["English + Urdu", "English", "Urdu"],
                    horizontal=True, label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**⚙️ OCR Engine**")
    engine_options = ["✨ Gemini AI (Default)"]
    if EASYOCR_AVAILABLE:
        engine_options.append("🔍 EasyOCR (Offline)")
    else:
        engine_options.append("🔍 EasyOCR (Unavailable)")
    engine = st.selectbox("engine", engine_options, label_visibility="collapsed")
    use_gemini = engine.startswith("✨")

    st.markdown("<br>", unsafe_allow_html=True)
    convert_btn = st.button(
        "▶  Convert to Text",
        type="primary",
        use_container_width=True,
        disabled=(not uploaded or (not use_gemini and not EASYOCR_AVAILABLE)),
    )
    st.markdown('</div>', unsafe_allow_html=True)


with col_right:

    # ── Step indicator ───────────────────────────────────────────────────────
    s1 = "done" if uploaded else ("active" if not uploaded else "")
    s2 = "active" if uploaded and not st.session_state.extracted_text else ("done" if st.session_state.extracted_text else "")
    s3 = "done" if st.session_state.extracted_text else ("active" if uploaded else "")
    div1 = "done" if uploaded else ""
    div2 = "done" if st.session_state.extracted_text else ""

    st.markdown(f"""
    <div class="steps-row">
        <div class="step">
            <div class="step-num {s1}">1</div>
            <span class="step-label {'active' if not uploaded else ''}">Upload Image</span>
        </div>
        <div class="step-divider {div1}"></div>
        <div class="step">
            <div class="step-num {s2}">2</div>
            <span class="step-label {'active' if uploaded and not st.session_state.extracted_text else ''}">Select Language</span>
        </div>
        <div class="step-divider {div2}"></div>
        <div class="step">
            <div class="step-num {s3}">3</div>
            <span class="step-label {'active' if st.session_state.extracted_text else ''}">Extract Text</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── No image uploaded ────────────────────────────────────────────────────
    if not uploaded:
        st.session_state.use_easyocr_fallback = False
        st.session_state.extracted_text = ""
        st.markdown("""
        <div style="text-align:center;padding:80px 20px;">
            <div style="font-size:4rem;margin-bottom:16px;filter:grayscale(0.3);">🖼️</div>
            <h3 style="color:#6366f1;margin:0 0 8px 0;font-weight:700;">No image uploaded yet</h3>
            <p style="font-size:0.9rem;margin:0;color:#94a3b8;">Upload a scanned document on the left to get started.</p>
        </div>
        """, unsafe_allow_html=True)

    # ── EasyOCR fallback ─────────────────────────────────────────────────────
    elif st.session_state.use_easyocr_fallback:
        with st.spinner("Running EasyOCR (offline)..."):
            t0 = time.time()
            text, err = run_easyocr(uploaded.read(), lang)
            elapsed = round(time.time() - t0, 2)
        if err:
            st.error(err)
        else:
            st.session_state.extracted_text = text
            show_result_panel(text, f"EasyOCR · {elapsed}s")
        if st.button("← Try Gemini Again", use_container_width=False):
            st.session_state.use_easyocr_fallback = False
            st.session_state.exhausted_keys = set()
            st.session_state.extracted_text = ""
            st.rerun()

    # ── Convert triggered ────────────────────────────────────────────────────
    elif convert_btn:
        image_bytes = uploaded.read()

        if use_gemini:
            with st.spinner("🤖 Gemini AI is reading your document..."):
                t0 = time.time()
                text, err = run_gemini(image_bytes, uploaded.name)
                elapsed = round(time.time() - t0, 2)

            if err == "ALL_KEYS_EXHAUSTED":
                show_quota_panel()
            elif err:
                st.error(err)
            else:
                st.session_state.extracted_text = text
                st.balloons()
                show_result_panel(text, f"Gemini AI · {elapsed}s")
        else:
            with st.spinner("🔍 EasyOCR is scanning your image..."):
                t0 = time.time()
                text, err = run_easyocr(image_bytes, lang)
                elapsed = round(time.time() - t0, 2)
            if err:
                st.error(err)
            else:
                st.session_state.extracted_text = text
                show_result_panel(text, f"EasyOCR · {elapsed}s")

    # ── Waiting for convert ──────────────────────────────────────────────────
    elif st.session_state.extracted_text:
        show_result_panel(st.session_state.extracted_text, st.session_state.last_engine)
    else:
        st.markdown("""
        <div style="text-align:center;padding:80px 20px;">
            <div style="font-size:3.5rem;margin-bottom:16px;">⚡</div>
            <h3 style="color:#6366f1;margin:0 0 8px 0;font-weight:700;">Ready to extract</h3>
            <p style="font-size:0.9rem;margin:0;color:#94a3b8;">Click <b style="color:#4f46e5;">Convert to Text</b> to begin.</p>
        </div>
        """, unsafe_allow_html=True)
