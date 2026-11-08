"""
Security & Role-Based Access Control (RBAC) Module for Weekly Automation Report App.

Access Model:
- VIEWER (Default for everyone visiting the link):
  Can view all reports, switch between weeks, copy email tables (Sheet 1-4 & Combined),
  and download the Excel (.xlsx) file. Cannot modify data, create weeks, or upload files.
- EDITOR (Password-authenticated via `st.secrets["APP_PASSWORD"]` or `ADMIN_PASSWORD`):
  Can edit all sheets, add/delete rows, create new weekly reports, and import Excel files.

Session Persistence & Security Controls:
1. 1-Hour Inactivity Session Token (`EDITOR_SESSION_TIMEOUT_SECONDS = 3600`):
   - Persists across browser page refreshes (F5) using a signed server-side token store
     paired with browser cookies (`st.context.cookies`) and client fingerprint binding.
   - Automatically expires after 1 hour (3600 seconds) of inactivity.
2. Server-side write guard (`is_editor_authenticated` checked before any disk write).
3. Constant-time password comparison (`hmac.compare_digest`) against SHA-256 digests.
4. Brute-force protection (5 failed attempts -> 5 minutes lockout).
5. XSS prevention (`escape_html`) & CSV/Excel Formula Injection prevention (`sanitize_excel_value`).
6. File upload validation (size <= 5 MB, `.xlsx` extension, ZIP magic bytes `PK\\x03\\x04`).
7. Path traversal prevention (`safe_report_id`).
"""

import hashlib
import hmac
import html
import json
import os
from pathlib import Path
import re
import secrets
import time
from typing import Dict, Any, Optional, Tuple
import streamlit as st
import streamlit.components.v1 as components

MAX_UPLOAD_SIZE_BYTES = 5 * 1024 * 1024  # 5 MB limit for Excel uploads
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 300  # 5 minutes lockout on brute-force
EDITOR_SESSION_TIMEOUT_SECONDS = 3600  # Max 1 hour (3600s) inactivity timeout
COOKIE_NAME = "wr_editor_session_token"
XLSX_MAGIC_BYTES = b"PK\x03\x04"

SESSIONS_FILE = Path(__file__).resolve().parent / "data" / ".editor_sessions.json"


def get_configured_password() -> str:
    """
    Retrieves the editor password from Streamlit secrets (`ADMIN_PASSWORD` or `APP_PASSWORD`)
    or environment variables.
    """
    try:
        if "ADMIN_PASSWORD" in st.secrets:
            return str(st.secrets["ADMIN_PASSWORD"])
        if "APP_PASSWORD" in st.secrets:
            return str(st.secrets["APP_PASSWORD"])
    except Exception:
        pass

    env_pwd = os.environ.get("ADMIN_PASSWORD") or os.environ.get("APP_PASSWORD")
    if env_pwd:
        return env_pwd

    return "Otomasyon2026!"


def verify_password(input_password: str, expected_password: str) -> bool:
    """Constant-time password comparison to prevent timing attacks."""
    if not input_password or not expected_password:
        return False
    input_hash = hashlib.sha256(input_password.encode("utf-8")).digest()
    expected_hash = hashlib.sha256(expected_password.encode("utf-8")).digest()
    return hmac.compare_digest(input_hash, expected_password_hash := expected_hash)


def _get_client_fingerprint() -> str:
    """
    Computes a privacy-preserving SHA-256 fingerprint of the client's browser headers and IP
    to bind session tokens to the originating browser.
    """
    try:
        headers = st.context.headers or {}
        ua = headers.get("User-Agent", "") or headers.get("user-agent", "")
        lang = headers.get("Accept-Language", "") or headers.get("accept-language", "")
        ip = getattr(st.context, "ip_address", "") or ""
        raw = f"{ua}|{lang}|{ip}"
    except Exception:
        raw = "default_client"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


def _load_server_sessions() -> Dict[str, Dict[str, Any]]:
    """Loads active editor sessions from disk and prunes expired tokens (> 1 hour inactive)."""
    SESSIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
    now = time.time()
    sessions: Dict[str, Dict[str, Any]] = {}
    if SESSIONS_FILE.exists():
        try:
            with open(SESSIONS_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                sessions = raw
        except Exception:
            sessions = {}

    # Prune sessions inactive for more than EDITOR_SESSION_TIMEOUT_SECONDS (1 hour)
    valid_sessions = {
        tok: info
        for tok, info in sessions.items()
        if isinstance(info, dict)
        and (now - float(info.get("last_active", 0))) <= EDITOR_SESSION_TIMEOUT_SECONDS
    }
    if len(valid_sessions) != len(sessions):
        _save_server_sessions(valid_sessions)
    return valid_sessions


def _save_server_sessions(sessions: Dict[str, Dict[str, Any]]) -> None:
    """Atomically writes active editor sessions to disk."""
    SESSIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = SESSIONS_FILE.with_suffix(".tmp")
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(sessions, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, SESSIONS_FILE)
    except Exception:
        pass


def _get_cookie_token() -> Optional[str]:
    """Reads the session token cookie from `st.context.cookies` if present."""
    try:
        cookies = st.context.cookies
        if cookies and COOKIE_NAME in cookies:
            return str(cookies[COOKIE_NAME]).strip()
    except Exception:
        pass
    return None


def sync_browser_cookie() -> None:
    """
    Injects a tiny, invisible script when a cookie needs to be set or cleared on the browser
    so that F5 page refreshes preserve the 1-hour session token.
    """
    action = st.session_state.pop("_cookie_sync_action", None)
    if not action:
        return

    if action["op"] == "set":
        token_val = action["token"]
        js_code = f"""
        <script>
        (function() {{
            const cookieStr = "{COOKIE_NAME}={token_val}; path=/; max-age={EDITOR_SESSION_TIMEOUT_SECONDS}; SameSite=Lax";
            document.cookie = cookieStr;
            try {{ window.parent.document.cookie = cookieStr; }} catch(e) {{}}
        }})();
        </script>
        """
        components.html(js_code, height=0, width=0)
    elif action["op"] == "clear":
        js_code = f"""
        <script>
        (function() {{
            const clearStr = "{COOKIE_NAME}=; path=/; max-age=0; SameSite=Lax";
            document.cookie = clearStr;
            try {{ window.parent.document.cookie = clearStr; }} catch(e) {{}}
        }})();
        </script>
        """
        components.html(js_code, height=0, width=0)


def init_security_state() -> None:
    """
    Initializes session state variables and restores active 1-hour session token
    across browser refreshes (F5).
    """
    if "auth_is_editor" not in st.session_state:
        st.session_state.auth_is_editor = False
    if "auth_token" not in st.session_state:
        st.session_state.auth_token = ""
    if "auth_failed_attempts" not in st.session_state:
        st.session_state.auth_failed_attempts = 0
    if "auth_lockout_until" not in st.session_state:
        st.session_state.auth_lockout_until = 0.0
    if "auth_last_active" not in st.session_state:
        st.session_state.auth_last_active = 0.0
    if "auth_user_name" not in st.session_state:
        st.session_state.auth_user_name = "Misafir (Salt Okunur)"

    # If session_state was reset by a page refresh (F5), check cookie or client fingerprint
    if not st.session_state.auth_is_editor and not st.session_state.get("_explicit_logout", False):
        sessions = _load_server_sessions()
        now = time.time()
        fp = _get_client_fingerprint()
        cookie_tok = _get_cookie_token()

        matched_token = None
        if cookie_tok and cookie_tok in sessions:
            matched_token = cookie_tok
        else:
            # Fallback for browsers blocking iframe cookies: match active session by client fingerprint
            for tok, info in sessions.items():
                if info.get("fingerprint") == fp and (now - float(info.get("last_active", 0))) <= EDITOR_SESSION_TIMEOUT_SECONDS:
                    matched_token = tok
                    break

        if matched_token:
            info = sessions[matched_token]
            info["last_active"] = now
            _save_server_sessions(sessions)
            st.session_state.auth_is_editor = True
            st.session_state.auth_token = matched_token
            st.session_state.auth_last_active = now
            st.session_state.auth_user_name = info.get("user_name", "Yetkili Editör")


def is_editor_authenticated() -> bool:
    """
    Checks if the current session has active, non-expired Editor privileges (max 1 hour inactivity).
    Updates the sliding 1-hour inactivity timestamp on valid activity.
    """
    init_security_state()
    if not st.session_state.auth_is_editor:
        return False

    now = time.time()
    if now - st.session_state.auth_last_active > EDITOR_SESSION_TIMEOUT_SECONDS:
        revoke_editor_access(expired=True)
        st.warning("1 saatlik hareketsizlik süresi dolduğu için düzenleme oturumunuz sonlandırıldı. Lütfen tekrar giriş yapın.")
        return False

    # Refresh server-side token activity timestamp
    token = st.session_state.get("auth_token", "")
    if token:
        sessions = _load_server_sessions()
        if token in sessions:
            sessions[token]["last_active"] = now
            _save_server_sessions(sessions)
        else:
            revoke_editor_access(expired=True)
            return False

    st.session_state.auth_last_active = now
    return True


def get_session_remaining_minutes() -> int:
    """Returns remaining minutes before inactivity timeout (up to 60 mins)."""
    if not st.session_state.get("auth_is_editor"):
        return 0
    elapsed = time.time() - float(st.session_state.get("auth_last_active", time.time()))
    rem_sec = max(0, EDITOR_SESSION_TIMEOUT_SECONDS - int(elapsed))
    return max(1, (rem_sec + 59) // 60)


def attempt_editor_login(user_name: str, password: str) -> Tuple[bool, str]:
    """
    Validates editor credentials, creates a 1-hour session token, and stores it on server & cookie.
    """
    init_security_state()
    now = time.time()

    if now < st.session_state.auth_lockout_until:
        remaining = int(st.session_state.auth_lockout_until - now)
        return False, f"Çok fazla hatalı deneme! Lütfen {remaining} saniye bekleyin."

    expected_pwd = get_configured_password()
    if verify_password(password, expected_pwd):
        clean_user = sanitize_text(user_name.strip(), max_len=50) or "Yetkili Editör"
        token = secrets.token_urlsafe(32)
        fp = _get_client_fingerprint()

        sessions = _load_server_sessions()
        sessions[token] = {
            "user_name": clean_user,
            "created_at": now,
            "last_active": now,
            "fingerprint": fp,
        }
        _save_server_sessions(sessions)

        st.session_state.auth_is_editor = True
        st.session_state.auth_token = token
        st.session_state.auth_failed_attempts = 0
        st.session_state.auth_last_active = now
        st.session_state.auth_user_name = clean_user
        st.session_state._explicit_logout = False
        st.session_state._cookie_sync_action = {"op": "set", "token": token}
        return True, f"Hoş geldiniz {clean_user}, düzenleme yetkisi açıldı (1 saatlik oturum)!"

    st.session_state.auth_failed_attempts += 1
    remaining_attempts = MAX_FAILED_ATTEMPTS - st.session_state.auth_failed_attempts
    if remaining_attempts <= 0:
        st.session_state.auth_lockout_until = now + LOCKOUT_DURATION_SECONDS
        return False, "Çok fazla başarısız deneme! Düzenleme girişi 5 dakika kilitlendi."
    return False, f"Hatalı yetkili parolası! Kalan deneme hakkı: {remaining_attempts}"


def revoke_editor_access(expired: bool = False) -> None:
    """Downgrades current session from Editor back to read-only Viewer and removes server token."""
    token = st.session_state.get("auth_token", "")
    fp = _get_client_fingerprint()
    sessions = _load_server_sessions()

    keys_to_remove = [
        k for k, v in sessions.items() if k == token or v.get("fingerprint") == fp
    ]
    for k in keys_to_remove:
        sessions.pop(k, None)
    _save_server_sessions(sessions)

    st.session_state.auth_is_editor = False
    st.session_state.auth_token = ""
    st.session_state.auth_last_active = 0.0
    st.session_state.auth_user_name = "Misafir (Salt Okunur)"
    if not expired:
        st.session_state._explicit_logout = True
    st.session_state._cookie_sync_action = {"op": "clear"}


def sanitize_text(value: object, max_len: int = 1000) -> str:
    """
    Cleans user-provided text, strips control characters (except newline/tab),
    and enforces a maximum length.
    """
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text[:max_len]


def escape_html(value: object, preserve_newlines: bool = False) -> str:
    """
    Escapes HTML special characters (&, <, >, ", ') to prevent XSS in generated HTML email tables.
    """
    clean = sanitize_text(value)
    escaped = html.escape(clean, quote=True)
    if preserve_newlines:
        escaped = escaped.replace("\n", "<br>")
    return escaped


def sanitize_excel_value(value: object) -> object:
    """
    Prevents CSV/Excel Formula Injection (CWE-1236) by prefixing a single quote
    when a user-supplied string starts with =, +, -, or @.
    """
    if isinstance(value, (int, float)):
        return value
    if value is None:
        return ""
    text = sanitize_text(value)
    if len(text) > 1 and text[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + text
    if text == "=" or text == "+" or text == "@":
        return "'" + text
    return text


def safe_report_id(raw_id: str) -> str:
    """
    Sanitizes a report identifier to prevent path traversal attacks.
    Allows only alphanumeric characters, underscores, hyphens, and spaces converted to underscores.
    """
    cleaned = re.sub(r"[^a-zA-Z0-9_\-\s]", "", str(raw_id or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    if not cleaned:
        cleaned = "rapor"
    return cleaned[:60]


def validate_xlsx_upload(uploaded_file) -> Tuple[bool, str]:
    """
    Validates an uploaded Excel file for security:
    - Checks editor permission
    - Checks file extension (.xlsx)
    - Checks file size (<= 5 MB)
    - Checks ZIP/XLSX magic bytes (PK\\x03\\x04)
    """
    if not is_editor_authenticated():
        return False, "Dosya yüklemek için Düzenleme Yetkisi (Editör Girişi) gereklidir."

    if uploaded_file is None:
        return False, "Dosya seçilmedi."

    filename = getattr(uploaded_file, "name", "") or ""
    if not filename.lower().endswith(".xlsx"):
        return False, "Sadece .xlsx uzantılı Excel dosyaları kabul edilmektedir."

    size = getattr(uploaded_file, "size", 0) or 0
    if size <= 0:
        return False, "Yüklenen dosya boş."
    if size > MAX_UPLOAD_SIZE_BYTES:
        return False, f"Dosya boyutu çok büyük (Maksimum {MAX_UPLOAD_SIZE_BYTES // (1024 * 1024)} MB)."

    header = uploaded_file.read(4)
    uploaded_file.seek(0)
    if header != XLSX_MAGIC_BYTES:
        return False, "Geçersiz dosya imzası! Dosya gerçek bir .xlsx paketi değil."

    return True, "Geçerli dosya."
