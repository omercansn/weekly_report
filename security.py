"""
Security & Role-Based Access Control (RBAC) Module for Weekly Automation Report App.

Access Model:
- VIEWER (Default for everyone visiting the link):
  Can view all reports, switch between weeks, copy email tables (Sheet 1-4 & Combined),
  and download the Excel (.xlsx) file. Cannot modify data, create weeks, or upload files.
- EDITOR (Password-authenticated via `st.secrets["APP_PASSWORD"]` or `ADMIN_PASSWORD`):
  Can edit all sheets, add/delete rows, create new weekly reports, and import Excel files.

Security Controls:
1. Server-side write guard (`is_editor_authenticated` checked before any disk write).
2. Constant-time password comparison (`hmac.compare_digest`) against SHA-256 digests.
3. Brute-force protection (5 failed attempts -> 5 minutes lockout).
4. Editor session inactivity timeout (2 hours).
5. XSS prevention (`escape_html`) & CSV/Excel Formula Injection prevention (`sanitize_excel_value`).
6. File upload validation (size <= 5 MB, `.xlsx` extension, ZIP magic bytes `PK\\x03\\x04`).
7. Path traversal prevention (`safe_report_id`).
"""

import hashlib
import hmac
import html
import os
import re
import time
from typing import Tuple
import streamlit as st

MAX_UPLOAD_SIZE_BYTES = 5 * 1024 * 1024  # 5 MB limit for Excel uploads
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 300  # 5 minutes lockout on brute-force
EDITOR_SESSION_TIMEOUT_SECONDS = 3600 * 2  # 2 hours editor inactivity timeout
XLSX_MAGIC_BYTES = b"PK\x03\x04"


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
    return hmac.compare_digest(input_hash, expected_hash)


def init_security_state() -> None:
    """Initializes session state variables for role-based access control."""
    if "auth_is_editor" not in st.session_state:
        st.session_state.auth_is_editor = False
    if "auth_failed_attempts" not in st.session_state:
        st.session_state.auth_failed_attempts = 0
    if "auth_lockout_until" not in st.session_state:
        st.session_state.auth_lockout_until = 0.0
    if "auth_last_active" not in st.session_state:
        st.session_state.auth_last_active = 0.0
    if "auth_user_name" not in st.session_state:
        st.session_state.auth_user_name = "Misafir (Salt Okunur)"


def is_editor_authenticated() -> bool:
    """
    Checks if the current session has active, non-expired Editor privileges.
    Automatically downgrades to Viewer if inactivity timeout has elapsed.
    """
    init_security_state()
    if not st.session_state.auth_is_editor:
        return False

    now = time.time()
    if now - st.session_state.auth_last_active > EDITOR_SESSION_TIMEOUT_SECONDS:
        st.session_state.auth_is_editor = False
        st.session_state.auth_user_name = "Misafir (Salt Okunur)"
        st.warning("Düzenleme oturumunuzun süresi doldu, Salt Okunur moda geçildi.")
        return False

    st.session_state.auth_last_active = now
    return True


def attempt_editor_login(user_name: str, password: str) -> Tuple[bool, str]:
    """
    Validates editor credentials with brute-force lockout protection.
    Returns (success, message).
    """
    init_security_state()
    now = time.time()

    if now < st.session_state.auth_lockout_until:
        remaining = int(st.session_state.auth_lockout_until - now)
        return False, f"Çok fazla hatalı deneme! Lütfen {remaining} saniye bekleyin."

    expected_pwd = get_configured_password()
    if verify_password(password, expected_pwd):
        st.session_state.auth_is_editor = True
        st.session_state.auth_failed_attempts = 0
        st.session_state.auth_last_active = now
        clean_user = sanitize_text(user_name.strip(), max_len=50) or "Yetkili Editör"
        st.session_state.auth_user_name = clean_user
        return True, f"Hoş geldiniz {clean_user}, düzenleme yetkisi açıldı!"

    st.session_state.auth_failed_attempts += 1
    remaining_attempts = MAX_FAILED_ATTEMPTS - st.session_state.auth_failed_attempts
    if remaining_attempts <= 0:
        st.session_state.auth_lockout_until = now + LOCKOUT_DURATION_SECONDS
        return False, "Çok fazla başarısız deneme! Düzenleme girişi 5 dakika kilitlendi."
    return False, f"Hatalı yetkili parolası! Kalan deneme hakkı: {remaining_attempts}"


def revoke_editor_access() -> None:
    """Downgrades current session from Editor back to read-only Viewer."""
    st.session_state.auth_is_editor = False
    st.session_state.auth_last_active = 0.0
    st.session_state.auth_user_name = "Misafir (Salt Okunur)"


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
