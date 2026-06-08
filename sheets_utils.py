import re
import streamlit as st

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as google_build
except Exception:
    service_account = None
    google_build = None

from config import load_service_account_info

@st.cache_resource(show_spinner=False)
def _get_google_sheets_service():
    if service_account is None or google_build is None:
        raise RuntimeError("Google Sheets 라이브러리가 설치되지 않았습니다. requirements.txt를 확인해 주세요.")
    info = load_service_account_info()
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return google_build("sheets", "v4", credentials=creds, cache_discovery=False)


def _extract_spreadsheet_id(value: str) -> str:
    """Google Sheet URL 또는 ID에서 spreadsheet_id만 추출합니다."""
    raw = (value or "").strip()
    if not raw:
        return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
    if m:
        return m.group(1)
    return raw.split("?", 1)[0].split("#", 1)[0].strip()

