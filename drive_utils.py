import io
import hashlib
from pathlib import Path
from datetime import datetime
from typing import Optional

import streamlit as st

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as google_build
    from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
except Exception:
    service_account = None
    google_build = None
    MediaIoBaseUpload = None
    MediaIoBaseDownload = None

from config import (
    KST_TZ, GOOGLE_SERVICE_ACCOUNT_JSON, GOOGLE_DRIVE_SALES_ROOT_FOLDER_ID,
    GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME, GOOGLE_DRIVE_XLSX_MIME, GOOGLE_DRIVE_FOLDER_MIME,
    load_service_account_info,
)

def _drive_date_folder_name(dt: Optional[datetime] = None) -> str:
    dt = dt or datetime.now(KST_TZ)
    return f"{dt.month}.{dt.day}"


def _drive_year_folder_name(dt: Optional[datetime] = None) -> str:
    dt = dt or datetime.now(KST_TZ)
    return str(dt.year)


def _drive_escape_query_value(value: str) -> str:
    return str(value).replace("\\", "\\\\").replace("'", "\\'")


def _safe_drive_filename(filename: str) -> str:
    name = (filename or "업로드엑셀.xlsx").strip()
    name = name.replace("/", "_").replace("\\", "_")
    return name or "업로드엑셀.xlsx"


def _get_google_drive_service():
    if service_account is None or google_build is None or MediaIoBaseUpload is None:
        raise RuntimeError("Google Drive 라이브러리가 설치되지 않았습니다. requirements.txt를 확인해 주세요.")
    info = load_service_account_info()
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return google_build("drive", "v3", credentials=creds, cache_discovery=False)


def _drive_list_files(service, query: str, page_size: int = 20) -> list[dict]:
    resp = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id,name,mimeType,modifiedTime,parents)",
        pageSize=page_size,
        orderBy="modifiedTime desc",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    return resp.get("files", []) or []


def _drive_list_files_all(service, query: str, page_size: int = 100) -> list[dict]:
    """Google Drive 검색 결과를 페이지 끝까지 모두 가져옵니다."""
    out: list[dict] = []
    page_token = None
    while True:
        resp = service.files().list(
            q=query,
            spaces="drive",
            fields="nextPageToken,files(id,name,mimeType,modifiedTime,parents,size)",
            pageSize=page_size,
            orderBy="name",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageToken=page_token,
        ).execute()
        out.extend(resp.get("files", []) or [])
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return out


def _drive_download_file_bytes(service, file_id: str) -> bytes:
    if MediaIoBaseDownload is None:
        raise RuntimeError("Google Drive 다운로드 라이브러리가 설치되지 않았습니다. requirements.txt를 확인해 주세요.")

    try:
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    except TypeError:
        request = service.files().get_media(fileId=file_id)

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue()


def _drive_find_folder_by_name(service, folder_name: str, parent_id: Optional[str] = None) -> Optional[str]:
    name_q = _drive_escape_query_value(folder_name)
    q = f"mimeType='{GOOGLE_DRIVE_FOLDER_MIME}' and trashed=false and name='{name_q}'"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    files = _drive_list_files(service, q, page_size=10)
    return files[0]["id"] if files else None


def _drive_create_folder(service, folder_name: str, parent_id: Optional[str] = None) -> str:
    metadata = {"name": folder_name, "mimeType": GOOGLE_DRIVE_FOLDER_MIME}
    if parent_id:
        metadata["parents"] = [parent_id]
    created = service.files().create(
        body=metadata,
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return created["id"]


def _drive_get_or_create_sales_root_folder(service) -> str:
    if GOOGLE_DRIVE_SALES_ROOT_FOLDER_ID:
        return GOOGLE_DRIVE_SALES_ROOT_FOLDER_ID
    folder_id = _drive_find_folder_by_name(service, GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME)
    if folder_id:
        return folder_id
    return _drive_create_folder(service, GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME)


def _drive_get_or_create_year_folder(service, root_folder_id: str, dt: Optional[datetime] = None) -> tuple[str, str]:
    """판매내역/연도 폴더를 찾고, 없으면 생성합니다. 반환: (연도폴더ID, 연도폴더명)."""
    year_folder_name = _drive_year_folder_name(dt)
    year_folder_id = _drive_find_folder_by_name(service, year_folder_name, parent_id=root_folder_id)
    if not year_folder_id:
        year_folder_id = _drive_create_folder(service, year_folder_name, parent_id=root_folder_id)
    return year_folder_id, year_folder_name


def _drive_find_year_date_folder(service, root_folder_id: str, dt_obj) -> tuple[str, str, Optional[str]]:
    """판매내역/연도/월.일 날짜 폴더를 찾습니다. 없으면 date_folder_id는 None입니다."""
    year_folder_name = str(dt_obj.year)
    date_folder_name = f"{dt_obj.month}.{dt_obj.day}"
    year_folder_id = _drive_find_folder_by_name(service, year_folder_name, parent_id=root_folder_id)
    if not year_folder_id:
        return year_folder_name, date_folder_name, None
    date_folder_id = _drive_find_folder_by_name(service, date_folder_name, parent_id=year_folder_id)
    return year_folder_name, date_folder_name, date_folder_id


def _drive_upload_or_replace_file(service, parent_folder_id: str, filename: str, file_bytes: bytes) -> tuple[str, str]:
    safe_name = _safe_drive_filename(filename)
    name_q = _drive_escape_query_value(safe_name)
    q = f"trashed=false and name='{name_q}' and '{parent_folder_id}' in parents"
    files = _drive_list_files(service, q, page_size=10)

    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes),
        mimetype=GOOGLE_DRIVE_XLSX_MIME,
        resumable=False,
    )

    if files:
        file_id = files[0]["id"]
        updated = service.files().update(
            fileId=file_id,
            body={"name": safe_name, "mimeType": GOOGLE_DRIVE_XLSX_MIME},
            media_body=media,
            fields="id,name",
            supportsAllDrives=True,
        ).execute()
        return updated["id"], "updated"

    created = service.files().create(
        body={"name": safe_name, "parents": [parent_folder_id], "mimeType": GOOGLE_DRIVE_XLSX_MIME},
        media_body=media,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    return created["id"], "created"


def save_excel_upload_to_drive_once(filename: str, file_bytes: bytes, target_dt: Optional[datetime] = None) -> tuple[bool, str]:
    """파일명+내용 해시 기준으로 세션 내 1회만 Google Drive에 저장합니다."""
    safe_name = _safe_drive_filename(filename)
    target_dt = target_dt or datetime.now(KST_TZ)
    year_folder_name = _drive_year_folder_name(target_dt)
    date_folder_name = _drive_date_folder_name(target_dt)
    digest_src = f"{year_folder_name}/{date_folder_name}/{safe_name}".encode("utf-8") + b"\0" + (file_bytes or b"")
    digest = hashlib.sha256(digest_src).hexdigest()
    state_key = f"drive_upload_excel_results_{digest}"

    if st.session_state.get(state_key):
        return True, st.session_state.get(f"{state_key}_msg", "이미 Google Drive에 저장된 업로드 파일입니다.")

    if not file_bytes:
        return False, "업로드 파일 내용이 비어 있어 Google Drive에 저장하지 않았습니다."

    try:
        service = _get_google_drive_service()
        root_id = _drive_get_or_create_sales_root_folder(service)

        # ✅ 저장 경로: 판매내역/연도/월.일/파일명.xlsx
        # 연도 폴더가 없으면 자동 생성하고, 날짜 폴더도 없으면 자동 생성합니다.
        year_folder_id, year_folder_name = _drive_get_or_create_year_folder(service, root_id, target_dt)
        date_folder_id = _drive_find_folder_by_name(service, date_folder_name, parent_id=year_folder_id)
        if not date_folder_id:
            date_folder_id = _drive_create_folder(service, date_folder_name, parent_id=year_folder_id)

        _, action = _drive_upload_or_replace_file(service, date_folder_id, safe_name, file_bytes)
        action_text = "덮어쓰기 완료" if action == "updated" else "새 파일 저장 완료"
        msg = f"Google Drive 저장 {action_text}: {GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME}/{year_folder_name}/{date_folder_name}/{safe_name}"
        st.session_state[state_key] = True
        st.session_state[f"{state_key}_msg"] = msg
        return True, msg
    except Exception as e:
        return False, f"Google Drive 저장 실패: {e}"

