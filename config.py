import os
import json
from pathlib import Path
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from typing import Optional

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    LongTable,
    TableStyle,
    Paragraph,
    Spacer,
    KeepTogether,
    HRFlowable,
)
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# =====================================================
# TIMEZONE
# =====================================================
KST_TZ = ZoneInfo("Asia/Seoul")
KST = timezone(timedelta(hours=9))  # for legacy functions


def now_prefix_kst() -> str:
    return datetime.now(KST).strftime("%Y%m%d_%H%M%S")


# =====================================================
# PATHS / STORAGE
# =====================================================
# 배포환경(Render 등)에서 재시작/재배포 후에도 설정/룰/내보내기 데이터가 유지되도록,
# 영구 저장 루트 폴더를 환경변수/디스크 마운트 경로로 자동 선택합니다.
#
# 권장(Render):
#  - Persistent Disk mount path: /var/data
#  - Environment Variable: APP_DATA_DIR=/var/data
#
# 선택 우선순위:
#  1) APP_DATA_DIR 환경변수
#  2) /var/data 가 존재하고 쓰기 가능하면 사용
#  3) 현재 작업 폴더(로컬 실행용)
_env_dir = (os.environ.get("APP_DATA_DIR") or "").strip()

def _pick_writable_dir(cands: list[Path]) -> Path:
    for p in cands:
        try:
            p.mkdir(parents=True, exist_ok=True)
            test = p / ".write_test"
            with open(test, "w", encoding="utf-8") as f:
                f.write("ok")
            try:
                test.unlink()
            except Exception:
                pass
            return p
        except Exception:
            continue
    return Path(".")

_candidates: list[Path] = []
if _env_dir:
    _candidates.append(Path(_env_dir))
_candidates.extend([Path("/var/data"), Path(".")])

APP_DATA_DIR = _pick_writable_dir(_candidates)
print(f"[BOOT] APP_DATA_DIR_ENV='{_env_dir}' -> USING='{APP_DATA_DIR}'")

# (1) 재고관리 저장
INVENTORY_FILE = str(APP_DATA_DIR / "inventory.csv")
EXPORT_ROOT = str(APP_DATA_DIR / "exports")

# (2) PACK/BOX/EA 규칙(제품별 합계 계산용)
RULES_FILE = str(APP_DATA_DIR / "rules.txt")
COUNT_UNITS = ["개", "통", "팩", "봉"]

# (3) 2번 코드(엑셀 업로드/매칭 규칙) 데이터 저장
DATA_DIR = APP_DATA_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

MAPPING_PATH = DATA_DIR / "name_mappings.json"
EXPR_RULES_PATH = DATA_DIR / "expression_rules.json"
BACKUP_DIR = DATA_DIR / "rules_backup"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

# ✅ TC 설정 저장 파일 (프로그램 껐다 켜도 유지)
TC_SETTINGS_PATH = DATA_DIR / "tc_settings.json"

# ✅ 스티커 제외 설정 저장 파일 (프로그램 껐다 켜도 유지)
STICKER_SETTINGS_PATH = DATA_DIR / "sticker_settings.json"

# ✅ 레포(앱 폴더)에 "TC주문_등록양식.xlsx" 파일을 같이 올려두면 업로드 없이 자동 사용
TC_TEMPLATE_DEFAULT_PATH = Path("TC주문_등록양식.xlsx")

# ✅ SmartStore 엑셀 비번
EXCEL_PASSWORD = "0000"

# (4) 2번 구글시트 Apps Script 웹앱 동기화
GOOGLE_SYNC_WEBAPP_URL = (os.environ.get("GOOGLE_SYNC_WEBAPP_URL") or "").strip()
GOOGLE_SYNC_TOKEN = (os.environ.get("GOOGLE_SYNC_TOKEN") or "").strip()
try:
    GOOGLE_SYNC_TIMEOUT_SEC = int((os.environ.get("GOOGLE_SYNC_TIMEOUT_SEC") or "20").strip())
except Exception:
    GOOGLE_SYNC_TIMEOUT_SEC = 20

# -------------------- Google Drive upload for Excel upload page --------------------
# 엑셀 업로드 페이지에서 업로드한 원본 .xlsx 파일을 Google Drive에 저장합니다.
# 저장 위치: 판매내역/2026/4.24/업로드파일명.xlsx
GOOGLE_SERVICE_ACCOUNT_JSON = (os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") or "").strip()
GOOGLE_DRIVE_SALES_ROOT_FOLDER_ID = (os.environ.get("GOOGLE_DRIVE_SALES_ROOT_FOLDER_ID") or "").strip()
GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME = (os.environ.get("GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME") or "판매내역").strip() or "판매내역"
GOOGLE_DRIVE_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GOOGLE_DRIVE_FOLDER_MIME = "application/vnd.google-apps.folder"

# 매출계산 결과를 기록할 Google Sheet ID/URL (선택)
# - 예: https://docs.google.com/spreadsheets/d/시트ID/edit
# - 해당 월 시트(예: 4월)가 없으면 기록하지 않고 계산 결과만 보여줍니다.
SALES_RESULT_SPREADSHEET_ID = (os.environ.get("SALES_RESULT_SPREADSHEET_ID") or "").strip()


def load_service_account_info() -> dict:
    raw = GOOGLE_SERVICE_ACCOUNT_JSON
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON 환경변수가 비어 있습니다.")
    if raw.startswith("{"):
        return json.loads(raw)
    path = Path(raw)
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON 값이 JSON 문자열도, 파일 경로도 아닙니다.")
