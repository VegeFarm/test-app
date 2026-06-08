import io
import re
from datetime import datetime, timedelta
from typing import Optional, Tuple, List, Dict

import pandas as pd
import streamlit as st

try:
    import msoffcrypto
except ModuleNotFoundError:
    msoffcrypto = None

from config import (
    KST_TZ, EXCEL_PASSWORD, GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME,
    SALES_RESULT_SPREADSHEET_ID, GOOGLE_DRIVE_XLSX_MIME,
)
from drive_utils import (
    _get_google_drive_service, _drive_get_or_create_sales_root_folder,
    _drive_find_year_date_folder, _drive_list_files_all, _drive_download_file_bytes,
)
from sheets_utils import _get_google_sheets_service, _extract_spreadsheet_id

def _sales_is_zip_xlsx(file_bytes: bytes) -> bool:
    # Normal xlsx starts with PK.. (zip)
    return file_bytes[:4] == b"PK\x03\x04"


def _sales_decrypt_excel_bytes(file_bytes: bytes, password: str = EXCEL_PASSWORD) -> io.BytesIO:
    """
    Returns a BytesIO that can be read by pandas/openpyxl.
    - If file is normal xlsx(zip), returns as-is.
    - If file is encrypted (OLE), decrypts using msoffcrypto.
    """
    if _sales_is_zip_xlsx(file_bytes):
        return io.BytesIO(file_bytes)

    if msoffcrypto is None:
        raise RuntimeError(
            "이 엑셀은 비밀번호로 암호화되어 있어요. requirements.txt에 'msoffcrypto-tool'을 추가해 설치해 주세요."
        )

    decrypted = io.BytesIO()
    office = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
    office.load_key(password=password)
    office.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def _sales_to_number(series: pd.Series) -> pd.Series:
    # 숫자/문자 섞여 있어도 안전하게 숫자로 변환 (콤마, 원, 공백 등 제거)
    return pd.to_numeric(
        series.astype(str).str.replace(r"[^\d\.-]", "", regex=True),
        errors="coerce",
    )


def _sales_normalize_text_series(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .replace({"nan": "", "None": ""})
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def _sales_norm_no_space(x: str) -> str:
    return re.sub(r"\s+", "", str(x or "")).strip()


def _sales_find_col(cols: List[str], candidates: List[str]) -> Optional[str]:
    # 1) exact match
    for c in candidates:
        if c in cols:
            return c

    # 2) normalized match (remove spaces/newlines)
    cols_norm = {_sales_norm_no_space(c): c for c in cols}
    for cand in candidates:
        n = _sales_norm_no_space(cand)
        if n in cols_norm:
            return cols_norm[n]

    # 3) substring match
    for cand in candidates:
        for col in cols:
            if str(cand) and str(cand) in str(col):
                return col

    return None


def _sales_detect_header_row(df: pd.DataFrame, max_scan: int = 30) -> int:
    """
    엑셀 상단에 안내문/요약 등이 섞여 있을 수 있어
    앞쪽 몇 줄 스캔 후 '구매자명/수취인명'이 함께 존재하는 줄을 헤더로 판단.
    """
    must_have = {_sales_norm_no_space("구매자명"), _sales_norm_no_space("수취인명")}

    scan_n = min(max_scan, len(df))
    for r in range(scan_n):
        row_vals = df.iloc[r].astype(str).tolist()
        row_norm_set = set(_sales_norm_no_space(v) for v in row_vals if str(v).strip() != "")
        if must_have.issubset(row_norm_set):
            return r

    return 0


def _sales_read_excel_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    bio = _sales_decrypt_excel_bytes(file_bytes, EXCEL_PASSWORD)
    raw = pd.read_excel(bio, sheet_name=None, header=None, engine="openpyxl")

    sheets: Dict[str, pd.DataFrame] = {}
    for name, df in raw.items():
        if df is None or df.empty:
            continue

        header_row = _sales_detect_header_row(df, max_scan=30)
        header = df.iloc[header_row].astype(str).str.strip().tolist()

        # make header unique (avoid duplicate col names)
        seen = {}
        new_cols = []
        for h in header:
            h2 = (h or "").strip()
            if h2.lower() == "nan" or h2 == "":
                h2 = "col"
            cnt = seen.get(h2, 0)
            new_cols.append(h2 if cnt == 0 else f"{h2}_{cnt}")
            seen[h2] = cnt + 1

        data = df.iloc[header_row + 1 :].copy()
        data.columns = new_cols
        data = data.dropna(how="all").reset_index(drop=True)
        sheets[name] = data

    return sheets


def _sales_compute_from_sheets(sheets: Dict[str, pd.DataFrame]) -> Tuple[float, set]:
    """
    Returns:
      (sum_of_final_order_amount, set_of_unique_keys_with_nonzero_shipping)
    """
    AMOUNT_CANDS = ["최종 상품별 총 주문금액"]
    SHIP_CANDS = ["배송비 합계"]
    BUYER_CANDS = ["구매자명"]
    RECIP_CANDS = ["수취인명"]
    ADDR_CANDS = ["통합배송지", "주소", "배송지", "수취인주소", "수령인주소", "수취인 주소", "수령인 주소"]

    total_amount = 0.0
    nonzero_people_keys: set = set()

    for _, df in sheets.items():
        cols = [str(c).strip() for c in df.columns]

        amount_col = _sales_find_col(cols, AMOUNT_CANDS)
        ship_col = _sales_find_col(cols, SHIP_CANDS)
        buyer_col = _sales_find_col(cols, BUYER_CANDS)
        recip_col = _sales_find_col(cols, RECIP_CANDS)
        addr_col = _sales_find_col(cols, ADDR_CANDS)

        if amount_col is not None:
            amt = _sales_to_number(df[amount_col])
            total_amount += float(amt.sum(skipna=True) or 0.0)

        if ship_col is not None:
            ship = _sales_to_number(df[ship_col]).fillna(0)
            nonzero_mask = ship != 0

            buyer = _sales_normalize_text_series(df[buyer_col]) if buyer_col else pd.Series([""] * len(df))
            recip = _sales_normalize_text_series(df[recip_col]) if recip_col else pd.Series([""] * len(df))
            addr = _sales_normalize_text_series(df[addr_col]) if addr_col else pd.Series([""] * len(df))

            keys = (buyer + "||" + recip + "||" + addr)
            keys = keys[nonzero_mask].dropna()

            # 빈 키 제거
            keys = keys[keys.str.replace("||", "", regex=False).str.strip() != ""]
            nonzero_people_keys.update(keys.tolist())

    return total_amount, nonzero_people_keys


def _sales_compute_from_file_bytes(file_bytes: bytes) -> tuple[float, set]:
    """기존 매출계산 로직을 파일 바이트 기준으로 재사용합니다."""
    sheets = _sales_read_excel_sheets(file_bytes)
    return _sales_compute_from_sheets(sheets)


def _sales_drive_date_folder_name(dt_obj) -> str:
    return f"{dt_obj.month}.{dt_obj.day}"


def _sales_result_month_sheet_name(dt_obj) -> str:
    return f"{dt_obj.month}월"


def _sales_result_date_label(dt_obj) -> str:
    return f"{dt_obj.month}/{dt_obj.day}"


def _sales_result_range_label_if_gap(target_date, previous_md: Optional[tuple[int, int]]) -> str:
    """
    매출기록시트에 새 최신 날짜를 추가할 때, 직전 기록일과 목표일 사이에
    빈 날짜가 있으면 A열 날짜를 범위로 표시합니다.
    예: 직전 4/24 + 목표 4/26 → 4/25~4/26
        직전 4/10 + 목표 4/26 → 4/11~4/26
        직전 4/25 + 목표 4/26 → 4/26
    """
    default_label = _sales_result_date_label(target_date)
    if not previous_md:
        return default_label

    try:
        target_dt = datetime(int(target_date.year), int(target_date.month), int(target_date.day))
        prev_dt = datetime(int(target_date.year), int(previous_md[0]), int(previous_md[1]))
    except Exception:
        return default_label

    # 해당 월 시트 안에서만 범위 표기를 적용합니다.
    if prev_dt.month != target_dt.month or prev_dt >= target_dt:
        return default_label

    # 직전 기록일 바로 다음날이면 범위가 아니라 목표 날짜만 씁니다.
    if (target_dt - prev_dt).days <= 1:
        return default_label

    start_dt = prev_dt + timedelta(days=1)
    return f"{start_dt.month}/{start_dt.day}~{target_dt.month}/{target_dt.day}"


def _sales_quote_sheet_name(sheet_name: str) -> str:
    return "'" + str(sheet_name).replace("'", "''") + "'"


def _sales_norm_date_label(value: str) -> str:
    """시트 A열 날짜값을 M/D 형태로 최대한 맞춰 비교합니다."""
    s = str(value or "").strip()
    if not s:
        return ""

    s2 = re.sub(r"\s+", "", s)
    s2 = s2.replace("년", "/").replace("월", "/").replace("일", "")
    s2 = s2.replace(".", "/").replace("-", "/")

    # 2026/4/25, 26/4/25 같은 연도 포함 형식이면 뒤의 월/일만 사용
    nums = re.findall(r"\d+", s2)
    if len(nums) >= 3:
        return f"{int(nums[-2])}/{int(nums[-1])}"
    if len(nums) >= 2:
        return f"{int(nums[0])}/{int(nums[1])}"

    return s2


def _sales_parse_month_day(value) -> Optional[tuple[int, int]]:
    """A열 날짜값을 (월, 일)로 파싱합니다. 예: 4/25, 4.25, 2026-04-25, 4월 25일"""
    key = _sales_norm_date_label(value)
    nums = re.findall(r"\d+", key)
    if len(nums) >= 2:
        try:
            return int(nums[0]), int(nums[1])
        except Exception:
            return None
    return None


def _sales_get_spreadsheet_sheet_properties(sheets_service, spreadsheet_id: str) -> dict[str, dict]:
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets.properties(sheetId,title)",
    ).execute()
    props = {}
    for s in (meta.get("sheets", []) or []):
        p = s.get("properties", {}) or {}
        title = p.get("title")
        if title:
            props[str(title)] = p
    return props


def _sales_get_spreadsheet_sheet_titles(sheets_service, spreadsheet_id: str) -> list[str]:
    return list(_sales_get_spreadsheet_sheet_properties(sheets_service, spreadsheet_id).keys())


def _sales_write_result_to_month_sheet(
    spreadsheet_id_or_url: str,
    target_date,
    total_amount: float,
    shipping_calc: float,
) -> dict:
    """
    지정 구글시트의 해당월 시트에 A=날짜, B=총 주문금액, C=인원×3,500 결과를 기록합니다.
    - 해당월 시트가 없으면 생성하지 않고 skipped 처리합니다.
    - 날짜 행이 있으면 B/C를 갱신합니다.
    - 날짜 행이 없으면 날짜 순서에 맞는 위치에 행을 삽입합니다.
    - 최신 날짜를 추가할 때 직전 기록일과 목표일 사이에 빈 날짜가 있으면 A열을 범위로 기록합니다.
      예: 직전 4/24 + 목표 4/26 → 4/25~4/26
    - '합계' 행이 있으면 합계 행 아래가 아니라 합계 행 위에 새 날짜 행을 삽입합니다.
    - 새 날짜 행을 만들 때 D/E열 수식은 인접 날짜 행에서 복사해 유지합니다.
    """
    spreadsheet_id = _extract_spreadsheet_id(spreadsheet_id_or_url)
    if not spreadsheet_id:
        return {"status": "skipped", "message": "결과 기록용 Google Sheet ID/URL이 비어 있어 시트 기록은 건너뛰었습니다."}

    sheets_service = _get_google_sheets_service()
    month_sheet = _sales_result_month_sheet_name(target_date)
    date_label = _sales_result_date_label(target_date)
    write_date_label = date_label
    range_label_applied = False
    existing_date_label = ""
    previous_md_for_range: Optional[tuple[int, int]] = None
    target_md = _sales_parse_month_day(date_label)
    target_key = _sales_norm_date_label(date_label)

    sheet_props = _sales_get_spreadsheet_sheet_properties(sheets_service, spreadsheet_id)
    if month_sheet not in sheet_props:
        return {
            "status": "skipped",
            "message": f"'{month_sheet}' 시트가 없어 Google Sheet 기록은 건너뛰었습니다.",
            "month_sheet": month_sheet,
            "date_label": date_label,
        }

    sheet_id = sheet_props[month_sheet].get("sheetId")
    if sheet_id is None:
        raise RuntimeError(f"'{month_sheet}' 시트 ID를 찾지 못했습니다.")

    quoted = _sales_quote_sheet_name(month_sheet)

    # A:C만 읽으면 '합계' 문구가 D/E열에 있는 시트에서 합계 행을 못 찾아서
    # 새 날짜가 합계 행 아래로 들어갈 수 있습니다. A:E까지 읽어 합계 행을 안정적으로 감지합니다.
    read_range = f"{quoted}!A:E"
    resp = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=read_range,
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    values = resp.get("values", []) or []

    target_row = None
    insert_needed = False
    summary_row = None
    last_used_row = 0
    date_rows: list[tuple[int, int, int]] = []

    def _row_values_at(row_no: int) -> list[str]:
        if 1 <= int(row_no) <= len(values):
            return [str(v).strip() for v in (values[int(row_no) - 1] or [])]
        return []

    def _row_has_ac_content(row_no: int) -> bool:
        row_values = _row_values_at(row_no)
        # A:C에 실제 값이 있으면 새 날짜를 바로 덮어쓰지 않고 행 삽입 대상으로 봅니다.
        return any(v for v in row_values[:3])

    def _is_summary_row(row_values: list[str]) -> bool:
        # 시트에서는 "합 계"처럼 글자 사이에 공백을 넣는 경우가 있어 공백 제거 후 판단합니다.
        normalized = [re.sub(r"\s+", "", str(v or "")) for v in row_values]
        return any(("합계" in v) or ("총합계" in v) for v in normalized)

    # 1) 합계 행/기존 날짜 행/마지막 사용 행 확인
    for idx, row in enumerate(values, start=1):
        row_values = [str(v).strip() for v in row]
        if any(row_values):
            last_used_row = idx

        if summary_row is None and _is_summary_row(row_values):
            summary_row = idx

        # 합계 행 이후는 날짜 데이터 영역으로 보지 않습니다.
        if summary_row is not None and idx >= summary_row:
            continue

        a_val = row[0] if row else ""
        md = _sales_parse_month_day(a_val)
        if md:
            date_rows.append((idx, md[0], md[1]))

        if _sales_norm_date_label(a_val) == target_key:
            target_row = idx
            existing_date_label = str(a_val or "").strip()
            break

    # 2) 기존 날짜가 없으면 날짜 순서 위치 찾기
    if target_row is None:
        if target_md:
            for row_idx, month_num, day_num in date_rows:
                if (month_num, day_num) > target_md:
                    # 중간 날짜는 다음 날짜 행 바로 위에 새 행을 삽입합니다.
                    target_row = row_idx
                    insert_needed = True
                    break

        if target_row is None:
            last_date_info = max(date_rows, key=lambda x: x[0], default=None)
            if last_date_info:
                last_date_row, last_m, last_d = last_date_info
                previous_md_for_range = (int(last_m), int(last_d))
                # 최신 날짜는 합계 바로 위가 아니라 마지막 날짜의 바로 다음 행에 넣습니다.
                # 예: 4/24 다음 4/26을 넣으면 빈 행들이 있더라도 4/24 바로 아래 행에 기록합니다.
                target_row = int(last_date_row) + 1
                # 단, 그 행의 A:C에 이미 값이 있으면 덮어쓰지 않고 행 삽입합니다.
                # 합계 행이 바로 다음 행인 경우도 합계 위에 삽입됩니다.
                insert_needed = bool(_row_has_ac_content(target_row))
            elif summary_row:
                # 날짜 데이터가 아직 없고 합계 행만 있는 새 월/템플릿 시트에서는
                # 합계 행 위에 바로 행을 삽입하지 않습니다.
                # 먼저 입력 영역(기본 4행~합계행 바로 위)의 비어 있는 첫 행을 사용합니다.
                # 예: 템플릿이 4~34행 빈칸, 35행 합계라면 첫 자동 기록은 4행에 작성되고,
                #     4~34행이 모두 찼을 때만 35행 합계 위에 새 행을 삽입합니다.
                first_data_row = 4
                blank_row = None
                for row_no in range(first_data_row, int(summary_row)):
                    if not _row_has_ac_content(row_no):
                        blank_row = row_no
                        break

                if blank_row is not None:
                    target_row = blank_row
                    insert_needed = False
                else:
                    target_row = summary_row
                    insert_needed = True
            else:
                # 합계 행도 날짜 행도 없으면 현재 데이터 마지막 행 아래에 작성합니다.
                target_row = max(last_used_row + 1, 1)
                insert_needed = False

        action = "created"
    else:
        action = "updated"
        # 기존 행이 4/25~4/26 같은 범위 날짜로 되어 있으면, 갱신할 때 A열 표기를 유지합니다.
        if existing_date_label:
            write_date_label = existing_date_label

    if action == "created" and previous_md_for_range:
        write_date_label = _sales_result_range_label_if_gap(target_date, previous_md_for_range)
        range_label_applied = (write_date_label != date_label)

    # 새 날짜 행의 D/E 수식을 복사할 기준 행을 정합니다.
    # 우선 이전 날짜 행을 사용하고, 이전 날짜가 없으면 다음 날짜 행을 사용합니다.
    formula_source_row = None
    if action == "created":
        previous_date_rows = [r for r, _m, _d in date_rows if r < target_row]
        next_date_rows = [r for r, _m, _d in date_rows if r >= target_row]
        if previous_date_rows:
            formula_source_row = max(previous_date_rows)
        elif next_date_rows:
            # 행 삽입 후에는 기존 target_row 위치의 날짜 행이 한 줄 아래로 밀립니다.
            formula_source_row = min(next_date_rows) + (1 if insert_needed else 0)

    # 3) 날짜 순서 중간/합계 행 위에 넣어야 하면 실제 행 삽입
    if insert_needed:
        insert_req = {
            "insertDimension": {
                "range": {
                    "sheetId": int(sheet_id),
                    "dimension": "ROWS",
                    "startIndex": int(target_row) - 1,
                    "endIndex": int(target_row),
                },
                "inheritFromBefore": bool(target_row > 1),
            }
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [insert_req]},
        ).execute()

    # 4) 새 날짜 행이면 D/E열 수식과 서식을 인접 날짜 행에서 복사합니다.
    # Google Sheets의 copyPaste는 상대참조 수식을 새 행 번호에 맞게 조정해 줍니다.
    formula_copied = False
    if action == "created" and formula_source_row and formula_source_row != target_row:
        copy_requests = [
            {
                "copyPaste": {
                    "source": {
                        "sheetId": int(sheet_id),
                        "startRowIndex": int(formula_source_row) - 1,
                        "endRowIndex": int(formula_source_row),
                        "startColumnIndex": 3,  # D열
                        "endColumnIndex": 5,    # E열까지
                    },
                    "destination": {
                        "sheetId": int(sheet_id),
                        "startRowIndex": int(target_row) - 1,
                        "endRowIndex": int(target_row),
                        "startColumnIndex": 3,
                        "endColumnIndex": 5,
                    },
                    "pasteType": "PASTE_FORMULA",
                    "pasteOrientation": "NORMAL",
                }
            },
            {
                "copyPaste": {
                    "source": {
                        "sheetId": int(sheet_id),
                        "startRowIndex": int(formula_source_row) - 1,
                        "endRowIndex": int(formula_source_row),
                        "startColumnIndex": 3,
                        "endColumnIndex": 5,
                    },
                    "destination": {
                        "sheetId": int(sheet_id),
                        "startRowIndex": int(target_row) - 1,
                        "endRowIndex": int(target_row),
                        "startColumnIndex": 3,
                        "endColumnIndex": 5,
                    },
                    "pasteType": "PASTE_FORMAT",
                    "pasteOrientation": "NORMAL",
                }
            },
        ]
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": copy_requests},
        ).execute()
        formula_copied = True

    write_range = f"{quoted}!A{target_row}:C{target_row}"
    row_values = [[write_date_label, int(round(float(total_amount or 0))), int(round(float(shipping_calc or 0)))]]
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"values": row_values},
    ).execute()

    if action == "updated":
        action_text = "기존 날짜 행 갱신"
    elif insert_needed:
        action_text = "날짜 순서에 맞춰 새 날짜 행 삽입"
    else:
        action_text = "새 날짜 행 생성"

    formula_text = " / D:E 수식 복사 완료" if formula_copied else ""
    range_text = " / 누락 기간 날짜범위 표기" if range_label_applied else ""

    return {
        "status": "written",
        "message": f"'{month_sheet}' 시트 {target_row}행에 {write_date_label} 기준 결과를 기록했습니다. ({action_text}{range_text}{formula_text})",
        "month_sheet": month_sheet,
        "date_label": write_date_label,
        "target_date_label": date_label,
        "row": target_row,
        "action": action,
        "inserted": insert_needed,
        "range_label_applied": range_label_applied,
        "formula_copied": formula_copied,
        "formula_source_row": formula_source_row,
    }

def _sales_list_drive_excels_for_date(target_date) -> tuple[str, list[dict]]:
    """판매내역/연도/날짜폴더 안의 xlsx 파일 목록을 전부 반환합니다."""
    service = _get_google_drive_service()
    root_id = _drive_get_or_create_sales_root_folder(service)
    year_folder_name, date_folder_name, date_folder_id = _drive_find_year_date_folder(service, root_id, target_date)
    folder_path = f"{year_folder_name}/{date_folder_name}"
    if not date_folder_id:
        return folder_path, []

    q = (
        f"trashed=false and '{date_folder_id}' in parents and "
        f"(mimeType='{GOOGLE_DRIVE_XLSX_MIME}' or name contains '.xlsx')"
    )
    files = _drive_list_files_all(service, q, page_size=100)
    # 임시/숨김 파일 제외
    files = [f for f in files if str(f.get("name", "")).lower().endswith(".xlsx") and not str(f.get("name", "")).startswith("~$")]
    return folder_path, files


def _sales_calc_from_drive_date_folder(target_date, spreadsheet_id_or_url: str = "") -> dict:
    """
    Google Drive 날짜 폴더의 엑셀을 전부 계산하고, 선택적으로 지정 Google Sheet에 기록합니다.
    """
    drive_service = _get_google_drive_service()
    root_id = _drive_get_or_create_sales_root_folder(drive_service)
    year_folder_name, date_folder_name, date_folder_id = _drive_find_year_date_folder(drive_service, root_id, target_date)
    folder_path = f"{year_folder_name}/{date_folder_name}"

    if not date_folder_id:
        return {
            "ok": False,
            "date_folder_name": folder_path,
            "files": [],
            "summary_df": pd.DataFrame(),
            "grand_amount": 0.0,
            "grand_unique_count_sum": 0,
            "grand_shipping_calc": 0,
            "sheet_write": {"status": "skipped", "message": "연도/날짜 폴더가 없어 계산/기록을 진행하지 않았습니다."},
            "message": f"Google Drive에서 '{GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME}/{folder_path}' 폴더를 찾지 못했습니다.",
        }

    q = (
        f"trashed=false and '{date_folder_id}' in parents and "
        f"(mimeType='{GOOGLE_DRIVE_XLSX_MIME}' or name contains '.xlsx')"
    )
    drive_files = _drive_list_files_all(drive_service, q, page_size=100)
    drive_files = [
        f for f in drive_files
        if str(f.get("name", "")).lower().endswith(".xlsx") and not str(f.get("name", "")).startswith("~$")
    ]

    if not drive_files:
        return {
            "ok": False,
            "date_folder_name": folder_path,
            "files": [],
            "summary_df": pd.DataFrame(),
            "grand_amount": 0.0,
            "grand_unique_count_sum": 0,
            "grand_shipping_calc": 0,
            "sheet_write": {"status": "skipped", "message": "날짜 폴더 안에 xlsx 파일이 없어 시트 기록은 건너뛰었습니다."},
            "message": f"'{GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME}/{folder_path}' 폴더 안에 .xlsx 파일이 없습니다.",
        }

    per_file_rows = []
    grand_amount = 0.0
    grand_unique_count_sum = 0

    for f in drive_files:
        try:
            file_bytes = _drive_download_file_bytes(drive_service, f["id"])
            amount_sum, keyset = _sales_compute_from_file_bytes(file_bytes)
            unique_count = len(keyset)
            shipping_calc = unique_count * 3500

            per_file_rows.append({
                "파일명": f.get("name", ""),
                "최종 상품별 총 주문금액 합계": amount_sum,
                "배송비≠0 (중복제거 인원수)": unique_count,
                "인원×3,500 합계": shipping_calc,
            })
            grand_amount += amount_sum
            grand_unique_count_sum += unique_count
        except Exception as e:
            per_file_rows.append({
                "파일명": f.get("name", ""),
                "최종 상품별 총 주문금액 합계": None,
                "배송비≠0 (중복제거 인원수)": None,
                "인원×3,500 합계": None,
                "오류": str(e),
            })

    grand_shipping_calc = grand_unique_count_sum * 3500
    summary_df = pd.DataFrame(per_file_rows)

    sheet_write = _sales_write_result_to_month_sheet(
        spreadsheet_id_or_url=spreadsheet_id_or_url,
        target_date=target_date,
        total_amount=grand_amount,
        shipping_calc=grand_shipping_calc,
    )

    return {
        "ok": True,
        "date_folder_name": folder_path,
        "files": drive_files,
        "summary_df": summary_df,
        "grand_amount": grand_amount,
        "grand_unique_count_sum": grand_unique_count_sum,
        "grand_shipping_calc": grand_shipping_calc,
        "sheet_write": sheet_write,
        "message": f"'{GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME}/{folder_path}' 폴더의 xlsx {len(drive_files)}개를 계산했습니다.",
    }


def _sales_fmt_commas(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass

    try:
        v = float(x)
    except Exception:
        return str(x)

    # integer-like
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}"

    # keep decimals (trim trailing zeros)
    s = f"{v:,.10f}"
    s = s.rstrip("0").rstrip(".")
    return s


def _sales_fmt_won(x) -> str:
    s = _sales_fmt_commas(x)
    return f"{s} 원" if s != "" else ""


def _sales_fmt_person(x) -> str:
    s = _sales_fmt_commas(x)
    return f"{s} 명" if s != "" else ""


def render_sales_calc_page():
    st.title("💰 매출계산")

    # 🔒 비밀번호 보호 (매출계산)    # 🔒 비밀번호 보호 (매출계산)
    # ⚠️ Streamlit 제약: 위젯이 생성된 뒤에는 동일 key의 session_state 값을 같은 run에서 직접 변경하면 오류가 납니다.
    # 그래서 비밀번호 입력칸(value)은 건드리지 않고, 인증 성공 시에는 rerun 후(위젯 미생성 상태)에서만 초기화합니다.
    if st.session_state.get("sales_authed", False):
        # 인증된 상태에서는 비밀번호 입력값을 지워둠(이 run에서는 입력 위젯이 없어서 안전)
        if "sales_password_input" in st.session_state:
            try:
                del st.session_state["sales_password_input"]
            except Exception:
                pass

        if st.button("🔓 로그아웃", use_container_width=False, key="sales_logout_btn"):
            st.session_state["sales_authed"] = False
            st.rerun()

    else:
        st.caption("이 페이지는 비밀번호가 필요합니다.")

        pw = None
        ok = False
        left_col, _ = st.columns([2, 8])
        with left_col:
            with st.form("sales_pw_form", clear_on_submit=False):
                pw = st.text_input(
                    "비밀번호",
                    type="password",
                    key="sales_password_input",
                    label_visibility="collapsed",
                    placeholder="비밀번호",
                )
                ok = st.form_submit_button("입장", use_container_width=False)

            if ok:
                if (pw or "").strip() == "1390":
                    st.session_state["sales_authed"] = True
                    st.rerun()
                else:
                    st.error("비밀번호가 올바르지 않습니다.")

        st.stop()

    mode = st.radio(
        "실행 방식",
        options=["수동", "자동"],
        horizontal=True,
        key="sales_calc_mode",
    )

    if mode == "수동":
        st.subheader("📊 네이버 매출 엑셀 합계 계산기")
        st.caption("엑셀을 직접 업로드해서 화면에서만 계산합니다. Google Sheet 기록은 자동 실행에서만 진행합니다.")

        uploaded_files = st.file_uploader(
            "엑셀 파일 업로드 (비밀번호 0000 고정) — 여러 개 업로드 가능",
            type=["xlsx"],
            accept_multiple_files=True,
            key="sales_uploaded_files",
        )

        left, _ = st.columns([1, 2])
        with left:
            calc_btn = st.button("✅ 계산", use_container_width=True, key="sales_calc_btn")

        if calc_btn:
            if not uploaded_files:
                st.warning("먼저 엑셀 파일을 업로드해 주세요.")
            else:
                per_file_rows = []
                grand_amount = 0.0

                # ✅ 전체 결과의 인원수 = "파일별(각 파일 내부 중복 제거) 인원수"를 합산
                grand_unique_count_sum = 0

                progress = st.progress(0)

                for i, f in enumerate(uploaded_files, start=1):
                    try:
                        sheets = _sales_read_excel_sheets(f.getvalue())
                        amount_sum, keyset = _sales_compute_from_sheets(sheets)

                        unique_count = len(keyset)  # 파일 내부(시트 포함) 중복 제거
                        shipping_calc = unique_count * 3500

                        per_file_rows.append(
                            {
                                "파일명": f.name,
                                "최종 상품별 총 주문금액 합계": amount_sum,
                                "배송비≠0 (중복제거 인원수)": unique_count,
                                "인원×3,500 합계": shipping_calc,
                            }
                        )

                        grand_amount += amount_sum
                        grand_unique_count_sum += unique_count  # ✅ 파일별 합산

                    except Exception as e:
                        per_file_rows.append(
                            {
                                "파일명": f.name,
                                "최종 상품별 총 주문금액 합계": None,
                                "배송비≠0 (중복제거 인원수)": None,
                                "인원×3,500 합계": None,
                                "오류": str(e),
                            }
                        )

                    progress.progress(i / len(uploaded_files))

                grand_shipping_calc = grand_unique_count_sum * 3500
                summary_df = pd.DataFrame(per_file_rows)

                st.session_state["sales_result"] = {
                    "summary_df": summary_df,
                    "grand_amount": grand_amount,
                    "grand_unique_count_sum": grand_unique_count_sum,
                    "grand_shipping_calc": grand_shipping_calc,
                }

    else:
        st.subheader("📁 자동 실행")
        st.caption(
            f"날짜를 따로 바꾸지 않으면 오늘 날짜 기준으로 "
            f"'{GOOGLE_DRIVE_SALES_ROOT_FOLDER_NAME}/연도/월.일' 폴더 안의 엑셀을 전부 가져와 계산합니다. "
            "결과는 지정 Google Sheet의 해당월 시트에 A=날짜, B=총 주문금액, C=인원×3,500원으로 기록합니다. "
            "해당월 시트가 없으면 새로 만들지 않고 기록하지 않습니다."
        )

        drive_col1, drive_col2 = st.columns([1, 2])
        with drive_col1:
            sales_drive_date = st.date_input(
                "불러올 날짜",
                value=datetime.now(KST_TZ).date(),
                key="sales_drive_date",
            )
        with drive_col2:
            sales_result_sheet_input = st.text_input(
                "결과 기록 Google Sheet ID 또는 URL",
                value=SALES_RESULT_SPREADSHEET_ID,
                placeholder="https://docs.google.com/spreadsheets/d/시트ID/edit 또는 시트ID",
                key="sales_result_spreadsheet_input",
            )

        drive_btn_col, _ = st.columns([1.6, 1])
        with drive_btn_col:
            drive_calc_btn = st.button(
                "▶ 자동 실행 시작",
                use_container_width=True,
                key="sales_drive_calc_btn",
            )

        if drive_calc_btn:
            if not (sales_result_sheet_input or "").strip():
                st.warning("결과를 넣을 Google Sheet ID 또는 URL을 입력해 주세요. Render 환경변수 SALES_RESULT_SPREADSHEET_ID로 고정해도 됩니다.")
            else:
                with st.spinner("Google Drive 날짜 폴더의 엑셀을 불러와 계산 중입니다..."):
                    try:
                        drive_res = _sales_calc_from_drive_date_folder(
                            target_date=sales_drive_date,
                            spreadsheet_id_or_url=sales_result_sheet_input,
                        )
                        st.session_state["sales_result"] = {
                            "summary_df": drive_res["summary_df"],
                            "grand_amount": drive_res["grand_amount"],
                            "grand_unique_count_sum": drive_res["grand_unique_count_sum"],
                            "grand_shipping_calc": drive_res["grand_shipping_calc"],
                        }
                        st.session_state["sales_drive_last_message"] = drive_res.get("message", "")
                        st.session_state["sales_drive_last_sheet_write"] = drive_res.get("sheet_write", {})

                        if drive_res.get("ok"):
                            st.success(drive_res.get("message", "계산 완료"))
                        else:
                            st.warning(drive_res.get("message", "계산할 파일이 없습니다."))

                        sheet_write = drive_res.get("sheet_write", {}) or {}
                        if sheet_write.get("status") == "written":
                            st.success(sheet_write.get("message", "Google Sheet 기록 완료"))
                        elif sheet_write.get("message"):
                            st.info(sheet_write.get("message"))
                    except Exception as e:
                        st.error(f"Drive 자동 계산/시트 기록 실패: {e}")

    if "sales_result" in st.session_state:
        res = st.session_state["sales_result"]
        summary_df = res["summary_df"]
        grand_amount = res["grand_amount"]
        grand_unique_count_sum = res["grand_unique_count_sum"]
        grand_shipping_calc = res["grand_shipping_calc"]

        st.subheader("✅ 전체 결과")

        amount_view = _sales_fmt_commas(grand_amount)
        shipping_view = _sales_fmt_commas(grand_shipping_calc)

        # ✅ 경고 방지 + 값 불일치 방지:
        # text_input에 value=를 주지 않고, session_state로만 값을 세팅
        st.session_state["sales_copy_total_amount_fmt_only"] = amount_view
        st.session_state["sales_copy_shipping_fmt_only"] = shipping_view

        # ✅ “📋 엑셀 복사용”을 맨 왼쪽으로 배치
        c_copy, c1, c2, c3 = st.columns([1.3, 1, 1, 1])

        with c_copy:
            st.caption("📋 엑셀 복사용 (클릭 → Ctrl+C)")
            st.text_input(
                "최종 상품별 총 주문금액 총합 (표시용 / 콤마)",
                key="sales_copy_total_amount_fmt_only",
            )
            st.text_input(
                "인원×3,500원 합계 (표시용 / 콤마)",
                key="sales_copy_shipping_fmt_only",
            )

        c1.metric("최종 상품별 총 주문금액 총합", f"{amount_view} 원")
        c2.metric("배송비≠0 인원수(파일별 합산)", f"{_sales_fmt_commas(grand_unique_count_sum)} 명")
        c3.metric("인원×3,500 합계", f"{shipping_view} 원")

        st.subheader("파일별 상세")

        # ✅ 파일별 상세에서 숫자를 통화로 표시
        display_df = summary_df.copy()

        if "최종 상품별 총 주문금액 합계" in display_df.columns:
            display_df["최종 상품별 총 주문금액 합계"] = display_df["최종 상품별 총 주문금액 합계"].apply(_sales_fmt_won)

        if "인원×3,500 합계" in display_df.columns:
            display_df["인원×3,500 합계"] = display_df["인원×3,500 합계"].apply(_sales_fmt_won)

        if "배송비≠0 (중복제거 인원수)" in display_df.columns:
            display_df["배송비≠0 (중복제거 인원수)"] = display_df["배송비≠0 (중복제거 인원수)"].apply(_sales_fmt_person)

        st.dataframe(display_df, use_container_width=True)



# =====================================================
# Page: 🚚 송장등록 (2번 코드 기능 이식)
