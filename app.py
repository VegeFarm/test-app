import io
import json
import re
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Optional, Tuple, List, Dict
from collections import OrderedDict

import pandas as pd
import streamlit as st
import openpyxl

# -----------------------------
# Excel decrypt
# -----------------------------
try:
    import msoffcrypto
except ModuleNotFoundError:
    msoffcrypto = None

# -----------------------------
# PDF
# -----------------------------
from reportlab.platypus import (
    SimpleDocTemplate,
    LongTable,
    TableStyle,
    Paragraph,
    Spacer,
    KeepTogether,
    HRFlowable,
)
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# =====================================================
# CONFIG
# =====================================================
KST = ZoneInfo("Asia/Seoul")
EXCEL_PASSWORD = "0000"

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

MAPPING_PATH = DATA_DIR / "name_mappings.json"
EXPR_RULES_PATH = DATA_DIR / "expression_rules.json"

BACKUP_DIR = DATA_DIR / "rules_backup"
BACKUP_DIR.mkdir(exist_ok=True)

# ✅ TC 설정 저장 파일 (프로그램 껐다 켜도 유지)
TC_SETTINGS_PATH = DATA_DIR / "tc_settings.json"

# ✅ 레포(앱 폴더)에 "TC주문_등록양식.xlsx" 파일을 같이 올려두면 업로드 없이 자동 사용
TC_TEMPLATE_DEFAULT_PATH = Path("TC주문_등록양식.xlsx")

# 수취인별 PDF 스타일
RECIPIENT_FONT_SIZE = 12
RECIPIENT_LEADING = 15
RECIPIENT_BLOCK_GAP_MM = 4.0
RECIPIENT_LINE_AFTER_MM = 4.0

# 스티커 용지 설정 (A4 / 65칸 / 38.2x21.1mm)
STICKER_COLS = 5
STICKER_ROWS = 13
STICKER_PER_PAGE = STICKER_COLS * STICKER_ROWS  # 65
STICKER_CELL_W_MM = 38.2
STICKER_CELL_H_MM = 21.1

# ✅ 스티커 텍스트 스타일 (Bold 없음 / 11pt)
STICKER_FONT_SIZE = 11
STICKER_LEADING = 13

# ✅ TC 양식 기본값
TC_PRODUCT_NAME_FIXED = "채소팜상품"
TC_ACCESS_FALLBACK = "경비실 호출"

# ✅ TC 배송유형 기본값(사이드바에서 수정 가능 + 저장됨)
TC_TYPE_DAWN_DEFAULT = "자동"
TC_TYPE_NEXT_DEFAULT = "택배대행"


# =====================================================
# VARIANT(단위) 추출
# =====================================================
UNIT_PATTERNS = [
    r"\d+(?:\.\d+)?kg\s*~\s*\d+(?:\.\d+)?kg",
    r"\d+(?:\.\d+)?kg",
    r"(?:약\s*)?\d+(?:\.\d+)?g",
    r"\d+개",
    r"\d+통",
    r"\d+단",
    r"\d+봉",
    r"\d+팩",
]
UNIT_RE = re.compile(r"(" + "|".join(UNIT_PATTERNS) + r")")


def extract_variant(name: str) -> str:
    s = (name or "").strip()
    m = UNIT_RE.search(s)
    if not m:
        return ""
    u = m.group(0)
    u = re.sub(r"\s+", "", u)
    u = u.replace("약", "")
    if "~" in u:
        u = u.split("~", 1)[1]
    return u


# =====================================================
# Helpers
# =====================================================
def normalize_text(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _safe_int(v) -> Optional[int]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        return int(v)
    except Exception:
        return None


def _xml_escape(s: str) -> str:
    s = str(s or "")
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _text_width_pt(text: str, font: str, size: float) -> float:
    try:
        w = pdfmetrics.stringWidth(text, font, size)
        if not w or w <= 0:
            return len(text) * size * 0.55
        return w
    except Exception:
        return len(text) * size * 0.55


def fmt_qty(x):
    try:
        x = float(x)
        return int(x) if x.is_integer() else x
    except Exception:
        return x


def _as_int_qty(v) -> int:
    try:
        f = float(v)
        if abs(f - round(f)) < 1e-9:
            return int(round(f))
        return int(round(f))
    except Exception:
        return 0


def _clean_access_message(msg: str) -> str:
    s = str(msg or "").strip()
    return s if s else TC_ACCESS_FALLBACK


# =====================================================
# ✅ TC Settings (persist)
# =====================================================
def load_tc_settings() -> Dict[str, str]:
    default = {"dawn": TC_TYPE_DAWN_DEFAULT, "next": TC_TYPE_NEXT_DEFAULT}
    if not TC_SETTINGS_PATH.exists():
        return default
    try:
        data = json.loads(TC_SETTINGS_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return default
        dawn = normalize_text(data.get("dawn", "")) or TC_TYPE_DAWN_DEFAULT
        nxt = normalize_text(data.get("next", "")) or TC_TYPE_NEXT_DEFAULT
        return {"dawn": dawn, "next": nxt}
    except Exception:
        return default


def save_tc_settings(dawn: str, nxt: str) -> None:
    dawn = normalize_text(dawn) or TC_TYPE_DAWN_DEFAULT
    nxt = normalize_text(nxt) or TC_TYPE_NEXT_DEFAULT
    TC_SETTINGS_PATH.write_text(json.dumps({"dawn": dawn, "next": nxt}, ensure_ascii=False, indent=2), encoding="utf-8")


# =====================================================
# 표현규칙 (통/개/팩/봉 같은 단위 관리)
# =====================================================
def default_expression_rules() -> Dict:
    return {
        "default_unit": "개",
        "units": [
            {"enabled": True, "unit": "개"},
            {"enabled": True, "unit": "봉"},
            {"enabled": True, "unit": "통"},
            {"enabled": True, "unit": "팩"},
        ],
        "note": "합산규칙(N)이 적용될 단위를 관리합니다.",
    }


def save_expression_rules(data: Dict) -> None:
    EXPR_RULES_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_expression_rules() -> Dict:
    if not EXPR_RULES_PATH.exists():
        data = default_expression_rules()
        save_expression_rules(data)
        return data
    try:
        data = json.loads(EXPR_RULES_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("invalid")

        if "units" not in data or not isinstance(data["units"], list):
            data["units"] = default_expression_rules()["units"]

        if "default_unit" not in data or not isinstance(data["default_unit"], str):
            data["default_unit"] = default_expression_rules()["default_unit"]

        cleaned_units = []
        for r in data["units"]:
            u = normalize_text(r.get("unit", ""))
            if not u:
                continue
            cleaned_units.append({"enabled": bool(r.get("enabled", True)), "unit": u})
        data["units"] = cleaned_units
        data["default_unit"] = normalize_text(data.get("default_unit", "개")) or "개"
        return data
    except Exception:
        data = default_expression_rules()
        save_expression_rules(data)
        return data


def get_bundle_units(expr: Dict) -> List[str]:
    units = []
    for r in expr.get("units", []):
        if r.get("enabled", True):
            u = normalize_text(r.get("unit", ""))
            if u:
                units.append(u)

    seen = set()
    out = []
    for u in units:
        if u not in seen:
            out.append(u)
            seen.add(u)
    return out


def build_bundle_re(bundle_units: List[str]) -> re.Pattern:
    if not bundle_units:
        bundle_units = ["개"]
    unit_alt = "|".join(map(re.escape, bundle_units))
    return re.compile(rf"^\s*(\d+)\s*({unit_alt})\s*$")


# =====================================================
# 상품명 매칭 규칙 (합산규칙 N 포함)
# =====================================================
def default_mapping_rules() -> List[Dict]:
    return [
        {
            "enabled": True,
            "match_type": "contains",
            "pattern": "와일드루꼴라",
            "display_name": "와일드",
            "sum_rule": None,
            "note": '예) "채소팜 와일드루꼴라 1kg ..." -> 와일드',
        },
        {
            "enabled": True,
            "match_type": "contains",
            "pattern": "라디치오",
            "display_name": "라디치오",
            "sum_rule": None,
            "note": '예) "채소팜 라디치오 1통 ..." -> 라디치오',
        },
        {
            "enabled": False,
            "match_type": "contains",
            "pattern": "오렌지",
            "display_name": "오렌지",
            "sum_rule": 5,
            "note": "예) 오렌지 합산규칙=5",
        },
    ]


def save_mapping_rules(rules: List[Dict]) -> None:
    MAPPING_PATH.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")


def load_mapping_rules() -> List[Dict]:
    if not MAPPING_PATH.exists():
        rules = default_mapping_rules()
        save_mapping_rules(rules)
        return rules

    try:
        raw = json.loads(MAPPING_PATH.read_text(encoding="utf-8"))
        if isinstance(raw, list):
            cleaned = []
            for r in raw:
                sr = _safe_int(r.get("sum_rule"))
                if sr is not None and sr < 2:
                    sr = None
                cleaned.append(
                    dict(
                        enabled=bool(r.get("enabled", True)),
                        match_type=normalize_text(r.get("match_type", "contains")) or "contains",
                        pattern=normalize_text(r.get("pattern", "")),
                        display_name=normalize_text(r.get("display_name", "")),
                        sum_rule=sr,
                        note=normalize_text(r.get("note", "")),
                    )
                )
            return cleaned
    except Exception:
        pass

    rules = default_mapping_rules()
    save_mapping_rules(rules)
    return rules


def apply_mapping(actual_name: str, rules: List[Dict]) -> Tuple[str, bool, Optional[int]]:
    actual = normalize_text(actual_name)
    if not actual:
        return "", False, None

    for r in rules:
        if not r.get("enabled", True):
            continue

        mt = normalize_text(r.get("match_type", "contains")) or "contains"
        pattern = normalize_text(r.get("pattern", ""))
        display = normalize_text(r.get("display_name", ""))
        sr = _safe_int(r.get("sum_rule"))

        if not pattern or not display:
            continue

        matched = False
        if mt == "exact":
            matched = (actual == pattern)
        elif mt == "contains":
            matched = (pattern in actual)
        elif mt == "regex":
            try:
                matched = bool(re.search(pattern, actual))
            except re.error:
                matched = False

        if matched:
            if sr is not None and sr < 2:
                sr = None
            return display, True, sr

    # fallback
    s = re.sub(r"^\s*채소팜\s*", "", actual)
    s = re.sub(r"\([^)]*\)", "", s).strip()
    s = re.sub(r"\s+", " ", s).strip()
    m = UNIT_RE.search(s)
    if m:
        s = s[: m.start()].strip()

    toks = s.split()
    if not toks:
        return actual, False, None

    PREFIX = {"생", "유기농", "국산", "수입", "냉동", "베이비", "프리미엄"}
    if len(toks) >= 2 and toks[0] in PREFIX:
        fallback = toks[0] + toks[1]
    else:
        fallback = toks[0]

    return fallback, False, None


def mapping_df_from_list(rules: List[Dict]) -> pd.DataFrame:
    df = pd.DataFrame(rules)
    keep = ["enabled", "match_type", "pattern", "display_name", "sum_rule", "note"]
    for c in keep:
        if c not in df.columns:
            df[c] = None
    return df[keep]


def mapping_list_from_df(edited: pd.DataFrame) -> List[Dict]:
    cleaned = []
    for _, row in edited.iterrows():
        pattern = normalize_text(row.get("pattern"))
        display = normalize_text(row.get("display_name"))
        if not pattern or not display:
            continue

        mt = normalize_text(row.get("match_type")) or "contains"
        if mt not in {"contains", "exact", "regex"}:
            mt = "contains"

        sr = _safe_int(row.get("sum_rule"))
        if sr is not None and sr < 2:
            sr = None

        cleaned.append(
            dict(
                enabled=bool(row.get("enabled", True)),
                match_type=mt,
                pattern=pattern,
                display_name=display,
                sum_rule=sr,
                note=normalize_text(row.get("note")),
            )
        )
    return cleaned


# =====================================================
# Backups (Excel)
# =====================================================
def backup_rules_to_excel(mapping_rules: List[Dict], expr_rules: Dict) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = BACKUP_DIR / f"rules_backup_{ts}.xlsx"

    df_map = mapping_df_from_list(mapping_rules).rename(
        columns={
            "enabled": "사용",
            "match_type": "매칭방식",
            "pattern": "실제상품명(패턴)",
            "display_name": "표시될상품명",
            "sum_rule": "합산규칙(N)",
            "note": "메모",
        }
    )

    units = expr_rules.get("units", [])
    df_expr = pd.DataFrame(units)
    if df_expr.empty:
        df_expr = pd.DataFrame([{"enabled": True, "unit": expr_rules.get("default_unit", "개")}])
    if "enabled" not in df_expr.columns:
        df_expr["enabled"] = True
    if "unit" not in df_expr.columns:
        df_expr["unit"] = ""
    df_expr = df_expr[["enabled", "unit"]].rename(columns={"enabled": "사용", "unit": "단위"})

    df_meta = pd.DataFrame(
        [
            {"키": "default_unit", "값": expr_rules.get("default_unit", "개")},
            {"키": "note", "값": expr_rules.get("note", "")},
        ]
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_map.to_excel(writer, sheet_name="상품명매칭", index=False)
        df_expr.to_excel(writer, sheet_name="표현규칙_단위", index=False)
        df_meta.to_excel(writer, sheet_name="표현규칙_설정", index=False)

    return out_path


# =====================================================
# Excel decrypt / read
# =====================================================
def decrypt_excel(uploaded_bytes: bytes, password: str = EXCEL_PASSWORD) -> io.BytesIO:
    if msoffcrypto is None:
        raise ModuleNotFoundError("msoffcrypto not installed")
    decrypted = io.BytesIO()
    office = msoffcrypto.OfficeFile(io.BytesIO(uploaded_bytes))
    office.load_key(password=password)
    office.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def _norm_col(x) -> str:
    s = str(x if x is not None else "")
    s = s.replace("\xa0", " ").replace("\n", " ").replace("\r", " ")
    return normalize_text(s)


def find_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    """컬럼명 탐색: 공백/개행/NBSP 등을 정규화해서 매칭합니다."""
    cols = list(df.columns)
    if not cols:
        return None

    kw_norm = [_norm_col(k) for k in (keywords or []) if _norm_col(k)]
    if not kw_norm:
        return None

    col_norm = [_norm_col(c) for c in cols]

    # 1) 정규화 후 완전일치
    for k in kw_norm:
        for c, cn in zip(cols, col_norm):
            if k == cn:
                return c

    # 2) 정규화 후 부분일치
    for c, cn in zip(cols, col_norm):
        for k in kw_norm:
            if k in cn:
                return c

    return None


# =====================================================
# ✅ Smart Excel header detection (안내문이 위에 있는 엑셀 자동 처리)
# =====================================================
REQUIRED_COL_GROUPS = OrderedDict(
    [
        ("상품명", ["상품명", "상품", "제품명"]),
        ("수량", ["수량", "주문수량", "구매수량", "개수"]),
        ("구매자명", ["구매자명", "구매자"]),
        ("수취인명", ["수취인명", "수령인", "받는사람"]),
        ("통합배송지", ["통합배송지", "배송지", "주소"]),
        ("옵션정보", ["옵션정보", "옵션", "선택옵션"]),
        ("수취인연락처", ["수취인연락처", "수령인연락처", "수취인 연락처", "수령인 연락처", "전화번호", "연락처"]),
        ("배송메세지", ["배송메세지", "배송메시지", "배송 메시지", "배송 메세지", "배송요청사항", "요청사항"]),
    ]
)


def _missing_required_cols(df: pd.DataFrame) -> List[str]:
    missing = []
    for k, kws in REQUIRED_COL_GROUPS.items():
        if find_col(df, kws) is None:
            missing.append(k)
    return missing


def _guess_header_row(preview: pd.DataFrame, scan_limit: int = 40) -> Tuple[Optional[int], int]:
    """header=None 로 읽은 preview에서 '헤더로 보이는 행'을 추정"""
    if preview is None or preview.empty:
        return None, 0

    best_i = None
    best_score = -1

    n = min(scan_limit, len(preview))
    for i in range(n):
        row = preview.iloc[i].tolist()
        row_strs = []
        for v in row:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            s = _norm_col(v)
            if s:
                row_strs.append(s)

        if not row_strs:
            continue

        score = 0
        for _, kws in REQUIRED_COL_GROUPS.items():
            hit = False
            for kw in kws:
                kw_n = _norm_col(kw)
                if not kw_n:
                    continue
                if any(kw_n in cell for cell in row_strs):
                    hit = True
                    break
            if hit:
                score += 1

        if score > best_score:
            best_score = score
            best_i = i

    return best_i, max(best_score, 0)


def smart_read_orders_excel(excel_bytes: bytes, min_score: int = 4) -> Tuple[pd.DataFrame, Dict]:
    """
    스마트스토어/일괄발송 엑셀처럼 상단에 안내문이 있는 경우,
    '헤더 행'을 자동으로 찾아서 DataFrame을 반환합니다.
    """
    if not excel_bytes:
        raise ValueError("empty excel bytes")

    def _read(sheet, header, nrows=None):
        bio = io.BytesIO(excel_bytes)
        bio.seek(0)
        return pd.read_excel(bio, sheet_name=sheet, header=header, nrows=nrows, engine="openpyxl")

    # 시트 목록
    bio0 = io.BytesIO(excel_bytes)
    bio0.seek(0)
    xls = pd.ExcelFile(bio0, engine="openpyxl")

    best_fallback = None  # (df, meta)

    for sheet in xls.sheet_names:
        # 1) 일반 header=0 시도
        try:
            df0 = _read(sheet, header=0)
            df0 = df0.dropna(how="all")
            missing0 = _missing_required_cols(df0)
            if not missing0:
                return df0, {"sheet": sheet, "header_row": 0, "method": "header=0"}
            if best_fallback is None:
                best_fallback = (df0, {"sheet": sheet, "header_row": 0, "method": "header=0", "missing": missing0})
        except Exception:
            pass

        # 2) header 행 추정
        try:
            preview = _read(sheet, header=None, nrows=60)
            header_row, score = _guess_header_row(preview, scan_limit=40)

            if header_row is None or score < min_score:
                continue

            df = _read(sheet, header=int(header_row))
            df = df.dropna(how="all")

            # 빈 컬럼 제거(Unnamed: n)
            try:
                df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
            except Exception:
                pass

            missing = _missing_required_cols(df)
            if not missing:
                return df, {"sheet": sheet, "header_row": int(header_row), "method": f"guessed(score={score})"}

            if best_fallback is None:
                best_fallback = (
                    df,
                    {"sheet": sheet, "header_row": int(header_row), "method": f"guessed(score={score})", "missing": missing},
                )
        except Exception:
            pass

    if best_fallback is not None:
        return best_fallback

    # 최후의 수단
    df_last = _read(0, header=0)
    return df_last, {"sheet": 0, "header_row": 0, "method": "fallback"}


# =====================================================
# 합산규칙 적용 (표현규칙에서 켠 단위에만)
# =====================================================
def parse_bundle_variant(variant: str, bundle_re: re.Pattern) -> Tuple[Optional[int], Optional[str]]:
    m = bundle_re.match((variant or "").strip())
    if not m:
        return None, None
    try:
        return int(m.group(1)), m.group(2)
    except Exception:
        return None, None


def explode_sum_rule_rows(
    df_rows: pd.DataFrame,
    bundle_units: List[str],
    default_unit: str,
) -> pd.DataFrame:
    bundle_units = bundle_units or [default_unit or "개"]
    default_unit = default_unit or "개"
    bundle_re = build_bundle_re(bundle_units)
    unit_set = set(bundle_units)

    out = []
    for _, r in df_rows.iterrows():
        product = r["제품명"]
        variant = (r.get("구분", "") or "").strip()
        qty = r.get("수량", None)
        rule_n = _safe_int(r.get("합산규칙", None))

        if rule_n is None or rule_n < 2:
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        # 구분이 비어 있으면 기본 단위 1개로 간주
        if variant == "":
            unit_size, unit_label = 1, default_unit
            is_bundle = unit_label in unit_set
        else:
            unit_size, unit_label = parse_bundle_variant(variant, bundle_re)
            is_bundle = (unit_size is not None and unit_label in unit_set)

        if not is_bundle:
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        try:
            total_units = int(round(float(qty))) * int(unit_size)
        except Exception:
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        if total_units <= 0:
            continue

        full = total_units // rule_n
        rem = total_units % rule_n

        if full > 0:
            out.append({"제품명": product, "구분": f"{rule_n}{unit_label}", "수량": full})
        if rem > 0:
            out.append({"제품명": product, "구분": f"{rem}{unit_label}", "수량": 1})

    return pd.DataFrame(out)


# =====================================================
# 배송 옵션 분류 & 그룹 규칙 (새벽 우선)
# =====================================================
def classify_delivery(opt: str) -> str:
    s = str(opt or "")
    if "새벽배송" in s:
        return "새벽배송"
    if "익일배송" in s:
        return "익일배송"
    return "기타"


def decide_group_delivery(deliv_set: set) -> str:
    if "새벽배송" in deliv_set:
        return "새벽배송"
    if "익일배송" in deliv_set:
        return "익일배송"
    return "기타"


# =====================================================
# PDF 1) 제품별 개수
# =====================================================
def build_summary_pdf(summary_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()

    font_name = "Helvetica"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
        font_name = "HYGothic-Medium"
    except Exception:
        pass

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=14,
        leading=18,
        spaceAfter=8,
    )

    elems = []
    elems.append(Paragraph("▣ 제품별 개수", title_style))
    elems.append(Spacer(1, 4))

    data = [["제품명", "구분", "수량"]]
    for _, row in summary_df.iterrows():
        data.append([str(row["제품명"]), str(row["구분"]), str(row["수량"])])

    table = LongTable(
        data,
        colWidths=[75 * mm, 60 * mm, 25 * mm],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (2, 1), (2, -1), "RIGHT"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("TOPPADDING", (0, 0), (-1, 0), 6),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("TOPPADDING", (0, 1), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ]
        )
    )

    elems.append(table)
    doc.build(elems)
    return buf.getvalue()


# =====================================================
# PDF 2) 수취인별 출력
# =====================================================
def build_recipient_pdf(entries: List[Dict[str, str]]) -> bytes:
    buf = io.BytesIO()

    font_name = "Helvetica"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
        font_name = "HYGothic-Medium"
    except Exception:
        pass

    left_margin = 12 * mm
    right_margin = 12 * mm
    top_margin = 12 * mm
    bottom_margin = 12 * mm

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )

    styles = getSampleStyleSheet()
    base_style = ParagraphStyle(
        "base",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=RECIPIENT_FONT_SIZE,
        leading=RECIPIENT_LEADING,
        spaceAfter=0,
    )

    usable_width = A4[0] - left_margin - right_margin

    elems = []
    for e in entries:
        recv = (e.get("수취인명") or "").strip() or " "
        items = (e.get("items_line") or "").strip() or " "

        name_token_plain = f"{recv} - "
        indent = _text_width_pt(name_token_plain, font_name, base_style.fontSize)
        indent_cap = usable_width * 0.55
        indent = min(max(indent, 40), indent_cap)

        line_style = ParagraphStyle(
            f"line_{abs(hash(recv)) % 10_000_000}",
            parent=base_style,
            leftIndent=indent,
            firstLineIndent=-indent,
        )

        text = f"<b>{_xml_escape(recv)}</b> - {_xml_escape(items)}"
        p = Paragraph(text, line_style)

        block = KeepTogether(
            [
                p,
                Spacer(1, RECIPIENT_BLOCK_GAP_MM * mm),
                HRFlowable(
                    width="100%",
                    thickness=0.4,
                    color=colors.lightgrey,
                    spaceBefore=0,
                    spaceAfter=RECIPIENT_LINE_AFTER_MM * mm,
                ),
            ]
        )
        elems.append(block)

    doc.build(elems)
    return buf.getvalue()


# =====================================================
# PDF 3) 스티커 용지 (Bold 없음 / 11pt)
# =====================================================
def _wrap_for_cell(txt: str, font_name: str, font_size: int, max_w_pt: float) -> List[str]:
    txt = (txt or "").strip()
    if not txt:
        return [""]

    def w(s: str) -> float:
        return _text_width_pt(s, font_name, font_size)

    if w(txt) <= max_w_pt:
        return [txt]

    if " " in txt:
        parts = txt.split()
        line1 = ""
        consumed = 0
        for p in parts:
            cand = (line1 + " " + p).strip()
            if w(cand) <= max_w_pt:
                line1 = cand
                consumed += 1
            else:
                break
        rest = " ".join(parts[consumed:]).strip()
        if not rest:
            return [line1]
        if w(rest) <= max_w_pt:
            return [line1, rest]
        trimmed = rest
        while trimmed and w(trimmed + "...") > max_w_pt:
            trimmed = trimmed[:-1]
        return [line1, (trimmed + "...") if trimmed else "..."]

    line1 = ""
    for ch in txt:
        if w(line1 + ch) <= max_w_pt:
            line1 += ch
        else:
            break
    rest = txt[len(line1) :].strip()
    if not rest:
        return [line1]
    if w(rest) <= max_w_pt:
        return [line1, rest]
    trimmed = rest
    while trimmed and w(trimmed + "...") > max_w_pt:
        trimmed = trimmed[:-1]
    return [line1, (trimmed + "...") if trimmed else "..."]


def _draw_center_text(c: canvas.Canvas, font_name: str, font_size: int, x_center: float, y: float, txt: str):
    txt = (txt or "").strip()
    if not txt:
        return
    w = _text_width_pt(txt, font_name, font_size)
    x_left = x_center - (w / 2.0)
    t = c.beginText()
    t.setTextOrigin(x_left, y)
    t.setFont(font_name, font_size)
    t.textOut(txt)
    c.drawText(t)


def build_sticker_pdf(label_texts: List[str]) -> bytes:
    buf = io.BytesIO()

    font_name = "Helvetica"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
        font_name = "HYGothic-Medium"
    except Exception:
        pass

    c = canvas.Canvas(buf, pagesize=A4)
    page_w_pt, page_h_pt = A4

    cell_w_pt = STICKER_CELL_W_MM * mm
    cell_h_pt = STICKER_CELL_H_MM * mm
    grid_w_pt = cell_w_pt * STICKER_COLS
    grid_h_pt = cell_h_pt * STICKER_ROWS

    x0 = (page_w_pt - grid_w_pt) / 2.0
    y0 = (page_h_pt - grid_h_pt) / 2.0

    total = len(label_texts)
    page_count = (total + STICKER_PER_PAGE - 1) // STICKER_PER_PAGE if total else 1

    pad_x = 2.0 * mm
    max_text_w = cell_w_pt - (pad_x * 2)

    for p in range(page_count):
        c.setFillColor(colors.black)
        c.setFont(font_name, STICKER_FONT_SIZE)

        for r in range(STICKER_ROWS):
            for col in range(STICKER_COLS):
                slot = r * STICKER_COLS + col
                global_i = p * STICKER_PER_PAGE + slot
                if global_i >= total:
                    continue

                text = (label_texts[global_i] or "").strip()

                x = x0 + col * cell_w_pt
                y = y0 + (STICKER_ROWS - 1 - r) * cell_h_pt

                lines = _wrap_for_cell(text, font_name, STICKER_FONT_SIZE, max_text_w)[:2]

                cx = x + cell_w_pt / 2.0
                if len(lines) == 1:
                    cy = y + (cell_h_pt / 2.0) - (STICKER_FONT_SIZE * 0.35)
                    _draw_center_text(c, font_name, STICKER_FONT_SIZE, cx, cy, lines[0])
                else:
                    center = y + (cell_h_pt / 2.0)
                    upper_y = center + (STICKER_LEADING * 0.25)
                    lower_y = center - (STICKER_LEADING * 0.95)
                    _draw_center_text(c, font_name, STICKER_FONT_SIZE, cx, upper_y, lines[0])
                    _draw_center_text(c, font_name, STICKER_FONT_SIZE, cx, lower_y, lines[1])

        if p < page_count - 1:
            c.showPage()

    c.save()
    return buf.getvalue()


# =====================================================
# TC 주문_등록양식 자동 채우기
# =====================================================
def _norm_header(s: str) -> str:
    s = str(s or "")
    s = s.replace("*", "")
    s = re.sub(r"\s+", "", s)
    return s.strip().lower()


def build_tc_excel_bytes(template_bytes: bytes, rows: List[Dict[str, str]]) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    if "양식" not in wb.sheetnames:
        raise ValueError("TC 템플릿에 '양식' 시트가 없습니다.")
    ws = wb["양식"]

    headers = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(1, col).value
        if v is None:
            continue
        headers[_norm_header(v)] = col

    def col_of(label_candidates: List[str]) -> int:
        for cand in label_candidates:
            key = _norm_header(cand)
            if key in headers:
                return headers[key]
        raise KeyError(f"필수 헤더를 찾지 못했습니다: {label_candidates}")

    c_req = col_of(["배송요청일", "배송요청일*"])
    c_orderer = col_of(["주문자", "주문자*"])
    c_receiver = col_of(["수령자", "수령자*"])
    c_addr = col_of(["수령자도로명주소", "수령자 도로명 주소", "수령자 도로명 주소*"])
    c_phone = col_of(["수령자연락처", "수령자 연락처", "수령자 연락처*"])
    c_in = col_of(["출입방법", "출입 방법"])
    c_prod = col_of(["상품명", "상품명*"])
    c_type = col_of(["배송유형", "배송 유형", "배송 유형*"])

    start_row = 2
    for i, r in enumerate(rows):
        rr = start_row + i
        ws.cell(rr, c_req).value = r.get("배송요청일", "")
        ws.cell(rr, c_orderer).value = r.get("주문자", "")
        ws.cell(rr, c_receiver).value = r.get("수령자", "")
        ws.cell(rr, c_addr).value = r.get("수령자도로명주소", "")
        ws.cell(rr, c_phone).value = r.get("수령자연락처", "")
        ws.cell(rr, c_in).value = r.get("출입방법", "")
        ws.cell(rr, c_prod).value = r.get("상품명", "")
        ws.cell(rr, c_type).value = r.get("배송유형", "")
        # 배송받을장소는 건드리지 않음

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =====================================================
# Sidebar (상품명 매칭 규칙 메뉴에서만): 백업폴더 + 표현규칙
# =====================================================
def sidebar_backup_folder():
    with st.sidebar.expander("📁 규칙 백업폴더", expanded=False):
        try:
            backups = sorted(BACKUP_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        except Exception:
            backups = []

        if not backups:
            st.caption("아직 백업 파일이 없습니다.")
            return

        for i, fp in enumerate(backups[:60]):
            cols = st.columns([6, 2, 2])
            cols[0].write(fp.name)

            try:
                b = fp.read_bytes()
                cols[1].download_button(
                    "다운",
                    data=b,
                    file_name=fp.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_bk_{i}_{fp.name}",
                    use_container_width=True,
                )
            except Exception:
                cols[1].write("")

            if cols[2].button("삭제", key=f"rm_bk_{i}_{fp.name}", use_container_width=True):
                try:
                    fp.unlink()
                    st.success(f"삭제 완료: {fp.name}")
                    st.rerun()
                except Exception as e:
                    st.error("삭제 실패")
                    st.exception(e)


def sidebar_expression_rules():
    expr = load_expression_rules()
    units = expr.get("units", [])
    default_unit = normalize_text(expr.get("default_unit", "개")) or "개"

    with st.sidebar.expander("🧩 표현규칙", expanded=False):
        st.caption("합산규칙(N)을 적용할 단위를 관리합니다. (통/개/팩/봉 등)")

        df = pd.DataFrame(units)
        if df.empty:
            df = pd.DataFrame([{"enabled": True, "unit": default_unit}])
        if "enabled" not in df.columns:
            df["enabled"] = True
        if "unit" not in df.columns:
            df["unit"] = ""
        df = df[["enabled", "unit"]]

        edited = st.data_editor(
            df,
            hide_index=True,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "enabled": st.column_config.CheckboxColumn("사용", default=True),
                "unit": st.column_config.TextColumn("단위"),
            },
            key="expr_units_editor",
        )

        enabled_units = []
        for _, r in edited.iterrows():
            u = normalize_text(r.get("unit", ""))
            if u:
                enabled_units.append((bool(r.get("enabled", True)), u))

        enabled_only = [u for en, u in enabled_units if en]
        if not enabled_only:
            enabled_only = ["개"]

        if default_unit not in enabled_only:
            default_unit = enabled_only[0]

        new_default = st.selectbox(
            "기본단위 (구분이 비어있을 때)",
            options=enabled_only,
            index=enabled_only.index(default_unit) if default_unit in enabled_only else 0,
            key="expr_default_unit",
        )

        if st.button("💾 표현규칙 저장", use_container_width=True, key="save_expr_rules_btn"):
            cleaned_units = []
            seen = set()
            for en, u in enabled_units:
                if u in seen:
                    continue
                cleaned_units.append({"enabled": bool(en), "unit": u})
                seen.add(u)

            data = {
                "default_unit": new_default,
                "units": cleaned_units,
                "note": expr.get("note", ""),
            }
            save_expression_rules(data)
            st.success("표현규칙 저장 완료")
            st.rerun()


# =====================================================
# Streamlit UI
# =====================================================
st.set_page_config(page_title="제품별 개수 & 수취인별 출력", page_icon="📄", layout="wide")
st.title("📄 제품별 개수 & 수취인별 출력")
st.caption("엑셀 업로드 → 제품별 집계 + 수취인별 PDF + 스티커용지 PDF + TC주문_등록양식 자동작성")

menu = st.sidebar.radio("메뉴", ["🧩 상품명 매칭 규칙", "⬆️ 엑셀 업로드 & 결과"], index=1)
st.sidebar.markdown("---")

# -----------------------------
# 1) 상품명 매칭 규칙
# -----------------------------
if menu == "🧩 상품명 매칭 규칙":
    sidebar_backup_folder()
    sidebar_expression_rules()

    mapping_rules = load_mapping_rules()
    expr = load_expression_rules()

    st.subheader("실제 상품명 → 표시될 상품명 (합산규칙 포함)")
    st.markdown(
        """
**매칭방식 설명**
- **contains**: `패턴`이 `엑셀 상품명` 안에 포함되면 매칭
- **exact**: `패턴`과 `엑셀 상품명`이 완전히 동일할 때만 매칭
- **regex**: `패턴`을 정규식으로 해석해 매칭

**합산규칙(N)**  
- N=5, 단위가 표현규칙에 포함된 경우(개/봉/통/팩 등) → 8개 주문 시 `5개 1개` + `3개 1개`로 표현
"""
    )

    df = mapping_df_from_list(mapping_rules)
    edited = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "enabled": st.column_config.CheckboxColumn("사용", default=True),
            "match_type": st.column_config.SelectboxColumn("매칭 방식", options=["contains", "exact", "regex"]),
            "pattern": st.column_config.TextColumn("실제 상품명(패턴)", width="large"),
            "display_name": st.column_config.TextColumn("표시될 상품명", width="medium"),
            "sum_rule": st.column_config.NumberColumn("합산규칙(N)", min_value=2, step=1),
            "note": st.column_config.TextColumn("메모", width="large"),
        },
        key="mapping_editor_main",
    )

    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("💾 저장", use_container_width=True):
            cleaned = mapping_list_from_df(edited)
            save_mapping_rules(cleaned)
            st.success(f"저장 완료! (규칙 {len(cleaned)}개)")
            st.rerun()

    with c2:
        if st.button("📗 엑셀로 저장하기(백업)", use_container_width=True):
            cleaned_map = mapping_list_from_df(edited)
            outp = backup_rules_to_excel(cleaned_map, expr)
            st.success(f"백업 저장 완료: {outp.name}")
            st.rerun()

# -----------------------------
# 2) 엑셀 업로드 & 결과
# -----------------------------
else:
    st.subheader("엑셀 업로드 → 제품별 집계 + 수취인별 출력(새벽/익일) + TC주문_등록 자동작성")

    if msoffcrypto is None:
        st.error("msoffcrypto가 설치되지 않았습니다. requirements.txt에 'msoffcrypto-tool'을 추가하고 재배포해 주세요.")
        st.stop()

    # ✅ 이 페이지의 사이드바에서만 TC 배송유형 설정 + 저장
    tc_saved = load_tc_settings()
    if "tc_type_dawn" not in st.session_state:
        st.session_state.tc_type_dawn = tc_saved["dawn"]
    if "tc_type_next" not in st.session_state:
        st.session_state.tc_type_next = tc_saved["next"]

    with st.sidebar.expander("🧾 TC주문_등록 설정", expanded=False):
        st.caption("변경 후 [저장]을 누르면 다음 실행에도 그대로 유지됩니다.")

        dawn_val = st.text_input(
            "새벽배송 → 배송유형",
            value=st.session_state.tc_type_dawn,
            key="tc_type_dawn_input",
        )

        next_val = st.text_input(
            "익일배송 → 배송유형",
            value=st.session_state.tc_type_next,
            key="tc_type_next_input",
        )

        if st.button("💾 TC 설정 저장", use_container_width=True, key="save_tc_settings_btn"):
            dawn_val = (dawn_val or "").strip() or TC_TYPE_DAWN_DEFAULT
            next_val = (next_val or "").strip() or TC_TYPE_NEXT_DEFAULT

            st.session_state.tc_type_dawn = dawn_val
            st.session_state.tc_type_next = next_val

            save_tc_settings(dawn_val, next_val)
            st.success("TC 설정 저장 완료")
            st.rerun()

    uploaded = st.file_uploader("비밀번호(0000) 엑셀 업로드 (.xlsx)", type=["xlsx"])
    if uploaded is None:
        st.info("엑셀을 업로드하면 결과 표와 다운로드가 나타납니다.")
        st.stop()

    upload_day = datetime.now(KST).date()
    req_day = upload_day + timedelta(days=1)
    req_day_str = req_day.strftime("%Y-%m-%d")

    try:
        decrypted_io = decrypt_excel(uploaded.getvalue(), password=EXCEL_PASSWORD)
        excel_bytes = decrypted_io.getvalue()
        raw_df, read_meta = smart_read_orders_excel(excel_bytes)

    except Exception as e:
        st.error('엑셀 읽기/복호화 실패: 비밀번호 "0000" 또는 파일 형식을 확인해 주세요.')
        st.exception(e)
        st.stop()

    # (디버그) 안내문이 위에 있는 엑셀은 헤더를 자동 탐지합니다.
    if isinstance(read_meta, dict) and read_meta.get("method") != "header=0":
        st.caption(f"📌 헤더 자동탐지: sheet={read_meta.get('sheet')} / header_row={read_meta.get('header_row')} / {read_meta.get('method')}")

    col_name = find_col(raw_df, ["상품명", "상품", "제품명"])
    col_qty = find_col(raw_df, ["수량", "주문수량", "구매수량", "개수"])
    col_buyer = find_col(raw_df, ["구매자명", "구매자"])
    col_recv = find_col(raw_df, ["수취인명", "수령인", "받는사람"])
    col_addr = find_col(raw_df, ["통합배송지", "배송지", "주소"])
    col_opt = find_col(raw_df, ["옵션정보", "옵션", "선택옵션"])
    col_recv_phone = find_col(raw_df, ["수취인연락처", "수령인연락처", "수취인 연락처", "수령인 연락처"])
    col_msg = find_col(raw_df, ["배송메세지", "배송메시지", "배송 메시지", "배송 메세지", "배송요청사항", "요청사항"])

    missing = [k for k, v in {
        "상품명": col_name,
        "수량": col_qty,
        "구매자명": col_buyer,
        "수취인명": col_recv,
        "통합배송지": col_addr,
        "옵션정보": col_opt,
        "수취인연락처": col_recv_phone,
        "배송메세지": col_msg,
    }.items() if v is None]
    if missing:
        st.error(f"필수 컬럼을 찾지 못했습니다: {', '.join(missing)}")
        st.write("현재 컬럼:", list(raw_df.columns))
        st.stop()

    mapping_rules = load_mapping_rules()
    expr = load_expression_rules()
    bundle_units = get_bundle_units(expr)
    default_unit = normalize_text(expr.get("default_unit", "개")) or "개"

    work = raw_df[[col_buyer, col_recv, col_addr, col_recv_phone, col_msg, col_opt, col_name, col_qty]].copy()
    work.columns = ["구매자명", "수취인명", "통합배송지", "수취인연락처", "배송메세지", "옵션정보", "상품명", "수량"]

    work["상품명"] = work["상품명"].astype(str)
    work["수량"] = pd.to_numeric(work["수량"], errors="coerce")
    work["구분"] = work["상품명"].apply(extract_variant)

    mapped = work["상품명"].apply(lambda x: apply_mapping(x, mapping_rules))
    work["제품명"] = mapped.apply(lambda t: t[0])
    work["매칭성공"] = mapped.apply(lambda t: t[1])
    work["합산규칙"] = mapped.apply(lambda t: t[2])

    base = work[(work["수량"].notna()) & (work["제품명"] != "")].copy()

    exploded = explode_sum_rule_rows(
        base[["제품명", "구분", "수량", "합산규칙"]],
        bundle_units=bundle_units,
        default_unit=default_unit,
    )

    summary = (
        exploded.groupby(["제품명", "구분"], as_index=False)["수량"]
        .sum()
        .sort_values(["제품명", "구분"], kind="mergesort")
        .reset_index(drop=True)
    )
    summary["수량"] = summary["수량"].apply(fmt_qty)

    st.markdown("---")
    with st.expander("✅ 결과 (제품명 / 구분 / 수량) 펼쳐보기", expanded=False):
        st.dataframe(summary, use_container_width=True, height=520)

    with st.expander("⚠️ 미매칭/누락 행 보기 (규칙 추가용)", expanded=False):
        bad = work[(work["매칭성공"] == False) | (work["수량"].isna())].copy()
        st.dataframe(bad.head(300), use_container_width=True)

    st.download_button(
        "⬇️ 제품별 개수 PDF 다운로드",
        data=build_summary_pdf(summary),
        file_name=f"제품별개수_{datetime.now(KST).strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # 스티커 PDF
    st.markdown("---")
    st.subheader("🏷️ 스티커용지 PDF (A4 / 65칸 / 38.2×21.1mm)")

    label_rows = []
    for _, r in summary.iterrows():
        name = str(r["제품명"]).strip()
        var = str(r["구분"]).strip()
        label = name if var in ("", "-", "nan", "None") else f"{name}{var}"
        qty = _as_int_qty(r["수량"])
        if qty > 0:
            label_rows.append((label, qty))
    label_rows.sort(key=lambda x: x[0])

    sticker_texts: List[str] = []
    for label, qty in label_rows:
        sticker_texts.extend([label] * qty)

    st.caption(f"총 {len(sticker_texts)}개 · 페이지당 65칸 · 글자 {STICKER_FONT_SIZE}pt")
    st.download_button(
        "⬇️ 스티커용지 PDF 다운로드",
        data=build_sticker_pdf(sticker_texts),
        file_name=f"스티커용지_65칸_{datetime.now(KST).strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # 수취인별 출력
    st.markdown("---")
    st.subheader("📄 수취인별 출력 - 새벽배송 / 익일배송 분리")

    base2 = base.copy()
    base2["배송구분"] = base2["옵션정보"].apply(classify_delivery)
    key_cols = ["구매자명", "수취인명", "통합배송지"]

    grp_deliv = (
        base2.groupby(key_cols)["배송구분"]
        .agg(lambda x: set(x))
        .apply(decide_group_delivery)
        .reset_index()
        .rename(columns={"배송구분": "그룹배송구분"})
    )
    base2 = base2.merge(grp_deliv, on=key_cols, how="left")

    def build_items_for_group(g: pd.DataFrame) -> Tuple[str, str]:
        g = g.sort_index()
        od = OrderedDict()
        for _, r in g.iterrows():
            prod = str(r["제품명"]).strip()
            var = str(r["구분"] or "").strip()
            qty = r["수량"]
            sr = _safe_int(r.get("합산규칙", None))
            if not prod:
                continue
            if var == "":
                var = "-"
            key = (prod, var, sr)
            od[key] = od.get(key, 0.0) + float(qty)

        rows = [{"제품명": p, "구분": v, "수량": q, "합산규칙": sr} for (p, v, sr), q in od.items()]
        rows_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["제품명", "구분", "수량", "합산규칙"])

        rows_ex = explode_sum_rule_rows(
            rows_df[["제품명", "구분", "수량", "합산규칙"]],
            bundle_units=bundle_units,
            default_unit=default_unit,
        ) if len(rows_df) else rows_df

        od2 = OrderedDict()
        for _, rr in rows_ex.iterrows():
            k2 = (str(rr["제품명"]), str(rr["구분"]))
            od2[k2] = od2.get(k2, 0.0) + float(rr["수량"])

        parts = [f"{pname}/{v} {fmt_qty(q2)}" for (pname, v), q2 in od2.items()]
        recv_name = str(g["수취인명"].iloc[0]).strip()
        return recv_name, ", ".join(parts)

    group_entries = []
    for _, g in base2.groupby(key_cols, sort=False):
        recv_name, items_line = build_items_for_group(g)
        group_entries.append({"그룹배송구분": str(g["그룹배송구분"].iloc[0]), "수취인명": recv_name, "items_line": items_line})

    dawn_entries = [e for e in group_entries if e["그룹배송구분"] == "새벽배송"]
    next_entries = [e for e in group_entries if e["그룹배송구분"] == "익일배송"]

    c1, c2 = st.columns(2)
    with c1:
        st.write(f"새벽배송: {len(dawn_entries)}명")
        st.download_button(
            "⬇️ 새벽배송 수취인별 PDF",
            data=build_recipient_pdf(dawn_entries),
            file_name=f"수취인별_새벽배송_{datetime.now(KST).strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with c2:
        st.write(f"익일배송: {len(next_entries)}명")
        st.download_button(
            "⬇️ 익일배송 수취인별 PDF",
            data=build_recipient_pdf(next_entries),
            file_name=f"수취인별_익일배송_{datetime.now(KST).strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    # TC 주문 등록
    st.markdown("---")
    st.subheader("🧾 TC주문_등록양식 자동작성 (새벽/익일 각각 엑셀 생성)")

    if not TC_TEMPLATE_DEFAULT_PATH.exists():
        st.error("앱 폴더에 'TC주문_등록양식.xlsx' 파일이 없습니다. GitHub에 app.py와 같이 올려주세요.")
    else:
        template_bytes = TC_TEMPLATE_DEFAULT_PATH.read_bytes()

        # 수취인별 출력 순서(원본 등장 순서)
        order_keys_df = base2[key_cols].drop_duplicates(keep="first").copy()

        def _first_nonempty(series: pd.Series) -> str:
            for v in series.tolist():
                if v is None:
                    continue
                s = str(v).strip()
                if s and s.lower() != "nan":
                    return s
            return ""

        grp_info_agg = (
            base2.groupby(key_cols, as_index=False)
            .agg(
                그룹배송구분=("그룹배송구분", "first"),
                수취인연락처=("수취인연락처", _first_nonempty),
                배송메세지=("배송메세지", _first_nonempty),
                구매자명=("구매자명", "first"),
                수취인명=("수취인명", "first"),
                통합배송지=("통합배송지", "first"),
            )
        )
        grp_info = order_keys_df.merge(grp_info_agg, on=key_cols, how="left")

        def make_tc_rows(df: pd.DataFrame, ship: str) -> List[Dict[str, str]]:
            out = []
            ship_type = st.session_state.tc_type_dawn if ship == "새벽배송" else st.session_state.tc_type_next
            ship_type = (ship_type or "").strip() or (TC_TYPE_DAWN_DEFAULT if ship == "새벽배송" else TC_TYPE_NEXT_DEFAULT)

            for _, r in df.iterrows():
                out.append(
                    {
                        "배송요청일": req_day_str,
                        "주문자": str(r["구매자명"] or "").strip(),
                        "수령자": str(r["수취인명"] or "").strip(),
                        "수령자도로명주소": str(r["통합배송지"] or "").strip(),
                        "수령자연락처": str(r.get("수취인연락처", "") or "").strip(),
                        "출입방법": _clean_access_message(r.get("배송메세지", "")),
                        "상품명": TC_PRODUCT_NAME_FIXED,
                        "배송유형": ship_type,
                    }
                )
            return out

        dawn_df = grp_info[grp_info["그룹배송구분"] == "새벽배송"].copy()
        next_df = grp_info[grp_info["그룹배송구분"] == "익일배송"].copy()

        cols = st.columns(2)
        with cols[0]:
            st.write(f"새벽배송 행: {len(dawn_df)} (배송유형: {st.session_state.tc_type_dawn})")
            if len(dawn_df):
                out_bytes = build_tc_excel_bytes(template_bytes, make_tc_rows(dawn_df, "새벽배송"))
                st.download_button(
                    "⬇️ TC주문_등록양식(새벽배송) 엑셀 다운로드",
                    data=out_bytes,
                    file_name=f"TC주문_등록양식_새벽배송_{datetime.now(KST).strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with cols[1]:
            st.write(f"익일배송 행: {len(next_df)} (배송유형: {st.session_state.tc_type_next})")
            if len(next_df):
                out_bytes = build_tc_excel_bytes(template_bytes, make_tc_rows(next_df, "익일배송"))
                st.download_button(
                    "⬇️ TC주문_등록양식(익일배송) 엑셀 다운로드",
                    data=out_bytes,
                    file_name=f"TC주문_등록양식_익일배송_{datetime.now(KST).strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
