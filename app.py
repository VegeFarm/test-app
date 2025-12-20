import io
import json
import re
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple, List, Dict
from collections import OrderedDict

import pandas as pd
import streamlit as st

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
EXCEL_PASSWORD = "0000"

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

MAPPING_PATH = DATA_DIR / "name_mappings.json"
SUMRULES_PATH = DATA_DIR / "sum_rules.json"

BACKUP_DIR = DATA_DIR / "rules_backup"
BACKUP_DIR.mkdir(exist_ok=True)

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


# =====================================================
# VARIANT(단위) 추출
# =====================================================
UNIT_PATTERNS = [
    r"\d+(?:\.\d+)?kg\s*~\s*\d+(?:\.\d+)?kg",
    r"\d+(?:\.\d+)?kg",
    r"(?:약\s*)?\d+(?:\.\d+)?g",
    r"\d+개", r"\d+통", r"\d+단", r"\d+봉", r"\d+팩",
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


# =====================================================
# RULES: 1) 상품명 매칭규칙
# =====================================================
def default_mapping_rules() -> List[Dict]:
    # ✅ 규칙은 "위에서 아래 순서대로" 매칭되며, 첫 매칭 1개만 적용
    return [
        {
            "enabled": True,
            "match_type": "contains",
            "pattern": "와일드루꼴라",
            "display_name": "와일드",
            "note": '예) "채소팜 와일드루꼴라 1kg ..." -> 와일드',
        },
        {
            "enabled": True,
            "match_type": "contains",
            "pattern": "라디치오",
            "display_name": "라디치오",
            "note": '예) "채소팜 라디치오 1통 ..." -> 라디치오',
        },
    ]


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
                # 과거 버전 sum_rule/priority 같은 필드가 있어도 무시
                cleaned.append(
                    dict(
                        enabled=bool(r.get("enabled", True)),
                        match_type=normalize_text(r.get("match_type", "contains")) or "contains",
                        pattern=normalize_text(r.get("pattern", "")),
                        display_name=normalize_text(r.get("display_name", "")),
                        note=normalize_text(r.get("note", "")),
                    )
                )
            return cleaned
    except Exception:
        pass

    rules = default_mapping_rules()
    save_mapping_rules(rules)
    return rules


def save_mapping_rules(rules: List[Dict]) -> None:
    MAPPING_PATH.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")


def apply_mapping(actual_name: str, rules: List[Dict]) -> Tuple[str, bool]:
    """
    return: (제품명, 매칭성공여부)
    """
    actual = normalize_text(actual_name)
    if not actual:
        return "", False

    for r in rules:
        if not r.get("enabled", True):
            continue

        mt = normalize_text(r.get("match_type", "contains")) or "contains"
        pattern = normalize_text(r.get("pattern", ""))
        display = normalize_text(r.get("display_name", ""))

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
            return display, True

    # --- fallback ---
    s = re.sub(r"^\s*채소팜\s*", "", actual)
    s = re.sub(r"\([^)]*\)", "", s).strip()
    s = re.sub(r"\s+", " ", s).strip()
    m = UNIT_RE.search(s)
    if m:
        s = s[: m.start()].strip()

    toks = s.split()
    if not toks:
        return actual, False

    PREFIX = {"생", "유기농", "국산", "수입", "냉동", "베이비", "프리미엄"}
    if len(toks) >= 2 and toks[0] in PREFIX:
        fallback = toks[0] + toks[1]
    else:
        fallback = toks[0]

    return fallback, False


def mapping_df_from_list(rules: List[Dict]) -> pd.DataFrame:
    df = pd.DataFrame(rules)
    keep = ["enabled", "match_type", "pattern", "display_name", "note"]
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

        cleaned.append(
            dict(
                enabled=bool(row.get("enabled", True)),
                match_type=mt,
                pattern=pattern,
                display_name=display,
                note=normalize_text(row.get("note")),
            )
        )
    return cleaned


# =====================================================
# RULES: 2) 합산규칙(별도 관리)
# =====================================================
def default_sum_rules() -> List[Dict]:
    # 제품명(표시될 상품명)에 대해 합산규칙 N을 매칭
    # (실제 분해/합산은 개/봉/통/팩 단위에만 적용됨)
    return [
        {
            "enabled": False,
            "match_type": "exact",
            "pattern": "오렌지",
            "sum_rule": 5,
            "note": "예: 오렌지 합산규칙 5",
        }
    ]


def load_sum_rules(mapping_rules: Optional[List[Dict]] = None) -> List[Dict]:
    # sum_rules.json이 없으면: 기존 name_mappings.json에 sum_rule이 들어있던 옛 버전을 자동 흡수(가능하면)
    if SUMRULES_PATH.exists():
        try:
            raw = json.loads(SUMRULES_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, list):
                out = []
                for r in raw:
                    sr = _safe_int(r.get("sum_rule"))
                    if sr is not None and sr < 2:
                        sr = None
                    out.append(
                        dict(
                            enabled=bool(r.get("enabled", True)),
                            match_type=normalize_text(r.get("match_type", "exact")) or "exact",
                            pattern=normalize_text(r.get("pattern", "")),
                            sum_rule=sr,
                            note=normalize_text(r.get("note", "")),
                        )
                    )
                return out
        except Exception:
            pass

    # ---- legacy absorb ----
    legacy = []
    try:
        if MAPPING_PATH.exists():
            raw = json.loads(MAPPING_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, list):
                for r in raw:
                    sr = _safe_int(r.get("sum_rule"))
                    dn = normalize_text(r.get("display_name", ""))
                    if sr and sr >= 2 and dn:
                        legacy.append(
                            dict(
                                enabled=bool(r.get("enabled", True)),
                                match_type="exact",
                                pattern=dn,
                                sum_rule=sr,
                                note="(구버전에서 자동 이관)",
                            )
                        )
    except Exception:
        legacy = []

    if legacy:
        save_sum_rules(legacy)
        return legacy

    rules = default_sum_rules()
    save_sum_rules(rules)
    return rules


def save_sum_rules(rules: List[Dict]) -> None:
    SUMRULES_PATH.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")


def apply_sum_rule(product_name: str, sum_rules: List[Dict]) -> Optional[int]:
    """
    제품명(표시될 상품명)을 기준으로 합산규칙 N 반환 (첫 매칭 1개)
    """
    p = normalize_text(product_name)
    if not p:
        return None

    for r in sum_rules:
        if not r.get("enabled", True):
            continue

        mt = normalize_text(r.get("match_type", "exact")) or "exact"
        pattern = normalize_text(r.get("pattern", ""))
        sr = _safe_int(r.get("sum_rule", None))

        if not pattern or sr is None or sr < 2:
            continue

        matched = False
        if mt == "exact":
            matched = (p == pattern)
        elif mt == "contains":
            matched = (pattern in p)
        elif mt == "regex":
            try:
                matched = bool(re.search(pattern, p))
            except re.error:
                matched = False

        if matched:
            return sr

    return None


def sumrules_df_from_list(rules: List[Dict]) -> pd.DataFrame:
    df = pd.DataFrame(rules)
    keep = ["enabled", "match_type", "pattern", "sum_rule", "note"]
    for c in keep:
        if c not in df.columns:
            df[c] = None
    return df[keep]


def sumrules_list_from_df(edited: pd.DataFrame) -> List[Dict]:
    cleaned = []
    for _, row in edited.iterrows():
        pattern = normalize_text(row.get("pattern"))
        sr = _safe_int(row.get("sum_rule"))

        if not pattern or sr is None or sr < 2:
            # pattern 또는 sum_rule이 없으면 스킵 (삭제 효과)
            continue

        mt = normalize_text(row.get("match_type")) or "exact"
        if mt not in {"contains", "exact", "regex"}:
            mt = "exact"

        cleaned.append(
            dict(
                enabled=bool(row.get("enabled", True)),
                match_type=mt,
                pattern=pattern,
                sum_rule=sr,
                note=normalize_text(row.get("note")),
            )
        )
    return cleaned


# =====================================================
# Backups (Excel)
# =====================================================
def backup_rules_to_excel(mapping_rules: List[Dict], sum_rules: List[Dict]) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = BACKUP_DIR / f"rules_backup_{ts}.xlsx"

    df_map = mapping_df_from_list(mapping_rules).rename(
        columns={
            "enabled": "사용",
            "match_type": "매칭방식",
            "pattern": "실제상품명(패턴)",
            "display_name": "표시될상품명",
            "note": "메모",
        }
    )
    df_sum = sumrules_df_from_list(sum_rules).rename(
        columns={
            "enabled": "사용",
            "match_type": "매칭방식",
            "pattern": "제품명(패턴)",
            "sum_rule": "합산규칙(N)",
            "note": "메모",
        }
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_map.to_excel(writer, sheet_name="상품명매칭", index=False)
        df_sum.to_excel(writer, sheet_name="합산규칙", index=False)

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


def find_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    cols = list(df.columns)
    for k in keywords:
        if k in cols:
            return k
    for c in cols:
        cs = str(c)
        for k in keywords:
            if k in cs:
                return c
    return None


# =====================================================
# 합산규칙(개/봉/통/팩) 적용
# =====================================================
BUNDLE_UNITS = {"개", "봉", "통", "팩"}
BUNDLE_RE = re.compile(r"^\s*(\d+)\s*(개|봉|통|팩)\s*$")


def parse_bundle_variant(variant: str) -> Tuple[Optional[int], Optional[str]]:
    m = BUNDLE_RE.match((variant or "").strip())
    if not m:
        return None, None
    try:
        return int(m.group(1)), m.group(2)
    except Exception:
        return None, None


def explode_sum_rule_rows(df_rows: pd.DataFrame) -> pd.DataFrame:
    """
    df_rows columns: 제품명, 구분, 수량, 합산규칙
    합산규칙은 개/봉/통/팩 단위에서만 분해됩니다.
    """
    out = []
    for _, r in df_rows.iterrows():
        product = r["제품명"]
        variant = (r.get("구분", "") or "").strip()
        qty = r.get("수량", None)
        rule_n = _safe_int(r.get("합산규칙", None))

        if rule_n is None or rule_n < 2:
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        if variant == "":
            unit_size, unit_label = 1, "개"
            is_bundle = True
        else:
            unit_size, unit_label = parse_bundle_variant(variant)
            is_bundle = (unit_size is not None and unit_label in BUNDLE_UNITS)

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
# PDF helpers
# =====================================================
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
    rest = txt[len(line1):].strip()
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
    try:
        t.setTextRenderMode(0)
    except Exception:
        pass

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
# Streamlit UI
# =====================================================
st.set_page_config(page_title="제품별 개수 & 수취인별 출력", page_icon="📄", layout="wide")
st.title("📄 제품별 개수 & 수취인별 출력")
st.caption('엑셀 업로드 → (상품명 매칭/합산규칙) → 제품별 집계 + 수취인별(새벽/익일) PDF + 스티커용지 PDF (엑셀 비밀번호 "0000" 고정)')

menu = st.sidebar.radio("메뉴", ["🧩 상품명 매칭 규칙", "⬆️ 엑셀 업로드 & 결과"], index=1)
st.sidebar.markdown("---")


# -----------------------------
# Sidebar: (상품명 매칭 규칙 메뉴에서만) 백업폴더 + 합산규칙 관리
# -----------------------------
def sidebar_backup_and_sumrules(mapping_rules: List[Dict], sum_rules: List[Dict]):
    # 1) 백업폴더 (펼쳐보기 + 다운로드 + 삭제)
    with st.sidebar.expander("📁 규칙 백업폴더", expanded=False):
        try:
            backups = sorted(BACKUP_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        except Exception:
            backups = []

        if not backups:
            st.caption("아직 백업 파일이 없습니다.")
        else:
            for i, fp in enumerate(backups[:50]):
                cols = st.columns([6, 2, 2])
                cols[0].write(fp.name)

                # 다운로드
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

                # 삭제
                if cols[2].button("삭제", key=f"rm_bk_{i}_{fp.name}", use_container_width=True):
                    try:
                        fp.unlink()
                        st.success(f"삭제 완료: {fp.name}")
                        st.rerun()
                    except Exception as e:
                        st.error("삭제 실패")
                        st.exception(e)

    # 2) 합산규칙 편집 (펼쳐보기 + 수정/추가/삭제 + 저장)
    with st.sidebar.expander("➕ 합산규칙 관리", expanded=False):
        st.caption("※ 합산규칙은 개/봉/통/팩 단위에서만 분해/합산됩니다.")
        df = sumrules_df_from_list(sum_rules)

        edited = st.data_editor(
            df,
            hide_index=True,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "enabled": st.column_config.CheckboxColumn("사용", default=True),
                "match_type": st.column_config.SelectboxColumn("매칭", options=["exact", "contains", "regex"]),
                "pattern": st.column_config.TextColumn("제품명(패턴)"),
                "sum_rule": st.column_config.NumberColumn("N", min_value=2, step=1),
                "note": st.column_config.TextColumn("메모"),
            },
            key="sumrules_editor_sidebar",
        )

        if st.button("💾 합산규칙 저장", use_container_width=True, key="save_sumrules_btn"):
            cleaned = sumrules_list_from_df(edited)
            save_sum_rules(cleaned)
            st.success(f"합산규칙 저장 완료 ({len(cleaned)}개)")
            st.rerun()

        # 편의: 현재 적용 예시 한 줄
        st.caption("예: 제품명=오렌지, N=5 → 8개 주문 시 5개 1개 + 3개 1개")


# -----------------------------
# 1) 규칙 관리
# -----------------------------
if menu == "🧩 상품명 매칭 규칙":
    mapping_rules = load_mapping_rules()
    sum_rules = load_sum_rules(mapping_rules)

    sidebar_backup_and_sumrules(mapping_rules, sum_rules)

    st.subheader("실제 상품명 → 표시될 상품명")

    st.markdown(
        """
**매칭방식 설명**
- **contains**: `패턴`이 `엑셀 상품명` 안에 **포함**되어 있으면 매칭
- **exact**: `패턴`과 `엑셀 상품명`이 **완전히 동일**할 때만 매칭
- **regex**: `패턴`을 **정규식**으로 해석해 매칭 (예: `와일드루꼴라\\s*(250g|500g|1kg)`)
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
            # sumrules는 사이드바 에디터가 따로 있으니 현재 파일 기준 로드해서 같이 백업
            cleaned_sum = load_sum_rules()
            outp = backup_rules_to_excel(cleaned_map, cleaned_sum)
            st.success(f"백업 저장 완료: {outp.name}")
            st.rerun()


# -----------------------------
# 2) 엑셀 업로드 & 결과
# -----------------------------
else:
    st.subheader("엑셀 업로드 → 제품별 집계 + 수취인별 출력(새벽/익일 분리)")

    if msoffcrypto is None:
        st.error("msoffcrypto가 설치되지 않았습니다. requirements.txt에 'msoffcrypto-tool'을 추가하고 재배포해 주세요.")
        st.stop()

    uploaded = st.file_uploader("비밀번호(0000) 엑셀 업로드 (.xlsx)", type=["xlsx"])
    if uploaded is None:
        st.info("엑셀을 업로드하면 결과 표와 PDF 다운로드가 나타납니다.")
        st.stop()

    try:
        decrypted = decrypt_excel(uploaded.getvalue(), password=EXCEL_PASSWORD)
        raw_df = pd.read_excel(decrypted, sheet_name=0, engine="openpyxl")
    except Exception as e:
        st.error('엑셀 읽기/복호화 실패: 비밀번호 "0000" 또는 파일 형식을 확인해 주세요.')
        st.exception(e)
        st.stop()

    col_name = find_col(raw_df, ["상품명", "상품", "제품명"])
    col_qty = find_col(raw_df, ["수량", "주문수량", "구매수량", "개수"])
    col_buyer = find_col(raw_df, ["구매자명", "구매자"])
    col_recv = find_col(raw_df, ["수취인명", "수령인", "받는사람"])
    col_addr = find_col(raw_df, ["통합배송지", "배송지", "주소"])
    col_opt = find_col(raw_df, ["옵션정보", "옵션", "선택옵션"])

    missing = [k for k, v in {
        "상품명": col_name,
        "수량": col_qty,
        "구매자명": col_buyer,
        "수취인명": col_recv,
        "통합배송지": col_addr,
        "옵션정보": col_opt,
    }.items() if v is None]

    if missing:
        st.error(f"필수 컬럼을 찾지 못했습니다: {', '.join(missing)}")
        st.write("현재 컬럼:", list(raw_df.columns))
        st.stop()

    mapping_rules = load_mapping_rules()
    sum_rules = load_sum_rules(mapping_rules)

    work = raw_df[[col_buyer, col_recv, col_addr, col_opt, col_name, col_qty]].copy()
    work.columns = ["구매자명", "수취인명", "통합배송지", "옵션정보", "상품명", "수량"]

    work["상품명"] = work["상품명"].astype(str)
    work["수량"] = pd.to_numeric(work["수량"], errors="coerce")
    work["구분"] = work["상품명"].apply(extract_variant)

    mapped = work["상품명"].apply(lambda x: apply_mapping(x, mapping_rules))
    work["제품명"] = mapped.apply(lambda t: t[0])
    work["매칭성공"] = mapped.apply(lambda t: t[1])

    # ✅ 합산규칙은 "제품명" 기준으로 별도 적용
    work["합산규칙"] = work["제품명"].apply(lambda p: apply_sum_rule(p, sum_rules))

    base = work[(work["수량"].notna()) & (work["제품명"] != "")].copy()

    exploded = explode_sum_rule_rows(base[["제품명", "구분", "수량", "합산규칙"]])
    summary = (
        exploded.groupby(["제품명", "구분"], as_index=False)["수량"]
        .sum()
        .sort_values(["제품명", "구분"], kind="mergesort")
        .reset_index(drop=True)
    )
    summary["수량"] = summary["수량"].apply(fmt_qty)

    st.markdown("---")
    st.subheader("✅ 결과 (제품명 / 구분 / 수량)")
    st.dataframe(summary, use_container_width=True, height=520)

    summary_pdf = build_summary_pdf(summary)
    st.download_button(
        "⬇️ 제품별 개수 PDF 다운로드",
        data=summary_pdf,
        file_name=f"제품별개수_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # -----------------------------
    # 스티커 PDF
    # -----------------------------
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

    label_rows.sort(key=lambda x: x[0])  # 가나다(유니코드) 순

    sticker_texts: List[str] = []
    for label, qty in label_rows:
        sticker_texts.extend([label] * qty)

    pages_needed = (len(sticker_texts) + STICKER_PER_PAGE - 1) // STICKER_PER_PAGE if sticker_texts else 0
    st.caption(f"총 스티커 {len(sticker_texts)}개 · {pages_needed}페이지 (페이지당 65칸 / 글자 {STICKER_FONT_SIZE}pt)")

    sticker_pdf = build_sticker_pdf(sticker_texts)
    st.download_button(
        "⬇️ 스티커용지 PDF 다운로드",
        data=sticker_pdf,
        file_name=f"스티커용지_65칸_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # -----------------------------
    # 수취인별 출력
    # -----------------------------
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
            if key not in od:
                od[key] = 0.0
            try:
                od[key] += float(qty)
            except Exception:
                pass

        rows = [{"제품명": p, "구분": v, "수량": q, "합산규칙": sr} for (p, v, sr), q in od.items()]
        rows_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["제품명", "구분", "수량", "합산규칙"])
        rows_ex = explode_sum_rule_rows(rows_df[["제품명", "구분", "수량", "합산규칙"]]) if len(rows_df) else rows_df

        od2 = OrderedDict()
        for _, rr in rows_ex.iterrows():
            k2 = (str(rr["제품명"]), str(rr["구분"]))
            if k2 not in od2:
                od2[k2] = 0.0
            try:
                od2[k2] += float(rr["수량"])
            except Exception:
                pass

        parts = [f"{pname}/{v} {fmt_qty(q2)}" for (pname, v), q2 in od2.items()]
        recv_name = str(g["수취인명"].iloc[0]).strip()
        return recv_name, ", ".join(parts)

    group_entries = []
    for _, g in base2.groupby(key_cols, sort=False):
        recv_name, items_line = build_items_for_group(g)
        group_entries.append(
            {
                "그룹배송구분": str(g["그룹배송구분"].iloc[0]),
                "수취인명": recv_name,
                "items_line": items_line if items_line else "",
            }
        )

    dawn_entries = [e for e in group_entries if e["그룹배송구분"] == "새벽배송"]
    next_entries = [e for e in group_entries if e["그룹배송구분"] == "익일배송"]

    c1, c2 = st.columns(2)
    with c1:
        st.write(f"새벽배송: {len(dawn_entries)}명")
        dawn_pdf = build_recipient_pdf(dawn_entries)
        st.download_button(
            "⬇️ 새벽배송 수취인별 PDF",
            data=dawn_pdf,
            file_name=f"수취인별_새벽배송_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with c2:
        st.write(f"익일배송: {len(next_entries)}명")
        next_pdf = build_recipient_pdf(next_entries)
        st.download_button(
            "⬇️ 익일배송 수취인별 PDF",
            data=next_pdf,
            file_name=f"수취인별_익일배송_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with st.expander("⚠️ 미매칭/누락 행 보기 (규칙 추가용)", expanded=False):
        bad = work[(work["매칭성공"] == False) | (work["수량"].isna())].copy()
        st.dataframe(bad.head(300), use_container_width=True)
