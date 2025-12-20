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

# 수취인별 PDF 스타일(원하면 여기만 수정)
RECIPIENT_FONT_SIZE = 12
RECIPIENT_LEADING = 15
RECIPIENT_BLOCK_GAP_MM = 4.0   # 한 사람 블록 아래 여백(Spacer)
RECIPIENT_LINE_AFTER_MM = 4.0  # 구분선 아래 여백(spaceAfter)

# 스티커 용지 설정 (A4 / 65칸 / 38.2x21.1mm)
STICKER_COLS = 5
STICKER_ROWS = 13
STICKER_PER_PAGE = STICKER_COLS * STICKER_ROWS  # 65
STICKER_CELL_W_MM = 38.2
STICKER_CELL_H_MM = 21.1

# ✅ 스티커 텍스트 스타일
STICKER_FONT_SIZE = 9     # 현재 글자 크기(포인트)
STICKER_LEADING = 11
STICKER_BOLD = True

# ✅ 깔끔한 Bold: fill + stroke 방식(겹쳐그리기 제거)
STICKER_BOLD_MODE = "stroke"    # "stroke" 권장
STICKER_STROKE_WIDTH = 0.12     # 0.18~0.30 사이에서 조절 (지저분하면 더 낮추기)


# =====================================================
# VARIANT(단위) 추출
# =====================================================
UNIT_PATTERNS = [
    r"\d+(?:\.\d+)?kg\s*~\s*\d+(?:\.\d+)?kg",  # 1.8kg~2kg
    r"\d+(?:\.\d+)?kg",                        # 1kg, 1.5kg
    r"(?:약\s*)?\d+(?:\.\d+)?g",               # 500g, 약350g
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
# RULES (상품명 매칭 + 합산규칙)
# =====================================================
def default_rules() -> List[Dict]:
    return [
        {
            "enabled": True,
            "priority": 10,
            "match_type": "contains",
            "pattern": "와일드루꼴라",
            "display_name": "와일드",
            "sum_rule": None,
            "note": '예) "채소팜 와일드루꼴라 1kg ..." -> 와일드',
        },
        {
            "enabled": True,
            "priority": 20,
            "match_type": "contains",
            "pattern": "라디치오",
            "display_name": "라디치오",
            "sum_rule": None,
            "note": '예) "채소팜 라디치오 1통 ..." -> 라디치오',
        },
        {
            "enabled": False,
            "priority": 30,
            "match_type": "contains",
            "pattern": "오렌지",
            "display_name": "오렌지",
            "sum_rule": 5,
            "note": "합산규칙=5 예시 (개/봉/통/팩)",
        },
    ]


def load_rules() -> List[Dict]:
    # 파일이 없으면 최초 1회만 예시 생성(초기화 버튼 없음)
    if not MAPPING_PATH.exists():
        rules = default_rules()
        save_rules(rules)
        return rules

    try:
        raw = json.loads(MAPPING_PATH.read_text(encoding="utf-8"))
        if isinstance(raw, list):
            return raw
    except Exception:
        pass

    rules = default_rules()
    save_rules(rules)
    return rules


def save_rules(rules: List[Dict]) -> None:
    MAPPING_PATH.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")


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


def apply_mapping(actual_name: str, rules: List[Dict]) -> Tuple[str, bool, Optional[int]]:
    """
    return: (제품명, 매칭성공여부, 합산규칙N or None)
    """
    actual = normalize_text(actual_name)
    if not actual:
        return "", False, None

    def prio(r):
        try:
            return int(r.get("priority", 9999))
        except Exception:
            return 9999

    for r in sorted(rules, key=prio):
        if not r.get("enabled", True):
            continue

        mt = normalize_text(r.get("match_type", "contains")) or "contains"
        pattern = normalize_text(r.get("pattern", ""))
        display = normalize_text(r.get("display_name", ""))
        sum_rule = _safe_int(r.get("sum_rule"))

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
            if sum_rule is not None and sum_rule < 2:
                sum_rule = None
            return display, True, sum_rule

    # --- fallback: 브랜드/괄호 제거 + 단위 앞까지만 + 접두어 처리(생/유기농...) ---
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
        fallback = toks[0] + toks[1]  # 예: "생 아스파라거스" -> "생아스파라거스"
    else:
        fallback = toks[0]

    return fallback, False, None


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
    columns required: 제품명, 구분, 수량, 합산규칙
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

        # 단위 판단
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
# PDF 2) 수취인별 출력 (행별로 수취인명 길이에 맞춰 붙이기)
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
# PDF 3) 스티커 용지 (Canvas로 직접 그림)
#   ✅ 페이지당 65칸(5x13) 고정 + 깔끔한 Bold(Stroke)
# =====================================================
def _wrap_for_cell(txt: str, font_name: str, font_size: int, max_w_pt: float) -> List[str]:
    """
    셀 내부에 들어가도록 최대 2줄로 래핑. 너무 길면 ... 처리
    """
    txt = (txt or "").strip()
    if not txt:
        return [""]

    def w(s: str) -> float:
        return _text_width_pt(s, font_name, font_size)

    if w(txt) <= max_w_pt:
        return [txt]

    # 공백이 있으면 공백 기준 래핑 시도
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

    # 공백이 없으면 글자 단위로 자르기
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


def _draw_center_text(
    c: canvas.Canvas,
    font_name: str,
    font_size: int,
    x_center: float,
    y: float,
    txt: str,
    bold: bool,
):
    """
    가운데 정렬 텍스트를 출력.
    bold=True이면 fill+stroke 방식(가능한 경우)으로 깔끔하게 두껍게 출력.
    """
    txt = (txt or "").strip()
    if not txt:
        return

    w = _text_width_pt(txt, font_name, font_size)
    x_left = x_center - (w / 2.0)

    t = c.beginText()
    t.setTextOrigin(x_left, y)
    t.setFont(font_name, font_size)

    if bold and STICKER_BOLD and STICKER_BOLD_MODE == "stroke":
        # 2 = fill + stroke
        try:
            t.setTextRenderMode(2)
            c.setStrokeColor(colors.black)
            c.setLineWidth(float(STICKER_STROKE_WIDTH))
        except Exception:
            # setTextRenderMode가 없는 환경이면 그냥 일반 출력(깔끔 우선)
            try:
                t.setTextRenderMode(0)
            except Exception:
                pass
    else:
        try:
            t.setTextRenderMode(0)
        except Exception:
            pass

    t.textOut(txt)
    c.drawText(t)


def build_sticker_pdf(label_texts: List[str], show_grid: bool = False) -> bytes:
    """
    A4 / 65칸(5x13) / 38.2x21.1mm
    label_texts는 '수량만큼 확장된 텍스트 리스트'
    """
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

    # 가운데 정렬(아래 기준)
    x0 = (page_w_pt - grid_w_pt) / 2.0
    y0 = (page_h_pt - grid_h_pt) / 2.0

    total = len(label_texts)
    page_count = (total + STICKER_PER_PAGE - 1) // STICKER_PER_PAGE if total else 1

    # 셀 안쪽 패딩
    pad_x = 2.0 * mm
    max_text_w = cell_w_pt - (pad_x * 2)

    for p in range(page_count):
        # 가이드선(테두리)
        if show_grid:
            c.setLineWidth(0.3)
            c.setStrokeColor(colors.lightgrey)
            for r in range(STICKER_ROWS):
                for col in range(STICKER_COLS):
                    x = x0 + col * cell_w_pt
                    y = y0 + (STICKER_ROWS - 1 - r) * cell_h_pt
                    c.rect(x, y, cell_w_pt, cell_h_pt, stroke=1, fill=0)

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

                # 최대 2줄
                lines = _wrap_for_cell(text, font_name, STICKER_FONT_SIZE, max_text_w)[:2]

                cx = x + cell_w_pt / 2.0
                if len(lines) == 1:
                    cy = y + (cell_h_pt / 2.0) - (STICKER_FONT_SIZE * 0.35)
                    _draw_center_text(c, font_name, STICKER_FONT_SIZE, cx, cy, lines[0], bold=True)
                else:
                    center = y + (cell_h_pt / 2.0)
                    upper_y = center + (STICKER_LEADING * 0.25)
                    lower_y = center - (STICKER_LEADING * 0.95)
                    _draw_center_text(c, font_name, STICKER_FONT_SIZE, cx, upper_y, lines[0], bold=True)
                    _draw_center_text(c, font_name, STICKER_FONT_SIZE, cx, lower_y, lines[1], bold=True)

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
st.sidebar.caption("규칙 파일: data/name_mappings.json")


# -----------------------------
# 1) 규칙 관리
# -----------------------------
if menu == "🧩 상품명 매칭 규칙":
    st.subheader("실제 상품명 → 표시될 상품명 + 합산규칙(개/봉/통/팩)")

    st.markdown(
        """
**매칭방식 설명**
- **contains**: `패턴`이 `엑셀 상품명` 안에 **포함**되어 있으면 매칭 (가장 많이 쓰는 방식)
- **exact**: `패턴`과 `엑셀 상품명`이 **완전히 동일**할 때만 매칭
- **regex**: `패턴`을 **정규식**으로 해석해 매칭 (예: `와일드루꼴라\\s*(250g|500g|1kg)`)

**우선순위(priority)**: 숫자가 **작을수록 먼저 적용**됩니다.
"""
    )

    rules = load_rules()
    df = pd.DataFrame(rules)
    for col in ["enabled", "priority", "match_type", "pattern", "display_name", "sum_rule", "note"]:
        if col not in df.columns:
            df[col] = None
    df = df[["enabled", "priority", "match_type", "pattern", "display_name", "sum_rule", "note"]]

    edited = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "enabled": st.column_config.CheckboxColumn("사용", default=True),
            "priority": st.column_config.NumberColumn("우선순위", help="작을수록 먼저 적용", min_value=0, step=1),
            "match_type": st.column_config.SelectboxColumn("매칭 방식", options=["contains", "exact", "regex"]),
            "pattern": st.column_config.TextColumn("실제 상품명(패턴)", width="large"),
            "display_name": st.column_config.TextColumn("표시될 상품명", width="medium"),
            "sum_rule": st.column_config.NumberColumn(
                "합산규칙(N)",
                help="개/봉/통/팩 상품을 N묶음으로 표현 (비우면 미적용)",
                min_value=2,
                step=1,
            ),
            "note": st.column_config.TextColumn("메모", width="large"),
        },
        key="mapping_editor",
    )

    c1, c2 = st.columns([1, 2])
    with c1:
        if st.button("💾 저장", use_container_width=True):
            cleaned = []
            for _, row in edited.iterrows():
                pattern = normalize_text(row.get("pattern"))
                display = normalize_text(row.get("display_name"))
                if not pattern or not display:
                    continue

                mt = normalize_text(row.get("match_type")) or "contains"
                if mt not in {"contains", "exact", "regex"}:
                    mt = "contains"

                try:
                    pr = int(row.get("priority", 9999))
                except Exception:
                    pr = 9999

                sr = _safe_int(row.get("sum_rule"))
                if sr is not None and sr < 2:
                    sr = None

                cleaned.append(
                    dict(
                        enabled=bool(row.get("enabled", True)),
                        priority=pr,
                        match_type=mt,
                        pattern=pattern,
                        display_name=display,
                        sum_rule=sr,
                        note=normalize_text(row.get("note")),
                    )
                )

            save_rules(cleaned)
            st.success(f"저장 완료! (규칙 {len(cleaned)}개)")

    with c2:
        export_bytes = json.dumps(load_rules(), ensure_ascii=False, indent=2).encode("utf-8")
        st.download_button(
            "⬇️ 규칙 내보내기(JSON)",
            data=export_bytes,
            file_name="name_mappings.json",
            mime="application/json",
            use_container_width=True,
        )

    st.info("합산규칙은 개/봉/통/팩 단위에만 적용됩니다. (kg/g 등은 미적용)")


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

    # 필요한 컬럼
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

    rules = load_rules()

    work = raw_df[[col_buyer, col_recv, col_addr, col_opt, col_name, col_qty]].copy()
    work.columns = ["구매자명", "수취인명", "통합배송지", "옵션정보", "상품명", "수량"]

    work["상품명"] = work["상품명"].astype(str)
    work["수량"] = pd.to_numeric(work["수량"], errors="coerce")

    work["구분"] = work["상품명"].apply(extract_variant)

    mapped = work["상품명"].apply(lambda x: apply_mapping(x, rules))
    work["제품명"] = mapped.apply(lambda t: t[0])
    work["매칭성공"] = mapped.apply(lambda t: t[1])
    work["합산규칙"] = mapped.apply(lambda t: t[2])

    # -----------------------------
    # (A) 제품별 집계
    # -----------------------------
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
    # (A-2) 스티커 용지 PDF (65칸 고정 + 깔끔한 Bold)
    # -----------------------------
    st.markdown("---")
    st.subheader("🏷️ 스티커용지 PDF (A4 / 65칸 / 38.2×21.1mm)")

    show_grid = st.checkbox("가이드선(테두리) 표시", value=False)

    # 가나다 순 정렬 후, 수량만큼 확장
    label_rows = []
    for _, r in summary.iterrows():
        name = str(r["제품명"]).strip()
        var = str(r["구분"]).strip()

        if var in ("", "-", "nan", "None"):
            label = name
        else:
            # 제품명 + 구분 합치기(공백 없이) 예: 가지1개, 건대추500g
            label = f"{name}{var}"

        qty = _as_int_qty(r["수량"])
        if qty > 0:
            label_rows.append((label, qty))

    label_rows.sort(key=lambda x: x[0])  # 가나다(유니코드) 순

    sticker_texts: List[str] = []
    for label, qty in label_rows:
        sticker_texts.extend([label] * qty)

    pages_needed = (len(sticker_texts) + STICKER_PER_PAGE - 1) // STICKER_PER_PAGE if sticker_texts else 0
    st.caption(
        f"총 스티커 {len(sticker_texts)}개 · {pages_needed}페이지 "
        f"(페이지당 65칸 고정 / 글자 {STICKER_FONT_SIZE}pt / Bold={STICKER_BOLD} / mode={STICKER_BOLD_MODE})"
    )

    sticker_pdf = build_sticker_pdf(sticker_texts, show_grid=show_grid)
    st.download_button(
        "⬇️ 스티커용지 PDF 다운로드",
        data=sticker_pdf,
        file_name=f"스티커용지_65칸_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # -----------------------------
    # (B) 수취인별 출력 (새벽/익일 분리 + 새벽 우선)
    # -----------------------------
    st.markdown("---")
    st.subheader("📄 수취인별 출력 - 새벽배송 / 익일배송 분리 (수취인명 길이에 맞춰 옆에 붙이기)")

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

        rows = []
        for (prod, var, sr), q in od.items():
            rows.append({"제품명": prod, "구분": var, "수량": q, "합산규칙": sr})

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

        parts = []
        for (pname, v), q2 in od2.items():
            parts.append(f"{pname}/{v} {fmt_qty(q2)}")

        recv_name = str(g["수취인명"].iloc[0]).strip()
        items_line = ", ".join(parts)
        return recv_name, items_line

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
