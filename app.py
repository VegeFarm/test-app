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
    SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, KeepInFrame, PageBreak
)
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
        # 예시(필요 시 켜서 사용)
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
            unit_size, unit_label = 1, "개"  # 비어있으면 1개로 가정(합산규칙 있을 때만)
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
    # 새벽+익일 둘 다면 -> 새벽
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
# PDF 2) 수취인명 - 주문상품 (2줄) / 20개 per page
# =====================================================
def build_recipient_pdf(entries: List[Dict[str, str]]) -> bytes:
    """
    entries: [{"name_line": "...", "items_line": "..."}]  # 2줄 구성
    A4 1페이지 20개, 각 2줄
    """
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
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    name_style = ParagraphStyle(
        "name",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=10,
        leading=12,
        spaceAfter=0,
    )
    item_style = ParagraphStyle(
        "item",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=9,
        leading=11,
        spaceAfter=0,
    )

    elems = []

    PER_PAGE = 20
    row_h = 13.0 * mm  # 2줄 기준 고정 높이(길면 shrink)

    def make_cell(name_line: str, items_line: str):
        p1 = Paragraph(name_line, name_style)
        p2 = Paragraph(items_line, item_style)
        kif = KeepInFrame(0, row_h - 2, [p1, p2], mode="shrink", hAlign="LEFT", vAlign="TOP")
        return kif

    # 페이지 단위로 테이블 생성
    for start in range(0, len(entries), PER_PAGE):
        chunk = entries[start:start + PER_PAGE]

        data = []
        for e in chunk:
            data.append([make_cell(e["name_line"], e["items_line"])])

        # 마지막 페이지가 20개 미만이면 빈 줄 채워서 20개 맞춤(요청대로)
        while len(data) < PER_PAGE:
            data.append([make_cell(" ", " ")])

        table = LongTable(
            data,
            colWidths=[A4[0] - (24 * mm)],
            rowHeights=[row_h] * PER_PAGE,
            repeatRows=0,
        )
        table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                ]
            )
        )

        elems.append(table)
        if start + PER_PAGE < len(entries):
            elems.append(PageBreak())

    doc.build(elems)
    return buf.getvalue()


# =====================================================
# Streamlit UI
# =====================================================
st.set_page_config(page_title="제품별 개수 & 수취인별 출력", page_icon="📄", layout="wide")
st.title("📄 제품별 개수 & 수취인별 출력")
st.caption('엑셀 업로드 → (상품명 매칭/합산규칙) → 제품별 집계 + 수취인별(20개/페이지) PDF (엑셀 비밀번호 "0000" 고정)')

menu = st.sidebar.radio("메뉴", ["🧩 상품명 매칭 규칙", "⬆️ 엑셀 업로드 & 결과"], index=1)
st.sidebar.markdown("---")
st.sidebar.caption("규칙 파일: data/name_mappings.json")


# -----------------------------
# 1) 규칙 관리
# -----------------------------
if menu == "🧩 상품명 매칭 규칙":
    st.subheader("실제 상품명 → 표시될 상품명 + 합산규칙(개/봉/통/팩)")

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

    c1, c2, c3 = st.columns([1, 1, 2])
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
        if st.button("♻️ 기본 예시로 초기화", use_container_width=True):
            save_rules(default_rules())
            st.success("기본 예시 규칙으로 초기화했습니다. (새로고침 시 반영)")

    with c3:
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
    st.subheader("엑셀 업로드 → 제품별 집계 + 수취인별(20개/페이지) 출력")

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

    # ✅ 원본 미리보기 기능 삭제(요청사항)

    # 필요한 컬럼
    col_name = find_col(raw_df, ["상품명", "상품", "제품명"])
    col_qty = find_col(raw_df, ["수량", "주문수량", "구매수량", "개수"])
    col_buyer = find_col(raw_df, ["구매자명", "구매자"])
    col_recv = find_col(raw_df, ["수취인명", "수령인", "받는사람"])
    col_addr = find_col(raw_df, ["통합배송지", "배송지", "주소"])
    col_opt = find_col(raw_df, ["옵션정보", "옵션", "선택옵션"])

    missing = [k for k, v in {
        "상품명": col_name, "수량": col_qty, "구매자명": col_buyer, "수취인명": col_recv, "통합배송지": col_addr, "옵션정보": col_opt
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
    # (A) 제품별 집계 (기존 결과)
    # -----------------------------
    base = work[(work["수량"].notna()) & (work["제품명"] != "")].copy()
    exploded = explode_sum_rule_rows(base[["제품명", "구분", "수량", "합산규칙"]])

    summary = (
        exploded.groupby(["제품명", "구분"], as_index=False)["수량"]
        .sum()
        .sort_values(["제품명", "구분"], kind="mergesort")
        .reset_index(drop=True)
    )

    def fmt_qty(x):
        try:
            x = float(x)
            return int(x) if x.is_integer() else x
        except Exception:
            return x

    summary["수량"] = summary["수량"].apply(fmt_qty)

    st.markdown("---")
    st.subheader("✅ 결과 (제품명 / 구분 / 수량)")
    st.dataframe(summary, use_container_width=True, height=520)

    # 제품별 PDF
    summary_pdf = build_summary_pdf(summary)
    st.download_button(
        "⬇️ 제품별 개수 PDF 다운로드",
        data=summary_pdf,
        file_name=f"제품별개수_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # -----------------------------
    # (B) 수취인별 출력 (새벽/익일 분리 + 새벽 우선)
    # -----------------------------
    st.markdown("---")
    st.subheader("📄 수취인별 출력 (A4 / 20개씩 / 2줄) - 새벽배송 / 익일배송 분리")

    # 배송 분류
    base2 = base.copy()
    base2["배송구분"] = base2["옵션정보"].apply(classify_delivery)

    # 같은 주문자 그룹 키: 구매자명 + 수취인명 + 통합배송지
    key_cols = ["구매자명", "수취인명", "통합배송지"]

    # 그룹별 배송구분 결정(새벽 우선)
    grp_deliv = (
        base2.groupby(key_cols)["배송구분"]
        .agg(lambda x: set(x))
        .apply(decide_group_delivery)
        .reset_index()
        .rename(columns={"배송구분": "그룹배송구분"})
    )

    base2 = base2.merge(grp_deliv, on=key_cols, how="left")

    # 그룹 단위로 아이템 문자열 만들기
    def build_items_for_group(g: pd.DataFrame) -> Tuple[str, List[Dict[str, str]]]:
        """
        return: (수취인명, items_list dicts)  # items_list는 최종 explode후 집계된 품목들
        """
        # 원본 순서 유지(엑셀 행 순서)
        g = g.sort_index()

        # (제품명, 구분, 합산규칙) 단위로 합산 (순서 보존)
        od = OrderedDict()
        for _, r in g.iterrows():
            prod = str(r["제품명"]).strip()
            var = str(r["구분"] or "").strip()
            qty = r["수량"]
            sr = _safe_int(r.get("합산규칙", None))

            if not prod:
                continue
            if var == "":
                # 구분이 없으면 표시용으로 "-" (슬래시 강제 때문에 빈칸 방지)
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
        # 합산규칙 행 분해(개/봉/통/팩만)
        rows_ex = explode_sum_rule_rows(rows_df[["제품명", "구분", "수량", "합산규칙"]]) if len(rows_df) else rows_df

        # 다시 (제품명, 구분)으로 합산 (표시 순서는 가능한 유지)
        od2 = OrderedDict()
        for _, r in rows_ex.iterrows():
            k2 = (str(r["제품명"]), str(r["구분"]))
            if k2 not in od2:
                od2[k2] = 0.0
            try:
                od2[k2] += float(r["수량"])
            except Exception:
                pass

        items = []
        for (p, v), q in od2.items():
            items.append({"제품명": p, "구분": v, "수량": fmt_qty(q)})

        recv_name = str(g["수취인명"].iloc[0]).strip()
        return recv_name, items

    # 그룹 엔트리 생성
    group_entries = []
    for _, g in base2.groupby(key_cols, sort=False):
        recv_name, items = build_items_for_group(g)
        # 주문상품 문자열: "상품/구분 수량, 상품/구분 수량"
        parts = [f"{it['제품명']}/{it['구분']} {it['수량']}" for it in items if it["제품명"]]
        items_line = ", ".join(parts)

        group_entries.append(
            {
                "그룹배송구분": str(g["그룹배송구분"].iloc[0]),
                "수취인명": recv_name,
                "name_line": f"{recv_name} -",
                "items_line": items_line if items_line else " ",
            }
        )

    # 새벽/익일로 분리 (새벽 우선 규칙은 그룹배송구분으로 이미 반영됨)
    dawn_entries = [e for e in group_entries if e["그룹배송구분"] == "새벽배송"]
    next_entries = [e for e in group_entries if e["그룹배송구분"] == "익일배송"]

    c1, c2 = st.columns(2)
    with c1:
        st.write(f"새벽배송: {len(dawn_entries)}명")
        dawn_pdf = build_recipient_pdf(dawn_entries)
        st.download_button(
            "⬇️ 새벽배송 수취인별 PDF (20개/페이지)",
            data=dawn_pdf,
            file_name=f"수취인별_새벽배송_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with c2:
        st.write(f"익일배송: {len(next_entries)}명")
        next_pdf = build_recipient_pdf(next_entries)
        st.download_button(
            "⬇️ 익일배송 수취인별 PDF (20개/페이지)",
            data=next_pdf,
            file_name=f"수취인별_익일배송_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with st.expander("⚠️ 미매칭/누락 행 보기 (규칙 추가용)", expanded=False):
        bad = work[(work["매칭성공"] == False) | (work["수량"].isna())].copy()
        st.dataframe(bad.head(300), use_container_width=True)
