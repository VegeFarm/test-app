import io
import json
import re
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple, List, Dict

import pandas as pd
import streamlit as st

# -----------------------------
# Optional import (Streamlit Cloud에서 requirements 누락 시 안내)
# -----------------------------
try:
    import msoffcrypto
except ModuleNotFoundError:
    msoffcrypto = None

# PDF
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# =====================================================
# 설정
# =====================================================
EXCEL_PASSWORD = "0000"

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
MAPPING_PATH = DATA_DIR / "name_mappings.json"


# =====================================================
# 구분(단위) 추출
# =====================================================
UNIT_PATTERNS = [
    r"\d+(?:\.\d+)?kg\s*~\s*\d+(?:\.\d+)?kg",  # 1.8kg~2kg
    r"\d+(?:\.\d+)?kg",                        # 1kg, 1.5kg
    r"(?:약\s*)?\d+(?:\.\d+)?g",               # 500g, 약350g
    r"\d+개", r"\d+통", r"\d+단", r"\d+봉", r"\d+팩",
]
UNIT_RE = re.compile(r"(" + "|".join(UNIT_PATTERNS) + r")")


def extract_variant(name: str) -> str:
    """상품명에서 구분(단위)을 하나 추출"""
    s = (name or "").strip()
    m = UNIT_RE.search(s)
    if not m:
        return ""
    u = m.group(0)
    u = re.sub(r"\s+", "", u)  # 공백 제거
    u = u.replace("약", "")    # 약350g -> 350g
    if "~" in u:               # 1.8kg~2kg -> 2kg
        u = u.split("~", 1)[1]
    return u


# =====================================================
# 매칭 규칙 로드/세이브
# =====================================================
def default_rules() -> List[Dict]:
    # 합산규칙 예시: 오렌지를 5개로 묶기
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
            "sum_rule": 5,  # ✅ 5개 이하 합산/묶음
            "note": "오렌지 합산규칙(5) 예시",
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
    MAPPING_PATH.write_text(
        json.dumps(rules, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def normalize_text(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _safe_int(v) -> Optional[int]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        x = int(v)
        return x
    except Exception:
        return None


def apply_mapping(actual_name: str, rules: List[Dict]) -> Tuple[str, bool, Optional[int]]:
    """
    실제 상품명 -> 표시될 상품명
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
            # 합산규칙은 2 이상일 때만 의미
            if sum_rule is not None and sum_rule < 2:
                sum_rule = None
            return display, True, sum_rule

    # 미매칭 fallback: 브랜드/괄호 제거 후 앞부분
    s = re.sub(r"^\s*채소팜\s*", "", actual)
    s = re.sub(r"\([^)]*\)", "", s).strip()
    s = re.sub(r"\s+", " ", s).strip()
    m = UNIT_RE.search(s)
    if m:
        s = s[: m.start()].strip()
    fallback = s.split(" ")[0] if s else actual
    return fallback, False, None


# =====================================================
# 엑셀(암호 0000) 읽기
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

    # 1) 정확 일치
    for k in keywords:
        if k in cols:
            return k

    # 2) 포함 검색
    for c in cols:
        cs = str(c)
        for k in keywords:
            if k in cs:
                return c
    return None


# =====================================================
# 합산규칙 적용 (개수 묶음)
# =====================================================
def parse_count_variant(variant: str) -> Optional[int]:
    """'3개' -> 3, 아니면 None"""
    m = re.fullmatch(r"(\d+)개", (variant or "").strip())
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def explode_sum_rule_rows(df_rows: pd.DataFrame) -> pd.DataFrame:
    """
    df_rows columns required:
      제품명, 구분, 수량, 합산규칙

    합산규칙(N)이 있는 경우:
      - 구분이 '개' 기반이면 (또는 구분이 비어있으면 1개로 가정)
      - 총 개수(total)를 N개 묶음 + 나머지(<=N) 묶음으로 분해
      - 예: total=8, N=5 => 5개 x1, 3개 x1
    """
    out = []

    for _, r in df_rows.iterrows():
        product = r["제품명"]
        variant = r["구분"]
        qty = r["수량"]
        rule = r.get("합산규칙", None)

        rule_n = _safe_int(rule)
        if rule_n is None or rule_n < 2:
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        # 개수 기반인지 판단
        if (variant or "").strip() == "":
            unit_size = 1  # 구분이 없으면 1개로 가정(합산규칙이 있을 때만)
            is_count = True
        else:
            unit_size = parse_count_variant(variant)
            is_count = unit_size is not None

        if not is_count:
            # g/kg/통/봉 같은 건 합산규칙 미적용
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        try:
            total_items = int(round(float(qty))) * int(unit_size)
        except Exception:
            # 수량이 이상하면 원본 유지
            out.append({"제품명": product, "구분": variant, "수량": qty})
            continue

        if total_items <= 0:
            continue

        full = total_items // rule_n
        rem = total_items % rule_n

        if full > 0:
            out.append({"제품명": product, "구분": f"{rule_n}개", "수량": full})
        if rem > 0:
            # ✅ 핵심: 5개 이하(=rem)는 "rem개 1"로 합산되어 표현
            out.append({"제품명": product, "구분": f"{rem}개", "수량": 1})

    return pd.DataFrame(out)


# =====================================================
# PDF 생성
# =====================================================
def build_pdf(summary_df: pd.DataFrame) -> bytes:
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
    for _, r in summary_df.iterrows():
        data.append([str(r["제품명"]), str(r["구분"]), str(r["수량"])])

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
# Streamlit UI
# =====================================================
st.set_page_config(page_title="제품별 개수 생성기", page_icon="📄", layout="wide")
st.title("📄 제품별 개수 생성기")
st.caption('엑셀 업로드 → (상품명 매칭 + 합산규칙) → 제품명/구분/수량 집계 → PDF 다운로드 (엑셀 비밀번호 "0000" 고정)')

menu = st.sidebar.radio("메뉴", ["🧩 상품명 매칭 규칙", "⬆️ 엑셀 업로드 & 결과"], index=1)
st.sidebar.markdown("---")
st.sidebar.caption("규칙 파일: data/name_mappings.json")


# -----------------------------
# 1) 규칙 관리
# -----------------------------
if menu == "🧩 상품명 매칭 규칙":
    st.subheader("실제 상품명 → 표시될 상품명 + 합산규칙(개수 묶음)")

    st.markdown(
        """
- **합산규칙(N)**: 해당 상품이 **개수(…개)** 기반일 때, 총 개수를 **N개 묶음 + 나머지(<=N)개 묶음**으로 표현합니다.
  - 예: 합산규칙=5, 총 8개 → `5개 수량 1` + `3개 수량 1`
  - 예: 합산규칙=5, 총 3개 → `3개 수량 1`
- g/kg/통/봉/팩 같은 단위에는 합산규칙이 적용되지 않습니다.
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
                help="개수 상품을 N개 묶음으로 표현. (비우면 미적용)",
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

    st.info('예) 오렌지를 5개 묶음으로 표현하려면 `pattern=오렌지`, `display_name=오렌지`, `합산규칙=5`로 설정하세요.')


# -----------------------------
# 2) 엑셀 업로드 & 결과
# -----------------------------
else:
    st.subheader("엑셀 업로드 → 제품명/구분/수량 집계 (합산규칙 포함)")

    if msoffcrypto is None:
        st.error("msoffcrypto가 설치되지 않았습니다. requirements.txt에 'msoffcrypto-tool'을 추가하고 재배포해 주세요.")
        st.stop()

    uploaded = st.file_uploader("비밀번호(0000) 엑셀 업로드 (.xlsx)", type=["xlsx"])
    if uploaded is None:
        st.info("엑셀을 업로드하면 결과 표와 PDF 다운로드가 나타납니다.")
        st.stop()

    # 복호화 + 로드
    try:
        decrypted = decrypt_excel(uploaded.getvalue(), password=EXCEL_PASSWORD)
        raw_df = pd.read_excel(decrypted, sheet_name=0, engine="openpyxl")
    except Exception as e:
        st.error('엑셀 읽기/복호화 실패: 비밀번호 "0000" 또는 파일 형식을 확인해 주세요.')
        st.exception(e)
        st.stop()

    # ✅ 원본 미리보기 삭제 (요청사항)

    # 컬럼 찾기
    name_col = find_col(raw_df, ["상품명", "상품", "제품명"])
    qty_col = find_col(raw_df, ["수량", "주문수량", "구매수량", "개수"])

    if not name_col or not qty_col:
        st.error("필수 컬럼을 찾지 못했습니다. (상품명/수량 계열 컬럼 필요)")
        st.write("현재 컬럼:", list(raw_df.columns))
        st.stop()

    rules = load_rules()

    work = raw_df[[name_col, qty_col]].copy()
    work.rename(columns={name_col: "상품명", qty_col: "수량"}, inplace=True)

    work["상품명"] = work["상품명"].astype(str)
    work["수량"] = pd.to_numeric(work["수량"], errors="coerce")

    # 구분 추출
    work["구분"] = work["상품명"].apply(extract_variant)

    # 매칭 적용 (제품명 + 합산규칙)
    mapped = work["상품명"].apply(lambda x: apply_mapping(x, rules))
    work["제품명"] = mapped.apply(lambda t: t[0])
    work["매칭성공"] = mapped.apply(lambda t: t[1])
    work["합산규칙"] = mapped.apply(lambda t: t[2])

    # 유효행
    base = work[(work["수량"].notna()) & (work["제품명"] != "")].copy()

    # ✅ 합산규칙 적용(행 분해)
    exploded = explode_sum_rule_rows(base[["제품명", "구분", "수량", "합산규칙"]])

    # 집계
    summary = (
        exploded.groupby(["제품명", "구분"], as_index=False)["수량"]
        .sum()
        .sort_values(["제품명", "구분"], kind="mergesort")
        .reset_index(drop=True)
    )

    # 수량 예쁘게
    def fmt_qty(x):
        try:
            x = float(x)
            return int(x) if x.is_integer() else x
        except Exception:
            return x

    summary["수량"] = summary["수량"].apply(fmt_qty)

    st.markdown("---")
    st.subheader("✅ 결과 (제품명 / 구분 / 수량)")
    st.dataframe(summary, use_container_width=True, height=600)

    with st.expander("⚠️ 미매칭/누락 행 보기 (규칙 추가용)", expanded=False):
        bad = work[(work["매칭성공"] == False) | (work["수량"].isna())].copy()
        st.dataframe(bad.head(300), use_container_width=True)

    # PDF 다운로드
    pdf_bytes = build_pdf(summary)
    filename = f"제품별개수_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"

    c1, c2 = st.columns([1, 1])
    with c1:
        st.download_button(
            "⬇️ PDF 다운로드",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "⬇️ 결과 CSV 다운로드",
            data=summary.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"제품별개수_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
