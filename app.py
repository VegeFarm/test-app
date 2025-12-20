import io
import json
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

# 엑셀(비번) 복호화
import msoffcrypto

# PDF 생성
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# =====================================================
# 저장 경로 (매칭 규칙)
# =====================================================
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
MAPPING_PATH = DATA_DIR / "name_mappings.json"


# =====================================================
# 단위(구분) 추출 정규식
# - 상품명 안의 g/kg/개/통/단/봉/팩 등을 찾아 "구분"으로 사용
# - (약350g) 같은 형태도 잡도록 "약" 허용
# =====================================================
UNIT_PATTERNS = [
    r"\d+(?:\.\d+)?kg\s*~\s*\d+(?:\.\d+)?kg",  # 1.8kg~2kg
    r"\d+(?:\.\d+)?kg",                        # 1kg, 1.5kg
    r"(?:약\s*)?\d+(?:\.\d+)?g",               # 500g, 약350g
    r"\d+개",
    r"\d+통",
    r"\d+단",
    r"\d+봉",
    r"\d+팩",
]
UNIT_RE = re.compile(r"(" + "|".join(UNIT_PATTERNS) + r")")


def extract_variant(name: str) -> str:
    """상품명에서 구분(단위)을 하나 추출"""
    s = (name or "").strip()
    m = UNIT_RE.search(s)
    if not m:
        return ""

    u = m.group(0)
    u = re.sub(r"\s+", "", u)       # 공백 제거
    u = u.replace("약", "")          # '약350g' -> '350g'

    # 범위: 1.8kg~2kg -> 오른쪽(2kg) 사용
    if "~" in u:
        u = u.split("~", 1)[1]

    return u


# =====================================================
# 매칭 규칙 로드/세이브
# =====================================================
def default_rules():
    # 요청 예시를 기본값으로 넣어둠
    return [
        {
            "enabled": True,
            "priority": 10,
            "match_type": "contains",
            "pattern": "와일드루꼴라",
            "display_name": "와일드",
            "note": '예) "채소팜 와일드루꼴라 1kg 베이비루꼴라" -> 와일드',
        },
        {
            "enabled": True,
            "priority": 20,
            "match_type": "contains",
            "pattern": "라디치오",
            "display_name": "라디치오",
            "note": '예) "채소팜 라디치오 1통 이탈리안치커리 (약350g)" -> 라디치오',
        },
    ]


def load_rules() -> list[dict]:
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


def save_rules(rules: list[dict]) -> None:
    MAPPING_PATH.write_text(
        json.dumps(rules, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def normalize_text(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def apply_mapping(actual_name: str, rules: list[dict]) -> tuple[str, bool]:
    """
    실제 상품명 -> 표시될 상품명
    반환: (표시명, 매칭성공여부)
    """
    actual = normalize_text(actual_name)
    if not actual:
        return "", False

    def prio(x):
        try:
            return int(x.get("priority", 9999))
        except Exception:
            return 9999

    for r in sorted(rules, key=prio):
        if not r.get("enabled", True):
            continue

        match_type = (r.get("match_type") or "contains").strip()
        pattern = normalize_text(r.get("pattern", ""))
        display = normalize_text(r.get("display_name", ""))

        if not pattern or not display:
            continue

        if match_type == "exact":
            if actual == pattern:
                return display, True

        elif match_type == "contains":
            if pattern in actual:
                return display, True

        elif match_type == "regex":
            try:
                if re.search(pattern, actual):
                    return display, True
            except re.error:
                continue

    # 미매칭이면 원본에서 브랜드/괄호 제거 후 첫 토큰 정도로 fallback (너무 지저분해지는 것 방지)
    s = re.sub(r"^\s*채소팜\s*", "", actual)
    s = re.sub(r"\([^)]*\)", "", s).strip()
    s = re.sub(r"\s+", " ", s).strip()
    # 단위 이전까지만 잘라서 첫 토큰 반환
    m = UNIT_RE.search(s)
    if m:
        s = s[: m.start()].strip()
    fallback = s.split(" ")[0] if s else actual
    return fallback, False


# =====================================================
# 엑셀(비번 0000) 로드
# =====================================================
EXCEL_PASSWORD = "0000"


def decrypt_excel(uploaded_bytes: bytes, password: str = EXCEL_PASSWORD) -> io.BytesIO:
    decrypted = io.BytesIO()
    office = msoffcrypto.OfficeFile(io.BytesIO(uploaded_bytes))
    office.load_key(password=password)
    office.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def find_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """
    컬럼명이 정확히 '상품명','수량'이 아닐 수도 있어서
    포함 키워드 기반으로 탐색
    """
    cols = list(df.columns)
    # 1) 정확히 일치 우선
    for k in keywords:
        if k in cols:
            return k
    # 2) 포함 탐색
    for c in cols:
        cs = str(c)
        for k in keywords:
            if k in cs:
                return c
    return None


# =====================================================
# PDF 생성 (예시 PDF 형태)
# =====================================================
def build_pdf(summary_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()

    # CID 폰트로 한글 깨짐 최소화
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
st.caption('엑셀 업로드 → (상품명 매칭 규칙 적용) → 제품명/구분/수량 집계 → PDF 다운로드 (엑셀 비밀번호 "0000" 고정)')

menu = st.sidebar.radio("메뉴", ["🧩 상품명 매칭 규칙", "⬆️ 엑셀 업로드 & 결과"], index=1)
st.sidebar.markdown("---")
st.sidebar.caption("규칙 파일: data/name_mappings.json")


# -----------------------------
# 1) 매칭 규칙 관리 페이지
# -----------------------------
if menu == "🧩 상품명 매칭 규칙":
    st.subheader("실제 상품명 → 표시될 상품명 규칙을 직접 만듭니다")
    st.markdown(
        """
- **contains(포함)**: 가장 추천 (1kg/500g/250g 옵션이 달라도 같은 상품으로 묶기 쉬움)
- **exact(정확히 일치)**: 상품명이 완전히 같은 경우만
- **regex(정규식)**: 패턴이 복잡할 때
"""
    )

    rules = load_rules()
    df = pd.DataFrame(rules)
    for col in ["enabled", "priority", "match_type", "pattern", "display_name", "note"]:
        if col not in df.columns:
            df[col] = None
    df = df[["enabled", "priority", "match_type", "pattern", "display_name", "note"]]

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
            "note": st.column_config.TextColumn("메모", width="large"),
        },
        key="mapping_editor",
    )

    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])

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

                cleaned.append(
                    dict(
                        enabled=bool(row.get("enabled", True)),
                        priority=pr,
                        match_type=mt,
                        pattern=pattern,
                        display_name=display,
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
            "⬇️ 규칙 내보내기",
            data=export_bytes,
            file_name="name_mappings.json",
            mime="application/json",
            use_container_width=True,
        )

    with c4:
        up = st.file_uploader("규칙 가져오기(JSON)", type=["json"], label_visibility="collapsed")
        if up is not None:
            try:
                imported = json.loads(up.getvalue().decode("utf-8"))
                if not isinstance(imported, list):
                    raise ValueError("JSON 배열(list) 형식이어야 합니다.")
                save_rules(imported)
                st.success("가져오기 완료! (새로고침 시 반영)")
            except Exception as e:
                st.error(f"가져오기 실패: {e}")

    st.markdown("---")
    st.info(
        '예) "채소팜 와일드루꼴라 1kg 베이비루꼴라" / "채소팜 와일드루꼴라 500g ..." 를 한 번에 묶으려면\n'
        'pattern="와일드루꼴라", match_type="contains", display_name="와일드" 로 두면 됩니다.'
    )


# -----------------------------
# 2) 엑셀 업로드 & 결과 페이지
# -----------------------------
else:
    st.subheader("엑셀 업로드 → 제품명/구분/수량 집계")

    uploaded = st.file_uploader("비밀번호(0000) 엑셀 업로드 (.xlsx)", type=["xlsx"])

    if uploaded is None:
        st.info("엑셀을 업로드하면 결과 표와 PDF 다운로드가 나타납니다.")
        st.stop()

    # 엑셀 복호화
    try:
        decrypted = decrypt_excel(uploaded.getvalue(), password=EXCEL_PASSWORD)
    except Exception as e:
        st.error('엑셀 복호화 실패: 비밀번호가 "0000"인지, 파일이 암호화된 xlsx인지 확인해 주세요.')
        st.exception(e)
        st.stop()

    # 엑셀 읽기
    try:
        raw_df = pd.read_excel(decrypted, sheet_name=0, engine="openpyxl")
    except Exception as e:
        st.error("엑셀 읽기 실패: 시트 구조/형식이 예상과 다를 수 있어요.")
        st.exception(e)
        st.stop()

    st.write("원본 일부 미리보기")
    st.dataframe(raw_df.head(20), use_container_width=True)

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

    work["수량"] = pd.to_numeric(work["수량"], errors="coerce")
    work["구분"] = work["상품명"].astype(str).apply(extract_variant)

    mapped = work["상품명"].astype(str).apply(lambda x: apply_mapping(x, rules))
    work["제품명"] = mapped.apply(lambda t: t[0])
    work["매칭성공"] = mapped.apply(lambda t: t[1])

    # 유효 행만
    ok = work[(work["수량"].notna()) & (work["구분"] != "") & (work["제품명"] != "")].copy()

    # 집계
    summary = (
        ok.groupby(["제품명", "구분"], as_index=False)["수량"]
        .sum()
        .sort_values(["제품명", "구분"], kind="mergesort")
        .reset_index(drop=True)
    )

    # 수량 예쁘게(정수면 정수로)
    def fmt_qty(x):
        try:
            x = float(x)
            return int(x) if x.is_integer() else x
        except Exception:
            return x

    summary["수량"] = summary["수량"].apply(fmt_qty)

    st.markdown("---")
    st.subheader("✅ 결과 (제품명 / 구분 / 수량)")
    st.dataframe(summary, use_container_width=True, height=560)

    # 미매칭 목록
    with st.expander("⚠️ 미매칭/누락 행 보기 (규칙 추가용)", expanded=False):
        # 미매칭(표시명이 fallback이거나 매칭성공 False) + 구분/수량 누락도 같이 보여줌
        bad = work[(work["매칭성공"] == False) | (work["구분"] == "") | (work["수량"].isna())].copy()
        bad = bad.sort_values(["매칭성공", "구분"], ascending=[True, True])
        st.write("아래를 보고 ‘상품명 매칭 규칙’에 pattern을 추가하면 결과가 예시 PDF처럼 깔끔해집니다.")
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
