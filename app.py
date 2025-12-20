# app.py
import json
import re
from pathlib import Path

import pandas as pd
import streamlit as st

# =========================
# 저장 경로
# =========================
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
MAPPING_PATH = DATA_DIR / "name_mappings.json"


# =========================
# 매칭 규칙 로드/세이브
# =========================
def default_rules():
    # ✅ 예시를 "포함" 규칙으로 넣어두면
    #    1kg/500g/250g 같은 옵션이 달라도 한 번에 묶입니다.
    return [
        {
            "enabled": True,
            "priority": 10,
            "match_type": "contains",
            "pattern": "와일드루꼴라",
            "display_name": "와일드",
            "note": "채소팜 와일드루꼴라 1kg/500g/250g ... 모두 와일드로",
        },
        {
            "enabled": True,
            "priority": 20,
            "match_type": "contains",
            "pattern": "라디치오",
            "display_name": "라디치오",
            "note": "채소팜 라디치오 1통 이탈리안치커리 (약350g)",
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

    # 파일이 깨졌거나 형식이 이상하면 기본값 복구
    rules = default_rules()
    save_rules(rules)
    return rules


def save_rules(rules: list[dict]) -> None:
    MAPPING_PATH.write_text(
        json.dumps(rules, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# =========================
# 매칭 로직
# =========================
def normalize_text(s: str) -> str:
    return (s or "").strip()


def apply_mapping(actual_name: str, rules: list[dict]) -> str:
    actual = normalize_text(actual_name)
    if not actual:
        return ""

    # priority 오름차순(작을수록 먼저 적용)
    def prio(x):
        try:
            return int(x.get("priority", 9999))
        except Exception:
            return 9999

    for r in sorted(rules, key=prio):
        if not r.get("enabled", True):
            continue

        match_type = (r.get("match_type") or "").strip()
        pattern = normalize_text(r.get("pattern", ""))
        display = normalize_text(r.get("display_name", ""))

        if not pattern or not display:
            continue

        if match_type == "exact":
            if actual == pattern:
                return display

        elif match_type == "contains":
            if pattern in actual:
                return display

        elif match_type == "regex":
            try:
                if re.search(pattern, actual):
                    return display
            except re.error:
                # 정규식이 깨져있으면 그냥 무시
                continue

    # 규칙에 안 걸리면 원본 그대로
    return actual


# =========================
# UI
# =========================
st.set_page_config(page_title="상품명 매칭 규칙 관리", page_icon="🧩", layout="wide")
st.title("🧩 상품명 매칭 규칙 관리")

rules = load_rules()

menu = st.sidebar.radio(
    "메뉴",
    ["📌 상품명 매칭 관리", "🧪 매칭 테스트"],
    index=0,
)

st.sidebar.markdown("---")
st.sidebar.caption("저장 위치: data/name_mappings.json")


if menu == "📌 상품명 매칭 관리":
    st.subheader("규칙을 직접 추가/수정해서 ‘표시될 상품명’을 통일합니다")

    st.markdown(
        """
- **포함(contains)**: 상품명에 특정 단어가 들어가면 매칭 (옵션이 달라도 한 번에 처리하기 좋아요)
- **정확히 일치(exact)**: 상품명이 완전히 같을 때만 매칭
- **정규식(regex)**: 패턴이 복잡할 때 사용 (예: `와일드루꼴라\\s+(?:250g|500g|1kg)` )
"""
    )

    df = pd.DataFrame(rules)
    # 컬럼 고정/정렬
    for col in ["enabled", "priority", "match_type", "pattern", "display_name", "note"]:
        if col not in df.columns:
            df[col] = None
    df = df[["enabled", "priority", "match_type", "pattern", "display_name", "note"]]

    edited = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "enabled": st.column_config.CheckboxColumn("사용", default=True),
            "priority": st.column_config.NumberColumn("우선순위", help="작을수록 먼저 적용", min_value=0, step=1),
            "match_type": st.column_config.SelectboxColumn(
                "매칭 방식",
                options=["contains", "exact", "regex"],
                help="contains=포함, exact=정확히 일치, regex=정규식",
            ),
            "pattern": st.column_config.TextColumn("실제 상품명(패턴)", width="large"),
            "display_name": st.column_config.TextColumn("표시될 상품명", width="medium"),
            "note": st.column_config.TextColumn("메모", width="large"),
        },
        hide_index=True,
        key="mapping_editor",
    )

    col1, col2, col3, col4 = st.columns([1, 1, 1, 2])

    with col1:
        if st.button("💾 저장", use_container_width=True):
            # 빈 행/필수값 누락 제거 + 타입 정리
            cleaned = []
            for _, row in edited.iterrows():
                pattern = normalize_text(row.get("pattern"))
                display = normalize_text(row.get("display_name"))
                if not pattern or not display:
                    continue

                match_type = normalize_text(row.get("match_type")) or "contains"
                if match_type not in {"contains", "exact", "regex"}:
                    match_type = "contains"

                try:
                    priority = int(row.get("priority", 9999))
                except Exception:
                    priority = 9999

                cleaned.append(
                    dict(
                        enabled=bool(row.get("enabled", True)),
                        priority=priority,
                        match_type=match_type,
                        pattern=pattern,
                        display_name=display,
                        note=normalize_text(row.get("note")),
                    )
                )

            save_rules(cleaned)
            st.success(f"저장 완료! (규칙 {len(cleaned)}개)")

    with col2:
        if st.button("♻️ 기본 예시로 초기화", use_container_width=True):
            rules0 = default_rules()
            save_rules(rules0)
            st.success("기본 예시 규칙으로 초기화했어요. 새로고침(F5)하면 반영됩니다.")

    with col3:
        # 내보내기(다운로드)
        export_bytes = json.dumps(load_rules(), ensure_ascii=False, indent=2).encode("utf-8")
        st.download_button(
            "⬇️ 규칙 내보내기(JSON)",
            data=export_bytes,
            file_name="name_mappings.json",
            mime="application/json",
            use_container_width=True,
        )

    with col4:
        # 가져오기(업로드)
        up = st.file_uploader("규칙 가져오기(JSON)", type=["json"], label_visibility="collapsed")
        if up is not None:
            try:
                imported = json.loads(up.getvalue().decode("utf-8"))
                if not isinstance(imported, list):
                    raise ValueError("리스트 형식(JSON 배열)이 아닙니다.")
                save_rules(imported)
                st.success("가져오기 완료! 새로고침(F5)하면 반영됩니다.")
            except Exception as e:
                st.error(f"가져오기 실패: {e}")

    st.markdown("---")
    st.caption("팁) ‘와일드루꼴라’처럼 **핵심 단어만 ‘포함’ 규칙으로** 넣으면 1kg/500g/250g 옵션이 달라도 자동으로 같은 표시명으로 묶입니다.")


elif menu == "🧪 매칭 테스트":
    st.subheader("실제 상품명을 붙여 넣고 매칭 결과를 바로 확인하세요")

    sample = """채소팜 와일드루꼴라 1kg 베이비루꼴라
채소팜 와일드루꼴라 500g 베이비루꼴라
채소팜 와일드루꼴라 250g 베이비루꼴라
채소팜 라디치오 1통 이탈리안치커리 (약350g)
"""
    text = st.text_area("실제 상품명 (줄바꿈으로 여러 개)", value=sample, height=180)

    lines = [l.strip() for l in text.splitlines() if l.strip()]
    out = []
    for a in lines:
        out.append({"실제 상품명": a, "표시될 상품명": apply_mapping(a, load_rules())})

    st.dataframe(pd.DataFrame(out), use_container_width=True, height=520)
