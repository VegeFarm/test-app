try:
    import msoffcrypto
except ModuleNotFoundError:
    msoffcrypto = None

def render_invoice_register_page():
    import io
    import re
    from decimal import Decimal, InvalidOperation
    from typing import Optional, Tuple, Dict

    import pandas as pd
    import streamlit as st

    INV_FIXED_PASSWORD = "0000"
    INV_ROMAN_MAP = str.maketrans({
        "Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5",
        "Ⅵ": "6", "Ⅶ": "7", "Ⅷ": "8", "Ⅸ": "9", "Ⅹ": "10",
        "ⅰ": "1", "ⅱ": "2", "ⅲ": "3", "ⅳ": "4", "ⅴ": "5",
        "ⅵ": "6", "ⅶ": "7", "ⅷ": "8", "ⅸ": "9", "ⅹ": "10",
    })

    INV_SMARTSTORE_REQUIRED = ("구매자명", "수취인명", "통합배송지", "상품주문번호")
    INV_TRACKING_REQUIRED = ("주문자", "수령자", "수령자주소(상세포함)", "운송장번호")

    def inv_norm_text(s) -> str:
        if s is None or (isinstance(s, float) and pd.isna(s)):
            return ""
        s = str(s).strip().translate(INV_ROMAN_MAP)
        s = re.sub(r"\s+", "", s)
        s = re.sub(r"[^0-9A-Za-z가-힣]", "", s)
        return s

    def inv_clean_header_text(s) -> str:
        if s is None:
            return ""
        try:
            if pd.isna(s):
                return ""
        except Exception:
            pass

        s = str(s)
        s = s.replace("\ufeff", "").replace("\u200b", "").replace("\xa0", " ")
        s = s.replace("\r", " ").replace("\n", " ").strip()
        s = re.sub(r"\s+", "", s)
        return s

    def inv_to_plain_number_str(x) -> str:
        if x is None:
            return ""
        try:
            if isinstance(x, float) and pd.isna(x):
                return ""
        except Exception:
            pass

        s = str(x).strip()
        if s == "" or s.lower() == "nan":
            return ""

        s = s.replace(",", "")
        if re.fullmatch(r"-?\d+\.0+", s):
            return s.split(".")[0]

        try:
            d = Decimal(s)
            if d == d.to_integral():
                return format(d.to_integral(), "f")
            return format(d, "f").rstrip("0").rstrip(".")
        except (InvalidOperation, ValueError):
            return s

    def inv_to_plain_tracking_str(x) -> str:
        if x is None:
            return ""
        try:
            if isinstance(x, float) and pd.isna(x):
                return ""
        except Exception:
            pass

        s = str(x).strip()
        if s == "" or s.lower() == "nan":
            return ""

        if "-" in s:
            return s
        return inv_to_plain_number_str(s)

    def inv_decrypt_office_excel(file_bytes: bytes, password: str) -> io.BytesIO:
        if msoffcrypto is None:
            raise ModuleNotFoundError("msoffcrypto not installed")
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted

    def inv_find_header_row(df: pd.DataFrame, must_have: Tuple[str, ...], max_scan: int = 50) -> int:
        required = [inv_clean_header_text(x) for x in must_have]
        scan = min(max_scan, len(df))

        best_idx = -1
        best_score = -1

        for i in range(scan):
            row_values = [inv_clean_header_text(v) for v in df.iloc[i].tolist()]
            row_values = [v for v in row_values if v]
            if not row_values:
                continue

            exact_set = set(row_values)
            exact_score = sum(1 for col in required if col in exact_set)
            contains_score = sum(1 for col in required if any(col in cell for cell in row_values))

            if exact_score == len(required):
                return i

            score = (exact_score * 10) + contains_score
            if score > best_score:
                best_score = score
                best_idx = i

        return -1 if best_score <= 0 else best_idx

    def inv_normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        cleaned = []
        used = {}
        for col in df.columns:
            base = inv_clean_header_text(col)
            if not base:
                base = "빈컬럼"
            if base in used:
                used[base] += 1
                base = f"{base}_{used[base]}"
            else:
                used[base] = 0
            cleaned.append(base)
        df.columns = cleaned
        return df

    def inv_read_excel_with_flexible_header(
        excel_source,
        required_columns: Tuple[str, ...],
        password: Optional[str] = None,
        max_scan: int = 50,
    ) -> Tuple[pd.DataFrame, int]:
        if password is not None:
            raw_source = inv_decrypt_office_excel(excel_source.read(), password)
        else:
            raw_source = excel_source

        raw_df = pd.read_excel(raw_source, header=None, dtype=object)
        header_idx = inv_find_header_row(raw_df, must_have=required_columns, max_scan=max_scan)
        if header_idx < 0:
            raise ValueError(f"컬럼명 행을 찾지 못했습니다. 필요한 컬럼: {', '.join(required_columns)}")

        header = [inv_clean_header_text(v) for v in raw_df.iloc[header_idx].tolist()]
        df = raw_df.iloc[header_idx + 1 :].copy()
        df.columns = header
        df = inv_normalize_columns(df).reset_index(drop=True)

        repeated_header_mask = pd.Series(False, index=df.index)
        for col in required_columns:
            if col in df.columns:
                repeated_header_mask = repeated_header_mask | (df[col].map(inv_clean_header_text) == col)
        if not df.empty:
            df = df.loc[~repeated_header_mask].reset_index(drop=True)

        return df, header_idx

    def inv_choose_tracking(series: pd.Series) -> Optional[str]:
        s = series.dropna().astype(str)
        if s.empty:
            return None
        vc = s.value_counts()
        top = vc.max()
        candidates = vc[vc == top].index.tolist()
        if len(candidates) == 1:
            return candidates[0]
        for v in s:
            if v in candidates:
                return v
        return candidates[0]

    def inv_build_output(df1: pd.DataFrame, df2: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        col_buyer = "구매자명"
        col_recv = "수취인명"
        col_addr = "통합배송지"
        col_po = "상품주문번호"

        col2_buyer = "주문자"
        col2_recv = "수령자"
        col2_addr = "수령자주소(상세포함)"
        col2_track = "운송장번호"

        df1 = df1.copy()
        df2 = df2.copy()

        df1["__key"] = (
            df1[col_buyer].map(inv_norm_text)
            + "|"
            + df1[col_recv].map(inv_norm_text)
            + "|"
            + df1[col_addr].map(inv_norm_text)
        )
        df2["__key"] = (
            df2[col2_buyer].map(inv_norm_text)
            + "|"
            + df2[col2_recv].map(inv_norm_text)
            + "|"
            + df2[col2_addr].map(inv_norm_text)
        )

        map_track: Dict[str, Optional[str]] = df2.groupby("__key")[col2_track].apply(inv_choose_tracking).to_dict()
        df1["송장번호"] = df1["__key"].map(map_track)

        dup_info = (
            df2.groupby("__key")[col2_track]
            .nunique(dropna=True)
            .reset_index(name="운송장번호_종류수")
            .query("운송장번호_종류수 > 1")
            .sort_values("운송장번호_종류수", ascending=False)
        )

        df1["_상품주문번호_plain"] = df1[col_po].apply(inv_to_plain_number_str)
        df1["_송장번호_plain"] = df1["송장번호"].apply(inv_to_plain_tracking_str)

        out = pd.DataFrame(
            {
                "상품주문번호": df1["_상품주문번호_plain"],
                "배송방법": ["택배,등기,소포"] * len(df1),
                "택배사": df1["_송장번호_plain"].apply(
                    lambda x: "컬리넥스트마일" if "-" in str(x) else ("롯데택배" if str(x).strip() else "")
                ),
                "송장번호": df1["_송장번호_plain"],
            }
        )
        return out, dup_info

    def inv_export_xls(out_df: pd.DataFrame) -> bytes:
        import xlwt

        wb = xlwt.Workbook()
        ws = wb.add_sheet("발송처리")

        header_style = xlwt.easyxf("font: bold on; align: horiz center, vert center;")
        center_style = xlwt.easyxf("align: horiz center, vert center;")
        left_style = xlwt.easyxf("align: horiz left, vert center;")

        col_widths = [24, 10, 16, 32]
        for c, w in enumerate(col_widths):
            ws.col(c).width = int(w * 256)

        for c, name in enumerate(out_df.columns):
            ws.write(0, c, name, header_style)

        for r, row in enumerate(out_df.itertuples(index=False), start=1):
            vals = list(row)
            for c, v in enumerate(vals):
                v_str = "" if v is None else str(v)
                if c in (0, 3):
                    ws.write(r, c, v_str, left_style)
                else:
                    ws.write(r, c, v_str, center_style)

        bio = io.BytesIO()
        wb.save(bio)
        return bio.getvalue()

    st.title("🚚 송장등록")
    st.caption("2번 코드의 송장일괄발송 기능을 1번 코드 안으로 이식한 페이지입니다.")

    st.markdown("- 1번 파일은 **비밀번호 0000 고정**으로 열어서 처리합니다.")
    st.markdown("- 스마트스토어 엑셀은 **1행에 안내문/메모가 있어도 자동으로 실제 헤더를 찾아 처리**합니다.")
    st.markdown("- 결과는 **xls** 형식으로 다운로드됩니다.")

    st.markdown(
        """
<style>
.upload-title { font-size: 20px; font-weight: 700; margin-bottom: 2px; }
.result-title { font-size: 22px; font-weight: 800; margin-top: 8px; }
</style>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="upload-title">1) 스마트스토어 엑셀(비번0000)</div>', unsafe_allow_html=True)
    f1 = st.file_uploader(
        label="스마트스토어 엑셀 업로드",
        type=["xlsx"],
        key="invoice_smartstore_file",
        label_visibility="collapsed",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="upload-title">2) 운송장/출고 엑셀</div>', unsafe_allow_html=True)
    f2 = st.file_uploader(
        label="운송장/출고 엑셀 업로드",
        type=["xlsx", "xls"],
        key="invoice_tracking_file",
        label_visibility="collapsed",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    run = st.button("자동 채우기", type="primary", disabled=(f1 is None or f2 is None), key="invoice_run_btn")

    if not run:
        return

    try:
        df1, smartstore_header_idx = inv_read_excel_with_flexible_header(
            f1,
            required_columns=INV_SMARTSTORE_REQUIRED,
            password=INV_FIXED_PASSWORD,
            max_scan=50,
        )
    except Exception as e:
        st.error("1번 파일의 실제 헤더 행을 찾지 못했습니다. 상단 안내문이 있어도 되지만, 필요한 컬럼은 있어야 합니다.")
        st.exception(e)
        return

    try:
        df2, tracking_header_idx = inv_read_excel_with_flexible_header(
            f2,
            required_columns=INV_TRACKING_REQUIRED,
            password=None,
            max_scan=30,
        )
    except Exception:
        try:
            f2.seek(0)
            df2 = pd.read_excel(f2, dtype=object)
            df2 = inv_normalize_columns(df2)
            tracking_header_idx = 0
        except Exception as e:
            st.error("2번 파일을 읽지 못했습니다.")
            st.exception(e)
            return

    need1 = set(INV_SMARTSTORE_REQUIRED)
    need2 = {"주문자", "수령자", "수령자주소(상세포함)", "운송장번호"}

    if not need1.issubset(set(df1.columns)):
        st.error(f"1번 파일에 필요한 컬럼이 없습니다: {sorted(list(need1 - set(df1.columns)))}")
        return
    if not need2.issubset(set(df2.columns)):
        st.error(f"2번 파일에 필요한 컬럼이 없습니다: {sorted(list(need2 - set(df2.columns)))}")
        return

    out_df, dup_info = inv_build_output(df1, df2)

    c_meta1, c_meta2 = st.columns(2)
    c_meta1.caption(f"스마트스토어 헤더 행: {smartstore_header_idx + 1}행")
    c_meta2.caption(f"운송장 파일 헤더 행: {tracking_header_idx + 1}행")

    with st.expander("미리보기 (상위 30건)", expanded=False):
        st.dataframe(out_df.head(30), use_container_width=True)

    miss = (out_df["송장번호"].isna() | (out_df["송장번호"].astype(str).str.strip() == "")).sum()
    st.write(f"총 {len(out_df)}건 / 송장번호 누락 {miss}건")

    if not dup_info.empty:
        with st.expander("⚠️ 같은 주문자/수령자/주소인데 운송장번호가 여러 개인 경우", expanded=False):
            st.dataframe(dup_info.head(50), use_container_width=True)

    st.markdown('<div class="result-title">3) 결과 다운로드</div>', unsafe_allow_html=True)

    try:
        xls_bytes = inv_export_xls(out_df)
        st.download_button(
            "✅ 일괄발송 엑셀 다운로드",
            data=xls_bytes,
            file_name="송장일괄발송.xls",
            mime="application/vnd.ms-excel",
            key="invoice_download_btn",
        )
    except Exception as e:
        st.error(f"결과 파일 생성 실패: {e}")


# =====================================================
# Page: 🧰 재고일괄변경 (2.py 기능 이식)
# =====================================================
