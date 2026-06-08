try:
    import msoffcrypto
except ModuleNotFoundError:
    msoffcrypto = None

def render_invoice_register_page():
    import io
    import json
    import re
    from decimal import Decimal, InvalidOperation
    from pathlib import Path
    from typing import Optional, Tuple, Dict

    import pandas as pd
    import streamlit as st

    INV_FIXED_PASSWORD = "0000"
    INV_TRACKING_PASSWORD_DEFAULT = "CU000640-master"
    INV_COURIER_NEXTMILE_DEFAULT = "컬리넥스트마일"
    INV_COURIER_LOTTE_DEFAULT = "롯데택배"

    try:
        from config import DATA_DIR
        INV_SETTINGS_PATH = DATA_DIR / "invoice_settings.json"
    except Exception:
        INV_SETTINGS_PATH = Path("invoice_settings.json")

    INV_ROMAN_MAP = str.maketrans({
        "Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5",
        "Ⅵ": "6", "Ⅶ": "7", "Ⅷ": "8", "Ⅸ": "9", "Ⅹ": "10",
        "ⅰ": "1", "ⅱ": "2", "ⅲ": "3", "ⅳ": "4", "ⅴ": "5",
        "ⅵ": "6", "ⅶ": "7", "ⅷ": "8", "ⅸ": "9", "ⅹ": "10",
    })

    INV_SMARTSTORE_REQUIRED = ("구매자명", "수취인명", "통합배송지", "상품주문번호")
    # 새 운송장/출고 엑셀은 "주문자 이름", "수령자 이름", "운송장 번호"처럼
    # 띄어쓰기/설명 문구가 달라질 수 있어 헤더 탐색은 핵심 컬럼만 사용합니다.
    INV_TRACKING_REQUIRED = ("운송장번호", "수령자주소(상세포함)")

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

    def inv_load_settings() -> dict:
        default = {
            "tracking_password": INV_TRACKING_PASSWORD_DEFAULT,
            "courier_nextmile": INV_COURIER_NEXTMILE_DEFAULT,
            "courier_lotte": INV_COURIER_LOTTE_DEFAULT,
        }
        try:
            if INV_SETTINGS_PATH.exists():
                data = json.loads(INV_SETTINGS_PATH.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    default.update({k: str(v) for k, v in data.items() if v is not None})
        except Exception:
            pass
        return default

    def inv_save_settings(data: dict) -> None:
        try:
            INV_SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
            INV_SETTINGS_PATH.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            # 설정 저장 실패가 송장 처리 자체를 막지 않도록 합니다.
            pass

    def inv_normalize_courier_output_setting(value: str, default_value: str) -> str:
        """왼쪽 옵션값을 네이버 송장일괄발송 결과물에 들어갈 택배사명으로 보정합니다.

        운송장/출고 엑셀에는 배송사가 '넥스트마일', '롯데'로 내려오지만,
        네이버 송장일괄발송 엑셀에는 보통 '컬리넥스트마일', '롯데택배'처럼
        등록 가능한 택배사명으로 넣어야 합니다.
        사용자가 옵션에 '롯데'만 저장해 둔 경우에도 결과물에는 '롯데택배'가 나오게 합니다.
        """
        s = "" if value is None else str(value).strip()
        if not s:
            return default_value

        normalized = inv_clean_header_text(s).lower()
        if normalized in {"롯데", "lotte"}:
            return "롯데택배"
        if normalized in {"넥스트마일", "nextmile", "kurlynextmile"}:
            return "컬리넥스트마일"
        return s

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

    def inv_get_column(df: pd.DataFrame, candidates, required: bool = True) -> Optional[str]:
        """정규화된 컬럼명에서 후보 컬럼을 찾습니다.

        예: 새 엑셀의 "주문자 이름"은 inv_clean_header_text 후 "주문자이름"이 됩니다.
        """
        cols = list(df.columns)
        cand_clean = [inv_clean_header_text(c) for c in candidates]

        for cand in cand_clean:
            if cand in cols:
                return cand

        for cand in cand_clean:
            for col in cols:
                if cand and cand in col:
                    return col

        if required:
            raise ValueError(f"필요한 컬럼을 찾지 못했습니다: {', '.join(candidates)}")
        return None

    def inv_first_nonempty(series: pd.Series) -> str:
        for v in series.tolist():
            if v is None:
                continue
            try:
                if isinstance(v, float) and pd.isna(v):
                    continue
            except Exception:
                pass
            s = str(v).strip()
            if s and s.lower() != "nan":
                return s
        return ""

    def inv_resolve_courier(raw_courier, tracking_no: str, courier_nextmile: str, courier_lotte: str) -> str:
        tracking_no = "" if tracking_no is None else str(tracking_no).strip()
        if not tracking_no:
            return ""

        raw = "" if raw_courier is None else str(raw_courier).strip()
        raw_norm = inv_clean_header_text(raw).lower()

        if "넥스트" in raw_norm or "nextmile" in raw_norm:
            return courier_nextmile
        if "롯데" in raw_norm or "lotte" in raw_norm:
            return courier_lotte
        if raw:
            return raw

        # 구형 운송장 파일처럼 배송사 컬럼이 없을 때는 기존 방식 유지
        return courier_nextmile if "-" in tracking_no else courier_lotte

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

    def inv_build_output(
        df1: pd.DataFrame,
        df2: pd.DataFrame,
        courier_nextmile: str,
        courier_lotte: str,
    ) -> Tuple[pd.DataFrame, pd.DataFrame, int]:
        col_buyer = "구매자명"
        col_recv = "수취인명"
        col_addr = "통합배송지"
        col_po = "상품주문번호"

        # 운송장/출고 엑셀은 구형/신형 컬럼명을 모두 허용합니다.
        col2_buyer = inv_get_column(df2, ["주문자이름", "주문자명", "주문자"])
        col2_recv = inv_get_column(df2, ["수령자이름", "수취인이름", "수령자", "수취인명"])
        col2_addr = inv_get_column(df2, ["수령자주소(상세포함)", "수취인주소(상세포함)", "통합배송지"])
        col2_track = inv_get_column(df2, ["운송장번호", "송장번호"])
        col2_status = inv_get_column(df2, ["운송장상태", "출고상태", "상태"], required=False)
        col2_courier = inv_get_column(df2, ["배송사", "택배사", "운송사"], required=False)

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

        skipped_cancelled = 0
        cancelled_keys = set()
        if col2_status:
            cancel_mask = df2[col2_status].map(inv_clean_header_text) == "배송취소"
            skipped_cancelled = int(cancel_mask.sum())
            cancelled_keys = set(
                k for k in df2.loc[cancel_mask, "__key"].dropna().astype(str).tolist()
                if k and k != "||"
            )
            df2 = df2.loc[~cancel_mask].reset_index(drop=True)

        df2["__송장번호_plain"] = df2[col2_track].apply(inv_to_plain_tracking_str)
        df2 = df2.loc[df2["__송장번호_plain"].astype(str).str.strip() != ""].reset_index(drop=True)

        track_records: Dict[str, Dict[str, str]] = {}
        for key, group in df2.groupby("__key", dropna=False):
            tracking = inv_choose_tracking(group["__송장번호_plain"])
            courier_raw = ""
            if col2_courier:
                same_tracking = group[group["__송장번호_plain"].astype(str) == str(tracking)]
                courier_raw = inv_first_nonempty(same_tracking[col2_courier])
                if not courier_raw:
                    courier_raw = inv_first_nonempty(group[col2_courier])
            track_records[key] = {"tracking": tracking or "", "courier_raw": courier_raw}

        df1["송장번호"] = df1["__key"].map(lambda k: track_records.get(k, {}).get("tracking", ""))
        df1["__배송사_raw"] = df1["__key"].map(lambda k: track_records.get(k, {}).get("courier_raw", ""))

        # 운송장/출고 엑셀에서 배송 취소로 내려온 주문은
        # 네이버 송장일괄발송 결과물에서도 아예 제외합니다.
        # 단, 같은 주문자/수령자/주소에 배송 취소와 출고 진행중이 함께 있으면
        # 출고 진행중 운송장번호가 존재하므로 제외하지 않고 정상 출력합니다.
        active_keys = {
            k for k, v in track_records.items()
            if str(v.get("tracking", "")).strip()
        }
        cancelled_only_keys = cancelled_keys - active_keys
        result_excluded_cancelled = 0
        if cancelled_only_keys:
            exclude_cancelled_mask = df1["__key"].isin(cancelled_only_keys)
            result_excluded_cancelled = int(exclude_cancelled_mask.sum())
            df1 = df1.loc[~exclude_cancelled_mask].reset_index(drop=True)

        dup_info = (
            df2.groupby("__key")["__송장번호_plain"]
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
                "택배사": df1.apply(
                    lambda row: inv_resolve_courier(
                        row.get("__배송사_raw", ""),
                        row.get("_송장번호_plain", ""),
                        courier_nextmile,
                        courier_lotte,
                    ),
                    axis=1,
                ),
                "송장번호": df1["_송장번호_plain"],
            }
        )
        return out, dup_info, skipped_cancelled, result_excluded_cancelled

    def inv_read_tracking_file(excel_source, password: str) -> Tuple[pd.DataFrame, int]:
        errors = []
        password = (password or "").strip()

        if password:
            try:
                excel_source.seek(0)
                return inv_read_excel_with_flexible_header(
                    excel_source,
                    required_columns=INV_TRACKING_REQUIRED,
                    password=password,
                    max_scan=50,
                )
            except Exception as e:
                errors.append(f"비밀번호 적용 읽기 실패: {e}")

        try:
            excel_source.seek(0)
            return inv_read_excel_with_flexible_header(
                excel_source,
                required_columns=INV_TRACKING_REQUIRED,
                password=None,
                max_scan=50,
            )
        except Exception as e:
            errors.append(f"일반 읽기 실패: {e}")

        raise ValueError(" / ".join(errors))

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
    st.markdown("- 2번 운송장/출고 엑셀은 왼쪽 **운송장 비밀번호 설정**에 저장된 비밀번호로 자동 복호화합니다.")
    st.markdown("- 운송장 상태가 **배송 취소**인 행은 송장 매칭에서 자동 제외합니다.")
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

    inv_settings = inv_load_settings()
    with st.sidebar:
        st.markdown("---")
        with st.expander("송장등록 옵션", expanded=False):
            st.subheader("운송장 비밀번호 설정")
            tracking_password_input = st.text_input(
                "운송장/출고 엑셀 비밀번호",
                value=str(inv_settings.get("tracking_password", INV_TRACKING_PASSWORD_DEFAULT)),
                type="password",
                key="invoice_tracking_password_setting",
                help="이 비밀번호로 2번 운송장/출고 엑셀을 자동으로 엽니다.",
            )

            st.subheader("운송장 배송사 출력 설정")
            st.caption("운송장/출고 엑셀의 배송사 값이 넥스트마일 또는 롯데일 때, 결과 엑셀의 택배사명을 아래 값으로 바꿔 넣습니다.")
            courier_nextmile_input = st.text_input(
                "넥스트마일 → 결과 택배사",
                value=str(inv_settings.get("courier_nextmile", INV_COURIER_NEXTMILE_DEFAULT)),
                key="invoice_courier_nextmile_setting",
            )
            courier_lotte_input = st.text_input(
                "롯데 → 결과 택배사",
                value=str(inv_settings.get("courier_lotte", INV_COURIER_LOTTE_DEFAULT)),
                key="invoice_courier_lotte_setting",
            )

            if st.button("💾 송장등록 설정 저장", use_container_width=True, key="invoice_save_settings_btn"):
                inv_save_settings(
                    {
                        "tracking_password": tracking_password_input.strip(),
                        "courier_nextmile": inv_normalize_courier_output_setting(
                            courier_nextmile_input,
                            INV_COURIER_NEXTMILE_DEFAULT,
                        ),
                        "courier_lotte": inv_normalize_courier_output_setting(
                            courier_lotte_input,
                            INV_COURIER_LOTTE_DEFAULT,
                        ),
                    }
                )
                st.success("송장등록 설정 저장 완료")

    courier_nextmile_input = inv_normalize_courier_output_setting(
        courier_nextmile_input,
        INV_COURIER_NEXTMILE_DEFAULT,
    )
    courier_lotte_input = inv_normalize_courier_output_setting(
        courier_lotte_input,
        INV_COURIER_LOTTE_DEFAULT,
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
        df2, tracking_header_idx = inv_read_tracking_file(f2, tracking_password_input)
    except Exception as e:
        st.error("2번 운송장/출고 엑셀을 읽지 못했습니다. 왼쪽의 운송장 비밀번호 설정 값을 확인해 주세요.")
        st.exception(e)
        return

    need1 = set(INV_SMARTSTORE_REQUIRED)

    if not need1.issubset(set(df1.columns)):
        st.error(f"1번 파일에 필요한 컬럼이 없습니다: {sorted(list(need1 - set(df1.columns)))}")
        return

    try:
        out_df, dup_info, skipped_cancelled, result_excluded_cancelled = inv_build_output(
            df1,
            df2,
            courier_nextmile=courier_nextmile_input,
            courier_lotte=courier_lotte_input,
        )
    except Exception as e:
        st.error("2번 운송장/출고 엑셀의 필수 컬럼을 찾지 못했거나 송장 매칭 중 오류가 발생했습니다.")
        st.exception(e)
        return

    c_meta1, c_meta2, c_meta3 = st.columns(3)
    c_meta1.caption(f"스마트스토어 헤더 행: {smartstore_header_idx + 1}행")
    c_meta2.caption(f"운송장 파일 헤더 행: {tracking_header_idx + 1}행")
    c_meta3.caption(f"배송 취소 행 제외: {skipped_cancelled}건 / 결과 제외: {result_excluded_cancelled}건")

    with st.expander("미리보기 (상위 30건)", expanded=False):
        st.dataframe(out_df.head(30), use_container_width=True)

    miss = (out_df["송장번호"].isna() | (out_df["송장번호"].astype(str).str.strip() == "")).sum()
    st.write(f"총 {len(out_df)}건 / 송장번호 누락 {miss}건")

    courier_counts = (
        out_df.loc[out_df["송장번호"].astype(str).str.strip() != "", "택배사"]
        .replace("", pd.NA)
        .dropna()
        .value_counts()
    )
    if not courier_counts.empty:
        st.caption("택배사 반영: " + " / ".join(f"{name} {cnt}건" for name, cnt in courier_counts.items()))

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
