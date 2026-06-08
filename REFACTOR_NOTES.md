# 리팩토링 적용 메모

기존 기능 유지를 우선으로 app.py의 기능을 단계적으로 분리했습니다.

## 분리된 파일

- `config.py`: 환경변수, 경로, 공통 상수, ReportLab import, `load_service_account_info`
- `drive_utils.py`: Google Drive 폴더/파일 저장·조회·다운로드
- `sheets_utils.py`: Google Sheets 서비스 생성, Sheet ID 추출
- `sales.py`: 매출계산 수동/자동 화면과 계산/기록 로직
- `invoice.py`: 송장등록 화면
- `stock.py`: 재고일괄변경 화면
- `app.py`: Streamlit 진입점, 라우팅, 그 외 페이지(주문엑셀/재고/총합/매핑) + PDF 생성

## 추가 정리 사항 (2차)

1. **`from config import *` 제거** → app.py에서 실제로 쓰는 32개 이름만 명시적 import
   - 코드를 읽을 때 어떤 상수가 어디서 왔는지 추적 가능
   - 이후 새 변수 만들 때 우연한 덮어쓰기 위험 제거

2. **중복 함수 정리**
   - `fmt_qty_no_zero`: app.py에서 삭제 (실제 사용처 없는 dead code, stock.py에 남김)
   - `load_service_account_info`: drive_utils.py의 중복 정의 제거 (config.py 버전을 import하여 사용)

3. **빈 placeholder 모듈 삭제**
   - `telegram_bot.py`, `naver_api.py` 제거 — 실제 텔레그램/네이버 코드는 `stock.py` 안에 있음
   - 향후 `render_bulk_stock_page` 분할 작업 시 함께 추출할 예정

## 중요

- Render 실행 명령은 기존처럼 `streamlit run app.py`를 유지하면 됩니다.
- 기존 환경변수 이름은 변경하지 않았습니다.
- `fonts/NanumGothic.ttf`는 기존 저장소에 있는 파일을 그대로 유지하세요. 이 ZIP에는 폰트 파일을 포함하지 않았습니다.

## 다음 단계 (선택)

- `render_bulk_stock_page` 내부의 63개 중첩 함수를 stock.py 모듈 레벨로 끌어올리기
- 텔레그램 / 네이버 API 호출 부분을 별도 모듈로 추출
- app.py에 남아있는 PDF 생성 함수(`build_sticker_pdf`, `build_recipient_pdf`, `build_summary_pdf`)를 `pdf_utils.py`로 분리
