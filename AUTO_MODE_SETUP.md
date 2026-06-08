# 자동모드 설정 순서

1. 이 폴더 전체를 GitHub 저장소에 업로드합니다.
2. Render 환경변수에 아래 값을 넣습니다.
   - NAVER_COMMERCE_CLIENT_ID
   - NAVER_COMMERCE_CLIENT_SECRET
   - NAVER_COMMERCE_BASE_URL=https://api.commerce.naver.com/external
   - TELEGRAM_BOT_TOKEN
   - TELEGRAM_CHAT_ID
   - TELEGRAM_AUTO_POLL_SECONDS=5
3. 필요하면 NAVER_PRODUCT_SEARCH_BODY 환경변수에 네이버 상품 조회용 JSON을 넣습니다.
   - 기본값은 {"page":1,"size":500} 입니다.
4. 배포 후 재고일괄변경으로 들어갑니다.
5. 입력 방식에서 자동을 고릅니다.
6. 자동 실행 시작을 누릅니다.
7. 텔레그램으로 온 메시지에 한 번만 답장합니다.
   - 예시
     1 5
     2 0
     3 10
8. 앱이 첫 번째 유효한 답장 1개를 자동 처리합니다.
9. 메모 텍스트를 다시 텔레그램으로 보냅니다.
10. 0보다 큰 값만 네이버 API로 자동 반영합니다.

주의:
- 자동모드는 첫 번째 유효 답장 1개만 처리합니다.
- 답장은 반드시 한 메시지로 보내세요.
- 형식은 번호 수량 입니다.
- 수동모드는 기존과 동일합니다.
