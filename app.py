import re
import json
from datetime import datetime
from urllib.parse import urlparse

import httpx
from bs4 import BeautifulSoup
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, RedirectResponse

app = FastAPI()

# 아주 간단한 인메모리 저장소(서버 재시작하면 초기화됨)
ITEMS = []  # [{"url":..., "title":..., "price":..., "currency":..., "updated_at":..., "error":...}]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0 Safari/537.36"
)

def _clean_price_number(text: str) -> str:
    """
    가격 문자열에서 숫자/구분자만 남기고 정규화.
    예: "12,900원" -> "12900"
        "$1,299.99" -> "1299.99"
    """
    if not text:
        return ""
    t = text.strip()
    # 통화 기호/문자 제거 (숫자, 콤마, 점만 유지)
    t = re.sub(r"[^\d,\.]", "", t)

    # 케이스 처리:
    # - "1,299.99" : 콤마는 천 단위, 점은 소수점 -> 콤마 제거
    # - "1.299,99" : 유럽식 -> 점 제거, 콤마를 점으로
    if "," in t and "." in t:
        # 마지막 구분자가 소수점일 가능성이 큼
        if t.rfind(".") > t.rfind(","):
            t = t.replace(",", "")
        else:
            t = t.replace(".", "").replace(",", ".")
    else:
        # 콤마만 있으면 천 단위로 보고 제거 (대부분 KR 사이트)
        if "," in t and "." not in t:
            t = t.replace(",", "")

    return t

def _iter_json_objects(obj):
    """JSON-LD가 list/dict 혼합인 경우 전부 순회."""
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from _iter_json_objects(v)
    elif isinstance(obj, list):
        for x in obj:
            yield from _iter_json_objects(x)

def extract_title_and_price(html: str, url: str):
    soup = BeautifulSoup(html, "lxml")

    # ---- title 추출(범용) ----
    title = None
    og_title = soup.select_one('meta[property="og:title"]')
    if og_title and og_title.get("content"):
        title = og_title["content"].strip()

    if not title:
        if soup.title and soup.title.string:
            title = soup.title.string.strip()

    # ---- price 추출(1) JSON-LD ----
    # schema.org Product / Offer / AggregateOffer에서 price 찾기
    for script in soup.select('script[type="application/ld+json"]'):
        raw = script.string
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue

        for node in _iter_json_objects(data):
            # offers가 dict 또는 list일 수 있음
            offers = node.get("offers") if isinstance(node, dict) else None
            if not offers:
                continue

            offer_nodes = offers if isinstance(offers, list) else [offers]
            for offer in offer_nodes:
                if not isinstance(offer, dict):
                    continue

                price = offer.get("price")
                currency = offer.get("priceCurrency")

                # AggregateOffer 케이스: lowPrice/highPrice
                if price is None:
                    price = offer.get("lowPrice") or offer.get("highPrice")

                if price is not None:
                    # price가 숫자거나 문자열일 수 있음
                    price_str = _clean_price_number(str(price))
                    if price_str:
                        return title, price_str, currency

    # ---- price 추출(2) meta 태그 ----
    meta_price = soup.select_one('meta[property="product:price:amount"], meta[itemprop="price"]')
    if meta_price and meta_price.get("content"):
        price_str = _clean_price_number(meta_price["content"])
        currency = None
        meta_currency = soup.select_one('meta[property="product:price:currency"], meta[itemprop="priceCurrency"]')
        if meta_currency and meta_currency.get("content"):
            currency = meta_currency["content"].strip()
        if price_str:
            return title, price_str, currency

    # ---- price 추출(3) 도메인별 선택자(필요시 추가) ----
    # 예시: 특정 쇼핑몰의 가격 영역이 확실할 때 여기에 추가
    # domain = urlparse(url).netloc.lower()
    # if "example.com" in domain:
    #     el = soup.select_one(".price strong")
    #     if el:
    #         price_str = _clean_price_number(el.get_text(" ", strip=True))
    #         if price_str:
    #             return title, price_str, None

    return title, None, None

async def fetch_html(url: str) -> str:
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8"}
    async with httpx.AsyncClient(follow_redirects=True, timeout=20) as client:
        r = await client.get(url, headers=headers)
        r.raise_for_status()
        return r.text

async def update_one(item: dict):
    try:
        html = await fetch_html(item["url"])
        title, price, currency = extract_title_and_price(html, item["url"])
        item["title"] = title or item.get("title") or "(제목 미확인)"
        item["price"] = price
        item["currency"] = currency
        item["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        item["error"] = None if price else "가격을 찾지 못함(동적 렌더링/선택자 필요 가능)"
    except Exception as e:
        item["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        item["error"] = f"가져오기 실패: {type(e).__name__}: {e}"

def render_page():
    rows = []
    for i, it in enumerate(ITEMS):
        price_display = it["price"] if it.get("price") else "-"
        cur = it.get("currency") or ""
        err = it.get("error") or ""
        rows.append(f"""
          <tr>
            <td>{i+1}</td>
            <td style="max-width:520px;word-break:break-all;">
              <a href="{it["url"]}" target="_blank" rel="noreferrer">{it["title"] or it["url"]}</a><br/>
              <small>{it["url"]}</small>
            </td>
            <td>{price_display} {cur}</td>
            <td><small>{it.get("updated_at","")}</small></td>
            <td><small style="color:#b00020;">{err}</small></td>
          </tr>
        """)

    html = f"""
    <!doctype html>
    <html lang="ko">
    <head>
      <meta charset="utf-8"/>
      <meta name="viewport" content="width=device-width,initial-scale=1"/>
      <title>상품 가격 트래커</title>
      <style>
        body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 24px; }}
        .box {{ padding: 16px; border: 1px solid #ddd; border-radius: 12px; margin-bottom: 16px; }}
        input[type=text] {{ width: 100%; padding: 10px; border-radius: 10px; border: 1px solid #ccc; }}
        button {{ padding: 10px 14px; border-radius: 10px; border: 1px solid #ccc; cursor: pointer; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border-bottom: 1px solid #eee; padding: 10px; vertical-align: top; text-align: left; }}
        th {{ background: #fafafa; }}
        .actions {{ display:flex; gap:8px; flex-wrap:wrap; }}
      </style>
    </head>
    <body>
      <h1>상품 가격 트래커</h1>

      <div class="box">
        <form method="post" action="/add">
          <label>상품 URL 추가</label><br/>
          <input type="text" name="url" placeholder="https://..." required/>
          <div style="height:10px;"></div>
          <div class="actions">
            <button type="submit">추가</button>
          </div>
        </form>
      </div>

      <div class="box">
        <form method="post" action="/refresh">
          <div class="actions">
            <button type="submit">전체 가격 새로고침</button>
            <a href="/" style="align-self:center;">새로고침</a>
          </div>
          <p style="margin:10px 0 0 0;"><small>
            ※ 가격을 못 찾는 경우: 사이트가 JS로 가격을 그리거나, 구조가 특수해서 도메인별 선택자/Playwright가 필요할 수 있어요.
          </small></p>
        </form>
      </div>

      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>상품</th>
            <th>가격</th>
            <th>업데이트</th>
            <th>상태</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows) if rows else '<tr><td colspan="5">아직 등록된 URL이 없어요.</td></tr>'}
        </tbody>
      </table>
    </body>
    </html>
    """
    return html

@app.get("/", response_class=HTMLResponse)
async def home():
    return render_page()

@app.post("/add")
async def add(url: str = Form(...)):
    url = url.strip()
    ITEMS.append({
        "url": url,
        "title": None,
        "price": None,
        "currency": None,
        "updated_at": None,
        "error": "아직 업데이트 안 됨",
    })
    return RedirectResponse("/", status_code=303)

@app.post("/refresh")
async def refresh():
    # 간단히 순차 업데이트(상품이 많아지면 병렬/큐/스케줄러 권장)
    for it in ITEMS:
        await update_one(it)
    return RedirectResponse("/", status_code=303)
