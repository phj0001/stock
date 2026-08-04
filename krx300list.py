import os
import datetime
import requests
import pandas as pd
from dotenv import load_dotenv

# .env 환경 변수 로드
load_dotenv()

# pykrx 모듈 임포트 시도
try:
    from pykrx import stock
    HAS_PYKRX = True
except ImportError:
    HAS_PYKRX = False


def get_recent_business_day(base_date=None):
    """주말 및 공휴일을 고려한 직전 영업일 YYYYMMDD 계산"""
    if base_date is None:
        base_date = datetime.datetime.now()
    
    target = base_date - datetime.timedelta(days=1)
    while target.weekday() >= 5:  # 토(5), 일(6) 제외
        target -= datetime.timedelta(days=1)
        
    return target.strftime("%Y%m%d")


def fetch_via_pykrx(target_date):
    """1차 시도: pykrx 및 KRX 로그인 세션 활용"""
    if not HAS_PYKRX:
        return None
    
    krx_id = os.getenv("KRX_ID")
    krx_pw = os.getenv("KRX_PW")
    if krx_id and krx_pw:
        print(f"🔒 [1차 시도] pykrx + KRX 계정 세션({krx_id[:3]}***)으로 조회 중...")
    else:
        print("⏳ [1차 시도] pykrx 지수 구성종목 조회 중...")

    try:
        # KRX 시장 지수 목록에서 'KRX 300' 지수코드 동적 탐색
        krx300_code = None
        indices = stock.get_index_ticker_list(target_date, market="KRX")
        for idx in indices:
            idx_name = stock.get_index_ticker_name(idx)
            if "KRX 300" in idx_name or "KRX300" in idx_name:
                krx300_code = idx
                break

        codes_to_try = [krx300_code, "5042", "1028"]
        codes_to_try = [c for c in codes_to_try if c]

        for code in codes_to_try:
            tickers = stock.get_index_portfolio_deposit_file(code, target_date)
            if isinstance(tickers, (pd.DataFrame, pd.Series)):
                tickers = tickers.iloc[:, 0].tolist() if isinstance(tickers, pd.DataFrame) else tickers.tolist()
            
            if isinstance(tickers, list) and len(tickers) >= 200:
                print(f"✅ pykrx 수집 성공! ({len(tickers)}개 종목)")
                data = []
                for t in tickers:
                    name = stock.get_market_ticker_name(t)
                    data.append({'종목코드': str(t).zfill(6), '종목명': name})
                return pd.DataFrame(data)
    except Exception as e:
        print(f"⚠️ 1차 시도(pykrx) 차단 또는 에러: {e}")
    
    return None


def fetch_via_naver_mobile():
    """2차 시도: 네이버 모바일 금융 API (우회 수집)"""
    print("🔄 [2차 시도] 네이버 모바일 금융 API로 전환 수집 중...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1',
        'Referer': 'https://m.stock.naver.com/'
    }
    
    data_list = []
    # 네이버 모바일 KRX300 구성종목 API (페이지당 100개씩 순회)
    for page in range(1, 6):
        url = f"https://m.stock.naver.com/api/index/KRX300/components?page={page}&pageSize=100"
        try:
            res = requests.get(url, headers=headers, timeout=10)
            if res.status_code != 200:
                break
            
            json_data = res.json()
            items = json_data.get('stocks', json_data.get('result', json_data.get('components', [])))
            
            if not items:
                break
                
            for item in items:
                code = item.get('itemCode', item.get('reutersCode', ''))
                name = item.get('stockName', item.get('itemName', ''))
                if code and name:
                    code_clean = str(code).replace('A', '').zfill(6)
                    data_list.append({'종목코드': code_clean, '종목명': str(name).strip()})
        except Exception as e:
            print(f"⚠️ 네이버 API {page}페이지 수집 중 에러: {e}")
            continue

    if len(data_list) >= 200:
        df = pd.DataFrame(data_list).drop_duplicates(subset=['종목코드'])
        print(f"✅ 네이버 모바일 API 수집 성공! ({len(df)}개 종목)")
        return df
    
    return None


def fetch_via_daum():
    """3차 시도: Daum 금융 API (추가 백업)"""
    print("🔄 [3차 시도] Daum 금융 API로 전환 수집 중...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://finance.daum.net/'
    }
    data_list = []
    for page in range(1, 6):
        url = f"https://finance.daum.net/api/quotes/stocks?pageSize=100&page={page}&market=KRX300"
        try:
            res = requests.get(url, headers=headers, timeout=10)
            if res.status_code != 200:
                break
            json_data = res.json()
            data = json_data.get('data', [])
            if not data:
                break
            for item in data:
                code = item.get('symbolCode', '')
                if code.startswith('A'):
                    code = code[1:]
                name = item.get('name', '')
                if code and name:
                    data_list.append({'종목코드': str(code).zfill(6), '종목명': str(name).strip()})
        except Exception:
            continue

    if len(data_list) >= 200:
        df = pd.DataFrame(data_list).drop_duplicates(subset=['종목코드'])
        print(f"✅ Daum API 수집 성공! ({len(df)}개 종목)")
        return df
    
    return None


def main():
    now = datetime.datetime.now()
    target_date = get_recent_business_day(now)
    print(f"🚀 [{target_date}] 기준 KRX 300 전체 종목 수집 파이프라인을 시작합니다.\n")

    # 1차: pykrx 시도
    df_result = fetch_via_pykrx(target_date)

    # 2차: 네이버 모바일 API 백업
    if df_result is None or len(df_result) < 200:
        df_result = fetch_via_naver_mobile()

    # 3차: Daum API 백업
    if df_result is None or len(df_result) < 200:
        df_result = fetch_via_daum()

    if df_result is None or df_result.empty:
        print("\n❌ 모든 수집 경로에서 데이터를 불러오는 데 실패했습니다.")
        return

    # krx300list.csv 저장 (인덱스 제외, 한글 깨짐 방지 utf-8-sig)
    filename = "krx300list.csv"
    df_result.to_csv(filename, index=False, encoding="utf-8-sig")

    print(f"\n🎉 성공적으로 수집 및 저장되었습니다: {filename}")
    print(f"📊 총 수집된 종목 수: {len(df_result)}개")
    print("\n--- 수집 결과 미리보기 (상위 10개) ---")
    print(df_result.head(10))


if __name__ == "__main__":
    main()