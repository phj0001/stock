import os
import datetime
import pandas as pd
from dotenv import load_dotenv
from pykrx import stock

# 1. .env 파일에서 KRX_ID, KRX_PW 환경 변수 불러오기
load_dotenv()

# 로그인 정보 로드 확인
krx_id = os.getenv("KRX_ID")
krx_pw = os.getenv("KRX_PW")

if krx_id and krx_pw:
    print(f"🔒 KRX 로그인 계정이 감지되었습니다: ID({krx_id[:3]}***)")
else:
    print("⚠️ WARNING: KRX 로그인 정보가 .env 파일에 설정되지 않았습니다.")


def get_krx300_tickers_safely(date_str):
    """
    KRX 300 종목 코드 리스트를 안전하게 가져오는 함수
    """
    try:
        tickers = stock.get_index_portfolio_deposit_file("1013", date_str)
        if isinstance(tickers, (pd.DataFrame, pd.Series)):
            if tickers.empty:
                return None
            return tickers.iloc[:, 0].tolist() if isinstance(tickers, pd.DataFrame) else tickers.tolist()
        
        if isinstance(tickers, list) and len(tickers) > 0:
            return tickers
            
        return None
    except Exception as e:
        print(f"⚠️ 지수 구성종목 조회 중 예외 발생: {e}")
        return None


def get_recent_business_day(base_date=None):
    """주말 및 휴일을 고려하여 이전 영업일 날짜를 계산"""
    if base_date is None:
        base_date = datetime.datetime.now()
    
    target = base_date - datetime.timedelta(days=1)
    while target.weekday() >= 5:
        target -= datetime.timedelta(days=1)
        
    return target.strftime("%Y%m%d")


def get_krx300_data():
    now = datetime.datetime.now()
    today_str = now.strftime("%Y%m%d")
    
    print(f"\n[{today_str}] KRX 300 종목 정보 수집을 시작합니다...")

    # 1차 시도: 오늘 날짜
    target_date = today_str
    krx300_tickers = get_krx300_tickers_safely(target_date)

    # 당일 데이터 생성 전(장 마감 전 등)이면 직전 영업일로 전환
    if not krx300_tickers:
        target_date = get_recent_business_day(now)
        print(f"🔄 당일({today_str}) 데이터 미생성으로 직전 영업일({target_date}) 데이터로 전환합니다...")
        krx300_tickers = get_krx300_tickers_safely(target_date)

    if not krx300_tickers:
        print(f"❌ [{target_date}] 기준 KRX 300 종목 정보를 가져오지 못했습니다.")
        return None

    try:
        # 시세 및 시가총액 데이터 수집
        df_price = stock.get_market_ohlcv_by_ticker(target_date, market="ALL")
        df_cap = stock.get_market_cap_by_ticker(target_date, market="ALL")

        if df_price.empty or df_cap.empty:
            print("⚠️ 시세 데이터가 비어 있습니다.")
            return None

        # 유효한 티커 필터링
        valid_tickers = [t for t in krx300_tickers if t in df_price.index]
        
        df_krx300_price = df_price.loc[valid_tickers].copy()
        df_krx300_cap = df_cap.loc[valid_tickers][['시가총액', '상장주식수']]

        # 데이터 병합
        df_result = pd.concat([df_krx300_price, df_krx300_cap], axis=1)

        # 종목명 추가
        df_result['종목명'] = [stock.get_market_ticker_name(ticker) for ticker in df_result.index]

        # 정렬 및 존재하는 컬럼만 동적 필터링
        df_result.index.name = '종목코드'
        target_cols = ['종목명', '종가', '등락률', '시가', '고가', '저가', '거래량', '거래대금', '시가총액', '상장주식수']
        existing_cols = [c for c in target_cols if c in df_result.columns]
        df_result = df_result[existing_cols]

        # CSV 저장
        filename = f"krx300_info_{target_date}.csv"
        df_result.to_csv(filename, encoding="utf-8-sig")
        
        print(f"✅ 데이터가 성공적으로 저장되었습니다: {filename}")
        print(f"총 수집된 종목 수: {len(df_result)}개")
        return df_result

    except Exception as e:
        print(f"❌ 데이터 처리 중 오류 발생: {e}")
        return None


if __name__ == "__main__":
    df = get_krx300_data()
    if df is not None and not df.empty:
        print("\n--- 상위 5개 종목 미리보기 ---")
        print(df.head())