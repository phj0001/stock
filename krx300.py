import datetime
import pandas as pd
from pykrx import stock

def get_krx300_data():
    # 오늘 날짜 가져오기
    today = datetime.datetime.now().strftime("%Y%m%d")
    print(f"[{today}] KRX 300 종목 정보를 수집 중입니다...")

    try:
        # 1. KRX 300 지수 구성 종목 코드 리스트 가져오기 (지수코드 '1013': KRX 300)
        krx300_tickers = stock.get_index_portfolio_deposit_file("1013", today)

        # 수집된 종목 리스트가 없는 경우 (주말, 공휴일, 장 개시 전)
        if not krx300_tickers:
            print("⚠️ 지정한 날짜의 KRX 300 종목 정보를 가져올 수 없습니다. (휴일/장 개시 전일 수 있습니다.)")
            return None

        # 2. 전체 종목 시세 및 시가총액 데이터 수집
        df_price = stock.get_market_ohlcv_by_ticker(today, market="ALL")
        df_cap = stock.get_market_cap_by_ticker(today, market="ALL")

        # 데이터가 비어 있는지 .empty 로 확인
        if df_price.empty or df_cap.empty:
            print("⚠️ 거래소 시세 데이터가 비어있습니다.")
            return None

        # 3. KRX 300 구성 종목만 필터링
        df_krx300_price = df_price.loc[krx300_tickers].copy()
        df_krx300_cap = df_cap.loc[krx300_tickers][['시가총액', '상장주식수']]

        # 4. 데이터 병합 (시세 + 시가총액)
        df_result = pd.concat([df_krx300_price, df_krx300_cap], axis=1)

        # 5. 종목명(Name) 컬럼 추가
        df_result['종목명'] = [stock.get_market_ticker_name(ticker) for ticker in df_result.index]

        # 6. 컬럼 순서 및 인덱스 정리
        df_result.index.name = '종목코드'
        cols = ['종목명', '종가', '대비', '등락률', '시가', '고가', '저가', '거래량', '거래대금', '시가총액', '상장주식수']
        df_result = df_result[cols]

        # 7. CSV 파일로 저장
        filename = f"krx300_info_{today}.csv"
        df_result.to_csv(filename, encoding="utf-8-sig")
        
        print(f"✅ 성공적으로 저장되었습니다: {filename}")
        print(f"총 수집된 종목 수: {len(df_result)}개")
        
        return df_result

    except Exception as e:
        print(f"❌ 데이터 수집 중 오류 발생: {e}")
        return None

if __name__ == "__main__":
    df = get_krx300_data()
    # DataFrame 검사 시 .empty 속성 사용 (.isnot None 조건 병행)
    if df is not None and not df.empty:
        print("\n--- 상위 5개 종목 미리보기 ---")
        print(df.head())