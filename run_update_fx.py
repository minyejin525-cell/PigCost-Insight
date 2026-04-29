from pathlib import Path
import io
import time
import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

FRED_URL = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DEXKOUS"
OUT_PATH = DATA_DIR / "fx_90d.csv"


def make_sample_fx():
    dates = pd.date_range(end=pd.Timestamp.today().normalize(), periods=90)
    values = [1450 + (i % 7) * 2 + (i // 15) for i in range(90)]
    return pd.DataFrame({
        "date": dates,
        "price": values
    })


def fetch_fx_data():
    last_error = None

    for attempt in range(1, 4):  # 3번 재시도
        try:
            print(f"[시도 {attempt}] FRED 환율 데이터 요청 중...")

            response = requests.get(
                FRED_URL,
                timeout=60,  # 20 -> 60으로 증가
                headers={"User-Agent": "Mozilla/5.0"},
                verify=False
            )
            response.raise_for_status()

            df = pd.read_csv(io.StringIO(response.text))
            df.columns = ["date", "price"]
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
            df["price"] = pd.to_numeric(df["price"], errors="coerce")

            df = df.dropna(subset=["date", "price"]).sort_values("date").tail(90)

            if df.empty:
                raise ValueError("가져온 환율 데이터가 비어 있습니다.")

            print("[성공] FRED 데이터 수집 완료")
            return df

        except Exception as e:
            last_error = e
            print(f"[실패] 시도 {attempt} 실패: {e}")
            if attempt < 3:
                print("5초 후 재시도합니다...")
                time.sleep(5)

    raise last_error


def save_fx_data(df):
    df.to_csv(OUT_PATH, index=False, encoding="utf-8-sig")
    print(f"[저장 완료] {OUT_PATH}")


def main():
    try:
        df = fetch_fx_data()
        save_fx_data(df)
        print(df.tail())

    except Exception as e:
        print(f"[최종 실패] FRED 데이터 수집 실패: {e}")

        # 1순위: 기존 파일 유지
        if OUT_PATH.exists():
            print("[대체] 기존 fx_90d.csv 파일이 있으므로 그대로 사용합니다.")
            old_df = pd.read_csv(OUT_PATH, encoding="utf-8-sig")
            print(old_df.tail())
            return

        # 2순위: 샘플 데이터 생성
        print("[대체] 기존 파일도 없어서 샘플 환율 데이터를 생성합니다.")
        sample_df = make_sample_fx()
        save_fx_data(sample_df)
        print(sample_df.tail())


if __name__ == "__main__":
    main()
