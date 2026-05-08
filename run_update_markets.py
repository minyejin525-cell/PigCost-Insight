from pathlib import Path
from urllib.parse import quote
import time
import urllib3
import requests
import pandas as pd

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

SERIES_CONFIG = {
    "corn_90d.csv": {
        "symbol": "ZC=F",
        "label": "옥수수",
    },
    "soymeal_90d.csv": {
        "symbol": "ZM=F",
        "label": "대두박",
    },
    "freight_90d.csv": {
        "symbol": "BDRY",
        "label": "해상운임(BDRY 프록시)",
    },
}

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

def fetch_yahoo_chart_90d(symbol: str) -> pd.DataFrame:
    last_error = None

    for attempt in range(1, 4):
        try:
            encoded_symbol = quote(symbol, safe="")
            url = (
                f"https://query1.finance.yahoo.com/v8/finance/chart/"
                f"{encoded_symbol}?range=6mo&interval=1d&includePrePost=false&events=div,splits"
            )

            print(f"[시도 {attempt}] {symbol} 요청 중...")

            response = requests.get(
                url,
                headers=HEADERS,
                timeout=30,
                verify=False
            )
            response.raise_for_status()

            data = response.json()

            chart = data.get("chart", {})
            error = chart.get("error")
            if error:
                raise ValueError(f"{symbol} Yahoo 응답 오류: {error}")

            result_list = chart.get("result")
            if not result_list:
                raise ValueError(f"{symbol} 결과(result)가 없습니다.")

            result = result_list[0]
            timestamps = result.get("timestamp", [])
            indicators = result.get("indicators", {})
            quote_list = indicators.get("quote", [])

            if not timestamps or not quote_list:
                raise ValueError(f"{symbol} 시계열 데이터가 없습니다.")

            closes = quote_list[0].get("close", [])
            if not closes:
                raise ValueError(f"{symbol} 종가 데이터가 없습니다.")

            rows = []
            for ts, close in zip(timestamps, closes):
                if close is None:
                    continue
                rows.append({
                    "date": pd.to_datetime(ts, unit="s").strftime("%Y-%m-%d"),
                    "price": float(close)
                })

            df = pd.DataFrame(rows)

            if df.empty:
                raise ValueError(f"{symbol} 정리 후 데이터가 비어 있습니다.")

            df["date"] = pd.to_datetime(df["date"], errors="coerce")
            df["price"] = pd.to_numeric(df["price"], errors="coerce")

            df = (
                df.dropna(subset=["date", "price"])
                  .sort_values("date")
                  .tail(90)
                  .reset_index(drop=True)
            )

            if df.empty:
                raise ValueError(f"{symbol} 최근 90일 데이터가 비어 있습니다.")

            df["date"] = df["date"].dt.strftime("%Y-%m-%d")
            return df

        except Exception as e:
            last_error = e
            print(f"[시도 {attempt}] {symbol} 가져오기 실패: {e}")
            time.sleep(3)

    raise last_error


def save_csv(df: pd.DataFrame, out_path: Path):
    df.to_csv(out_path, index=False, encoding="utf-8-sig")


def main():
    print("=== 시장 데이터 업데이트 시작 ===")

    for filename, cfg in SERIES_CONFIG.items():
        symbol = cfg["symbol"]
        label = cfg["label"]
        out_path = DATA_DIR / filename

        try:
            print(f"{label} 요청 중... ({symbol})")
            df = fetch_yahoo_chart_90d(symbol)
            save_csv(df, out_path)

            latest = df.iloc[-1]
            print(f"저장 완료: {out_path}")
            print(f"최신값: {latest['date']} / {latest['price']:.2f}")

        except Exception as e:
            print(f"{label} 업데이트 실패: {e}")

            if out_path.exists():
                print(f"기존 파일 유지: {out_path}")
            else:
                print(f"기존 파일도 없음: {out_path}")

    print("=== 시장 데이터 업데이트 종료 ===")

# force redeploy
if __name__ == "__main__":
    main()
