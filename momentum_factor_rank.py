
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import os

CURRENCY_THRESHOLDS = {
    "USD": 2_000_000_000,
    "CAD": 2_700_000_000,
    "EUR": 1_850_000_000,
    "JPY": 300_000_000_000,
    "GBP": 1_600_000_000
}

# Dynamic exchange suffix logic based on currency/country assumptions
def infer_yf_ticker(symbol, currency):
    currency = currency.upper()

    if currency == "CAD":
        return f"{symbol}.TO"
    elif currency == "JPY":
        return f"{symbol}.T"
    elif currency == "EUR":
        # Try both .DE and .PA — fallback logic below will try each
        return [f"{symbol}.PA", f"{symbol}.DE"]
    elif currency == "GBP":
        return f"{symbol}.L"
    else:
        return symbol

def resolve_market_cap_ticker(possible_symbols):
    if isinstance(possible_symbols, str):
        possible_symbols = [possible_symbols]
    for symbol in possible_symbols:
        try:
            info = yf.Ticker(symbol).info
            market_cap = info.get("marketCap", None)
            if market_cap is not None:
                return symbol, market_cap
        except Exception:
            continue
    return None, None

def calculate_momentum(symbol):
    try:
        df = yf.download(symbol, period="7mo", interval="1d", auto_adjust=True)
        df = df.dropna()
        if len(df) <= 21:
            return -1
        df = df.iloc[:-21]
        cumulative_return = (df["Close"].iloc[-1] - df["Close"].iloc[0]) / df["Close"].iloc[0]
        return float(cumulative_return.iloc[0]) if isinstance(cumulative_return, pd.Series) else float(cumulative_return)
    except Exception as e:
        print(f"[ERROR] Momentum calc failed for {symbol}: {e}")
        return -1

def main():
    print("[INFO] Loading tickers...")
    csv_path = r"C:\Users\murph\OneDrive\Documents\Python Scripts\Momentum Factor Rank\all_tickers.csv"
    df = pd.read_csv(csv_path)
    df = df.dropna(subset=["Symbol"])

    results = []
    skipped = []

    for _, row in df.iterrows():
        raw_symbol = str(row["Symbol"]).replace(".", "-")
        currency = str(row.get("Price Currency", "USD")).upper()
        possible_symbols = infer_yf_ticker(raw_symbol, currency)

        print(f"[INFO] Processing {raw_symbol} ({currency})")

        resolved_symbol, market_cap = resolve_market_cap_ticker(possible_symbols)
        if not resolved_symbol or market_cap is None:
            print(f"[INFO] Skipping {raw_symbol}: market cap not available.")
            skipped.append(raw_symbol)
            continue

        threshold = CURRENCY_THRESHOLDS.get(currency, 2_000_000_000)
        if market_cap < threshold:
            print(f"[INFO] Skipping {resolved_symbol}: market cap {market_cap} below threshold.")
            skipped.append(resolved_symbol)
            continue

        momentum = calculate_momentum(resolved_symbol)
        if momentum == -1:
            skipped.append(resolved_symbol)
            continue

        results.append({
            "Symbol": resolved_symbol,
            "Momentum": momentum
        })

    if skipped:
        print("[INFO] Skipped symbols:")
        for sym in skipped:
            print(f"  - {sym}")

    if not results:
        print("[ERROR] No valid data collected.")
        return

    result_df = pd.DataFrame(results)
    result_df["Z_Momentum"] = (result_df["Momentum"] - result_df["Momentum"].mean()) / result_df["Momentum"].std()
    result_df.sort_values("Z_Momentum", ascending=False, inplace=True)
    result_df.reset_index(drop=True, inplace=True)

    excel_path = r"C:\Users\murph\OneDrive\Documents\Python Scripts\Momentum Factor Rank\momentum_ranked_stocks.xlsx"
    result_df.to_excel(excel_path, index=False)
    print(f"[SUCCESS] Exported to {excel_path}")

    plt.figure(figsize=(10, 6))
    top_n = result_df.head(10)
    plt.barh(top_n["Symbol"], top_n["Z_Momentum"], color="skyblue")
    plt.xlabel("Momentum Z-Score")
    plt.title("Top 10 Stocks by Momentum (6M excl. last month)")
    plt.gca().invert_yaxis()
    plt.tight_layout()
    chart_path = r"C:\Users\murph\OneDrive\Documents\Python Scripts\Momentum Factor Rank\momentum_chart.png"
    plt.savefig(chart_path)
    print(f"[INFO] Bar chart saved as {chart_path}")

    try:
        os.startfile(excel_path)
    except Exception as e:
        print(f"[WARN] Could not open Excel file: {e}")

if __name__ == "__main__":
    main()
