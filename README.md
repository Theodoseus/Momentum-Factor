# Momentum Factor Ranker

A Python-based tool that ranks global large-cap stocks using a 6-month momentum factor (excluding the most recent month), adjusted for minimum market cap by currency and region. Results are exported to Excel and visualized in a chart.

## 📈 Overview

This script is designed for portfolio managers, analysts, and data-driven investors looking to systematically rank stocks by recent price momentum. It:

* Loads a list of global stock tickers and associated price currencies from a CSV file
* Infers the correct Yahoo Finance ticker format for each name
* Screens for minimum market cap thresholds based on currency
* Calculates 6-month momentum (excluding the most recent month to avoid short-term noise)
* Ranks all qualifying stocks by Z-scored momentum
* Exports results to Excel
* Generates a bar chart of the top 10 stocks

## ✅ Example Use Case

You maintain a curated universe of global large-cap stocks (e.g., from analyst "top picks" or internal lists) and want a fast way to quantify recent momentum across those names. This script lets you rank and visualize them with minimal manual effort.

## 🔧 Requirements

* Python 3.8+
* `pandas`
* `matplotlib`
* `yfinance`

You can install dependencies via:

```bash
pip install pandas matplotlib yfinance
```

## 📂 Input Format

The script expects an input CSV file with at least two columns:

```
Symbol,Price Currency
AAPL,USD
TD,CAD
BMW,EUR
```

> 📍 The file path to this CSV is currently hardcoded as:

```python
C:\Users\murph\OneDrive\Documents\Python Scripts\Momentum Factor Rank\all_tickers.csv
```

Update this path as needed to match your environment.

## 💡 How It Works

1. **Symbol Mapping:** Uses suffix rules to convert symbols to Yahoo Finance-compatible tickers based on currency/region.
2. **Market Cap Filter:** Applies minimum market cap screens:

   * USD: $2B
   * CAD: $2.7B
   * EUR: €1.85B
   * GBP: £1.6B
   * JPY: ¥300B
3. **Momentum Calculation:** Computes price return over a 6-month period ending one month ago.
4. **Ranking:** Standardizes momentum scores via Z-score, then sorts descending.
5. **Output:** Saves results to Excel and charts the top 10 stocks by momentum Z-score.

## 📊 Output

* `momentum_ranked_stocks.xlsx` — Excel file with symbols, raw momentum, and Z-scores
* `momentum_chart.png` — horizontal bar chart of top 10 momentum stocks

Both files are saved in the working directory.

## 📌 Notes

* Symbols that fail to resolve or lack sufficient data are skipped with warnings.
* The script is currently configured for local Windows paths and uses `os.startfile()` to open the Excel output.
* Consider adapting the code to accept dynamic input/output paths or CLI arguments for more flexibility.

## 🚀 Next Steps

* Add multi-factor support (e.g. value, quality)
* Enable CLI arguments or GUI
* Integrate live data refresh and reporting via email or dashboards
