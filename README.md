student_tribe_assesment_task_1
# Live Cryptocurrency Data Fetcher & Analyzer

## Project Overview

This Python script fetches live data for the top 50 cryptocurrencies by market capitalization using the CoinGecko API. It performs basic analysis and continuously updates an Excel file every 2 minutes with fresh cryptocurrency data and insights.

## Output

- **Console Output:**
  - Top 5 cryptocurrencies by market capitalization.
  - Average price of the top 50 cryptocurrencies.
  - Highest and lowest 24-hour price changes.

- **Excel File (`live_crypto_data.xlsx`):**
  - Columns include:
    - Cryptocurrency Name
    - Symbol
    - Current Price (USD)
    - Market Capitalization (USD)
    - 24-hour Trading Volume (USD)
    - Price Change (24h, %)
  - The file is updated every 2 minutes with the latest data.

## Conclusion

The script provides live, continuous updates on cryptocurrency market data, stored in an Excel file for easy analysis and reference.
