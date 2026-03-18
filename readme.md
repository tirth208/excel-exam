# Mercedes Stock Analysis Dashboard

**Author:** Tirth Virani

## Project Overview
This project features a comprehensive Excel-based data analysis and dashboard designed to track and analyze the historical stock performance of Mercedes. Focusing primarily on the late 1990s (1996–1999), this dataset provides daily trading metrics alongside aggregated summaries to help visualize market trends, calculate profit/loss margins, and analyze trading volumes.

## Included Files and Data Structure
The project is organized into several key datasets exported as CSV files:

* **`Mercedes_Stock.csv` (Main Dataset):** The core dataset containing daily historical stock records.
  * **Key Metrics:** `Date`, `Open`, `High`, `Low`, `Close`, `Adj Close`, and `Volume`.
  * **Custom Calculations:** Features built-in logic (including nested functions and VLOOKUPs) to evaluate daily performance and automatically categorize each trading day as a `Profit` or `Loss`.

* **`Sheet1.csv` (Trend Data):**
  A focused subset of the main dataset tracking specific daily indicators (`Open`, `Volume`, and `Close`). This data is optimized for feeding simplified charts and visualizing trendlines.

* **`Sheet2.csv` (Aggregated Pivot Data):**
  Contains time-based aggregations summarizing overall market activity.
  * Groups data hierarchically by **Year** (1996–1999) and **Month**.
  * Summarizes crucial metrics: `Sum of High`, `Sum of Volume`, `Sum of Close`, and `Sum of Open`.

* **`Sheet3.csv` (Dashboard Totals):**
  Houses high-level summary metrics and grand totals (e.g., total trading volume of 298,169,517). These figures power the top-line KPI (Key Performance Indicator) cards on the main visual dashboard.

## 📷Excel Preview.

Excel Preview(DashboardPhoto.png) 

## Key Features
* **📈 Historical Trend Analysis:** Monitor opening and closing price fluctuations over a multi-year period.
* **💹 Automated Profit/Loss Tracking:** Built-in spreadsheet logic to dynamically categorize daily market performance.
* **📊 Volume Insights:** Aggregated trading volume broken down by month and year to highlight periods of significant market activity.

## How to Use
1. Clone or download this repository to your local machine.
2. Open the main Excel file (`tirthvirani_exceldashbord.xlsx`) in Microsoft Excel.
3. Navigate to the main Dashboard tab to view the interactive visualizations.
4. **Updating Data:** If you append or update raw data in the `Mercedes_Stock` tab, go to **Data > Refresh All** in the Excel ribbon. This ensures that the Pivot Tables (Sheet2 and Sheet3) and the main dashboard visualizations update automatically to reflect the newest data.

## Requirements
* Microsoft Excel (2016 or newer recommended for optimal Pivot Table and Dashboard compatibility).
* Alternatively, CSV files can be imported into Python (Pandas), R, or SQL for further programmatic analysis.
