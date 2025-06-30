## 📊 Equal-Weight S\&P 500 Portfolio Builder

This project calculates how to build an equal-weighted version of the S\&P 500 index using a subset of 200 selected companies, with live stock data pulled from the **Finnhub API**.

Instead of mimicking the market-cap weighted structure of the actual S\&P 500, this approach invests the same amount of money into each company, which is a strategy often preferred by individual investors looking for simplicity and diversification.

---

### 🚀 What This Project Does

* Accepts a user-defined total portfolio value
* Pulls live **stock price** and **market capitalization** for each of the 200 companies
* Calculates how many shares to buy for each company to maintain equal weighting
* Stores the final recommendation in a clean, formatted **Excel spreadsheet**

---

### 🔧 Tech Stack

* **Python**
* **Pandas** (for data processing)
* **Requests** (for API calls)
* **Finnhub API** (for real-time stock data)
* **XlsxWriter** (for Excel formatting)

---

### 📁 Project Workflow

#### 📥 Importing Stock List

I used a static list of 200 companies from the S\&P 500. This list was loaded into a Pandas DataFrame to serve as the working stock universe for the portfolio.

#### 🔐 Acquiring the API Token

I used the [Finnhub API](https://finnhub.io) to retrieve stock data. The API key was stored in a variable directly within the notebook. (Optionally, this could be stored in a `secrets.py` file for better security.)

#### 📡 Making API Calls

For each stock, I made two API calls:

* `/quote` for the latest stock price
* `/stock/profile2` for the company’s market capitalization

Since Finnhub does not support batch API calls via REST, I looped through the 200 tickers individually.

#### 🧮 Parsing the API Response

The raw JSON response from the API was parsed to extract only the required fields: stock price and market capitalization.

#### 📊 Building the DataFrame

Each stock’s data was added to a Pandas DataFrame. The number of shares to buy was calculated by dividing an equal portion of the total portfolio value by the stock price.

#### 📁 Excel Output

The final output was exported to an Excel file named `recommended trades.xlsx`. I used the `xlsxwriter` engine with Pandas' `ExcelWriter` to:

* Create reusable formatting templates (e.g., for currency and integers)
* Apply column widths and styles for better readability

---

### 📄 Example Excel Output

The generated Excel file includes the following columns:

* **Ticker**
* **Stock Price**
* **Market Capitalization**
* **Number of Shares to Buy**

Each column is formatted for clarity, with appropriate widths, currency symbols, and number formatting.

---

### ⚠️ Notes

* The project uses a fixed list of 200 companies instead of dynamically fetching the full S\&P 500.
* Finnhub’s REST API does not support batch quotes or company lookups, so performance depends on API rate limits.
* The API token should be kept private and not uploaded to public repositories.

---

### 📌 Future Improvements

* Add async support for faster API requests
* Allow user to upload their own stock list
* Integrate a full S\&P 500 dataset with dynamic fetching (if access is available)
* Add error handling and logging for API failures

---
