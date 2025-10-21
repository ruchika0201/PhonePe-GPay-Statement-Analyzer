# PhonePe-GPay-Statement-Analyzer

A Python tool to parse, analyze, and visualize transaction statements from PhonePe and Google Pay. Supports password-protected PDFs (PhonePe), extracts debit and credit transactions, provides last-30-days and multi-month summaries, top merchants, daily and monthly spending charts, and key financial insights.

## 🌟 Key Features

- **Automatic PDF Detection**: Intelligently detects whether your PDF is from PhonePe or Google Pay
- **Password Protection Support**: Handles password-protected PDFs securely
- **Dual Analysis Modes**: 
  - Single Month Analysis (≤30 days)
  - Multi-Month Analysis (>30 days)
- **Comprehensive Financial Insights**: 20+ different analyses and metrics
- **Interactive Visualizations**: View charts interactively with zoom/pan capabilities
- **Excel Export**: All analysis results and charts in one timestamped Excel file
- **Smart Name Formatting**: Automatically adds spaces to concatenated merchant names (e.g., "MissRUCHIKAPANDE" → "Miss RUCHIKA PANDE")

## 📊 Analysis Features

### Single Month Analysis (Last 30 Days)

#### Summary Statistics
- Total debit/credit amounts
- Average and median transaction values
- Net cash flow (surplus/deficit)
- Largest and smallest transactions
- Average daily spending
- Transaction counts by type

#### Spending Patterns
- **Top Merchants**: Identify where you spend the most (with transaction counts and averages)
- **Spending Categories**: Transactions grouped by amount ranges:
  - Under ₹100
  - ₹100-500
  - ₹500-1000
  - ₹1000-5000
  - Above ₹5000
- **Weekday Analysis**: Discover which days you spend most/least
- **Time of Day Analysis**: Spending patterns by time periods:
  - Morning (5AM-12PM)
  - Afternoon (12PM-5PM)
  - Evening (5PM-9PM)
  - Night (9PM-5AM)

#### Transaction Insights
- **Transaction Frequency**: Average, max, min transactions per day
- **Most Active Days**: Days with highest transaction activity
- **Top Expensive Transactions**: List of your biggest purchases
- **Savings Potential**: 
  - Analysis of small transactions that add up
  - Frequent merchants for subscription optimization

#### Visualizations
- Daily spending bar chart
- Debit vs Credit comparison
- Spending categories dual chart
- Weekday spending patterns
- Time of day distribution

### Multi-Month Analysis

#### Overall Summary
- Total debit/credit across all months
- Net cash flow
- Average monthly and daily spending
- Total transaction counts
- Date range coverage

#### Month-by-Month Breakdown
- **Monthly Spending**: Detailed statistics for each month (total, average, median, max, min)
- **Spending Trends**: Month-over-month changes with percentage calculations
- **Trend Indicators**: Visual indicators (📈 Increase, 📉 Decrease, ➡️ Same)
- **Highest/Lowest Months**: Identify spending peaks and valleys

#### Merchant Analysis
- **Top Merchants (All Time)**: Overall spending leaders with totals, counts, and averages
- **Top Merchants Per Month**: Top 3 merchants for each individual month
- **Biggest Transaction Per Month**: Largest transaction details for each month

#### Comparative Analysis
- **Monthly Comparison**: Side-by-side comparison of key metrics
- **Spending Categories**: Overall category distribution across all months
- **Savings Insights**: Cumulative small transactions and frequent merchant analysis

#### Visualizations
- Monthly debit vs credit comparison (side-by-side bars)
- Spending trend line chart with value labels
- Cumulative spending over time
- Category distribution bar chart
- Overall debit vs credit pie chart

## 🚀 Installation

### Prerequisites
- Python 3.6 or higher
- pip (Python package manager)

### Install Dependencies

```bash
pip install pdfplumber pandas matplotlib xlsxwriter
```

Or use the requirements file:

```bash
pip install -r requirements.txt
```

## 📖 Usage

### Method 1: Command-Line Argument

```bash
python statement_analyser.py /path/to/your/statement.pdf
```

### Method 2: Interactive Input

```bash
python statement_analyser.py
```
Then enter the path when prompted.

### Method 3: Direct Execution (Unix/Mac)

```bash
chmod +x statement_analyser.py
./statement_analyser.py /path/to/your/statement.pdf
```

### Password-Protected PDFs

If your PDF is password-protected (PhonePe PDFs typically are), you'll be prompted:

```
Enter PDF password (press Enter if no password): ••••••••
```

The password input is secure and won't be displayed on screen.

## 📁 Output Files

The script generates multiple output files:

### PNG Image Files
- `daily_spend_last_30_days.png` - Daily spending chart (single month)
- `debit_vs_credit_last_30_days.png` - Debit vs credit comparison (single month)
- `spending_categories.png` - Category analysis chart
- `weekday_spending.png` - Weekday patterns chart
- `time_of_day_spending.png` - Time-based spending chart
- `monthly_debit_vs_credit.png` - Monthly comparison (multi-month)
- `spending_trend.png` - Trend line chart (multi-month)
- `cumulative_spending.png` - Cumulative chart (multi-month)
- `category_distribution.png` - Category distribution (multi-month)
- `debit_credit_ratio.png` - Pie chart (multi-month)

### Excel File
**`statement_analysis_YYYYMMDD_HHMMSS.xlsx`** - Comprehensive Excel workbook containing:

#### Sheets Included:
1. **All Transactions** - Complete transaction data
2. **Summary** - Key financial metrics
3. **Top Merchants** - Merchant spending analysis
4. **Spending Categories** - Category breakdowns
5. **Weekday Spending** - Day-of-week analysis
6. **Time Of Day Spending** - Time-based patterns
7. **Transaction Frequency** - Frequency metrics
8. **Top Expensive Transactions** - Biggest purchases
9. **Savings Insights** - Optimization opportunities
10. **Chart Sheets** - All visualizations embedded as images

For multi-month analysis, additional sheets include:
- Overall Summary
- Monthly Detailed breakdown
- Spending Trends
- Monthly Comparison
- Biggest Transactions per month
- And more...

## 💡 Example Workflow

```bash
$ python statement_analyser.py PhonePe_Statement.pdf
============================================================
PhonePe + Google Pay Statement Analyser
============================================================
Enter PDF password (press Enter if no password): ••••••••

Loading PDF...
PDF opened successfully!

Detecting PDF type...
Detected PDF type: PhonePe

Parsing transactions...

Found 127 transactions

First few transactions:
         Date      Time           Merchant   Type   Amount Transaction_ID  Account
0  2025-06-15  10:30 AM    Amazon India   Debit   1299.00      ABC123XYZ  XX1234
...

Date range: 2025-06-01 to 2025-09-30 (122 days)

============================================================
ANALYSIS RESULTS
============================================================

Running Multi-Month Analysis...

🔍 🔍 🔍 ... (analysis output) ... 🔍 🔍 🔍

📊 📊 📊 ... (generating visualizations) ... 📊 📊 📊

======================================================================
SAVING ANALYSIS TO EXCEL
======================================================================
📄 Filename: statement_analysis_20251021_143052.xlsx

✅ Successfully saved analysis!
──────────────────────────────────────────────────────────────────────
📊 File contains:
   • All transactions
   • Summary statistics
   • Analysis tables
   • All visualization charts
──────────────────────────────────────────────────────────────────────
📁 Location: /Users/you/Desktop/StatementAnalyser/statement_analysis_20251021_143052.xlsx
======================================================================

============================================================
Analysis complete!
============================================================
```

## 🔧 Technical Details

### PDF Parsing
- Uses `pdfplumber` for robust PDF text extraction
- Regex-based transaction parsing for both PhonePe and Google Pay formats
- Handles merged keywords and formatting inconsistencies
- Intelligent merchant name cleaning and spacing

### Data Processing
- Pandas DataFrames for efficient data manipulation
- Chronological sorting for multi-month analysis
- Date/time parsing and categorization
- Statistical calculations (mean, median, sum, count)

### Visualizations
- Matplotlib for chart generation
- Interactive display with zoom/pan capabilities
- Professional styling with colors and labels
- In-memory buffer storage for Excel embedding

### Excel Export
- XlsxWriter engine for advanced Excel features
- Automatic sheet name management (handles duplicates)
- Image embedding for charts
- Formatted columns for readability
- Timestamped filenames to prevent overwrites

## 🐛 Troubleshooting

### "No transactions found in the PDF!"
- Verify the PDF is from PhonePe or Google Pay
- Check if the PDF has readable text (not scanned images)
- Ensure the PDF isn't corrupted

### "Incorrect PDF password"
- Double-check your password
- PhonePe PDFs typically use your date of birth (DDMMYYYY) as password

### "Error saving to Excel"
- Ensure `xlsxwriter` is installed: `pip install xlsxwriter`
- Check if you have write permissions in the current directory
- Close any open Excel files with the same name

### Merchant names still concatenated
- This is rare but can happen with unusual formatting
- The script handles most cases automatically
- Check the Excel file for the cleaned data

## 📝 Notes

- The script automatically determines whether to run single-month or multi-month analysis based on date range
- All monetary values are in INR (₹)
- Transaction data is never sent to any external server - all processing is local
- Charts are displayed interactively - close each window to continue to the next
- The Excel file is your complete analysis report - perfect for sharing or archiving

## 🔒 Privacy & Security

- All data processing happens locally on your machine
- No data is transmitted to any external servers
- Password input is secure (not displayed on screen)
- Generated files remain on your local system

## 📄 License

This tool is provided as-is for personal use. Feel free to modify and adapt it to your needs.

## 🤝 Contributing

Suggestions and improvements are welcome! If you encounter any issues or have feature requests, please report them.

## ⚡ Version

**Version 2.0** - Enhanced with comprehensive multi-month analysis, smart name formatting, and improved visualizations.

---

**Happy Analyzing! 📊💰**
