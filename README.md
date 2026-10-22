<div align="center">  

# 📊 Excel Sales Analysis & Business Intelligence Dashboard

### 🧮 From Raw Data to Actionable Insights with Microsoft Excel

**An Excel-based analytics project focused on customer behavior, sales performance, profitability, and business insights.**

</div>

---

## 📌 Project Overview

This project uses Microsoft Excel to transform raw sales data into useful business insights. It combines formulas, Conditional Formatting, What-If Analysis, the Analysis ToolPak, PivotTables, charts, and a dashboard to explore sales and profitability.

> **Project status:** Analysis workflow documented. Add actual results, screenshots, and data-backed findings after completing the workbook.

## 🎯 Project Objectives

- 🏆 Identify the top 10 customers by total purchase.
- 💸 Analyze how changing discount rates may affect total profit.
- 📈 Run Linear Regression to study the relationship between Sales and Profit.
- 📊 Generate Descriptive Statistics for important numeric fields.
- ↕️ Show month-over-month sales growth with up/down arrows.
- 🕒 Add a timestamp using `NOW()`.
- 🔎 Identify high-value customers with `INDEX()`, `MATCH()`, and filters.
- 🌍 Compare total sales by region and product using a PivotTable.
- 📉 Build bar, line, and pie charts to visualize KPIs.
- 🧭 Summarize key findings in an interactive dashboard.

## 🛠️ Tools & Excel Features

| Tool / Feature | Purpose |
|---|---|
| 🟩 Microsoft Excel | Main data analysis and reporting tool |
| 🎨 Conditional Formatting | Highlight top customers and important changes |
| 🧪 What-If Analysis | Compare discount scenarios and their profit impact |
| 📐 Analysis ToolPak | Run Regression and Descriptive Statistics |
| 🧮 Excel formulas | Calculate customer totals, growth, and KPIs |
| 🔍 `INDEX()` + `MATCH()` | Return customer details from summarized data |
| 🧱 PivotTables | Summarize sales by region and product |
| 📊 Charts | Visualize sales trends and category comparisons |
| 🖥️ Dashboard | Present KPIs and findings in one place |

## 📂 Analysis Tasks

### 1. 🏆 Top 10 Customers — Conditional Formatting
Summarize total purchases by customer, then use **Home → Conditional Formatting → Top/Bottom Rules → Top 10 Items** to highlight the highest customer totals.

### 2. 💸 What-If Analysis — Discount vs. Profit
Use **Data → What-If Analysis** to test different discount assumptions. A Data Table can compare scenarios; Goal Seek can find an input that reaches a target profit. Ensure the workbook's profit formula correctly reflects the discount.

### 3. 📈 Linear Regression — Profit vs. Sales
Use **Data → Data Analysis → Regression**.

- **Input Y Range:** Profit
- **Input X Range:** Sales
- Tick **Labels** if the selected ranges include headers.
- Review R-squared, coefficients, p-values, and residuals.

Regression can show an association between Sales and Profit; it does not by itself prove that sales changes cause profit changes.

### 4. 📊 Descriptive Statistics
Use **Data → Data Analysis → Descriptive Statistics** on relevant numeric columns such as Sales, Profit, Discount, and Quantity. Select **Summary statistics** and choose an output location.

### 5. ↕️ Monthly Sales Growth Arrows
Calculate month-over-month growth:

```excel
=IFERROR((B3-B2)/B2,0)
```

Format the result as a percentage, then apply **Home → Conditional Formatting → Icon Sets**. Adjust the rule thresholds so increases show an up arrow and decreases show a down arrow. The formula returns 0 if the previous month's sales are zero; review such cases separately if needed.

### 6. 🕒 Timestamp with `NOW()`
Use:

```excel
=NOW()
```

`NOW()` returns the current date and time and updates when Excel recalculates. To preserve a fixed timestamp, copy the result and use **Paste Special → Values**.

### 7. 🔎 High-Value Customers — `INDEX()`, `MATCH()`, and Filters
After creating a customer summary, use this example to return the customer name associated with the highest total purchase:

```excel
=INDEX(A2:A100,MATCH(MAX(B2:B100),B2:B100,0))
```

Here, `A2:A100` contains customer names and `B2:B100` contains their total purchases. Adjust the ranges to match your workbook. Use filters or a defined business threshold to identify additional high-value customers.

### 8. 🌍 PivotTable — Sales by Region and Product
1. Select the dataset, including column headers.
2. Choose **Insert → PivotTable**.
3. Place **Region** in Rows.
4. Place **Product** in Columns.
5. Place **Sales** in Values and summarize by **Sum**.
6. Format the values as currency; add slicers if useful.

### 9. 📊 KPI Charts
Create charts from summarized data:

- **Bar/Column Chart:** Compare sales by region, product, or customer.
- **Line Chart:** Track monthly sales and growth.
- **Pie Chart:** Show the share of sales by a small number of regions or product categories.

Use clear titles, readable labels, and consistent number formats.

### 10. 🖥️ Dashboard — KPIs & Insights
Create a dedicated worksheet named `Dashboard`.

**Suggested KPI cards**
- 💰 Total Sales
- 💵 Total Profit
- 🧾 Total Orders
- 👥 Total Customers
- 🏷️ Average Discount
- 📈 Profit Margin

**Suggested visuals**
- Monthly sales trend
- Sales by region
- Sales by product
- Top 10 customers
- Monthly growth indicators

**Dashboard guidelines**
- Keep the layout clean and easy to scan.
- Use PivotCharts and slicers for interactive analysis where appropriate.
- Add a short insights section supported by actual results.
- State the date range and filters used.
- Do not add conclusions or numbers that have not been verified in the dataset.

## 🗂️ Recommended Workbook Structure

```text
📁 Excel-Sales-Analysis/
├── 📄 README.md
└── 📊 Sales_Analysis.xlsx
    ├── Raw Data
    ├── Customer Analysis
    ├── What-If Analysis
    ├── Regression
    ├── Descriptive Statistics
    ├── Pivot Analysis
    └── Dashboard
```

## 📋 Data Requirements

For the full workflow, the dataset should ideally include:

`Order ID` · `Order Date` · `Customer` · `Region` · `Product` · `Sales` · `Profit` · `Discount` · `Quantity`

Column names may differ. Map the formulas and analysis inputs to the fields in your actual dataset.

## 🚀 How to Use This Project

1. Open the Excel workbook and review the source data.
2. Keep a clean copy of the raw data.
3. Complete each analysis task and place outputs on the recommended worksheets.
4. Refresh PivotTables and charts after updating the data.
5. Check formulas and validate outputs.
6. Add screenshots and a concise summary of verified findings to this README.

### ⚙️ Enable the Analysis ToolPak

If **Data Analysis** is missing from the Data tab:

1. Go to **File → Options → Add-ins**.
2. At the bottom, choose **Excel Add-ins** and click **Go**.
3. Tick **Analysis ToolPak** and click **OK**.

Menu names may vary slightly by Excel version.

## 📦 Deliverables

- ✅ Excel workbook with analysis worksheets
- ✅ Top 10 customer highlighting
- ✅ Discount What-If comparison
- ✅ Regression output and Descriptive Statistics
- ✅ Monthly sales growth indicators
- ✅ Region/product PivotTable
- ✅ KPI charts and dashboard
- 📝 Written insights based on the actual dataset

## 📌 Key Takeaway

This project demonstrates how Excel can be used for **data cleaning, analysis, visualization, and business reporting**—turning raw sales records into clear information that supports better decisions.

---
## 👨‍💻 Author

## Kush Kumar ##


<div align="center">

**Made with 💚 Microsoft Excel | 📊 Data Analysis | 💡 Business Insights**

</div>

