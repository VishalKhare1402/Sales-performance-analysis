# 📊 Sales Performance & Data Quality Analysis

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Tool-Power%20Query-F2C811?style=flat&logo=powerbi&logoColor=black)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)

> An end-to-end sales data analysis project using Excel and Power Query — covering data cleaning, validation, and interactive dashboard creation for business performance tracking.

---

## 📌 Project Overview

This project analyzes regional and category-wise sales performance using structured datasets. The goal was to identify revenue trends, clean inconsistent data, and build an interactive Excel dashboard that enables quick business decision-making.

---

## 🎯 Objectives

- Clean and standardize raw sales data using Power Query
- Identify and resolve data quality issues (duplicates, missing values, inconsistencies)
- Analyze performance across regions and product categories
- Build an interactive dashboard with Pivot Tables and Slicers

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|------|---------|
| Microsoft Excel | Core analysis and dashboarding |
| Power Query | Data cleaning & transformation |
| Pivot Tables | Aggregation and summarization |
| Slicers & Charts | Interactive filtering and visualization |

---

## 📂 Project Structure

```
sales-performance-analysis/
│
├── data/
│   ├── raw_sales_data.xlsx          # Original uncleaned dataset
│   └── cleaned_sales_data.xlsx      # Post Power Query cleaned data
│
├── dashboard/
│   └── sales_dashboard.xlsx         # Final interactive Excel dashboard
│
├── documentation/
│   └── data_cleaning_steps.md       # Step-by-step cleaning log
│
└── README.md
```

---

## 🔍 Key Steps

### 1. Data Cleaning (Power Query)
- Removed duplicate order entries
- Standardized inconsistent region and category names (e.g., `"north"` → `"North"`)
- Handled missing values in revenue and quantity columns
- Corrected data types (dates, currency, integers)
- Trimmed whitespace and fixed text casing across columns

### 2. Data Validation
- Cross-validated totals against summary rows
- Flagged outliers in sales figures using conditional formatting
- Verified foreign key consistency (product IDs, region codes)

### 3. Analysis
- **Regional Performance:** Compared revenue and order volume across North, South, East, West
- **Category-wise Breakdown:** Identified top and bottom performing product categories
- **Monthly Trend Analysis:** Tracked revenue growth/decline month-over-month
- **Discount Impact:** Assessed correlation between discounts applied and margin compression

### 4. Dashboard (Excel)
- Pivot Table with Region × Category cross-tab
- Monthly revenue trend chart (line)
- Top 5 products bar chart
- Interactive Slicers for Region, Category, and Month filters

---

## 📈 Key Insights

- **Top Region:** [North region contributed ~38% of total revenue]
- **Best Category:** [Electronics had highest average order value]
- **Data Quality:** ~12% of records had missing or inconsistent values before cleaning
- **Trend:** Revenue showed consistent growth in Q3 compared to Q1-Q2

---

## 💡 Skills Demonstrated

- Data cleaning and transformation using Power Query
- Handling real-world data quality issues
- Building business-ready dashboards in Excel
- Performance analysis with Pivot Tables and dynamic charts

---

## 📸 Dashboard Preview

> *(Add a screenshot of your Excel dashboard here)*
> To add: Screenshot your dashboard → save as `dashboard_preview.png` → place in `/images/` folder → update this line:
```
![Dashboard Preview](images/dashboard_preview.png)
```

---

## 🚀 How to Use

1. Clone or download this repository
2. Open `data/raw_sales_data.xlsx` to view the original dataset
3. Open `dashboard/sales_dashboard.xlsx` to explore the interactive dashboard
4. Use the Slicers on the dashboard to filter by Region, Category, or Month

---

## 👤 Author

**Vishal Khare**
📧 vishalkhare54@gmail.com
🔗 [LinkedIn](https://www.linkedin.com/in/vishalkhare9bb31436b)
🐙 [GitHub](https://github.com/VishalKhare1402)
