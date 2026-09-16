# 📊 PR Final Project — Data Intelligence Dashboard

![Project Badge](https://img.shields.io/badge/Project-Data%20Intelligence%20Dashboard-0F766E?style=for-the-badge)
![Excel Badge](https://img.shields.io/badge/Made%20with-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Status Badge](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

> **A complete Excel-based customer transaction analytics project** with raw data preparation, formula-driven analysis, What-If Analysis, Analysis ToolPak regression, Pivot Tables, Pivot Charts, KPI cards, and storytelling.

---

## 🎥 Learning / Presentation Video

📺 **Excel Dashboard & Data Analysis Tutorial**

[▶️ Watch the video here]
https://drive.google.com/file/d/1UxSqiM1vGSvKH7yrZYgR6KBjhnuHoDOX/view?usp=sharing

---

## 🚀 Project Overview

This project converts customer transaction data into a practical business intelligence workbook.

### 🎯 Objective

- Analyze customer transactions and sales performance
- Identify high-value customers and frequently purchased products
- Compare regional and product-wise sales
- Build KPI cards and visual dashboards
- Apply advanced Excel formulas and analysis tools
- Present business insights through data storytelling

---

## 📁 Workbook Structure

| Sheet | Purpose |
|---|---|
| `Raw data` | Original transaction dataset |
| `Analysis` | Formula-driven calculations and derived columns |
| `WHAT-IF Analysis` | Goal Seek and scenario-style calculations |
| `Analysis Toolpak` | Regression summary and statistical output |
| `Pivot table` | Customer, product, and regional summaries |
| `Charts` | Supporting charts |
| `Dashboards` | KPI cards and visual dashboard |
| `Data storytelling` | Business interpretation and insights |

---

## 📌 Dataset Snapshot

| Metric | Output |
|---|---:|
| Total records | **250** |
| Unique customers | **50** |
| Unique orders | **250** |
| Total quantity sold | **753 units** |
| Total sales | **₹229,192.47** |
| Average order value | **₹916.77** |
| Data period | **11 Apr 2024 – 11 Apr 2025** |

---

## 🧮 Excel Formulas Implemented

> Formula references below are based on the `Analysis` sheet structure.

### 1. 📅 Date & Time Functions

#### Extract month

```excel
=TEXT(B2,"mmm")
```

**Output:** `Apr`

#### Extract year

```excel
=YEAR(B2)
```

**Output:** `2024`

#### Customer age / days since customer joined

```excel
=DATEDIF(M2,TODAY(),"D")
```

**Output:** Number of days since `Customer_Since`.

#### Month-end date

```excel
=EOMONTH(B2,0)
```

**Output:** `30-Apr-2024`

#### Current timestamp

```excel
=NOW()
```

**Output:** Current date and time.

---

### 2. 💰 Sales Classification

```excel
=IF(N2>1000,"High Sale","Low Sale")
```

**Output examples:**

| Total Amount | Result |
|---:|---|
| ₹749.97 | Low Sale |
| ₹1,499.94 | High Sale |

---

### 3. 🧾 Customer Total Spend

```excel
=SUMIFS($N:$N,$D:$D,D2)
```

**Purpose:** Calculates the total amount spent by each customer.

**Output examples:**

| Customer | Total Spend |
|---|---:|
| Mark Carter | ₹15,659.65 |
| Edward Mitchell | ₹11,919.77 |
| Barbara Young | ₹10,649.80 |

---

### 4. 🔁 Customer Purchase Frequency

```excel
=COUNTIF($D$2:$D$171,D2)
```

**Purpose:** Counts how many transactions are associated with a customer.

**Business use:** Helps identify repeat customers.

---

### 5. 🛍️ Most Purchased Product

Example dynamic-array approach:

```excel
=LET(
customer,D2,
products,FILTER($F$2:$F$251,$D$2:$D$251=customer),
counts,COUNTIF(products,products),
INDEX(products,MATCH(MAX(counts),counts,0))
)
```

**Purpose:** Returns the most frequently purchased product for a customer.

> If your Excel version does not support `LET` or `FILTER`, use a Pivot Table or helper-column method.

---

### 6. 🔎 FILTER Function — Multi-value Return

```excel
=FILTER(A2:N251,K2:K251="East","No records found")
```

**Purpose:** Returns all transactions from the East region.

---

### 7. 🧠 High-Value Customer Identification

```excel
=IF(U2>=10000,"High Value Customer","Regular Customer")
```

**Purpose:** Flags customers whose total spend is at least ₹10,000.

> The threshold can be changed according to business requirements.

---

### 8. 🔤 TEXT Functions / Name Abbreviations

```excel
=LEFT(A2,1)&"."&RIGHT(A2,1)
```

**Purpose:** Creates a short identifier from text.

Example:

| Input | Output |
|---|---|
| Rahul | R.l |
| Priya | P.a |

---

### 9. 🔤 Name Matching / Comparison

```excel
=IF(A2=B2,"Match","Not Match")
```

**Purpose:** Compares two names or text values.

---

## 📊 Key Business Outputs

### 🏆 Product-wise Sales

| Rank by sales | Product | Sales |
|---:|---|---:|
| 1 | Laptop | ₹67,499.25 |
| 2 | Smartphone | ₹67,199.04 |
| 3 | Desk | ₹23,399.22 |
| 4 | Monitor | ₹21,999.12 |
| 5 | Bookshelf | ₹15,298.98 |

### 🌍 Regional Sales

| Region | Sales |
|---|---:|
| East | ₹59,288.39 |
| North | ₹50,808.31 |
| West | ₹41,408.68 |
| Central | ₹41,288.34 |
| South | ₹36,398.75 |

### 👑 Top Customers by Total Spend

| Customer | Total Spend |
|---|---:|
| Mark Carter | ₹15,659.65 |
| Edward Mitchell | ₹11,919.77 |
| Barbara Young | ₹10,649.80 |
| Patricia Moore | ₹9,799.73 |
| Paul Baker | ₹8,309.74 |

---

## 🧪 What-If Analysis

### Goal Seek

Goal Seek is used to find the required quantity for a target sales amount.

Example formula:

```excel
=B5*C5
```

Where:

- `B5` = Unit Price
- `C5` = Quantity
- Result = Total

### Example

If the target is ₹5,000:

```text
Required Quantity = Target Amount / Unit Price
```

For a Monitor priced at ₹249.99:

```text
Required Quantity ≈ 20 units
```

---

## 📈 Analysis ToolPak — Regression

The workbook includes a regression summary generated through Excel’s Analysis ToolPak.

### Included outputs

- Multiple R
- R Square
- Adjusted R Square
- Standard Error
- Observations
- Regression coefficients
- ANOVA summary

### Interpretation

Regression can be used to understand the relationship between transaction-related variables and sales amount. The output should be interpreted as an analytical relationship, not automatically as proof of causation.

---

## 📊 Dashboard Components

- KPI Cards
  - Total Sales
  - Total Orders
  - Average Sale
- Product-wise sales chart
- Regional sales chart
- Customer segment visualization
- Pivot-based summaries
- Conditional formatting
- Business storytelling section

---


## 💡 Business Insights

- Laptop and Smartphone contribute the largest product-wise sales.
- East region generates the highest regional sales in this dataset.
- High-value customer analysis helps identify customers for retention campaigns.
- Purchase frequency can support loyalty-program decisions.
- KPI cards make the workbook easier for managers to understand quickly.
- Pivot Tables and charts convert raw transactions into actionable summaries.

---

## 🛠️ Tools & Skills

![Excel](https://img.shields.io/badge/Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Pivot Tables](https://img.shields.io/badge/Pivot%20Tables-Data%20Analysis-blue?style=flat-square)
![Dashboard](https://img.shields.io/badge/Dashboard-Visualization-purple?style=flat-square)
![What If](https://img.shields.io/badge/What--If%20Analysis-Goal%20Seek-orange?style=flat-square)
![Regression](https://img.shields.io/badge/Analysis%20ToolPak-Regression-red?style=flat-square)

- Microsoft Excel
- Excel Tables
- Date & Time Functions
- `IF`, `SUMIFS`, `COUNTIF`, `FILTER`, `TEXT`
- Pivot Tables
- Pivot Charts
- Conditional Formatting
- Goal Seek
- Analysis ToolPak
- Data storytelling

---

## ✅ Final Deliverable Checklist

- [x] Raw transaction data
- [x] Analysis sheet with formulas
- [x] High-value customer logic
- [x] Product and regional summaries
- [x] What-If Analysis
- [x] Regression output
- [x] Pivot Tables
- [x] Charts and dashboard
- [x] Data storytelling
- [x] Dashboard screenshot
- [x] Learning video link

---

## 👨‍💻 Author

**Simran Gohel**

> Built as a practical Excel Data Analytics / Business Intelligence project.

⭐ If you found this project useful, consider starring the repository and sharing your feedback.
