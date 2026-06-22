# 📊 Sales Performance Dashboard in Excel

## 📌 Project Overview

This project demonstrates end-to-end data analysis and dashboard development using Microsoft Excel. The objective was to transform raw sales data into actionable business insights through data cleaning, data integration, statistical analysis, Pivot Tables, Pivot Charts, and an interactive dashboard.

The dashboard provides a comprehensive view of sales performance, profitability, customer behavior, product performance, regional trends, and discount effectiveness.

---

## 🎯 Project Objectives

- Import and organize raw data into structured Excel tables
- Clean and validate data for analysis
- Remove duplicates and handle missing values
- Integrate multiple datasets using XLOOKUP/VLOOKUP
- Perform descriptive statistical analysis
- Create Pivot Tables for advanced summarization
- Build interactive visualizations and dashboards
- Generate business insights and recommendations

---

## 📂 Dataset Structure

The project consists of four relational tables.

### 1. Customers Table

| Customer_ID | Name | Age | Gender | City | State | Country | Loyalty_Level |
|------------|------|-----|--------|------|--------|---------|---------------|
| CUST1000 | Sheena Bonilla | 55 | Male | Donaldton | North Carolina | USA | Platinum |
| CUST1001 | Eugene Nelson | 26 | Female | Phillipsmouth | North Dakota | USA | Platinum |

**Purpose:**
- Customer segmentation
- Demographic analysis
- Loyalty analysis

---

### 2. Products Table

| Product_ID | Product_Name | Category | Sub_Category | Brand | Cost | Stock |
|------------|-------------|----------|--------------|--------|------|-------|
| PROD0 | Brother Laser Printer | Electronics | Computer | Hunt Inc | 112.42 | 416 |
| PROD1 | Men's Graphic T-Shirt | Clothing | Tops | Terry-Moore | 342.56 | 181 |

**Purpose:**
- Product performance analysis
- Profitability calculations
- Inventory tracking

---

### 3. Stores Table

| Store_ID | Store_Name | Region | City | Store_Type |
|----------|------------|--------|------|------------|
| STORE0 | Rivera-Warner | East | Washingtonfurt | Flagship |
| STORE1 | Martinez, Adkins and Johnson | North | Danielmouth | Flagship |

**Purpose:**
- Regional sales analysis
- Store performance evaluation

---

### 4. Sales Table

| Sales_ID | Order_Date | Customer_ID | Product_ID | Store_ID | Quantity | Unit_Price | Discount | Payment_Type | Total_Amount |
|----------|------------|-------------|------------|----------|----------|------------|----------|--------------|-------------|
| SALE0 | 45480 | CUST1409 | PROD97 | STORE3 | 1 | 218.04 | 0.09 | Credit Card | 198.42 |
| SALE1 | 45285 | CUST1467 | PROD74 | STORE17 | 3 | 104.47 | 0.27 | PayPal | 228.79 |

**Purpose:**
- Revenue analysis
- Sales trend analysis
- Profit calculations

---

## 🧹 Data Cleaning

The following data cleaning techniques were applied:

### ✅ Duplicate Removal
- Identified and removed duplicate records.
- Used Excel's Remove Duplicates feature.

### ✅ Missing Value Handling
- Filled missing loyalty levels.
- Imputed missing cost values where necessary.

### ✅ Data Validation
- Gender validation
- Region validation
- Payment type validation
- Discount range validation

### ✅ Formatting Corrections
- Standardized text formats
- Corrected date formats
- Applied currency formatting

### ✅ Removal of Irrelevant Columns
- Removed unused helper columns
- Removed redundant fields that did not contribute to analysis

---

## 🔗 Data Integration

The datasets were combined using Excel lookup functions.

### XLOOKUP Example

```excel
=XLOOKUP(Customer_ID,
Customers[Customer_ID],
Customers[Loyalty_Level])
```

### Product Lookup

```excel
=XLOOKUP(Product_ID,
Products[Product_ID],
Products[Category])
```

### Store Lookup

```excel
=XLOOKUP(Store_ID,
Stores[Store_ID],
Stores[Region])
```

---

## 🧮 Excel Formulas Used

### Revenue

```excel
=Quantity*Unit_Price*(1-Discount)
```

### Cost

```excel
=Quantity*Cost
```

### Profit

```excel
=Revenue-Cost
```

### Profit Margin

```excel
=Profit/Revenue
```

### Logical Functions

```excel
=IF(Profit>0,"Profit","Loss")
```

### Text Functions

```excel
=TRIM(Name)
=PROPER(Name)
=LEFT(Customer_ID,4)
```

---

## 📈 Descriptive Statistics

Descriptive statistics were calculated using Excel formulas and Analysis ToolPak.

### Metrics Calculated

- Mean
- Median
- Mode
- Standard Deviation
- Maximum
- Minimum

Examples:

```excel
=AVERAGE(Total_Amount)
=MEDIAN(Total_Amount)
=MODE.SNGL(Total_Amount)
=STDEV.P(Total_Amount)
```

---

## 📊 Pivot Table Analysis

Pivot Tables were created to summarize and analyze data.

### Revenue by Region
- Sum of Revenue by Region

### Profit by Region
- Sum of Profit by Region

### Revenue by Loyalty Level
- Revenue grouped by customer loyalty tiers

### Top 10 Products by Revenue
- Highest revenue-generating products

### Top 10 Products by Profit
- Most profitable products

---

## 📉 Dashboard Features

### KPI Cards

- Total Revenue
- Total Profit
- Profit Margin %
- Units Sold
- Average Discount %

### Interactive Slicers

- Region
- Year

### Visualizations

#### 📈 Line Chart
- Revenue Trend Over Time

#### 📊 Clustered Column Chart
- Revenue & Profit by Region

#### 🍩 Doughnut Chart
- Revenue by Age Group

#### 📊 Horizontal Bar Chart
- Top 10 Products by Revenue

#### 📊 Horizontal Bar Chart
- Top 10 Products by Profit

#### 📊 Clustered Column Chart
- Loyalty Impact on Revenue & Profit

#### 📊 Clustered Column Chart
- Discount Impact on Revenue & Profit

---

## 📌 Dashboard KPIs

| KPI | Value |
|------|--------|
| Revenue | $2,174,394.70 |
| Profit | $888,327.80 |
| Profit Margin | 40.90% |
| Units Sold | 4,964 |
| Average Discount | 14.80% |

---

## 💡 Key Findings

- Generated over **$2.17 Million** in revenue.
- Achieved **$888K** in profit.
- Maintained a strong **40.90% profit margin**.
- Customers aged **36–45** contributed the highest revenue.
- Platinum loyalty members generated the highest sales.
- Top-performing products drove a large share of total revenue.
- Moderate discounts improved sales while preserving profitability.

---

## 🛠 Tools Used

- Microsoft Excel
- Excel Tables
- XLOOKUP / VLOOKUP
- Pivot Tables
- Pivot Charts
- Analysis ToolPak
- Conditional Formatting
- Data Validation
- Slicers
- Dashboard Design

---

## 🚀 Conclusion

This project demonstrates how Microsoft Excel can be used as a complete Business Intelligence and Data Analytics tool. Through data cleaning, integration, statistical analysis, Pivot Tables, and dashboard development, meaningful business insights were generated to support data-driven decision-making.
