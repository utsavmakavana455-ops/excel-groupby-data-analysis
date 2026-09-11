# Excel GROUPBY – Data Analysis Practice

## 📌 Overview

This project is a practical exercise focused on learning and applying the **GROUPBY function in Microsoft Excel**.

The dataset contains sales transactions with information about regions, salespersons, products, categories, quantities, and sales amounts.

The goal of this project is to understand how Excel can be used to **group, summarize, and analyze data efficiently without relying on PivotTables**.

## 📊 Dataset

The dataset includes the following columns:

* Date
* Region
* Salesperson
* Product
* Category
* Quantity
* Sales

## 🛠️ Excel Skills Practiced

* GROUPBY function
* SUM
* AVERAGE
* COUNT
* Data aggregation
* Grouping data
* Sales analysis
* Product analysis
* Regional analysis
* Salesperson performance analysis

## 🔍 Analysis Performed

### 1. Total Sales by Region

```excel
=GROUPBY(B2:B21,G2:G21,SUM)
```

This groups the sales data by region and calculates total sales for each region.

### 2. Total Quantity by Product

```excel
=GROUPBY(D2:D21,F2:F21,SUM)
```

This calculates the total quantity sold for each product.

### 3. Average Sales by Salesperson

```excel
=GROUPBY(C2:C21,G2:G21,AVERAGE)
```

This calculates the average sales value for each salesperson.

### 4. Total Sales by Category

```excel
=GROUPBY(E2:E21,G2:G21,SUM)
```

This summarizes total sales for each product category.

### 5. Number of Transactions by Region

```excel
=GROUPBY(B2:B21,G2:G21,COUNT)
```

This counts the number of sales transactions in each region.

## 🎯 Learning Objectives

Through this project, I practiced:

* Understanding the purpose of GROUPBY in Excel
* Summarizing large datasets
* Applying aggregation functions
* Extracting useful business insights
* Building formula-based analysis
* Improving Excel data-analysis skills

## 💡 Key Learning

The basic structure I learned is:

```text
GROUPBY → Group the data → Select values → Apply calculation
```


This makes it easier to transform raw transaction data into meaningful summaries.

## 🚀 Future Improvements

I plan to extend this project by practicing:

* PIVOTBY
* PivotTables
* XLOOKUP
* INDEX + MATCH
* VLOOKUP
* IF / IFS
* SUMIFS / COUNTIFS
* Data Cleaning
* Excel Dashboards
* Power Query
* Power BI

---

**Tools:** Microsoft Excel
**Focus:** Data Analysis & Excel Functions


