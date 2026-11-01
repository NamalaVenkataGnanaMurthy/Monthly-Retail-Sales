# Monthly Retail Sales Analysis & Interactive Dashboard

## Project Overview

This project focuses on analyzing retail sales data and creating an
interactive dashboard using **Microsoft Excel 2024**. The dataset was
cleaned and analyzed to understand sales, profit, product performance,
customer segments, and other business metrics.

The dashboard uses PivotTables, PivotCharts, KPI cards, and month
slicers to present the results clearly.

## Project Details

-   **Project:** Monthly Retail Sales Analysis
-   **Tool Used:** Microsoft Excel 2024
-   **Project Type:** Data Cleaning, Data Analysis, and Dashboard
    Development
-   **Duration:** 2 days

## Objectives

-   Clean the raw retail sales dataset.
-   Remove duplicate records and handle blank values.
-   Correct inconsistent text and category values.
-   Analyze sales, profit, cost, quantity, and customer ratings.
-   Create PivotTables and PivotCharts.
-   Build an interactive dashboard with KPI cards and month slicers.
-   Identify useful business insights from the sales data.

## Dataset Description

The dataset contains 20 columns:

`Order_ID`, `Order_Date`, `Customer_Name`, `Customer_Age`, `City`,
`State`, `Category`, `Product`, `Customer_Segment`, `Quantity`,
`Unit_Price`, `Discount`, `Sales`, `Cost`, `Profit`, `Rating`,
`Payment_Method`, `Sales_Channel`, `Order_Status`, and `Salesperson`.

## Data Cleaning

The original dataset contained **10,031 rows**. Data cleaning was
performed to remove duplicates and handle blank values.

  Cleaning Step                            Remaining Rows
  -------------------------------------- ----------------
  Initial dataset                                  10,031
  After removing duplicates                        10,001
  After deleting blank customer names               9,881
  After deleting blank cities                       9,801
  After deleting blank discounts                    9,761
  After deleting blank ratings                      9,721
  After deleting blank payment methods              9,663

**Final cleaned dataset: 9,663 rows**

### Cleaning activities and Excel functions

-   Removed duplicate records.
-   Checked and handled blank values.
-   Corrected inconsistent text and category values.
-   Standardized date values.
-   Cleaned text using `TRIM`, `CLEAN`, and `PROPER`.
-   Used `COUNTIF` to identify repeated values.
-   Used `IFERROR` and `DATEVALUE` to help handle date values.
-   Used Excel filters and Find and Replace where required.

Example formula:

``` excel
=PROPER(TRIM(CLEAN(A2)))
```

## KPI Cards

Five KPI cards were created:

1.  **Total Sales**
2.  **Total Profit**
3.  **Total Cost**
4.  **Total Quantity**
5.  **Average Rating**

The KPI values were linked to summarized PivotTable cells using cell
reference formulas. Shapes and text boxes were used to display the KPI
values on the dashboard.

## Dashboard Charts

The dashboard contains eight charts:

1.  **Sales by Category --- Clustered Bar Chart:** Compares total sales
    across product categories.
2.  **Profit by State --- Clustered Bar Chart:** Compares profit across
    different states.
3.  **Monthly Sales Trend --- Column Chart:** Displays monthly sales and
    helps identify changes over time.
4.  **Sales by Customer Segment --- Doughnut Chart:** Shows how sales
    are divided among customer segments.
5.  **Top 10 Products by Profit --- Clustered Bar Chart:** Displays the
    ten products with the highest total profit.
6.  **Sales by Salesperson --- Clustered Bar Chart:** Compares sales
    performance across salespeople.
7.  **Order Status Distribution --- Doughnut Chart:** Shows the number
    of orders in each order status.
8.  **Sales by Payment Method --- Clustered Column Chart:** Compares
    sales across different payment methods.

## Interactive Features

-   Month slicers were added to filter the dashboard by month.
-   Slicers were connected to relevant PivotTables and charts using
    Excel's Report Connections feature.
-   Selecting a month updates the connected dashboard visuals and KPI
    values.
-   KPI cards, charts, and slicers are combined in one place for easier
    analysis.

## Key Business Insight

-   **Electronics** recorded the highest sales among the categories,
    with total sales of approximately **₹12.15 crore**.
-   The category sales analysis helps compare the contribution of each
    product category.
-   The dashboard makes it easier to review monthly sales and compare
    performance across products, states, customer segments, salespeople,
    and payment methods.

## Project Outcome

This project provided practical experience in cleaning retail data,
creating PivotTables and PivotCharts, using Excel formulas, and building
an interactive dashboard. It also helped develop skills in presenting
data clearly and identifying business insights from a dataset.

## Skills Practiced

-   Microsoft Excel 2024
-   Data Cleaning
-   Data Analysis
-   PivotTables
-   PivotCharts
-   KPI Cards
-   Slicers
-   Excel Formulas
-   Dashboard Development
-   Business Insights

## Project Files
Repository structure:

``` text
Monthly-Retail-Sales-Analysis/
├── README.md
├── Monthly_Retail_Sales_Dataset_Cleaned_Data.xlsx
├── Monthly_Retail_Sales_Raw_Data.xlsx
├── Monthly_Retail_Sales_Dashboard_Report.pdf
├── Dashboard_Screenshot.png
└── Dashboard_Demo_Video.mp4

```

Replace the example filenames with the actual filenames used in the
repository.
