# Coffee_Shop
An Excel base sales performance analysis for insight and maximation of profits.
### An Excel-based sales analysis and dashboard project built on 149,116 transactions from a three-location coffee shop business in New York City (January to June 2023). The project covers data preparation, PivotTable analysis, dashboard design and business recommendations.


## 1. PROJECT OVERVIEW.

This project analyses six months of point-of-sale data from a coffee shop business operating three stores in Manhattan and Queens: **Lower Manhattan, Hell's Kitchen and Astoria**. Raw transaction records were cleaned and structured in Excel, summarised with PivotTables, and presented in an interactive dashboard that shows when, where and what customers buy.

| Item|Detail|
| Period covered| 1 January 2023 to 30 June 2023 (181 days) |
Records analysed |149,116 transactions |
| Store | 3 |
|Product | 80 product items across 9 categories |
| Tool | Microsoft Excel (formulas, Tables, PivotTables, PivotCharts) |


## 2. BUSINESS PROBLEM

Management of a multi-location coffee business needs to understand what drives revenue, but raw transaction logs of nearly 150,000 rows cannot be read directly. Key questions the business could not answer from the raw data alone:

- Which products and categories generate the most revenue?
- At what times of day and on which days is demand highest?
- How is sales performance trending month to month?
- Do the three stores perform equally?

Without answers, decisions on staffing, stock and promotions rely on guesswork.

## 3. PROJECT OBJECTIVES

1. Clean and structure the raw transaction data for analysis.
2. Calculate core sales metrics: total revenue, quantity sold, transaction count and average transaction value.
3. Analyse sales by time (month, weekday, hour), product (category, type) and store.
4. Build a dashboard that presents these insights clearly to non-technical stakeholders.
5. Provide data-driven recommendations on staffing, product focus and promotions.

## 4. DATASET DESCRIPTION
The workbook 'COFFEE_SHOP.xlsx' contains two sheets:

| Sheet | Rows | Columns | Description |
|--------|-----|---------|-------------|
| 'ORIGINAL' | 149,116 | 11 | Untouched raw data, kept for reference |
| 'TABLE ' | 149, 116 | 17 | Working dataset formatted as an Excel Table ('Table3') with six calculated columns added | 

Data quality checks on the dataset found **no missing values, no duplicate rows and unique transaction IDs**. Calculated `Revenue` matches quantity C  unit price on every row.

## 5. DATA DICTIONARY

| Column | Type | Description |
|---|---|---|
| `Transaction_ID` | Integer | Unique ID for each transaction line |
| `Transaction_Date` | Date | Date of sale (1 Jan to 30 Jun 2023) |
| `Transaction_Time` | Time | Time of sale (hh:mm:ss) |
| `Transaction_Quantity` | Integer | Number of units sold in the transaction |
| `Store_ID` | Integer | Store identifier (3, 5, 8) |
| `Store_Location` | Text | Lower Manhattan, Hell's Kitchen, Astoria |
| `Product_ID` | Integer | Unique product identifier (80 products) |
| `Unit_Price` | Decimal | Price per unit, ranging from 0.80 to 45.00 |
| `Product_Category` | Text | High-level group (9 categories, e.g. Coffee, Tea, Bakery) |
| `Product_Type` | Text | Sub-group within a category (29 types) |
| `Product_Details` | Text | Specific product name and size |
| `Revenue` *(calculated)* | Decimal | `Unit_Price C  Transaction_Quantity` |
| `Month` *(calculated)* | Integer | Month number (1 to 6) |
| `Weekday_Name` *(calculated)* | Text | Day name (Monday to Sunday) |
| `Month_Nmae` *(calculated)* | Text | Intended as month name (see [Data Limitations](#14-data-limitations)) |
| `Hour` *(calculated)* | Integer | Hour of the day (6 to 20) |
| `Weekday` *(calculated)* | Integer | Intended as weekday number (see [Data Limitations](#14-data-limitations)) |


## 6. DATA PREPARATION

1. **Preserved the raw data.** The original data was kept on its own sheet (`ORIGINAL`) and all work was done on a copy (`TABLE`).
2. **Standardised column names.** Lowercase headers (e.g. `transaction_qty`, `product_detail`) were renamed to a consistent capitalised format.
3. **Set correct data types.** Date and time values were stored as Excel serial numbers and were formatted as proper dates and times.
4. **Converted the range to an Excel Table.** This allows structured references, automatic formula fill and reliable PivotTable refresh.
5. **Created calculated columns** used by the analysis:

| Column | Formula |
|---|---|
| `Revenue` | `=H2*D2` |
| `Month` | `=MONTH(B2)` |
| `Weekday_Name` | `=TEXT(B2,"dddd")` |
| `Hour` | `=HOUR(C2)` |

6. **Validated the data.** Checked for blanks and duplicates, and confirmed that `Revenue`, `Month`, `Hour` and `Weekday_Name` agree with the source date, time, price and quantity fields.


## 7. EXCEL ANALYSIS

Analysis followed a top-down approach: overall performance first, then breakdowns by time, product and location.

- **Overall metrics:** total revenue, units sold, transaction count, average transaction value
- **Time analysis:** revenue and quantity by month, weekday and hour of day
- **Product analysis:** revenue and quantity by product category and product type
- **Store analysis:** revenue, transactions and average transaction value by location

Excel functions and features used: `SUM`, `MONTH`, `HOUR`, `TEXT`, calculated Table columns, PivotTables and PivotCharts.

## 8. PIVOT_TABLES

Each PivotTable summarises one question and feeds one dashboard chart.

| PivotTable | Rows | Values | Business question |
|---|---|---|---|
| Revenue by Month | Month | Sum of Revenue | How is revenue trending? |
| Quantity by Month | Month | Sum of Quantity | Is volume growing with revenue? |
| Sales by Weekday | Weekday_Name | Revenue / Quantity | Which days are busiest? |
| Sales by Hour | Hour | Revenue / Quantity | When is peak demand? |
| Sales by Product Category | Product_Category | Revenue / Quantity | Which categories lead? |
| Sales by Product Type | Product_Type | Revenue / Quantity | Which product types lead? |

TOTAL REVENUE BY MONTHS
<img width="527" height="338" alt="image" src="https://github.com/user-attachments/assets/5bda719b-43d3-41a9-80ab-f2c9a24f9c3d" /> 

TOTAL REVENUE BY WEEKDAY

<img width="561" height="337" alt="image" src="https://github.com/user-attachments/assets/858fc59a-828c-4c24-a1e7-8ed3cbbdab25" />

REVENUE BY PRODUCT CATEGORY

<img width="626" height="342" alt="image" src="https://github.com/user-attachments/assets/76cc16fc-6db2-43e9-a503-2f2b6830b7e6" />


WEEKDAY WITH TRANSACTIONS

<img width="537" height="372" alt="image" src="https://github.com/user-attachments/assets/46cb74fd-a374-444c-ace8-dcbc37e5b186" />


REVENUE BY PRODUCT_CATEGORY
<img width="563" height="369" alt="image" src="https://github.com/user-attachments/assets/511f92b1-8a63-42ca-90b6-dac1209e6956" />

HOURS WITH TRANSACTIONS

<img width="619" height="370" alt="image" src="https://github.com/user-attachments/assets/482d8256-a550-4745-83bc-eae7c5d2bf91" />






## 9. DASHBOARD DEVELOPMENT


1. Built PivotTables from `Table3` so the dashboard updates when data is refreshed.
2. Created a PivotChart from each PivotTable, choosing the chart type to suit the data (line for trends over time, column or bar for category comparisons).
3. Placed KPI summary figures at the top, with charts arranged below in a consistent grid.
4. Applied a single colour scheme, clear titles and axis labels, and removed gridline and chart clutter.
5. Reviewed the dashboard against the project objectives so each chart answers a defined business question.


## 10. DASHBOARD FEATURES

- Six charts covering revenue and quantity by **month, weekday, hour, product category and product type**
- KPI summary cards for headline metrics
- Charts linked to PivotTables that refresh when the source data changes
- Consistent layout and colour scheme for easy reading

## 11. KEY PERFORMANCE INDICATORS


| KPI | Value |
|---|---|
| Total Revenue | **$698,812** |
| Total Units Sold | **214,470** |
| Total Transactions | **149,116** |
| Average Transaction Value | **$4.69** |
| Average Units per Transaction | **1.44** |
| Average Daily Revenue | **$3,861** |


## 12. KEY FINDINGS

**Overall growth**
- Monthly revenue roughly doubled over the period, rising from **$81,678 in January to $166,486 in June** (about +104%). June was the strongest month at 24% of total revenue.

| Month | Revenue | Units |
|---|---|---|
| January | $81,678 | 24,870 |
| February | $76,145 | 23,550 |
| March | $98,835 | 30,406 |
| April | $118,941 | 36,469 |
| May | $156,728 | 48,233 |
| June | $166,486 | 50,942 |

**Products**
- **Coffee** is the top category with $269,952 (39% of revenue), followed by **Tea** at $196,406 (28%). Together they make up about two thirds of revenue.
- **Bakery** (12%) and **Drinking Chocolate** (10%) follow. The remaining five categories (coffee beans, branded goods, loose tea, flavours and packaged chocolate) contribute about 11% combined.
- **Barista Espresso** is the single best product type ($91,406), ahead of Brewed Chai tea ($77,082), Hot chocolate ($72,416) and Gourmet brewed coffee ($70,035).
- Syrup flavours sell many units (10,511) but earn little revenue ($8,409) because of their low unit price.

**Time of day**
- Demand peaks in the morning. Hours **7am to 10am generate about 46% of all revenue**, with 10am the busiest hour ($88,673).
- From noon to 5pm, revenue is flat at about $40,000 per hour. Sales fall after 6pm and are very low at 8pm ($2,936).

**Day of week**
- Sales are very evenly spread across the week. Monday is highest ($101,677) and Saturday lowest ($96,894), a gap of under 5%.

**Stores**
- The three stores perform almost equally: Hell's Kitchen $236,511 (34%), Astoria $232,244 (33%), Lower Manhattan $230,057 (33%).
- Lower Manhattan has the fewest transactions (47,782) but the highest average transaction value ($4.81, against $4.66 in Hell's Kitchen and $4.59 in Astoria).


## 13. RECOMMENDATIONS

1. **Staff for the morning rush.** Concentrate the strongest staffing between 7am and 10am, where nearly half of revenue is earned. Ensure espresso and brewed-coffee stock is full before opening.
2. **Run targeted promotions in quiet periods.** Hours after 6pm and the early afternoon are the weakest. Test offers such as a bundled afternoon drink and pastry deal.
3. **Increase pastry attach rate.** Bakery is 12% of revenue and the average order is only 1.44 units. Pairing bakery items with morning coffee orders is the clearest way to raise order value.
4. **Grow retail products selectively.** Coffee beans and branded goods have high unit prices but low volume. Better in-store display and bundling could be tested, especially at Lower Manhattan, where average transaction value is highest.
5. **Plan inventory for continued growth.** Revenue doubled in six months. Use the monthly trend to plan stock and staffing ahead of demand, and check whether the trend reflects seasonality (e.g. summer iced drinks) by collecting a full year of data.
6. **Review low-performing products.** Products and sizes with consistently low sales could be simplified to reduce waste and complexity.


## 14. DATA LIMITATIONS

1. **Short time window.** Only six months (Jan to Jun 2023) are covered, so seasonality and year-on-year growth cannot be assessed. The January to June increase may be seasonal rather than a lasting trend.
2. **No cost or profit data.** Analysis is based on revenue only. High-revenue products may not be the most profitable.
3. **No customer data.** Customer IDs are not recorded, so repeat purchases, loyalty and customer segments cannot be analysed.
4. **Transaction-line granularity.** Each row is one product line, not a full customer basket, so "average transaction value" here is the average per line item and understates the true basket size.
5. **No external context.** Promotions, weather, holidays and store events are not recorded and may explain some of the patterns.
6. **Known formula errors in the working table (to be corrected):**
   - `Month_Nmae` uses `=TEXT(2,"mmmm")`, a hard-coded value that returns "January" for every row. It should be `=TEXT(B2,"mmmm")`. Use the numeric `Month` column in the meantime.
   - `Weekday` uses `=WEEKDAY(A2,2)`, which points at `Transaction_ID` instead of the date, producing meaningless values. It should be `=WEEKDAY(B2,2)`. The `Weekday_Name` column is correct and should be used for weekday analysis.
   - The column header `Month_Nmae` also contains a spelling error and should read `Month_Name`.

## 15. CONCLUSION

The analysis shows a coffee business that is growing fast, driven by coffee and tea, with a strong morning peak and consistent performance across all three stores and all days of the week. The clearest opportunities lie in managing the morning rush, lifting sales in quieterhours, and increasing order value through bakery and retail add-ons. Extending the dataset to a full year and adding cost data would allow the next phase of analysis to move from revenue to profitability and seasonality.


## 16. TOOLS AND SKILLS

**Tools**
- Microsoft Excel (Tables, PivotTables, PivotCharts)
- GitHub (project documentation and version control)

**Skills demonstrated**
- Data cleaning and validation
- Calculated columns and Excel formulas (`MONTH`, `HOUR`, `TEXT`)
- PivotTable and PivotChart analysis
- Dashboard design and data visualisation
- KPI definition and business insight generation
- Technical documentation and business recommendations










