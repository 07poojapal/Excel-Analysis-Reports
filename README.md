# Excel-Analysis-Reports
This repository contains all the excel reports on sales, financial analysis.

The reports contain the below reports
1. Customer Performance Report
2. Market Performance vs Target Report
3. P&L Statement by FiscalYear
4. P&L Statement by Markets
5. P&L Statement by Months
   
# Tools used : 
             Excel 
             Power Query - To Transform and connect data from different sources such as databases, spreadsheets, into a format suitable for analysis and reporting. 
             Power Pivot 

# Important Functions: 
                     VLOOKUP, XLOOKUP, SUMIF, UNIQUE, IF, INDEX, MATCH, COUNT, COUNTIF
                     
# Statistical Funcations: 
Mean, Median, Mode, Var, Std deviation, Correlation, PMI

## Steps : 
ETL
1. Loaded Dataset into excel(Power Query)
2. Data Cleaning, transformation and modeling done in Power Query -
   1. This inclused spell checking and replacing the correct values/names in power query, replaced incorrect data while exporting to correct one. For example: market region for Canada 
      and USA was nan instead NA. Converted Qty values which were negative to absolute values.
   2. Connected tables using Data Modelling (Fact and Dimension tables - connecting tables mentioned in step 4 below using star schema. That is joining dimension tables with fact tables)
      To verify proper data mdelling:
      I was able to get customer (customer name) from dim_customer into fact_sales_monthly.
      for this i added customer column (formula = Related(dim_customer[customer]))
   3. Created new table dim_date (formula = {Number.From(#date(2018,9,1))..Number.From(#date(2021,8,1))}
      Then added Start of the Month, and Year from the column dervied using above forumla in dim_date.
      Then added Custom Column for Fiscal Month (FY_Month) (formula = Date.AddMonths([month],4)
      Added one more column for Fiscal Year (FY_Year) (formula = Date.Year([FY_Month])
      
   5. Created dim_date using Power Query
4. Used Power Pivot to summarize and analyze large amounts of data quickly and easily.
CSV Files
dim_customer : customer_code, customer, market(country), platform(Ecommerce,Brick Mortar), channel(distributor, direct, retailer) [190 rows]
dim_market: market, subzone, region [24 rows]
dim_product: product_code, division, segment, category, product, variant [299 rows]
fact_sales_monthly: Date (2018 - 2021), product_Code, customer_code, Qty, net_sales_amount [799963 Rows]
ns_target : market, date, ns_target 

## Customer Performance Report 
1. This report shows Netsales for indivial "Customer" for 2019,2020,2021and 2021 vs 2020
   Tool - Pivot Table
   DAX Measure in Power Pivot for 2021 vs 2020 (For this we need NetSales for 2019, 2020 and 2021)
   Formula (NetSales = SUM(fact_sales_monthly[net_sales_amount])
   Get the FY column in fact_sales_monthly to get NetSales in 2019,2020 and 2021 (forumla RELATED(dim_date[FY]))
   For this I have used CALCULATE([Net Sales],dim_date[FY]='2019') (same for 2020 and 2021)
   Then to get 2021 vs 2020 (formula DIVIDE([Net Sales 21],[Net Sales 20],0)

   Used conditional formating, Pivot table design, color formating for showing the databars and color scale
      

## Market Performance Report
1. This report shows Netsales and Target for "Country"/"Market" for 2019, 2020, 2021, 2021-Target, 2021-Target%
   For this report we need to add target into fact_sales_monthly. We can establish relation between ns_target and fact_sales_monthly using diagram view in power pivot.
   target21 (formula SUM(ns_targets_2021[ns_target])
   2021-target21 (formula [NetSales21] - [target21])
   2021-target% (formula DIVIDE([2021-target],[NetSales 21],0)

   Used conditional formating, Pivot table design, color formating for showing the databars and color scale
   India is lagging behind in terms of target.
   Target is missing in almost all the countries as per the data in 2021-target%
   We can see the value lagging by dollars values and %

## Profit and Loss Statement
1. This report shows NetSales, COGS, Gross Margin, GM% for 2019,2020,2021 and 2021 vs 2020
   COGS = Manufacturing Cost + Freight (Transportation) + Other Cost
   Gross Margin = NetSales - COGS
   Gross Margin % = Gross Margin / Net Sales
   fact_sales_monthly with two additional colums (freight_cost and manufacturing_cost)
   Added column Total COGS = fact_sales_monthly[freight_cost]+fact_sales_monthly[manufacturing_cost])
   In Pivot table add DAX (SUM(fact_sales_monthly[Total_COGS])
   Gross Margin = [NetSales - COGs]
   GM% =DIVIDE([Gross Margin],[Net Sales],0)

## Profit and Loss Statement by Month and Quarters
1. For this we need month from the date column. For this I used formula (FORMAT([date],"mmm"))
2. To get Fiscal month number MONTH(DATE(YEAR([date]),MONTH([date])+4,1)
3. To get quarter "Q" & ROUNDUP([fy_month_no]/3,0)
4. Then added Fiscal month and Quarter in the report.

   GM, Sales are higher in Nov and Dec as compared to the other months.
   Net_Sales Comparision
   21 vs 20 (For this we can divide each metric for 21 and 20)
   20 vs 19

   

