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
                     VLOOKUP 
                     XLOOKUP
                     SUMIF
                     UNIQUE
                     IF
                     INDEX 
                     MATCH
                     COUNT
                     COUNTIF

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
      
   4. Created dim_date using Power Query
4. Used Power Pivot to summarize and analyze large amounts of data quickly and easily.
CSV Files
dim_customer : customer_code, customer, market(country), platform(Ecommerce,Brick Mortar), channel(distributor, direct, retailer) [190 rows]
dim_market: market, subzone, region [24 rows]
dim_product: product_code, division, segment, category, product, variant [299 rows]
fact_sales_monthly: Date (2018 - 2021), product_Code, customer_code, Qty, net_sales_amount [799963 Rows]

## Customer Performance Report 
1. This report shows Netsales for indivial "Customer" for 2019,2020,2021and 2021 vs 2020

## Market Performance Report
1. This report shows Netsales and Target for "Country"/"Market" for 2019, 2020, 2021, 2021-Target, 2021-Target%
   

