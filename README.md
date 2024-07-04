# Excel Data Modeling and Analysis

## Table of Contents

1. [Project Overview](#project-overview)
2. [Data Import Process](#data-import-process)
   - [Import Sales Data from SQL Server Database](#import-sales-data-from-sql-server-database)
   - [Import Product Data from CSV](#import-product-data-from-csv)
   - [Import Store Data from PDF](#import-store-data-from-pdf)
   - [Create a Reference Dates Table](#create-a-reference-dates-table)
3. [Establish Relationships](#establish-relationships)
4. [Pivot Table](#pivot-table)
5. [Required KPIs and Data Insights](#required-kpis-and-data-insights)
   - [Total Orders](#total-orders)
   - [Revenue](#revenue)
6. [Interactive Dashboard](#interactive-dashboard)

## Project Overview

This GitHub project demonstrates a comprehensive Excel data analysis effort, focusing on effective data modeling and collection from various sources. The project includes a well-designed data model that captures relationships between entities such as customers, products, and orders. Data collection is achieved through multiple methods: extracting data from an SQL Server database through a connection link, importing supplementary details or historical records from CSV files, and parsing structured data from PDF documents using Excel's PDF tool. The primary tool for data manipulation, visualization, and analysis is Excel. Explore the repository for insights! 😊

## Data Import Process

To import data into Excel for this project, follow these steps:

### Import Sales Data from SQL Server Database

1. Navigate to the `Data` tab.
2. Select `Get Data`, then `From Database`, and finally `From SQL Server Database`.
3. Establish a connection to the SQL Server database and select the `Sales` table for import.

![pic1](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/2ec6db96-2f10-4086-8709-570a2354ce78)

4. Once you finish, click on the *only create connection* and *add this to the data model* to establish the connection.

![pic2](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/a5e2b219-8abb-4b78-ba55-8bf2dc1ccbf4)

### Import Product Data from CSV

1. Go to the `Data` tab.
2. Choose `Get Data`, then `From Text/CSV`.
3. Establish a connection to the CSV file containing the product data.

### Import Store Data from PDF

1. Navigate to the `Data` tab.
2. Select `Get Data`, then `From File`, and finally `From PDF`.
3. Establish a connection to the PDF file containing the store data.

### Create a Reference Dates Table

1. On the Queries & Connections dialog box that appears on the right, click on the Sales and duplicate.

![calendar](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/bc758d0d-f271-45ff-8898-d0e3d03976d6)

2. It opens up a duplicate **Power Query Editor**. On the left, change the name from Sales(2) to Dates.

![change name](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/adecfe2a-60b6-4df0-8db7-19654b9f79dd)

3. On this new table, select the `Order_Dates` column, right-click, and select `Remove other columns`.

![remove other columns](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/84149a1e-dc0c-4bbc-be40-a9d0bf7ed977)

4. On the Home tab, select `Remove Rows` drop-down and choose `Remove Duplicates`, ensuring that you only have dates for the transaction sales, renaming it **Dates** from **Order_Dates**.

5. Go to the **Add New Column** and Dates tab, select `Day` and `Name of Day`, and do the same for `Start of Week`, `Start of Month`, and `Year`, creating a new Calendar Table with those five columns.

![name of day](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/edb9211f-c100-4883-8b32-6e28c4f65f6a)

Step 1

![5 columns](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/bdc17d2a-7422-40da-b4d8-17b08df61a4c)

All five columns

## Establish Relationships

1. Go to the `Data` tab, and in the Data Tools section, click the arrow on `Data Model` and select `Manage Data Model`. It will open the following dialog box as shown in the image below.

![dialog box1](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/938b9741-96dc-4d0b-a51a-7ec53a3a085e)

2. Click on the diagram view as shown in the above image to open the relationships section.

![dialog box2](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/bbe51dd5-97b2-434b-872c-41179950880f)

3. Create the relationships by dragging:
   - `ProductKey` (Products Table) to `ProductKey` (Sales Table)
   - `StoreKey` (Sales Table) to `StoreKey` (Sales Table)
   - `Order_Date` (Sales Table) to `Date` (Calendar Table)

## Pivot Table

1. Go to the `Insert` tab, select `PivotTable`, and then insert from Data Model.

![data model](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/0f162191-c798-4fd7-bc8d-1047271a8a17)

2. This opens a PivotTable with connections to the four tables, instead of from one.

![tables](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/5ab5791d-f3b5-44ae-94b5-bef7516459ef)

## Required KPIs and Data Insights

Once this is done, begin the EDA to answer KPIs and required data insights.

![question](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/26b65d72-e651-4128-a7d4-1eed6beec801)

### Total Orders

1. Go to the `Power Pivot` tab and select `Measures`, then `New Measures`.

![new measure](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/146d922d-ac66-4a41-81b9-3ce02386b67c)

2. Label the new measure as **Total Orders** using the following DAX expression:

```DAX
=DISTINCTCOUNT(Sales[Order_Number])
```

3. Format the DAX to a Whole Number and a 1000 separator as shown below.

![dax1](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/74177546-67c1-4ff9-9b06-054cdcc92f94)

4. The output:
   - Arranged Rows by Category and Subcategory
   - Column by Total Orders

![totalorders](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/fa533037-1333-4941-8635-7c0668576b0f)

### Revenue

1. For **REVENUE**, use the following DAX expression:

```DAX
=SUMX(Sales[Quantity]*RELATED(Products[Unit Price USD]))
```

![REVENUE](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/3d97c7b1-9203-46eb-b644-197384cfd1ab)

#### The SUMX function

1. **SUMX:**
   - SUMX is a DAX function that calculates the sum of a specified expression over a table.
   - In this case, it sums up the result of the expression inside the parentheses.

2. **Sales[Quantity]:**
   - Sales[Quantity] refers to the column named “Quantity” in the “Sales” table.
   - It represents the quantity of items sold.

3. **RELATED(Products[Unit Price USD]):**
   - RELATED is used to follow relationships between tables.
   - Products[Unit Price USD] refers to the “Unit Price USD” column in the related “Products” table.
   - It retrieves the unit price of the product associated with the sales transaction.

4. **Expression:**
   - The expression inside the parentheses multiplies the quantity by the related unit price for each row in the “Sales” table.
   - The result is summed up across all rows.

In summary, this DAX expression calculates the total sales amount by multiplying the quantity sold with the corresponding unit price for each product. It’s commonly used for aggregating data in Power BI or other DAX-supported tools. 😊

## Interactive Dashboard

1. To make it interactive, first organize the data, ensuring it is presented

 by month (Start of Month).

2. Insert a **PivotChart**, a line graph.

![revenuechart](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/c3165bd0-6e2d-404a-be25-24deaa7a3a6d)

3. The Revenue Line Chart.

![chart1](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/1dc1f983-f7d7-4999-b70d-396bb62faa5a)

4. Add another PivotTable:
   - Column is Total Orders (initially created)
   - Rows is Category

5. Sort Total Orders in descending order to start with the highest Total Orders to the lowest.

![desc](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/3d0e9dfc-e03d-40d8-a18e-018b1f415f06)

6. Add a visual slicer.

7. Create the Interactive Dashboard.

![dashboard](https://github.com/leon-madara/Excel_Data_Modelling/assets/147078093/033dffbd-f2cf-42d3-a53b-d36739d34956)

---
