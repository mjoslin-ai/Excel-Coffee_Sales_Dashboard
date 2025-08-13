# Excel-Coffee_Sales_Dashboard

summary

## Objective

objective

## Analysis Steps

### Step 1: XLOOKUP

- From customers and products to orders sheet
- Customer Name: =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)
- Email: =IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0))
- Country: =XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)

### Step 2: INDEX MATCH

- Coffee Type, Roast Type, Size, Unit Price: =INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))
- Note: the $ before D in orders!$D2 for example means D2 is fixed when copied horizontally while orders!I$1 is fixed when copied vertically

### Step 3: Multiplication formula for Sales

- Sales: ==L2*E2

### Step 4: Multiple IF functions

- Create columns Coffee Type Name and Roast Type Name
- Coffee Type Name: =IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica",""))))
- Roast Type Name: =IF(J2="M","Medium",IF(J2="L","Light",IF(J2="D","Dark","")))

### Step 5: Date Formatting

- Order Date: dd-mmm-yyyy

### Step 6: Number Formatting

- Size: 0.0 "kg"
- Unit Price, Sales: $ US

### Step 7: Check For Duplicates

- Data > Data Tools > Remove Duplicates

### Step 8: Convert Range to Table

- Improves managing and refreshing of pivot tables when new data is added

### Step 9: Pivot Tables and Pivot Charts + Formatting

- Create TotalSales pivot table
- Rows: Order Date
- Group by months and years in tabular form
- Columns: Coffee Type Name
- Values: Sum of Sales with 0 decimal places
- Insert line chart: Insert > Charts > Line chart
- Hide field buttons
- Change colors for font, fill, legend, and gridlines
- Add vertical axis (USD) and chart title (Total Sales Over Time)

### Step 10: Insert Timeline + Formatting

- Insert timeline: PivotChart Analyze > Insert Timeline
- New timeline style (Green Timeeline Style)

### Step 11: Insert Slicers + Formatting

- Insert slicer: PivotChart Analyze > Insert Slicer
- One for Roast Type Name and Size

### Step 12: Updating the Pivot Table Data Source + Formatting

- Add Loyalty Card from customers to orders sheet
- Loyalty Card: =XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,,0)
- Add Loyalty Card slicer
- New slicer style (Green Slicer Style)
- Change Roast Type Name Layout to 3 columns and Size slicer to 2 columns
- Copy Total Sales worksheet named Country Bar Chart
- Modify TotalSales pivot table with only Country in Axis and Sum of Sales in Values
- Insert bar chart sorting by Sum of Sales
- Add data labels and put units in US dollars
- Copy Country Bar Chart worksheet named Top 5 Customers
- Modify TotalSales pivot table with Customer Name in Axis
- Filter to show top 5 customers by Sum of Sales
- Sort by Sum of Sales

### Step 13: Building the Dashboard

- Create Dashboard worksheet
- Change column witch and row height to 1 and 5 respectfully
- Add shape for dashboard title (hold ALT to snap to grid)
- Move visuals from Total Sales, Country Bar chart, and Top 5 Customers worksheet (CTRL X and CTRL V)
- Arrange visuals to make an elegant dashboard
- Connect the timeline and slicers to each visual using Report Connections
- Remove gridlines: View > uncheck gridlines
- In File > Options > Advanced uncheck: show formula bar, show sheet tabs, and show row and column headers

## Insights








