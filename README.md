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



### Step 5: Date Formatting

### Step 6: Number Formatting

### Step 7: Check For Duplicates

### Step 8: Convert Range to Table

### Step 9: Pivot Tables and Pivot Charts + Formatting

### Step 10: Insert Timeline + Formatting

### Step 11: Insert Slicers + Formatting

### Step 12: Updating the Pivot Table Data Source

### Step 13: Building the Dashboard

## 
