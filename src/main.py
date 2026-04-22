

from csv import writer
import os
import pandas as pd

#File paths
input_file=os.path.join("data", "sales_data.csv")
output_dir="output"

#Create output folder if not exists
os.makedirs(output_dir, exist_ok=True)

#--------------------------
#Step 1: Read the sales data
#---------------------------

print("Reading sales data...")
#df=pd.read_csv(input_file, sep=",", encoding="utf-8")

#Read raw file as single column first to handle messy data
df=pd.read_csv(input_file,header=None)

#Split into proper columns (single column into multiple columns)
df=df[0].str.split(",", expand=True)

#Set column names manually
df.columns=[ 
    "Order ID", "Order Date", "Region", "Sales Person",
    "Product", "Category", "Units Sold", "Unit Price",
    "Total Sales", "Payment Method"
]

#Remove Spaces
df=df.apply(lambda col: col.str.strip() if col.dtype=="object" else col)

#Remove repeated header row
df=df[df.iloc[:,0].str.lower()!="order id"]

#Reset index
df=df.reset_index(drop=True)

print("\nCleaned Data Preview:")
print(df.head())

#print("\nColumns count:", len(df.columns))
#print("\nDataset Info:")
#print(df.info())

print("\nShape:", df.shape)
print("\nMissing Values:")
print(df.isnull().sum())

print("\nDuplicte Rows:", df.duplicated().sum())

#--------------------------
#Step 2: Data Cleaning
#--------------------------

print("\nStarting Data Cleaning....")

#Treat empty strings as missing values
df=df.replace("",pd.NA)

#1. Convert Order Date to proper format
df["OrderDate"]=pd.to_datetime(df["Order Date"], errors="coerce")

#2. Convert numeric columns
df["Units Sold"]=pd.to_numeric(df["Units Sold"], errors="coerce")
df["Unit Price"]=pd.to_numeric(df["Unit Price"], errors="coerce")
df["Total Sales"]=pd.to_numeric(df["Total Sales"], errors="coerce")

#3. Fill missing values
df["Sales Person"]=df["Sales Person"].fillna("Unknown")
df["Units Sold"]=df["Units Sold"].fillna(0)

#4. Remove duplicate rows
df=df.drop_duplicates()

#5. Recalculate Total Sales
df["Total Sales"]=df["Units Sold"] * df["Unit Price"]

#6. Reset index after cleaning
df=df.reset_index(drop=True)

print("\nAfter Cleaning Data Preview:")
print(df.head())

print("\nFinal Shape:", df.shape)

#-----------------------------
#Step 3: Generate Excel Report
#-----------------------------

print("\nGenerating Excel Report...")

output_file=os.path.join(output_dir,"sales_report.xlsx")

#Summary
total_sales=df["Total Sales"].sum()
total_orders=len(df)
average_order_value=df["Total Sales"].mean()

summary_df=pd.DataFrame({
    "Metric": ["Total Sales", "Total Orders", "Average Order Value"],
    "Value": [total_sales, total_orders, average_order_value]
})

#Region-wise Report
region_df=df.groupby("Region", as_index=False)["Total Sales"].sum()

#Product-wise Report
product_df=df.groupby("Product", as_index=False)["Total Sales"].sum()

#Sales Person Report
salesperson_df=df.groupby("Sales Person", as_index=False)["Total Sales"].sum()

#Write to Excel
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    region_df.to_excel(writer, sheet_name="Region-wise", index=False)
    product_df.to_excel(writer, sheet_name="Product-wise", index=False)
    salesperson_df.to_excel(writer, sheet_name="Sales Person-wise", index=False)

print(f"\nReport generated successfully: {output_file}")


