
import os
import pandas as pd

print("Step 1: Script started")

input_file = os.path.join("data", "sales_data.csv")
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

print("Step 2: File path created:", input_file)

df = pd.read_csv(input_file, header=None)
print("Step 3: Raw file read successfully")

df = df[0].str.split(",", expand=True)
print("Step 4: Data split into columns")

df.columns = [
    "Order ID", "Order Date", "Region", "Sales Person",
    "Product", "Category", "Units Sold", "Unit Price",
    "Total Sales", "Payment Method"
]
print("Step 5: Column names assigned")

df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
df = df[df.iloc[:, 0].str.lower() != "order id"]
df = df.reset_index(drop=True)
print("Step 6: Basic cleaning done")

df = df.replace("", pd.NA)
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
df["Units Sold"] = pd.to_numeric(df["Units Sold"], errors="coerce")
df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce")
df["Total Sales"] = pd.to_numeric(df["Total Sales"], errors="coerce")

df["Sales Person"] = df["Sales Person"].fillna("Unknown")
df["Units Sold"] = df["Units Sold"].fillna(0)

df = df.drop_duplicates()
df["Total Sales"] = df["Units Sold"] * df["Unit Price"]
df = df.reset_index(drop=True)

print("Step 7: Data cleaning completed")
print(df.head())

output_file = os.path.join(output_dir, "sales_report.xlsx")
print("Step 8: Preparing summary tables")

summary_df = pd.DataFrame({
    "Metric": ["Total Sales", "Total Orders", "Average Order Value"],
    "Value": [df["Total Sales"].sum(), len(df), df["Total Sales"].mean()]
})

region_df = df.groupby("Region", as_index=False)["Total Sales"].sum()
product_df = df.groupby("Product", as_index=False)["Total Sales"].sum()
salesperson_df = df.groupby("Sales Person", as_index=False)["Total Sales"].sum()

print("Step 9: Writing Excel file")

output_file = os.path.join(output_dir, "sales_report.xlsx")

# if old file exists, remove it first
if os.path.exists(output_file):
    os.remove(output_file)
    
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    region_df.to_excel(writer, sheet_name="Region Report", index=False)
    product_df.to_excel(writer, sheet_name="Product Report", index=False)
    salesperson_df.to_excel(writer, sheet_name="Salesperson Report", index=False)

print("Step 10: Report generated successfully")
print("Output file:", output_file)