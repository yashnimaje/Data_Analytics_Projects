import pandas as pd
import numpy as np
import os
import sys

try:
    import openpyxl
except ImportError:
    print("[SYSTEM NOTICE] 'openpyxl' dependency missing. Auto-installing environment packages...")
    os.system(f'{sys.executable} -m pip install openpyxl')

input_file = "C:/Users/yashn/OneDrive/Desktop/Data Analytics/Task-1/raw_sales_data.xlsx"
output_file = "C:/Users/yashn/OneDrive/Desktop/Data Analytics/cleaned_sales_data.xlsx"

print("--- Starting Data Wrangling & Cleaning Pipeline ---")
print(f"Connecting to source: {input_file}\n")

if not os.path.exists(input_file):
    raise FileNotFoundError(f"❌ Error: Cannot find file at path: '{input_file}'. Ensure the name matches perfectly.")

print("Processing Worksheet [Raw_Sales_Data]...")
df_sales = pd.read_excel(input_file, sheet_name="Raw_Sales_Data")

initial_rows = len(df_sales)
df_sales = df_sales[df_sales['Date'].astype(str).str.strip() != "00;00;00;00"]
print(f" -> Timecode Filter: Removed {initial_rows - len(df_sales)} corrupt '00;00;00;00' rows.")

sales_dedup_rows = len(df_sales)
df_sales.drop_duplicates(inplace=True)
print(f" -> Deduplication: Extracted {sales_dedup_rows - len(df_sales)} identical duplicate rows.")

df_sales['Gender'] = df_sales['Gender'].astype(str).str.strip().str.title().replace({'M': 'Male', 'F': 'Female'})
df_sales['Region'] = df_sales['Region'].astype(str).str.strip().str.title().replace({'Central!': 'Central', 'Soutth': 'South', 'Others': 'Other'})
df_sales['Category'] = df_sales['Category'].astype(str).str.strip().str.title()
df_sales['Product Name'] = df_sales['Product Name'].astype(str).str.strip()

df_sales['Sales'] = pd.to_numeric(df_sales['Sales'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')
df_sales['Quantity'] = pd.to_numeric(df_sales['Quantity'].astype(str).str.replace(' pcs', '', case=False), errors='coerce')

df_sales['Quantity'] = df_sales['Quantity'].fillna(df_sales['Quantity'].median())
df_sales['Sales'] = df_sales['Sales'].fillna(df_sales['Sales'].median())

df_sales = df_sales[df_sales['Quantity'] >= 0]
df_sales = df_sales[df_sales['Sales'] <= df_sales['Sales'].quantile(0.99)]

df_sales['Date'] = pd.to_datetime(df_sales['Date'].replace('missing', np.nan), errors='coerce', format='mixed').ffill()
print(" -> Finished processing Sales Data successfully.")

print(f"\nExporting results back to spreadsheet structure...")

with pd.ExcelWriter(output_file, 
                    engine='openpyxl', 
                    date_format='DD-MM-YYYY', 
                    datetime_format='DD-MM-YYYY') as writer:
    df_sales.to_excel(writer, sheet_name="Cleaned_Sales_Data", index=False)

print(f"\n--- Data Cleaning Complete And Target saved to: {output_file} ---")