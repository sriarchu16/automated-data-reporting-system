# Automated Data Reporting System

This project automates the process of cleaning, analyzing, and generating reports from raw sales data using Python.

## Features:
-Data cleaning using Pandas
-Handling missing values and inconsistent formats
-Duplicate removal and data validation
-Business metric calculations (Total Sales, Average Order Value)
-Automated Excel report generation with multiple sheets

## Technologies Used:
-Python
-Pandas
-OpenPyXL

## Project Structure:
automated-data-reporting-system/
│
├── data/              # Raw input data
├── output/            # Generated reports
├── src/
│   └── main.py        # Main script
└── README.md

## How to Run:
-pip install pandas openpyxl
-python src/main.py

## Output:
The script generates an Excel file with:
-Cleaned Data
-Summary Report
-Region-wise Sales
-Product-wise Sales
-Salesperson-wise Report