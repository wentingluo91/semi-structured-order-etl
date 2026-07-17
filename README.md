# 📊 Semi-Structured Order ETL Pipeline

A Python-based ETL pipeline designed to extract, transform, and standardize data from semi-structured Excel order forms into analytics-ready datasets.

## 📑 Overview

In many business environments, order information is stored in Excel files with inconsistent layouts and manual formatting. This project demonstrates how Python can automate the extraction and transformation of semi-structured business data.

The pipeline converts raw Excel order forms into structured CSV data for reporting, analysis, and system integration.

## 📌 Workflow


Excel Order Forms
↓
Data Extraction (Python)
↓
Data Cleaning & Standardization
↓
Structured Order Dataset
↓
CSV Output


## 🚀 Features

- Batch processing of multiple Excel order files
- Extraction of customer, order, and cost information
- Handling of semi-structured Excel templates
- Data normalization and cleaning
- Automated order number generation
- Export of structured datasets for analytics

## 🛠 Technologies

- Python
- Pandas
- OpenPyXL
- Regular Expressions
- Excel Data Processing


## 💡 Project Structure

```text
semi-structured-order-etl
│
├── semi_structured_order_etl.py
│
├── sample_data
│   ├── sample_order_standard_01.xlsx
│   ├── sample_order_standard_02.xlsx
│   └── ...
│
└── output
    └── orders.csv
```

## 📖 How to Run

Install required packages:

```bash
pip install pandas openpyxl

Run the ETL pipeline:

python semi_structured_order_etl.py

The processed dataset will be generated as:

orders.csv
Data Privacy

This project is based on a real-world business data processing scenario.
All company-specific information has been removed and replaced with synthetic sample data for demonstration purposes.

## 👤 Author

Wenting Luo
