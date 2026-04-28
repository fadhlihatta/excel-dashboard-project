# 📊 Sales Data Analysis Dashboard (Excel + VBA)

## 📌 Overview

This project showcases an end-to-end data analytics workflow built entirely in Microsoft Excel, enhanced with VBA automation.

The solution transforms raw CSV data into a clean dataset and presents insights through an interactive dashboard.

---

## ⚙️ Key Features

* 🔄 **Automated Data Cleaning (VBA)**

  * Handles incremental data updates
  * Removes duplicates
  * Standardizes text and data types

* 📊 **Interactive Dashboard**

  * KPI metrics (Revenue, Profit, Quantity Sold)
  * Sales trend analysis (monthly/yearly)
  * Product & category performance
  * Sales channel distribution
  * Payment method insights
  * Delivery time comparison

* 📁 **Structured Data Pipeline**

  * Raw data → Cleaned data → Pivot → Dashboard

---

## 🧱 Tech Stack

* Microsoft Excel
* VBA (Visual Basic for Applications)
* Pivot Tables
* Data Visualization (Charts, KPI Cards)

---

## 🔄 Data Workflow

1. Import raw CSV data into `RawData` sheet
2. Run VBA macro (`CleanData`)
3. Cleaned data stored in `CleanedData`
4. Pivot tables update automatically
5. Dashboard reflects updated insights

---

## 🧠 Data Cleaning Logic

The VBA script performs:

* Trimming text values
* Handling empty and error cells
* Converting data types (date & numeric)
* Removing duplicate records
* Incremental data processing (only new rows)

---

## 📊 Dashboard Preview

![Dashboard](docs/dashboard_preview.png)

---

## 📁 Project Structure

```
excel-sales-dashboard/
│
├── data/
├── excel/
├── vba/
├── docs/
└── README.md
```

---

## 🚀 How to Use

1. Open `sales_dashboard.xlsm`
2. Paste new data into `RawData`
3. Run macro: `CleanData`
4. Refresh Pivot Tables
5. View updated dashboard

---

## 🎯 Business Questions Answered

* What is the total revenue and profit?
* Which products and categories perform best?
* How do sales trends evolve over time?
* What are the dominant sales channels?
* How do payment methods distribute?
* Which distributor has the fastest delivery time?

---

## ⚠️ Notes

* Data used in this project is synthetic (AI-generated)
* This project focuses on demonstrating data workflow and analytics capability

---

## 📌 Future Improvements

* Integrate SQL or Python for scalable data processing
* Add forecasting model for sales prediction
* Migrate dashboard to Power BI for advanced interactivity

---

## 👤 Author

Fadhli Hatta
