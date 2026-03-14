# Apple Product Sales Data Automation using Power Query

## Overview

This project demonstrates how Microsoft Excel Power Query can automate the process of combining multiple monthly sales reports into a single structured dataset.

In many organizations, analysts receive sales data in separate Excel files every month. Manually cleaning and consolidating these files is time-consuming.

This solution automates the process by connecting to a folder containing multiple Excel files, cleaning the raw data, and producing a structured dataset ready for analysis.

The final output includes a clean sales dataset and a product sales report by month.

---

## Tools Used

* Microsoft Excel
* Power Query
* Folder Data Connection
* Data Transformation Techniques

---

## Project Pipeline

1. Connect Power Query to a folder containing multiple Excel sales files.
2. Extract the reporting month from the file names.
3. Expand workbook sheets and filter valid sheet objects.
4. Remove junk rows and promote correct headers.
5. Combine data from all monthly files.
6. Apply correct data types and calculate **Sales Amount**.
7. Generate a structured dataset for reporting.

---

## Final Dataset

| Column Name  |
| ------------ |
| Report Month |
| Order ID     |
| Order Date   |
| Country      |
| City         |
| Product      |
| Category     |
| Units Sold   |
| Unit Price   |
| Sales Amount |

---

## Project Output

### Product Sales by Month Report

<img src="https://github.com/Vyaskaushal201/combine-excel-files-from-folder-power-query/blob/cb198b2e2dba99ee088ad2f3efa4aa5e5eb404c1/Images/Monthly-Sales-Report-Chart.png.png" alt="Image Description" width="600">

---

### Power Query Transformation Steps

<img src="https://github.com/Vyaskaushal201/combine-excel-files-from-folder-power-query/blob/6c03bec0fd331a99ef2d02cf3723c658f450659a/Images/Power-Query-Transformation-Steps.png.png" alt="Image Description" width="600">

---

### Final Clean Dataset

<img src="https://github.com/Vyaskaushal201/combine-excel-files-from-folder-power-query/blob/98ba3c0408009a5d7114117db29ff4d78b42e7d3/Images/Final-Sales-Dataset.png.png" alt="Image Description" width="600">

---

## Key Features

* Combine multiple Excel files automatically
* Remove junk rows and clean raw data
* Extract report month from file names
* Calculate sales metrics automatically
* Generate monthly product sales report

---

## Disclaimer

This project is a practice data analytics project using a simulated dataset created for learning purposes.
It is not affiliated with or based on real data from Apple Inc.

---

## Author

Kaushal
|
Aspiring Data Analyst
