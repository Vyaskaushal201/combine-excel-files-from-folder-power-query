# Apple Product Sales Data Automation using Power Query

## Overview

This project demonstrates how Microsoft Excel Power Query can automate the process of combining multiple monthly sales reports into a single structured dataset.

The solution reads multiple Excel files from a folder, cleans the raw data, standardizes the structure, and produces a dataset ready for analysis and reporting.

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

<img src="https://github.com/Vyaskaushal201/Gmail-To-Excel-Automation/blob/e87784d210005aea2a88de8fc343e31020e2bc0e/PowerQuery_Transformation_Steps.png" alt="Image Description" width="600">

---

### Power Query Transformation Steps

![Power Query Steps](images/power_query_steps.png)

---

### Final Clean Dataset

![Clean Dataset](images/final_dataset.png)

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
Aspiring Data Analyst
