# Excel Data Cleaning & Transformation — Product Dataset

## 📌 Project Overview

This project demonstrates my practical skills in **data cleaning, data preprocessing, transformation, and formatting using Microsoft Excel**.

Real-world datasets often contain missing values, inconsistent text, duplicate records, poorly structured fields, and formatting issues. Before performing meaningful analysis, these data quality issues need to be identified and addressed.

In this project, I worked with a **Product Dataset** containing product information such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

The objective was to transform the raw dataset into a cleaner and more structured format suitable for further analysis.

---

## 🎯 Project Objectives

The main objectives of this project were to:

* Identify and handle missing values
* Standardize inconsistent text
* Identify and correct category inconsistencies
* Remove duplicate records
* Split structured information from Product ID
* Merge product and brand information
* Apply appropriate number and date formatting
* Use conditional formatting to improve data readability
* Apply Excel formulas for data transformation

---

## 📊 Dataset Description

The dataset contains product-level information with the following attributes:

| Column             | Description                                                      |
| ------------------ | ---------------------------------------------------------------- |
| Product ID         | Unique identifier containing manufacturing date and country code |
| Product Name       | Name of the product                                              |
| Brand Name         | Brand associated with the product                                |
| Quantity           | Number of units available                                        |
| Category           | Product category                                                 |
| Price              | Product price                                                    |
| Manufacturing Date | Date extracted from Product ID                                   |
| Country Code       | Country code extracted from Product ID                           |
| Product Brand      | Combined Product Name and Brand Name                             |

The dataset contains **31 populated product records** used for the cleaning and transformation exercise.

---

# 🧹 Data Cleaning Process

## 1. Handling Missing Values

### Price

The Price column was checked for missing or invalid values.

For missing price information, an appropriate approach is to use a representative value such as the **average price of available products**, while also considering whether category- or product-level averages would provide a more accurate result.

An Excel formula was used to check for blank or zero values and substitute an average value when required.

```excel
=IF(OR(ISBLANK(H2),H2=0),AVERAGE(H2:H35),H2)
```

This demonstrates the use of:

* `IF()`
* `OR()`
* `ISBLANK()`
* `AVERAGE()`

### Category

The Category column was checked for missing values.

Where category information is unavailable, `"Unknown"` can be used as a temporary placeholder. In a real-world project, category assignment could also be based on:

* Product Name
* Brand
* Product description
* Existing business classification rules

Formula used:

```excel
=IF(OR(ISBLANK(J2),J2=""),"Unknown",J2)
```

---

# ✨ 2. Correcting Inconsistent Data

## Product Name Standardization

Product names were checked for inconsistent capitalization and unnecessary spaces.

The following formula was used to standardize the text:

```excel
=PROPER(TRIM(E2))
```

### Functions used

* `TRIM()` — removes unnecessary spaces
* `PROPER()` — standardizes capitalization

This helps ensure that values such as inconsistent capitalization or extra spaces do not create separate categories during analysis.

---

## Category Standardization

The Category column was also reviewed for inconsistent formatting.

Formula used:

```excel
=PROPER(TRIM(J2))
```

This ensures consistent capitalization and removes unnecessary spaces.

---

## Find & Replace

Excel's **Find & Replace** feature was also used to standardize text and correct inconsistent entries.

### Shortcut

```text
Ctrl + H
```

This is useful when the same typo or formatting inconsistency appears multiple times across a dataset.

---

# 🗑️ 3. Removing Duplicate Records

Duplicate records were checked based on the **entire row**, rather than relying only on Product ID.

Excel's built-in:

**Data → Remove Duplicates**

feature was used to identify and remove duplicate records.

This ensures that the same product record is not counted multiple times during future analysis.

---

# 🔄 4. Splitting and Merging Data

## Splitting Product ID

The original Product ID follows a structure similar to:

```text
28-JAN-US
```

The Product ID contains two useful pieces of information:

```text
28-JAN → Manufacturing Date
US     → Country Code
```

### Manufacturing Date

The date portion was extracted using:

```excel
=LEFT(A2,6)
```

Example:

```text
28-JAN
```

The extracted value was then converted into a proper Excel date using:

```excel
=DATEVALUE(B2&"-2026")
```

The final date was formatted as:

```text
DD-MM-YYYY
```

---

## Country Code

The country code was extracted using:

```excel
=RIGHT(A2,2)
```

Example:

```text
28-JAN-US → US
```

This transformation separates the country information from the original Product ID and makes it easier to analyze by country.

---

# 🔗 5. Merging Product and Brand Information

The **Product Name** and **Brand Name** columns were combined into a new column called:

### Product Brand

Formula used:

```excel
=E2&" "&F2
```

Example:

```text
Product Name: Laptop
Brand Name: Dell

Product Brand: Laptop Dell
```

This creates a more descriptive field that combines both pieces of information.

---

# 💰 6. Number Formatting

## Price

The Price column was formatted using **Currency ($)** formatting.

This improves readability and clearly communicates that the values represent monetary amounts.

Example:

```text
1000 → $1,000.00
80   → $80.00
130  → $130.00
```

---

## Manufacturing Date

The Manufacturing Date was converted into a proper Excel date and formatted using:

```text
DD-MM-YYYY
```

Example:

```text
28-JAN-2026 → 28-01-2026
```

This creates a standardized date format suitable for analysis and sorting.

---

# 🎨 7. Conditional Formatting

Conditional formatting was applied to improve the visual readability of the dataset.

## Price Column

A **Data Bar / Color Scale** was applied to the Price column.

This makes it easier to visually compare:

* Lower-priced products
* Medium-priced products
* Higher-priced products

without manually sorting the data.

---

## Category Column

A custom conditional formatting rule was created to highlight products belonging to:

```text
Electronics
```

This makes it easier to identify and analyze products within the Electronics category.

---

# 🧮 Excel Functions & Features Used

The following Excel functions and features were used throughout the project:

| Function / Feature     | Purpose                            |
| ---------------------- | ---------------------------------- |
| `IF()`                 | Conditional logic                  |
| `OR()`                 | Check multiple conditions          |
| `ISBLANK()`            | Identify blank cells               |
| `AVERAGE()`            | Calculate average price            |
| `PROPER()`             | Standardize capitalization         |
| `TRIM()`               | Remove unnecessary spaces          |
| `LEFT()`               | Extract manufacturing date portion |
| `RIGHT()`              | Extract country code               |
| `DATEVALUE()`          | Convert text into a valid date     |
| `&`                    | Merge text values                  |
| Find & Replace         | Correct and standardize text       |
| Remove Duplicates      | Remove duplicate records           |
| Currency Formatting    | Format prices                      |
| Date Formatting        | Standardize dates                  |
| Conditional Formatting | Highlight and visualize values     |
| AutoFilter             | Filter and inspect dataset         |

---

# 📁 Project Structure

```text
excel-data-cleaning-product-dataset/
│
├── README.md
│
├── Assignment 2 - Data Cleaning and Transformation with Answers.xlsx
│
└── screenshots/
    ├── data-cleaning.png
    ├── data-transformation.png
    └── conditional-formatting.png
```

> File names can be adjusted based on the final GitHub repository structure.

---

# 🛠️ Tools Used

* **Microsoft Excel**
* Excel Formulas & Functions
* Find & Replace
* Remove Duplicates
* Conditional Formatting
* Data Formatting
* Data Transformation

---

# 📈 Key Skills Demonstrated

Through this project, I demonstrated practical knowledge of:

* Data Cleaning
* Data Preprocessing
* Data Transformation
* Missing Value Handling
* Text Standardization
* Duplicate Detection
* Data Formatting
* Excel Functions
* Conditional Formatting
* Data Quality Improvement
* Basic Data Preparation for Analysis

---

# 💡 Key Learning Outcomes

This project helped me understand that **data analysis starts with data quality**.

Before building reports, dashboards, or performing statistical analysis, raw data should be checked for:

* Missing values
* Duplicate records
* Incorrect formats
* Inconsistent text
* Invalid values
* Poorly structured fields

I also gained practical experience using Excel formulas to automate repetitive cleaning and transformation tasks.

---

# 🚀 Future Improvements

As part of my continued development as a Data Analyst, I plan to extend this project by:

* Performing exploratory data analysis
* Creating Excel dashboards
* Building PivotTables and PivotCharts
* Analyzing sales and product trends
* Comparing prices across categories
* Analyzing country-level product distribution
* Identifying high-value products
* Creating automated data-cleaning workflows using Power Query
* Recreating the analysis using SQL and Python

---

# 👨‍💻 About Me

I am an aspiring **Data Analyst** building my portfolio through practical projects involving **Excel, SQL, Python, data cleaning, data visualization, and exploratory data analysis**.

This project is part of my learning journey toward developing practical, job-ready data analytics skills.

---

## ⭐ Portfolio Note

This project demonstrates my ability to take a structured but imperfect dataset and prepare it for reliable analysis using Microsoft Excel.

**From raw data → clean data → analysis-ready data.**
