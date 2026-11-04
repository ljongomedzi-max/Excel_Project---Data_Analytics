# Project 1 – Interactive Salary Dashboard Using Excel

## Table of Contents
[1. Executive Summary](#1-executive-summary)  
[2. Project Objective](#2-project-objective)  
[3. Business Questions](#3-business-questions)  
[4. Dataset](#4-dataset)  
[5. Database Design](#5-database-design)  
[6. Tools and Technologies](#6-tools-and-technologies)  
[7. Exploratory Data Analysis](#7-exploratory-data-analysis)  
[8. Key Findings](#8-key-findings)  
[9. Career Insights](#9-career-insights)  
[10. Recommendations](#10-recommendations)  
[11. Limitations](#11-limitations)  
[12. Conclusion](#12-conclusion)  
[13. Dashboard and Workbook Structure](#13-dashboard-and-workbook-structure)  
[14. Project Skills Demonstrated](#14-project-skills-demonstrated)  
[15. Portfolio Project Summary](#15-portfolio-project-summary)

---

## 1. Executive Summary

Project 1 is an **interactive salary dashboard built in Microsoft Excel** using job-posting data. The workbook provides a way to explore salary information across job titles, countries, and job schedule types.

The supplied workbook contains **32,672 job-posting records** covering the period from **1 January to 31 December 2023**. The data includes job titles, locations, job platforms, work arrangements, posting dates, country, salary rate, annual and hourly salary values, company names, and listed skills.

The dashboard uses three selection fields:

- Job title
- Country
- Job schedule type

The dashboard includes charts designed to compare salary information across job-title categories and job schedule types. Supporting worksheets contain the data and the calculation ranges used by the dashboard.

This project demonstrates how Excel tables, formulas, data validation, and charts can be combined to turn a job-postings dataset into an interactive analytical tool.

---

## 2. Project Objective

The main objective is to develop an Excel dashboard that makes job-market salary information easier to explore.

The project focuses on:

- Comparing salary levels across job titles.
- Exploring salary differences between countries.
- Examining salary patterns across job schedule types.
- Allowing users to change dashboard selections and explore different views.
- Presenting results through charts rather than relying only on raw data.
- Organising the workbook so that source data, supporting calculations, and dashboard visuals work together.

---

## 3. Business Questions

The dashboard is designed to help users investigate questions such as:

1. How do salary levels differ across job-title categories?
2. How does the selected country's salary profile compare with other countries represented in the data?
3. How do salary levels vary by job schedule type?
4. How does selecting a different job title affect the salary comparison?
5. Which job schedule types are represented in the selected view?
6. How can users interactively explore salary information without manually filtering the source data?

The answers depend on the selections made in the dashboard and on the salary observations available for each selection.

---

## 4. Dataset

### Dataset Overview

The workbook contains a table named `jobs` on the `Data` worksheet.

| Dataset attribute | Description |
|---|---|
| Number of job postings | 32,672 |
| Period covered | 1 January–31 December 2023 |
| Job-title categories | 10 |
| Countries represented | 111 |
| Source-data worksheet | `Data` |
| Excel table name | `jobs` |

### Main Fields

The source table contains 16 columns:

| Field | Description |
|---|---|
| `job_title_short` | Standardised job-title category |
| `job_title` | Full job title from the posting |
| `job_location` | Location stated in the posting |
| `job_via` | Job platform or source |
| `job_schedule_type` | Employment or schedule type |
| `job_work_from_home` | Indicates whether the posting is marked as work from home |
| `search_location` | Location used in the job search |
| `job_posted_date` | Date and time the job was posted |
| `job_no_degree_mention` | Indicates whether a degree requirement is not mentioned |
| `job_health_insurance` | Indicates whether health insurance is mentioned |
| `job_country` | Country associated with the posting |
| `salary_rate` | Salary payment basis, such as year or hour |
| `salary_year_avg` | Average annual salary value, where available |
| `salary_hour_avg` | Average hourly salary value, where available |
| `company_name` | Company associated with the posting |
| `job_skills` | Skills listed for the job |

### Workbook Worksheets

| Worksheet | Purpose |
|---|---|
| `Data` | Source job-posting records stored in the `jobs` table |
| `Data Validation` | Supporting lists and values used by dashboard selections |
| `Title` | Job-title salary calculations and chart-supporting data |
| `Country` | Country-level salary calculations and supporting data |
| `Type` | Job schedule-type salary calculations and chart-supporting data |
| `Platform` | Job-platform counts and supporting calculations |
| `Salary Dashboard` | Main interactive dashboard |

---

## 5. Database Design

This project uses an **Excel workbook structure rather than a relational database**.

The source data is organised as a structured Excel table named `jobs`. Supporting worksheets separate the dashboard calculations and validation lists from the source records.

The workbook follows this general flow:

**Data worksheet → Supporting calculation worksheets → Salary Dashboard**

The supporting worksheets have distinct responsibilities:

- `Data Validation` supplies values for the dashboard selection fields.
- `Title` supports job-title salary comparisons.
- `Country` supports country-level salary calculations.
- `Type` supports job schedule-type salary comparisons.
- `Platform` supports job-platform calculations.

This organisation makes the workbook easier to maintain and helps keep the dashboard separate from the underlying source data.

---

## 6. Tools and Technologies

### Spreadsheet Application
Microsoft Excel

### Main Excel Features

The workbook uses or is structured around:

- Excel Tables
- Structured references
- Formulas and conditional calculations
- Data validation and dropdown selections
- Sorting and ranking
- `COUNTIFS` for conditional counts
- `XLOOKUP` for retrieving corresponding values
- `IF` for selection-dependent calculations
- `NA()` for excluding non-selected categories from chart series
- Bar charts
- Multiple worksheets for data, calculations, and presentation

The dashboard's supporting calculations use the selected job title, country, and schedule type to determine the values displayed in the charts.

---

## 7. Exploratory Data Analysis

### 7.1 Job-Title Categories

The dataset contains 10 standardised job-title categories.

The largest categories by record count in the supplied workbook are:

| Job-title category | Postings |
|---|---:|
| Data Analyst | 9,606 |
| Data Scientist | 8,495 |
| Data Engineer | 6,783 |
| Senior Data Engineer | 2,014 |
| Senior Data Scientist | 2,009 |
| Senior Data Analyst | 1,484 |
| Business Analyst | 1,001 |
| Machine Learning Engineer | 621 |
| Software Engineer | 573 |
| Cloud Engineer | 86 |

Data Analyst is the largest category in this workbook's source data.

### 7.2 Salary Information

The dataset contains both annual and hourly salary fields.

- 22,036 records contain a numeric value in `salary_year_avg`.
- 10,636 records contain a numeric value in `salary_hour_avg`.

Salary comparisons should use the appropriate salary basis. Annual and hourly values should not be combined as though they were expressed in the same unit.

### 7.3 Country Coverage

The source data represents 111 countries. The United States has the largest number of postings in the supplied workbook, followed by Sudan, India, Canada, and the United Kingdom.

Country-level comparisons depend on the number of salary observations available for each country.

### 7.4 Job Schedule Types

Full-time is the most common schedule type in the workbook, followed by Contractor and Full-time and Part-time.

Some records contain combined schedule descriptions, such as Full-time and Part-time. These should be treated as the categories recorded in the source data rather than assumed to be separate individual arrangements.

### 7.5 Dashboard Salary Comparisons

<img width="1757" height="847" alt="Screenshot 2026-10-31 160330" src="https://github.com/user-attachments/assets/a7d5408a-ca21-413a-b6dd-55a7015f596c" />

The `Salary Dashboard` worksheet provides selection fields for:

- Job title
- Country
- Job schedule type

The workbook includes two bar charts. Their supporting data is held on the `Title` and `Type` worksheets:

1. A job-title comparison chart using the values on the `Title` worksheet.
2. A job schedule-type comparison chart using the values on the `Type` worksheet.

The supporting `Country` worksheet contains country-level salary calculations. The dashboard is intended to let users change the selection fields and examine the corresponding salary comparisons.

---

## 8. Key Findings

The supplied workbook supports the following observations:

### Finding 1 — Data Analyst Is the Largest Job Category

Data Analyst postings account for 9,606 of the 32,672 records in the workbook.

### Finding 2 — Salary Data Has Two Different Payment Bases

The source data contains annual and hourly salary values. These measures need to be analysed separately to avoid misleading comparisons.

### Finding 3 — The United States Has the Largest Country Representation

The United States accounts for the largest number of job postings in the supplied dataset. The number of records varies substantially across countries.

### Finding 4 — Full-Time Is the Most Common Schedule Category

Full-time postings are the largest schedule-type group in the workbook, with contractor postings forming the next largest group.

### Finding 5 — Interactive Selections Support Different Comparisons

The dashboard's job-title, country, and schedule-type selections are intended to let users explore salary comparisons under different combinations of filters.

Salary conclusions should always be interpreted in the context of the selected filters and the number of salary observations available.

---

## 9. Career Insights

The workbook can help job seekers explore how salary information varies across job-title categories, countries, and work arrangements.

Useful considerations include:

- Compare salaries for relevant job-title categories rather than relying on a single overall figure.
- Consider country and location when interpreting salary differences.
- Distinguish annual salary from hourly salary.
- Review job schedule types to understand the work arrangements represented in the data.
- Consider the number of available salary observations before drawing conclusions from a comparison.

The dashboard is an exploratory tool. It does not establish that a particular job title, country, or schedule type causes a higher salary.

---

## 10. Recommendations

### For Job Seekers

- Use the job-title selection to compare relevant roles.
- Explore country selections to understand geographic differences in the dataset.
- Compare schedule types where sufficient salary observations are available.
- Check whether salary values are annual or hourly before making comparisons.
- Consider job responsibilities, experience, and other factors not fully represented by a salary figure.

### For Analysts and Dashboard Users

- Change one selection at a time when investigating an unexpected result.
- Check the underlying calculation worksheets when a chart appears blank or inconsistent.
- Confirm that formulas and chart series respond to the selected values.
- Treat missing salary information separately from a salary of zero.
- Use clear chart titles, labels, and number formats so that salary units are immediately understandable.

### For Future Development

The dashboard could be extended with:

- KPI cards for the number of matching postings and available salary observations.
- A clearer distinction between annual and hourly salary analysis.
- Additional visualisations for remote and non-remote work.
- Job-platform comparisons.
- Salary distribution charts, such as box plots or histograms.
- More detailed filters for date, company, and remote-work status.
- A dedicated explanation of how each salary measure is calculated.

---

## 11. Limitations

Several limitations should be considered when interpreting the workbook.

### Salary Coverage

Not every job posting contains salary information. Results are therefore based on the records with usable salary values for the relevant measure.

### Different Salary Units

The dataset contains annual and hourly salary values. They should not be compared directly without an appropriate conversion and clear assumptions.

### Job Posting Representation

A posting does not necessarily represent a unique vacancy or an actual hire. The dataset may contain repeated or syndicated advertisements.

### Geographic Differences

Salary comparisons across countries may be affected by currency, cost of living, local labour-market conditions, and differences in how salary information is reported. The workbook should not be treated as a cost-of-living-adjusted salary comparison unless such adjustments are added.

### Job-Title and Schedule Categories

Standardised job-title categories may group together postings with different responsibilities or seniority. Schedule descriptions can also contain combined categories.

### Historical Data

The workbook covers postings dated in 2023. It describes the period represented by the dataset and should not automatically be interpreted as a current salary-market report.

### Dashboard Calculation and Filter Behaviour

Dashboard results depend on the supporting formulas, validation lists, and chart source ranges. If a selection produces no matching salary observations, the charts may show blank or unavailable values. Such a result should not be interpreted as a salary of zero.

---

## 12. Conclusion

Project 1 demonstrates how Microsoft Excel can be used to organise job-posting data and present salary information through an interactive dashboard.

The workbook combines a structured source-data table with supporting worksheets for job-title, country, schedule-type, and platform calculations. The main dashboard provides selection fields and charts that help users explore salary comparisons without working directly with the full source table.

The project highlights the importance of checking salary units, data coverage, filters, and calculation logic before drawing conclusions. It also demonstrates a practical workflow:

**Source Data → Structured Excel Table → Supporting Calculations → Interactive Dashboard → Exploration and Interpretation**

Overall, the project provides a practical example of using Excel to transform job-market data into a more accessible analytical presentation.

---

## 13. Dashboard and Workbook Structure

[Dashboard workbook](https://github.com/ljongomedzi-max/Excel_Project_-_Data_Job_Analysis/blob/main/Project%201%20-%20Final%20Dashboard/Project%201%20-%20Final%20Dashboard.xlsx)

### Main Dashboard

The `Salary Dashboard` worksheet is the main user-facing sheet.

It contains three selection fields:

- **Job title**
- **Country**
- **Job Schedule Type**

It also contains two bar charts:

- Job-title salary comparison.
- Job schedule-type salary comparison.

### Supporting Worksheets

| Worksheet | Role in the dashboard |
|---|---|
| `Data` | Stores the source records in the `jobs` Excel table |
| `Data Validation` | Provides selection lists and supporting values |
| `Title` | Holds job-title comparison calculations |
| `Country` | Holds country-level salary calculations |
| `Type` | Holds schedule-type comparison calculations |
| `Platform` | Holds job-platform calculations |
| `Salary Dashboard` | Presents the selections and charts |

### Recommended User Workflow

1. Open the `Salary Dashboard` worksheet.
2. Select a job-title category.
3. Select a country.
4. Select a job schedule type.
5. Review the charts and compare the displayed salary values.
6. Change a selection to explore another view.
7. If a chart does not respond as expected, inspect the relevant supporting worksheet and its chart source ranges.

The dashboard should be opened in a version of Excel that supports the formulas and chart features used in the workbook. Recalculate the workbook if formula results appear outdated.

---

## 14. Project Skills Demonstrated

### Excel and Spreadsheet Skills

- Structuring data as an Excel Table
- Using multiple worksheets
- Writing and reviewing formulas
- Using structured references
- Applying data validation
- Building interactive dropdown selections
- Using conditional calculations
- Creating and configuring charts
- Linking chart series to supporting calculation ranges

### Data Analysis

- Exploring job-title categories
- Examining salary coverage
- Comparing salary measures
- Analysing country representation
- Comparing job schedule types
- Interpreting filtered results

### Dashboard Development

- Separating source data from calculations and presentation
- Designing a user-facing dashboard
- Connecting selection fields to supporting calculations
- Presenting comparisons visually
- Considering blank or unavailable results

### Analytical Communication

- Identifying patterns in job-posting data
- Explaining the limits of salary comparisons
- Presenting findings in a structured format
- Translating data into practical observations

---

## 15. Portfolio Project Summary

### Project Title

**Project 1 – Interactive Salary Dashboard Using Excel**

### Tools

Microsoft Excel | Excel Tables | Formulas | Data Validation | Charts

### Dataset

32,672 job postings dated in 2023.

### Primary Focus

Salary exploration by job title, country, and job schedule type.

### Main Deliverable

An Excel workbook with a source-data table, supporting calculation worksheets, dashboard selection fields, and salary comparison charts.

### Main Outcome

A practical demonstration of using Excel to organise job-market data, build interactive dashboard elements, and communicate salary comparisons in a visual and accessible format.
