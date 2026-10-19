# Employee Dashboard: Excel Project

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![Pivot Tables](https://img.shields.io/badge/Pivot%20Tables-25-blue)
![Charts](https://img.shields.io/badge/Charts-7-green)

An Excel dashboard built for a Module 1 Data Analytics assignment. It uses Excel formulas, pivot tables, pivot charts and slicers to analyse salary, experience and gender across 100 employees.

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Dataset](#2-dataset)
3. [Data Preparation](#3-data-preparation)
4. [Method](#4-method)
5. [Dashboard](#5-dashboard)
6. [Key Insights](#6-key-insights)
7. [Data Notes](#7-data-notes)
8. [Repository Structure](#8-repository-structure)
9. [How to Use](#9-how-to-use)
10. [Author](#10-author)

---

## 1. Project Overview

**Objective:** Show how much the organisation spends on salaries and how pay, experience and gender are spread across job titles, cities and departments.

**Questions the dashboard answers**
- What is the total salary bill, and how is it split between male and female employees?
- Which job titles and cities account for the most salary?
- How much experience does the workforce have, and where is it concentrated?
- How is gender balanced across departments and job titles?

---

## 2. Dataset

**Source:** Employee table (`Task_Data`) with **8 columns and 100 rows**. No missing values and no duplicate Employee IDs.

| Column | Description |
|---|---|
| Employee ID | Unique ID (E001 to E100) |
| Full Name | Employee name |
| Department | IT, HR, Sales, Marketing, Finance |
| Job Title | 10 job titles |
| Gender | Male or Female |
| Salary | Salary paid |
| Experience (Years) | Years of experience |
| City | Lagos, Ibadan, Kano, Port Harcourt, Abuja |

---

## 3. Data Preparation

The data did not need clean-up. Extra columns were added with Excel formulas:

| Task | Formula used |
|---|---|
| Split names | First Name and Last Name columns created from Full Name |
| Email address | `=CONCATENATE(J3,".",K3,"@dame.com")` |
| Name length | `=LEN(B3)` |
| Department in capitals, lowercase, title case | `=UPPER(C3)`, `=LOWER(C3)`, `=PROPER(C3)` |
| Remove extra spaces | `=TRIM(B3)` |
| Total and average salary | `=SUM(F3:F102)`, `=AVERAGE(F3:F102)` |
| Count of employees | `=COUNTA(A3:A102)` |
| Salary by department | `=SUMIF(C3:C102,C5,F3:F102)`, `=AVERAGEIF(C3:C102,C12,F3:F102)` |
| Salary by city and experience | `=SUMIFS(F3:F102,G3:G102,">3",H3:H102,H7)` |
| Average salary by gender and city | `=AVERAGEIFS(F3:F102,E3:E102,E3,H3:H102,H7)` |
| Headcount by gender and city | `=COUNTIFS(E3:E102,E95,H3:H102,H99)` |
| Highest salary in a department | `=MAXIFS(F3:F102,C3:C102,C3)` |

---

## 4. Method

| Sheet | Purpose |
|---|---|
| Home Page | Cover page |
| Assignment-Module 1 | Source table and formula work |
| Pivot Table | 25 pivot tables that feed the charts |
| Employee Dashboard | KPIs, 7 pivot charts and 2 slicers |

---

## 5. Dashboard

![Employee Dashboard](images/employee_dashboard.png)

**Slicers:** City, Gender

### KPIs

| KPI | Value |
|---|---:|
| Total Salary Paid to Employees | 34,862,380 |
| Total Salary Paid to Male Employees | 15,649,032 |
| Total Salary Paid to Female Employees | 19,213,348 |
| Total Years of Experience | 535 |
| Total Years of Experience, Male | 264 |
| Total Years of Experience, Female | 271 |

### Charts

| Chart | Type |
|---|---|
| Total Salary Earned by Job Title | Bar |
| Total Years of Experience by Job Title | Column |
| Percentage of Salary Earned by Names | Bar |
| Total Salary Earned by City | Bar |
| Total Salary Earned by Gender | Pie |
| Proportion of Gender by Department | Doughnut |
| Gender by Job Title | Line |

---

## 6. Key Insights

- **Salary bill:** Total salary is 34.86M, an average of about 348,600 per employee (range 200,754 to 493,395).
- **Gender:** Women are 54% of staff and receive 55.1% of salary (19.21M vs 15.65M). Their average salary is about 4.6% higher (355,803 vs 340,196).
- **Job titles:** Recruiters take the largest share of salary (6.48M, 18.6%) and have the highest average (381,431). HR Managers are the smallest group (1.53M, 4.4%) with the lowest average (306,383).
- **Cities:** Ibadan (24.0%) and Lagos (23.4%) account for almost half of salary. Abuja is lowest (13.1%). Lagos has the most employees (25) but the lowest average salary (326,818); Port Harcourt has the highest (366,750).
- **Departments:** IT is the largest, with 30 employees and 32.4% of salary. IT has more women than men (18 vs 12); HR is the only department with more men (11 vs 9).
- **Experience:** The workforce has 535 years of experience in total, about 5.4 years per employee, split almost evenly by gender (271 female, 264 male).

---

## 7. Data Notes

- Only 10 distinct full names appear across the 100 employees, and each name is recorded under both genders. Employee ID is the only unique identifier, so the "by Names" chart shows totals per shared name and not per person. Email addresses repeat for the same reason.
- The formula examples (SUMIF, SUMIFS, COUNTIFS and so on) each use a single criteria cell, so they show the technique rather than a full summary table. The pivot tables carry the full breakdowns.

---

## 8. Repository Structure

```
employee-dashboard-excel/
├── README.md
├── Adebowale_Assignment 1.xlsx
└── images/
    └── employee_dashboard.png
```

---

## 9. How to Use

1. Download or clone this repository.
2. Open `Adebowale_Assignment 1.xlsx` in Excel (slicers need Excel 2013 or later).
3. Go to the **Employee Dashboard** sheet and use the City and Gender slicers to filter the charts.
4. See the **Pivot Table** sheet for the underlying pivots and **Assignment-Module 1** for the formulas.

---

## 10. Author

**Adebowale Sulaimon Adeyinka**, Data Analyst
[LinkedIn](https://www.linkedin.com/) | [Email](mailto:you@example.com)
