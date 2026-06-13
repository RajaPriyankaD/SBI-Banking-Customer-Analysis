# SBI Banking Customer Analysis Dashboard

## Project Overview

This project presents an interactive Excel Dashboard for analyzing SBI banking customer data. The dashboard provides insights into customer demographics, account activity, loan behavior, deposits, customer satisfaction, regional performance, and digital banking adoption.

The project demonstrates advanced Excel skills including Formulas, Pivot Tables, Pivot Charts, Conditional Formatting, Slicers, and Dashboard Design.

---

## Business Problem

Banks manage large volumes of customer data across multiple branches and regions. The objective of this project is to transform raw banking data into actionable insights that help management:

* Understand customer distribution
* Monitor loan performance
* Analyze deposits and balances
* Track customer satisfaction
* Evaluate digital banking adoption
* Identify high-performing regions

---

## Dataset

The dataset contains customer-level banking information with fields such as:

| Column Name        | Description                             |
| ------------------ | --------------------------------------- |
| Customer_ID        | Unique customer identifier              |
| Age                | Customer age                            |
| Gender             | Male/Female/Others                      |
| State              | Customer location                       |
| Account_Type       | Savings, Current, Salary, Fixed Deposit |
| Balance            | Account balance                         |
| Loan_Status        | Active / No Loan                        |
| Loan_Amount        | Total loan amount                       |
| Credit_Score       | Customer credit score                   |
| Transaction_Count  | Number of transactions                  |
| Internet_Banking   | Yes/No                                  |
| Mobile_Banking     | Yes/No                                  |
| Satisfaction_Score | Customer satisfaction rating            |

---

## Dashboard KPIs

The dashboard includes the following KPI cards:

* Total Customers
* Total Deposits
* Total Loan Amount
* Average Balance
* Active Loan Customers
* Average Credit Score
* Digital Banking Users %
* Customer Satisfaction Score

### Sample Excel Formulas

```excel
=COUNTA(Customer_ID)
=SUM(Balance)
=AVERAGE(Balance)
=COUNTIF(Loan_Status,"Active")
```

---

## Dashboard Components

### Customer Distribution

* Age Group Analysis
* Gender-wise Distribution
* Occupation Analysis

### Account Activity

* Account Type Analysis
* Transaction Trends

### Loan Behavior

* Active Loan Customers
* Loan Amount by Branch
* Risk Customer Analysis

### Deposits & Balances

* Deposit Distribution
* Average Balance Analysis
* High-Value Customers

### Customer Satisfaction & Risk

* Satisfaction Score Trends
* Credit Score Analysis

### Regional Performance

* State-wise Customer Distribution
* Branch Performance
* Deposit Contribution by Region

### Digital Banking Adoption

* Mobile Banking Users
* Internet Banking Users
* Digital vs Non-Digital Customers

---

## Excel Features Used

* Excel Formulas
* Pivot Tables
* Pivot Charts
* Slicers
* Conditional Formatting
* Data Validation
* Dashboard Design
* Interactive Filtering

---

## Dashboard Layout

```text
--------------------------------------------------
| SBI Banking Customer Analysis Dashboard        |
--------------------------------------------------
| KPI Cards                                      |
--------------------------------------------------
| Age Analysis | Gender Split | Account Types    |
--------------------------------------------------
| State Analysis | Loan Analysis                 |
--------------------------------------------------
| Digital Banking | Satisfaction Trend           |
--------------------------------------------------
| Slicers                                        |
--------------------------------------------------
```

---

## Key Insights

* Savings accounts represent the largest customer segment.
* Digital banking adoption continues to grow across regions.
* Customers with higher credit scores maintain larger balances.
* Certain branches contribute significantly to total loan portfolios.
* Customer satisfaction is positively correlated with digital banking usage.

---

## Tools & Technologies

* Microsoft Excel
* Pivot Tables
* Pivot Charts
* Data Visualization
* Business Analytics

---

## Project Files

* SBI_Banking_Customer_Analysis_Dashboard.xlsx
* SBI_Banking_Dashboard_Presentation.pptx

---

## Learning Outcomes

Through this project, I gained hands-on experience in:

* Data Cleaning
* Dashboard Development
* Banking Analytics
* KPI Design
* Data Visualization
* Business Intelligence Reporting
* Excel Automation Techniques
