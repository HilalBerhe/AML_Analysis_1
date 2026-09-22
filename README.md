# AML Transaction Monitoring & Investigation Case Study

## Project Overview

This project is a synthetic Anti-Money Laundering (AML) transaction monitoring and investigation case study developed using **Microsoft Excel** and **Power Query**.

The purpose of the project was to review a customer's transaction history, compare observed activity against the customer's Know Your Customer (KYC) profile, identify potentially suspicious transaction patterns, and document an investigation conclusion.

The review covered **1,374 transactions from March 1, 2026 through August 31, 2026**.

### Key Results

- Reviewed **1,374 transactions**
- Identified **10 near-$10K transactions totaling $93,280**
- Identified **7 cash deposits totaling $67,080**
- Cash deposits ranged from approximately **$9,150 to $9,890**
- Identified **3 international wire transfers totaling $26,200**
- International wires were sent to **Orion Trading LLC in the United Arab Emirates**
- Cash deposits occurred across multiple Northern Virginia branch locations
- Observed repeated cash-deposit-to-international-wire patterns
- Identified activity materially inconsistent with the customer's expected KYC profile

---

## 1. Customer Profile & KYC Review

![Customer Profile](images/01-customer-profile.png)

The investigation begins with the customer's KYC profile to establish a baseline for expected account activity.

The synthetic customer is an **IT Project Manager** earning approximately **$92,000 annually** and maintaining a personal checking account.

### Expected Customer Activity

- **Expected Monthly Credits:** $6,000–$10,000
- **Expected Monthly Debits:** $4,000–$9,000
- **Expected Cash Activity:** Occasional, typically under $1,000 per month
- **Expected International Activity:** None disclosed
- **Primary Channels:** Mobile, Debit Card, and ACH
- **Initial Risk Rating:** Low

This profile was used throughout the investigation to compare the customer's expected behavior against observed transaction activity.

---

## 2. Raw Transaction Data

![Transactions](images/02-transactions.png)

The **Transactions** worksheet contains the complete transaction-level dataset used for the investigation.

The dataset includes fields such as:

- Transaction ID
- Transaction Date
- Transaction Time
- Account ID
- Customer ID
- Direction
- Transaction Type
- Amount
- Currency
- Counterparty Name
- Counterparty Country
- Channel
- Branch City
- Branch State
- Merchant Category
- Description
- Balance After Transaction

The account contained **1,374 transactions** across several transaction types, including:

- Card Purchases
- ACH Debits
- ACH Credits
- Payroll
- P2P Transfers
- ATM Withdrawals
- Cash Deposits
- Wire Transfers
- Cashier's Checks

The raw transaction data served as the source dataset for the Power Query transformation and subsequent AML analysis.

---

## 3. Transaction Transformation & AML Flags

![Transactions Table](images/03-transactions-table.png)

The raw transaction data was cleaned, transformed, and organized into a structured analytical table using **Power Query** and **Excel**.

Additional analytical fields were created to support transaction monitoring and investigation.

Examples include:

- Signed transaction amounts
- Large transaction indicators
- Near-$10K transaction flags
- Cash activity indicators
- Transaction categorization
- Direction-based transaction analysis
- AML review fields

Power Query helped separate the original source data from the analytical dataset and created a repeatable workflow for transaction review.

### Power Query Techniques Used

- Loading transaction data into Power Query
- Reviewing and correcting data types
- Creating conditional columns
- Transforming transaction fields
- Creating AML flag columns
- Organizing transaction categories
- Preparing data for PivotTable analysis
- Loading transformed data back into Excel

---

## 4. Transaction Analysis Dashboard

![Analysis Dashboard](images/04-analysis.png)

The **Analysis** worksheet summarizes transaction activity and highlights transactions requiring additional review.

The analysis included:

- Transaction count by transaction type
- Total dollar amount by transaction type
- Flagged transaction count
- Flagged transaction amount
- Flagged transaction percentage
- Near-$10K activity
- Cash deposit analysis
- International wire analysis
- Comparison of total activity versus flagged activity
- PivotTable-based transaction summaries
- Charts and visualizations

Cash deposits were identified as the primary source of flagged activity.

All **7 reviewed cash deposits were flagged**, totaling **$67,080**.

The customer's KYC profile indicated expected cash activity of typically less than **$1,000 per month**, making the observed cash activity significantly inconsistent with the customer's expected behavior.

---

## 5. Near-$10K Transaction Review

![Near $10K Review](images/05-near-10k-review.png)

A focused investigation was conducted on transactions occurring close to the $10,000 threshold.

The review identified **10 transactions totaling $93,280** between **July 6 and July 18, 2026**.

### Cash Deposits

The transaction sequence included:

- **7 cash deposits**
- **Total cash deposited: $67,080**
- Individual cash deposits ranged from approximately **$9,150 to $9,890**

The deposits occurred across multiple Northern Virginia branch locations, including:

- Alexandria
- Arlington
- Falls Church
- Springfield
- Fairfax
- Annandale
- Tysons

The repeated use of different branch locations combined with repeated near-threshold cash deposits represented an important transaction-monitoring indicator.

### International Wire Transfers

The review also identified **3 outbound international wire transfers totaling $26,200**.

The wires were sent to **Orion Trading LLC in the United Arab Emirates**.

The wire amounts were:

- **$8,200**
- **$8,800**
- **$9,200**

The customer had disclosed **no expected international activity** in the KYC profile.

### Transaction Pattern

The review identified repeated sequences involving:

**Cash Deposit → International Wire Transfer**

The activity occurred within relatively short time intervals and repeated more than once during the review period.

The combination of:

- Near-$10K cash deposits
- Repeated cash activity
- Multiple branch locations
- International wire transfers
- Rapid movement of funds
- Activity inconsistent with the KYC profile

created indicators consistent with potential **structuring and rapid movement of funds**.

A single transaction does not establish suspicious activity. The concern resulted from the overall transaction pattern and its inconsistency with the customer's stated profile.

---

## 6. Final Investigation Summary

![Analysis Summary](images/06-analysis-summary.png)

The **Analysis Summary** worksheet consolidates the results of the investigation into a final AML review.

The summary includes:

- Customer overview
- Expected customer activity
- Activity reviewed
- Key findings
- Risk indicators
- Profile comparison
- Risk assessment
- Investigation disposition

### Key Findings

#### Repeated Near-$10K Cash Deposits

Seven cash deposits totaling **$67,080** occurred near but below the $10,000 threshold.

#### Activity Inconsistent With KYC Profile

The customer's expected cash activity was typically under **$1,000 per month**, while the reviewed transaction sequence included approximately **$67,080 in cash deposits**.

#### International Wire Activity

The customer disclosed no expected international activity, yet **3 international wires totaling $26,200** were sent to the United Arab Emirates.

#### Rapid Movement of Funds

Multiple cash deposits were followed by outbound international wire transfers during the reviewed period.

#### Multiple Branch Locations

Cash deposits were made across several Northern Virginia branch locations.

#### Repeated Transaction Pattern

The activity occurred multiple times rather than appearing as a single isolated transaction.

---

# Excel Functions & Techniques Used

## COUNTIF

Used to count transactions meeting a single condition.

Example:

`=COUNTIF(Transaction_Type_Range,"Cash Deposit")`

Applications included:

- Counting cash deposits
- Counting wire transfers
- Counting transaction categories
- Counting flagged transactions

---

## COUNTIFS

Used to count transactions meeting multiple conditions.

Example:

`=COUNTIFS(Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")`

Applications included:

- Counting near-$10K cash deposits
- Counting transactions within specific dollar ranges
- Counting transactions matching multiple AML criteria

---

## SUMIF

Used to calculate the total dollar amount associated with a specific transaction category.

Example:

`=SUMIF(Transaction_Type_Range,"Cash Deposit",Amount_Range)`

Applications included calculating total amounts for:

- Cash Deposits
- Wire Transfers
- ACH Transactions
- Card Purchases
- P2P Transfers

---

## SUMIFS

Used to calculate transaction amounts meeting multiple conditions.

Example:

`=SUMIFS(Amount_Range,Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")`

Applications included:

- Near-$10K transaction totals
- Cash deposit totals
- Flagged transaction totals
- Transaction totals based on multiple AML criteria

---

## IF

Used to classify transactions based on transaction characteristics.

Example:

`=IF(Amount>=10000,"Large Transaction","Normal")`

Applications included:

- Large transaction flags
- AML classification
- Transaction categorization
- Flagged versus normal activity

---

## AND

Used with `IF` statements when multiple conditions had to be satisfied simultaneously.

Example:

`=IF(AND(Transaction_Type="Cash Deposit",Amount>=8000,Amount<10000),"Near 10K","Normal")`

This was useful for identifying transactions that met multiple AML review criteria.

---

# Additional Excel Skills Demonstrated

This project also used:

- Microsoft Excel Tables
- Power Query
- PivotTables
- PivotCharts
- Conditional Columns
- Filtering
- Sorting
- Transaction Categorization
- AML Flag Creation
- Data Aggregation
- Percentage Calculations
- KYC-to-Transaction Comparison
- Transaction Drill-Down Analysis
- Dashboard Creation
- Data Visualization
- Summary Tables
- Transaction Count Analysis
- Transaction Amount Analysis
- Investigation Documentation

---

# AML Skills Demonstrated

This project demonstrates practical exposure to:

- AML Transaction Monitoring
- KYC Profile Review
- Customer Activity Analysis
- Transaction Pattern Analysis
- Potential Structuring Identification
- Near-Threshold Transaction Analysis
- Cash Activity Review
- International Wire Review
- Rapid Movement of Funds Analysis
- Multiple-Branch Activity Review
- AML Alert Investigation
- Risk Indicator Identification
- Profile-to-Activity Comparison
- Investigation Documentation
- Enhanced Review Escalation
- SAR Consideration

---

# Investigation Conclusion

The reviewed transaction activity was materially inconsistent with the customer's stated KYC profile.

The customer was expected to conduct limited cash activity and had disclosed no expected international activity. However, the investigation identified repeated cash deposits immediately below $10,000, substantial cash activity relative to the customer's expected profile, use of multiple branch locations, and subsequent international wire transfers.

The combination of these factors created multiple AML risk indicators that warranted further investigation.

## Disposition

**Escalate for enhanced review and SAR consideration in accordance with institutional procedures.**

The analysis does not conclude that money laundering occurred. Instead, the identified patterns warrant additional review to determine whether the activity has a reasonable explanation or requires further escalation.

---

# Workbook Structure

| Worksheet | Purpose |
|---|---|
| **Customer Profile** | Contains KYC information and expected customer activity |
| **Transactions** | Contains the original transaction-level dataset |
| **Transactions Table** | Contains cleaned and transformed transaction data |
| **Analysis** | Contains transaction metrics, PivotTable analysis, flags, and visualizations |
| **Near-10K Review** | Contains detailed review of near-threshold cash deposits and international wires |
| **Analysis Summary** | Contains final findings, red flags, risk assessment, and disposition |

---

# Repository Structure

    AML-Transaction-Monitoring-Case-Study/
    │
    ├── README.md
    ├── AML-Analysis.xlsx
    │
    └── images/
        ├── 01-customer-profile.png
        ├── 02-transactions.png
        ├── 03-transactions-table.png
        ├── 04-analysis.png
        ├── 05-near-10k-review.png
        └── 06-analysis-summary.png

---

# Disclaimer

This project was created for **educational and portfolio purposes only**.

All customer names, account information, counterparties, transaction records, and financial activity shown in this project are **synthetic and do not represent real individuals or financial accounts**.
