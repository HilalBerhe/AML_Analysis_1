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

![Customer Profile](AML%20Analysis%201%20-%20Customer%20Profile.png)

The investigation begins with the customer's KYC profile to establish a baseline for expected account activity.

The synthetic customer is an **IT Project Manager** earning approximately **$92,000 annually** with a personal checking account.

### Expected Customer Activity

- **Expected Monthly Credits:** $6,000–$10,000
- **Expected Monthly Debits:** $4,000–$9,000
- **Expected Cash Activity:** Occasional, typically under $1,000 per month
- **Expected International Activity:** None disclosed
- **Primary Channels:** Mobile, Debit Card, and ACH
- **Initial Risk Rating:** Low

The KYC profile established the expected baseline used to evaluate whether observed transaction activity was reasonable for the customer.

---

## 2. Raw Transaction Data

![Transactions](AML%20Analysis%201%20-%20Transactions.png)

The **Transactions** worksheet contains the complete transaction-level dataset used for the investigation.

The dataset includes:

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

The account contained **1,374 transactions** across multiple transaction types, including card purchases, ACH activity, payroll, P2P transfers, ATM withdrawals, cash deposits, wire transfers, and cashier's checks.

This worksheet served as the raw source data for the analytical workflow.

---

## 3. Transaction Transformation & AML Flags

![Transactions Table](AML%20Analysis%201%20-%20Transactions%20Table.png)

The transaction data was organized and transformed into a structured analytical table using **Excel and Power Query**.

Additional fields were created to support AML transaction monitoring and investigation.

These included:

- Signed transaction amounts
- Large transaction indicators
- Near-$10K indicators
- Transaction classifications
- AML review flags
- Transaction direction analysis

### Power Query Techniques Used

- Loading source transaction data
- Reviewing and assigning data types
- Creating conditional columns
- Transforming transaction fields
- Creating analytical flags
- Organizing transaction categories
- Preparing data for PivotTable analysis
- Loading transformed data back into Excel

Power Query created a repeatable process that separated the raw transaction data from the analytical dataset.

---

## 4. Transaction Analysis Dashboard

![Analysis](AML%20Analysis%201%20-%20Analysis.png)

The **Analysis** worksheet summarizes customer activity by transaction type and highlights activity requiring further review.

The analysis included:

- Transaction count by transaction type
- Total dollar amount by transaction type
- Flagged transaction count
- Flagged transaction amount
- Flagged transaction percentage
- Cash deposit analysis
- Near-$10K activity
- International wire activity
- PivotTable analysis
- Charts comparing total and flagged transaction activity

Cash deposits were identified as the primary source of flagged activity.

All **7 reviewed cash deposits were flagged**, totaling **$67,080**.

The customer's KYC profile indicated expected cash activity of typically less than **$1,000 per month**, making the observed activity significantly inconsistent with the stated profile.

---

## 5. Near-$10K Transaction Review

![Near 10K Review](AML%20Analysis%201%20-%20Near%2010K%20Review.png)

A focused transaction review was performed on activity occurring close to the $10,000 threshold.

The review identified **10 transactions totaling $93,280** between **July 6 and July 18, 2026**.

### Cash Deposit Activity

The reviewed activity included:

- **7 cash deposits**
- **$67,080 total deposited**
- Individual deposits ranging from approximately **$9,150 to $9,890**

The deposits occurred across several Northern Virginia branch locations, including:

- Alexandria
- Arlington
- Falls Church
- Springfield
- Fairfax
- Annandale
- Tysons

### International Wire Activity

The investigation also identified **3 outbound international wires totaling $26,200**.

All three wires were sent to **Orion Trading LLC in the United Arab Emirates**.

The wire amounts were:

- **$8,200**
- **$8,800**
- **$9,200**

The customer's KYC profile disclosed **no expected international activity**.

### Transaction Pattern

The review identified repeated patterns involving:

**Cash Deposit → International Wire Transfer**

The combination of repeated near-$10K deposits, multiple branch locations, subsequent international wires, rapid movement of funds, and significant deviation from expected customer behavior created indicators consistent with potential **structuring and rapid movement of funds**.

A single transaction alone does not establish suspicious activity. The risk assessment was based on the combined transaction pattern and its inconsistency with the customer's expected profile.

---

## 6. Final Investigation Summary

![Analysis Summary](AML%20Analysis%201%20-%20Analysis%20Summary.png)

The **Analysis Summary** worksheet consolidates the investigation into a final AML review.

The summary includes:

- Customer overview
- Expected customer activity
- Activity reviewed
- Key findings
- Risk indicators
- KYC comparison
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

Multiple cash deposits were followed by outbound international wire transfers during the review period.

#### Multiple Branch Locations

Cash deposits were conducted across several Northern Virginia branch locations.

#### Repeated Transaction Pattern

The behavior occurred multiple times rather than appearing as a single isolated transaction.

---

# Excel Functions & Techniques Used

## COUNTIF

Used to count transactions meeting a single condition.

Example:

`=COUNTIF(Transaction_Type_Range,"Cash Deposit")`

Used for:

- Cash deposit counts
- Wire transfer counts
- Transaction category counts
- Flagged transaction counts

---

## COUNTIFS

Used to count transactions meeting multiple criteria.

Example:

`=COUNTIFS(Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")`

Used for:

- Near-$10K transaction counts
- Multi-condition AML flags
- Transaction type and amount analysis

---

## SUMIF

Used to calculate the total amount associated with a specific transaction category.

Example:

`=SUMIF(Transaction_Type_Range,"Cash Deposit",Amount_Range)`

Used to calculate totals for:

- Cash Deposits
- Wire Transfers
- ACH Transactions
- Card Purchases
- P2P Transfers

---

## SUMIFS

Used to calculate dollar amounts satisfying multiple criteria.

Example:

`=SUMIFS(Amount_Range,Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")`

Used for:

- Near-$10K transaction totals
- Flagged transaction totals
- Cash deposit totals
- Multi-condition transaction analysis

---

## IF

Used to classify transactions based on specific characteristics.

Example:

`=IF(Amount>=10000,"Large Transaction","Normal")`

Used for:

- AML flags
- Large transaction classifications
- Flagged vs. normal activity
- Transaction categorization

---

## AND

Used with `IF` when multiple conditions needed to be satisfied.

Example:

`=IF(AND(Transaction_Type="Cash Deposit",Amount>=8000,Amount<10000),"Near 10K","Normal")`

This was useful for creating transaction-level AML indicators based on multiple criteria.

---

# Additional Excel Skills Demonstrated

- Microsoft Excel Tables
- Power Query
- PivotTables
- PivotCharts
- Conditional Columns
- Filtering
- Sorting
- Data Cleaning
- Transaction Categorization
- AML Flag Creation
- Data Aggregation
- Percentage Calculations
- Transaction Count Analysis
- Transaction Amount Analysis
- KYC-to-Transaction Comparison
- Transaction Drill-Down Analysis
- Dashboard Creation
- Data Visualization
- Summary Tables
- Investigation Documentation

---

# AML Skills Demonstrated

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

The customer was expected to conduct limited cash activity and disclosed no expected international activity. However, the investigation identified repeated cash deposits immediately below $10,000, substantially higher-than-expected cash activity, use of multiple branch locations, and subsequent international wire transfers.

The combination of these factors created multiple AML risk indicators warranting further investigation.

## Disposition

**Escalate for enhanced review and SAR consideration in accordance with institutional procedures.**

The analysis does not conclude that money laundering occurred. Instead, the identified transaction patterns warrant additional review to determine whether the activity has a reasonable explanation or requires further escalation.

---

# Workbook Structure

| Worksheet | Purpose |
|---|---|
| **Customer Profile** | KYC information and expected customer activity |
| **Transactions** | Original transaction-level dataset |
| **Transactions Table** | Cleaned and transformed analytical transaction data |
| **Analysis** | Transaction metrics, flagged activity, PivotTables, and visualizations |
| **Near-10K Review** | Detailed review of near-threshold cash deposits and international wires |
| **Analysis Summary** | Final findings, risk indicators, risk assessment, and disposition |

---

# Project Files

- `AML Analysis.xlsx`
- `AML Analysis 1 - Customer Profile.png`
- `AML Analysis 1 - Transactions.png`
- `AML Analysis 1 - Transactions Table.png`
- `AML Analysis 1 - Analysis.png`
- `AML Analysis 1 - Near 10K Review.png`
- `AML Analysis 1 - Analysis Summary.png`
- `README.md`

---

# Disclaimer

This project was created for **educational and portfolio purposes only**.

All customer names, account information, counterparties, transaction records, and financial activity shown in this project are **synthetic and do not represent real individuals or financial accounts**.
