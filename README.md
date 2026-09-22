AML Transaction Monitoring & Investigation Case Study
Project Overview

This project is a synthetic Anti-Money Laundering (AML) transaction monitoring and investigation case study built in Microsoft Excel and Power Query.

The objective was to review a customer's transaction history, compare observed activity against the customer's KYC profile, identify unusual or potentially suspicious transaction patterns, and document an investigation conclusion.

The analysis covered 1,374 transactions from March 1, 2026 through August 31, 2026.

The investigation identified:

10 near-$10,000 transactions totaling $93,280
7 cash deposits totaling $67,080
Cash deposits ranging from approximately $9,150 to $9,890
3 international wire transfers totaling $26,200
International wires sent to Orion Trading LLC in the United Arab Emirates
Cash deposits occurring across several Northern Virginia branch locations
Repeated cash-deposit-to-international-wire activity
Activity materially inconsistent with the customer's stated KYC profile
1. Customer Profile / KYC Review

The investigation begins with the customer's KYC profile to establish a baseline for expected account activity.

The synthetic customer is an IT Project Manager earning approximately $92,000 annually with a personal checking account.

Expected activity included:

Monthly credits of approximately $6,000–$10,000
Monthly debits of approximately $4,000–$9,000
Cash activity typically below $1,000 per month
No disclosed international activity
Primary use of mobile banking, debit card, and ACH
Initial customer risk rating of Low

This information was used throughout the investigation to compare expected behavior against actual transaction activity.

2. Raw Transaction Data

The Transactions worksheet contains the full raw transaction history used for the investigation.

The dataset includes fields such as:

Transaction ID
Transaction Date
Transaction Time
Account ID
Customer ID
Direction
Transaction Type
Amount
Currency
Counterparty
Counterparty Country
Channel
Branch City
Branch State
Merchant Category
Description
Balance After Transaction

The account contained 1,374 transactions across several transaction types including debit card purchases, ACH transactions, P2P transfers, ATM withdrawals, cash deposits, payroll, international wires, and cashier's checks.

3. Transaction Transformation & AML Flags

The raw transaction data was processed and organized into a structured analysis table using Excel and Power Query.

Additional fields were created to make transaction monitoring and aggregation easier.

Examples include:

Signed transaction amounts
Large transaction flags
Cash activity indicators
Near-$10K activity indicators
Transaction categorization
Transaction direction analysis

Power Query was used to transform the source data into a cleaner dataset suitable for AML analysis.

This separated the raw transaction data from the analytical dataset and created a repeatable workflow for reviewing the account.

4. Transaction Analysis Dashboard

The Analysis worksheet summarizes transaction activity by transaction type and highlights activity requiring additional review.

The analysis included:

Transaction count by transaction type
Total transaction amount
Flagged transaction count
Flagged transaction amount
Flagged transaction percentage
Cash deposit analysis
Near-$10K activity
International wire activity
Charts comparing total activity against flagged activity

The analysis identified cash deposits as the primary source of flagged activity.

All 7 reviewed cash deposits were flagged, totaling $67,080.

The customer had stated that cash activity was normally below approximately $1,000 per month, making the observed cash activity significantly inconsistent with the expected customer profile.

5. Near-$10K Transaction Review

A focused review was performed on transactions occurring close to the $10,000 threshold.

The review identified 10 transactions totaling $93,280 between July 6 and July 18, 2026.

The activity included:

7 cash deposits totaling $67,080
3 international wire transfers totaling $26,200

Cash deposits ranged from approximately $9,150 to $9,890 and occurred at multiple Northern Virginia branch locations including Alexandria, Arlington, Falls Church, Springfield, Fairfax, Annandale, and Tysons.

Three outbound international wires were sent to Orion Trading LLC in the United Arab Emirates:

$8,200
$8,800
$9,200

Total international wire activity reviewed:

$26,200

The repeated deposits immediately below $10,000, use of multiple branch locations, subsequent international wires, and short time intervals between transactions created indicators consistent with potential structuring and rapid movement of funds.

A single transaction alone would not establish suspicious activity. The concern resulted from the combined transaction pattern and inconsistency with the customer's KYC profile.

6. Final AML Investigation Summary

The final worksheet consolidates the customer profile, expected activity, reviewed transactions, key findings, risk indicators, and investigation disposition.

Key Findings

The investigation identified several notable risk indicators:

Repeated Near-$10K Cash Deposits

Seven cash deposits totaling $67,080 were made near but below the $10,000 threshold.

Activity Inconsistent With KYC Profile

The customer's expected cash activity was typically below $1,000 per month, while approximately $67,080 in cash deposits occurred during the reviewed transaction sequence.

International Wire Activity

The customer disclosed no expected international activity, yet three international wires totaling $26,200 were sent to the UAE.

Rapid Movement of Funds

Cash deposits were followed by outbound international wire transfers during the reviewed period.

Multiple Branch Locations

Cash deposits were conducted across several different branch locations.

Repeated Transaction Pattern

The behavior occurred multiple times rather than appearing as one isolated transaction.

Excel Functions & Techniques Used
COUNTIF

Used to count transactions meeting a single condition.

Example use cases:

Counting cash deposits
Counting transaction types
Counting flagged transactions

Example:

=COUNTIF(Transaction_Type_Range,"Cash Deposit")

COUNTIFS

Used when transactions had to satisfy multiple conditions.

Example use cases:

Counting near-$10K cash deposits
Counting transactions within a particular amount range
Counting transactions matching both transaction type and AML flag

Example:

=COUNTIFS(Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")

SUMIF

Used to calculate the total dollar amount associated with a particular transaction category.

Example:

=SUMIF(Transaction_Type_Range,"Cash Deposit",Amount_Range)

This was useful for calculating total activity for transaction types such as cash deposits, wires, ACH transactions, and card purchases.

SUMIFS

Used to calculate transaction amounts meeting multiple AML criteria.

Example:

=SUMIFS(Amount_Range,Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")

This allowed transaction counts and dollar amounts to be analyzed separately.

IF

Used to classify transactions based on transaction characteristics.

Example:

=IF(Amount>=10000,"Large Transaction","Normal")

IF statements were useful for creating transaction-level AML flags.

AND

Used with IF statements when multiple conditions had to be satisfied simultaneously.

Example:

=IF(AND(Transaction_Type="Cash Deposit",Amount>=8000,Amount<10000),"Near 10K","Normal")

This allowed transactions to be categorized based on combinations of transaction type and dollar amount.

Additional Excel Skills Used

The project also incorporated:

Power Query
Excel Tables
PivotTables
PivotCharts
Transaction filtering
Sorting
Conditional logic
AML flag creation
Data aggregation
Transaction categorization
Percentage calculations
KYC-to-transaction comparison
Transaction drill-down analysis
Dashboard creation
Data visualization
Investigation documentation
Risk indicator analysis
Power Query

Power Query was used to organize and transform the transaction dataset before performing the investigation.

The workflow included:

Loading the raw transaction dataset
Reviewing data types
Creating analytical columns
Creating transaction flags
Preparing transaction data for aggregation
Producing a structured transaction table for downstream Excel analysis

Using Power Query made the workflow more scalable than performing every transformation manually inside the raw dataset.

AML Skills Demonstrated

This project demonstrates practical experience with:

AML Transaction Monitoring
KYC Profile Review
Customer Activity Analysis
Transaction Pattern Analysis
Potential Structuring Identification
Near-Threshold Transaction Analysis
International Wire Review
Rapid Movement of Funds Analysis
Multiple-Branch Activity Review
AML Alert Investigation
Risk Indicator Identification
Excel-Based Data Analysis
Power Query
Investigation Documentation
SAR Escalation Consideration
Investigation Conclusion

The transaction activity reviewed was materially inconsistent with the customer's stated KYC profile.

The combination of repeated cash deposits immediately below $10,000, substantially higher-than-expected cash activity, multiple branch locations, undisclosed international activity, and subsequent outbound international wires created sufficient AML risk indicators to warrant additional investigation.

Disposition

Escalate for enhanced review and SAR consideration in accordance with institutional procedures.

The analysis does not conclude that money laundering occurred. Rather, the identified transaction patterns warrant additional review to determine whether the activity has a reasonable explanation or requires further escalation.

Project Structure

Customer Profile
KYC information and expected account behavior.

Transactions
Original transaction-level dataset.

Transactions Table
Cleaned and transformed analytical transaction dataset.

Analysis
Transaction counts, amounts, AML metrics, PivotTable analysis, and visualizations.

Near-10K Review
Focused investigation of near-threshold cash deposits and related international wire activity.

Analysis Summary
Customer comparison, key findings, red flags, risk assessment, and final investigation disposition.

Disclaimer

This project was created for educational and portfolio purposes only.

All customer names, transaction information, counterparties, account numbers, and financial activity shown in this project are synthetic and do not represent real individuals or financial accounts
