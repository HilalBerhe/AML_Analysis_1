AML Transaction Monitoring & Investigation — Excel Case Study

Project Overview

This project is a synthetic Anti-Money Laundering (AML) transaction-monitoring and investigation case study built in Microsoft Excel and Power Query. The objective was to review a customer's transaction history, compare observed activity with the customer's KYC profile, identify unusual patterns, and document an investigation-style disposition.

The review covered 1,374 transactions from March 1 through August 31, 2026. The analysis identified repeated cash deposits just below the $10,000 threshold, activity across multiple branch locations, and international wire transfers inconsistent with the customer's expected profile.

Disclaimer: All names, account information, counterparties, transaction data, and findings in this project are entirely synthetic and were created for educational and portfolio purposes.

1. Customer Profile / KYC Review



The investigation begins with the customer's KYC profile. The customer is a low-risk personal checking customer with annual income of $92,000, expected monthly credits of $6,000–$10,000, expected monthly debits of $4,000–$9,000, and expected cash activity typically below $1,000 per month. No international activity was disclosed.

This profile establishes the expected baseline used to compare the customer's actual transaction behavior.

2. Raw Transaction Data



The raw transaction history contains the underlying account activity reviewed during the investigation. Fields include transaction date and time, account and customer IDs, transaction direction, transaction type, amount, counterparty, country, channel, branch location, merchant category, description, and balance after the transaction.

The dataset contains 1,374 transactions and includes normal day-to-day activity such as card purchases, ACH transactions, P2P transfers, ATM withdrawals, payroll credits, cash deposits, wire transfers, and cashier's checks.

3. Power Query / Transaction Table



The transaction data was cleaned and transformed into a structured analysis table using Excel Power Query. Additional analytical fields were created to support AML screening and transaction-level review.

Examples of transformations and analysis fields include:

Signed transaction amount

Large-transaction flag

Near-$10K transaction identification

Transaction type classification

Direction-based credit/debit analysis

Country and international activity review

Cash transaction identification

Sorting transactions chronologically

Filtering transaction populations for targeted investigation

The structured table made it easier to apply Excel formulas, PivotTables, filters, and AML review logic consistently across the full transaction population.

4. AML Analysis Dashboard



The analysis sheet summarizes transaction activity by type and measures both total activity and flagged activity.

Key Metrics Identified

1,374 total transactions reviewed

10 near-$10K transactions totaling $93,280

7 cash deposits totaling $67,080

3 international wires totaling $26,200

Cash deposits represented the primary source of flagged activity

The dashboard uses PivotTables, summary calculations, and charts to compare total transaction activity against flagged transaction activity and to highlight the transaction categories contributing most to the investigation.

5. Near-$10K Transaction Review



A targeted review was performed on transactions occurring between July 6 and July 18, 2026.

The review identified 10 transactions totaling $93,280, consisting of:

7 cash deposits totaling $67,080

3 outbound international wire transfers totaling $26,200

The cash deposits ranged from approximately $9,150 to $9,890 and were made across multiple Northern Virginia branch locations. Three outbound wires — $8,200, $8,800, and $9,200 — were sent to Orion Trading LLC in the United Arab Emirates.

The activity was compared directly with the KYC profile. Expected cash activity was typically below $1,000 per month, while observed cash deposits totaled $67,080 in approximately 12 days. The customer also disclosed no expected international activity.

Risk Indicators Reviewed

Repeated cash deposits below the $10,000 threshold

Significant deviation from expected cash activity

Multiple branch locations

Outbound international wires

Cash-to-wire transaction sequences

Rapid movement of funds

Repeated pattern over a short period

These indicators do not independently establish suspicious activity, but together they warranted additional investigation and escalation under the assumptions of this synthetic case.

6. Final Investigation Summary



The final summary consolidates the customer's KYC profile, expected activity, reviewed activity, key findings, AML risk indicators, and investigative disposition.

Key Findings

Repeated near-$10K cash deposits totaling $67,080

Three international wires totaling $26,200

Repeated cash-deposit-to-international-wire sequences

Observed cash activity materially exceeded the customer's expected profile

International activity was inconsistent with the customer's disclosed activity

Cash deposits occurred across multiple branch locations

Investigation Conclusion

The reviewed activity was inconsistent with the customer's expected cash activity and disclosed international activity. The combination of repeated near-threshold cash deposits, multiple branch locations, subsequent outbound international wires, and rapid movement of funds presented indicators associated with potential structuring and layering/rapid movement of funds.

Disposition: Escalate for enhanced review and SAR consideration in accordance with institutional policies and procedures.

Excel Functions & Techniques Used

Excel Formulas

COUNTIF — counted transactions meeting a single AML criterion

COUNTIFS — counted transactions meeting multiple conditions

SUMIF — calculated transaction amounts for a specific category or flag

SUMIFS — aggregated transaction amounts using multiple AML criteria

IF — created transaction-level classifications and flags

AND — combined multiple conditions within AML flagging logic

Percentage calculations — calculated flagged transaction rates and category-level exposure

Excel / Data Analysis Features

Power Query for data cleaning and transformation

Excel Tables for structured transaction data

PivotTables for transaction-type aggregation

PivotCharts / Excel Charts for visual analysis

Filters and sorting for targeted transaction review

Conditional / rule-based flags for AML screening

KYC-to-transaction profile comparison

Transaction-level drill-down analysis

Cash-flow pattern analysis

International transaction review

Branch-location analysis

Dashboard and investigation-summary reporting

Example AML Logic

Examples of the types of Excel logic used in the project include:

=COUNTIF(Transaction_Type_Range,"Cash Deposit")

=COUNTIFS(Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")

=SUMIFS(Amount_Range,Transaction_Type_Range,"Cash Deposit",Amount_Range,">=8000",Amount_Range,"<10000")

=IF(AND(Transaction_Type="Cash Deposit",Amount>=8000,Amount<10000),"Near 10K","Normal")

These formulas were used alongside Power Query transformations and PivotTable analysis to identify and summarize transactions requiring additional review.

Skills Demonstrated

AML transaction monitoring

KYC profile analysis

Transaction investigation

Identification of potential structuring indicators

Rapid movement of funds analysis

International wire review

Excel data analysis

Power Query

PivotTables and charts

COUNTIF / COUNTIFS

SUMIF / SUMIFS

IF / AND logic

Data cleaning and transformation

Investigative documentation

Risk assessment and escalation rationale

Repository Structure

AML-Transaction-Monitoring-Case-Study/
│
├── README.md
├── workbook/
│   └── AML-Analysis.xlsx
│
└── images/
    ├── 01-customer-profile.png
    ├── 02-transactions.png
    ├── 03-transactions-table.png
    ├── 04-analysis.png
    ├── 05-near-10k-review.png
    └── 06-analysis-summary.png

Portfolio Purpose

This project demonstrates the use of Excel, Power Query, transaction analysis, and AML investigative reasoning to move from raw transaction data to a documented case conclusion. It is intended to demonstrate practical skills relevant to entry-level AML Analyst, AML Investigator, Transaction Monitoring Analyst, Financial Crime Analyst, and Financial Analyst roles.
