# 📊 SaaS Customer Revenue & Churn Analytics Project

A comprehensive data analytics project analyzing subscriber retention, revenue trends (MRR), unit economics (CLV vs. CAC), and behavioral churn risk indicators for a B2B/B2C SaaS company across a 4-year period (2022–2025).

---

## 📌 Project Overview

This project evaluates subscription performance, financial sustainability, and customer attrition risks. By combining historical subscription data with monthly revenue records, the project identifies key drivers of churn, pinpoints high-risk customer segments, evaluates plan profitability, and establishes early warning indicators to proactively prevent customer churn.

---

## 📁 Dataset Structure

The analysis is powered by two primary datasets:

1. **`subscriptions.csv` / `Subscriptions` Sheet (600 Records):**
   * **`customer_id`**: Unique subscriber identifier (`CUST-1001` to `CUST-1600`).
   * **`plan`**: Subscription tier (`Starter`, `Professional`, `Business`, `Enterprise`).
   * **`billing_cycle`**: Commitment terms (`Monthly` vs. `Annual`).
   * **`company_size`**: Employee brackets (`1-10`, `11-50`, `51-200`, `201-500`, `500+`).
   * **`acquisition_channel`**: Channel origin (`Organic Search`, `Paid Ads`, `Social Media`, `Referral`, `Partner`, `Direct Sales`).
   * **`monthly_price`**: Recurring plan cost.
   * **`signup_date` & `churn_date`**: Customer lifecycle dates.
   * **`churned`**: Subscription status (`Yes` / `No`).
   * **`churn_reason`**: Primary exit driver (e.g., `Budget Cuts`, `Price Too High`, `Company Closed`).
   * **`support_tickets_12mo`**: Support volume over 12 months.
   * **`nps_score`**: Net Promoter Score (1–10).
   * **`feature_usage_pct`**: Product feature adoption percentage (0–100%).

2. **`monthly_revenue.csv` / `monthly_revenue` Sheet (48 Months):**
   * **`month`**: Monthly recording period (Jan 2022 – Dec 2025).
   * **`active_customers`**: Total active subscribers at month-end.
   * **`new_customers`**: Newly acquired customers during the month.
   * **`churned_customers`**: Customers lost during the month.
   * **`total_mrr`**: Total Monthly Recurring Revenue generated ($).
   * **`cac`**: Customer Acquisition Cost per user ($).

---

## 💡 Key Analytical Findings & Answers

### 1. Overall Churn & Historical Trend
* **Overall Churn Rate:** **52.17%** (313 total churned out of 600 subscribers).
* **Trendline Direction:** Monthly churn has trended **upward over the 4-year period**. 
  * In **2022**, monthly churn peaked at 6 customers (July 2022).
  * By **2025**, monthly churn spiked significantly to **17 customers in April 2025**.
* **Verdict:** Churn is **getting worse over time**. The business faces an accelerating revenue leak if immediate retention interventions are not deployed.

---

### 2. Segment & Billing Churn Breakdown
* **Highest-Risk Plan:** **Starter Plan (70.51% churn)**, followed by Professional (47.98%), Business (41.25%), and Enterprise (22.00%).
* **Billing Impact:** **Monthly billing (60.51% churn)** suffers a far higher attrition rate than **Annual billing (40.32% churn)** due to lower switching friction.
* **Highest-Risk Segment Intersection:** **Starter Plan on Monthly Billing** exhibits a **76.87% churn rate** (103 churned out of 134 accounts).

---

### 3. Top Reasons for Customer Churn
 across the dataset, the top 3 drivers for cancellations are:
1. 🥊 **Budget Cuts:** 53 customers
2. 💸 **Price Too High:** 51 customers
3. 🏢 **Company Closed:** 48 customers

#### Differences by Plan vs. Company Size:
* **By Plan:** Reasons align closely with overall results—Starter accounts churn heavily due to *Budget Cuts* and *Price Too High*.
* **By Company Size:** Drivers diverge—smaller and mid-sized companies churn primarily due to **No Longer Needed**, **Company Closed**, and **Poor Support**.

---

### 4. Unit Economics (CLV vs. CAC Analysis)

* **Average Acquisition Cost (CAC):** **$200.04** (derived across all 48 historical months).
* **Benchmark Target:** A healthy SaaS business targets a **CLV:CAC ratio ≥ 3.0x**.

| Plan | Avg MRR ($) | Churn Rate (%) | Lifespan (Months) | Customer Lifetime Value (CLV) | CAC ($) | CLV:CAC Ratio | Financial Evaluation |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **Starter** | $215.54 | 70.51% | 1.42 | **$305.69** | $200.04 | **1.53 : 1** | ⚠️ **Unprofitable / High Risk** |
| **Professional** | $497.04 | 47.98% | 2.08 | **$1,035.93** | $200.04 | **5.18 : 1** | ✅ **Healthy** |
| **Business** | $1,303.64 | 41.25% | 2.42 | **$3,160.33** | $200.04 | **15.80 : 1** | ✅ **Highly Profitable** |
| **Enterprise** | $2,984.99 | 22.00% | 4.55 | **$13,568.15** | $200.04 | **67.83 : 1** | 🚀 **Top Performer** |

* **Most Profitable:** **Enterprise Plan ($13,568.15 CLV | 67.83:1 ratio)** due to high MRR and long retention.
* **Least Profitable:** **Starter Plan ($305.69 CLV | 1.53:1 ratio)**—barely covers acquisition costs before churning out at 1.42 months.

---

### 5. At-Risk Indicators & Predictive Thresholds

By correlating customer behavior with past cancellations:
* **Feature Usage Impact:** Churned customers averaged **27.45% feature usage** vs. **55.02% for active accounts**. Zero customers with usage >50% churned.
* **NPS Impact:** Churned customers averaged an **NPS score of 3.04** (Detractors) vs. **5.81 for active accounts**.
* **At-Risk Threshold Defined:** Active customers with **`feature_usage_pct < 35%`**.
* **Current Vulnerable Base:** Exactly **67 active customers** (23.3% of current active subscribers) fall below this threshold and require immediate Customer Success intervention.

---

## 🛠️ Step-by-Step Implementation Guide

### Excel Formulas (2016 Edition)

1. **Helper Churn Flag Column (`Subscriptions!R2`):**
   ```excel
   =IF(K2="Yes", 1, 0)
   ```
2. **Average Customer Lifespan (`LifeSpan(Month)` in cell `D2`):**
   ```excel
   =1 / (C2 / 100)
   ```
3. **Customer Lifetime Value (`CLV` in cell `E2`):**
   ```excel
   =B2 * D2
   ```
4. **CLV:CAC Ratio (`CLV:CAC` in cell `F2`):**
   ```excel
   =E2 / 200.04
   ```
5. **Active At-Risk Customer Flag (`Subscriptions!S2`):**
   ```excel
   =IF(AND(K2="No", P2<35), "At-Risk", "Healthy")
   ```
6. **Total At-Risk Count Summary:**
   ```excel
   =COUNTIF(Subscriptions!S2:S601, "At-Risk")
   ```

---

## 📈 Strategic Recommendations for Management

1. **Fix or Pivot the Starter Plan:**
   * Introduce annual-only billing options or minimum 3-month commitments for low-tier plans.
   * Improve initial onboarding sequences to push `feature_usage_pct` above 35% within the first 14 days.
2. **Double Down on High-Value Tiers:**
   * Reallocate marketing acquisition spend away from low-intent channels toward **Direct Sales** targeting **Enterprise** and **Business** accounts.
3. **Automate Churn Early-Warning Systems:**
   * Deploy automated alerts in CRM/CS platforms when an active subscriber’s feature usage drops below 35% or NPS score falls to ≤ 4.

---

## 🚀 How to Run the Analysis

1. Clone or download this repository.
2. Open `Project_dataset_cleaned.xlsx` in Microsoft Excel (2016 or newer).
3. Navigate to the `Subscriptions` sheet to inspect raw customer records and risk flags.
4. Review calculated Pivot Tables and summary metrics in the `Churn Analysis` and `Unit Economics` sheets.