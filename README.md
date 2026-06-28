<div align="center">

# 📊 Sales & Marketing Performance Data Model
### End-to-End Excel Analytics — Power Query · SUMIFS · Pivot Analysis · KPI Dashboard

![Excel](https://img.shields.io/badge/Microsoft_Excel-Advanced-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power_Query-ETL_Pipeline-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen?style=for-the-badge)

</div>

---

## 📌 Project Overview

A production-grade sales and marketing analytics model built entirely in Microsoft Excel, covering **725 raw transaction records** across **12 months (2024–2025)**, **4 U.S. regions**, **4 sales channels**, and **6 marketing channels**.

The model spans the full analytics lifecycle — from raw data ingestion and cleaning through multi-criteria formula logic, pivot analysis, and executive KPI reporting — demonstrating how Excel can serve as a complete, auditable analytics platform without external tooling.

---

## 🏆 Key Results at a Glance

| Metric | Value |
|---|---|
| 💰 Total Net Sales | **$506,167** |
| 📈 Gross Profit | **$210,147** |
| 🎯 Gross Margin | **41.5%** |
| 🛒 Total Valid Orders | **720** |
| 🧾 Average Order Value | **$703** |
| 📅 Months Above Target | **8 of 12 (66.7%)** |
| 📊 Best Month vs. Target | **+63.5% (October, +$23,617)** |
| 📦 Return Rate | **5.0%** |
| 🚚 Late Shipment Rate | **4.7%** |
| 🌱 Best Channel ROAS | **Organic — 1.51x blended; peak 8.28x** |

---

## 🗂️ What's Inside the Workbook

| Sheet | Purpose |
|---|---|
| `Raw_Orders` | 725 source transactions — untouched |
| `audit` | Data quality log (duplicates, nulls, invalid records) |
| `Cleaned data` | Standardised, enriched orders with 36 calculated fields |
| `Customers` | 165-row lookup table with loyalty tier and signup date |
| `Products` | 27 SKUs with standard cost, list price, supplier, and status |
| `Monthly_Targets` | 385-row target table by Month × Region × Category |
| `Marketing_Spend` | 481-row spend log by Month × Region × Channel |
| `KPI summary` | 7 headline KPIs via SUMIFS / COUNTIFS |
| `Pivot analysis` | 6-dimension pivot summaries |
| `Target variance analysis table` | Month-level actuals vs. targets with variance $ and % |
| `Dashboard` *(T09)* | One-page executive view with KPI cards and charts |

---

## ⚙️ Technical Implementation

### 1 — Data Pipeline (Power Query)

Extracted, cleaned, and standardised **725 raw records**, resolving:

- **5 duplicate Order IDs** (e.g., ORD-01212, ORD-01028, ORD-01119) — excluded from all KPI calculations
- **107 missing values** across Customer ID, Customer Rating, and other fields
- **10 inconsistent text values** — e.g., `Technlogy` → `Technology`, `EAST` → `East`, `Store` → `Retail Store`
- **3 SKU-to-category mismatches** — Laptops filed under Office Supplies, Air Purifier under Technology, etc.
- **6 invalid dates** and **4 invalid quantity records** (zero or negative) — flagged and excluded with a documented rule, preserving source rows

Final working dataset: **720 valid orders**.

---

### 2 — Calculated Fields (36 Columns)

All computed fields use formula logic against the cleaned table — no hardcoded values:

```excel
Gross Sales      = Quantity × Unit_Price
Discount Amount  = Gross_Sales × Discount_Pct
Net Sales        = Gross_Sales − Discount_Amount
COGS             = Quantity × Product_Cost          ← VLOOKUP from Products table
Gross Profit     = Net_Sales − COGS
Margin %         = Gross_Profit / Net_Sales          ← IFERROR-wrapped for $0 price records
Delivery Days    = Ship_Date − Order_Date
Late Shipment    = IF(Delivery_Days > 10, "late", "on time")
```

---

### 3 — Multi-Criteria Target Variance (SUMIFS)

Matched **720 cleaned orders** against a **385-row target table** using three simultaneous criteria to prevent double-counting:

```excel
Actual Sales =
  SUMIFS(
    Cleaned[Net_Sales],
    Cleaned[Order_Month], [@Month],
    Cleaned[Region],      [@Region],
    Cleaned[Category],    [@Category]
  )

Variance %   = (Actual_Sales − Target_Sales) / Target_Sales
```

**Result:** 8 of 12 months beat target. Full monthly breakdown:

| Month | Actual | Target | Variance $ | Variance % | Status |
|---|---|---|---|---|---|
| January | $52,560 | $32,968 | +$19,592 | **+59.4%** | ✅ Above |
| February | $34,530 | $33,383 | +$1,147 | +3.4% | ✅ Above |
| March | $42,403 | $37,183 | +$5,220 | +14.0% | ✅ Above |
| April | $22,605 | $37,539 | −$14,934 | **−39.8%** | ❌ Below |
| May | $46,016 | $37,180 | +$8,836 | +23.8% | ✅ Above |
| June | $44,497 | $37,415 | +$7,082 | +18.9% | ✅ Above |
| July | $34,361 | $37,418 | −$3,057 | −8.2% | ❌ Below |
| August | $46,656 | $38,159 | +$8,497 | +22.3% | ✅ Above |
| September | $42,464 | $37,784 | +$4,680 | +12.4% | ✅ Above |
| October | $60,812 | $37,194 | **+$23,617** | **+63.5%** | ✅ Above |
| November | $36,784 | $45,175 | −$8,392 | −18.6% | ❌ Below |
| December | $42,479 | $45,073 | −$2,593 | −5.8% | ❌ Below |

---

### 4 — Marketing ROI Table (3-Criteria SUMIFS + COUNTIFS)

Joined **actual revenue and order counts** into the Marketing Spend table using three criteria simultaneously:

```excel
Actual Revenue =
  SUMIFS(
    Cleaned[Net_Sales],
    Cleaned[Order_Month],       [@Month],
    Cleaned[Region],            [@Region],
    Cleaned[Marketing_Channel], [@Marketing_Channel]
  )

Total Orders =
  COUNTIFS(
    Cleaned[Order_Month],       [@Month],
    Cleaned[Region],            [@Region],
    Cleaned[Marketing_Channel], [@Marketing_Channel]
  )

ROAS = Actual_Revenue / Marketing_Spend
```

**Channel Performance Summary:**

| Marketing Channel | Total Spend | Total Revenue | Blended ROAS | Orders |
|---|---|---|---|---|
| Organic | $109,808 | $166,217 | **1.51x** | 224 |
| Email | $180,508 | $178,054 | 0.99x | 242 |
| Affiliate | $246,134 | $165,203 | 0.67x | 226 |
| Social | $315,962 | $134,523 | 0.43x | 224 |
| Paid Search | $459,759 | $186,544 | 0.41x | 252 |

> **Peak ROAS discovered:** Organic / West / October reached **8.28x** ($1,445 spend → $11,971 revenue) — the single highest-efficiency marketing combination in the dataset.

---

### 5 — Pivot Analysis (6 Dimensions)

Built dynamic summaries across Month, Region, Category, Customer Segment, Sales Channel, and Loyalty Tier. Key findings:

**By Category:**
| Category | Net Sales | Share |
|---|---|---|
| Technology | $247,266 | **48.9%** |
| Appliances | $111,860 | 22.1% |
| Furniture | $104,299 | 20.6% |
| Office Supplies | $42,742 | 8.4% |

> Laptops alone ($159,339) accounted for **31.5% of all net sales** — the largest single subcategory by a wide margin.

**By Region:**
| Region | Net Sales | Share |
|---|---|---|
| South | $147,097 | 29.1% |
| East | $127,658 | 25.2% |
| West | $125,578 | 24.8% |
| Central | $105,833 | 20.9% |

**By Customer Segment:**
| Segment | Net Sales | Share |
|---|---|---|
| Corporate | $208,261 | **41.1%** |
| Small Business | $165,774 | 32.8% |
| Consumer | $132,132 | 26.1% |

---

### 6 — KPI Dashboard

Seven headline KPIs built with formula logic (no typed values) and presented on a single executive-facing sheet:

```excel
Total Net Sales   = SUMIFS(Cleaned[Net_Sales], Cleaned[data quality], "Valid")
Gross Profit      = SUMIFS(Cleaned[Gross_Profit], Cleaned[data quality], "Valid")
Margin %          = Gross_Profit / Total_Net_Sales
Order Count       = COUNTIFS(Cleaned[data quality], "Valid")
AOV               = Total_Net_Sales / Order_Count
Return Rate       = COUNTIFS(Cleaned[Returned], "Yes") / Order_Count
Late Shipment %   = COUNTIFS(Cleaned[Late_Shipment], "late") / Order_Count
```

---

## 💡 Business Insights Discovered

**1. October Surge, April Gap**
October was the strongest month at $60,812 — 63.5% above its $37,194 target. April was the sharpest miss at −39.8% (−$14,934 below target). The four below-target months are concentrated in Q2 and Q4 end-periods, suggesting seasonal demand patterns that targets may not currently reflect.

**2. Laptop Concentration Risk**
A single SKU subcategory — Laptops — drives 31.5% ($159,339) of total net sales. This represents meaningful revenue concentration in one product line; a supply disruption or pricing shift would materially impact the P&L.

**3. Organic Is the Highest-Efficiency Marketing Channel**
Organic search returned a **1.51x blended ROAS** on $109,808 in spend — the only channel to consistently return more revenue than cost. The West region's Organic channel peaked at **8.28x ROAS** in October. Paid Search, by contrast, delivered just 0.41x on $459,759 in spend — the largest channel investment with the lowest return.

**4. Corporate Segment Is the Revenue Anchor**
Corporate customers account for 41.1% ($208,261) of all revenue despite representing one of three segments, implying an average order value significantly above the blended $703 AOV. The Loyalty Tier distribution (Gold 27.3%, Silver 26.6%, Bronze 23.3%, Platinum 22.9%) is unusually flat, suggesting the tier programme may not currently be differentiating customer behaviour effectively.

---

## 🔍 Data Limitations

1. **Lead count is a proxy.** The Lead-to-Order Conversion Rate uses a lead estimate from the Marketing Spend table rather than actual tracked leads. True conversion rates may differ if lead quality or volume varies by channel beyond what the data captures.

2. **3 orders with unknown return status** were excluded from return rate calculations. If these represent actual returns, the 5.0% rate is marginally understated.

3. **3 orders with missing Customer IDs** could not be enriched with loyalty tier or signup date data, slightly understating loyalty-tier revenue breakdowns.

---

## 🛠 Skills Demonstrated

`Power Query (ETL)` · `SUMIFS / COUNTIFS (multi-criteria)` · `VLOOKUP / XLOOKUP` · `IFERROR error handling` · `Pivot Tables` · `Calculated fields` · `Data validation` · `KPI dashboard design` · `Variance analysis` · `Marketing ROI modelling` · `Data cleaning & documentation`

---

<div align="center">

*Built with Microsoft Excel · Dataset: 725 transactions · Period: 2024–2025*

</div>
