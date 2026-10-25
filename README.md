# E-Commerce 360° Analytics — Excel + Power BI

## 1. What this project is

E-Commerce 360° is a business analytics project I built starting from raw exports — the kind of messy files you'd actually get from different parts of a business: customers, products, orders, payments, shipping, and returns.

The whole point was to go through the full workflow myself: take messy operational data, clean and standardize it in Excel, validate and reconcile the numbers, build an interactive Power BI dashboard on top of it, and then actually turn all of that into business insights instead of just charts.

## 2. Why this matters (the business problem)

In a real e-commerce business, useful information ends up scattered across a bunch of different files. And if those files have inconsistent formatting, blank cells, duplicate rows, or mismatched data types, you can't really trust whatever report comes out the other end.

So I followed a simple workflow: understand the raw data first, clean it in Excel, validate and reconcile everything, build the Power BI model, create the measures and visuals, and finally explain what the numbers are actually saying.

## 3. Where the data came from

I started with six raw operational exports:

- `customers_raw.csv` — customer master data
- `products_raw.csv` — product catalog
- `orders_raw.csv` — order transactions
- `payments_raw.txt` — payment export, pipe-delimited text
- `shipping_raw.xlsx` — shipping/logistics data
- `returns_raw.csv` — return records

These came from different operational systems, so nothing about them matched — different formats, different casing, different levels of completeness, even different file types.

## 4. The end-to-end workflow

```text
Raw Operational Files
        ↓
Understand columns and business meaning
        ↓
Excel Cleaning & Standardization
        ↓
Clean Excel Tables
        ↓
Checks, reconciliation & exploratory analysis
        ↓
Power BI Data Model
        ↓
DAX Measures / KPIs
        ↓
Interactive Dashboard
        ↓
Business Insights
        ↓
Business Questions & Next Steps

```

## 5. Step 1 — Understanding the raw data first

Before I calculated anything, I went through each source file and made sure I actually understood it:

- what one row represents
- which column is the key
- which fields are dates, numbers, categories, or IDs
- where the blanks, duplicates, or messy formatting are
- how the files connect to each other through the order data

I did this on purpose before touching any formulas — it's easy to start "cleaning" a column before you actually know what it means, and that's how you end up making things worse.

## 6. Step 2 — Cleaning each dataset from scratch in Excel

I didn't just dump all six files into one giant table. I cleaned each one separately first, and only connected them through shared keys afterward. I also kept the original raw files untouched and worked off cleaned copies for everything else.

### 6.1 Customers

`customers_raw.csv` → cleaned into `Customers_Clean`.

- Checked for duplicate `Customer_ID` values. Where a duplicate was an exact copy, I removed it. Where duplicate IDs had conflicting information, I kept one consistent record instead of just deleting one and hoping for the best.
- Cleaned up `Customer_Name`, `City`, and `State` — trimmed extra spaces and standardized everything to title case.
- Checked `Region` and it was already fine — no changes needed there.
- `Phone` and `Email` had some blanks. I left those as missing rather than making up data to fill the gaps.
- Standardized `Registration_Date` into one consistent date format.
- Rechecked the cleaned table one more time before using `Customer_ID` to link it to orders.

### 6.2 Products

`products_raw.csv` → cleaned into `Products_Clean`.

- Checked that `Product_ID` was actually unique before relying on it as the key.
- Went through product names and the main descriptive fields to make sure the product table could safely feed category/product info into the order data.
- Standardized the `Category` text so casing differences (like "electronics" vs "Electronics") didn't accidentally split into separate categories in the analysis. In Power BI I also normalized this field — trimmed and consistently cased — before using it in slicers and visuals.
- This cleaned product table became the lookup source for anything product- or category-level.

### 6.3 Orders

`orders_raw.csv` → cleaned into `Orders_Clean`. This was the main transaction table, so I spent the most time here.

- The raw file had 8,021 rows including the header. I found and removed 20 exact duplicate rows, which left 8,001 rows including the header.
- Converted `Order_Date` into one consistent date format.
- Cleaned up `Order_Status` and `City` for spacing and capitalization.
- `Unit_Price` had currency text mixed into it (things like `Rs.`), so I stripped that out so it could actually be used as a number.
- The only field with real blanks left after cleaning was `Discount`. I treated those blanks as zero for revenue purposes, since a blank discount just means no discount was applied — not that the data is broken.
- Re-ran the duplicate and blank checks after cleaning to make sure everything came back clean.

The main formula I used later for revenue:

```text
Net Revenue = Unit_Price × Quantity − Discount

```

For revenue analysis specifically, I excluded cancelled orders and zero-quantity rows — including them would have distorted the numbers.

### 6.4 Payments

`payments_raw.txt` came in as pipe-delimited text, so before I could clean anything I first had to split it into actual columns. Then it became `Payments_Clean`.

- Split the pipe-delimited file into proper columns.
- Standardized the payment method labels — for example, `Upi` became `UPI`.
- Cleaned up the `Amount` field so stray text/blanks wouldn't break the numeric analysis.
- Checked the keys and values before connecting payments back to orders.

### 6.5 Shipping

`shipping_raw.xlsx` → cleaned into `Shipping_Clean`.

- The date fields came in with inconsistent text/date formatting, so I had to check those carefully.
- Converted them into real Excel dates and applied one consistent display format across the board.
- Checked shipment IDs and delivery-status fields before using this table for operational analysis.
- Ended up with 7,248 shipment records. `Actual_Delivery_Date` is blank for 1,033 of them — those are in-transit or pending shipments, and I kept the blanks as they are rather than treating them like errors, since the delivery genuinely hasn't happened yet.

### 6.6 Returns

`returns_raw.csv` → cleaned into `Returns_Clean`.

- Checked `Return_ID` for uniqueness so every return could be counted safely.
- Kept returns as their own transaction table, linked back through `Order_ID`.
- 65 records had a blank `Return_Reason`. I left those as missing instead of guessing a reason — for the return-reason chart specifically, I excluded those blanks from that one visual, but didn't delete them from the actual dataset.
- Ended up with 646 return records total.

### 6.7 Final check before combining everything

Once all six datasets were cleaned individually, I went through them one more time together before using them side by side:

- checked keys for duplicates
- made sure dates were actual dates, not text
- confirmed numeric fields were usable for calculations
- confirmed category/status labels were standardized
- made sure expected blanks were still blanks, not filled with made-up values
- reconciled row counts before importing into Power BI

The final cleaned tables I ended up with:

- `Customers_Clean`
- `Products_Clean`
- `Orders_Clean`
- `Payments_Clean`
- `Shipping_Clean`
- `Returns_Clean`

## 7. Step 3 — Excel analysis and reconciliation

Once everything was clean, I used the workbook to answer the main business questions and sanity-check that the numbers behaved the way I'd expect.

The central revenue formula, again:

```text
Net Revenue = Unit Price × Quantity − Discount

```

Same rule as before — cancelled orders and zero-quantity rows are excluded from the main revenue analysis.

In Excel, I covered:

- monthly performance
- category performance
- regional performance
- top products
- customer behaviour
- returns
- shipping performance
- payment outcomes
- product-level profitability and margin

## 8. Step 4 — Building the Power BI data model

I imported the cleaned Excel tables into Power BI and connected them through the business keys.

The core relationships:

```text
Customers_Clean  → Orders_Clean
Products_Clean   → Orders_Clean
Orders_Clean     ↔ Payments_Clean
Orders_Clean     ↔ Shipping_Clean
Orders_Clean     ↔ Returns_Clean

```

Customers and Products feed descriptive info into Orders, while Payments, Shipping, and Returns extend the order-level analysis outward. I checked the model for duplicate keys and made sure the relationship cardinalities were actually correct before building anything on top of it.

## 9. Step 5 — Measures and KPIs

I built everything as Power BI measures so the dashboard reacts to filters properly instead of relying on hardcoded numbers.

The key measures:

- Total Orders
- Revenue-Eligible Orders
- Net Revenue
- Total Customers
- Average Order Value
- Units Sold
- Estimated Margin
- Margin %
- Returned Orders
- Return Rate %
- Total Shipments
- Late Deliveries
- Late Delivery %
- Cancelled Orders
- Cancelled + Successfully Paid Orders
- Amount Involved

## 10. Step 6 — Building the dashboard

The final report is split across five pages.

### Page 1 — Executive Overview

High-level KPIs, revenue trend, regional/category performance — the big picture.

### Page 2 — Product & Category Performance

Category contribution, top products, product performance, and drill-through detail.

### Page 3 — Customer & Returns

Repeat vs one-time customer revenue, customer KPIs, return rate by category, and return reasons.

### Page 4 — Operations & Risk

Shipping performance, late delivery rate, payment performance, cancellation metrics, and the cancelled-but-paid anomaly.

### Page 5 — Product Detail

A drill-through page — product info, revenue, units, orders, margin, return info, and order-status breakdown for whatever product you click into.

## 11. How the dashboard is meant to be used

I built this to be explored, not just looked at:

- Region slicer
- Date slicer
- Category slicer
- Order-status filtering where it's relevant
- Cross-filtering between visuals
- Product drill-through from the Top 10 chart
- Page navigation across the sections

## 14. Tools I used

- **Microsoft Excel** — cleaning the raw data, standardizing it, validation checks, pivot tables, calculations, reconciliation, exploratory analysis
- **Power BI** — data modeling, DAX measures, interactive filtering, drill-through, KPIs, and pulling it all together into a dashboard

## 📂 Project Files & Dashboard

The complete project files, including the raw dataset, cleaned datasets, Excel analysis workbook, and Power BI dashboard, are available below.

https://drive.google.com/drive/folders/1Tz6OJ0SVBzO_D3p9k9weAEPZaYe6WsAh?usp=sharing

The project explores e-commerce sales performance, customer behaviour, product categories, returns, and operational trends using Excel and Power BI.

