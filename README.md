# Zepto Sales Analysis (Excel)

An Excel-based analysis of quick-commerce (Zepto-style) order data covering **January to March 2025** across **Bangalore, Delhi and Mumbai**. The workbook holds the raw transaction data, pivot tables, charts, slicers and a KPI summary.

## Dataset at a glance

| Metric | Value |
|---|---|
| Orders | 1,498 (unique `Order_ID`s, no duplicates) |
| Customers | 300 |
| Products / SKUs | 200 |
| Categories | 10 |
| Date range | 01-01-2025 to 01-03-2025 (60 distinct dates) |
| Total sales | 563,332.74 |
| Average order value | 376.06 |

## Workbook structure

| Sheet | Contents |
|---|---|
| `Sheet1` | Raw order data (1,498 rows × 23 columns) |
| `Sheet2` | Pivot tables: highest-value orders, lowest-value (cancelled) orders, top and bottom customers by sales, breakdowns by city and payment method |
| `Sheet3` | KPI summary: total sales, average sales, total orders, customers, products, quantity |
| `Sheet5` | Category-wise sales share and distinct customers per city |
| `Sheet4` | Empty |

The workbook contains about 24 pivot tables, charts and slicers. Some summary cells use `UNIQUE` and `XLOOKUP`, so it needs **Excel 365 / 2021 or later**.

## Data dictionary (`Sheet1`)

| Column | Description |
|---|---|
| `Order_ID` | Unique order identifier |
| `Order_Date` | Order date, stored as text in `DD-MM-YYYY` |
| `Day`, `Days`, `Months` | Derived by formula: day of month, weekday name, month name |
| `Customer_Name` | Anonymised customer ID (`Cust_0` … `Cust_299`) |
| `Customer_Type` | Bulk, New or Returning |
| `Customer_Segment` | Office, Cloud Kitchen, Household, Tea Stall, Café or Bakery |
| `City` | Bangalore, Delhi or Mumbai (28 blanks) |
| `Payment_Method` | UPI, Netbanking, GPay, Card, PhonePe, Wallet or COD |
| `SKU`, `Product_Name` | Product identifiers (`SKU0` … `SKU199`) |
| `Category` | Beverages, Frozen, Veg, Fruits, Snacks, Dairy, Essentials, Other, Baby or Bakery |
| `Qty` | Units ordered |
| `Price` | Unit price |
| `Discount` | Discount per unit |
| `Subtotal` | `Qty × Price − Discount` |
| `Tax` | Tax on the order |
| `Delivery_Fee`, `Surge_Fee` | Delivery and surge charges |
| `Sales` | `Subtotal + Tax + Delivery_Fee + Surge_Fee` |
| `Delivery_Status` | Delivered, Cancelled, Failed or Returned |
| `Delivery_Time` | Delivery time in minutes |

## Key findings

- **Top categories by sales:** Beverages (14.7%), Frozen (13.4%) and Veg (13.1%). Bakery is the smallest (5.3%).
- **City-wise sales:** Bangalore leads (about 201K), followed by Delhi (about 181K) and Mumbai (about 169K).
- **Order outcomes:** 1,254 orders were delivered (about 84%). 114 were cancelled, 82 failed and 48 were returned. Sales tied up in non-delivered orders total about 94K.
- **Customer mix:** Bulk customers place 1,245 of the 1,498 orders.
- **Payments:** UPI is the most used method (491 orders), well ahead of Netbanking (186) and GPay (182).
- **Average delivery time:** about 25 minutes.

## Data notes and limitations

- 28 orders (about 1.9%) have no `City` and appear as `N/A` in the city pivot.
- March has only 33 orders (1 March only), so monthly comparisons should treat January and February as the full months.
- Currency is not specified in the data, so values are shown without a symbol.
- Customer, product and SKU names are placeholders, so the dataset appears to be sample or synthetic data.
- `Sales` includes orders that were cancelled, failed or returned. Filter on `Delivery_Status = Delivered` for realised revenue (about 468.9K).

## How to use

1. Open `zepto_sales_raw.xlsx` in Excel 365 or 2021+.
2. Go to `Sheet3` for headline KPIs and `Sheet5` for category and city views.
3. Use the slicers on the pivot sheets to filter by city, category and other fields.
4. To refresh after adding data, paste new rows into `Sheet1`, then use **Data → Refresh All**. The `Sheet3` formulas reference `U2:U1499`, so extend those ranges for new rows.

## Tools

Microsoft Excel: PivotTables, charts, slicers, `XLOOKUP`, `UNIQUE` and text/date functions.

## Possible next steps

- Compare delivery time and cancellation rate by city and customer segment.
- Analyse discount impact on order value.
- Build a dashboard sheet combining the main charts and slicers.
- Reproduce the analysis in Python (pandas) or SQL.

## Author

Yeshwanth
