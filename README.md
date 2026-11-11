# Hotel_Booking_Excel_Analysis
Cleaned and analyzed 500 hotel bookings, built a dashboard with pivot tables and charts
Hotel Bookings Analysis (Excel)

Project Overview
A regional hotel manager asked for a review of booking performance across three
hotels (Chicago, Madison, Milwaukee) for **August–October 2026**. The raw data came
straight from the booking system and needed cleaning before analysis.
I cleaned the data, answered 7 business questions, and built a one-page dashboard.

Business Questions
1. What is total revenue and number of bookings?
2. Which city brings the most revenue?
3. Which room type earns the most — is it worth promoting Suites?
4. How is revenue trending month to month?
5. Which booking channel is growing?
6. Which day of the week is busiest, and which earns the most?
7. Do members spend more than non-members?

Dataset
| Sheet | Rows | Description |
|---|---|---|
| Bookings | 512 (raw) → 500 (clean) | Booking ID, guest, city, room type, channel, dates, amount |
| Guests | 180 | Guest ID, name, membership level, home state |

Data Cleaning Steps
- Kept an untouched **Raw_Data** copy before making any changes
- Removed **12 duplicate** bookings (512 → 500)
- Fixed extra spaces and inconsistent capitals using **TRIM** and **PROPER**, then Paste Values
- Found that cleaned columns still held formulas (not values) — converted them to values
- Converted **numbers stored as text** in the Amount column
- Replaced placeholder **"N/A"** cities with "Unknown"
- Flagged **5 missing amounts** in yellow instead of guessing values
- Kept a **cleaning log** of every change

Analysis Steps
- Added calculated columns: **Nights** (check-out − check-in), **Month**, **Day** (TEXT function)
- Brought in membership level from the Guests table using **XLOOKUP** on Guest_ID
- Built **pivot tables** for each question, using Sum, Count, Average and % of Total
- Created **KPI** formulas: total revenue, bookings, average per booking, nights
- Built a **dashboard** with pivot charts and a **City slicer**

 Key Results
| KPI | Value |
|---|---|
| Total Revenue | $146,046 |
| Total Bookings | 500 |
| Avg Revenue per Booking | $295 |
| Avg Nights per Stay | 2.0 |

Key Insights
- **Chicago** generates about **50%** of total revenue.
- **Suites** are only 17% of bookings but bring **27% of revenue** (~$458 per booking vs $278 King, $245 Double).
- Revenue **dipped in September** (even per day) and **October** was the best month.
- **Website bookings nearly doubled** (55 → 98) and became over half of October bookings, while walk-ins declined every month.
- **Saturday** has the most bookings, but **Friday earns ~39% more revenue** because Friday guests stay longer (2.1 vs 1.4 nights).
- **Members spend more per booking** than non-members (Silver $311, Gold $298, None $285).

Recommendations
- Invest in the **website and online promotions**, the fastest-growing channel.
- Promote **Suites**, which earn the most per booking.
- Offer **weekend packages** to turn Saturday one-night stays into two-night stays.
- Plan **September promotions** to soften the seasonal dip.
- Encourage non-members to **join the loyalty program**.

What I Learned
- Never sort data in the middle of cleaning — it can misalign columns.
- A cell can look like text but secretly contain a formula — always check the formula bar.
- Count rows with an ID column, and look up by **ID, not name**.
- Compare groups using **averages or rates**, not just totals.
- Cross-check KPIs against each other (average × count = total).

Tools & Skills
Excel · Data Cleaning · TRIM / PROPER · XLOOKUP · SUMIFS / COUNTIFS · IFS ·
TEXT (dates) · Pivot Tables · Pivot Charts · Slicers · Conditional Formatting · Dashboard Design

Files
- `Hotel_Project_Jhansi.xlsx` — full workbook (raw data, cleaned data, pivots, dashboard)
- `dashboard.png` — screenshot of the final dashboard

Dashboard
![Dashboard](dashboard.png)

Author
**Jhansi Rapeti** — aspiring Data Analyst
