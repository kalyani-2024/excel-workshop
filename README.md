# Day 1 — Excel Crash Course

**From messy rows to a working dashboard.**

You start with 1,000 rows of scruffy sales data and finish with a one-screen dashboard that filters live. Work along in the file as you read — every formula here is copy-ready.

| | |
|---|---|
| **Level** | Beginner → confident |
| **You need** | Excel 2021 or 365 |
| **Working file** | `Day1_Sales_Raw.xlsx` |
| **You leave with** | Your own dashboard |

### What's covered

| Module | Topic |
|---|---|
| 1 | Data cleaning |
| 2 | Formulas & functions |
| 3 | XLOOKUP |
| 4 | PivotTables |
| 5 | Charts |
| 6 | Slicers |
| 7 | Conditional formatting |
| — | Dashboard exercise |

---

## Before you start

**Check your Excel version.** Type `=XLOOKUP(` into any cell. If Excel autocompletes it, you're fine. If it doesn't, you're on Excel 2019 or older and you'll use the `INDEX/MATCH` fallback given in Module 3. Excel for the web and Mac 365 both work.

**Open the file, don't copy the data out.** Download `Day1_Sales_Raw.xlsx` and work inside it. Three sheets: `Orders` (the mess), `Reps` (the lookup table), `Dashboard` (empty, for later).

**Turn on two things.** Formula bar and gridlines: `View` ribbon, both checkboxes. Then `File → Options → Formulas` and confirm calculation is **Automatic**. Half of all "my formula is broken" moments are this setting.

**Save a checkpoint now.** Press `F12` and save as `Day1_Working.xlsx`. You'll break something along the way, and reopening the raw file is much faster than undoing 40 steps.

---

## The dataset

A fictional consumer-electronics distributor: 1,000 orders across four regions for FY 2025–26. The `Orders` sheet is deliberately dirty.

### Sheet: Orders · A1:K1001

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| Order ID | Order Date | Rep ID | Region | City | Category | Product | Units | Unit Price | Discount | Revenue |
| ORD-1001 | 04-07-2025 | R-07 | North | Delhi | Audio | Bass Pods 2 | 12 | 4,299 | 0.10 | 46,429 |
| ORD-1002 | ⚠ 12-04-2025 | R-03 | ⚠ north | ⚠ `Delhi  ` | Wearables | PulseBand X | 4 | 7,999 | 0.00 | 31,996 |
| ORD-1003 | 18-07-2025 | R-11 | South | Chennai | Audio | Bass Pods 2 | ⚠ *(blank)* | 4,299 | 0.05 | ⚠ #VALUE! |
| ⚠ ORD-1003 | 18-07-2025 | R-11 | South | Chennai | Audio | Bass Pods 2 | 9 | 4,299 | 0.05 | 36,757 |
| ORD-1004 | 02-08-2025 | R-07 | ⚠ NORTH | Jaipur | Computing | Slate Air 13 | 2 | ⚠ ₹ 61,990 | 0.08 | ⚠ #VALUE! |
| ORD-1005 | 09-08-2025 | R-22 | West | Pune | Wearables | PulseBand X | 15 | 7,999 | 0.12 | 1,05,587 |

⚠ marks the six defects: a text-formatted date, inconsistent region casing, trailing spaces, a duplicated `Order ID`, a blank `Units` value, and a price typed with a currency symbol so Excel reads it as text.

### Sheet: Reps · A1:D25

| Rep ID | Rep Name | Home Region | Monthly Target |
|---|---|---|---|
| R-03 | Anjali Rao | North | 8,00,000 |
| R-07 | Devika Nair | North | 12,00,000 |
| R-11 | Imran Sheikh | South | 9,50,000 |
| R-22 | Farhan Qureshi | West | 7,00,000 |

24 reps in total. This is your lookup table for Module 3 — and the source of the **% of target** measure on the final dashboard.

---

## Module 1 — Clean the data first, always

Nothing downstream survives bad input. A PivotTable will happily report "North" and "north" as two separate regions and never warn you.

### Fix it in this order

1. **Kill merged cells.** Select all with `Ctrl+A`, then `Home → Merge & Center` to toggle every merge off. Merged cells break sorting, filtering and PivotTables. One header row, one cell per column.
2. **Remove duplicates.** `Data → Remove Duplicates`, tick only `Order ID`. Excel tells you how many rows it dropped — note that number, it's your record of what changed.
3. **Trim the whitespace.** Add a helper column, pull it down, then paste back as values:

   ```excel
   =TRIM(CLEAN(E2))
   ```
   `TRIM` strips leading, trailing and doubled spaces; `CLEAN` removes non-printing characters that arrive from exported systems.

4. **Normalise the casing** so "north", "North" and "NORTH" collapse into one value:

   ```excel
   =PROPER(TRIM(D2))
   ```

5. **Convert text dates to real dates.** Text dates sit left-aligned in the cell — that's your tell. Select column `B`, then `Data → Text to Columns → Next → Next → Date: DMY → Finish`. For stubborn cases:

   ```excel
   =DATEVALUE(B2)
   ```
   Then format the result as a date (`Ctrl+1`). A real date is a number; a text date is a wall you'll hit again in Module 4 when you try to group by month.

6. **Strip the currency symbol** out of `Unit Price`: `Ctrl+H`, find `₹`, replace with nothing, `Replace All`. Format the column as Number, not Currency — you'll apply currency formatting once, at the end, on the dashboard.
7. **Handle the blanks.** `Ctrl+G → Special → Blanks` highlights every empty cell at once. Then decide deliberately: delete the row, or enter `0`. Deleting silently is how you lose ₹3 lakh of revenue.
8. **Convert the range to a Table:** `Ctrl+T`. Name it `tblOrders` in `Table Design → Table Name`.

> **Why the Table matters**
> A Table auto-expands when rows are added, so your formulas, PivotTables and charts pick up new data without you re-pointing a single range. It also gives you readable formulas: `tblOrders[Revenue]` instead of `$K$2:$K$1001`. This one keystroke is what makes the dashboard refreshable.

> **Watch for**
> **Flash Fill** (`Ctrl+E`) looks magical for splitting names or cities — type the first result, press it, done. But it produces static text, not formulas. If the source data changes, Flash Fill output does not. Use it for one-off cleanup only.

---

## Module 2 — The eight functions that cover most work

Not eighty. These eight, used well, handle the overwhelming majority of day-to-day business spreadsheets.

### Build the calculated columns

```excel
=[@Units]*[@[Unit Price]]*(1-[@Discount])
```
Revenue, column `K`. Inside a Table, typing `[@` lets you pick columns by name, and the formula fills all 1,000 rows the moment you press `Enter`.

```excel
=TEXT([@[Order Date]],"mmm yyyy")
```
A **Month** label column. Useful for charts — but note it sorts alphabetically (Apr, Aug, Dec…). For correct chronological order, let the PivotTable group the real dates instead, as in Module 4.

### Conditional aggregation — the workhorses

```excel
=SUMIFS(tblOrders[Revenue], tblOrders[Region], "North", tblOrders[Category], "Audio")
```
Total Audio revenue in the North. Read it as: **sum this column, where that column equals this, and that other column equals that.** The sum range comes first; the criteria come in pairs after it.

```excel
=COUNTIFS(tblOrders[Region], "West", tblOrders[Units], ">=10")
```
Order count for bulk orders in the West. Comparison operators go inside the quotes.

```excel
=SUMIFS(tblOrders[Revenue], tblOrders[Order Date], ">="&DATE(2025,7,1), tblOrders[Order Date], "<="&EOMONTH(DATE(2025,7,1),0))
```
Revenue for July 2025. The `&` joins the operator to the date, because a criterion is always one text string. `EOMONTH` finds the last day of the month so you never hand-count 30 or 31.

### Logic and rounding

```excel
=IF([@Revenue]>=100000,"Large",IF([@Revenue]>=25000,"Medium","Small"))
=IFS([@Revenue]>=100000,"Large",[@Revenue]>=25000,"Medium",TRUE,"Small")
```
Identical results. `IFS` reads left to right and stops at the first `TRUE`, so order your tests from most to least restrictive. The final `TRUE` is the catch-all.

```excel
=ROUND([@Revenue],0)
```
Formatting a cell to 0 decimals only *displays* a rounded number — the underlying value keeps its decimals and your totals drift by a rupee or two. `ROUND` changes the value itself.

> **Absolute references in 15 seconds**
> `A1` moves when you copy it. `$A$1` never moves. `$A1` locks the column, `A$1` locks the row. Press `F4` while the cursor is on a reference to cycle through all four. You'll need `$A1` specifically for the conditional formatting rule in Module 7.

---

## Module 3 — XLOOKUP, and why VLOOKUP retires today

Pull the rep's name and target from the `Reps` sheet into `Orders`, so every row can be measured against quota.

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```
Three arguments required, three optional. Say it out loud as **"find this, in here, give me that."**

### The four you'll actually use

```excel
=XLOOKUP([@[Rep ID]], tblReps[Rep ID], tblReps[Rep Name], "Unknown rep")
```
**Exact match with a safety net.** The fourth argument replaces the `#N/A` that would otherwise litter your dashboard — and unlike `IFERROR`, it only catches genuine misses, not your own typos elsewhere in the formula.

```excel
=XLOOKUP([@[Rep ID]], tblReps[Rep ID], tblReps[[Rep Name]:[Monthly Target]])
```
**Return three columns at once.** Give the return argument a multi-column range and the result spills sideways into the neighbouring cells. VLOOKUP cannot do this at all.

```excel
=XLOOKUP([@Revenue], tblTiers[Floor], tblTiers[Tier], "", -1)
```
**Approximate match.** `-1` means "exact match, or the next smaller item" — the right behaviour for banding revenue into tiers or commission slabs. `1` takes the next larger instead. Your tier table needs a `Floor` column: 0, 25000, 100000.

```excel
=XLOOKUP(A2, tblOrders[Rep ID], tblOrders[Order Date], "No orders", 0, -1)
```
**Search bottom-up.** The last argument `-1` reverses the search direction, so on a date-sorted table this returns that rep's *most recent* order. Searching backwards is impossible in VLOOKUP.

### Why the upgrade is worth it

| | VLOOKUP | XLOOKUP |
|---|---|---|
| **Looks left** | No — key must be the leftmost column | Yes, any direction |
| **Column inserted** | Breaks silently, returns wrong data | Unaffected |
| **Default match** | Approximate — a classic source of silent errors | Exact |
| **Not found** | `#N/A`, needs wrapping | Built-in argument |
| **Search from bottom** | No | Yes |

The second row is the dangerous one: VLOOKUP counts columns by position, so inserting a column anywhere in the source shifts every result without raising an error.

> **On Excel 2019 or older**
> Use this equivalent everywhere the handbook says XLOOKUP. It looks left, and it survives inserted columns — the only thing it lacks is the `if_not_found` argument.
> ```excel
> =IFERROR(INDEX(tblReps[Rep Name], MATCH([@[Rep ID]], tblReps[Rep ID], 0)), "Unknown rep")
> ```

---

## Module 4 — PivotTables: summarise 1,000 rows in four drags

Everything you built with SUMIFS in Module 2, rebuilt in seconds — and re-sliceable without touching a formula.

### Your first pivot

1. Click any cell inside `tblOrders`, then `Insert → PivotTable → New Worksheet`. Rename that sheet `Pivots`.
2. Drag `Region` into **Rows**.
3. Drag `Revenue` into **Values**. It should read "Sum of Revenue" — if it says "Count of", you have text in that column and Module 1 isn't finished.
4. Drag `Category` into **Columns**. You now have a cross-tab.

> **The four boxes, in one line each**
> **Rows** — what you're breaking the numbers down by, going down.
> **Columns** — the same, going across. Use sparingly; two or three values at most.
> **Values** — the number being measured.
> **Filters** — a whole-pivot cutter. Here, slicers replace it.

### Five moves that make it presentable

- **Rename the ugly header.** Click the "Sum of Revenue" cell and type `Revenue ` — with a trailing space, because Excel refuses a name that matches the source field exactly.
- **Format the numbers once.** Right-click a value → `Number Format → Currency, 0 decimals`. Do it here, not by selecting cells — this way it survives a refresh.
- **Group dates into months.** Put `Order Date` in Rows, right-click any date → `Group → Months and Years`. Only works if Module 1 converted them to real dates.
- **Show a percentage.** Drag `Revenue` into Values a second time, right-click → `Show Values As → % of Grand Total`. Two views of the same number, side by side.
- **Top 5 only.** Right-click a row label → `Filter → Top 10` → change 10 to 5. Instant leaderboard.

### Build these three — you need them for the dashboard

| | Pivot |
|---|---|
| Pivot 1 | Revenue by Region |
| Pivot 2 | Revenue by Month |
| Pivot 3 | Top 5 Reps by Revenue |

> **Two things that trip everyone**
> **Pivots do not update themselves.** Change the source data and the pivot keeps showing the old numbers until you press `Alt+F5`, or `Data → Refresh All` for the lot.
> **GETPIVOTDATA** hijacks your formula when you click a pivot cell. Turn it off: `PivotTable Analyze → Options ▾ → untick Generate GetPivotData`.

---

## Module 5 — Charts that answer a question

Pick the chart from the question you're answering, not from the ribbon gallery.

| The question | The chart | Built from |
|---|---|---|
| Who's biggest? | Horizontal bar, sorted descending | Pivot 3 — Top 5 Reps |
| Is it going up? | Line, time on the x-axis | Pivot 2 — Revenue by Month |
| How do groups compare? | Clustered column | Pivot 1 — Revenue by Region |
| Two different units? | Combo with a secondary axis | Revenue bars + % of target line |
| Trend inside a table row? | Sparkline | `Insert → Sparklines → Line` |

Click any pivot, then `PivotTable Analyze → PivotChart` — the chart is now wired to the pivot and will react to slicers in Module 6.

### Clean up every chart

1. **Delete the gridlines.** Click one, press `Delete`. They compete with your data for attention.
2. **Delete the legend** when there's only one series. It's telling you something you already know.
3. **Add data labels** instead of making people read values off an axis: `Chart Design → Add Chart Element → Data Labels → Outside End`. Then you can usually delete the axis too.
4. **Sort bar charts** by value, not alphabetically. Sort the pivot and the chart follows.
5. **Write a real title.** "Revenue by Region" is a label. "North delivers 41% of revenue" is a finding.

> **Don't**
> No 3-D anything — the perspective distorts the very lengths you're asking people to compare. No pie charts beyond three slices. No dual axes unless the two series genuinely use different units; otherwise you can make any two lines cross wherever you like.

---

## Module 6 — Slicers: the part that makes it interactive

This is the single step that turns a page of charts into a dashboard someone else can use without asking you questions.

1. Click any PivotTable, then `PivotTable Analyze → Insert Slicer`. Tick `Region` and `Category`.
2. Add a **Timeline**: `Insert Timeline` → tick `Order Date` → switch the dropdown to **Months**. Timelines only accept real date fields.
3. **Wire each slicer to every pivot.** Right-click the slicer → `Report Connections` → tick all three pivots. Repeat for the second slicer and the timeline.
4. Style it: `Slicer → Columns: 2` so it sits in a tidy block, and drop the height to match your KPI row.

> **Don't skip step 3**
> Without **Report Connections**, each slicer controls only the pivot it was created from. Someone clicks "South", one chart changes, the other two don't, and the whole dashboard looks broken. Check all three boxes, on every slicer.

Test it: click **South** and confirm all three charts move together. Then `Clear Filter` (the funnel icon, top-right of the slicer) to reset.

---

## Module 7 — Conditional formatting, used with restraint

Colour should point at the exceptions. If everything is coloured, nothing is.

### Three built-ins worth knowing

- **Data bars** — `Home → Conditional Formatting → Data Bars`. An in-cell bar chart; best on a column of revenue beside its labels. Tick `Show Bar Only` to hide the numbers and keep it clean.
- **Colour scales** — good for a dense grid like month × region, poor for a single column where a bar reads faster.
- **Top/Bottom rules** — `Top 10%` on your rep list surfaces the outperformers without any sorting.

### The one that matters: a formula rule

Highlight the **entire row** of any rep who beat their target. Select `A2:E25` first — the selection defines the rule's scope — then `Conditional Formatting → New Rule → Use a formula`:

```excel
=$D2>=$C2
```

Write the formula **for the top-left cell of your selection only**; Excel applies the same logic to the rest relatively. The `$` before the column letters is what makes the whole row highlight — without it, each cell checks against itself and you get a diagonal stripe.

> **Rules for using colour**
> Pick one hue for "good" and one for "needs attention", and reuse them across the whole workbook. Never rely on colour alone — pair it with an icon or a text label, because roughly 1 in 12 men cannot distinguish your red from your green. Check your rules afterwards in `Conditional Formatting → Manage Rules`: duplicated and overlapping rules are the usual cause of a file that has become mysteriously slow.

---

## Exercise — Build the Day 1 dashboard

On the `Dashboard` sheet, using only what you built in the modules above. Everything you need already exists on the `Pivots` sheet — this is assembly, not invention.

### Required layout

```
┌─────────────────────────────────────────────────────────────┐
│  FY 2025–26 Sales Overview                                  │
├──────────────┬──────────────┬──────────────┬────────────────┤
│  KPI 1       │  KPI 2       │  KPI 3       │  KPI 4         │
│  Total Rev   │  Units Sold  │  Avg Order   │  Reps @ Target │
├──────────────┼──────────────┴──────┬───────┴────────────────┤
│  SLICERS     │  Chart A            │  Chart B               │
│  Region      │  Revenue by Month   │  Revenue by Region     │
│  Category    │  (line)             │  (column)              │
│  + Timeline  │                     │                        │
├──────────────┴─────────────────────┴────────────────────────┤
│  Chart C — Top 5 Reps (sorted bar + data bars on the table) │
└─────────────────────────────────────────────────────────────┘
```

Everything on one screen. If a reader has to scroll, it isn't a dashboard yet.

### The four KPI formulas

```excel
=SUM(tblOrders[Revenue])
=SUM(tblOrders[Units])
=AVERAGE(tblOrders[Revenue])
=COUNTIFS(tblReps[Achieved],">=1")&" of "&COUNTA(tblReps[Rep ID])
```

The fourth reads as **"9 of 24"**, where `Achieved` is revenue ÷ target. Note that these four are static totals — they won't respond to slicers, which is fine for a headline. Making KPIs follow the slicer needs a pivot or a CUBE function, which is beyond what this session covers.

### Finishing checklist

- [ ] Gridlines off — `View → uncheck Gridlines`. The single biggest visual upgrade available.
- [ ] All three charts are PivotCharts, all connected to all three slicers via Report Connections.
- [ ] No raw data on the dashboard sheet. Right-click the `Pivots` tab → Hide.
- [ ] Charts aligned to each other — select two, then `Shape Format → Align → Align Top`.
- [ ] Currency formatting applied through the pivot's Number Format, so it survives a refresh.
- [ ] One conditional formatting rule visible — data bars on the rep table.
- [ ] Sheets renamed, tab colours set, cursor parked on `A1` before the final save.
- [ ] Click a slicer, confirm all three charts move, then clear the filter.

### What a finished dashboard has

Use this to check your own work.

| | Weight |
|---|---|
| **Data is genuinely clean** — no duplicate order IDs, one spelling per region, dates are real dates, no text in numeric columns | 20 |
| **Four KPIs, correct and formatted** — right values, currency formatting, labelled so a stranger knows what they're looking at | 15 |
| **At least one XLOOKUP in use** — rep name or target pulled from the Reps table, with an `if_not_found` value set | 15 |
| **Three PivotCharts, each answering its question** — chart type matches the question; bars sorted; titles written, not defaulted | 20 |
| **Slicers connected to every pivot** — the interactivity test: one click moves the whole dashboard | 15 |
| **Conditional formatting that adds information** — applied to the exceptions, not to everything | 10 |
| **It fits on one screen** — no scrolling, no gridlines, nothing overlapping, supporting sheets hidden | 5 |
| **Total** | **100** |

Unfinished is fine and expected — you've had two hours. Save the file and keep it; it's a working template you can point at your own data by swapping what's in the `Orders` sheet and refreshing.

---

## Shortcuts

**Moving and selecting**
`Ctrl+↓` jump to the last row · `Ctrl+Shift+↓` select to it · `Ctrl+A` select the whole table · `Ctrl+Home` back to `A1` · `Ctrl+PgDn` next sheet

**Doing the work**
`Ctrl+T` make a Table · `Ctrl+Shift+L` toggle filters · `Alt+=` AutoSum · `Ctrl+E` Flash Fill · `Alt+F5` refresh a pivot · `F4` cycle `$` anchors

**Formatting**
`Ctrl+1` Format Cells · `Ctrl+Shift+4` currency · `Ctrl+Shift+5` percent · `Ctrl+Shift+V` paste special · `Alt+Enter` line break in a cell

## When it goes wrong

| Symptom | Cause and fix |
|---|---|
| "Sum of" says **Count** | Text hiding in your number column — often a stray space, or a `-` used for zero. Filter the column and look at the bottom of the list. |
| Dates won't group by month | They're text. Left-aligned by default is the giveaway. Go back to `Text to Columns` in Module 1, then refresh the pivot. |
| The slicer only moves one chart | Report Connections. Right-click each slicer and tick every pivot. It has to be done per slicer, not once for the sheet. |
| `#REF!` appeared everywhere | You deleted rows or columns a formula pointed at. `Ctrl+Z` immediately — `#REF!` destroys the original reference, and undo is the only way back. |
| The file crawls | Usually overlapping conditional formatting rules, or formulas pointing at whole columns (`A:A`) instead of a Table. Check `Manage Rules` first. |

---

**If you want to go further on your own:** Power Query handles repeatable cleaning — everything in Module 1, recorded once and replayed on next month's file with one click. From there, the data model and relationships let you combine tables without a single XLOOKUP.
