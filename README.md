# Day 1 — Excel Crash Course

**From a messy measurement file to a working dashboard.**

You start with 1,000 rows of raw test results and finish with a one-screen dashboard that filters live and states what the data actually shows. Work along in the file as you read — every formula here is copy-ready.


### How to use this handbook

Sections tagged **`OPTIONAL`** are worth knowing but aren't needed to finish the exercise. Skip them while working along; read them afterwards.

### What's covered

| Module | Topic |
|---|---|
| — | How an analysis dataset is shaped |
| 1 | Data cleaning |
| 2 | Formulas & functions |
| 3 | Describing the data, and finding anomalies |
| 4 | XLOOKUP |
| 5 | PivotTables |
| 6 | Charts |
| 7 | Slicers |
| 8 | Conditional formatting |
| 9 | Turning numbers into findings |
| — | Dashboard exercise |

---

## Before you start

**Check your Excel version.** Type `=XLOOKUP(` into any cell. If Excel autocompletes it, you're fine. If it doesn't, you're on Excel 2019 or older and you'll use the `INDEX/MATCH` fallback given in Module 4. Excel for the web and Mac 365 both work.

**Open the file, don't copy the data out.** Download `Day1_BatteryTest_Raw.xlsx` and work inside it. Three sheets: `Runs` (the mess), `Cells` (the lookup table), `Dashboard` (empty, for later).

**Turn on two things.** Formula bar and gridlines: `View` ribbon, both checkboxes. Then `File → Options → Formulas` and confirm calculation is **Automatic**. Half of all "my formula is broken" moments are this setting.

**Save a checkpoint now.** Press `F12` and save as `Day1_Working.xlsx`. You'll break something along the way, and reopening the raw file is much faster than undoing 40 steps.

---

## How an analysis dataset is shaped

Almost every dataset worth analysing has the same four kinds of column. Learn to spot them and any new file becomes readable in about thirty seconds.

| Role | What it is | In this file |
|---|---|---|
| **Identifier** | One row, one thing. Unique, never analysed. | `Run ID` |
| **Group** | Which category this row belongs to. A handful of repeated text values. | `Cell Model`, `Rig ID` |
| **Condition** | What was set or varied deliberately. Usually a number with a few levels. | `Charge Rate`, `Ambient Temp` |
| **Measure** | What came out. The number you're actually studying. | `Capacity Retention`, `Internal Resistance` |

The whole of analysis is: **split the measures by the groups and conditions, compare, and explain the gaps.**

Two consequences that drive everything below:

- **Groups go on chart axes and in slicers. Measures go in Values.** Never the other way round.
- **Measures are averaged, not added.** A total of every capacity reading is a meaningless number. This is the most common way a good-looking dashboard ends up wrong — see Module 5.

> **It's the same shape everywhere**
> Rows might be customers, machines, field plots, survey respondents, delivery routes or lab samples. Rename the columns and the analysis is identical. Nothing you learn today is specific to batteries.

### Repeats and controls

Two ideas from experimental work that make any dataset easier to reason about:

- **Replicates** — the same condition tested more than once. Here, every combination of cell model and charge rate is run on several rigs. Replicates are why you report an *average* and a *spread*, never a single reading.
- **Control** — the baseline condition everything else is compared against. Here it's **0.5C at 25 °C**, the gentlest setting. A result only means something relative to a baseline: "82% retention" says little; "12 points below control" is a finding.

---

## The dataset

A fictional battery test lab. Three cell models were run to 500 charge cycles at four charge rates and four ambient temperatures, across eight test rigs — 1,000 runs. The `Runs` sheet is deliberately dirty.

### Sheet: Runs · A1:K1001

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| Run ID | Test Date | Cell Model | Rig ID | Charge Rate | Ambient Temp | Cycles | Capacity Retention % | Internal Resistance | QC Flag | Notes |
| RUN-1001 | 04-07-2025 | Model B | RIG-02 | 1.0 | 25 | 500 | 91.4 | 38.2 | Pass | |
| RUN-1002 | ⚠ 12-04-2025 | ⚠ model b | RIG-02 | 1.0 | 25 | 500 | 90.8 | 39.1 | Pass | |
| RUN-1003 | 18-07-2025 | Model A | ⚠ `RIG-04  ` | 2.0 | 45 | 500 | ⚠ *(blank)* | 61.7 | Pass | |
| ⚠ RUN-1003 | 18-07-2025 | Model A | RIG-04 | 2.0 | 45 | 500 | 74.2 | 61.7 | Pass | |
| RUN-1004 | 02-08-2025 | ⚠ MODEL C | RIG-07 | 0.5 | 25 | 500 | ⚠ 104.6 | 31.0 | Pass | rig recalibrated |
| RUN-1005 | 09-08-2025 | Model C | RIG-07 | 1.5 | 35 | 500 | 86.9 | ⚠ `42.5 mΩ` | Pass | |

⚠ marks the six defects: a text-formatted date, inconsistent model casing, trailing spaces, a duplicated `Run ID`, a blank measurement, an impossible value (retention above 100%), and a number typed with its unit so Excel reads it as text.

### Sheet: Cells · A1:D4

| Cell Model | Chemistry | Rated Capacity (mAh) | Spec Floor % |
|---|---|---|---|
| Model A | NMC | 4,200 | 80 |
| Model B | LFP | 3,600 | 85 |
| Model C | NMC | 5,000 | 80 |

Your lookup table for Module 4, and the source of the **pass/fail against spec** measure on the final dashboard.

### What you're trying to find out

Three questions, which is what a dashboard should be built to answer:

1. Does charging faster degrade cells more? By how much?
2. Do the three models behave differently — and is the difference bigger than the noise between repeat runs?
3. Is anything in this data wrong, rather than merely surprising?

---

## Module 1 — Clean the data first, always

Nothing downstream survives bad input. A PivotTable will happily report "Model B" and "model b" as two separate models and never warn you.

### Fix it in this order

1. **Kill merged cells.** Select all with `Ctrl+A`, then `Home → Merge & Center` to toggle every merge off. Merged cells break sorting, filtering and PivotTables. One header row, one cell per column.
2. **Remove duplicates.** `Data → Remove Duplicates`, tick only `Run ID`. Excel tells you how many rows it dropped — note that number, it's your record of what changed.
3. **Trim the whitespace.** Add a helper column, pull it down, then paste back as values:

   ```excel
   =TRIM(CLEAN(D2))
   ```
   `TRIM` strips leading, trailing and doubled spaces; `CLEAN` removes non-printing characters that arrive from exported instruments.

4. **Normalise the casing** so "model b", "Model B" and "MODEL B" collapse into one value:

   ```excel
   =PROPER(TRIM(C2))
   ```

5. **Convert text dates to real dates.** Text dates sit left-aligned in the cell — that's your tell. Select column `B`, then `Data → Text to Columns → Next → Next → Date: DMY → Finish`. For stubborn cases:

   ```excel
   =DATEVALUE(B2)
   ```
   Then format the result as a date (`Ctrl+1`). A real date is a number; a text date is a wall you'll hit again in Module 5.

6. **Strip units out of numbers.** `Internal Resistance` contains entries like `42.5 mΩ`. `Ctrl+H`, find ` mΩ`, replace with nothing, `Replace All`. Put the unit in the column header where it belongs, never in the cells.
7. **Handle the blanks.** `Ctrl+G → Special → Blanks` highlights every empty cell at once. Then decide deliberately — and for a *measurement*, the decision is different from a sales figure:

   > A blank measurement is **missing**, not zero. Entering `0` tells Excel that cell retained none of its capacity, which will drag every average down and invent a catastrophic failure that never happened. Leave it blank, or delete the row. `AVERAGE` skips blanks correctly; it does not skip zeroes.

8. **Catch impossible values.** Capacity retention above 100% or below 0 is physically meaningless — an instrument or entry error. Flag them before they reach a chart:

   ```excel
   =IF(OR([@[Capacity Retention]]>100,[@[Capacity Retention]]<=0),"Impossible","OK")
   ```
   Do this for every measure, using whatever bounds the quantity actually has. Percentages sit in 0–100, physical quantities are rarely negative, counts are whole numbers.

9. **Convert the range to a Table:** `Ctrl+T`. Name it `tblRuns` in `Table Design → Table Name`.

> **Why the Table matters**
> A Table auto-expands when rows are added, so your formulas, PivotTables and charts pick up new data without you re-pointing a single range. It also gives you readable formulas: `tblRuns[Capacity Retention]` instead of `$H$2:$H$1001`. This one keystroke is what makes the dashboard refreshable.

> **Watch for**
> **Flash Fill** (`Ctrl+E`) looks magical for splitting codes or labels — type the first result, press it, done. But it produces static text, not formulas. If the source data changes, Flash Fill output does not. Use it for one-off cleanup only.

### Wide data won't pivot — reshape it first

Instruments and survey tools usually export **wide**: one column per timepoint or per condition.

| Run ID | Cycle 100 | Cycle 250 | Cycle 500 |
|---|---|---|---|
| RUN-1001 | 98.1 | 94.6 | 91.4 |

PivotTables cannot use this. "Cycle 250" is a *value* of a condition, not a separate measure, and a pivot can only group by things that live in one column. You need it **long**:

| Run ID | Cycle Point | Retention |
|---|---|---|
| RUN-1001 | 100 | 98.1 |
| RUN-1001 | 250 | 94.6 |
| RUN-1001 | 500 | 91.4 |

One row per observation; the condition becomes a column of its own. If a file arrives wide, reshape it **before** anything else — every later step depends on it. For three or four columns, copy and stack them by hand. For more, `Data → Get & Transform → From Table/Range → Unpivot Columns` does it in two clicks and remembers the steps.

> **`OPTIONAL` — the one-minute version**
> Select the columns you want stacked in Power Query, right-click → `Unpivot Columns`, rename the two new columns, `Close & Load`. Re-running it next month is one click on `Refresh`.

---

## Module 2 — The functions that cover most work

Roughly a dozen, used well, handle the overwhelming majority of day-to-day analysis.

### Build the calculated columns

```excel
=[@[Capacity Retention]]/100*XLOOKUP([@[Cell Model]],tblCells[Cell Model],tblCells[Rated Capacity])
```
Remaining capacity in mAh — a measure the raw file doesn't contain. Inside a Table, typing `[@` lets you pick columns by name, and the formula fills all 1,000 rows the moment you press `Enter`.

```excel
=IF([@[Capacity Retention]]>=XLOOKUP([@[Cell Model]],tblCells[Cell Model],tblCells[Spec Floor]),"Pass","Fail")
```
Pass or fail against that model's own spec floor. Note the threshold is looked up per model rather than typed in — hard-coding `80` here would quietly mis-grade every Model B.

### Conditional aggregation — the workhorses

```excel
=AVERAGEIFS(tblRuns[Capacity Retention], tblRuns[Cell Model], "Model B", tblRuns[Charge Rate], 1)
```
**The one you'll use most.** Mean retention for Model B at 1.0C. Read it as: **average this column, where that column equals this, and that other column equals that.** The average range comes first; the criteria come in pairs after it.

```excel
=COUNTIFS(tblRuns[Cell Model], "Model B", tblRuns[Charge Rate], 1)
```
How many runs that average is based on — its **n**. Always report it next to the average. A mean of three runs and a mean of ninety are not the same claim, and nothing on the face of a chart tells them apart.

```excel
=SUMIFS(tblRuns[Cycles], tblRuns[Rig ID], "RIG-04")
```
`SUMIFS` works identically but adds instead of averaging. Use it only for things that genuinely accumulate — cycles run, units sold, hours logged — never for a measured level like retention or resistance.

> **`OPTIONAL` — criteria on a range of values**
> ```excel
> =AVERAGEIFS(tblRuns[Capacity Retention], tblRuns[Ambient Temp], ">=35", tblRuns[Charge Rate], "<2")
> ```
> Comparison operators go inside the quotes. To compare against a date or a cell, join it with `&`:
> ```excel
> =AVERAGEIFS(tblRuns[Capacity Retention], tblRuns[Test Date], ">="&DATE(2025,7,1), tblRuns[Test Date], "<="&EOMONTH(DATE(2025,7,1),0))
> ```
> `EOMONTH` finds the last day of that month, so you never hand-count 30 or 31.

### Logic and rounding

```excel
=IF([@[Capacity Retention]]>=95,"Excellent",IF([@[Capacity Retention]]>=85,"Good","Degraded"))
=IFS([@[Capacity Retention]]>=95,"Excellent",[@[Capacity Retention]]>=85,"Good",TRUE,"Degraded")
```
Identical results. `IFS` reads left to right and stops at the first `TRUE`, so order your tests from most to least restrictive. The final `TRUE` is the catch-all.

```excel
=ROUND([@[Capacity Retention]],1)
```
Formatting a cell to one decimal only *displays* a rounded number — the underlying value keeps its decimals. `ROUND` changes the value itself. **Round for display, never before calculating**: rounding your raw readings first and then averaging them introduces error that wasn't in the measurement.

> **Absolute references in 15 seconds**
> `A1` moves when you copy it. `$A$1` never moves. `$A1` locks the column, `A$1` locks the row. Press `F4` while the cursor is on a reference to cycle through all four. You'll need `$A1` specifically for the conditional formatting rule in Module 8.

---

## Module 3 — Describe the data, then find what's wrong with it

Before a single chart. This module is where the anomalies come from, and anomalies are usually the most interesting thing in a dataset.

### The six numbers that describe any measure

Build this block once, somewhere on your `Pivots` sheet:

```excel
=COUNT(tblRuns[Capacity Retention])
=AVERAGE(tblRuns[Capacity Retention])
=MEDIAN(tblRuns[Capacity Retention])
=STDEV.S(tblRuns[Capacity Retention])
=MIN(tblRuns[Capacity Retention])
=MAX(tblRuns[Capacity Retention])
```

Read them together, because each one answers a different question:

| | What it tells you | What to notice |
|---|---|---|
| `COUNT` | How many actual readings there are | Lower than your row count? Blanks or text are hiding in the column. |
| `AVERAGE` | The centre, pulled by extremes | |
| `MEDIAN` | The centre, ignoring extremes | **Far from the average? The data is skewed or has outliers.** This one comparison is the fastest anomaly detector there is. |
| `STDEV.S` | Typical distance from the average — the spread | Large relative to the average means your groups may not be genuinely different. |
| `MIN` / `MAX` | The extremes | Check both against what's physically possible before anything else. |

Then repeat it **per group** — a mean for the whole file hides everything interesting:

```excel
=AVERAGEIFS(tblRuns[Capacity Retention], tblRuns[Cell Model], $A2)
=COUNTIFS(tblRuns[Cell Model], $A2)
```

Put the model names down column `A` and drag across. The `$` locks the column so the formula stays pointed at the label as you copy sideways.

### Flag the outliers

An outlier is a reading far from the rest of its group. The standard rule: **more than two standard deviations from the mean**.

```excel
=ABS([@[Capacity Retention]]-AVERAGE(tblRuns[Capacity Retention]))/STDEV.S(tblRuns[Capacity Retention])
```

That's a **z-score** — how many standard deviations from the mean this reading sits, ignoring direction. Then:

```excel
=IF([@[Z Score]]>2,"Review","OK")
```

Roughly 5% of normal data scores above 2, so expect some hits. You are looking for *clusters* of them — all in one rig, one model, one date — not individual rows.

> **`OPTIONAL` — the IQR method**
> More robust when the data is skewed, because it doesn't use the mean at all. Anything outside these two fences is an outlier:
> ```excel
> =QUARTILE.INC(tblRuns[Capacity Retention],1)-1.5*(QUARTILE.INC(tblRuns[Capacity Retention],3)-QUARTILE.INC(tblRuns[Capacity Retention],1))
> =QUARTILE.INC(tblRuns[Capacity Retention],3)+1.5*(QUARTILE.INC(tblRuns[Capacity Retention],3)-QUARTILE.INC(tblRuns[Capacity Retention],1))
> ```
> This is exactly what the whiskers on a box plot mark.

### Three kinds of anomaly, and what each means

Not every anomaly is an error. Sorting them is the analytical work:

1. **Impossible values** — retention above 100%, negative resistance. Always an error. Fix or remove, and say in your summary that you did.
2. **Extreme but possible values** — one run at 61% when the rest sit near 90%. Might be a genuine early failure, which is a finding, not dirt. Never delete these silently.
3. **Systematic offsets** — a whole group reading consistently low. Filter to one rig at a time and compare its mean against the others:

   ```excel
   =AVERAGEIFS(tblRuns[Capacity Retention], tblRuns[Rig ID], $A2)
   ```
   If one rig sits several points below every other across all models and all charge rates, that isn't a bad battery — it's a bad instrument. **This is the highest-value thing you can find in a dataset**, because it invalidates every conclusion drawn from that rig until it's corrected for.

> **Never delete an outlier just because it's inconvenient**
> Removing points until the pattern looks clean is how analyses become wrong. Flag, investigate, then decide — and state what you excluded and why. An analysis that says "we dropped 14 runs from RIG-04 for suspected calibration drift" is far stronger than one that quietly reports a tidier number.

### Does the condition actually relate to the measure?

```excel
=CORREL(tblRuns[Charge Rate], tblRuns[Capacity Retention])
```

Returns between **−1 and +1**. Near `+1`: they rise together. Near `−1`: one rises as the other falls. Near `0`: no straight-line relationship.

Rough reading: below 0.3 is weak, 0.3–0.7 moderate, above 0.7 strong. Negative values are read on the same scale — `−0.8` is a strong relationship.

> **Correlation is not causation, and Excel cannot tell the difference**
> Two things can move together because one causes the other, because something else drives both, or by chance. Here, faster charging and higher temperature both climb together in the test schedule, so a correlation between rate and degradation can't by itself tell you which one is responsible. Say what you found; be careful about why.

---

## Module 4 — XLOOKUP, and why VLOOKUP retires today

Pull each model's chemistry and spec floor from the `Cells` sheet into `Runs`, so every row can be judged against its own standard.

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```
Three arguments required, three optional. Say it out loud as **"find this, in here, give me that."**

### The two you'll definitely use

```excel
=XLOOKUP([@[Cell Model]], tblCells[Cell Model], tblCells[Spec Floor], "Unknown model")
```
**Exact match with a safety net.** The fourth argument replaces the `#N/A` that would otherwise litter your dashboard — and unlike `IFERROR`, it only catches genuine misses, not your own typos elsewhere in the formula. If "Unknown model" appears anywhere, your cleaning in Module 1 missed a spelling.

```excel
=XLOOKUP([@[Cell Model]], tblCells[Cell Model], tblCells[[Chemistry]:[Spec Floor]])
```
**Return three columns at once.** Give the return argument a multi-column range and the result spills sideways into the neighbouring cells. VLOOKUP cannot do this at all.

> **`OPTIONAL` — two more forms**
> **Banding a continuous value.** `-1` means "exact match, or the next smaller item" — the right behaviour for turning a number into a category. Your band table needs a `Floor` column: 0, 80, 85, 95.
> ```excel
> =XLOOKUP([@[Capacity Retention]], tblBands[Floor], tblBands[Band], "", -1)
> ```
> **Searching bottom-up.** The last argument `-1` reverses the search direction, so on a date-sorted table this returns that rig's *most recent* run. Impossible in VLOOKUP.
> ```excel
> =XLOOKUP(A2, tblRuns[Rig ID], tblRuns[Test Date], "No runs", 0, -1)
> ```

### Why the upgrade is worth it

| | VLOOKUP | XLOOKUP |
|---|---|---|
| **Looks left** | No — key must be the leftmost column | Yes, any direction |
| **Column inserted** | Breaks silently, returns wrong data | Unaffected |
| **Default match** | Approximate — a classic source of silent errors | Exact |
| **Not found** | `#N/A`, needs wrapping | Built-in argument |
| **Search from bottom** | No | Yes |

The second row is the dangerous one: VLOOKUP counts columns by position, so inserting a column anywhere in the source shifts every result without raising an error.

> **`OPTIONAL` — on Excel 2019 or older**
> Use this equivalent everywhere the handbook says XLOOKUP. It looks left, and it survives inserted columns — the only thing it lacks is the `if_not_found` argument.
> ```excel
> =IFERROR(INDEX(tblCells[Spec Floor], MATCH([@[Cell Model]], tblCells[Cell Model], 0)), "Unknown model")
> ```

---

## Module 5 — PivotTables: summarise 1,000 rows in four drags

Everything you built with AVERAGEIFS in Modules 2 and 3, rebuilt in seconds — and re-sliceable without touching a formula.

### Your first pivot

1. Click any cell inside `tblRuns`, then `Insert → PivotTable → New Worksheet`. Rename that sheet `Pivots`.
2. Drag `Cell Model` into **Rows**.
3. Drag `Capacity Retention` into **Values**.
4. **Stop and read what it says.** It will read "Sum of Capacity Retention" and show a number in the tens of thousands.

### Fix the aggregation — the most important step in this handbook

**Excel defaults to Sum, and Sum is wrong for almost every measurement.** Adding up 340 capacity readings produces a number with no physical meaning that changes size purely with how many runs a model happened to get. Slice it and the "total" moves for reasons that have nothing to do with the batteries.

Right-click the value → `Value Field Settings` → **Average**. Then:

- Drag `Capacity Retention` into **Values** a second time → set it to **StdDev** → this is the spread.
- Drag `Run ID` into **Values** → set it to **Count** → this is your **n**.

Now you have mean, spread and count side by side, which is the minimum honest summary of a group. A mean alone is a claim without evidence.

| Use | When |
|---|---|
| **Average** | Any measured level — retention, resistance, temperature, score, price |
| **Sum** | Only things that accumulate — cycles, units, revenue, hours |
| **Count / CountA** | How many rows are behind each cell. Always show it somewhere. |
| **StdDev** | How much the readings disagree with each other |
| **Max / Min** | Worst and best case, useful beside the average |

> **Check this before every screenshot**
> "Sum of" in a value header on a measurement column means the dashboard is wrong. It is the single most common serious error in Excel analysis, and it is invisible unless you read the header.

### Bin a continuous condition

`Charge Rate` has four levels here, so it groups naturally. When a condition has many distinct values — temperature, dose, age, price — put it in **Rows**, right-click any number → `Group`, and set **Starting at / Ending at / By**. Excel creates bands like `0.5–1.0`, `1.0–1.5`.

Bands are a judgement call: too few hides the pattern, too many leaves one row per value. Start with five or six and adjust.

Dates work the same way: right-click → `Group → Months and Years`. Only works if Module 1 converted them to real dates.

### Compare everything against the control

This is what turns a table of numbers into a result. Drag `Capacity Retention` into **Values** again, then right-click → `Show Values As` → **% Difference From** → Base field: `Charge Rate`, Base item: `0.5`.

Every cell now reads as its distance from the control condition — `−9.2%` instead of `82.6`. Same data, and now it states a finding directly.

`% of Grand Total` answers a different question — each group's share of the whole — and is the right choice when the measure is a total rather than a level.

### Five moves that make it presentable

- **Rename the ugly header.** Click "Average of Capacity Retention" and type `Mean retention ` — with a trailing space, because Excel refuses a name that matches the source field exactly.
- **Format the numbers once.** Right-click a value → `Number Format → Number, 1 decimal`. Do it here, not by selecting cells — this way it survives a refresh. Match the precision of the instrument: don't report six decimals from a reading taken to one.
- **Top 5 only.** Right-click a row label → `Filter → Top 10` → change 10 to 5.
- **Sort by value**, not alphabetically — right-click → `Sort → Largest to Smallest`. Charts follow the pivot's order.
- **Kill the blank row.** A `(blank)` label means unhandled empty cells upstream. Go back to Module 1.

### Build these four — you need them for the dashboard

| | Pivot | Values |
|---|---|---|
| Pivot 1 | Mean retention by Cell Model | Average, StdDev, Count |
| Pivot 2 | Mean retention by Charge Rate | Average, and % Difference From control |
| Pivot 3 | Mean retention by Rig ID | Average, Count — this is your anomaly check |
| Pivot 4 | Count of runs by Pass/Fail | Count |

> **Two things that trip everyone**
> **Pivots do not update themselves.** Change the source data and the pivot keeps showing the old numbers until you press `Alt+F5`, or `Data → Refresh All` for the lot.
> **GETPIVOTDATA** hijacks your formula when you click a pivot cell. Turn it off: `PivotTable Analyze → Options ▾ → untick Generate GetPivotData`.

---

## Module 6 — Charts that answer a question

Pick the chart from the question you're answering, not from the ribbon gallery.

| The question | The chart | Built from |
|---|---|---|
| Does this number drive that one? | **Scatter, with a trendline** | Raw rows — not a pivot |
| How do groups compare? | Clustered column, sorted | Pivot 1 — by Cell Model |
| Does it change across a condition? | Line or column across the condition | Pivot 2 — by Charge Rate |
| Which group is out of line? | Horizontal bar, sorted | Pivot 3 — by Rig ID |
| How is one measure distributed? | Histogram | Raw rows |
| Two different units together? | Combo with a secondary axis | Retention bars + resistance line |
| Trend inside a table row? | Sparkline | `Insert → Sparklines → Line` |

Click any pivot, then `PivotTable Analyze → PivotChart` — the chart is wired to that pivot and will react to slicers in Module 7.

### The scatter plot, in detail

For "does charge rate affect retention", nothing else will do. A bar chart of averages hides how much the individual runs disagree; a scatter shows every run.

1. Select the two numeric columns — condition first, measure second. **Excel always plots the left column on the x-axis**, so column order decides the chart.
2. `Insert → Scatter (Markers only)`. Never the joined-dots variant; connecting unordered observations draws a line that means nothing.
3. Right-click any point → `Add Trendline` → **Linear**.
4. In the trendline pane, tick **Display R-squared value on chart**.

**Reading R².** It runs 0 to 1 and says how much of the variation in the measure the line accounts for. `0.85` — the relationship explains most of what you see. `0.10` — a line through noise. It is the square of the correlation from Module 3, so a `CORREL` of `−0.8` gives an R² of `0.64`.

A low R² is a real result, not a failure. "Charge rate alone explains only a third of the variation, so something else is driving the rest" is a genuine finding.

> **`OPTIONAL` — show the spread on a bar chart**
> Select the chart → `Chart Design → Add Chart Element → Error Bars → More Options → Custom` → point both fields at your StdDev column from Pivot 1. Two bars whose error bars overlap heavily are not convincingly different, however far apart their tops look.

> **`OPTIONAL` — histogram**
> Select one measure column → `Insert → Statistic Chart → Histogram`. Shows the shape of a single variable: one hump, two humps (two populations mixed together), or a long tail. Right-click the axis to set bin width.

### Clean up every chart

1. **Delete the gridlines.** Click one, press `Delete`. They compete with your data for attention.
2. **Delete the legend** when there's only one series. It's telling you something you already know.
3. **Label the axes with their units.** `Capacity retention (%)`, `Charge rate (C)`. An unlabelled axis on a measurement chart is unreadable to anyone but you.
4. **Don't start a bar chart's axis above zero.** Truncating the axis makes a two-point difference look like a collapse. Line and scatter charts may start elsewhere; bars may not, because the bar's *length* is the comparison.
5. **Sort bar charts** by value, not alphabetically. Sort the pivot and the chart follows.
6. **Write a real title.** "Retention by Charge Rate" is a label. "Retention falls 9 points from 0.5C to 2.0C" is a finding.

> **Don't**
> No 3-D anything — the perspective distorts the very lengths you're asking people to compare. No pie charts beyond three slices. No dual axes unless the two series genuinely use different units; otherwise you can make any two lines cross wherever you like.

---

## Module 7 — Slicers: the part that makes it interactive

This is the single step that turns a page of charts into a dashboard someone else can use without asking you questions.

1. Click any PivotTable, then `PivotTable Analyze → Insert Slicer`. Tick `Cell Model` and `Rig ID`.
2. Add a **Timeline**: `Insert Timeline` → tick `Test Date` → switch the dropdown to **Months**. Timelines only accept real date fields.
3. **Wire each slicer to every pivot.** Right-click the slicer → `Report Connections` → tick all four pivots. Repeat for the second slicer and the timeline.
4. Style it: `Slicer → Columns: 2` so it sits in a tidy block, and drop the height to match your KPI row.

> **Don't skip step 3**
> Without **Report Connections**, each slicer controls only the pivot it was created from. Someone clicks "Model A", one chart changes, the others don't, and the whole dashboard looks broken. Check every box, on every slicer.

**Slicers are also an analysis tool, not just decoration.** Excluding one rig with a single click and watching whether the headline pattern survives is a real test of whether your finding is robust. Do that before you present anything.

Test it: click a model and confirm all charts move together. Then `Clear Filter` (the funnel icon, top-right of the slicer) to reset.

---

## Module 8 — Conditional formatting, used with restraint

Colour should point at the exceptions. If everything is coloured, nothing is.

### Three built-ins worth knowing

- **Data bars** — `Home → Conditional Formatting → Data Bars`. An in-cell bar chart; best on a column of values beside its labels. Tick `Show Bar Only` to hide the numbers and keep it clean.
- **Colour scales** — good for a dense grid like charge rate × temperature, poor for a single column where a bar reads faster.
- **Top/Bottom rules** — `Bottom 10%` on your run list surfaces the worst performers without any sorting.

### The one that matters: a formula rule

Highlight the **entire row** of any run that fell below its model's spec floor. Select your data range first — the selection defines the rule's scope — then `Conditional Formatting → New Rule → Use a formula`:

```excel
=$H2<$L2
```

Write the formula **for the top-left cell of your selection only**; Excel applies the same logic to the rest relatively. The `$` before the column letters is what makes the whole row highlight — without it, each cell checks against itself and you get a diagonal stripe.

Use the same technique to mark your flagged outliers, pointing the rule at the z-score column:

```excel
=$M2>2
```

> **Rules for using colour**
> Pick one hue for "good" and one for "needs attention", and reuse them across the whole workbook. Never rely on colour alone — pair it with an icon or a text label, because roughly 1 in 12 men cannot distinguish your red from your green. Check your rules afterwards in `Conditional Formatting → Manage Rules`: duplicated and overlapping rules are the usual cause of a file that has become mysteriously slow.

---

## Module 9 — Turning numbers into findings

A dashboard that only shows numbers makes the reader do the analysis. Say what you found.

### What a finding looks like

A finding has three parts: **the pattern, the size, and the caveat.**

| Not a finding | A finding |
|---|---|
| "Retention varies by charge rate." | "Retention falls from 94% at 0.5C to 85% at 2.0C — a 9-point drop, consistent across all three models." |
| "Model A performed differently." | "Model A averages 3 points below Model B, but the spread within each model is about 4 points, so the gap isn't convincing on this data." |
| "There were some outliers." | "RIG-04 reads about 8 points low across every model and rate, which looks like calibration drift rather than cell behaviour. Excluding it, the rate effect holds." |

The pattern says what moves. The size makes it checkable. The caveat is what separates analysis from assertion — and it's usually the part that earns the most credit, because it shows you looked for reasons to doubt yourself.

### Write four or five bullets, in this order

1. **The headline** — the strongest relationship in the data, with its number.
2. **The comparison between groups** — and whether the difference is big relative to the spread within them.
3. **The anomaly** — what's wrong or suspicious, what you did about it, and whether it changes the conclusion.
4. **The limit** — what this data cannot tell you. Uneven group sizes, conditions that vary together, a measure recorded only at one timepoint.
5. **`OPTIONAL` — what you'd do next** if you had more data or more time.

### Phrases that give you away

- **"This proves…"** — it doesn't. Data shows, suggests, is consistent with.
- **"X causes Y"** when you only measured that they move together. See Module 3.
- **"A significant difference"** used loosely. Significance is a specific statistical test. Say "a difference of 9 points" instead.
- **Any number with more decimals than the instrument produced.**
- **An average with no n and no spread beside it.**

Put these bullets in a text box directly on the dashboard — `Insert → Text Box`. A dashboard nobody has to interpret is worth several that look impressive.

---

## Exercise — Build the dashboard

On the `Dashboard` sheet, using only what you built in the modules above. Everything you need already exists on the `Pivots` sheet — this is assembly, not invention.

### Required layout

```
┌─────────────────────────────────────────────────────────────┐
│  Battery Endurance Test — 500 cycles                        │
├──────────────┬──────────────┬──────────────┬────────────────┤
│  KPI 1       │  KPI 2       │  KPI 3       │  KPI 4         │
│  Runs (n)    │  Mean        │  Spread      │  Flagged for   │
│              │  retention   │  (SD)        │  review        │
├──────────────┼──────────────┴──────┬───────┴────────────────┤
│  SLICERS     │  Chart A            │  Chart B               │
│  Cell Model  │  Charge rate vs     │  Mean retention        │
│  Rig ID      │  retention          │  by cell model         │
│  + Timeline  │  (scatter + R²)     │  (column, sorted)      │
├──────────────┼─────────────────────┴────────────────────────┤
│  Chart C     │  KEY FINDINGS                                │
│  Mean by rig │  • …                                         │
│  (bar,       │  • …                                         │
│  sorted)     │  • …                                         │
└──────────────┴──────────────────────────────────────────────┘
```

Everything on one screen. If a reader has to scroll, it isn't a dashboard yet.

### The four KPI formulas

```excel
=COUNT(tblRuns[Capacity Retention])
=ROUND(AVERAGE(tblRuns[Capacity Retention]),1)
=ROUND(STDEV.S(tblRuns[Capacity Retention]),1)
=COUNTIFS(tblRuns[Outlier Flag],"Review")&" of "&COUNTA(tblRuns[Run ID])
```

Note the first KPI is the count of *readings*, not rows — it tells you immediately if blanks are eating your data. These four are static totals and won't respond to slicers, which is fine for a headline; making KPIs follow the slicer needs a pivot or a CUBE function, beyond what this session covers.


## Shortcuts

**Moving and selecting**
`Ctrl+↓` jump to the last row · `Ctrl+Shift+↓` select to it · `Ctrl+A` select the whole table · `Ctrl+Home` back to `A1` · `Ctrl+PgDn` next sheet

**Doing the work**
`Ctrl+T` make a Table · `Ctrl+Shift+L` toggle filters · `Alt+=` AutoSum · `Ctrl+E` Flash Fill · `Alt+F5` refresh a pivot · `F4` cycle `$` anchors

**Formatting**
`Ctrl+1` Format Cells · `Ctrl+Shift+1` number with 2 decimals · `Ctrl+Shift+5` percent · `Ctrl+Shift+V` paste special · `Alt+Enter` line break in a cell

## When it goes wrong

| Symptom | Cause and fix |
|---|---|
| Pivot header says **"Sum of"** | The default. Right-click → `Value Field Settings → Average`. Check this on every value field before you trust any number on the dashboard. |
| "Sum of" says **Count** instead | Text hiding in your number column — often a stray space, or a unit typed into the cell. Filter the column and look at the bottom of the list. |
| An average looks impossibly low | Blanks were filled with `0`. `AVERAGE` skips blanks but counts zeroes. Undo, leave them blank. |
| `(blank)` appears as a row label | Unhandled empty cells in a grouping column. Back to Module 1. |
| Dates won't group by month | They're text. Left-aligned by default is the giveaway. Go back to `Text to Columns` in Module 1, then refresh the pivot. |
| Scatter plots the wrong variable on x | Excel always uses the left column for x. Reorder the two columns and re-insert. |
| The slicer only moves one chart | Report Connections. Right-click each slicer and tick every pivot. It has to be done per slicer, not once for the sheet. |
| `#REF!` appeared everywhere | You deleted rows or columns a formula pointed at. `Ctrl+Z` immediately — `#REF!` destroys the original reference, and undo is the only way back. |
| `#DIV/0!` from an average | No rows match those criteria. Check your spelling against the cleaned values, not the raw ones. |
| The file crawls | Usually overlapping conditional formatting rules, or formulas pointing at whole columns (`A:A`) instead of a Table. Check `Manage Rules` first. |

---


