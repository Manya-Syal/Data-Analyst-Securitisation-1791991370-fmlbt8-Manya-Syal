# Project 1A – Securitisation Risk Analytics: Build Guide

Built on your four ZeTheta files and the brief (483553A). **Deadline: 14/10/2026 (extended 29/10/2026).** Everything is private and confidential under the portal rules – keep all repositories private.

## 1. What is in the package

| File | Use |
|---|---|
| `securitisation_pipeline.py` | Reads the 4 CSVs from `input_data/`, cleans them, builds the star schema, PD term structure, scenarios, waterfall, validation, data-quality log |
| `powerbi_data.zip` | 25 CSVs produced by the pipeline – load these into Power BI |
| `securitisation_measures.dax.md` | **222 DAX measures** (brief asks 80+), grouped by Day 3–12, plus parameter-table DAX and the relationship map |
| `483553A_ManyaSyal_DAX_Measure_Dictionary.xlsx` | Measure dictionary (the brief's required Excel deliverable) with an *Expected value* column to test against |
| `483553A_ManyaSyal_ECL_Waterfall_Validation.xlsx` | ECL cross-check and a formula rebuild of the 36-month waterfall (5,683 live formulas, zero errors) |
| `DAXStudio_Benchmark.dax` | One query per measure for the <100 ms performance test |
| `build_measures.py`, `build_excel_validation.py` | Regenerate the DAX file and both workbooks if you change anything |

**Not included (still to do):** the `.pbix` itself (I cannot create one), the 40+ page report, the 10-slide deck and the GitHub repository. See section 9.

## 2. What your data says (headline numbers)

- Pool: **₹317.8M** outstanding on 500 loans (476 active, 24 with zero balance), WAC 10.95%, WAM 45 months, pool factor 58%.
- Delinquency: **30+ DPD 14.1%, 60+ DPD 10.6%, 90+ DPD 5.6%** of balance.
- Tape ECL: **₹11.66M = 3.58% of EAD** (Stage 1 ₹2.8M, Stage 2 ₹1.6M, Stage 3 ₹7.2M). It reconciles exactly to PD × LGD × EAD for every loan.
- Calibration from `dynamic_loss_monthly`: annual default rate **3.58%**, net loss severity **41.5%**, CPR **9.2%**.
- **Challenger ECL (roll-rate PD) = ₹29.5M, 2.5× the tape.** The tape's Stage 1 PD is about 2%, but your DPD history implies about 11% for loans starting "Current". Caveats: only 11 monthly transitions and 31 entries into 90+. Present this as a model-risk finding, not a corrected provision.
- Concentration: top 5 states hold 37.7% of balance.
- Scenario write-downs over 36 months (assumed structure): Base none; Moderate Stress 34% of equity; Severe Stress wipes equity and 83% of Class C; Crisis also takes 82% of Class B; Class A is untouched in all six.
- The 60+ DPD ratio (10.6%) is close to the assumed 12% delinquency trigger – thin headroom.

## 3. Data quality log (you must disclose these)

| ID | Severity | Finding | Treatment |
|---|---|---|---|
| DQ05 | High | 12 DPD snapshots per loan but only 11 distinct dates; **30-May-2024 appears twice** | All 12 used as sequential snapshots for roll rates; second May row flagged `IncludeInTrend = 0` for date-based visuals |
| DQ06 | High | 588 snapshot rows (102 loans) are dated **before the loan was originated** | Flagged `PreOrigination = 1`, excluded from transitions and DPD measures |
| DQ08 | High | Dynamic loss file has the same duplicated 30-May row | Same treatment |
| DQ10 | High | Tape balance (₹317.8M) ≠ dynamic file closing balance (₹470.1M): **₹152M break** | Reported; measure `Reconciliation Break (Tape vs Dynamic)` |
| DQ01 | Medium | 24 loans have zero balance but still carry DPD status, stage, PD, LGD | `Active Loans` measure excludes them |
| DQ03 | Medium | Tape staging (S2 = 30–89 DPD, S3 = 90+) differs from the brief's rule (S2 = 60–90, S3 = 120+) | `ECL (Alt Staging)` with adjustable thresholds shows the impact |
| DQ09 | Medium | Each month's opening balance ≠ prior month's closing balance | Used for rates only |
| DQ02 / DQ04 / DQ07 | Low | EAD > balance; only 2 loans have realised losses; prior-bucket field consistent | Documented |
| DQ11 | Info | Static pool original balance (₹546,073,700) **reconciles exactly** to the tape | – |

Also ask ZeTheta which month the duplicated 30-May snapshot really is (hr@zetheta.com or your supervisor).

## 4. Set up the Power BI model

1. **Get Data → Text/CSV**: load every CSV from the zip. Set `Date`, `MonthEnd`, `SnapshotDate`, `ReportingDate` etc. to *Date*; money columns to *Fixed decimal*.
2. **Relationships** – exactly as listed in section 1 of `securitisation_measures.dax.md`. `DimLoan` is the hub; DimBorrower, DimVehicle, DimGeography, DimVintage connect to it. The date→DimLoan[OriginationDate] link must be **inactive**. Do **not** relate DimBucket or DimDate to FactLoanTape (creates ambiguous paths).
3. **Mark `DimDate` as the date table.** Sort: `DelinquencyStatus` and `DPD_Bucket` by `BucketOrder`; `SnapshotLabel` by `SnapshotSeq`; scenario by `scenario_id`; tranche by `seniority`.
4. **Parameter tables:** paste the 9 `GENERATESERIES` tables from section 2 of the DAX file (Modeling → New table). Add a slicer for each on the stress page.
5. **Measures:** create a blank `_Measures` table, then add the measures. 222 is a lot to paste one by one – use *DAX query view* (`DEFINE MEASURE`) or Tabular Editor (it is in your brief's tool list). I can generate a Tabular Editor script if you want.
6. Hide key and helper columns; set formats from the dictionary's *Format* column.
7. **Test:** put a card for each headline measure and compare with the dictionary's *Expected value* column (paste your value in column J; column K shows MATCH / CHECK).

## 5. Dashboard pages (brief requires 8+; this gives 10)

1. **Executive Summary** – cards: Total Outstanding Balance, Pool Factor, Loan Count, WAC, WAM, WA CIBIL, Delinquent 30+/60+/90+ %, Total ECL (Tape), ECL Coverage %; donut by stage; slicers: servicer, region, vintage, scenario.
2. **Pool Stratification** – tables/bars by LTV band, CIBIL band, state, vehicle make, employment; `Pool Share %`, `Top 3 States`, `HHI State`, `Top 10 Loan Concentration`.
3. **Delinquency & Roll Rates** – DPD distribution by `SnapshotLabel` (stacked column, `Bucket Balance %`); roll-rate **matrix** (rows `PriorBucketSeq`, columns `DimBucket[DPD_Bucket]`, value `Roll Rate %`); cure/roll-forward cards; ageing bars; trend alert card.
4. **Vintage / Static Pool** – line chart: axis `MonthsOnBook`, legend `VintageID`, value `CNL Rate @ MOB`; a second for cumulative default rate; loss-timing curve; table of `CNL Rate (Latest MOB)` and `Projected Ultimate CNL (MOB 36)` by vintage.
5. **Loss & Prepayment Trend** – monthly net loss, default rate, CPR, collection efficiency; MTD/QTD/FYTD cards; rolling 3M/12M default rate; Z-score alert.
6. **IFRS 9 ECL** – ECL and EAD by stage, coverage, stage migration table, tape vs recomputed vs challenger bars, PD term-structure line chart (`PD Term Structure` by `HorizonMonths`, legend bucket), alt-staging card with the two threshold slicers.
7. **Waterfall & Triggers** – single-select scenario slicer; pool balance and cumulative loss lines; tranche balance stacked area; matrix of interest due/paid, principal, write-down, `Tranche Status`; trigger cards with status colours; `CNL Trigger Breach Month`.
8. **Stress Testing** – four what-if sliders → `Stress ECL (What-If)`, uplift; scenario bar chart of `Scenario ECL` with `Probability-Weighted Scenario ECL`; heat-map scenario × tranche (`Writedown % of Tranche`).
9. **Securitisation Risk Arena** (brief Part B) – tile slicer of the six scenarios; `Arena Score`, `Arena Rating`, `Arena Message`; gauge; tranche shields. Add bookmarks for "Round 1–4" (Base → Moderate → Severe → Crisis). Add a case-study panel on the 2008 US sub-prime crisis and the 2018 IL&FS episode (Part C of the brief) – quote only facts you verify from public sources.
10. **Investor / Servicer Report & Data Tape** – the six report sections (pool summary, delinquency, loss & recovery, prepayment, waterfall, triggers) and a rating-agency tape table; `Data Tape Field Completeness %`; data-quality log table. Drill-through page: loan detail.

Use one palette (navy `#0B2545`, teal `#13A89E`, amber `#F4A261`, red `#C0392B`), consistent titles, tooltips and a navigation bar.

## 6. Row-level security (brief: 5+ roles)

Create the five roles under *Modeling → Manage roles*. Roles are mapped to users through `DimSecurity` (replace the sample emails). Apply this one filter to **DimLoan** in every role:

```dax
VAR u  = USERPRINCIPALNAME ()
VAR sv = LOOKUPVALUE ( DimSecurity[ServicerID], DimSecurity[UserEmail], u, "NONE" )
VAR rg = LOOKUPVALUE ( DimSecurity[Region],     DimSecurity[UserEmail], u, "NONE" )
RETURN ( sv = "ALL" || DimLoan[ServicerID] = sv )
    && ( rg = "ALL" || RELATED ( DimGeography[Region] ) = rg )
```

| Role | Sees | Extra control |
|---|---|---|
| PortfolioManager | all loans | all pages |
| RiskAnalyst | all loans | all pages |
| SeniorManagement | all loans | summary pages only (separate app audience); hide `DimBorrower[AnnualIncome_INR]` with Tabular Editor column security |
| Auditor | one region (sample: North) | read-only app audience; hide borrower income |
| Servicer | own servicer's loans (sample: SVC001) | servicer-report pages only |

Test with *View as*. Expected `Loans Visible to User`: PortfolioManager 500, RiskAnalyst 500, SeniorManagement 500, **Auditor (North) 103, Servicer (SVC001) 156**. Note: RLS cannot hide columns – that needs object-level security.

## 7. Validation and performance

- **Excel workbook:** `ECL_Summary` shows PASS/FAIL checks. On `Waterfall`, change `Assumptions!B10` (1–6) – the Excel rebuild matches the Python engine to ₹0.005 in all six scenarios, and cash conservation holds.
- **Accuracy target:** the brief wants ECL to match the Excel cross-check within 0.01%. Tape ECL vs recomputed differs by ₹0.05 on ₹11.66M (0.0004%).
- **Performance:** open DAX Studio, connect to the open report, switch on Server Timings, run `DAXStudio_Benchmark.dax`, and record milliseconds in the dictionary (column L). I could not run Power BI here, so **no measure has been executed against a live model** – the first run may reveal syntax or context errors. Send me any error text and I will fix it.
- **Hard-coded values:** the brief penalises hard-coded values in measures. Thresholds that are part of a definition (30/60/90 DPD, 12-month horizon) are written in the formula; every tunable assumption (triggers, stress levers, Arena weights, pool LGD, CPR) is a parameter table or lives in `DimParameters`. Keep it that way if you edit anything.

## 8. Assumptions you must state in the report

Tranche sizes (80/12/5/3%), coupons, 0.5% servicing fee, the trigger levels (12% / 2% / 1.05), scenario multipliers and weights, the 10% PD elasticity per 100 bps of rate shock, and the Arena weights are **mine**, because the brief gives none. Change them in `securitisation_pipeline.py` / `DimParameters` if ZeTheta gives guidance.

## 9. Self-check against the 1,000-point rubric and what remains

| Category (points) | Status |
|---|---|
| DAX accuracy & complexity (300) | 222 measures with iterators, virtual tables, context transition – **needs your live test** |
| Query performance (150) | Benchmark file ready – **not yet measured** |
| Waterfall & tranche allocation (200) | Done in DAX + Python + Excel; triggers included |
| IFRS 9 ECL (150) | Done; ties to Excel; challenger adds depth |
| Investor reporting (100) | Measures ready; **pages to build** |
| Dashboard quality (100) | Specified; **.pbix to build** |

**Remaining deliverables:** (a) the `.pbix`; (b) main report, 40+ pages – I can draft it from the findings above; (c) 10-slide deck; (d) the GitHub repository. Suggested plan to hit 14 Oct: 6–7 Oct load model and paste measures; 8–10 Oct dashboards and RLS; 11 Oct validation and DAX Studio timings; 12–13 Oct report and deck; 14 Oct repo transfer. The extended date of 29 Oct is your buffer.

## 10. Submission checklist (from the portal rules)

- Create a **private** GitHub repository from day one; add only authorised collaborators; enable branch protection; **disable forking**; no public wikis or discussions.
- **Transfer ownership to `@ZethetaIntern` between day 15 and day 30** – this counts as the submission. Also upload files through the portal's *Submit* box.
- Use clear file names in the form `ProjectCode_Name_Deliverable`.
- Do **not** post code, concepts or results on social media, portfolios or public platforms. Only the project title and task may be mentioned.
- The checklist in Part F of the brief contains items that look carried over from another project; ask your supervisor to confirm which apply.
- The rules reward originality: understand every measure, run the validation, and write the report in your own words.
