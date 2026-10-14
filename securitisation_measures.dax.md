# Securitisation Risk Analytics - DAX Measure Library

**222 measures** across Days 3-12 of the brief. Every measure is formula-driven (no hard-coded results). Generated from `build_measures.py`; the same list feeds the Excel measure dictionary and the DAX Studio benchmark file.

## 1. Model and relationships

Load the CSVs from `powerbi_data/`. Create a blank table `_Measures` and paste measures into it.

| From (one side) | To (many side) | Notes |
|---|---|---|
| DimLoan[LoanID] | FactLoanTape[LoanID] | one-to-many, single direction |
| DimLoan[LoanID] | FactMonthlyPerformance[LoanID] | one-to-many, single |
| DimBorrower[BorrowerID] | DimLoan[BorrowerID] | one-to-many, single (borrower IDs are unique) |
| DimVehicle[LoanID] | DimLoan[LoanID] | one-to-many, single |
| DimGeography[State] | DimLoan[State] | one-to-many, single |
| DimVintage[VintageID] | DimLoan[VintageID] | one-to-many, single |
| DimVintage[VintageID] | FactStaticPool[VintageID] | one-to-many, single |
| DimStage[IFRS9_Stage] | FactLoanTape[IFRS9_Stage] | one-to-many, single |
| DimBucket[DPD_Bucket] | FactMonthlyPerformance[DPD_Bucket] | one-to-many, single |
| DimDate[Date] | FactMonthlyPerformance[MonthEnd] | one-to-many, single |
| DimDate[Date] | FactDynamicLoss[MonthEnd] | one-to-many, single |
| DimDate[Date] | DimLoan[OriginationDate] | **inactive** role-playing date (use USERELATIONSHIP) |
| DimScenario[scenario_id] | FactPool[scenario_id], FactWaterfall[scenario_id] | one-to-many, single |
| DimPeriod[period] | FactPool[period], FactWaterfall[period] | one-to-many, single |
| DimTranche[tranche] | FactWaterfall[tranche] | one-to-many, single |

**Disconnected tables (no relationships):** FactRollRate, FactPDTermStructure, DimParameters, DimSecurity, DataFieldMapping, DataQualityLog, ExpectedMeasureValues.
**Do not** relate DimBucket to FactLoanTape or DimDate to FactLoanTape: together with DimLoan they would create an ambiguous path. Slice the tape by `DelinquencyStatus` (sorted by `BucketOrder`).

Mark DimDate as the date table (column Date). Sort DelinquencyStatus / DPD_Bucket by BucketOrder, SnapshotLabel by SnapshotSeq, scenario by scenario_id, tranche by seniority.

## 2. Parameter tables (Modeling > New table)

```dax
ParamDefaultMult = GENERATESERIES ( 0.5, 5, 0.1 )    // Default-rate multiplier (x)
ParamRecoveryHaircut = GENERATESERIES ( 0, 0.5, 0.05 )    // Recovery haircut = additive LGD (pts)
ParamCPRChange = GENERATESERIES ( -0.1, 0.1, 0.01 )    // Change in annual CPR (pts)
ParamRateShock = GENERATESERIES ( 0, 500, 25 )    // Interest-rate shock (bps)
ParamStage2DPD = GENERATESERIES ( 0, 120, 5 )    // Stage 2 DPD threshold (tape = 30, brief = 60)
ParamStage3DPD = GENERATESERIES ( 60, 180, 5 )    // Stage 3 DPD threshold (tape = 90, brief = 120)
ParamDelinqTrigger = GENERATESERIES ( 0.05, 0.25, 0.005 )    // Delinquency trigger level
ParamCNLTrigger = GENERATESERIES ( 0.005, 0.1, 0.0025 )    // Cumulative net loss trigger level
ParamOCTrigger = GENERATESERIES ( 1, 1.3, 0.01 )    // OC ratio trigger (minimum)
```

Each table is a single column named `Value`; attach a slicer to it. The matching `... Value` measures below read the selection with SELECTEDVALUE and a default.

## 3. Day 3 - Core pool, delinquency and weighted-average measures

### D03-01 - Total Outstanding Balance
*Pool. Current principal outstanding on the pool.*  Format: `#,##0`

```dax
Total Outstanding Balance =
SUM ( FactLoanTape[CurrentBalance] )
```

### D03-02 - Original Balance
*Pool. Original principal advanced at origination.*  Format: `#,##0`

```dax
Original Balance =
SUM ( DimLoan[OriginalLoanAmount] )
```

### D03-03 - Pool Factor
*Pool. Share of original balance still outstanding.*  Format: `0.0%`

```dax
Pool Factor =
DIVIDE ( [Total Outstanding Balance], [Original Balance] )
```

### D03-04 - Total EAD
*Pool. Exposure at default used for ECL.*  Format: `#,##0`

```dax
Total EAD =
SUM ( FactLoanTape[EAD] )
```

### D03-05 - Loan Count
*Counts. All loans on the tape.*  Format: `#,##0`

```dax
Loan Count =
COUNTROWS ( FactLoanTape )
```

### D03-06 - Active Loans
*Counts. Loans with a positive balance (24 loans carry zero balance).*  Format: `#,##0`

```dax
Active Loans =
CALCULATE ( COUNTROWS ( FactLoanTape ), FactLoanTape[CurrentBalance] > 0 )
```

### D03-07 - Defaulted Loans
*Counts. Loans flagged Default or Repossessed.*  Format: `#,##0`

```dax
Defaulted Loans =
CALCULATE ( COUNTROWS ( FactLoanTape ), FactLoanTape[IsDefaulted] = 1 )
```

### D03-08 - Prepaid Loans
*Counts. Loans with a voluntary prepayment.*  Format: `#,##0`

```dax
Prepaid Loans =
CALCULATE ( COUNTROWS ( FactLoanTape ), FactLoanTape[PrepaymentAmount] > 0 )
```

### D03-09 - Defaulted Balance
*Balances. Outstanding balance of defaulted loans.*  Format: `#,##0`

```dax
Defaulted Balance =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[IsDefaulted] = 1 )
```

### D03-10 - Balance Current
*Delinquency. Balance with no arrears.*  Format: `#,##0`

```dax
Balance Current =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyStatus] = "Current" )
```

### D03-11 - Balance 1-29 DPD
*Delinquency. Balance 1-29 days past due.*  Format: `#,##0`

```dax
Balance 1-29 DPD =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyStatus] = "1-29 DPD" )
```

### D03-12 - Balance 30 DPD
*Delinquency. Balance 30-59 days past due.*  Format: `#,##0`

```dax
Balance 30 DPD =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyStatus] = "30-59 DPD" )
```

### D03-13 - Balance 60 DPD
*Delinquency. Balance 60-89 days past due.*  Format: `#,##0`

```dax
Balance 60 DPD =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyStatus] = "60-89 DPD" )
```

### D03-14 - Balance 90 DPD
*Delinquency. Balance 90-119 days past due.*  Format: `#,##0`

```dax
Balance 90 DPD =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyStatus] = "90-119 DPD" )
```

### D03-15 - Balance 120+ DPD
*Delinquency. Balance 120+ DPD plus Default/Repossessed.*  Format: `#,##0`

```dax
Balance 120+ DPD =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyStatus] IN { "120+ DPD", "Default", "Repossessed" } )
```

### D03-16 - Delinquent 30+ Balance
*Delinquency. Balance 30 or more days past due.*  Format: `#,##0`

```dax
Delinquent 30+ Balance =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyDays] >= 30 )
```

### D03-17 - Delinquent 30+ %
*Delinquency. 30+ DPD as % of pool balance (SMA-1 and worse).*  Format: `0.0%`

```dax
Delinquent 30+ % =
DIVIDE ( [Delinquent 30+ Balance], [Total Outstanding Balance] )
```

### D03-18 - Delinquent 60+ Balance
*Delinquency. Balance 60 or more days past due.*  Format: `#,##0`

```dax
Delinquent 60+ Balance =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyDays] >= 60 )
```

### D03-19 - Delinquent 60+ %
*Delinquency. 60+ DPD as % of pool; drives the delinquency trigger.*  Format: `0.0%`

```dax
Delinquent 60+ % =
DIVIDE ( [Delinquent 60+ Balance], [Total Outstanding Balance] )
```

### D03-20 - Delinquent 90+ Balance
*Delinquency. Balance 90 or more days past due (NPA).*  Format: `#,##0`

```dax
Delinquent 90+ Balance =
CALCULATE ( [Total Outstanding Balance], FactLoanTape[DelinquencyDays] >= 90 )
```

### D03-21 - Delinquent 90+ %
*Delinquency. NPA ratio.*  Format: `0.0%`

```dax
Delinquent 90+ % =
DIVIDE ( [Delinquent 90+ Balance], [Total Outstanding Balance] )
```

### D03-22 - WAC
*Weighted avg. Weighted-average coupon (rate in tape is in % points).*  Format: `0.00%`

```dax
WAC =
DIVIDE ( SUMX ( FactLoanTape, RELATED ( DimLoan[InterestRate] ) * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] ) / 100
```

### D03-23 - WAM
*Weighted avg. Weighted-average remaining maturity (months).*  Format: `0.0`

```dax
WAM =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[RemainingTerm] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-24 - WALA
*Weighted avg. Weighted-average loan age (months on book).*  Format: `0.0`

```dax
WALA =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[MonthsOnBook] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-25 - WALT
*Weighted avg. Weighted-average original loan term (months).*  Format: `0.0`

```dax
WALT =
DIVIDE ( SUMX ( FactLoanTape, RELATED ( DimLoan[OriginalTerm] ) * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-26 - WA LTV Origination
*Weighted avg. Balance-weighted LTV at origination.*  Format: `0.0%`

```dax
WA LTV Origination =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[LTV_AtOrigination] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-27 - WA LTV Current
*Weighted avg. Balance-weighted current LTV.*  Format: `0.0%`

```dax
WA LTV Current =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[LTV_Current] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-28 - WA CIBIL
*Weighted avg. Balance-weighted current CIBIL score.*  Format: `0`

```dax
WA CIBIL =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[CIBIL_Score_Current] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-29 - WA CIBIL Origination
*Weighted avg. Balance-weighted CIBIL at origination.*  Format: `0`

```dax
WA CIBIL Origination =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[CIBIL_Score_Origination] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-30 - WA DTI
*Weighted avg. Balance-weighted debt-to-income.*  Format: `0.0%`

```dax
WA DTI =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[DTI_Ratio] * FactLoanTape[CurrentBalance] ), [Total Outstanding Balance] )
```

### D03-31 - Average Loan Balance
*Pool. Mean balance of active loans.*  Format: `#,##0`

```dax
Average Loan Balance =
DIVIDE ( [Total Outstanding Balance], [Active Loans] )
```

### D03-32 - Largest Loan Balance
*Pool. Single largest exposure.*  Format: `#,##0`

```dax
Largest Loan Balance =
MAX ( FactLoanTape[CurrentBalance] )
```

### D03-33 - Insured Loan %
*Pool. Balance share with any vehicle insurance.*  Format: `0.0%`

```dax
Insured Loan % =
DIVIDE ( CALCULATE ( [Total Outstanding Balance], DimLoan[HasInsurance] = 1 ), [Total Outstanding Balance] )
```

### D03-34 - New Vehicle %
*Pool. Balance share financing new vehicles.*  Format: `0.0%`

```dax
New Vehicle % =
DIVIDE ( CALCULATE ( [Total Outstanding Balance], DimVehicle[IsNewVehicle] = 1 ), [Total Outstanding Balance] )
```

### D03-35 - Modified Loan %
*Pool. Share of loans restructured (deferral, extension, rate cut).*  Format: `0.0%`

```dax
Modified Loan % =
DIVIDE ( CALCULATE ( COUNTROWS ( FactLoanTape ), DimLoan[IsModified] = 1 ), [Loan Count] )
```

### D03-36 - Total Prepayments
*Loss. Voluntary prepayments recorded on the tape.*  Format: `#,##0`

```dax
Total Prepayments =
SUM ( FactLoanTape[PrepaymentAmount] )
```

### D03-37 - Total Loss Amount
*Loss. Gross realised loss on the tape.*  Format: `#,##0`

```dax
Total Loss Amount =
SUM ( FactLoanTape[LossAmount] )
```

### D03-38 - Total Recoveries
*Loss. Recoveries on the tape.*  Format: `#,##0`

```dax
Total Recoveries =
SUM ( FactLoanTape[RecoveryAmount] )
```

### D03-39 - Total Net Loss (Tape)
*Loss. Net realised loss on the tape.*  Format: `#,##0`

```dax
Total Net Loss (Tape) =
SUM ( FactLoanTape[NetLoss] )
```

### D03-40 - Payment Collection Rate
*Collections. Last payment received vs scheduled payment due.*  Format: `0.0%`

```dax
Payment Collection Rate =
DIVIDE ( SUM ( FactLoanTape[LastPaymentAmount] ), SUM ( FactLoanTape[ScheduledPaymentDue] ) )
```

## 4. Day 4 - Time intelligence and trend alerts

### D04-01 - Net Loss (Monthly)
*Base. Monthly net loss (second 30-May row excluded - see Data Quality DQ08).*  Format: `#,##0`

```dax
Net Loss (Monthly) =
CALCULATE ( SUM ( FactDynamicLoss[NetLoss_ThisMonth] ), FactDynamicLoss[IncludeInTrend] = 1 )
```

### D04-02 - Gross Loss (Monthly)
*Base. Monthly gross loss.*  Format: `#,##0`

```dax
Gross Loss (Monthly) =
CALCULATE ( SUM ( FactDynamicLoss[GrossLoss_ThisMonth] ), FactDynamicLoss[IncludeInTrend] = 1 )
```

### D04-03 - Recoveries (Monthly)
*Base. Monthly recoveries.*  Format: `#,##0`

```dax
Recoveries (Monthly) =
CALCULATE ( SUM ( FactDynamicLoss[Recoveries_ThisMonth] ), FactDynamicLoss[IncludeInTrend] = 1 )
```

### D04-04 - New Defaults Balance (Monthly)
*Base. Balance newly defaulted in the month.*  Format: `#,##0`

```dax
New Defaults Balance (Monthly) =
CALCULATE ( SUM ( FactDynamicLoss[NewDefaults_Balance] ), FactDynamicLoss[IncludeInTrend] = 1 )
```

### D04-05 - BOP Balance (Monthly)
*Base. Beginning-of-period pool balance.*  Format: `#,##0`

```dax
BOP Balance (Monthly) =
CALCULATE ( SUM ( FactDynamicLoss[BOP_Balance] ), FactDynamicLoss[IncludeInTrend] = 1 )
```

### D04-06 - Monthly Default Rate
*Base. New defaults / opening balance.*  Format: `0.000%`

```dax
Monthly Default Rate =
DIVIDE ( [New Defaults Balance (Monthly)], [BOP Balance (Monthly)] )
```

### D04-07 - Monthly Net Loss Rate
*Base. Net loss / opening balance.*  Format: `0.000%`

```dax
Monthly Net Loss Rate =
DIVIDE ( [Net Loss (Monthly)], [BOP Balance (Monthly)] )
```

### D04-08 - Net Loss MTD
*To-date. Month-to-date net loss.*  Format: `#,##0`

```dax
Net Loss MTD =
TOTALMTD ( [Net Loss (Monthly)], DimDate[Date] )
```

### D04-09 - Net Loss QTD
*To-date. Quarter-to-date net loss.*  Format: `#,##0`

```dax
Net Loss QTD =
TOTALQTD ( [Net Loss (Monthly)], DimDate[Date] )
```

### D04-10 - Net Loss YTD
*To-date. Calendar year-to-date net loss.*  Format: `#,##0`

```dax
Net Loss YTD =
TOTALYTD ( [Net Loss (Monthly)], DimDate[Date] )
```

### D04-11 - Net Loss FYTD
*To-date. Indian fiscal year-to-date (April-March) net loss; built on DimDate[FiscalYear] so it does not depend on locale date strings.*  Format: `#,##0`

```dax
Net Loss FYTD =
VAR fy = MAX ( DimDate[FiscalYear] )
VAR d  = MAX ( DimDate[Date] )
RETURN CALCULATE ( [Net Loss (Monthly)], FILTER ( ALL ( DimDate ), DimDate[FiscalYear] = fy && DimDate[Date] <= d ) )
```

### D04-12 - Net Loss Prior Month
*Comparison. Net loss one month earlier.*  Format: `#,##0`

```dax
Net Loss Prior Month =
CALCULATE ( [Net Loss (Monthly)], DATEADD ( DimDate[Date], -1, MONTH ) )
```

### D04-13 - Net Loss MoM %
*Comparison. Month-on-month change in net loss.*  Format: `0.0%`

```dax
Net Loss MoM % =
DIVIDE ( [Net Loss (Monthly)] - [Net Loss Prior Month], [Net Loss Prior Month] )
```

### D04-14 - Default Rate Prior Quarter
*Comparison. Monthly default rate one quarter earlier.*  Format: `0.000%`

```dax
Default Rate Prior Quarter =
CALCULATE ( [Monthly Default Rate], DATEADD ( DimDate[Date], -1, QUARTER ) )
```

### D04-15 - Default Rate QoQ Change
*Comparison. Quarter-on-quarter change (pts).*  Format: `0.000%`

```dax
Default Rate QoQ Change =
[Monthly Default Rate] - [Default Rate Prior Quarter]
```

### D04-16 - Default Rate vs Prior Year
*Comparison. Prior-year comparator (blank: file covers < 12 months).*  Format: `0.000%`

```dax
Default Rate vs Prior Year =
CALCULATE ( [Monthly Default Rate], SAMEPERIODLASTYEAR ( DimDate[Date] ) )
```

### D04-17 - Rolling 3M Default Rate
*Rolling. Balance-weighted default rate over the trailing 3 months.*  Format: `0.000%`

```dax
Rolling 3M Default Rate =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
VAR Win = DATESINPERIOD ( DimDate[Date], LastD, -3, MONTH )
RETURN DIVIDE ( CALCULATE ( [New Defaults Balance (Monthly)], Win ), CALCULATE ( [BOP Balance (Monthly)], Win ) )
```

### D04-18 - Rolling 12M Default Rate
*Rolling. Balance-weighted default rate over the trailing 12 months.*  Format: `0.000%`

```dax
Rolling 12M Default Rate =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
VAR Win = DATESINPERIOD ( DimDate[Date], LastD, -12, MONTH )
RETURN DIVIDE ( CALCULATE ( [New Defaults Balance (Monthly)], Win ), CALCULATE ( [BOP Balance (Monthly)], Win ) )
```

### D04-19 - Rolling 3M Net Loss Rate
*Rolling. Trailing 3-month net loss rate.*  Format: `0.000%`

```dax
Rolling 3M Net Loss Rate =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
VAR Win = DATESINPERIOD ( DimDate[Date], LastD, -3, MONTH )
RETURN DIVIDE ( CALCULATE ( [Net Loss (Monthly)], Win ), CALCULATE ( [BOP Balance (Monthly)], Win ) )
```

### D04-20 - Default Rate 12M Average
*Alerts. Mean monthly default rate over 12 months.*  Format: `0.000%`

```dax
Default Rate 12M Average =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
VAR Months12 =
    CALCULATETABLE ( VALUES ( DimDate[Date] ), DATESINPERIOD ( DimDate[Date], LastD, -12, MONTH ), DimDate[IsMonthEnd] = 1 )
RETURN AVERAGEX ( Months12, [Monthly Default Rate] )
```

### D04-21 - Default Rate 12M StdDev
*Alerts. Standard deviation of monthly default rate.*  Format: `0.000%`

```dax
Default Rate 12M StdDev =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
VAR Months12 =
    CALCULATETABLE ( VALUES ( DimDate[Date] ), DATESINPERIOD ( DimDate[Date], LastD, -12, MONTH ), DimDate[IsMonthEnd] = 1 )
RETURN STDEVX.P ( Months12, [Monthly Default Rate] )
```

### D04-22 - Default Rate Z-Score
*Alerts. Standard deviations from the 12-month average.*  Format: `0.00`

```dax
Default Rate Z-Score =
DIVIDE ( [Monthly Default Rate] - [Default Rate 12M Average], [Default Rate 12M StdDev] )
```

### D04-23 - Default Rate Trend Alert
*Alerts. Flags metrics moving more than 2 std dev from their 12M average (brief Day 4).*  

```dax
Default Rate Trend Alert =
IF ( ABS ( [Default Rate Z-Score] ) > 2, "ALERT: > 2 std dev", "Within range" )
```

### D04-24 - Delinquency 30+ Z-Score
*Alerts. Z-score of 30+ DPD share vs its 12M history.*  Format: `0.00`

```dax
Delinquency 30+ Z-Score =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
VAR Months12 =
    CALCULATETABLE ( VALUES ( DimDate[Date] ), DATESINPERIOD ( DimDate[Date], LastD, -12, MONTH ), DimDate[IsMonthEnd] = 1 )
VAR Cur = [DPD 30+ % (Date View)]
RETURN DIVIDE ( Cur - AVERAGEX ( Months12, [DPD 30+ % (Date View)] ), STDEVX.P ( Months12, [DPD 30+ % (Date View)] ) )
```

## 5. Day 5 - Advanced DAX, virtual tables, concentration

### D05-01 - Pool Share %
*Strat. Share of the selected pool - use with any stratification axis (LTV band, CIBIL band, vintage).*  Format: `0.0%`

```dax
Pool Share % =
DIVIDE ( [Total Outstanding Balance], CALCULATE ( [Total Outstanding Balance], ALLSELECTED () ) )
```

### D05-02 - LTV > 90% Share
*Strat. Balance share with current LTV above 90%.*  Format: `0.0%`

```dax
LTV > 90% Share =
DIVIDE ( CALCULATE ( [Total Outstanding Balance], FactLoanTape[LTV_Current] > 0.9 ), [Total Outstanding Balance] )
```

### D05-03 - CIBIL < 650 Share
*Strat. Balance share with sub-650 CIBIL.*  Format: `0.0%`

```dax
CIBIL < 650 Share =
DIVIDE ( CALCULATE ( [Total Outstanding Balance], FactLoanTape[CIBIL_Score_Current] < 650 ), [Total Outstanding Balance] )
```

### D05-04 - Top 3 States
*Concentration. Text label of the three largest states with their share.*  

```dax
Top 3 States =
CONCATENATEX (
    TOPN ( 3, VALUES ( DimGeography[State] ), [Total Outstanding Balance], DESC ),
    DimGeography[State] & " " & FORMAT ( DIVIDE ( [Total Outstanding Balance], CALCULATE ( [Total Outstanding Balance], ALLSELECTED ( DimGeography ) ) ), "0.0%" ),
    ", ", [Total Outstanding Balance], DESC )
```

### D05-05 - Top 5 State Concentration
*Concentration. Share of the pool in the five largest states (TOPN virtual table).*  Format: `0.0%`

```dax
Top 5 State Concentration =
DIVIDE ( SUMX ( TOPN ( 5, VALUES ( DimGeography[State] ), [Total Outstanding Balance], DESC ), [Total Outstanding Balance] ), [Total Outstanding Balance] )
```

### D05-06 - Top 10 Loan Concentration
*Concentration. Share held by the ten largest loans.*  Format: `0.0%`

```dax
Top 10 Loan Concentration =
DIVIDE ( SUMX ( TOPN ( 10, VALUES ( DimLoan[LoanID] ), [Total Outstanding Balance], DESC ), [Total Outstanding Balance] ), [Total Outstanding Balance] )
```

### D05-07 - Top 10 Borrower Concentration
*Concentration. Share held by the ten largest borrowers.*  Format: `0.0%`

```dax
Top 10 Borrower Concentration =
DIVIDE ( SUMX ( TOPN ( 10, VALUES ( DimBorrower[BorrowerID] ), [Total Outstanding Balance], DESC ), [Total Outstanding Balance] ), [Total Outstanding Balance] )
```

### D05-08 - HHI State
*Concentration. Herfindahl-Hirschman Index by state (0-1).*  Format: `0.0000`

```dax
HHI State =
VAR Tot = [Total Outstanding Balance]
RETURN SUMX ( VALUES ( DimGeography[State] ), POWER ( DIVIDE ( [Total Outstanding Balance], Tot ), 2 ) )
```

### D05-09 - HHI Servicer
*Concentration. HHI by servicer.*  Format: `0.0000`

```dax
HHI Servicer =
VAR Tot = [Total Outstanding Balance]
RETURN SUMX ( VALUES ( DimLoan[ServicerName] ), POWER ( DIVIDE ( [Total Outstanding Balance], Tot ), 2 ) )
```

### D05-10 - HHI Vehicle Make
*Concentration. HHI by vehicle manufacturer.*  Format: `0.0000`

```dax
HHI Vehicle Make =
VAR Tot = [Total Outstanding Balance]
RETURN SUMX ( VALUES ( DimVehicle[VehicleMake] ), POWER ( DIVIDE ( [Total Outstanding Balance], Tot ), 2 ) )
```

### D05-11 - HHI Borrower
*Concentration. HHI by borrower (borrower IDs are unique in this tape).*  Format: `0.0000`

```dax
HHI Borrower =
VAR Tot = [Total Outstanding Balance]
RETURN SUMX ( VALUES ( DimBorrower[BorrowerID] ), POWER ( DIVIDE ( [Total Outstanding Balance], Tot ), 2 ) )
```

### D05-12 - State Rank
*Ranking. Rank of each state by balance.*  Format: `0`

```dax
State Rank =
RANKX ( ALLSELECTED ( DimGeography[State] ), [Total Outstanding Balance], , DESC, DENSE )
```

### D05-13 - Distinct States
*Ranking. Number of states in the selection.*  Format: `0`

```dax
Distinct States =
DISTINCTCOUNT ( DimLoan[State] )
```

### D05-14 - Portfolio ECL (Context Transition)
*Context transition. Context-transition demonstration: CALCULATE inside SUMX turns each row into a filter context (brief Day 5).*  Format: `#,##0`

```dax
Portfolio ECL (Context Transition) =
SUMX (
    FactLoanTape,
    CALCULATE ( AVERAGE ( FactLoanTape[PD_Estimate] ) * AVERAGE ( FactLoanTape[LGD_Estimate] ) ) * FactLoanTape[EAD]
)
```

## 6. Day 6 - Waterfall, tranche allocation and trigger testing

### D06-01 - Ending Pool Balance
*Pool cash flow. Pool balance at the last period in context.*  Format: `#,##0`

```dax
Ending Pool Balance =
VAR p = MAX ( FactPool[period] )
RETURN CALCULATE ( SUM ( FactPool[ending_pool] ), FactPool[period] = p )
```

### D06-02 - Interest Collected
*Pool cash flow. Interest collections.*  Format: `#,##0`

```dax
Interest Collected =
SUM ( FactPool[interest_collected] )
```

### D06-03 - Servicing Fees
*Pool cash flow. Senior servicing fee (0.5% p.a., assumption).*  Format: `#,##0`

```dax
Servicing Fees =
SUM ( FactPool[servicing_fees] )
```

### D06-04 - Net Loss (Projected)
*Pool cash flow. Projected net loss under the scenario.*  Format: `#,##0`

```dax
Net Loss (Projected) =
SUM ( FactPool[net_loss] )
```

### D06-05 - Excess Spread
*Pool cash flow. Interest left after fees and note coupons.*  Format: `#,##0`

```dax
Excess Spread =
SUM ( FactPool[excess_spread] )
```

### D06-06 - Excess Spread % (annualised)
*Pool cash flow. Annualised excess spread / pool balance.*  Format: `0.00%`

```dax
Excess Spread % (annualised) =
DIVIDE ( SUM ( FactPool[excess_spread] ), SUM ( FactPool[beginning_pool] ) ) * 12
```

### D06-07 - Cumulative Loss % (Projected)
*Pool cash flow. Cumulative projected net loss / opening pool at the period in context.*  Format: `0.00%`

```dax
Cumulative Loss % (Projected) =
VAR p = MAX ( FactPool[period] )
RETURN CALCULATE ( MAX ( FactPool[cumulative_loss_pct] ), FactPool[period] = p )
```

### D06-08 - OC Ratio
*Pool cash flow. Pool balance / (Class A+B+C notes).*  Format: `0.000`

```dax
OC Ratio =
VAR p = MAX ( FactPool[period] )
RETURN CALCULATE ( MAX ( FactPool[oc_ratio] ), FactPool[period] = p )
```

### D06-09 - Tranche Beginning Balance
*Tranche. Opening note balance for the first period in context.*  Format: `#,##0`

```dax
Tranche Beginning Balance =
VAR p = MIN ( FactWaterfall[period] )
RETURN CALCULATE ( SUM ( FactWaterfall[beginning_balance] ), FactWaterfall[period] = p )
```

### D06-10 - Tranche Ending Balance
*Tranche. Closing note balance at the last period in context.*  Format: `#,##0`

```dax
Tranche Ending Balance =
VAR p = MAX ( FactWaterfall[period] )
RETURN CALCULATE ( SUM ( FactWaterfall[ending_balance] ), FactWaterfall[period] = p )
```

### D06-11 - Interest Due
*Tranche. Coupon due.*  Format: `#,##0`

```dax
Interest Due =
SUM ( FactWaterfall[interest_due] )
```

### D06-12 - Interest Paid
*Tranche. Coupon paid out of available interest.*  Format: `#,##0`

```dax
Interest Paid =
SUM ( FactWaterfall[interest_paid] )
```

### D06-13 - Interest Shortfall
*Tranche. Unpaid coupon.*  Format: `#,##0`

```dax
Interest Shortfall =
[Interest Due] - [Interest Paid]
```

### D06-14 - Principal Paid
*Tranche. Principal distributed (sequential A > B > C > Equity).*  Format: `#,##0`

```dax
Principal Paid =
SUM ( FactWaterfall[principal_paid] )
```

### D06-15 - Tranche Writedown
*Tranche. Losses allocated to the tranche (Equity first, Class A last).*  Format: `#,##0`

```dax
Tranche Writedown =
SUM ( FactWaterfall[writedown] )
```

### D06-16 - Writedown % of Tranche
*Tranche. Share of original tranche balance written down.*  Format: `0.0%`

```dax
Writedown % of Tranche =
DIVIDE ( [Tranche Writedown], CALCULATE ( SUM ( FactWaterfall[beginning_balance] ), FactWaterfall[period] = 1 ) )
```

### D06-17 - Credit Enhancement %
*Tranche. Subordination below the tranche (attachment point).*  Format: `0.0%`

```dax
Credit Enhancement % =
MAX ( DimTranche[attachment] )
```

### D06-18 - Tranche Loss % (Attach-Detach)
*Tranche. Closed-form tranche loss from pool loss and attachment/detachment points; cross-checks the cash-flow write-down.*  Format: `0.0%`

```dax
Tranche Loss % (Attach-Detach) =
VAR CumLoss = [Cumulative Loss % (Projected)]
VAR Attach  = MAX ( DimTranche[attachment] )
VAR Detach  = MAX ( DimTranche[detachment] )
RETURN DIVIDE ( MAX ( 0, MIN ( CumLoss, Detach ) - Attach ), Detach - Attach )
```

### D06-19 - Principal Available
*Tranche. Cash available for sequential principal pay-down.*  Format: `#,##0`

```dax
Principal Available =
SUM ( FactPool[scheduled_principal] ) + SUM ( FactPool[prepayments] ) + SUM ( FactPool[recoveries] ) + SUM ( FactPool[excess_spread_used_for_losses] )
```

### D06-20 - Principal to Class A
*Tranche. Sequential allocation: senior class receives principal first.*  Format: `#,##0`

```dax
Principal to Class A =
VAR BalA = CALCULATE ( SUM ( FactWaterfall[beginning_balance] ), DimTranche[tranche] = "Class A (Senior)" )
RETURN MIN ( [Principal Available], BalA )
```

### D06-21 - Principal to Class B
*Tranche. Mezzanine receives what remains after Class A.*  Format: `#,##0`

```dax
Principal to Class B =
VAR BalB = CALCULATE ( SUM ( FactWaterfall[beginning_balance] ), DimTranche[tranche] = "Class B (Mezzanine)" )
RETURN MIN ( MAX ( [Principal Available] - [Principal to Class A], 0 ), BalB )
```

### D06-22 - Tranche Status
*Tranche. Intact / impaired / wiped out label.*  

```dax
Tranche Status =
SWITCH ( TRUE (),
    [Writedown % of Tranche] = 0, "Intact",
    [Writedown % of Tranche] < 0.5, "Impaired",
    "Wiped out" )
```

### D06-23 - Delinq Trigger Value
*Triggers. What-if driven delinq trigger level (default is an ASSUMPTION).*  Format: `0.0%`

```dax
Delinq Trigger Value =
SELECTEDVALUE ( ParamDelinqTrigger[Value], 0.12 )
```

### D06-24 - CNL Trigger Value
*Triggers. What-if driven cnl trigger level (default is an ASSUMPTION).*  Format: `0.0%`

```dax
CNL Trigger Value =
SELECTEDVALUE ( ParamCNLTrigger[Value], 0.02 )
```

### D06-25 - OC Trigger Value
*Triggers. What-if driven oc trigger level (default is an ASSUMPTION).*  Format: `0.00`

```dax
OC Trigger Value =
SELECTEDVALUE ( ParamOCTrigger[Value], 1.05 )
```

### D06-26 - Delinquency Trigger Status
*Triggers. Tape 60+ DPD % vs delinquency trigger.*  

```dax
Delinquency Trigger Status =
IF ( [Delinquent 60+ %] > [Delinq Trigger Value], "BREACH", "OK" )
```

### D06-27 - Delinquency Trigger Headroom
*Triggers. Trigger level minus actual (negative = breach).*  Format: `0.0%`

```dax
Delinquency Trigger Headroom =
[Delinq Trigger Value] - [Delinquent 60+ %]
```

### D06-28 - CNL Trigger Status
*Triggers. Projected cumulative net loss vs trigger.*  

```dax
CNL Trigger Status =
IF ( [Cumulative Loss % (Projected)] > [CNL Trigger Value], "BREACH", "OK" )
```

### D06-29 - OC Trigger Status
*Triggers. OC ratio vs minimum.*  

```dax
OC Trigger Status =
IF ( [OC Ratio] < [OC Trigger Value], "BREACH", "OK" )
```

### D06-30 - Trigger Breach Count
*Triggers. Number of triggers breached.*  Format: `0`

```dax
Trigger Breach Count =
( [Delinquency Trigger Status] = "BREACH" ) + ( [CNL Trigger Status] = "BREACH" ) + ( [OC Trigger Status] = "BREACH" )
```

### D06-31 - CNL Trigger Breach Month
*Triggers. First projected month in which cumulative loss exceeds the trigger (blank = never).*  Format: `0`

```dax
CNL Trigger Breach Month =
VAR t = [CNL Trigger Value]
RETURN CALCULATE ( MIN ( DimPeriod[period] ), FILTER ( ALL ( DimPeriod[period] ), [Cumulative Loss % (Projected)] > t ) )
```

## 7. Day 7 - DPD distribution, roll rates and ageing

### D07-01 - Snapshot Loans
*Distribution. Loans with a post-origination snapshot (use SnapshotLabel on the axis).*  Format: `#,##0`

```dax
Snapshot Loans =
CALCULATE ( DISTINCTCOUNT ( FactMonthlyPerformance[LoanID] ), FactMonthlyPerformance[PreOrigination] = 0 )
```

### D07-02 - Snapshot Balance
*Distribution. Balance on the snapshots in context (pre-origination rows excluded).*  Format: `#,##0`

```dax
Snapshot Balance =
CALCULATE ( SUM ( FactMonthlyPerformance[CurrentBalance] ), FactMonthlyPerformance[PreOrigination] = 0 )
```

### D07-03 - DPD 30+ Balance (Snapshot)
*Distribution. 30+ DPD balance by snapshot.*  Format: `#,##0`

```dax
DPD 30+ Balance (Snapshot) =
CALCULATE ( [Snapshot Balance], FactMonthlyPerformance[BucketOrder] >= 2 )
```

### D07-04 - DPD 30+ %
*Distribution. 30+ DPD share by snapshot (axis = SnapshotLabel).*  Format: `0.0%`

```dax
DPD 30+ % =
DIVIDE ( [DPD 30+ Balance (Snapshot)], [Snapshot Balance] )
```

### D07-05 - DPD 60+ %
*Distribution. 60+ DPD share by snapshot.*  Format: `0.0%`

```dax
DPD 60+ % =
DIVIDE ( CALCULATE ( [Snapshot Balance], FactMonthlyPerformance[BucketOrder] >= 3 ), [Snapshot Balance] )
```

### D07-06 - DPD 90+ %
*Distribution. 90+ DPD share by snapshot.*  Format: `0.0%`

```dax
DPD 90+ % =
DIVIDE ( CALCULATE ( [Snapshot Balance], FactMonthlyPerformance[BucketOrder] >= 4 ), [Snapshot Balance] )
```

### D07-07 - Bucket Balance %
*Distribution. Share of snapshot balance by DPD bucket.*  Format: `0.0%`

```dax
Bucket Balance % =
DIVIDE ( [Snapshot Balance], CALCULATE ( [Snapshot Balance], ALL ( DimBucket ) ) )
```

### D07-08 - DPD 30+ % (Date View)
*Distribution. 30+ DPD share on a calendar axis (duplicate May snapshot excluded).*  Format: `0.0%`

```dax
DPD 30+ % (Date View) =
CALCULATE ( [DPD 30+ %], FactMonthlyPerformance[IncludeInTrend] = 1 )
```

### D07-09 - Transition Count
*Roll rates. Loan-month transitions between consecutive snapshots.*  Format: `#,##0`

```dax
Transition Count =
CALCULATE ( COUNTROWS ( FactMonthlyPerformance ), FactMonthlyPerformance[ValidTransition] = 1 )
```

### D07-10 - Roll Rate %
*Roll rates. Matrix measure: rows = FactMonthlyPerformance[PriorBucketSeq], columns = DimBucket[DPD_Bucket]; % of the prior bucket moving to each bucket.*  Format: `0.0%`

```dax
Roll Rate % =
DIVIDE ( [Transition Count], CALCULATE ( [Transition Count], ALL ( DimBucket[DPD_Bucket] ) ) )
```

### D07-11 - Worsened Transitions
*Roll rates. Transitions to a worse bucket.*  Format: `#,##0`

```dax
Worsened Transitions =
CALCULATE ( [Transition Count], FILTER ( FactMonthlyPerformance, FactMonthlyPerformance[BucketOrder] > FactMonthlyPerformance[PriorBucketOrder] ) )
```

### D07-12 - Cured Transitions
*Roll rates. Transitions to a better bucket.*  Format: `#,##0`

```dax
Cured Transitions =
CALCULATE ( [Transition Count], FILTER ( FactMonthlyPerformance, FactMonthlyPerformance[BucketOrder] < FactMonthlyPerformance[PriorBucketOrder] ) )
```

### D07-13 - Roll Forward Rate
*Roll rates. Share of transitions that worsen.*  Format: `0.0%`

```dax
Roll Forward Rate =
DIVIDE ( [Worsened Transitions], [Transition Count] )
```

### D07-14 - Cure Rate
*Roll rates. Share of transitions that improve.*  Format: `0.0%`

```dax
Cure Rate =
DIVIDE ( [Cured Transitions], [Transition Count] )
```

### D07-15 - Cure Rate (30+ DPD)
*Roll rates. Cure rate for loans already 30+ DPD (zero in this data).*  Format: `0.0%`

```dax
Cure Rate (30+ DPD) =
DIVIDE ( CALCULATE ( [Cured Transitions], FactMonthlyPerformance[PriorBucketOrder] >= 2 ), CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketOrder] >= 2 ) )
```

### D07-16 - Roll Rate Current to Current
*Roll rates. Current loans staying current.*  Format: `0.0%`

```dax
Roll Rate Current to Current =
DIVIDE ( CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "Current", FactMonthlyPerformance[DPD_Bucket] = "Current" ), CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "Current" ) )
```

### D07-17 - Roll Rate Current to 1-29
*Roll rates. Current loans slipping into 1-29 DPD.*  Format: `0.0%`

```dax
Roll Rate Current to 1-29 =
DIVIDE ( CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "Current", FactMonthlyPerformance[DPD_Bucket] = "1-29 DPD" ), CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "Current" ) )
```

### D07-18 - Roll Rate 30-59 to 60-89
*Roll rates. 30-59 DPD loans rolling to 60-89.*  Format: `0.0%`

```dax
Roll Rate 30-59 to 60-89 =
DIVIDE ( CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "30-59 DPD", FactMonthlyPerformance[DPD_Bucket] = "60-89 DPD" ), CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "30-59 DPD" ) )
```

### D07-19 - Roll Rate 60-89 to 90+
*Roll rates. 60-89 DPD loans rolling into 90+ / default.*  Format: `0.0%`

```dax
Roll Rate 60-89 to 90+ =
DIVIDE ( CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "60-89 DPD", FactMonthlyPerformance[BucketOrder] >= 4 ), CALCULATE ( [Transition Count], FactMonthlyPerformance[PriorBucketSeq] = "60-89 DPD" ) )
```

### D07-20 - Balance Delinquent 1-2 Months
*Ageing. Balance delinquent 1-2 consecutive months.*  Format: `#,##0`

```dax
Balance Delinquent 1-2 Months =
CALCULATE ( [Snapshot Balance], FactMonthlyPerformance[ConsecutiveMonthsDelinquent] >= 1, FactMonthlyPerformance[ConsecutiveMonthsDelinquent] <= 2 )
```

### D07-21 - Balance Delinquent 3-5 Months
*Ageing. Balance delinquent 3-5 consecutive months.*  Format: `#,##0`

```dax
Balance Delinquent 3-5 Months =
CALCULATE ( [Snapshot Balance], FactMonthlyPerformance[ConsecutiveMonthsDelinquent] >= 3, FactMonthlyPerformance[ConsecutiveMonthsDelinquent] <= 5 )
```

### D07-22 - Balance Delinquent 6+ Months
*Ageing. Balance delinquent 6+ consecutive months.*  Format: `#,##0`

```dax
Balance Delinquent 6+ Months =
CALCULATE ( [Snapshot Balance], FactMonthlyPerformance[ConsecutiveMonthsDelinquent] >= 6 )
```

### D07-23 - Avg Months Delinquent
*Ageing. Average consecutive months delinquent among delinquent loans.*  Format: `0.0`

```dax
Avg Months Delinquent =
CALCULATE ( AVERAGE ( FactMonthlyPerformance[ConsecutiveMonthsDelinquent] ), FactMonthlyPerformance[ConsecutiveMonthsDelinquent] > 0, FactMonthlyPerformance[PreOrigination] = 0 )
```

### D07-24 - Loans Delinquent 3+ Months (Latest)
*Ageing. Loans in arrears 3+ months at the latest snapshot.*  Format: `#,##0`

```dax
Loans Delinquent 3+ Months (Latest) =
VAR s = MAX ( FactMonthlyPerformance[SnapshotSeq] )
RETURN CALCULATE ( DISTINCTCOUNT ( FactMonthlyPerformance[LoanID] ), FactMonthlyPerformance[SnapshotSeq] = s, FactMonthlyPerformance[ConsecutiveMonthsDelinquent] >= 3 )
```

## 8. Day 8 - Vintage / static pool analysis

### D08-01 - Vintage Original Balance
*Vintage. Original pool balance (counted once per vintage at MOB 0).*  Format: `#,##0`

```dax
Vintage Original Balance =
CALCULATE ( SUM ( FactStaticPool[OriginalPoolBalance] ), FactStaticPool[MonthsOnBook] = 0 )
```

### D08-02 - CNL Rate @ MOB
*Vintage. Cumulative net loss rate - put MonthsOnBook on the axis (one MOB per point).*  Format: `0.00%`

```dax
CNL Rate @ MOB =
DIVIDE ( SUM ( FactStaticPool[CumulativeNetLoss] ), SUM ( FactStaticPool[OriginalPoolBalance] ) )
```

### D08-03 - Cumulative Default Rate @ MOB
*Vintage. Cumulative default rate by months on book.*  Format: `0.00%`

```dax
Cumulative Default Rate @ MOB =
DIVIDE ( SUM ( FactStaticPool[CumulativeDefaults_Balance] ), SUM ( FactStaticPool[OriginalPoolBalance] ) )
```

### D08-04 - Cumulative Gross Loss Rate @ MOB
*Vintage. Cumulative gross loss rate by months on book.*  Format: `0.00%`

```dax
Cumulative Gross Loss Rate @ MOB =
DIVIDE ( SUM ( FactStaticPool[CumulativeGrossLoss] ), SUM ( FactStaticPool[OriginalPoolBalance] ) )
```

### D08-05 - Cumulative Prepayment Rate @ MOB
*Vintage. Cumulative prepayment rate by months on book.*  Format: `0.00%`

```dax
Cumulative Prepayment Rate @ MOB =
DIVIDE ( SUM ( FactStaticPool[CumulativePrepayments] ), SUM ( FactStaticPool[OriginalPoolBalance] ) )
```

### D08-06 - Pool Factor @ MOB
*Vintage. Remaining balance / original by months on book.*  Format: `0.0%`

```dax
Pool Factor @ MOB =
DIVIDE ( SUM ( FactStaticPool[RemainingPoolBalance] ), SUM ( FactStaticPool[OriginalPoolBalance] ) )
```

### D08-07 - Latest MOB
*Vintage. Most recent months-on-book observed.*  Format: `0`

```dax
Latest MOB =
MAX ( FactStaticPool[MonthsOnBook] )
```

### D08-08 - CNL Rate (Latest MOB)
*Vintage. Cumulative net loss to date, each vintage at its own latest MOB (virtual-table iteration).*  Format: `0.00%`

```dax
CNL Rate (Latest MOB) =
DIVIDE (
    SUMX ( VALUES ( FactStaticPool[VintageID] ),
        VAR m = CALCULATE ( MAX ( FactStaticPool[MonthsOnBook] ) )
        RETURN CALCULATE ( SUM ( FactStaticPool[CumulativeNetLoss] ), FactStaticPool[MonthsOnBook] = m ) ),
    [Vintage Original Balance] )
```

### D08-09 - Loss Timing % (of CNL to date)
*Vintage. Share of loss-to-date incurred by each MOB - the loss timing curve rating agencies ask for.*  Format: `0.0%`

```dax
Loss Timing % (of CNL to date) =
DIVIDE ( [CNL Rate @ MOB], [CNL Rate (Latest MOB)] )
```

### D08-10 - Marginal Loss Rate (Wtd)
*Vintage. Monthly marginal loss rate weighted by remaining balance.*  Format: `0.000%`

```dax
Marginal Loss Rate (Wtd) =
DIVIDE ( SUMX ( FactStaticPool, FactStaticPool[MarginalLossRate] * FactStaticPool[RemainingPoolBalance] ), SUM ( FactStaticPool[RemainingPoolBalance] ) )
```

### D08-11 - Vintage 30+ Delinquency (Wtd)
*Vintage. 30+ delinquency weighted by remaining balance.*  Format: `0.00%`

```dax
Vintage 30+ Delinquency (Wtd) =
DIVIDE ( SUMX ( FactStaticPool, FactStaticPool[CurrentDelinq30Plus] * FactStaticPool[RemainingPoolBalance] ), SUM ( FactStaticPool[RemainingPoolBalance] ) )
```

### D08-12 - Projected Ultimate CNL (MOB 36)
*Vintage. Chain-ladder style projection: scales the vintage's CNL by the mature-vintage (2021) development from its MOB to MOB 36.*  Format: `0.00%`

```dax
Projected Ultimate CNL (MOB 36) =
VAR MaxMOB  = MAX ( FactStaticPool[MonthsOnBook] )
VAR Cur     = CALCULATE ( [CNL Rate @ MOB], FactStaticPool[MonthsOnBook] = MaxMOB )
VAR RefNow  = CALCULATE ( [CNL Rate @ MOB], FactStaticPool[MonthsOnBook] = MaxMOB, FILTER ( ALL ( FactStaticPool[VintageID] ), FactStaticPool[VintageID] IN { "2021-Q1", "2021-Q2", "2021-Q3", "2021-Q4" } ) )
VAR Ref36   = CALCULATE ( [CNL Rate @ MOB], FactStaticPool[MonthsOnBook] = 36, FILTER ( ALL ( FactStaticPool[VintageID] ), FactStaticPool[VintageID] IN { "2021-Q1", "2021-Q2", "2021-Q3", "2021-Q4" } ) )
RETURN IF ( MaxMOB >= 36, Cur, DIVIDE ( Cur * Ref36, RefNow ) )
```

### D08-13 - Vintage WAC
*Vintage. Vintage-level WAC: filter flows DimVintage > DimLoan > tape, so the base measure is reused.*  Format: `0.00%`

```dax
Vintage WAC =
[WAC]
```

### D08-14 - Vintage WALA
*Vintage. Vintage-level WALA via the same filter path.*  Format: `0.0`

```dax
Vintage WALA =
[WALA]
```

## 9. Day 9 - IFRS 9 ECL framework

### D09-01 - Total ECL (Tape)
*ECL (tape). Provision as supplied on the tape.*  Format: `#,##0`

```dax
Total ECL (Tape) =
SUM ( FactLoanTape[ECL_Provision] )
```

### D09-02 - ECL Coverage %
*ECL (tape). Provision / EAD.*  Format: `0.00%`

```dax
ECL Coverage % =
DIVIDE ( [Total ECL (Tape)], [Total EAD] )
```

### D09-03 - ECL Stage 1
*ECL (tape). Provision on Stage 1 loans.*  Format: `#,##0`

```dax
ECL Stage 1 =
CALCULATE ( [Total ECL (Tape)], FactLoanTape[IFRS9_Stage] = 1 )
```

### D09-04 - ECL Stage 2
*ECL (tape). Provision on Stage 2 loans.*  Format: `#,##0`

```dax
ECL Stage 2 =
CALCULATE ( [Total ECL (Tape)], FactLoanTape[IFRS9_Stage] = 2 )
```

### D09-05 - ECL Stage 3
*ECL (tape). Provision on Stage 3 loans.*  Format: `#,##0`

```dax
ECL Stage 3 =
CALCULATE ( [Total ECL (Tape)], FactLoanTape[IFRS9_Stage] = 3 )
```

### D09-06 - EAD Stage 1
*ECL (tape). EAD of Stage 1 loans.*  Format: `#,##0`

```dax
EAD Stage 1 =
CALCULATE ( [Total EAD], FactLoanTape[IFRS9_Stage] = 1 )
```

### D09-07 - EAD Stage 2
*ECL (tape). EAD of Stage 2 loans.*  Format: `#,##0`

```dax
EAD Stage 2 =
CALCULATE ( [Total EAD], FactLoanTape[IFRS9_Stage] = 2 )
```

### D09-08 - EAD Stage 3
*ECL (tape). EAD of Stage 3 loans.*  Format: `#,##0`

```dax
EAD Stage 3 =
CALCULATE ( [Total EAD], FactLoanTape[IFRS9_Stage] = 3 )
```

### D09-09 - Stage Mix % (EAD)
*ECL (tape). Stage share of EAD (DimStage on axis).*  Format: `0.0%`

```dax
Stage Mix % (EAD) =
DIVIDE ( [Total EAD], CALCULATE ( [Total EAD], ALL ( DimStage ) ) )
```

### D09-10 - Avg PD (EAD-weighted)
*ECL (tape). EAD-weighted PD.*  Format: `0.00%`

```dax
Avg PD (EAD-weighted) =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[PD_Estimate] * FactLoanTape[EAD] ), [Total EAD] )
```

### D09-11 - Avg LGD (EAD-weighted)
*ECL (tape). EAD-weighted LGD.*  Format: `0.0%`

```dax
Avg LGD (EAD-weighted) =
DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[LGD_Estimate] * FactLoanTape[EAD] ), [Total EAD] )
```

### D09-12 - ECL Recomputed (PD x LGD x EAD)
*ECL (tape). Independent loan-level recomputation of the provision.*  Format: `#,##0`

```dax
ECL Recomputed (PD x LGD x EAD) =
SUMX ( FactLoanTape, FactLoanTape[PD_Estimate] * FactLoanTape[LGD_Estimate] * FactLoanTape[EAD] )
```

### D09-13 - ECL Recompute Variance
*ECL (tape). Should be zero - tape ECL = PD x LGD x EAD.*  Format: `#,##0`

```dax
ECL Recompute Variance =
[ECL Recomputed (PD x LGD x EAD)] - [Total ECL (Tape)]
```

### D09-14 - Realised Recovery Rate (Tape)
*LGD. Recoveries / gross loss on the tape (2 loans only).*  Format: `0.0%`

```dax
Realised Recovery Rate (Tape) =
DIVIDE ( SUM ( FactLoanTape[RecoveryAmount] ), SUM ( FactLoanTape[LossAmount] ) )
```

### D09-15 - Net Loss Severity (Dynamic)
*LGD. Net loss / defaulted balance over 12 months - the pool LGD used by the challenger model.*  Format: `0.0%`

```dax
Net Loss Severity (Dynamic) =
DIVIDE ( SUM ( FactDynamicLoss[NetLoss_ThisMonth] ), SUM ( FactDynamicLoss[NewDefaults_Balance] ) )
```

### D09-16 - Gross Loss Severity (Dynamic)
*LGD. Gross loss / defaulted balance.*  Format: `0.0%`

```dax
Gross Loss Severity (Dynamic) =
DIVIDE ( SUM ( FactDynamicLoss[GrossLoss_ThisMonth] ), SUM ( FactDynamicLoss[NewDefaults_Balance] ) )
```

### D09-17 - PD Term Structure
*PD. Cumulative PD from the empirical transition matrix (axis: HorizonMonths, legend: DPD_Bucket).*  Format: `0.0%`

```dax
PD Term Structure =
MAX ( FactPDTermStructure[CumulativePD] )
```

### D09-18 - 12M PD from Transition Matrix
*PD. 12-month PD by starting bucket.*  Format: `0.0%`

```dax
12M PD from Transition Matrix =
CALCULATE ( MAX ( FactPDTermStructure[CumulativePD] ), FactPDTermStructure[HorizonMonths] = 12 )
```

### D09-19 - Param Pool LGD
*PD. Pool LGD read from DimParameters (calibrated on dynamic_loss_monthly).*  Format: `0.0%`

```dax
Param Pool LGD =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Pool net LGD (severity)" )
```

### D09-20 - Param Base CPR
*PD. Base annual CPR read from DimParameters.*  Format: `0.0%`

```dax
Param Base CPR =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Base annual CPR" )
```

### D09-21 - Param Rate Elasticity
*PD. ASSUMPTION: PD uplift per 100bps rate shock.*  Format: `0.00`

```dax
Param Rate Elasticity =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Rate shock PD elasticity (per 100bps)" )
```

### D09-22 - Challenger ECL (DAX)
*Challenger. Independent IFRS 9 ECL: Stage 1 = 12M PD, Stage 2 = lifetime PD (capped 60M) from the roll-rate matrix, Stage 3 = 100%; LGD from the dynamic loss data; discounted at EIR.*  Format: `#,##0`

```dax
Challenger ECL (DAX) =
VAR LGDp = [Param Pool LGD]
RETURN
SUMX (
    FactLoanTape,
    VAR stg = FactLoanTape[IFRS9_Stage]
    VAR h   = IF ( stg = 1, 12, MIN ( MAX ( FactLoanTape[RemainingTerm], 1 ), 60 ) )
    VAR pdv = IF ( stg = 3, 1,
                LOOKUPVALUE ( FactPDTermStructure[CumulativePD],
                              FactPDTermStructure[DPD_Bucket], FactLoanTape[DelinquencyStatus],
                              FactPDTermStructure[HorizonMonths], h ) )
    VAR hh  = IF ( stg = 3, 6, h )
    VAR dfc = POWER ( 1 + RELATED ( DimLoan[InterestRate] ) / 100, - ( hh / 12 ) / 2 )
    RETURN IF ( FactLoanTape[EAD] <= 0, 0, pdv * LGDp * FactLoanTape[EAD] * dfc )
)
```

### D09-23 - Challenger vs Tape Variance %
*Challenger. How far the data-driven ECL sits from the tape provision.*  Format: `0.0%`

```dax
Challenger vs Tape Variance % =
DIVIDE ( [Challenger ECL (DAX)] - [Total ECL (Tape)], [Total ECL (Tape)] )
```

### D09-24 - Stage 2 DPD Value
*Staging. Stage 2 DPD threshold (tape rule 30; brief 60).*  Format: `0`

```dax
Stage 2 DPD Value =
SELECTEDVALUE ( ParamStage2DPD[Value], 30 )
```

### D09-25 - Stage 3 DPD Value
*Staging. Stage 3 DPD threshold (tape rule 90; brief 120).*  Format: `0`

```dax
Stage 3 DPD Value =
SELECTEDVALUE ( ParamStage3DPD[Value], 90 )
```

### D09-26 - ECL (Alt Staging)
*Staging. Re-stages loans with adjustable DPD thresholds and applies stage-average PDs; compare the brief's 60/120 rule with the tape's 30/90 (approximation - use for impact, not as the provision).*  Format: `#,##0`

```dax
ECL (Alt Staging) =
VAR S2  = [Stage 2 DPD Value]
VAR S3  = [Stage 3 DPD Value]
VAR pd1 = CALCULATE ( DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[PD_Estimate] * FactLoanTape[EAD] ), SUM ( FactLoanTape[EAD] ) ), REMOVEFILTERS (), FactLoanTape[IFRS9_Stage] = 1 )
VAR pd2 = CALCULATE ( DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[PD_Estimate] * FactLoanTape[EAD] ), SUM ( FactLoanTape[EAD] ) ), REMOVEFILTERS (), FactLoanTape[IFRS9_Stage] = 2 )
VAR pd3 = CALCULATE ( DIVIDE ( SUMX ( FactLoanTape, FactLoanTape[PD_Estimate] * FactLoanTape[EAD] ), SUM ( FactLoanTape[EAD] ) ), REMOVEFILTERS (), FactLoanTape[IFRS9_Stage] = 3 )
RETURN
SUMX ( FactLoanTape,
    VAR s = IF ( FactLoanTape[IsDefaulted] = 1 || FactLoanTape[DelinquencyDays] >= S3, 3, IF ( FactLoanTape[DelinquencyDays] >= S2, 2, 1 ) )
    RETURN SWITCH ( s, 1, pd1, 2, pd2, pd3 ) * FactLoanTape[LGD_Estimate] * FactLoanTape[EAD] )
```

## 10. Day 10 - Stress testing and what-if

### D10-01 - Default Rate Multiplier Value
*What-if. PD multiplier slider.*  Format: `0.0`

```dax
Default Rate Multiplier Value =
SELECTEDVALUE ( ParamDefaultMult[Value], 1 )
```

### D10-02 - Recovery Haircut Value
*What-if. Additive LGD increase slider (recovery haircut).*  Format: `0%`

```dax
Recovery Haircut Value =
SELECTEDVALUE ( ParamRecoveryHaircut[Value], 0 )
```

### D10-03 - CPR Change Value
*What-if. Change in annual CPR slider.*  Format: `0%`

```dax
CPR Change Value =
SELECTEDVALUE ( ParamCPRChange[Value], 0 )
```

### D10-04 - Rate Shock Value
*What-if. Interest-rate shock (bps) slider.*  Format: `0`

```dax
Rate Shock Value =
SELECTEDVALUE ( ParamRateShock[Value], 0 )
```

### D10-05 - Effective PD Multiplier
*What-if. Default multiplier compounded with the rate-shock PD elasticity (assumption).*  Format: `0.00`

```dax
Effective PD Multiplier =
[Default Rate Multiplier Value] * ( 1 + [Param Rate Elasticity] * [Rate Shock Value] / 100 )
```

### D10-06 - Stress ECL (What-If)
*What-if. ECL under the four what-if levers (equals tape ECL when all levers are at their defaults). Faster prepayment shortens exposure (half-year average run-off).*  Format: `#,##0`

```dax
Stress ECL (What-If) =
VAR PDM = [Effective PD Multiplier]
VAR Hc  = [Recovery Haircut Value]
VAR cpr0 = [Param Base CPR]
VAR RunOff = POWER ( DIVIDE ( 1 - ( cpr0 + [CPR Change Value] ), 1 - cpr0 ), 0.5 )
RETURN
SUMX ( FactLoanTape, MIN ( 1, FactLoanTape[PD_Estimate] * PDM ) * MIN ( 1, FactLoanTape[LGD_Estimate] + Hc ) * FactLoanTape[EAD] * RunOff )
```

### D10-07 - Stress ECL Uplift
*What-if. Increase vs tape provision.*  Format: `#,##0`

```dax
Stress ECL Uplift =
[Stress ECL (What-If)] - [Total ECL (Tape)]
```

### D10-08 - Stress ECL Uplift %
*What-if. Percentage increase vs tape provision.*  Format: `0.0%`

```dax
Stress ECL Uplift % =
DIVIDE ( [Stress ECL Uplift], [Total ECL (Tape)] )
```

### D10-09 - Stress Coverage %
*What-if. Stressed provision / EAD.*  Format: `0.00%`

```dax
Stress Coverage % =
DIVIDE ( [Stress ECL (What-If)], [Total EAD] )
```

### D10-10 - Scenario ECL
*Scenarios. ECL under the scenario in context (DimScenario on axis); rate-shock effect is embedded in the scenario PD multiplier.*  Format: `#,##0`

```dax
Scenario ECL =
VAR PDM = SELECTEDVALUE ( DimScenario[pd_multiplier], 1 )
VAR Hc  = SELECTEDVALUE ( DimScenario[lgd_add], 0 )
RETURN SUMX ( FactLoanTape, MIN ( 1, FactLoanTape[PD_Estimate] * PDM ) * MIN ( 1, FactLoanTape[LGD_Estimate] + Hc ) * FactLoanTape[EAD] )
```

### D10-11 - Probability-Weighted Scenario ECL
*Scenarios. IFRS 9 probability-weighted ECL across scenarios.*  Format: `#,##0`

```dax
Probability-Weighted Scenario ECL =
SUMX ( ALL ( DimScenario ), DimScenario[weight] * [Scenario ECL] )
```

### D10-12 - Weighted ECL vs Tape
*Scenarios. Scenario-weighting overlay vs tape provision.*  Format: `#,##0`

```dax
Weighted ECL vs Tape =
[Probability-Weighted Scenario ECL] - [Total ECL (Tape)]
```

### D10-13 - Cum Loss % at Horizon
*Scenarios. Cumulative projected net loss at the end of the 36-month horizon.*  Format: `0.00%`

```dax
Cum Loss % at Horizon =
VAR p = CALCULATE ( MAX ( FactPool[period] ), ALL ( FactPool ) )
RETURN CALCULATE ( MAX ( FactPool[cumulative_loss_pct] ), FactPool[period] = p )
```

### D10-14 - Senior Tranche Loss
*Scenarios. Loss absorbed by Class A.*  Format: `#,##0`

```dax
Senior Tranche Loss =
CALCULATE ( [Tranche Writedown], DimTranche[tranche] = "Class A (Senior)" )
```

### D10-15 - Mezzanine Tranche Loss
*Scenarios. Loss absorbed by Class B.*  Format: `#,##0`

```dax
Mezzanine Tranche Loss =
CALCULATE ( [Tranche Writedown], DimTranche[tranche] = "Class B (Mezzanine)" )
```

### D10-16 - Junior Tranche Loss %
*Scenarios. Class C write-down %.*  Format: `0.0%`

```dax
Junior Tranche Loss % =
CALCULATE ( [Writedown % of Tranche], DimTranche[tranche] = "Class C (Junior)" )
```

### D10-17 - Equity Loss %
*Scenarios. Equity write-down %.*  Format: `0.0%`

```dax
Equity Loss % =
CALCULATE ( [Writedown % of Tranche], DimTranche[tranche] = "Equity / Residual" )
```

### D10-18 - Minimum OC Ratio
*Scenarios. Lowest OC ratio over the horizon.*  Format: `0.000`

```dax
Minimum OC Ratio =
CALCULATE ( MIN ( FactPool[oc_ratio] ) )
```

### D10-19 - CNL Trigger Headroom
*Scenarios. Distance to the CNL trigger at horizon (negative = breached).*  Format: `0.00%`

```dax
CNL Trigger Headroom =
[CNL Trigger Value] - [Cum Loss % at Horizon]
```

### D10-20 - OC Trigger Headroom
*Scenarios. Distance to the OC trigger at the tightest point.*  Format: `0.000`

```dax
OC Trigger Headroom =
[Minimum OC Ratio] - [OC Trigger Value]
```

### D10-21 - Arena W Equity
*Arena. Points lost per full equity loss / breach (DimParameters).*  Format: `0`

```dax
Arena W Equity =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena weight - Equity" )
```

### D10-22 - Arena W Junior
*Arena. Points lost per full junior loss / breach (DimParameters).*  Format: `0`

```dax
Arena W Junior =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena weight - Junior" )
```

### D10-23 - Arena W Mezzanine
*Arena. Points lost per full mezzanine loss / breach (DimParameters).*  Format: `0`

```dax
Arena W Mezzanine =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena weight - Mezzanine" )
```

### D10-24 - Arena W Senior
*Arena. Points lost per full senior loss / breach (DimParameters).*  Format: `0`

```dax
Arena W Senior =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena weight - Senior" )
```

### D10-25 - Arena W Trigger breach
*Arena. Points lost per full trigger breach loss / breach (DimParameters).*  Format: `0`

```dax
Arena W Trigger breach =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena weight - Trigger breach" )
```

### D10-26 - Arena Min Fortress
*Arena. Minimum score for the Fortress rating (DimParameters).*  Format: `0`

```dax
Arena Min Fortress =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena rating - Fortress (min score)" )
```

### D10-27 - Arena Min Resilient
*Arena. Minimum score for the Resilient rating (DimParameters).*  Format: `0`

```dax
Arena Min Resilient =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena rating - Resilient (min score)" )
```

### D10-28 - Arena Min Stressed
*Arena. Minimum score for the Stressed rating (DimParameters).*  Format: `0`

```dax
Arena Min Stressed =
CALCULATE ( MAX ( DimParameters[Value] ), DimParameters[Parameter] = "Arena rating - Stressed (min score)" )
```

### D10-29 - Mezzanine Loss %
*Arena. Class B write-down %.*  Format: `0.0%`

```dax
Mezzanine Loss % =
CALCULATE ( [Writedown % of Tranche], DimTranche[tranche] = "Class B (Mezzanine)" )
```

### D10-30 - Senior Loss %
*Arena. Class A write-down %.*  Format: `0.0%`

```dax
Senior Loss % =
CALCULATE ( [Writedown % of Tranche], DimTranche[tranche] = "Class A (Senior)" )
```

### D10-31 - Arena Score
*Arena. Crisis-survival score 0-100 for the scenario in context: starts at 100, loses points for write-downs by seniority and for each trigger breached.*  Format: `0`

```dax
Arena Score =
MAX ( 0, ROUND ( 100
    - [Arena W Equity]         * [Equity Loss %]
    - [Arena W Junior]         * [Junior Tranche Loss %]
    - [Arena W Mezzanine]      * [Mezzanine Loss %]
    - [Arena W Senior]         * [Senior Loss %]
    - [Arena W Trigger breach] * [Trigger Breach Count], 0 ) )
```

### D10-32 - Arena Rating
*Arena. Rating band from the score.*  

```dax
Arena Rating =
SWITCH ( TRUE (),
    [Arena Score] >= [Arena Min Fortress],  "Fortress",
    [Arena Score] >= [Arena Min Resilient], "Resilient",
    [Arena Score] >= [Arena Min Stressed],  "Stressed",
    "Crisis" )
```

### D10-33 - Arena Message
*Arena. Narrative line for the Arena scoreboard card.*  

```dax
Arena Message =
"Scenario: " & SELECTEDVALUE ( DimScenario[scenario], "All" )
    & " | Cum. loss " & FORMAT ( [Cum Loss % at Horizon], "0.0%" )
    & " | Score " & [Arena Score] & " - " & [Arena Rating]
```

## 11. Day 11 - Investor reporting and data tape

### D11-01 - Cumulative Net Loss (Reporting)
*Loss & recovery. Cumulative net loss over the reporting history.*  Format: `#,##0`

```dax
Cumulative Net Loss (Reporting) =
VAR LastD = MAX ( FactDynamicLoss[MonthEnd] )
RETURN CALCULATE ( [Net Loss (Monthly)], FILTER ( ALL ( DimDate[Date] ), DimDate[Date] <= LastD ) )
```

### D11-02 - Cumulative Net Loss % of Original
*Loss & recovery. Cumulative net loss / original pool balance.*  Format: `0.00%`

```dax
Cumulative Net Loss % of Original =
DIVIDE ( [Cumulative Net Loss (Reporting)], CALCULATE ( [Original Balance], REMOVEFILTERS () ) )
```

### D11-03 - Annualised Default Rate (All Months)
*Loss & recovery. Annualised default rate using all 12 monthly rows (used to calibrate the base CDR).*  Format: `0.00%`

```dax
Annualised Default Rate (All Months) =
1 - POWER ( 1 - DIVIDE ( SUM ( FactDynamicLoss[NewDefaults_Balance] ), SUM ( FactDynamicLoss[BOP_Balance] ) ), 12 )
```

### D11-04 - SMM
*Prepayment. Single monthly mortality.*  Format: `0.000%`

```dax
SMM =
DIVIDE ( SUM ( FactDynamicLoss[Prepayments_ThisMonth] ), SUM ( FactDynamicLoss[BOP_Balance] ) )
```

### D11-05 - CPR (Annualised)
*Prepayment. Annualised conditional prepayment rate.*  Format: `0.0%`

```dax
CPR (Annualised) =
1 - POWER ( 1 - [SMM], 12 )
```

### D11-06 - Collection Efficiency
*Collections. Collections / billing.*  Format: `0.0%`

```dax
Collection Efficiency =
DIVIDE ( SUM ( FactDynamicLoss[CollectionsTotal] ), SUM ( FactDynamicLoss[BillingAmount] ) )
```

### D11-07 - Excess Spread (Reported)
*Collections. Excess spread as reported in the servicer file.*  Format: `#,##0`

```dax
Excess Spread (Reported) =
SUM ( FactDynamicLoss[ExcessSpread_Monthly] )
```

### D11-08 - Excess Spread % (Reported, annualised)
*Collections. Annualised reported excess spread / opening balance.*  Format: `0.00%`

```dax
Excess Spread % (Reported, annualised) =
DIVIDE ( SUM ( FactDynamicLoss[ExcessSpread_Monthly] ), SUM ( FactDynamicLoss[BOP_Balance] ) ) * 12
```

### D11-09 - Report Month Label
*Report header. Header label for servicer-report pages.*  

```dax
Report Month Label =
FORMAT ( MAX ( FactDynamicLoss[MonthEnd] ), "mmm yyyy" )
```

### D11-10 - Data Tape Field Completeness %
*Data tape. Share of rating-agency fields (Brief A2.1) present or computable from the supplied data.*  Format: `0.0%`

```dax
Data Tape Field Completeness % =
DIVIDE ( COUNTROWS ( FILTER ( DataFieldMapping, DataFieldMapping[Status] IN { "Present", "Computed" } ) ), COUNTROWS ( DataFieldMapping ) )
```

## 12. Day 12 - Security, governance and validation

### D12-01 - Current User
*Security. Signed-in user (for RLS testing).*  

```dax
Current User =
USERPRINCIPALNAME ()
```

### D12-02 - Current User Role
*Security. Role of the signed-in user per DimSecurity.*  

```dax
Current User Role =
LOOKUPVALUE ( DimSecurity[Role], DimSecurity[UserEmail], USERPRINCIPALNAME (), "Unmapped" )
```

### D12-03 - Loans Visible to User
*Security. Row-count check under 'View as role' (must differ by role).*  Format: `#,##0`

```dax
Loans Visible to User =
[Loan Count]
```

### D12-04 - Reconciliation Break (Tape vs Dynamic)
*Validation. Tape balance minus the dynamic file's closing balance (reconciliation break, DQ10).*  Format: `#,##0`

```dax
Reconciliation Break (Tape vs Dynamic) =
VAR LastSeq = CALCULATE ( MAX ( FactDynamicLoss[ReportingSeq] ), REMOVEFILTERS () )
RETURN CALCULATE ( [Total Outstanding Balance], REMOVEFILTERS () )
     - CALCULATE ( SUM ( FactDynamicLoss[EOP_Balance] ), REMOVEFILTERS (), FactDynamicLoss[ReportingSeq] = LastSeq )
```

### D12-05 - High-Severity Data Issues
*Validation. Number of high-severity data-quality findings.*  Format: `0`

```dax
High-Severity Data Issues =
CALCULATE ( COUNTROWS ( DataQualityLog ), DataQualityLog[severity] = "High" )
```

### D12-06 - ECL Check Flag
*Validation. Pass/fail flag for the PD x LGD x EAD recomputation.*  

```dax
ECL Check Flag =
IF ( ABS ( [ECL Recompute Variance] ) < 1, "PASS", "FAIL" )
```

## Notes and assumptions

- **Trigger levels** (12% 60+ DPD, 2% CNL, 1.05 OC), tranche structure, coupons and the 0.5% servicing fee are **assumptions** - the brief supplies none.
- **Duplicate 30-May snapshot**: roll-rate and PD measures use all 12 sequential snapshots; date-based trend measures use IncludeInTrend = 1.
- **CNL Rate @ MOB / Cumulative Default Rate @ MOB** are cumulative series: always put MonthsOnBook on the axis (or filter to one MOB) to avoid summing cumulative values across MOBs.
- **Challenger vs tape ECL** differ materially because PD in the tape (about 2% Stage 1) is far below the roll-rate-implied PD; this is reported as a finding, not corrected.
- Measures could not be run here (no Power BI engine). Use `ExpectedMeasureValues.csv` / the dictionary's *Expected value* column and DAX Studio to confirm each headline figure.