"""
Securitisation Risk Analytics - Project 1A  |  Data pipeline (REAL ZeTheta datasets)
------------------------------------------------------------------------------------
Input  (./input_data/, the 4 CSVs downloaded from the portal):
    auto_loan_securitisation_data.csv   500 loans x 58 fields (tape at cutoff 2024-10-31)
    dpd_snapshot_history.csv            6,000 loan-month DPD snapshots
    dynamic_loss_monthly.csv            pool-level monthly loss / prepayment / collections
    static_pool_vintage_data.csv        15 quarterly vintages x months-on-book

Output (./powerbi_data/): star-schema CSVs for Power BI + validation + data-quality log.

    python securitisation_pipeline.py [--input input_data] [--out powerbi_data]
"""
import argparse
from pathlib import Path

import numpy as np
import pandas as pd

ap = argparse.ArgumentParser()
ap.add_argument("--input", default="input_data")
ap.add_argument("--out", default="powerbi_data")
ARGS, _ = ap.parse_known_args()
INPUT, OUT = Path(ARGS.input), Path(ARGS.out)
OUT.mkdir(exist_ok=True)

HORIZON = 36
SERVICING_FEE = 0.005
BUCKETS = ["Current", "1-29 DPD", "30-59 DPD", "60-89 DPD", "90-119 DPD", "120+ DPD", "Default", "Repossessed"]
BORD = {b: i for i, b in enumerate(BUCKETS)}
PD_STATES = ["Current", "1-29 DPD", "30-59 DPD", "60-89 DPD", "D"]      # D = 90+ / default (absorbing)

TRANCHES = pd.DataFrame(
    [("Class A (Senior)", 1, 0.20, 1.00, 0.070, "AAA"),
     ("Class B (Mezzanine)", 2, 0.08, 0.20, 0.085, "A"),
     ("Class C (Junior)", 3, 0.03, 0.08, 0.115, "BB"),
     ("Equity / Residual", 4, 0.00, 0.03, 0.000, "NR")],
    columns=["tranche", "seniority", "attachment", "detachment", "coupon", "rating"])
TRANCHES["size_pct"] = TRANCHES["detachment"] - TRANCHES["attachment"]
ORDER = TRANCHES.sort_values("seniority")["tranche"].tolist()
COUPON = dict(zip(TRANCHES["tranche"], TRANCHES["coupon"]))


def read(name):
    f = next((p for p in INPUT.glob("*.csv") if p.stem.lower() == name.lower()), None)
    if f is None:
        raise SystemExit(f"Missing {name}.csv in {INPUT.resolve()}")
    return pd.read_csv(f)


def eom(s):
    return pd.to_datetime(s) + pd.offsets.MonthEnd(0)


DQ = []
def dq(dataset, issue, rows, severity, treatment):
    DQ.append({"id": f"DQ{len(DQ) + 1:02d}", "dataset": dataset, "issue": issue, "rows_affected": rows,
               "severity": severity, "treatment_in_model": treatment})


# =============================================================================
# 1. Load + clean
# =============================================================================
raw_L, raw_D = read("auto_loan_securitisation_data"), read("dpd_snapshot_history")
raw_Y, raw_S = read("dynamic_loss_monthly"), read("static_pool_vintage_data")

L = raw_L.copy()
for c in ["OriginationDate", "CutoffDate", "MaturityDate", "LastPaymentDate"]:
    L[c] = pd.to_datetime(L[c])
for c in ["IsDefaulted", "IsModified", "HasInsurance", "IsNewVehicle"]:
    L[c] = L[c].astype(int)
L["ModificationType"] = L["ModificationType"].fillna("None")
L["VintageID"] = L["OriginationDate"].dt.year.astype(str) + "-Q" + L["OriginationDate"].dt.quarter.astype(str)
L["LTVBand"] = pd.cut(L["LTV_Current"], [0, .5, .7, .8, .9, 1.0, 9], labels=["<50%", "50-70%", "70-80%", "80-90%", "90-100%", ">100%"]).astype(str)
L["CIBILBand"] = pd.cut(L["CIBIL_Score_Current"], [0, 650, 700, 750, 900], labels=["<650", "650-699", "700-749", "750+"]).astype(str)
L["BucketOrder"] = L["DelinquencyStatus"].map(BORD)
L["ReportingDate"] = L["CutoffDate"]

# ---- data quality findings on the loan tape
n0 = int((L.CurrentBalance == 0).sum())
dq("auto_loan_securitisation_data", "Loans with CurrentBalance = 0 still carry a DPD status / IFRS 9 stage / PD & LGD (EAD = 0, so ECL = 0)", n0, "Medium",
   "Kept in tape; 'Active Loans' measure filters CurrentBalance > 0")
n_eadgt = int((L.EAD > L.CurrentBalance).sum())
dq("auto_loan_securitisation_data", "EAD exceeds CurrentBalance for most loans (EAD appears to include accrued interest / buffer; no field explains it)", n_eadgt, "Low",
   "ECL uses EAD as supplied; balance-based metrics use CurrentBalance")
tape_rule = L.groupby("IFRS9_Stage").DelinquencyDays.agg(["min", "max"]).to_dict("index")
dq("auto_loan_securitisation_data", "Tape staging (S2 = 30-89 DPD, S3 = 90+/default) differs from the brief's Day 9 rule (S1 = current/30 DPD, S2 = 60-90 DPD, S3 = 120+/default)",
   int(len(L)), "Medium", "Tape stage is the primary basis; DAX 'ECL (Alt Staging)' re-stages with adjustable DPD thresholds so the impact can be shown")
dq("auto_loan_securitisation_data", "7 'Default' loans carry no LossAmount/RecoveryAmount; only 2 repossessed loans have realised losses", 7, "Low",
   "Realised-loss measures from the tape are thin; LGD calibration uses dynamic_loss_monthly instead")

# =============================================================================
# 2. DPD snapshots
# =============================================================================
D = raw_D.copy()
D["SnapshotDate"] = pd.to_datetime(D["SnapshotDate"])
D["SnapshotSeq"] = D.groupby("LoanID").cumcount() + 1
D["MonthEnd"] = eom(D["SnapshotDate"])
D["DupDate"] = D.groupby("LoanID")["SnapshotDate"].diff().dt.days.eq(0)
D["IncludeInTrend"] = (~D["DupDate"]).astype(int)
D = D.merge(L[["LoanID", "OriginationDate"]], on="LoanID")
D["PreOrigination"] = (D["SnapshotDate"] < D["OriginationDate"]).astype(int)
D["BucketOrder"] = D["DPD_Bucket"].map(BORD)

tr = D.sort_values(["LoanID", "SnapshotSeq"]).copy()
tr["PriorBucketSeq"] = tr.groupby("LoanID")["DPD_Bucket"].shift(1)
tr["PriorPre"] = tr.groupby("LoanID")["PreOrigination"].shift(1)
tr["ValidTransition"] = (tr["PriorBucketSeq"].notna() & (tr["PreOrigination"] == 0) & (tr["PriorPre"] == 0)).astype(int)
D = D.merge(tr[["LoanID", "SnapshotSeq", "PriorBucketSeq", "ValidTransition"]], on=["LoanID", "SnapshotSeq"], how="left")
D["PriorBucketSeq"] = D["PriorBucketSeq"].fillna("None")
D["ValidTransition"] = D["ValidTransition"].fillna(0).astype(int)
D["PriorBucketOrder"] = D["PriorBucketSeq"].map(BORD).fillna(-1).astype(int)
D["SnapshotLabel"] = "S" + D["SnapshotSeq"].astype(str).str.zfill(2) + " " + D["SnapshotDate"].dt.strftime("%d-%b-%y")
D = D.drop(columns=["OriginationDate", "DupDate"]).sort_values(["LoanID", "SnapshotSeq"]).reset_index(drop=True)

dq("dpd_snapshot_history", "12 snapshots per loan but only 11 distinct dates: 2024-05-30 appears twice (Dec-23 to Oct-24). The supplied DPD_Bucket_Prior chains through both rows, so they are consecutive observations with one wrong date label", int((D.IncludeInTrend == 0).sum()), "High",
   "All 12 kept as sequential snapshots (SnapshotSeq) for roll rates / PD; the second 30-May row is flagged IncludeInTrend = 0 so date-based trend measures do not double-count May. Confirm the correct month with ZeTheta")
npre = int(D.PreOrigination.sum())
dq("dpd_snapshot_history", "Snapshots dated before the loan's OriginationDate (balance = original amount, DPD 0)", npre, "High",
   f"{L.shape[0] and int(D[D.PreOrigination == 1].LoanID.nunique())} loans affected; flagged PreOrigination = 1 and excluded from transitions")
m = tr[tr.PriorBucketSeq.notna() & (tr.PreOrigination == 0)]
mismatch = int((m["DPD_Bucket_Prior"] != m["PriorBucketSeq"]).sum())
dq("dpd_snapshot_history", "Supplied DPD_Bucket_Prior differs from the previous snapshot's bucket (post-origination rows)", mismatch, "Low",
   "Transitions rebuilt from consecutive snapshots (PriorBucketSeq); supplied column kept for reference")

# ---- empirical transition matrix (5 states, 90+ absorbing) and PD term structure
V = D[D.ValidTransition == 1].copy()
st = lambda b: b if BORD[b] < 4 else "D"
V["f"], V["t"] = V["PriorBucketSeq"].map(st), V["DPD_Bucket"].map(st)
cnt = pd.crosstab(V["f"], V["t"]).reindex(index=PD_STATES[:4], columns=PD_STATES, fill_value=0)
P = np.zeros((5, 5))
P[:4] = (cnt.div(cnt.sum(axis=1), axis=0)).to_numpy()
P[4, 4] = 1.0
pdts = []
Ph = np.eye(5)
for h in range(1, 61):
    Ph = Ph @ P
    for i, s in enumerate(PD_STATES[:4]):
        pdts.append({"DPD_Bucket": s, "HorizonMonths": h, "CumulativePD": round(Ph[i, 4], 6)})
PDTS = pd.DataFrame(pdts)

roll_rows = []
full = pd.crosstab(V["PriorBucketSeq"], V["DPD_Bucket"]).reindex(index=BUCKETS, columns=BUCKETS, fill_value=0)
for fb in BUCKETS:
    for tb in BUCKETS:
        roll_rows.append({"FromBucket": fb, "ToBucket": tb, "FromOrder": BORD[fb], "ToOrder": BORD[tb], "LoanMonths": int(full.loc[fb, tb])})
ROLL = pd.DataFrame(roll_rows)

# =============================================================================
# 3. Dynamic loss + static pool
# =============================================================================
Y = raw_Y.copy()
Y["ReportingDate"] = pd.to_datetime(Y["ReportingDate"])
Y["ReportingSeq"] = np.arange(1, len(Y) + 1)
Y["MonthEnd"] = eom(Y["ReportingDate"])
Y["IncludeInTrend"] = (~Y["ReportingDate"].duplicated(keep="first")).astype(int)
Y["NetLossSeverity"] = (Y["NetLoss_ThisMonth"] / Y["NewDefaults_Balance"]).round(5)
Y["BalanceContinuity_Gap"] = (Y["BOP_Balance"] - Y["EOP_Balance"].shift(1)).round(2)
dq("dynamic_loss_monthly", "ReportingDate 2024-05-30 appears twice (12 rows, 11 distinct months)", 1, "High",
   "All 12 rows used for rate calibration; second occurrence flagged IncludeInTrend = 0 for date-based trend measures (matches the DPD duplicate)")
gap = Y["BalanceContinuity_Gap"].abs().iloc[1:]
dq("dynamic_loss_monthly", "BOP_Balance of each month does not equal prior month EOP_Balance (average gap shown in BalanceContinuity_Gap)", int((gap > 1).sum()), "Medium",
   "Used as supplied for rates; roll-forward within a month (BOP - defaults - prepay - amort = EOP) reconciles and is tested")
tape_bal, dyn_eop = L.CurrentBalance.sum(), Y["EOP_Balance"].iloc[-1]
dq("Cross-file", f"Loan tape CurrentBalance ({tape_bal / 1e6:,.1f}M) does not reconcile to dynamic_loss_monthly closing EOP_Balance ({dyn_eop / 1e6:,.1f}M)", 1, "High",
   "Reported as a reconciliation break; tape used for pool metrics, dynamic file used for rates only")

S = raw_S.copy()
S["VintageStartDate"] = pd.to_datetime(S["VintageStartDate"])
S["CumDefaultRate"] = (S["CumulativeDefaults_Balance"] / S["OriginalPoolBalance"]).round(6)
S["CumGrossLossRate"] = (S["CumulativeGrossLoss"] / S["OriginalPoolBalance"]).round(6)
orig_static = S.drop_duplicates("VintageID").OriginalPoolBalance.sum()
dq("Cross-file", "static_pool OriginalPoolBalance (sum of vintages) vs tape OriginalLoanAmount", 0,
   "Info", f"Reconciles exactly ({orig_static:,.0f})" if abs(orig_static - L.OriginalLoanAmount.sum()) < 1 else "DOES NOT reconcile")

# =============================================================================
# 4. Calibration, scenarios, challenger ECL
# =============================================================================
Yt = Y   # all 12 monthly rows are used for calibration
m_def = Yt.NewDefaults_Balance.sum() / Yt.BOP_Balance.sum()
base_cdr = 1 - (1 - m_def) ** 12
base_lgd = Yt.NetLoss_ThisMonth.sum() / Yt.NewDefaults_Balance.sum()
gross_lgd = Yt.GrossLoss_ThisMonth.sum() / Yt.NewDefaults_Balance.sum()
smm = Yt.Prepayments_ThisMonth.sum() / Yt.BOP_Balance.sum()
base_cpr = 1 - (1 - smm) ** 12
act = L[L.CurrentBalance > 0]
pool0 = L.CurrentBalance.sum()
wac = np.average(L.InterestRate, weights=L.CurrentBalance) / 100
wam = np.average(L.RemainingTerm, weights=L.CurrentBalance)

PARAMS = pd.DataFrame([
    ("Base annual CDR", base_cdr, "dynamic_loss_monthly: NewDefaults_Balance / BOP_Balance, annualised"),
    ("Pool net LGD (severity)", base_lgd, "dynamic_loss_monthly: NetLoss / NewDefaults_Balance"),
    ("Pool gross LGD", gross_lgd, "dynamic_loss_monthly: GrossLoss / NewDefaults_Balance"),
    ("Base annual CPR", base_cpr, "dynamic_loss_monthly: Prepayments / BOP_Balance, annualised"),
    ("Pool WAC", wac, "tape: InterestRate weighted by CurrentBalance"),
    ("Pool WAM (months)", wam, "tape: RemainingTerm weighted by CurrentBalance"),
    ("Pool balance", pool0, "tape: sum CurrentBalance"),
    ("Rate shock PD elasticity (per 100bps)", 0.10, "ASSUMPTION - PD multiplier rises 10% per 100bps shock"),
    ("Delinquency trigger (60+ DPD % of pool)", 0.12, "ASSUMPTION - brief gives no trigger levels"),
    ("Cumulative net loss trigger", 0.02, "ASSUMPTION - brief gives no trigger levels"),
    ("OC ratio trigger (min)", 1.05, "ASSUMPTION - brief gives no trigger levels"),
    ("Arena weight - Equity", 10, "ASSUMPTION - Risk Arena points lost per 100% equity write-down"),
    ("Arena weight - Junior", 25, "ASSUMPTION - points lost per 100% Class C write-down"),
    ("Arena weight - Mezzanine", 35, "ASSUMPTION - points lost per 100% Class B write-down"),
    ("Arena weight - Senior", 30, "ASSUMPTION - points lost per 100% Class A write-down"),
    ("Arena weight - Trigger breach", 5, "ASSUMPTION - points lost per trigger breached"),
    ("Arena rating - Fortress (min score)", 95, "ASSUMPTION"),
    ("Arena rating - Resilient (min score)", 80, "ASSUMPTION"),
    ("Arena rating - Stressed (min score)", 60, "ASSUMPTION"),
], columns=["Parameter", "Value", "Source"])

SC = pd.DataFrame([
    (1, "Base", 0.50, 1.0, 0.00, 1.0, 0.00, 0), (2, "Upside", 0.10, 0.8, -0.03, 0.6, 0.02, 0),
    (3, "Moderate Stress", 0.25, 1.5, 0.07, 2.0, -0.02, 0), (4, "Severe Stress", 0.15, 2.5, 0.15, 3.5, -0.04, 100),
    (5, "Crisis (GFC-style)", 0.00, 4.0, 0.25, 6.0, -0.06, 200), (6, "Rate Shock", 0.00, 1.6, 0.03, 2.0, -0.03, 300)],
    columns=["scenario_id", "scenario", "weight", "pd_multiplier", "lgd_add", "cdr_multiple", "cpr_change", "rate_shock_bps"])
SC["annual_cdr"] = (base_cdr * SC.cdr_multiple).clip(upper=0.6).round(5)
SC["annual_cpr"] = (base_cpr + SC.cpr_change).clip(lower=0.01).round(5)
SC["pool_lgd"] = (base_lgd + SC.lgd_add).clip(0.10, 0.95).round(4)
assert abs(SC.weight.sum() - 1) < 1e-9


def challenger_ecl(df):
    ptab = PDTS.set_index(["DPD_Bucket", "HorizonMonths"])["CumulativePD"]
    out = []
    for r in df.itertuples():
        if r.EAD <= 0:
            out.append(0.0); continue
        if r.IFRS9_Stage == 3:
            pdv, h = 1.0, 6
        else:
            h = 12 if r.IFRS9_Stage == 1 else int(min(max(r.RemainingTerm, 1), 60))
            pdv = ptab[(r.DelinquencyStatus, h)]
        dfac = (1 + r.InterestRate / 100) ** (-(h / 12) / 2)
        out.append(pdv * base_lgd * r.EAD * dfac)
    return np.array(out)


L["ECL_Challenger"] = challenger_ecl(L).round(2)

# =============================================================================
# 5. Waterfall (calibrated from the data above)
# =============================================================================
def run_waterfall(sc):
    mdr, smm_ = 1 - (1 - sc.annual_cdr) ** (1 / 12), 1 - (1 - sc.annual_cpr) ** (1 / 12)
    bal = {r.tranche: pool0 * r.size_pct for r in TRANCHES.itertuples()}
    pool, cum_loss = pool0, 0.0
    prow, trow = [], []
    for t in range(1, HORIZON + 1):
        beg, beg_bal = pool, dict(bal)
        default = beg * mdr
        perf = beg - default
        prepay = perf * smm_
        sched = (perf - prepay) / max(wam - t + 1, 1)
        int_coll = perf * wac / 12
        recov, loss = default * (1 - sc.pool_lgd), default * sc.pool_lgd
        fees = beg * SERVICING_FEE / 12
        avail = max(int_coll - fees, 0.0)
        idue, ipaid = {}, {}
        wd, prin = {k: 0.0 for k in ORDER}, {k: 0.0 for k in ORDER}
        for k in ORDER[:3]:
            due = bal[k] * COUPON[k] / 12
            pay = min(due, avail); avail -= pay
            idue[k], ipaid[k] = due, pay
        idue[ORDER[3]] = 0.0
        excess = avail
        cover = min(excess, loss)
        ipaid[ORDER[3]] = excess - cover
        rem = loss - cover
        for k in reversed(ORDER):
            w = min(bal[k], rem); wd[k], bal[k], rem = w, bal[k] - w, rem - w
        pav = sched + prepay + recov + cover
        for k in ORDER:
            p = min(bal[k], pav); prin[k], bal[k], pav = p, bal[k] - p, pav - p
        pool = perf - prepay - sched
        cum_loss += loss
        notes = sum(bal[k] for k in ORDER[:3])
        prow.append({"scenario_id": sc.scenario_id, "period": t, "beginning_pool": beg, "defaults": default,
                     "prepayments": prepay, "scheduled_principal": sched, "interest_collected": int_coll,
                     "servicing_fees": fees, "recoveries": recov, "net_loss": loss, "excess_spread": excess,
                     "excess_spread_used_for_losses": cover, "ending_pool": pool, "cumulative_loss": cum_loss,
                     "cumulative_loss_pct": cum_loss / pool0, "oc_ratio": pool / notes if notes > 0 else np.nan})
        for k in ORDER:
            trow.append({"scenario_id": sc.scenario_id, "period": t, "tranche": k, "beginning_balance": beg_bal[k],
                         "interest_due": idue[k], "interest_paid": ipaid[k], "principal_paid": prin[k],
                         "writedown": wd[k], "ending_balance": bal[k]})
    return prow, trow


pr, tw = [], []
for s in SC.itertuples():
    a, b = run_waterfall(s); pr += a; tw += b
FP, FW = pd.DataFrame(pr), pd.DataFrame(tw)

# =============================================================================
# 6. Dimensions
# =============================================================================
DimLoan = L[["LoanID", "PoolID", "BorrowerID", "ServicerID", "ServicerName", "OriginationChannel", "LoanPurpose",
             "OriginationDate", "MaturityDate", "OriginalLoanAmount", "OriginalTerm", "InterestRate", "VintageID",
             "State", "InsuranceType", "HasInsurance", "IsModified", "ModificationType"]]
tape_cols = ["LoanID", "BorrowerID", "State", "ReportingDate", "CurrentBalance", "RemainingTerm", "MonthsOnBook", "MonthlyEMI",
             "ScheduledPaymentDue", "LastPaymentAmount", "LastPaymentDate", "DelinquencyStatus", "DelinquencyDays", "BucketOrder",
             "TotalPaymentsDue", "TotalPaymentsMade", "Times30DPD_Last12M", "Times60DPD_Last12M", "Times90DPD_Last12M",
             "LTV_AtOrigination", "LTV_Current", "LTVBand", "DTI_Ratio", "CIBIL_Score_Origination", "CIBIL_Score_Current", "CIBILBand",
             "IFRS9_Stage", "PD_Estimate", "LGD_Estimate", "EAD", "ECL_Provision", "ECL_Challenger", "IsDefaulted",
             "LossAmount", "RecoveryAmount", "NetLoss", "PrepaymentAmount"]
FactLoanTape = L[tape_cols]
DimBorrower = L[["BorrowerID", "BorrowerAge", "EmploymentType", "AnnualIncome_INR"]].copy()
DimBorrower["AgeBand"] = pd.cut(DimBorrower.BorrowerAge, [0, 30, 40, 50, 60, 120], labels=["<30", "30-39", "40-49", "50-59", "60+"]).astype(str)
DimBorrower["IncomeBand"] = pd.cut(DimBorrower.AnnualIncome_INR, [0, 5e5, 1e6, 2e6, 4e6, 1e10], labels=["<5L", "5-10L", "10-20L", "20-40L", "40L+"]).astype(str)
DimVehicle = L[["LoanID", "VehicleMake", "VehicleModel", "VehicleYear", "VehicleType", "IsNewVehicle", "OriginalVehicleValue", "CurrentVehicleValue"]]
DimGeography = L[["State", "Region"]].drop_duplicates().sort_values("State")
dates = pd.date_range("2021-01-01", "2027-12-31", freq="D")
DimDate = pd.DataFrame({"Date": dates})
DimDate["Year"], DimDate["Quarter"] = dates.year, dates.quarter
DimDate["MonthNum"], DimDate["MonthName"] = dates.month, dates.strftime("%b")
DimDate["MonthYear"] = dates.strftime("%b-%y")
DimDate["YearMonthKey"] = dates.year * 100 + dates.month
DimDate["MonthEndDate"] = dates + pd.offsets.MonthEnd(0)
DimDate["IsMonthEnd"] = (dates == DimDate["MonthEndDate"]).astype(int)
DimDate["FiscalYear"] = np.where(dates.month >= 4, dates.year + 1, dates.year)       # India FY Apr-Mar
DimDate["FiscalQuarter"] = ((dates.month - 4) % 12 // 3 + 1)
sma = D.groupby("DPD_Bucket").RBI_SMA_Class.agg(lambda s: s.mode().iat[0]).to_dict()
DimBucket = pd.DataFrame({"DPD_Bucket": BUCKETS, "BucketOrder": range(8)})
DimBucket["IsDelinquent30Plus"] = (DimBucket.BucketOrder >= 2).astype(int)
DimBucket["IsNPA90Plus"] = (DimBucket.BucketOrder >= 4).astype(int)
DimBucket["RBI_SMA_Class"] = DimBucket.DPD_Bucket.map(sma)
DimStage = pd.DataFrame({"IFRS9_Stage": [1, 2, 3], "StageName": ["Stage 1 - Performing", "Stage 2 - Underperforming", "Stage 3 - Credit-impaired"]})
DimVintage = S.drop_duplicates("VintageID")[["VintageID", "VintageStartDate", "OriginalLoanCount", "OriginalPoolBalance"]].copy()
DimVintage["VintageYear"] = DimVintage.VintageStartDate.dt.year
DimPeriod = pd.DataFrame({"period": range(1, HORIZON + 1), "period_date": pd.date_range("2024-11-30", periods=HORIZON, freq="ME")})
DimSecurity = pd.DataFrame([
    ("portfolio.manager@example.com", "PortfolioManager", "ALL", "ALL"),
    ("risk.analyst@example.com", "RiskAnalyst", "ALL", "ALL"),
    ("senior.mgmt@example.com", "SeniorManagement", "ALL", "ALL"),
    ("auditor@example.com", "Auditor", "ALL", "North"),
    ("servicer.hdfc@example.com", "Servicer", "SVC001", "ALL")], columns=["UserEmail", "Role", "ServicerID", "Region"])

# DATA FIELD MAPPING (Day 1): brief Part A2.1 -> provided tape
FM = [
 ("Loan ID", "Identification", "LoanID", "Present"), ("Pool ID", "Identification", "PoolID", "Present"),
 ("Originator ID", "Identification", "-", "Missing (ServicerID only)"), ("Servicer ID", "Identification", "ServicerID", "Present"),
 ("Account Status", "Identification", "DelinquencyStatus", "Present"), ("Origination Date", "Origination", "OriginationDate", "Present"),
 ("Original Balance", "Origination", "OriginalLoanAmount", "Present"), ("Original Term", "Origination", "OriginalTerm", "Present"),
 ("Interest Rate", "Origination", "InterestRate", "Present"), ("Rate Type", "Origination", "-", "Missing"),
 ("Origination Channel", "Origination", "OriginationChannel", "Present"), ("Loan Purpose", "Origination", "LoanPurpose", "Present"),
 ("Vehicle Make / Model / Year / Type", "Collateral", "VehicleMake, VehicleModel, VehicleYear, VehicleType", "Present"),
 ("New/Used Flag", "Collateral", "IsNewVehicle", "Present"), ("Original / Current Vehicle Value", "Collateral", "OriginalVehicleValue, CurrentVehicleValue", "Present"),
 ("Mileage at Origination", "Collateral", "-", "Missing"), ("Vehicle Registration State", "Collateral", "State (borrower state)", "Proxy"),
 ("CIBIL Score at Origination / Current", "Borrower", "CIBIL_Score_Origination, CIBIL_Score_Current", "Present"),
 ("Employment Type", "Borrower", "EmploymentType", "Present"), ("Annual Income", "Borrower", "AnnualIncome_INR", "Present"),
 ("DTI Ratio", "Borrower", "DTI_Ratio", "Present"), ("Borrower Age", "Borrower", "BorrowerAge", "Present"),
 ("Geographic Region", "Borrower", "Region, State", "Present"), ("Years at Current Employment", "Borrower", "-", "Missing"),
 ("Reporting Date", "Performance", "CutoffDate", "Present"), ("Current Balance", "Performance", "CurrentBalance", "Present"),
 ("Scheduled Payment", "Performance", "ScheduledPaymentDue", "Present"), ("Actual Payment", "Performance", "LastPaymentAmount", "Present"),
 ("Delinquency Status / DPD", "Performance", "DelinquencyStatus, DelinquencyDays", "Present"),
 ("Cumulative Loss / Recovery", "Performance", "LossAmount, RecoveryAmount, NetLoss", "Present"),
 ("Prepayment Amount", "Performance", "PrepaymentAmount", "Present"), ("Modification Flag", "Performance", "IsModified, ModificationType", "Present"),
 ("IFRS 9 Stage", "Performance", "IFRS9_Stage", "Present"),
 ("RBI Priority Sector Classification", "India-specific", "-", "Missing"), ("GST Registration", "India-specific", "-", "Missing"),
 ("Aadhaar / KYC Status", "India-specific", "-", "Missing"), ("NACH Mandate Status", "India-specific", "-", "Missing"),
 ("Insurance Coverage", "India-specific", "InsuranceType, HasInsurance", "Present"), ("Hypothecation Registration", "India-specific", "-", "Missing"),
 ("Moratorium History", "India-specific", "ModificationType (partial)", "Partial"),
 ("WAC / WAM / WALA / WA LTV / WA CIBIL", "Computed in DAX", "see measure dictionary", "Computed"),
 ("HHI, Top-10 concentration, stratification", "Computed in DAX", "see measure dictionary", "Computed")]
FieldMap = pd.DataFrame(FM, columns=["RequiredField", "Group", "SourceColumns", "Status"])

# =============================================================================
# 7. Validation (Day 12/14)
# =============================================================================
checks = []
def chk(name, ok, detail="", info=False):
    checks.append({"check": name, "result": "INFO" if info else ("PASS" if ok else "FAIL"), "detail": detail})

chk("Loan IDs unique", L.LoanID.is_unique)
chk("Tape ECL_Provision = PD x LGD x EAD (all loans)", bool(np.allclose(L.PD_Estimate * L.LGD_Estimate * L.EAD, L.ECL_Provision, atol=0.02)))
chk("Tape stage consistent with DPD status (S1 <30 DPD, S2 30-89, S3 90+/default)",
    bool(((L.IFRS9_Stage == 1) & (L.DelinquencyDays < 30)).sum() + ((L.IFRS9_Stage == 2) & L.DelinquencyDays.between(30, 89)).sum()
         + ((L.IFRS9_Stage == 3) & ((L.DelinquencyDays >= 90) | (L.IsDefaulted == 1))).sum() == len(L)))
chk("Every DPD LoanID exists in tape", set(D.LoanID) <= set(L.LoanID))
chk("Latest DPD snapshot bucket = tape DelinquencyStatus", bool(
    (D[D.SnapshotSeq == D.SnapshotSeq.max()].set_index("LoanID").DPD_Bucket == L.set_index("LoanID").DelinquencyStatus).all()))
chk("Static pool vintages = tape vintages (IDs)", set(S.VintageID) == set(L.VintageID))
chk("Static pool loan counts = tape loans per vintage", bool((DimVintage.set_index("VintageID").OriginalLoanCount == L.groupby("VintageID").size()).all()))
chk("Static pool original balance = tape original amount", abs(orig_static - L.OriginalLoanAmount.sum()) < 1, f"{orig_static:,.0f}")
chk("Static pool CNL never decreases with months on book",
    bool((S.sort_values("MonthsOnBook").groupby("VintageID")["CumulativeNetLoss"].diff().dropna() >= -0.01).all()))
chk("Dynamic loss: Net = Gross - Recoveries", bool(np.allclose(Y.NetLoss_ThisMonth, Y.GrossLoss_ThisMonth - Y.Recoveries_ThisMonth, atol=0.02)))
chk("Dynamic loss: BOP - defaults - prepay - amort = EOP", bool(np.allclose(
    Y.BOP_Balance - Y.NewDefaults_Balance - Y.Prepayments_ThisMonth - Y.ScheduledAmort, Y.EOP_Balance, atol=1)))
chk("Transition matrix rows sum to 1", bool(np.allclose(P.sum(1), 1)))
chk("PD term structure monotone non-decreasing", bool((PDTS.sort_values("HorizonMonths").groupby("DPD_Bucket").CumulativePD.diff().dropna() >= -1e-12).all()))
chk("Weights sum to 1", abs(SC.weight.sum() - 1) < 1e-9)
init = FW[FW.period == 1].groupby("scenario_id").beginning_balance.sum()
chk("Waterfall: initial notes = pool balance", bool(np.allclose(init, pool0)))
en = FW[FW.period == HORIZON].groupby("scenario_id").ending_balance.sum().sort_index()
ep = FP[FP.period == HORIZON].set_index("scenario_id").ending_pool.sort_index()
chk("Waterfall: notes = pool at month 36 (cash conservation)", bool(np.allclose(en, ep, rtol=1e-6)))
chk("Waterfall: no negative tranche balances", bool((FW.ending_balance >= -1e-6).all()))
tot_t, tot_c = L.ECL_Provision.sum(), L.ECL_Challenger.sum()
chk("Tape ECL vs challenger (transition-matrix PD) ECL", True, f"tape {tot_t:,.0f} vs challenger {tot_c:,.0f} ({tot_c / tot_t - 1:+.1%})", info=True)
chk("Calibration (dynamic loss)", True, f"CDR {base_cdr:.2%}, net LGD {base_lgd:.1%}, CPR {base_cpr:.1%}", info=True)

# =============================================================================
# 8. Expected values for headline measures (for DAX cross-validation, Day 12)
# =============================================================================
w = L.CurrentBalance
def wavg(x): return float((x * w).sum() / w.sum())
EXP = {
 "Total Outstanding Balance": w.sum(), "Original Balance": L.OriginalLoanAmount.sum(), "Pool Factor": w.sum() / L.OriginalLoanAmount.sum(),
 "Total EAD": L.EAD.sum(), "Loan Count": len(L), "Active Loans": int((w > 0).sum()), "Defaulted Loans": int(L.IsDefaulted.sum()),
 "WAC": wavg(L.InterestRate) / 100, "WAM": wavg(L.RemainingTerm), "WALA": wavg(L.MonthsOnBook), "WALT": wavg(L.OriginalTerm),
 "WA LTV Current": wavg(L.LTV_Current), "WA CIBIL": wavg(L.CIBIL_Score_Current), "WA DTI": wavg(L.DTI_Ratio),
 "Delinquent 30+ %": w[L.DelinquencyDays >= 30].sum() / w.sum(), "Delinquent 60+ %": w[L.DelinquencyDays >= 60].sum() / w.sum(),
 "Delinquent 90+ %": w[L.DelinquencyDays >= 90].sum() / w.sum(),
 "Total ECL (Tape)": L.ECL_Provision.sum(), "ECL Coverage %": L.ECL_Provision.sum() / L.EAD.sum(),
 "ECL Stage 1": L[L.IFRS9_Stage == 1].ECL_Provision.sum(), "ECL Stage 2": L[L.IFRS9_Stage == 2].ECL_Provision.sum(), "ECL Stage 3": L[L.IFRS9_Stage == 3].ECL_Provision.sum(),
 "Challenger ECL (DAX)": L.ECL_Challenger.sum(), "Realised Recovery Rate (Tape)": L.RecoveryAmount.sum() / L.LossAmount.sum(),
 "Net Loss Severity (Dynamic)": base_lgd, "Annualised Default Rate (All Months)": 1 - (1 - m_def) ** 12,
 "HHI State": float(((L.groupby("State").CurrentBalance.sum() / w.sum()) ** 2).sum()),
 "Top 5 State Concentration": float(L.groupby("State").CurrentBalance.sum().nlargest(5).sum() / w.sum()),
 "Roll Rate Current to Current": float(full.loc["Current", "Current"] / full.loc["Current"].sum()),
 "Roll Rate Current to 1-29": float(full.loc["Current", "1-29 DPD"] / full.loc["Current"].sum()),
 "Roll Rate 60-89 to 90+": float(full.loc["60-89 DPD", BUCKETS[4:]].sum() / full.loc["60-89 DPD"].sum()),
 "Roll Rate 30-59 to 60-89": float(full.loc["30-59 DPD", "60-89 DPD"] / full.loc["30-59 DPD"].sum()),
 "Reconciliation Break (Tape vs Dynamic)": float(w.sum() - Y.EOP_Balance.iloc[-1]),
 "Cumulative Net Loss (Reporting)": float(Y[Y.IncludeInTrend == 1].NetLoss_ThisMonth.sum()),
}
EXPV = pd.DataFrame({"Measure": list(EXP), "ExpectedValue_Python": [round(float(v), 6) for v in EXP.values()]})

# =============================================================================
# 9. Write
# =============================================================================
def out(df, name, **kw):
    df.to_csv(OUT / f"{name}.csv", index=False, **kw)

out(DimLoan, "DimLoan"); out(FactLoanTape, "FactLoanTape"); out(DimBorrower, "DimBorrower"); out(DimVehicle, "DimVehicle")
out(DimGeography, "DimGeography"); out(DimDate, "DimDate"); out(DimBucket, "DimBucket"); out(DimStage, "DimStage")
out(DimVintage, "DimVintage"); out(D, "FactMonthlyPerformance"); out(S, "FactStaticPool"); out(Y, "FactDynamicLoss")
out(ROLL, "FactRollRate"); out(PDTS, "FactPDTermStructure"); out(PARAMS, "DimParameters"); out(SC, "DimScenario")
out(TRANCHES, "DimTranche"); out(DimPeriod, "DimPeriod"); out(FP.round(2), "FactPool"); out(FW.round(2), "FactWaterfall")
out(DimSecurity, "DimSecurity"); out(FieldMap, "DataFieldMapping"); out(pd.DataFrame(DQ), "DataQualityLog")
out(pd.DataFrame(checks), "ValidationReport"); out(EXPV, "ExpectedMeasureValues")

print(f"Loans {len(L)} | pool {pool0:,.0f} | active {int((w > 0).sum())} | WAC {wac:.2%} | WAM {wam:.1f}")
print(f"Calibration: CDR {base_cdr:.2%}  net LGD {base_lgd:.1%}  gross LGD {gross_lgd:.1%}  CPR {base_cpr:.1%}")
print(f"ECL tape {tot_t:,.0f} ({tot_t / L.EAD.sum():.2%} of EAD) | challenger {tot_c:,.0f} ({tot_c / tot_t - 1:+.1%})")
print("Stage ECL:", L.groupby("IFRS9_Stage").ECL_Provision.sum().round(0).to_dict())
print("PD12 by bucket:", PDTS[PDTS.HorizonMonths == 12].set_index("DPD_Bucket").CumulativePD.round(4).to_dict())
print("Transition counts:\n", cnt.to_string())
tl = FW.groupby(["scenario_id", "tranche"]).writedown.sum().unstack().div(pool0 * TRANCHES.set_index("tranche").size_pct, axis=1)
print("Writedown % of tranche (36m):\n", tl[ORDER].round(3).to_string())
print("\nValidation:\n", pd.DataFrame(checks).to_string(index=False))
print("\nData quality log:\n", pd.DataFrame(DQ)[["id", "dataset", "severity", "rows_affected"]].to_string(index=False))
