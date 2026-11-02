"""
AnalystLab Africa - Week 3 (Track A: Data Analytics)
Hospital Operations, Cost & Patient Flow Intelligence

Reproducible script covering:
  Task 1  - data profiling, cleaning rules, Power BI-ready data model tables
  Task 2  - all eight analysis questions + supplementary patient-behaviour questions

Usage:
    python week3_analysis.py <path-to-merged-workbook.xlsx> <output-folder>

Notes
  * No rows are deleted and no missing values are imputed. Derived columns are
    ADDED; raw values are never overwritten.
  * Patient names, addresses, coordinates etc. are NOT carried into the model
    tables (privacy by design). Patient IDs are shown shortened in outputs.
"""
import sys, os
import numpy as np
import pandas as pd

SRC = sys.argv[1] if len(sys.argv) > 1 else "Untitled_spreadsheet__1_.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "outputs"
RES = os.path.join(OUT, "results")
MOD = os.path.join(OUT, "model_tables")
os.makedirs(RES, exist_ok=True)
os.makedirs(MOD, exist_ok=True)


def save(df, name, index=False):
    df.to_csv(os.path.join(RES, name + ".csv"), index=index)
    return df


# ---------------------------------------------------------------- 1. LOAD
raw = pd.read_excel(SRC, sheet_name=None)
raw = {k.strip().lower(): v for k, v in raw.items()}          # sheet names have trailing spaces
for k in raw:
    raw[k].columns = [str(c).strip() for c in raw[k].columns]  # normalise headers
enc, proc = raw["encounters"].copy(), raw["procedures"].copy()
pat, pay, org = raw["patients"].copy(), raw["payers"].copy(), raw["organizations"].copy()

# ---------------------------------------------------------------- 2. DERIVED FIELDS (additive only)
for df in (enc, proc):
    df["START_TS"] = pd.to_datetime(df["START"], utc=True)
    df["STOP_TS"] = pd.to_datetime(df["STOP"], utc=True)

enc["DURATION_HOURS"] = (enc.STOP_TS - enc.START_TS).dt.total_seconds() / 3600
enc["DURATION_BAND"] = np.select(
    [enc.DURATION_HOURS < 24, enc.DURATION_HOURS == 24],
    ["Under 24h", "Exactly 24h"], default="Over 24h")
enc["START_YEAR"] = enc.START_TS.dt.year
enc["START_QUARTER"] = enc.START_TS.dt.year.astype(str) + "-Q" + enc.START_TS.dt.quarter.astype(str)
enc["UNCOVERED_AMOUNT"] = (enc.TOTAL_CLAIM_COST - enc.PAYER_COVERAGE).round(2)
enc["ZERO_COVERAGE_FLAG"] = (enc.PAYER_COVERAGE == 0).astype(int)
enc["ZERO_COST_FLAG"] = (enc.TOTAL_CLAIM_COST == 0).astype(int)
enc = enc.merge(pay[["Id", "NAME"]].rename(columns={"Id": "PAYER", "NAME": "PAYER_NAME"}), on="PAYER", how="left")

# ---------------------------------------------------------------- 3. DATA-QUALITY LOG
dq = []
def log(table, check, result, note=""):
    dq.append({"table": table, "check": check, "result": result, "note": note})

for name, df in raw.items():
    log(name, "rows x columns", f"{df.shape[0]:,} x {df.shape[1]}")
    log(name, "fully duplicated rows", int(df.duplicated().sum()))
    miss = df.isna().sum()
    log(name, "columns with missing values", ", ".join(f"{c}={n:,}" for c, n in miss[miss > 0].items()) or "none")

log("encounters", "duplicate Id", int(enc.Id.duplicated().sum()))
log("encounters", "unparseable START/STOP", int(enc.START_TS.isna().sum() + enc.STOP_TS.isna().sum()))
log("encounters", "STOP earlier than START", int((enc.STOP_TS < enc.START_TS).sum()))
log("encounters", "date range", f"{enc.START_TS.min():%Y-%m-%d} to {enc.START_TS.max():%Y-%m-%d}")
log("encounters", "orphan PATIENT / ORGANIZATION / PAYER keys",
    f"{(~enc.PATIENT.isin(pat.Id)).sum()} / {(~enc.ORGANIZATION.isin(org.Id)).sum()} / {(~enc.PAYER.isin(pay.Id)).sum()}")
log("encounters", "negative cost or coverage values",
    int((enc[["BASE_ENCOUNTER_COST", "TOTAL_CLAIM_COST", "PAYER_COVERAGE"]] < 0).sum().sum()))
log("encounters", "PAYER_COVERAGE > TOTAL_CLAIM_COST", int((enc.PAYER_COVERAGE > enc.TOTAL_CLAIM_COST).sum()))
log("encounters", "TOTAL_CLAIM_COST = 0", int(enc.ZERO_COST_FLAG.sum()))
log("encounters", "TOTAL_CLAIM_COST < BASE_ENCOUNTER_COST", int((enc.TOTAL_CLAIM_COST < enc.BASE_ENCOUNTER_COST).sum()))
log("encounters", "same CODE with >1 DESCRIPTION", int((enc.groupby("CODE").DESCRIPTION.nunique() > 1).sum()))
log("encounters", "duplicate PATIENT+START", int(enc.duplicated(["PATIENT", "START"]).sum()),
    "kept: multiple concurrent records for one patient are plausible")
log("encounters", "encounters with duration > 365 days", int((enc.DURATION_HOURS > 24 * 365).sum()))

log("procedures", "orphan ENCOUNTER / PATIENT keys",
    f"{(~proc.ENCOUNTER.isin(enc.Id)).sum()} / {(~proc.PATIENT.isin(pat.Id)).sum()}")
log("procedures", "BASE_COST exactly 431",
    f"{(proc.BASE_COST == 431).sum():,} ({(proc.BASE_COST == 431).mean():.1%})", "looks like a default/flat price")
log("procedures", "same CODE with >1 DESCRIPTION", int((proc.groupby("CODE").DESCRIPTION.nunique() > 1).sum()))
pe = proc.merge(enc[["Id", "START_TS", "STOP_TS"]], left_on="ENCOUNTER", right_on="Id", suffixes=("", "_E"))
log("procedures", "procedure START before encounter START", int((pe.START_TS < pe.START_TS_E).sum()))
log("procedures", "procedure START after encounter STOP", int((pe.START_TS > pe.STOP_TS_E).sum()))

sp = proc.groupby("ENCOUNTER").BASE_COST.agg(psum="sum", pn="size")
chk = enc.join(sp, on="Id")
chk[["psum", "pn"]] = chk[["psum", "pn"]].fillna(0)
short = chk.BASE_ENCOUNTER_COST + chk.psum - chk.TOTAL_CLAIM_COST
tol = 0.5 * chk.pn + 0.01                       # procedure costs are whole numbers in the source
log("encounters/procedures", "TOTAL_CLAIM_COST < base + sum(procedure costs), beyond rounding tolerance",
    f"{(short > tol).sum():,}", "claim totals cannot always be rebuilt from the supplied tables")
log("encounters/procedures", "encounters with at least one procedure", f"{enc.Id.isin(proc.ENCOUNTER).sum():,}")

pat["DEATH_TS"] = pd.to_datetime(pat.DEATHDATE)
x = enc.merge(pat[["Id", "DEATH_TS"]], left_on="PATIENT", right_on="Id", suffixes=("", "_P"))
late = x[x.START_TS.dt.tz_localize(None) > x.DEATH_TS + pd.Timedelta(days=1)]
log("patients", "encounters starting after death date", f"{len(late)} (wellness={int((late.ENCOUNTERCLASS=='wellness').sum())})",
    "mostly death-certification records dated within 14 days of death")
log("patients", "patients without any encounter", int((~pat.Id.isin(enc.PATIENT)).sum()))
log("patients", "ZIP missing / stored as number (leading zeros lost)", int(pat.ZIP.isna().sum()))
log("organizations", "distinct organizations", int(org.Id.nunique()), "single-hospital dataset: no cross-hospital comparison possible")
save(pd.DataFrame(dq), "00_data_quality_log")

# ---------------------------------------------------------------- 4. MODEL TABLES (Power BI ready)
dim_patient = pat[["Id", "BIRTHDATE", "DEATHDATE", "GENDER", "RACE", "ETHNICITY", "MARITAL", "CITY", "COUNTY", "STATE"]].copy()
dim_patient["ZIP"] = pat.ZIP.apply(lambda z: f"{int(z):05d}" if pd.notna(z) else None)
dim_patient["IS_DECEASED"] = pat.DEATHDATE.notna().astype(int)
dim_payer = pay[["Id", "NAME", "STATE_HEADQUARTERED"]].copy()
dim_org = org[["Id", "NAME", "CITY", "STATE"]].copy()
rng = pd.date_range(enc.START_TS.min().floor("D").tz_localize(None), enc.START_TS.max().ceil("D").tz_localize(None))
dim_date = pd.DataFrame({"DATE": rng})
dim_date["YEAR"] = dim_date.DATE.dt.year
dim_date["QUARTER"] = "Q" + dim_date.DATE.dt.quarter.astype(str)
dim_date["YEAR_QUARTER"] = dim_date.YEAR.astype(str) + "-Q" + dim_date.DATE.dt.quarter.astype(str)
dim_date["MONTH_NUM"] = dim_date.DATE.dt.month
dim_date["MONTH_NAME"] = dim_date.DATE.dt.strftime("%b")
fact_enc = enc[["Id", "PATIENT", "ORGANIZATION", "PAYER", "ENCOUNTERCLASS", "CODE", "DESCRIPTION", "START_TS", "STOP_TS",
                "DURATION_HOURS", "DURATION_BAND", "START_YEAR", "START_QUARTER", "BASE_ENCOUNTER_COST",
                "TOTAL_CLAIM_COST", "PAYER_COVERAGE", "UNCOVERED_AMOUNT", "ZERO_COVERAGE_FLAG", "ZERO_COST_FLAG",
                "REASONCODE", "REASONDESCRIPTION"]].copy()
fact_enc["START_DATE"] = fact_enc.START_TS.dt.tz_localize(None).dt.normalize()
fact_enc["START_TS"] = fact_enc.START_TS.dt.tz_localize(None)
fact_enc["STOP_TS"] = fact_enc.STOP_TS.dt.tz_localize(None)
fact_proc = proc[["ENCOUNTER", "PATIENT", "CODE", "DESCRIPTION", "START_TS", "STOP_TS", "BASE_COST",
                  "REASONCODE", "REASONDESCRIPTION"]].copy()
fact_proc["START_TS"] = fact_proc.START_TS.dt.tz_localize(None)
fact_proc["STOP_TS"] = fact_proc.STOP_TS.dt.tz_localize(None)
for n, d in {"dim_patient": dim_patient, "dim_payer": dim_payer, "dim_organization": dim_org, "dim_date": dim_date,
             "fact_encounters": fact_enc, "fact_procedures": fact_proc}.items():
    d.to_csv(os.path.join(MOD, n + ".csv"), index=False)

# ---------------------------------------------------------------- 5. ANALYSIS
# Q1 / Obj1a - volumes by year
yr = enc.groupby("START_YEAR").size().rename("encounters").reset_index()
yr["yoy_change_pct"] = (yr.encounters.pct_change() * 100).round(1)
last = enc.START_TS.max()
yr["note"] = ""
yr.loc[yr.START_YEAR == last.year, "note"] = f"PARTIAL YEAR (data ends {last:%d %b %Y})"
save(yr, "q01_encounters_by_year")
qt = enc.groupby("START_QUARTER").size().rename("encounters").reset_index()
save(qt, "q01b_encounters_by_quarter")

# Q2 / Obj1b - class mix
cls = enc.ENCOUNTERCLASS.value_counts().rename("encounters").reset_index()
cls["pct"] = (cls.encounters / cls.encounters.sum() * 100).round(2)
save(cls, "q02_class_share_overall")
mix = pd.crosstab(enc.START_YEAR, enc.ENCOUNTERCLASS)
mixpct = (mix.div(mix.sum(axis=1), axis=0) * 100).round(1)
save(mix.reset_index(), "q02b_class_counts_by_year")
save(mixpct.reset_index(), "q02c_class_pct_by_year")

# Q3 / Obj1c - duration
band = enc.DURATION_BAND.value_counts().rename("encounters").reset_index()
band["pct"] = (band.encounters / len(enc) * 100).round(2)
save(band, "q03_duration_bands")
save(pd.crosstab(enc.ENCOUNTERCLASS, enc.DURATION_BAND).reset_index(), "q03b_duration_by_class")
over = enc[enc.DURATION_HOURS > 24].DURATION_HOURS
save(pd.DataFrame({"stat": ["n", "median_h", "p90_h", "max_h", "n_over_365d"],
                   "value": [len(over), over.median(), over.quantile(.9), over.max(), int((over > 8760).sum())]}),
     "q03c_over24h_profile")

# Q4 - claim cost by class
c4 = enc.groupby("ENCOUNTERCLASS").TOTAL_CLAIM_COST.agg(encounters="size", total="sum", mean="mean", median="median",
                                                        p90=lambda s: s.quantile(.9), max="max").round(2).reset_index()
c4["share_of_total_cost_pct"] = (c4.total / c4.total.sum() * 100).round(1)
save(c4.sort_values("total", ascending=False), "q04_claim_cost_by_class")

# Q5 / Obj2a - coverage
tot_claim, tot_cov = enc.TOTAL_CLAIM_COST.sum(), enc.PAYER_COVERAGE.sum()
z = enc[enc.ZERO_COVERAGE_FLAG == 1]
cov = pd.DataFrame({
    "metric": ["total_claim_cost", "total_payer_coverage", "coverage_ratio_pct", "uncovered_amount",
               "encounters_zero_coverage", "pct_encounters_zero_coverage",
               "  of which NO_INSURANCE", "  of which insured payer, zero coverage", "  of which zero-cost encounters",
               "zero-coverage with claim cost > 0"],
    "value": [round(tot_claim, 2), round(tot_cov, 2), round(tot_cov / tot_claim * 100, 2), round(tot_claim - tot_cov, 2),
              len(z), round(len(z) / len(enc) * 100, 2),
              int((z.PAYER_NAME == "NO_INSURANCE").sum()), int((z.PAYER_NAME != "NO_INSURANCE").sum()),
              int((z.TOTAL_CLAIM_COST == 0).sum()), int((z.TOTAL_CLAIM_COST > 0).sum())]})
save(cov, "q05_coverage_summary")
c5 = enc.groupby("PAYER_NAME").agg(encounters=("Id", "size"), zero_cov=("ZERO_COVERAGE_FLAG", "sum"),
                                   claim=("TOTAL_CLAIM_COST", "sum"), covered=("PAYER_COVERAGE", "sum")).reset_index()
c5["zero_cov_pct"] = (c5.zero_cov / c5.encounters * 100).round(1)
c5["coverage_ratio_pct"] = (c5.covered / c5.claim * 100).round(1)
save(c5.sort_values("encounters", ascending=False), "q05b_coverage_by_payer")
c5c = enc.groupby("ENCOUNTERCLASS").agg(encounters=("Id", "size"), zero_cov=("ZERO_COVERAGE_FLAG", "sum"),
                                        claim=("TOTAL_CLAIM_COST", "sum"), covered=("PAYER_COVERAGE", "sum")).reset_index()
c5c["zero_cov_pct"] = (c5c.zero_cov / c5c.encounters * 100).round(1)
c5c["coverage_ratio_pct"] = (c5c.covered / c5c.claim * 100).round(1)
save(c5c, "q05c_coverage_by_class")

# Q6 / Q7 / Obj2b-c - procedures (grouped by SNOMED CODE; label = most frequent description)
label = proc.groupby("CODE").DESCRIPTION.agg(lambda s: s.value_counts().index[0])
pg = proc.groupby("CODE").BASE_COST.agg(times_performed="size", avg_base_cost="mean", min_cost="min", max_cost="max",
                                        total_cost="sum").reset_index()
pg["procedure"] = pg.CODE.map(label)
pg["pct_of_all_procedures"] = (pg.times_performed / len(proc) * 100).round(2)
pg["avg_base_cost"] = pg.avg_base_cost.round(2)
cols = ["CODE", "procedure", "times_performed", "pct_of_all_procedures", "avg_base_cost", "min_cost", "max_cost", "total_cost"]
save(pg.sort_values("times_performed", ascending=False).head(10)[cols], "q06_top10_procedures_by_frequency")
save(pg.sort_values("avg_base_cost", ascending=False).head(10)[cols], "q07_top10_by_avg_cost_unfiltered")
save(pg[pg.times_performed >= 30].sort_values("avg_base_cost", ascending=False).head(10)[cols],
     "q07b_top10_by_avg_cost_min30")
save(pg.sort_values("total_cost", ascending=False).head(10)[cols], "q07c_top10_by_total_spend")
save(pd.DataFrame({"metric": ["distinct procedure codes", "procedures at exactly 431", "share at exactly 431",
                              "codes performed fewer than 30 times", "codes performed once"],
                   "value": [len(pg), int((proc.BASE_COST == 431).sum()), round((proc.BASE_COST == 431).mean() * 100, 1),
                             int((pg.times_performed < 30).sum()), int((pg.times_performed == 1).sum())]}),
     "q07d_procedure_cost_context")

# Q8 / Obj2d - claim cost by payer (+ class-mix standardised view)
c8 = enc.groupby("PAYER_NAME").TOTAL_CLAIM_COST.agg(encounters="size", mean="mean", median="median", total="sum").round(2).reset_index()
std_w = enc.ENCOUNTERCLASS.value_counts(normalize=True)
cm = enc.groupby(["PAYER_NAME", "ENCOUNTERCLASS"]).TOTAL_CLAIM_COST.mean().unstack()
c8["class_mix_adjusted_mean"] = c8.PAYER_NAME.map(
    lambda p: float((cm.loc[p].reindex(std_w.index).fillna(cm.mean()) * std_w).sum())).round(2)
mixp = pd.crosstab(enc.PAYER_NAME, enc.ENCOUNTERCLASS, normalize="index") * 100
c8["pct_inpatient"] = c8.PAYER_NAME.map(mixp["inpatient"]).round(1)
c8["pct_urgentcare"] = c8.PAYER_NAME.map(mixp["urgentcare"]).round(1)
save(c8.sort_values("mean", ascending=False), "q08_claim_cost_by_payer")

# ---- Supplementary (SQL Objective 3)
inp = enc[enc.ENCOUNTERCLASS == "inpatient"]
q_any = enc.groupby("START_QUARTER").PATIENT.nunique().rename("unique_patients_any_encounter")
q_inp = inp.groupby("START_QUARTER").PATIENT.nunique().rename("unique_patients_inpatient_admission")
save(pd.concat([q_any, q_inp], axis=1).fillna(0).astype(int).reset_index(), "s01_unique_patients_by_quarter")

e2 = enc.sort_values(["PATIENT", "START_TS"]).copy()
e2["prev_stop"] = e2.groupby("PATIENT").STOP_TS.shift()
e2["gap_days"] = (e2.START_TS - e2.prev_stop).dt.total_seconds() / 86400
anyre = e2[(e2.gap_days >= 0) & (e2.gap_days <= 30)]
i2 = inp.sort_values(["PATIENT", "START_TS"]).copy()
i2["prev_stop"] = i2.groupby("PATIENT").STOP_TS.shift()
i2["gap_days"] = (i2.START_TS - i2.prev_stop).dt.total_seconds() / 86400
inre = i2[(i2.gap_days >= 0) & (i2.gap_days <= 30)]
save(pd.DataFrame({
    "definition": ["A. ANY encounter starting within 30 days of the end of the patient's previous encounter (literal SQL wording)",
                   "B. INPATIENT admission starting within 30 days of the end of the patient's previous inpatient stay"],
    "index_encounters": [len(enc), len(inp)],
    "readmission_events": [len(anyre), len(inre)],
    "events_pct_of_index": [round(len(anyre) / len(enc) * 100, 1), round(len(inre) / len(inp) * 100, 1)],
    "distinct_patients": [anyre.PATIENT.nunique(), inre.PATIENT.nunique()],
    "patients_pct_of_974": [round(anyre.PATIENT.nunique() / pat.shape[0] * 100, 1), round(inre.PATIENT.nunique() / pat.shape[0] * 100, 1)]}),
    "s02_readmission_summary")
top = inre.groupby("PATIENT").size().rename("inpatient_readmissions_30d").sort_values(ascending=False).head(10).reset_index()
top["patient_ref"] = "PT-" + top.PATIENT.str[:8]
tot = inp.groupby("PATIENT").size()
top["total_inpatient_stays"] = top.PATIENT.map(tot)
save(top[["patient_ref", "inpatient_readmissions_30d", "total_inpatient_stays"]], "s03_top_readmitted_patients_B")
topa = anyre.groupby("PATIENT").size().rename("encounters_within_30d_of_previous").sort_values(ascending=False).head(10).reset_index()
topa["patient_ref"] = "PT-" + topa.PATIENT.str[:8]
topa["total_encounters"] = topa.PATIENT.map(enc.groupby("PATIENT").size())
save(topa[["patient_ref", "encounters_within_30d_of_previous", "total_encounters"]], "s03b_top_patients_A")
save(pd.DataFrame({"inpatient_readmission_dist": inre.groupby("PATIENT").size().value_counts().sort_index().index,
                   "patients": inre.groupby("PATIENT").size().value_counts().sort_index().values}), "s03c_readmission_distribution_B")
# readmission rate by inpatient stays per patient
save(inp.groupby("PATIENT").size().value_counts().sort_index().rename_axis("inpatient_stays").rename("patients").reset_index(),
     "s04_inpatient_stays_per_patient")
print("done ->", OUT)
