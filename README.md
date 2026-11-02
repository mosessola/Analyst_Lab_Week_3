# Hospital Operations, Cost & Patient Flow Intelligence

**AnalystLab Africa Experience Lab | Week 3 | Track A: Data Analytics**

An analysis of hospital encounters, costs, payer coverage and procedures (January 2011 to February 2022), with a data model, a written report with recommendations, and an interactive dashboard.

## Business problem

A healthcare organisation wants to understand its encounter volumes, costs, payer coverage, procedure use and patient return visits. This project asks: *how have encounters, costs and payer coverage behaved over time, where is money going unrecovered, and which patterns need management attention?*

The intended users are hospital operations managers and finance or revenue-cycle leads. All figures are billed amounts for a single hospital, not cash collected, and nothing here is a clinical claim.

## Key findings

- Volume is stable at roughly 2,200 to 2,500 encounters a year, apart from two spikes: February 2014 (1,179 encounters in a month) and 2021 (1,001 vaccine-administration encounters). 2022 holds only five weeks of data.
- Ambulatory and outpatient care make up 67.5% of encounters; inpatient is 4.1%.
- 13,586 encounters (48.71%) show zero payer coverage. Payers cover 30.6% of the $101.5M billed.
- Five commercial payers (Anthem, Cigna, Aetna, Humana, UnitedHealthcare) show $14.1M billed against $8,640 recorded coverage. Whether this is non-payment or a data-capture gap needs client confirmation.
- Urgent care is 13.1% of encounters but 23.0% of billed cost, driven by electrical cardioversion procedures.
- Ten patients (1% of 974) account for 20.6% of encounters; 17 patients account for 397 of 409 inpatient returns within 30 days.
- Several fields are unreliable as headline metrics: every inpatient stay lasts exactly 24 hours, 65% of procedures are priced at a flat $431, and claim totals cannot be rebuilt from procedure costs.

## What is in this repository

| Path | Description |
|---|---|
| `report/Week3_TrackA_Analysis_and_Recommendations.docx` | Full report: data-quality assessment, data model, all eight analysis questions, validation, six recommendations, limitations |
| `dashboard/hospital_dashboard.html` | Interactive dashboard (open in any browser; needs internet access to load Chart.js from cdnjs) |
| `src/week3_analysis.py` | Builds the data model tables and every result from the source workbook |
| `src/week3_independent_verification.py` | Recomputes headline figures in SQL (SQLite) to check the Python results |

## Data model

A star schema with two fact tables and four dimensions:

- `fact_encounters` (one row per encounter) and `fact_procedures` (one row per procedure)
- `dim_patient`, `dim_payer`, `dim_organization`, `dim_date`

Procedures link to patients through their encounter only; the direct procedure-to-patient relationship is deliberately left inactive to avoid ambiguous filtering in Power BI. Patient names, addresses and coordinates are excluded from the model.

## Methodology and decisions

- No rows were deleted and no missing values were imputed. Derived columns (duration band, quarter, uncovered amount, zero-coverage and zero-cost flags) are added alongside the raw values.
- "Over 24 hours" is reported in three bands (under, exactly, over 24h) because 1,077 encounters last exactly 24 hours.
- Procedures are grouped by SNOMED code, since some codes carry more than one description.
- Costs are skewed, so medians are shown next to averages.
- Readmission is reported two ways: any encounter within 30 days (not meaningful as a readmission rate) and an inpatient-to-inpatient return within 30 days.

## Validation

- Eight headline calculations were recomputed independently in SQL and matched exactly.
- The dashboard's filter logic was tested against pandas for six filter combinations (year, class and payer, alone and together): encounters, unique patients, billed, covered, zero-coverage, over-24h and median all matched.
- The source workbook was compared with the SQL setup script: row counts agree, except the script has 973 patients against 974 in the workbook.

## How to reproduce

Requirements: Python 3.10+, `pandas`, `openpyxl`.

```bash
pip install pandas openpyxl
python src/week3_analysis.py <path-to-merged-workbook.xlsx> <output-folder>
python src/week3_independent_verification.py <path-to-merged-workbook.xlsx> <output-folder>
```

The first script writes `model_tables/` (six CSVs) and `results/` (one CSV per question plus the data-quality log) to the output folder. The workbook is the six dataset sheets merged into one file (encounters, patients, procedures, payers, organizations, data dictionary).

The setup script `create_hospital_db.sql` drops and recreates a database named `hospital_db`; it was read but never executed. Run it only in an isolated practice environment.

## Data and privacy

The source data comes from the AnalystLab Africa healthcare dataset pack and is **not included** in this repository. The model tables contain patient-level records (pseudonymous IDs, birth dates, ZIP codes), so keep them out of any public repository. The dashboard contains aggregate data and anonymous patient indexes only, with no names or identifiers.

## Limitations

- Single hospital, no outcome or diagnosis data, so no benchmarking and no clinical conclusions.
- Costs are billed amounts, not payments received.
- Timestamps are UTC, so a few encounters near midnight on 31 December may belong to a different local year.
- A handful of patients carry large weight in totals and averages.

## Next steps

Power BI version of the dashboard, executive summary, and client answers to the open questions in the report (the February 2014 spike, near-zero commercial coverage, flat procedure pricing, and cardioversion recorded in urgent care).

## Author

[Your name] | AnalystLab Africa Experience Lab, Week 3 | [LinkedIn profile] | #AnalystLabAfrica
