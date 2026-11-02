"""Independent check of headline figures via SQL (SQLite). Usage: python week3_independent_verification.py <workbook.xlsx> <output-folder-used-by-week3_analysis.py>"""
import sys, sqlite3, pandas as pd
SRC = sys.argv[1]; m = sys.argv[2].rstrip('/') + '/model_tables/'
raw = {k.strip(): v for k, v in pd.read_excel(SRC, sheet_name=None).items()}
for k in raw: raw[k].columns=[str(c).strip() for c in raw[k].columns]
con=sqlite3.connect(':memory:')
raw['encounters'].to_sql('encounters',con,index=False)
raw['procedures'].to_sql('procedures',con,index=False)
raw['Payers'].to_sql('payers',con,index=False)
q=lambda s: pd.read_sql(s,con)
print(q("select substr(START,1,4) y,count(*) n from encounters group by 1").T.to_string())
print(q("select count(*) zero_cov, round(100.0*count(*)/(select count(*) from encounters),2) pct from encounters where PAYER_COVERAGE=0"))
print(q("select ENCOUNTERCLASS,count(*) n, round(100.0*count(*)/(select count(*) from encounters),2) pct from encounters group by 1 order by 2 desc"))
print(q("select p.NAME, round(avg(e.TOTAL_CLAIM_COST),2) avg_claim from encounters e join payers p on p.Id=e.PAYER group by 1 order by 2 desc"))
print(q("select CODE, count(*) n, round(avg(BASE_COST),2) a from procedures group by 1 order by 2 desc limit 10"))
print(q("""select sum(case when (julianday(replace(replace(STOP,'T',' '),'Z',''))-julianday(replace(replace(START,'T',' '),'Z','')))*24>24 then 1 else 0 end) over24,
 sum(case when abs((julianday(replace(replace(STOP,'T',' '),'Z',''))-julianday(replace(replace(START,'T',' '),'Z','')))*24-24)<0.001 then 1 else 0 end) eq24 from encounters"""))
# readmission B independent via SQL self-join (inpatient)
print(q("""select count(*) events, count(distinct a.PATIENT) pts from encounters a
 join encounters b on a.PATIENT=b.PATIENT and a.ENCOUNTERCLASS='inpatient' and b.ENCOUNTERCLASS='inpatient'
 and b.START>=a.STOP and julianday(replace(replace(b.START,'T',' '),'Z',''))-julianday(replace(replace(a.STOP,'T',' '),'Z',''))<=30
 and not exists (select 1 from encounters c where c.PATIENT=a.PATIENT and c.ENCOUNTERCLASS='inpatient' and c.START>a.START and c.START<b.START)"""))
# model-table integrity
fe=pd.read_csv(m+'fact_encounters.csv'); fp=pd.read_csv(m+'fact_procedures.csv'); dp=pd.read_csv(m+'dim_patient.csv'); dy=pd.read_csv(m+'dim_payer.csv'); dd=pd.read_csv(m+'dim_date.csv')
print('rows',len(fe),len(fp),len(dp),len(dy),len(dd))
print('fk',fe.PATIENT.isin(dp.Id).all(),fe.PAYER.isin(dy.Id).all(),fp.ENCOUNTER.isin(fe.Id).all(), pd.to_datetime(fe.START_DATE).isin(pd.to_datetime(dd.DATE)).all())
print('sums',round(fe.TOTAL_CLAIM_COST.sum(),2),round(fe.PAYER_COVERAGE.sum(),2),fp.BASE_COST.sum())
print(list(dp.columns)); print(dp.ZIP.head(3).tolist())
