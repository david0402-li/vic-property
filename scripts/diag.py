import os, io
from curl_cffi import requests as cr
import pandas as pd
out=[]
def rec(*a):
    s=" ".join(map(str,a)); print(s,flush=True); out.append(s)
u="https://files.crimestatistics.vic.gov.au/2026-09/Data_Tables_LGA_Recorded_Offences_Year_Ending_June_2026.xlsx"
r=cr.get(u,impersonate="chrome",timeout=180); rec("crime",r.status_code,len(r.content),r.content[:4])
x=pd.read_excel(io.BytesIO(r.content),sheet_name=None,header=None)
for sn,df in x.items():
    rec(f"=== {sn} {df.shape}"); rec(df.head(8).to_string()[:2500])
    m=df.apply(lambda row: row.astype(str).str.contains("Glen Waverley",case=False).any(),axis=1)
    if m.any(): rec("GW rows:\n"+df[m].head(12).to_string()[:3000])
for u in ["https://content.vic.gov.au/DE-School-Locations%20-2026","https://www.education.vic.gov.au/Documents/about/research/datavic/dv402-SchoolLocations2025.csv"]:
    r=cr.get(u,impersonate="chrome",timeout=120); rec("school",u,r.status_code,len(r.content),r.headers.get("content-type"),r.content[:200])
    try:
        df=pd.read_csv(io.BytesIO(r.content),encoding="latin1"); rec(list(df.columns)); rec(df.head(3).to_string()[:1500])
        rec(df[df.apply(lambda row: row.astype(str).str.contains("Glen Waverley",case=False).any(),axis=1)].to_string()[:3000])
    except Exception as e: rec("ERR",e)
open("debug/diag.txt","w").write("\n".join(out))
