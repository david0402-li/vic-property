import os, io, re, traceback
from curl_cffi import requests as cr
import pandas as pd
out=[]
def rec(*a):
    s=" ".join(map(str,a)); print(s,flush=True); out.append(s)
def save(): os.makedirs("debug",exist_ok=True); open("debug/diag.txt","w").write("\n".join(out))
def fetch(u):
    r=cr.get(u,impersonate="chrome",timeout=180); rec("GET",u,r.status_code,len(r.content),r.content[:4])
    if r.content[:2]==b"PK": return r.content
    j=cr.get("https://archive.org/wayback/available?url="+re.sub(r"^https?://","",u),impersonate="chrome",timeout=60).json()
    snap=(j.get("archived_snapshots") or {}).get("closest"); rec("wayback",snap)
    if snap:
        raw=re.sub(r"/web/(\d+)/",r"/web/\1id_/",snap["url"]); r=cr.get(raw,impersonate="chrome",timeout=180); rec("WB",r.status_code,len(r.content),r.content[:4])
        if r.content[:2]==b"PK": return r.content
for u in ["https://files.crimestatistics.vic.gov.au/2026-09/Data_Tables_LGA_Recorded_Offences_Year_Ending_June_2026.xlsx",
          "https://files.crimestatistics.vic.gov.au/2026-06/Data_Tables_LGA_Recorded_Offences_Year_Ending_March_2026.xlsx"]:
    try:
        b=fetch(u)
        if not b: continue
        x=pd.read_excel(io.BytesIO(b),sheet_name=None,header=None)
        for sn,df in x.items():
            rec(f"=== {sn} {df.shape}"); rec(df.head(6).to_string()[:2000])
            m=df.apply(lambda row: row.astype(str).str.contains("Glen Waverley",case=False).any(),axis=1)
            if m.any(): rec("GW rows:\n"+df[m].head(10).to_string()[:2500])
        break
    except Exception: rec(traceback.format_exc())
    save()
for u in ["https://content.vic.gov.au/DE-School-Locations%20-2026","https://www.education.vic.gov.au/Documents/about/research/datavic/dv402-SchoolLocations2025.csv"]:
    try:
        r=cr.get(u,impersonate="chrome",timeout=120); rec("school",u,r.status_code,len(r.content),r.headers.get("content-type"),r.content[:150])
        df=pd.read_csv(io.BytesIO(r.content),encoding="latin1"); rec(list(df.columns)); rec(df.head(3).to_string()[:1500])
        rec(df[df.apply(lambda row: row.astype(str).str.contains("Glen Waverley",case=False).any(),axis=1)].to_string()[:3000])
    except Exception: rec(traceback.format_exc())
    save()
save()
