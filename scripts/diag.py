import os, io, re, traceback
from curl_cffi import requests as cr
import pandas as pd
out=[]
def rec(*a):
    s=" ".join(map(str,a)); print(s,flush=True); out.append(s)
def save(): os.makedirs("debug",exist_ok=True); open("debug/diag.txt","w").write("\n".join(out))
def get(u,t=120): return cr.get(u,impersonate="chrome",timeout=t)
try:
    r=get("https://data.api.abs.gov.au/rest/dataflow/ABS?detail=allstubs")
    for fid,ver,name in re.findall(r'<structure:Dataflow id="([^"]+)"[^>]*version="([^"]+)"[^>]*>.*?<common:Name xml:lang="en">([^<]+)<', r.text, re.S):
        if re.search(r"dwelling|property price|RPPI|house price|RES_|value of", fid+" "+name, re.I): rec("FLOW",fid,ver,name)
except Exception: rec(traceback.format_exc())
save()
for flow in ["RES_DWELL","RES_DWELL_ST"]:
    try:
        r=get(f"https://data.api.abs.gov.au/rest/data/{flow}/all?format=csvfilewithlabels&startPeriod=2024-Q1",300); rec("##",flow,r.status_code,len(r.content))
        if r.status_code==200:
            df=pd.read_csv(io.BytesIO(r.content)); rec(list(df.columns))
            for c in df.columns:
                if df[c].nunique()<40: rec("  ",c,list(df[c].unique()))
    except Exception: rec(traceback.format_exc())
    save()
for pkg in ["victorian-property-sales-report-time-series","rental-report-quarterly-data-tables"]:
    try:
        j=get("https://discover.data.vic.gov.au/api/3/action/package_show?id="+pkg).json()["result"]["resources"]
        j.sort(key=lambda x:(x.get("period_end") or "",x.get("created") or ""),reverse=True)
        for x in j[:3]: rec("RES",pkg,x.get("name"),x.get("url"))
        u=j[0]["url"]; r=get(u); rec("GET",u,r.status_code,len(r.content),r.content[:4])
        if r.content[:2]!=b"PK" and r.content[:4]!=bytes.fromhex("D0CF11E0"):
            jj=get("https://archive.org/wayback/available?url="+re.sub(r"^https?://","",u),60).json(); snap=(jj.get("archived_snapshots") or {}).get("closest")
            if snap: r=get(re.sub(r"/web/(\d+)/",r"/web/\1id_/",snap["url"])); rec("WB",r.status_code,r.content[:4])
        x=pd.read_excel(io.BytesIO(r.content),sheet_name=None,header=None)
        for sn,df in x.items(): rec(f"=== {pkg} :: {sn} {df.shape}"); rec(df.head(14).to_string()[:2200])
    except Exception: rec(traceback.format_exc())
    save()
