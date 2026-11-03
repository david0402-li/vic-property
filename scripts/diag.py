import os, re, io, json
from curl_cffi import requests as cr
import pandas as pd
out=[]
def rec(*a):
    s=" ".join(map(str,a)); print(s,flush=True); out.append(s)
def get(u,t=90): return cr.get(u,impersonate="chrome",timeout=t)
for q in ["school locations","crime suburb recorded offences","crime statistics postcode suburb"]:
    r=get("https://discover.data.vic.gov.au/api/3/action/package_search?rows=8&q="+q.replace(" ","%20"))
    for p in r.json()["result"]["results"]:
        rec("## pkg",p["name"],"|",p["title"],"|",p.get("organization",{}).get("title"))
        for x in sorted(p["resources"],key=lambda x:x.get("created") or "",reverse=True)[:4]:
            rec("    ",x.get("format"),x.get("name"),x.get("url"))
# try crime stats site
for u in ["https://www.crimestatistics.vic.gov.au/crime-statistics/latest-victorian-crime-data/download-data",
          "https://files.crimestatistics.vic.gov.au/"]:
    try:
        r=get(u); rec("##",u,r.status_code,len(r.content)); 
        for m in re.findall(r'href="([^"]+\.(?:xlsx|zip|csv))"',r.text)[:30]: rec("   link",m)
    except Exception as e: rec(u,"ERR",e)
os.makedirs("debug",exist_ok=True); open("debug/diag.txt","w").write("\n".join(out))
