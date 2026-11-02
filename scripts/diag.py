"""Discovery round 3: ABS API data formats + ASGS geo service."""
import os, re, io, time
from curl_cffi import requests as cr
import pandas as pd

out = []
def rec(*a):
    s = " ".join(map(str, a)); print(s, flush=True); out.append(s)

prev = open("debug/diag.txt").read() if os.path.exists("debug/diag.txt") else ""
for line in prev.splitlines():
    if re.match(r"C21_T\d+_SA2", line) and re.search(r"dwelling|structure", line, re.I):
        rec("T-table:", line)

API = "https://data.api.abs.gov.au/rest/data/"
def csv(flow, key="all", extra=""):
    url = f"{API}{flow}/{key}?format=csvfilewithlabels{extra}"
    t = time.time()
    r = cr.get(url, impersonate="chrome", timeout=300)
    rec(f"\n## {flow} {key} -> {r.status_code} {len(r.content)/1e6:.1f}MB {time.time()-t:.0f}s")
    if r.status_code != 200:
        rec(r.text[:500]); return None
    df = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    rec("columns:", list(df.columns))
    for c in df.columns:
        if df[c].nunique() < 60:
            rec(f"  {c}: {df[c].nunique()} uniq e.g. {list(df[c].unique()[:12])}")
        else:
            rec(f"  {c}: {df[c].nunique()} uniq e.g. {list(df[c].unique()[:5])}")
    rec(df.head(6).to_string())
    return df

for flow in ["C21_G01_SAL", "C21_G36_SAL", "ABS_ANNUAL_ERP_ASGS2021", "C21_T01_SA2"]:
    try:
        df = csv(flow)
        if df is not None:
            m = df.apply(lambda r: r.astype(str).str.contains("Glen Waverley").any(), axis=1)
            rec("Glen Waverley rows:\n" + df[m].head(25).to_string())
    except Exception as e:
        rec(flow, "ERR", e)

# T table with dwelling structure, try T14..T20
for n in range(13, 21):
    flow = f"C21_T{n:02d}_SA2"
    try:
        r = cr.get(f"https://data.api.abs.gov.au/rest/dataflow/ABS/{flow}", impersonate="chrome", timeout=60)
        nm = re.search(r'<common:Name xml:lang="en">([^<]+)<', r.text)
        rec(flow, nm.group(1) if nm else r.status_code)
    except Exception as e:
        rec(flow, e)

# ASGS geo service
for u in ["https://geo.abs.gov.au/arcgis/rest/services/ASGS2021/SAL/MapServer?f=json",
          "https://geo.abs.gov.au/arcgis/rest/services/ASGS2021/SA2/MapServer?f=json"]:
    try:
        r = cr.get(u, impersonate="chrome", timeout=60)
        rec("##", u, r.status_code, r.text[:300])
    except Exception as e:
        rec(u, "ERR", e)

open("debug/diag.txt", "w").write("\n".join(out))
