"""Discovery round 2: ABS Data API dataflows + Wayback availability + data.vic alternatives."""
import os, re, json, time
from curl_cffi import requests as cr

out = []
def rec(*a):
    s = " ".join(map(str, a)); print(s, flush=True); out.append(s)

# 1. ABS dataflows
r = cr.get("https://data.api.abs.gov.au/rest/dataflow/ABS?detail=allstubs", impersonate="chrome", timeout=60)
flows = re.findall(r'<structure:Dataflow id="([^"]+)"[^>]*version="([^"]+)"[^>]*>.*?<common:Name xml:lang="en">([^<]+)<', r.text, re.S)
rec(f"## ABS dataflows: {len(flows)}")
for fid, ver, name in flows:
    if re.search(r"C21|C16|C11|CENSUS|ERP|SAL|SA2|POP|DWELL|TSP", fid + name, re.I):
        rec(f"{fid} v{ver} :: {name}")

# 2. wayback
for u in ["www.land.vic.gov.au/__data/assets/excel_doc/0033/775617/houses-by-suburb-2015-2025.xlsx",
          "www.land.vic.gov.au/__data/assets/excel_doc/0034/775618/units-by-suburb-2015-2025.xlsx",
          "www.land.vic.gov.au/__data/assets/excel_doc/0033/756582/units-by-suburb-2014-2024.xlsx"]:
    r = cr.get("https://archive.org/wayback/available?url=" + u, impersonate="chrome", timeout=60)
    rec("## wayback", u, r.text[:400])
    try:
        snap = r.json()["archived_snapshots"].get("closest")
        if snap:
            raw = re.sub(r"/web/(\d+)/", r"/web/\1id_/", snap["url"])
            t = cr.get(raw, impersonate="chrome", timeout=90)
            rec("   fetch", raw, t.status_code, t.headers.get("content-type"), len(t.content), t.content[:4])
    except Exception as e:
        rec("   err", e)

# 3. land.vic other paths / data.vic datastore
for u in ["https://discover.data.vic.gov.au/api/3/action/package_search?q=median%20house%20suburb&rows=20",
          "https://discover.data.vic.gov.au/api/3/action/package_search?q=property%20sales%20report&rows=30"]:
    r = cr.get(u, impersonate="chrome", timeout=60)
    try:
        for p in r.json()["result"]["results"]:
            rec("## pkg", p["name"], "|", p["title"])
            for x in p["resources"][:3]:
                rec("     ", x.get("format"), x.get("url"))
    except Exception as e:
        rec("pkgsearch err", r.status_code, e)

os.makedirs("debug", exist_ok=True)
open("debug/diag.txt", "w").write("\n".join(out))
