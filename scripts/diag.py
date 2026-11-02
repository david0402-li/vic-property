"""Try several HTTP client strategies against each blocked host; write debug/diag.txt."""
import os, subprocess, time, traceback
import requests

URLS = {
    "landvic": "https://www.land.vic.gov.au/__data/assets/excel_doc/0033/775617/houses-by-suburb-2015-2025.xlsx",
    "dffh": "https://www.dffh.vic.gov.au/moving-annual-rent-suburb-september-quarter-2025-excel",
    "abs": "https://www.abs.gov.au/census/find-census-data/datapacks/download/2021_TSP_SAL_for_VIC_short-header.zip",
    "abs_api": "https://data.api.abs.gov.au/rest/dataflow/ABS?detail=allstubs",
    "wayback_landvic": "https://archive.org/wayback/available?url=www.land.vic.gov.au/__data/assets/excel_doc/0033/775617/houses-by-suburb-2015-2025.xlsx",
}
H = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-AU,en;q=0.9",
}
out = []
def rec(*a):
    s = " ".join(map(str, a)); print(s, flush=True); out.append(s)

for k, u in URLS.items():
    rec(f"\n## {k} {u}")
    t = time.time()
    try:
        r = requests.get(u, headers=H, timeout=40, stream=True)
        b = next(r.iter_content(2048), b"")
        rec(f"requests+headers: {r.status_code} {r.headers.get('content-type')} len={r.headers.get('content-length')} server={r.headers.get('server')} first={b[:80]!r} {time.time()-t:.1f}s")
        if k.startswith("wayback") or k == "abs_api":
            rec("  body:", r.text[:600] if k.startswith("wayback") else b[:300])
    except Exception as e:
        rec(f"requests+headers: ERR {type(e).__name__} {e} {time.time()-t:.1f}s")
    t = time.time()
    try:
        from curl_cffi import requests as cr
        r = cr.get(u, impersonate="chrome", timeout=40)
        rec(f"curl_cffi chrome: {r.status_code} {r.headers.get('content-type')} len={len(r.content)} first={r.content[:60]!r} {time.time()-t:.1f}s")
    except Exception as e:
        rec(f"curl_cffi chrome: ERR {type(e).__name__} {e} {time.time()-t:.1f}s")
    t = time.time()
    p = subprocess.run(["curl", "-sS", "-L", "--http2", "-m", "40", "-o", "/tmp/x", "-w", "%{http_code} %{content_type} %{size_download}",
                        "-A", H["User-Agent"], "-H", "Accept-Language: en-AU", u], capture_output=True, text=True)
    rec(f"curl cli: {p.stdout} {p.stderr.strip()[:200]} {time.time()-t:.1f}s")

os.makedirs("debug", exist_ok=True)
open("debug/diag.txt", "w").write("\n".join(out))
