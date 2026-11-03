#!/usr/bin/env python3
"""
VIC suburb property data pipeline.

Downloads public Victorian Government + ABS data, normalises it per suburb,
and writes site/data/vic.json for the static website.

Sources (all free, CC-BY):
  - Victorian Property Sales Report (Valuer-General Victoria) via data.vic CKAN API
      * Median house by suburb, 10-yr time series
      * Median unit by suburb, 10-yr time series
  - DFFH Rental Report: moving annual median rents by suburb (quarterly since 2000)
  - ABS Census 2021 Time Series Profile (TSP) for Suburbs and Localities (SAL), VIC
      * population 2011/2016/2021, dwelling structure 2011/2016/2021

Each source is optional: if one fails, the others still build and the failure is
recorded in vic.json -> meta.errors (and in debug/log.txt).
"""
import io, json, os, re, sys, time, zipfile, datetime, traceback
from collections import defaultdict

import pandas as pd
from curl_cffi import requests as cr

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, "cache")
DEBUG = os.path.join(ROOT, "debug")
OUT = os.path.join(ROOT, "site", "data")
STORE = os.path.join(ROOT, "data_store")  # last successfully parsed source data (committed)
for d in (CACHE, DEBUG, OUT, STORE):
    os.makedirs(d, exist_ok=True)

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36 vic-property-bot"}
CKAN = "https://discover.data.vic.gov.au/api/3/action/package_show?id="

PKG_HOUSE = "victorian-property-sales-report-median-house-by-suburb-time-series"
PKG_UNIT = "victorian-property-sales-report-median-unit-by-suburb-time-series"
PKG_RENT = "rental-report-quarterly-moving-annual-rents-by-suburb"
ABS_TSP_URLS_UNUSED = [
    "https://www.abs.gov.au/census/find-census-data/datapacks/download/2021_TSP_SAL_for_VIC_short-header.zip",
    "https://www.abs.gov.au/census/find-census-data/datapacks/download/2021_TSP_all_for_VIC_short-header.zip",
]

LOG = []
ERRORS = []


def log(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    LOG.append(s)


def debug_write(name, text):
    with open(os.path.join(DEBUG, name), "w", encoding="utf-8") as f:
        f.write(text)


# ---------------------------------------------------------------- download
def get(url, tries=3, timeout=120):
    last = None
    for i in range(tries):
        try:
            r = cr.get(url, impersonate="chrome", timeout=timeout, allow_redirects=True)
            if r.status_code >= 400:
                raise RuntimeError(f"HTTP {r.status_code} for {url}")
            return r
        except Exception as e:  # noqa
            last = e
            time.sleep(3 * (i + 1))
    raise last


def ckan_latest(pkg):
    """Return (resource_dict, all_resources) for the newest data file in a CKAN package."""
    r = get(CKAN + pkg)
    res = r.json()["result"]["resources"]
    def key(x):
        return (x.get("period_end") or "", x.get("release_date") or "", x.get("created") or "")
    files = [x for x in res if re.search(r"xls", (x.get("format") or "") + x.get("url", ""), re.I)]
    files.sort(key=key, reverse=True)
    lines = [f"{key(x)} | {x.get('name')} | {x.get('format')} | {x.get('url')}" for x in sorted(res, key=key, reverse=True)]
    debug_write(f"ckan_{pkg}.txt", "\n".join(lines))
    if not files:
        raise RuntimeError(f"no xls resources in {pkg}")
    return files[0], res


def is_excel(b):
    return b[:2] == b"PK" or b[:4] == bytes.fromhex("D0CF11E0")


def wayback(url):
    bare = re.sub(r"^https?://", "", url)
    j = get("https://archive.org/wayback/available?url=" + bare, tries=2, timeout=60).json()
    snap = (j.get("archived_snapshots") or {}).get("closest")
    if not snap:
        raise RuntimeError("no wayback snapshot for " + url)
    raw = re.sub(r"/web/(\d+)/", r"/web/\1id_/", snap["url"]).replace("http://", "https://", 1)
    return raw, snap.get("timestamp")


def download_cached(url, fname, allow_wayback=True):
    """Download url; if blocked (e.g. Cloudflare challenge) fall back to the Internet Archive copy.
    Returns (path, via) where via is 'direct' or 'wayback:<timestamp>'."""
    path = os.path.join(CACHE, fname)
    via = "direct"
    try:
        r = get(url, tries=2, timeout=90)
        if not is_excel(r.content):
            raise RuntimeError(f"not an Excel file ({r.headers.get('content-type')}): likely bot challenge")
        content = r.content
    except Exception as e:
        if not allow_wayback:
            raise
        log(f"  direct download failed ({e}); trying Wayback Machine")
        raw, ts = wayback(url)
        r = get(raw, tries=3, timeout=120)
        if not is_excel(r.content):
            raise RuntimeError("wayback copy is not an Excel file")
        content, via = r.content, f"wayback:{ts}"
    with open(path, "wb") as f:
        f.write(content)
    log(f"  downloaded {url} -> {fname} ({len(content)/1024:.0f} KB) via {via}")
    return path, via


def read_all_sheets(path):
    try:
        return pd.read_excel(path, sheet_name=None, header=None)
    except Exception:
        # some 'xls' files are really xlsx or html
        for eng in ("openpyxl", "xlrd"):
            try:
                return pd.read_excel(path, sheet_name=None, header=None, engine=eng)
            except Exception:
                pass
        raise


def dump_sheets(sheets, name, rows=18):
    out = []
    for sn, df in sheets.items():
        out.append(f"=== sheet '{sn}' shape={df.shape}")
        with pd.option_context("display.max_columns", 40, "display.width", 400):
            out.append(df.head(rows).to_string())
        out.append("")
    debug_write(name, "\n".join(out))


# ---------------------------------------------------------------- names
def norm(name):
    s = str(name).upper().strip()
    s = re.sub(r"\(VIC\.?\)", "", s)
    s = re.sub(r"\s+", " ", s).replace("’", "'").strip(" -*")
    return s


def slug(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def title(name):
    return " ".join(w.capitalize() if not re.match(r"^(MC|MAC)", w) else w.capitalize() for w in name.lower().split())


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if pd.isna(v) else float(v)
    s = re.sub(r"[,$\s]", "", str(v))
    try:
        return float(s)
    except ValueError:
        return None


# ---------------------------------------------------------------- VPSR price time series
def parse_vpsr(path, label):
    sheets = read_all_sheets(path)
    dump_sheets(sheets, f"raw_{label}.txt")
    result = {}
    for sn, df in sheets.items():
        # header row = row with >=5 year-like cells
        hdr_idx, years = None, {}
        for i in range(min(30, len(df))):
            yrs = {}
            for j, v in enumerate(df.iloc[i].tolist()):
                m = re.fullmatch(r"(19|20)\d{2}(\.0)?", str(v).strip())
                if m:
                    yrs[j] = int(float(str(v)))
            if len(yrs) >= 5:
                hdr_idx, years = i, yrs
                break
        if hdr_idx is None:
            continue
        # suburb column: the column left of first year col with most string values
        first_year_col = min(years)
        best, best_cnt = 0, -1
        for j in range(first_year_col):
            col = df.iloc[hdr_idx + 1:, j]
            cnt = sum(isinstance(x, str) and len(x.strip()) > 2 for x in col)
            if cnt > best_cnt:
                best, best_cnt = j, cnt
        for i in range(hdr_idx + 1, len(df)):
            nm = df.iat[i, best]
            if not isinstance(nm, str) or not nm.strip():
                continue
            key = norm(nm)
            if key in ("LOCALITY", "SUBURB", "VICTORIA", "METROPOLITAN", "COUNTRY") or key.startswith("SOURCE"):
                continue
            series = {}
            for j, y in years.items():
                val = num(df.iat[i, j])
                if val and val > 1000:
                    series[y] = round(val)
            if series:
                result[key] = series
        log(f"  {label}: sheet '{sn}' header row {hdr_idx}, years {min(years.values())}-{max(years.values())}, "
            f"suburb col {best}, parsed {len(result)} suburbs")
    return result


# ---------------------------------------------------------------- DFFH rents
RENT_TYPES = {
    "1 bedroom flat": "f1", "2 bedroom flat": "f2", "3 bedroom flat": "f3",
    "2 bedroom house": "h2", "3 bedroom house": "h3", "4 bedroom house": "h4",
    "all properties": "all",
}
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def parse_qdate(v):
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.year, v.month
    s = str(v).strip().lower()
    m = re.match(r"([a-z]{3})[a-z]*[\s\-']*(\d{2,4})", s)
    if m and m.group(1) in MONTHS:
        y = int(m.group(2)); y = y + 2000 if y < 100 else y
        return y, MONTHS[m.group(1)]
    return None


def parse_rents(path):
    sheets = read_all_sheets(path)
    dump_sheets(sheets, "raw_rents.txt", rows=8)
    out = defaultdict(dict)  # group -> type -> {"YYYY-MM": median}
    for sn, df in sheets.items():
        t = None
        for k, code in RENT_TYPES.items():
            if k in str(sn).lower():
                t = code
        if t is None:
            continue
        # find date row and Count/Median row
        date_row = med_row = None
        for i in range(min(10, len(df))):
            row = df.iloc[i].tolist()
            if sum(parse_qdate(x) is not None for x in row) >= 10:
                date_row = i
            if sum(str(x).strip().lower() == "median" for x in row) >= 10:
                med_row = i
        if date_row is None or med_row is None:
            log(f"  rents: sheet '{sn}' skipped (date_row={date_row}, med_row={med_row})")
            continue
        dates = df.iloc[date_row].tolist()
        # forward-fill dates across merged cells
        cur = None
        filled = []
        for x in dates:
            d = parse_qdate(x)
            cur = d or cur
            filled.append(cur)
        med_cols = [(j, filled[j]) for j, x in enumerate(df.iloc[med_row].tolist())
                    if str(x).strip().lower() == "median" and filled[j]]
        name_col = 1 if df.shape[1] > 1 else 0
        n = 0
        for i in range(med_row + 1, len(df)):
            nm = df.iat[i, name_col]
            if not isinstance(nm, str) or not nm.strip():
                continue
            g = nm.strip()
            if g.lower() in ("group total", "total") or g.lower().startswith("source"):
                continue
            ser = {}
            for j, (y, m) in med_cols:
                val = num(df.iat[i, j])
                if val and 30 < val < 10000:
                    ser[f"{y}-{m:02d}"] = round(val)
            if ser:
                out[g][t] = ser
                n += 1
        log(f"  rents: sheet '{sn}' -> {t}: {n} suburb groups, {len(med_cols)} quarters")
    return out


def rent_group_index(groups):
    """Map normalised suburb name -> DFFH group name. Groups look like 'Glen Waverley-Mulgrave'."""
    idx = {}
    for g in groups:
        for p in re.split(r"\s*-\s*", g):
            p = norm(p).replace("MT ", "MOUNT ")
            idx.setdefault(p, g)
        idx.setdefault(norm(g), g)
    # known spelling quirks in DFFH file
    if "WANAGARATTA" in idx:
        idx.setdefault("WANGARATTA", idx["WANAGARATTA"])
    return idx


COMPASS = r"(NORTH|SOUTH|EAST|WEST|CENTRAL|UPPER|LOWER|NORTH EAST|NORTH WEST|SOUTH EAST|SOUTH WEST)"


def rent_match(key, gidx):
    """Return (group, how) where how is 'exact' or 'approx'."""
    k = re.sub(r"\s*\(.*?\)", "", key).strip()
    if k in gidx:
        return gidx[k], "exact" if k == key else "approx"
    cands = []
    m = re.match(rf"^(.*?)\s+{COMPASS}$", k)
    if m:
        base, d = m.group(1), m.group(2)
        cands += [f"{d} {base}", base]
    m = re.match(rf"^{COMPASS}\s+(.*)$", k)
    if m:
        cands += [f"{m.group(2)} {m.group(1)}", m.group(2)]
    if k.startswith("MOUNT "):
        cands.append("MT " + k[6:])
    for c in cands:
        if c in gidx:
            return gidx[c], "approx"
    return None, None


# ---------------------------------------------------------------- ABS Data API
ABS_API = "https://data.api.abs.gov.au/rest/data/"
STATE_CODE = "2"  # Victoria


def abs_csv(flow, key="all"):
    """Fetch an ABS SDMX dataflow as CSV-with-labels. Falls back to 'all' if the filtered key fails."""
    for k in ([key, "all"] if key != "all" else ["all"]):
        url = f"{ABS_API}{flow}/{k}?format=csvfilewithlabels"
        try:
            r = get(url, tries=2, timeout=300)
            df = pd.read_csv(io.BytesIO(r.content), low_memory=False, dtype=str)
            log(f"  ABS {flow} key={k}: {len(df)} rows")
            return df
        except Exception as e:
            log(f"  ABS {flow} key={k} failed: {e}")
    raise RuntimeError(f"ABS {flow} unavailable")


def clean_sal(name):
    return norm(re.sub(r"\s*\((Vic\.?|VIC)\)\s*$", "", str(name)))


def sa2_base(name):
    """'Glen Waverley - West' -> ['GLEN WAVERLEY']; 'Ashburton (Vic.)' -> ['ASHBURTON'];
    'Brunswick East' -> ['BRUNSWICK EAST']; 'Albert Park' -> ['ALBERT PARK']"""
    n = re.sub(r"\s*\((Vic\.?|VIC)\)", "", name)
    parts = [p.strip() for p in re.split(r"\s+-\s+", n)]
    base = parts[0]
    # 'A - B' where B is a compass/qualifier -> base A ; 'A - B' where B is another place -> both
    quals = {"NORTH", "SOUTH", "EAST", "WEST", "CENTRAL", "INNER", "OUTER", "NORTH EAST", "NORTH WEST",
             "SOUTH EAST", "SOUTH WEST", "UPPER", "LOWER", "TOWNSHIP", "BUSHLAND", "RURAL", "ESTATE"}
    out = [norm(base)]
    for q in parts[1:]:
        if norm(q) not in quals and not re.match(r"^(NORTH|SOUTH|EAST|WEST)", norm(q)):
            out.append(norm(q))
    return out


def load_abs():
    """Returns dict with:
       sal: {SUBURB: {"pop2021": int, "dw2021": {...}}}
       sa2: {code: {"name":..., "erp": {year: n}, "dw": {year: {...}}}}
       sa2_by_suburb: {SUBURB: [codes]}"""
    out = {"sal": {}, "sa2": {}, "sa2_by_suburb": defaultdict(list)}

    # SAL 2021 population (G01: SEXP.PCHAR.REGION.REGION_TYPE.STATE)
    try:
        df = abs_csv("C21_G01_SAL", f"3.P_1..SAL.{STATE_CODE}")
        df = df[(df["STATE"] == STATE_CODE) & (df["PCHAR"] == "P_1") & (df["SEXP"] == "3")]
        for _, r in df.iterrows():
            out["sal"].setdefault(clean_sal(r["Region"]), {})["pop2021"] = int(float(r["OBS_VALUE"]))
        log(f"  SAL population: {len(df)} suburbs")
    except Exception as e:
        ERRORS.append(f"ABS SAL population: {e}"); log(traceback.format_exc())

    # SAL 2021 dwelling structure (G36: DWTSTRD.SUM.REGION.REGION_TYPE.STATE)
    try:
        df = abs_csv("C21_G36_SAL", f".D..SAL.{STATE_CODE}")
        df = df[(df["STATE"] == STATE_CODE) & (df["SUM"] == "D")]
        piv = df.pivot_table(index="Region", columns="DWTSTRD", values="OBS_VALUE", aggfunc="first")
        for reg, r in piv.iterrows():
            f = lambda c: float(r[c]) if c in r and pd.notna(r[c]) else 0.0  # noqa
            occ = f("OCC")
            if occ > 0:
                out["sal"].setdefault(clean_sal(reg), {})["dw2021"] = {
                    "flat": round(f("3") / occ * 100, 1), "semi": round(f("2") / occ * 100, 1),
                    "house": round(f("11") / occ * 100, 1), "n": int(occ)}
        log(f"  SAL dwelling structure: {len(piv)} suburbs")
    except Exception as e:
        ERRORS.append(f"ABS SAL dwellings: {e}"); log(traceback.format_exc())

    # SAL 2021 median weekly rent (G02 selected medians)
    try:
        df = abs_csv("C21_G02_SAL")
        df = df[df["STATE"] == STATE_CODE]
        labcol = [c for c in df.columns if not c.isupper() and df[c].astype(str).str.contains("rent", case=False).any()]
        if labcol:
            df = df[df[labcol[0]].str.contains("rent", case=False, na=False)]
            for _, r in df.iterrows():
                v = float(r["OBS_VALUE"]) if pd.notna(r["OBS_VALUE"]) else 0
                if v > 0:
                    out["sal"].setdefault(clean_sal(r["Region"]), {})["rent2021"] = v
            log(f"  SAL census median rent: {len(df)} rows via label col {labcol[0]}")
    except Exception as e:
        ERRORS.append(f"ABS SAL rent: {e}"); log(traceback.format_exc())

    # SA2 ERP annual (MEASURE.REGION_TYPE.ASGS_2021.FREQ)
    try:
        df = abs_csv("ABS_ANNUAL_ERP_ASGS2021", "ERP.SA2..A")
        df = df[(df["REGION_TYPE"] == "SA2") & df["ASGS_2021"].str.startswith(STATE_CODE)]
        for _, r in df.iterrows():
            d = out["sa2"].setdefault(r["ASGS_2021"], {"name": r["Region"], "erp": {}, "dw": {}})
            d["erp"][int(r["TIME_PERIOD"])] = int(float(r["OBS_VALUE"]))
        log(f"  SA2 ERP: {len(out['sa2'])} SA2s")
    except Exception as e:
        ERRORS.append(f"ABS ERP: {e}"); log(traceback.format_exc())

    # SA2 dwelling structure time series 2011/2016/2021 (T18: tenure x dwelling structure)
    try:
        df = abs_csv("C21_T18_SA2")
        debug_write("abs_t18_columns.txt", str(list(df.columns)) + "\n\n" +
                    "\n".join(f"{c}: {sorted(df[c].dropna().unique().tolist())[:40]}"
                              for c in df.columns if df[c].nunique() < 80))
        if "REGION_TYPE" in df:
            df = df[df["REGION_TYPE"] == "SA2"]
        df = df[df["REGION"].astype(str).str.startswith(STATE_CODE)]
        df = df[df["TENLLD"] == "_T"]
        code = {"11": "house", "2": "semi", "3": "flat", "_T": "tot"}
        df = df[df["STRD"].isin(code)].assign(_c=lambda d: d["STRD"].map(code))
        g = df.pivot_table(index=["REGION", "TIME_PERIOD"], columns="_c", values="OBS_VALUE",
                           aggfunc=lambda x: float(x.iloc[0]))
        for (reg, yr), r in g.iterrows():
            tot = r.get("tot")
            if tot and pd.notna(tot) and reg in out["sa2"]:
                out["sa2"][reg]["dw"][int(yr)] = {"flat": float(r.get("flat") or 0), "semi": float(r.get("semi") or 0),
                                                 "tot": float(tot)}
        log(f"  SA2 dwelling TS: {g.shape[0]} region-years, categories {list(g.columns)}")
    except Exception as e:
        ERRORS.append(f"ABS T18: {e}"); log(traceback.format_exc())

    for code, d in out["sa2"].items():
        for b in sa2_base(d["name"]):
            out["sa2_by_suburb"][b].append(code)
    return out


# ---------------------------------------------------------------- postcodes
POSTCODE_URL = "https://raw.githubusercontent.com/matthewproctor/australianpostcodes/master/australian_postcodes.csv"
POSTCODE_MANUAL = {  # VPSR sub-localities that aren't official localities in the postcode list
    "WESTGARTH": "3070", "JOLIMONT": "3002", "FISHERMANS BEND": "3207", "SANCTUARY LAKES": "3030",
    "WESTALL": "3169", "LARA LAKE": "3212", "WOOLAMAI WATERS": "3925", "COWES WEST": "3922",
    "HASTINGS WEST": "3915", "ROSEBUD SOUTH": "3939", "WENDOUREE WEST": "3355",
    "KANGAROO GROUND SOUTH": "3097",
}


def load_postcodes(state="VIC"):
    r = get(POSTCODE_URL, tries=3, timeout=120)
    df = pd.read_csv(io.BytesIO(r.content), dtype=str)
    df = df[df["state"] == state]
    out = {}
    for loc, g in df.groupby("locality"):
        da = g["type"].fillna("").str.contains("Delivery Area")
        d = pd.concat([g[da], g[~da]])
        cands = []
        for pcode in dict.fromkeys(d["postcode"]):
            row = d[d["postcode"] == pcode].iloc[0]
            cands.append({"pc": pcode, "lat": float(row["lat"]) if pd.notna(row["lat"]) else None,
                          "lng": float(row["long"]) if pd.notna(row["long"]) else None,
                          "lga": row.get("lgaregion") if pd.notna(row.get("lgaregion")) else None,
                          "_ctx": norm(" ".join(str(row.get(c, "")) for c in ("lgaregion", "dc")))})
        out[norm(loc)] = cands
    log(f"  postcodes: {len(out)} {state} localities")
    return out


def postcode_for(key, pcs):
    k = re.sub(r"\s*\(.*?\)", "", key).strip()
    qual = re.search(r"\((.*?)\)", key)
    for c in (key, k):
        if c in pcs:
            cands = pcs[c]
            if qual and len(cands) > 1:
                words = [w for w in qual.group(1).split() if len(w) > 2]
                for x in cands:
                    if any(w in x.get("_ctx", "") for w in words if w != "GREATER"):
                        return x
            return cands[0]
    if k in POSTCODE_MANUAL:
        return {"pc": POSTCODE_MANUAL[k], "lat": None, "lng": None, "lga": None}
    m = re.match(rf"^(.*?)\s+{COMPASS}$", k)
    if m and m.group(1) in pcs:
        return dict(pcs[m.group(1)][0], lat=None, lng=None)
    return {}


# ---------------------------------------------------------------- last-good store
def store_save(name, obj, meta):
    with open(os.path.join(STORE, name + ".json"), "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "data": obj}, f, ensure_ascii=False)


def store_load(name):
    path = os.path.join(STORE, name + ".json")
    if not os.path.exists(path):
        return None, None
    j = json.load(open(path, encoding="utf-8"))
    data = j["data"]
    if name in ("house", "unit"):
        data = {k: {int(y): v for y, v in ser.items()} for k, ser in data.items()}
    return data, j["meta"]


# ---------------------------------------------------------------- build
def annual_from_quarterly(ser):
    """quarterly {'YYYY-MM': v} -> annual {YYYY: v} using latest quarter in each year."""
    out = {}
    for k in sorted(ser):
        out[int(k[:4])] = ser[k]
    return out


def growth(series, n):
    if not series:
        return None
    last = max(series)
    if last - n in series and series[last - n]:
        return round(((series[last] / series[last - n]) ** (1 / n) - 1) * 100, 1)
    return None


def main():
    meta = {"generated": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%MZ"), "state": "VIC",
            "sources": {}, "errors": []}
    house, unit, rents = {}, {}, {}
    absd = {"sal": {}, "sa2": {}, "sa2_by_suburb": {}}

    for label, pkg in (("house", PKG_HOUSE), ("unit", PKG_UNIT)):
        try:
            log(f"[{label}] CKAN {pkg}")
            res, _ = ckan_latest(pkg)
            p, via = download_cached(res["url"], f"vpsr_{label}.xlsx")
            data = parse_vpsr(p, label)
            if len(data) < 100:
                raise RuntimeError(f"only {len(data)} suburbs parsed")
            store_save(label, data, {"name": res.get("name"), "url": res["url"], "via": via,
                                     "fetched": meta["generated"]})
            if label == "house":
                house = data
            else:
                unit = data
            meta["sources"][label] = {"name": res.get("name"), "url": res["url"], "via": via}
        except Exception as e:
            log(traceback.format_exc())
            data, m = store_load(label)
            if data:
                house, unit = (data, unit) if label == "house" else (house, data)
                meta["sources"][label] = dict(m, stale=True)
                ERRORS.append(f"{label}: 本次下载失败，沿用 {m.get('fetched','')} 的数据")
            else:
                ERRORS.append(f"{label}: {e}")

    try:
        log("[rent] CKAN " + PKG_RENT)
        res, _ = ckan_latest(PKG_RENT)
        p, via = download_cached(res["url"], "dffh_rents.xlsx")
        rents = parse_rents(p)
        if len(rents) < 50:
            raise RuntimeError(f"only {len(rents)} rent groups parsed")
        store_save("rent", rents, {"name": res.get("name"), "url": res["url"], "via": via, "fetched": meta["generated"]})
        meta["sources"]["rent"] = {"name": res.get("name"), "url": res["url"], "via": via}
    except Exception as e:
        log(traceback.format_exc())
        data, m = store_load("rent")
        if data:
            rents = data
            meta["sources"]["rent"] = dict(m, stale=True)
            ERRORS.append(f"rent: 本次下载失败，沿用 {m.get('fetched','')} 的数据")
        else:
            ERRORS.append(f"rent: {e}")

    log("[abs] ABS Data API")
    absd = load_abs()
    if absd["sal"] and absd["sa2"]:
        store_save("abs", {"sal": absd["sal"], "sa2": absd["sa2"], "sa2_by_suburb": dict(absd["sa2_by_suburb"])},
                   {"fetched": meta["generated"]})
    else:
        old, m = store_load("abs")
        if old:
            for d in old["sa2"].values():
                d["erp"] = {int(k): v for k, v in d["erp"].items()}
                d["dw"] = {int(k): v for k, v in d["dw"].items()}
            absd = old
            ERRORS.append(f"ABS: 本次下载不完整，沿用 {m.get('fetched','')} 的数据")
    meta["sources"]["abs"] = {"name": "ABS Data API (Census 2021 G01/G36 SAL, T18 SA2; ERP by SA2)",
                              "url": "https://data.api.abs.gov.au"}

    try:
        log("[postcodes]")
        pcs = load_postcodes()
        store_save("postcodes", pcs, {"fetched": meta["generated"]})
        meta["sources"]["postcode"] = {"name": "matthewproctor/australianpostcodes", "url": POSTCODE_URL}
    except Exception as e:
        log(traceback.format_exc())
        pcs, _ = store_load("postcodes")
        pcs = pcs or {}
        ERRORS.append(f"postcodes: {e}")

    gidx = rent_group_index(rents.keys())
    # statewide rent index: median across DFFH groups of 'all properties' per quarter
    qs = defaultdict(list)
    for g, t in rents.items():
        for q, v in (t.get("all") or {}).items():
            qs[q].append(v)
    state_idx = {q: sorted(v)[len(v) // 2] for q, v in qs.items() if len(v) > 20}
    base_q = "2021-09" if "2021-09" in state_idx else None
    names = set(house) | set(unit)
    suburbs = []
    for key in sorted(names):
        h, u = house.get(key, {}), unit.get(key, {})
        g, how = rent_match(key, gidx)
        rq = rents.get(g, {}) if g else {}
        sal = absd["sal"].get(key, {})
        if not rq and sal.get("rent2021") and base_q:
            # estimate: 2021 census median rent (all dwellings) indexed by statewide DFFH movement
            f = sal["rent2021"] / state_idx[base_q]
            est = {q: round(v * f) for q, v in state_idx.items() if q >= "2016"}
            rq = {"est": est}
            how = "census-est"
        rh = annual_from_quarterly(rq.get("h3", {}) or rq.get("all", {}) or rq.get("est", {}))
        ru = annual_from_quarterly(rq.get("f2", {}) or rq.get("all", {}) or rq.get("est", {}))
        yh = {y: round(rh[y] * 52 / h[y] * 100, 2) for y in h if y in rh and h[y]}
        yu = {y: round(ru[y] * 52 / u[y] * 100, 2) for y in u if y in ru and u[y]}

        codes = absd["sa2_by_suburb"].get(key, [])
        sa2s = [absd["sa2"][c] for c in codes]
        erp = defaultdict(int)
        for d in sa2s:
            for y, v in d["erp"].items():
                erp[y] += v
        dws = defaultdict(lambda: {"flat": 0.0, "semi": 0.0, "tot": 0.0})
        for d in sa2s:
            for y, v in d["dw"].items():
                for k in ("flat", "semi", "tot"):
                    dws[y][k] += v[k]
        dw_ts = {y: {"flat": round(v["flat"] / v["tot"] * 100, 1), "semi": round(v["semi"] / v["tot"] * 100, 1)}
                 for y, v in sorted(dws.items()) if v["tot"]}

        latest = lambda s: s[max(s)] if s else None  # noqa
        dw21 = sal.get("dw2021")
        pcinfo = {k: v for k, v in postcode_for(key, pcs).items() if not k.startswith("_")}
        rec = {
            "id": slug(key), "name": title(key), "postcode": pcinfo.get("pc"),
            "lat": pcinfo.get("lat"), "lng": pcinfo.get("lng"), "lga": pcinfo.get("lga"), "rentGroup": g, "rentMatch": how,
            "house": h, "unit": u, "rent": {t: {q: v for q, v in ser.items() if q >= "2010"} for t, ser in rq.items()},
            "yieldHouse": yh, "yieldUnit": yu,
            "sa2": [d["name"] for d in sa2s],
            "erp": dict(sorted(erp.items())) or None,
            "dwTs": dw_ts or None,
            "pop2021": sal.get("pop2021"), "dw2021": dw21,
            "latest": {
                "house": latest(h), "unit": latest(u), "year": max(h) if h else (max(u) if u else None),
                "rentH3": latest(rq.get("h3", {})), "rentF2": latest(rq.get("f2", {})), "rentAll": latest(rq.get("all", {})), "rentEst": latest(rq.get("est", {})), "rentMatch": how,
                "rentQ": max(rq.get("all", rq.get("h3", {})) or {"": 0}) or None,
                "yieldHouse": latest(yh), "yieldUnit": latest(yu),
                "g1": growth(h, 1), "g5": growth(h, 5), "g10": growth(h, 10), "gu5": growth(u, 5),
                "pop": sal.get("pop2021"), "erp": latest(erp) if erp else None,
                "popGrowth5": growth(dict(erp), 5) if erp else None,
                "flatShare": dw21["flat"] if dw21 else None,
                "unitShare": round(dw21["flat"] + dw21["semi"], 1) if dw21 else None,
            },
        }
        suburbs.append(rec)

    meta["errors"] = ERRORS
    meta["counts"] = {"suburbs": len(suburbs), "rentGroups": len(rents),
                      "withRent": sum(1 for s in suburbs if s["rentGroup"]),
                      "withRentEst": sum(1 for s in suburbs if s["rentMatch"] == "census-est"),
                      "withYield": sum(1 for s in suburbs if s["latest"]["yieldHouse"] is not None),
                      "withCensus": sum(1 for s in suburbs if s["pop2021"]),
                      "withErp": sum(1 for s in suburbs if s["erp"]),
                      "withPostcode": sum(1 for s in suburbs if s["postcode"])}
    log("counts:", meta["counts"])
    with open(os.path.join(OUT, "vic.json"), "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "suburbs": suburbs}, f, ensure_ascii=False, separators=(",", ":"))
    unmatched = [s["name"] for s in suburbs if not s["rentGroup"]][:300]
    debug_write("unmatched_rent.txt", "\n".join(unmatched) + "\n\nRENT GROUPS:\n" + "\n".join(sorted(rents)))
    debug_write("log.txt", "\n".join(LOG) + "\n\nERRORS:\n" + "\n".join(ERRORS))
    debug_write("sample_glen_waverley.json",
                json.dumps(next((s for s in suburbs if s["id"] == "glen-waverley"), {}), indent=1, ensure_ascii=False))
    import shutil
    os.makedirs(os.path.join(DEBUG, "raw"), exist_ok=True)
    for fn in os.listdir(CACHE):
        if fn.endswith(".xlsx") and os.path.getsize(os.path.join(CACHE, fn)) < 5e6:
            shutil.copy(os.path.join(CACHE, fn), os.path.join(DEBUG, "raw", fn))
    if not suburbs:
        sys.exit(1)


if __name__ == "__main__":
    main()
