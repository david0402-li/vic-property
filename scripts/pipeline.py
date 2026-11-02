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

import requests
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, "cache")
DEBUG = os.path.join(ROOT, "debug")
OUT = os.path.join(ROOT, "site", "data")
for d in (CACHE, DEBUG, OUT):
    os.makedirs(d, exist_ok=True)

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36 vic-property-bot"}
CKAN = "https://discover.data.vic.gov.au/api/3/action/package_show?id="

PKG_HOUSE = "victorian-property-sales-report-median-house-by-suburb-time-series"
PKG_UNIT = "victorian-property-sales-report-median-unit-by-suburb-time-series"
PKG_RENT = "rental-report-quarterly-moving-annual-rents-by-suburb"
ABS_TSP_URLS = [
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
            r = requests.get(url, headers=UA, timeout=timeout, allow_redirects=True)
            r.raise_for_status()
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


def download_cached(url, fname):
    path = os.path.join(CACHE, fname)
    r = get(url)
    with open(path, "wb") as f:
        f.write(r.content)
    log(f"  downloaded {url} -> {fname} ({len(r.content)/1024:.0f} KB, {r.headers.get('content-type')})")
    return path


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
    """Map normalised suburb name -> DFFH group name. Groups look like 'Glen Waverley' or 'Albert Park-Middle Park-West St Kilda'."""
    idx = {}
    for g in groups:
        parts = re.split(r"\s*[-/&]\s*|\s*,\s*", g)
        for p in parts:
            p = norm(p)
            if p and p not in idx:
                idx[p] = g
        idx.setdefault(norm(g), g)
    return idx


# ---------------------------------------------------------------- ABS census TSP
def parse_census():
    zpath = None
    for url in ABS_TSP_URLS:
        try:
            zpath = download_cached(url, "abs_tsp_vic.zip")
            break
        except Exception as e:
            log(f"  ABS url failed {url}: {e}")
    if not zpath:
        raise RuntimeError("could not download ABS TSP datapack")
    z = zipfile.ZipFile(zpath)
    names = z.namelist()
    debug_write("abs_zip_listing.txt", "\n".join(names))

    # SAL code -> name from metadata workbook
    sal_names = {}
    meta = [n for n in names if re.search(r"geog_desc.*\.xlsx$", n, re.I)]
    for m in meta:
        sheets = pd.read_excel(io.BytesIO(z.read(m)), sheet_name=None, header=None)
        for sn, df in sheets.items():
            for _, row in df.iterrows():
                vals = [str(x) for x in row.tolist()]
                for k, v in enumerate(vals):
                    if re.fullmatch(r"SAL\d{5}", v) and k + 1 < len(vals):
                        sal_names[v] = vals[k + 1]
    log(f"  ABS: {len(sal_names)} SAL names from metadata")

    csvs = [n for n in names if n.lower().endswith(".csv") and "_sal" in n.lower()]
    pop, dw = {}, {}
    hdr_dump = []
    for c in csvs:
        df = pd.read_csv(io.BytesIO(z.read(c)), dtype={0: str})
        cols = list(df.columns)
        hdr_dump.append(f"{c}: {cols[:80]}")
        code_col = cols[0]
        # population: Tot_P_C11 style or Tot_persons_C11_P
        pcols = {}
        for col in cols:
            m = re.fullmatch(r"Tot(?:al)?_(?:P|persons|Persons)_C(\d{2})(?:_P)?", col) or \
                re.fullmatch(r"Tot_persons_C(\d{2})_P", col)
            if m:
                pcols[2000 + int(m.group(1))] = col
        if len(pcols) >= 2 and not pop:
            for _, r in df.iterrows():
                pop[r[code_col]] = {y: int(r[c_]) for y, c_ in pcols.items() if pd.notna(r[c_])}
            log(f"  ABS: population from {c} cols {pcols}")
        # dwelling structure: look for flat/apartment + separate house + total per census year
        if not dw and any(re.search(r"flat|apart", col, re.I) for col in cols):
            groups = defaultdict(dict)
            for col in cols:
                m = re.search(r"C(11|16|21)", col)
                if not m:
                    continue
                y = 2000 + int(m.group(1))
                lc = col.lower()
                if re.search(r"sep_?house|separate", lc):
                    groups[y].setdefault("sep", col)
                elif re.search(r"semi|sd_r_t|row|terr|townh", lc):
                    groups[y].setdefault("semi", col)
                elif re.search(r"flat|apart", lc):
                    groups[y].setdefault("flat", col)
                elif re.search(r"^tot|total", lc):
                    groups[y].setdefault("tot", col)
            if len(groups) >= 2 and all("flat" in g for g in groups.values()):
                for _, r in df.iterrows():
                    d = {}
                    for y, g in groups.items():
                        vals = {k: (float(r[v]) if pd.notna(r[v]) else 0) for k, v in g.items()}
                        tot = vals.get("tot") or sum(vals.get(k, 0) for k in ("sep", "semi", "flat"))
                        if tot > 0:
                            d[y] = {"flat": round(vals.get("flat", 0) / tot * 100, 1),
                                    "semi": round(vals.get("semi", 0) / tot * 100, 1),
                                    "n": int(tot)}
                    dw[r[code_col]] = d
                log(f"  ABS: dwelling structure from {c} cols {dict(groups)}")
    debug_write("abs_csv_headers.txt", "\n".join(hdr_dump))
    by_name = {}
    for code, nm in sal_names.items():
        key = norm(nm)
        by_name[key] = {"pop": pop.get(code), "dw": dw.get(code), "sal": code}
    return by_name


# ---------------------------------------------------------------- build
def annual_from_quarterly(ser):
    """quarterly {'YYYY-MM': v} -> annual {YYYY: v} using latest quarter in each year."""
    out = {}
    for k in sorted(ser):
        out[int(k[:4])] = ser[k]
    return out


def main():
    meta = {"generated": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%MZ"), "sources": {}, "errors": []}

    house = unit = {}
    rents = {}
    census = {}

    for label, pkg in (("house", PKG_HOUSE), ("unit", PKG_UNIT)):
        try:
            log(f"[{label}] CKAN {pkg}")
            res, _ = ckan_latest(pkg)
            p = download_cached(res["url"], f"vpsr_{label}" + os.path.splitext(res["url"].split("?")[0])[1])
            data = parse_vpsr(p, label)
            if label == "house":
                house = data
            else:
                unit = data
            meta["sources"][label] = {"name": res.get("name"), "url": res["url"], "period_end": res.get("period_end")}
        except Exception as e:
            ERRORS.append(f"{label}: {e}")
            log(traceback.format_exc())

    try:
        log("[rent] CKAN " + PKG_RENT)
        res, _ = ckan_latest(PKG_RENT)
        p = download_cached(res["url"], "dffh_rents.xlsx")
        rents = parse_rents(p)
        meta["sources"]["rent"] = {"name": res.get("name"), "url": res["url"], "period_end": res.get("period_end")}
    except Exception as e:
        ERRORS.append(f"rent: {e}")
        log(traceback.format_exc())

    try:
        log("[census] ABS TSP")
        census = parse_census()
        meta["sources"]["census"] = {"name": "ABS 2021 Census Time Series Profile (SAL, VIC)", "url": ABS_TSP_URLS[0]}
    except Exception as e:
        ERRORS.append(f"census: {e}")
        log(traceback.format_exc())

    gidx = rent_group_index(rents.keys())
    names = set(house) | set(unit)
    suburbs = []
    unmatched_rent = 0
    for key in sorted(names):
        h = house.get(key, {})
        u = unit.get(key, {})
        g = gidx.get(key)
        r = rents.get(g, {}) if g else {}
        if not g:
            unmatched_rent += 1
        rq = {t: s for t, s in r.items()}
        rh = annual_from_quarterly(rq.get("h3", {}))
        ru = annual_from_quarterly(rq.get("f2", {}))
        yh = {y: round(rh[y] * 52 / h[y] * 100, 2) for y in h if y in rh and h[y]}
        yu = {y: round(ru[y] * 52 / u[y] * 100, 2) for y in u if y in ru and u[y]}
        c = census.get(key) or {}
        latest = lambda s: s[max(s)] if s else None  # noqa
        rec = {
            "id": slug(key), "name": title(key), "rentGroup": g,
            "house": h, "unit": u,
            "rent": {t: s for t, s in rq.items()},
            "yieldHouse": yh, "yieldUnit": yu,
            "pop": c.get("pop"), "dw": c.get("dw"),
            "latest": {
                "house": latest(h), "unit": latest(u),
                "houseYear": max(h) if h else None,
                "rentH3": latest(rq.get("h3", {})), "rentF2": latest(rq.get("f2", {})),
                "rentAll": latest(rq.get("all", {})),
                "yieldHouse": latest(yh), "yieldUnit": latest(yu),
                "pop": latest(c.get("pop") or {}),
                "flatShare": (c.get("dw") or {}).get(2021, {}).get("flat") if c.get("dw") else None,
            },
        }
        # 5y / 10y growth
        if len(h) >= 2:
            ys = sorted(h)
            last = ys[-1]
            for n in (1, 5, 10):
                if last - n in h and h[last - n]:
                    rec["latest"][f"g{n}"] = round((h[last] / h[last - n]) ** (1 / n) * 100 - 100, 1) if n > 1 \
                        else round(h[last] / h[last - n] * 100 - 100, 1)
        suburbs.append(rec)

    log(f"suburbs: {len(suburbs)}; with rent group: {len(suburbs)-unmatched_rent}; "
        f"with census: {sum(1 for s in suburbs if s['pop'])}")
    meta["errors"] = ERRORS
    meta["counts"] = {"suburbs": len(suburbs), "rentGroups": len(rents),
                      "withRent": len(suburbs) - unmatched_rent,
                      "withCensus": sum(1 for s in suburbs if s["pop"])}
    with open(os.path.join(OUT, "vic.json"), "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "suburbs": suburbs}, f, ensure_ascii=False, separators=(",", ":"))
    debug_write("log.txt", "\n".join(LOG) + "\n\nERRORS:\n" + "\n".join(ERRORS))
    debug_write("sample_glen_waverley.json",
                json.dumps(next((s for s in suburbs if s["id"] == "glen-waverley"), {}), indent=1, ensure_ascii=False))
    if not suburbs:
        sys.exit(1)


if __name__ == "__main__":
    main()
