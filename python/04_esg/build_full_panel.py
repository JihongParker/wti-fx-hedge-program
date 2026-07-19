"""
build_full_panel.py — Stage 2 of the executed study.
For every sampled firm x fiscal year 2017–2024:
  * locate the annual report (사업보고서) via one list.json sweep per firm,
  * download document.xml, parse:
      - HedgeAccountingDummy  (designation read from the operative sentence
                               — the pilot-validated parser),
      - DerivUse dummy        (any true derivative-instrument mention),
      - FX-forward mention count (descriptive),
  * fetch total assets / liabilities (fnlttSinglAcnt, CFS->OFS fallback).
Checkpointed: appends one JSON line per firm-year to panel.jsonl and skips
rows already present, so the run is resumable.
"""
import json, os, re, time
import urllib.request

from derivative_parser import fetch_doc, designation_status  # validated parser

KEY = "e5d6c99ddded63ed330b7c230e7c333f276014fa"
BASE = "https://opendart.fss.or.kr/api"
YEARS = list(range(2016, 2025))          # fiscal years; report filed year+1
DERIV_KW = ["통화선도", "통화스왑", "이자율스왑", "통화옵션", "선물환", "상품스왑", "원자재스왑"]


def jget(url, tries=3):
    for i in range(tries):
        try:
            return json.loads(urllib.request.urlopen(url, timeout=40).read().decode("utf-8"))
        except Exception:
            time.sleep(1.0+i)
    return {}


def annual_reports(cc):
    """one sweep: fiscal-year -> rcept_no of the (latest-corrected) 사업보고서"""
    out = {}
    for pg in (1, 2):
        js = jget(f"{BASE}/list.json?crtfc_key={KEY}&corp_code={cc}"
                  f"&bgn_de=20170101&end_de=20250701&pblntf_detail_ty=A001"
                  f"&page_no={pg}&page_count=100")
        for it in js.get("list", []):
            nm = it.get("report_nm", "")
            m = re.search(r"사업보고서\s*\((\d{4})\.", nm)
            if m:
                fy = int(m.group(1))
                # keep the latest rcept (corrections supersede)
                if fy not in out or it["rcept_no"] > out[fy]:
                    out[fy] = it["rcept_no"]
        if not js.get("list") or len(js.get("list", [])) < 100:
            break
    return out


def financials(cc, year):
    js = jget(f"{BASE}/fnlttSinglAcnt.json?crtfc_key={KEY}&corp_code={cc}"
              f"&bsns_year={year}&reprt_code=11011")
    a = l = None
    if js.get("status") == "000":
        for div in ("CFS", "OFS"):
            for r in js.get("list", []):
                if r.get("fs_div") != div:
                    continue
                amt = (r.get("thstrm_amount") or "").replace(",", "")
                try:
                    v = float(amt)
                except ValueError:
                    continue
                if r.get("account_nm") == "자산총계" and a is None:
                    a = v
                if r.get("account_nm") == "부채총계" and l is None:
                    l = v
            if a is not None:
                break
    return a, l


def parse_year(cc, rcept):
    try:
        full = fetch_doc(rcept)
    except Exception as e:
        return dict(error=str(e)[:60])
    applied, snip = designation_status(full)
    plain = re.sub(r"<[^>]+>", " ", full)
    deriv_hits = sum(plain.count(k) for k in DERIV_KW)
    return dict(ha=applied, deriv_use=int(deriv_hits > 0),
                fxfwd=plain.count("통화선도"), doc_chars=len(plain),
                snippet=snip[:80])


def main():
    sample = json.load(open("sample.json"))
    firms = [dict(f, cohort=g) for g, fs in sample.items() for f in fs]
    done = set()
    if os.path.exists("panel.jsonl"):
        for ln in open("panel.jsonl"):
            r = json.loads(ln)
            done.add((r["corp_code"], r["year"]))
    out = open("panel.jsonl", "a")
    t0 = time.time()
    for fi, f in enumerate(firms):
        cc = f["corp_code"]
        if all((cc, y) in done for y in YEARS):
            continue
        reps = annual_reports(cc)
        for y in YEARS:
            if (cc, y) in done:
                continue
            row = dict(corp_code=cc, name=f["name"], cohort=f["cohort"],
                       induty=f["induty"], assets2018=f["assets2018"], year=y)
            rc = reps.get(y)
            if rc:
                row["rcept"] = rc
                row.update(parse_year(cc, rc))
                time.sleep(0.15)
            else:
                row["error"] = "no annual report"
            a, l = financials(cc, y)
            row["assets"], row["liab"] = a, l
            out.write(json.dumps(row, ensure_ascii=False)+"\n")
            out.flush()
        el = time.time()-t0
        print(f"[{fi+1}/{len(firms)}] {f['name'][:12]:12s} cohort={f['cohort']:>4} "
              f"({el/60:.1f} min elapsed)", flush=True)
    out.close()
    print("panel complete.")


if __name__ == "__main__":
    main()
