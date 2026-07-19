"""
build_sample.py — Stage 1 of the executed study.
Select a stratified sample of KOSPI non-financial firms and assign
governance-report-mandate cohorts by FY2018 consolidated assets:
  g=2019 : assets >= 2.0tn KRW   (mandatory from FY2018 report, filed 2019)
  g=2022 : 1.0tn <= assets < 2.0tn
  g=2024 : 0.5tn <= assets < 1.0tn
  g=inf  : assets < 0.5tn        (not yet covered through 2024)
Assignment uses FY2018 assets only (intention-to-treat; avoids endogenous
threshold crossing). Financial firms (KSIC 64-66) excluded.
Output: sample.json
"""
import json, time, random
import urllib.request
import xml.etree.ElementTree as ET

KEY = "e5d6c99ddded63ed330b7c230e7c333f276014fa"
BASE = "https://opendart.fss.or.kr/api"
TARGET = {"2019": 100, "2022": 80, "2024": 70, "inf": 150}
random.seed(20260713)


def get(url, tries=3):
    for i in range(tries):
        try:
            return urllib.request.urlopen(url, timeout=25).read()
        except Exception:
            time.sleep(1.0+i)
    return None


def jget(url):
    b = get(url)
    return json.loads(b.decode("utf-8")) if b else {}


def listed_corps():
    rows = []
    for x in ET.parse("CORPCODE.xml").getroot().findall("list"):
        sc = (x.findtext("stock_code") or "").strip()
        if sc:
            rows.append((x.findtext("corp_code"), sc, x.findtext("corp_name")))
    random.shuffle(rows)
    return rows


def assets_2018(cc):
    js = jget(f"{BASE}/fnlttSinglAcnt.json?crtfc_key={KEY}&corp_code={cc}&bsns_year=2018&reprt_code=11011")
    if js.get("status") != "000":
        return None
    for div in ("CFS", "OFS"):
        for r in js.get("list", []):
            if r.get("fs_div") == div and r.get("account_nm") == "자산총계":
                a = (r.get("thstrm_amount") or "").replace(",", "")
                try:
                    return float(a)
                except ValueError:
                    pass
    return None


def main():
    corps = listed_corps()
    got = {k: [] for k in TARGET}
    scanned = 0
    for cc, sc, nm in corps:
        if all(len(got[k]) >= TARGET[k] for k in TARGET):
            break
        scanned += 1
        info = jget(f"{BASE}/company.json?crtfc_key={KEY}&corp_code={cc}")
        if info.get("corp_cls") != "Y":            # KOSPI only
            continue
        ind = (info.get("induty_code") or "")[:2]
        if ind in ("64", "65", "66"):              # financials out
            continue
        a = assets_2018(cc)
        if a is None:                              # no FY2018 CFS/OFS -> skip
            continue
        g = ("2019" if a >= 2e12 else
             "2022" if a >= 1e12 else
             "2024" if a >= 5e11 else "inf")
        if len(got[g]) < TARGET[g]:
            got[g].append(dict(corp_code=cc, stock=sc, name=nm.strip(),
                               induty=info.get("induty_code"), assets2018=a))
            print(f"[{g:>4}] {nm.strip()[:14]:14s} assets {a/1e12:6.2f}tn  "
                  f"({sum(len(v) for v in got.values())} banked / {scanned} scanned)")
        time.sleep(0.12)
    json.dump(got, open("sample.json", "w"), ensure_ascii=False, indent=1)
    print("\ncohort counts:", {k: len(v) for k, v in got.items()})
    print("wrote sample.json")


if __name__ == "__main__":
    main()
