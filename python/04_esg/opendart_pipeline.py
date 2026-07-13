"""
opendart_pipeline.py
===================================================================
REAL OpenDART extraction for the ESG-disclosure/hedging paper.

Executes the front of the pipeline described in the paper's Section 6:
  corpCode  ->  resolve target firms by stock_code
            ->  fnlttSinglAcntAll (annual consolidated statements)
            ->  extract total assets / revenue / debt / equity (real KRW)
            ->  build a firm-year panel of the DENOMINATORS and controls.

Derivative notional and hedge-accounting designation live in the filing
NOTES (주석), which are in the filing document, NOT in the structured
account API -- so this script builds the clean structured layer and, in a
separate step (notes_probe.py), demonstrates the document-retrieval path
for the notes. This is exactly the split the paper flags: the structured
denominators are easy and reproducible; the notes-level outcomes are the
labour-intensive remaining step.
"""
import json, time, io, zipfile
import urllib.request, urllib.parse
import xml.etree.ElementTree as ET

KEY = "e5d6c99ddded63ed330b7c230e7c333f276014fa"
BASE = "https://opendart.fss.or.kr/api"

# --- target sample: the paper's commodity-exposed sectors + broad/FX panel ---
TARGETS = {
    "010950": ("S-Oil", "refining"),
    "096770": ("SK이노베이션", "refining"),
    "003490": ("대한항공", "airline"),
    "089590": ("제주항공", "airline"),
    "011200": ("HMM", "shipping"),
    "028670": ("팬오션", "shipping"),
    "005880": ("대한해운", "shipping"),
    "005490": ("POSCO홀딩스", "steel"),
    "004020": ("현대제철", "steel"),
    "016380": ("KG스틸", "steel"),
    "005930": ("삼성전자", "broad_fx"),
    "005380": ("현대차", "broad_fx"),
    "051910": ("LG화학", "broad_fx"),
    "000270": ("기아", "broad_fx"),
}
YEARS = [2020, 2021, 2022, 2023]

# accounts we want, with common aliases (account_nm as reported by K-IFRS)
WANT = {
    "total_assets": ["자산총계"],
    "total_liab":   ["부채총계"],
    "total_equity": ["자본총계"],
    "revenue":      ["수익(매출액)", "매출액", "영업수익", "수익"],
}


def resolve_corp_codes():
    m = {}
    for x in ET.parse("CORPCODE.xml").getroot().findall("list"):
        sc = (x.findtext("stock_code") or "").strip()
        if sc in TARGETS:
            m[sc] = x.findtext("corp_code")
    return m


def fetch_fs(corp_code, year):
    q = urllib.parse.urlencode(dict(
        crtfc_key=KEY, corp_code=corp_code, bsns_year=str(year),
        reprt_code="11011", fs_div="CFS"))
    url = f"{BASE}/fnlttSinglAcntAll.json?{q}"
    with urllib.request.urlopen(url, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def pick(rows, aliases):
    # prefer 재무상태표 (BS) / 손익계산서 (IS); match by account_nm, take current-term amount
    for al in aliases:
        for row in rows:
            if (row.get("account_nm") or "").strip() == al:
                amt = (row.get("thstrm_amount") or "").replace(",", "").strip()
                if amt not in ("", "-"):
                    try:
                        return float(amt)
                    except ValueError:
                        pass
    return None


def main():
    codes = resolve_corp_codes()
    print("resolved corp_codes:")
    for sc, cc in codes.items():
        print(f"   {sc} {TARGETS[sc][0]:16s} -> {cc}")

    panel = []
    for sc, cc in codes.items():
        name, sector = TARGETS[sc]
        for y in YEARS:
            try:
                js = fetch_fs(cc, y)
            except Exception as e:
                print(f"   [ERR] {name} {y}: {e}"); continue
            if js.get("status") != "000":
                print(f"   [{js.get('status')}] {name} {y}: {js.get('message')}");
                time.sleep(0.25); continue
            rows = js["list"]
            rec = dict(stock_code=sc, corp_code=cc, name=name, sector=sector, year=y)
            for k, al in WANT.items():
                rec[k] = pick(rows, al)
            # simple derived ratios (denominators for the paper's proxies)
            if rec["total_assets"] and rec["total_liab"] is not None:
                rec["leverage"] = rec["total_liab"]/rec["total_assets"]
            panel.append(rec)
            print(f"   OK  {name:16s} {y}  assets={rec['total_assets']}  rev={rec['revenue']}")
            time.sleep(0.25)

    with open("real_panel.json", "w") as f:
        json.dump(panel, f, ensure_ascii=False, indent=2)
    print(f"\nwrote real_panel.json  ({len(panel)} firm-years, "
          f"{len({r['name'] for r in panel})} firms)")
    return panel


if __name__ == "__main__":
    main()
