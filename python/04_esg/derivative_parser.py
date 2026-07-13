"""
derivative_parser.py
===================================================================
Parse REAL derivative notes from OpenDART FY2023 annual-report documents
for the pilot firms, extracting -- per firm -- what is actually there:

  * derivative fair-value ASSET and LIABILITY, by risk type
    (FX / commodity / interest / other), summed from the note table;
  * the hedge-accounting DESIGNATION STATUS, read from the note's own
    sentence ("위험회피회계를 적용 ... 없으며" = NOT applied, vs
    "적용하고 있습니다"/"위험회피수단으로 지정" = applied) -- NOT from a
    keyword count, which the pilot showed is contaminated by policy
    boilerplate;
  * the notional/계약금액 where a derivative-notional table genuinely
    exists (many issuers disclose only fair value in the report body).

This is the "remaining 30%" done for real on the pilot: it turns the
reachable notes (notes_probe) into measured variables, and documents,
firm by firm, exactly how clean each field is.
"""
import json, io, zipfile, re, time
import urllib.request

KEY = "e5d6c99ddded63ed330b7c230e7c333f276014fa"
BASE = "https://opendart.fss.or.kr/api"

# strict, multi-word derivative instrument names only (bare 선도/선물/옵션/swap are
# too greedy and match convertibles, prepayments, futures margins, etc.)
INSTR = ["통화선도", "통화스왑", "통화옵션", "통화금리스왑", "이자율스왑", "선물환",
         "commodity swap", "상품스왑", "원자재스왑", "금스왑", "cross currency", "ccs",
         "통화선물", "이자율옵션"]
STOP = ["합계", "소계", "총계", "total", "유동", "비유동", "부문", "순액", "차감"]
CMD_KW  = ["commodity", "상품", "원자재", "귀금속", "금스왑", "원유"]
INT_KW  = ["이자율", "irs", "금리", "cross currency", "ccs", "통화금리"]
FX_KW   = ["통화", "선물환", "forward", "currency", "fx"]
HEAD_RE = re.compile(r"파생(금융)?상품[^<]{0,30}(거래\s*현황|계약\s*체결|평가|자산|부채|내역|명세)")


def _get(url, tries=4):
    last = None
    for _ in range(tries):
        try:
            return urllib.request.urlopen(url, timeout=60).read()
        except Exception as e:
            last = e; time.sleep(1.5)
    raise last


def fetch_doc(rcept):
    z = zipfile.ZipFile(io.BytesIO(_get(f"{BASE}/document.xml?crtfc_key={KEY}&rcept_no={rcept}")))
    return "\n".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist())


def find_annual(corp_code):
    q = (f"{BASE}/list.json?crtfc_key={KEY}&corp_code={corp_code}"
         f"&bgn_de=20240101&end_de=20240630&pblntf_detail_ty=A001&page_count=20")
    js = json.loads(_get(q).decode("utf-8"))
    if js.get("status") != "000":
        return None
    cands = [it for it in js.get("list", []) if "사업보고서" in it.get("report_nm", "")]
    return cands[0]["rcept_no"] if cands else None


def cells(tr):
    return [re.sub(r"<[^>]+>", "", c).strip()
            for c in re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.S | re.I)]


def to_num(s):
    s = s.replace(",", "").replace(" ", "")
    if s in ("", "-", "－"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


def classify(name):
    n = name.lower()
    if any(k in n for k in CMD_KW):  return "commodity"
    if any(k in n for k in INT_KW):  return "interest"
    if any(k in n for k in FX_KW):   return "fx"
    return "other"


def designation_status(text):
    """Read hedge-accounting designation from the note's own wording (whole doc)."""
    plain = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", text))
    # the OPERATIVE negative statement ("no hedge-accounting derivatives") is
    # authoritative: a firm that applies hedge accounting never carries it.
    neg_op = re.search(r"위험회피회계를?\s*적용(하는|한|하고\s*있는)\s*파생(금융)?상품[은는이]?\s*없"
                       r"|위험회피회계를?\s*적용하지\s*않"
                       r"|위험회피회계를?\s*적용하고\s*있지\s*않", plain)
    pos = re.search(r"위험회피수단으로\s*지정"
                    r"|현금흐름위험회피회계?를?\s*적용하고\s*있"
                    r"|위험회피관계로\s*지정"
                    r"|위험회피대상[^.]{0,20}지정", plain)
    if neg_op:
        applied = False
    elif pos:
        applied = True
    else:
        applied = None
    anchor = neg_op or pos
    snap = plain[anchor.start():anchor.start()+180] if anchor else ""
    return applied, snap


def note_profile(full):
    """
    Profile the derivative-note tables a firm actually files: which TABLE TYPE
    (fair-value / P&L-impact / notional / hedge-holdings) and which UNIT (KRW mn,
    USD th/mn) each uses. This documents, data-drivenly, why one generic parser
    cannot yield comparable numbers across firms.
    """
    tables = re.findall(r"<TABLE.*?</TABLE>", full, re.S | re.I)
    types, units = set(), set()
    for tb in tables:
        pos = full.find(tb)
        ctx = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", full[max(0, pos-600):pos]))
        txt = re.sub(r"<[^>]+>", " ", tb)
        if not (HEAD_RE.search(ctx) or "파생" in ctx[-200:]):
            continue
        if not _instr_rows(tb):
            continue
        if re.search(r"미친\s*영향|당기손익|평가손익|거래손익", ctx):        types.add("pl_impact")
        elif re.search(r"명목|계약금액|약정금액", txt):                    types.add("notional")
        elif re.search(r"공정가치|자산.{0,3}부채|평가금액", ctx+txt):        types.add("fair_value")
        elif re.search(r"위험회피목적으로\s*보유|위험회피\s*내역", ctx):     types.add("hedge_holdings")
        else:                                                            types.add("other")
        if re.search(r"USD\s*백만|외화\s*백만|백만\s*USD", ctx):           units.add("USD_mn")
        elif re.search(r"USD\s*천|외화\s*천|천\s*USD|외화-?천", ctx):       units.add("USD_th")
        elif re.search(r"백만원|단위\s*:\s*백만", ctx):                    units.add("KRW_mn")
        else:                                                            units.add("unspecified")
    return sorted(types), sorted(units)


def _instr_rows(tb):
    """Clean derivative-instrument rows (subtotals excluded), POSITIONAL:
    '-'/blank map to 0.0 so column alignment (asset, liability, ...) is kept."""
    out = []
    for tr in re.findall(r"<TR.*?</TR>", tb, re.S | re.I):
        r = cells(tr)
        if not r or not r[0]:
            continue
        name = r[0]
        if any(s in name.lower() for s in [x.lower() for x in STOP]):
            continue
        if not any(k.lower() in name.lower() for k in INSTR):
            continue
        vals = []
        for x in r[1:]:
            if x.strip() in ("", "-", "－", "&nbsp;"):
                vals.append(0.0)
            else:
                n = to_num(x)
                vals.append(n if n is not None else None)      # None = non-numeric label cell
        nums = [v for v in vals if v is not None]
        if any(v != 0 for v in nums):
            out.append((name, nums))
    return out


def parse_derivatives(full, assets_krw):
    """
    Heading-anchored: scan tables in document order; the FIRST table after a
    derivative-note heading that carries clean instrument rows is the fair-value
    table; a later table additionally tagged 명목/계약금액 is the notional table.
    One fair-value table per firm (no cross-table summing). Sanity cap: any
    single figure exceeding total assets is rejected as a mis-parse.
    """
    cap = assets_krw / 1e6 if assets_krw else 1e18   # figures are in mn KRW
    tables = re.findall(r"<TABLE.*?</TABLE>", full, re.S | re.I)
    fv = {"fx": [0.0, 0.0], "commodity": [0.0, 0.0], "interest": [0.0, 0.0], "other": [0.0, 0.0]}
    notional = {"fx": 0.0, "commodity": 0.0, "interest": 0.0, "other": 0.0}
    found_fv = found_notional = False
    fv_done = False
    for tb in tables:
        pos = full.find(tb)
        ctx = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", full[max(0, pos-600):pos]))
        txt = re.sub(r"<[^>]+>", " ", tb)
        if not HEAD_RE.search(ctx) and "파생" not in ctx[-200:]:
            continue
        rows = _instr_rows(tb)
        if not rows:
            continue
        is_notional = bool(re.search(r"명목|계약금액|약정금액", txt)) and not re.search(r"공정가치", txt)
        if is_notional:
            for name, nums in rows:
                v = abs(nums[0])
                if v <= cap:
                    notional[classify(name)] += v; found_notional = True
        elif not fv_done:                       # FIRST fair-value table only
            for name, nums in rows:
                asset = nums[0] if len(nums) >= 1 else 0.0
                liab  = nums[1] if len(nums) >= 2 else 0.0
                if abs(asset) <= cap and abs(liab) <= cap:
                    fv[classify(name)][0] += asset
                    fv[classify(name)][1] += liab
                    found_fv = True
            if found_fv:
                fv_done = True
    return dict(fair_value=fv, found_fv=found_fv,
                notional=notional, found_notional=found_notional)


def main():
    panel = json.load(open("real_panel.json"))
    # one rcept per firm (FY2023 annual report); resolve from corp_code
    firms = {}
    for r in panel:
        firms.setdefault(r["name"], dict(corp_code=r["corp_code"], sector=r["sector"],
                                         assets2023=None))
        if r["year"] == 2023:
            firms[r["name"]]["assets2023"] = r["total_assets"]

    out = []
    for name, meta in firms.items():
        rcept = find_annual(meta["corp_code"])
        rec = dict(name=name, sector=meta["sector"], assets2023=meta["assets2023"], rcept=rcept)
        if not rcept:
            rec["error"] = "no annual report"; out.append(rec)
            print(f"[{name}] no report"); time.sleep(0.3); continue
        try:
            full = fetch_doc(rcept)
        except Exception as e:
            rec["error"] = str(e); out.append(rec); print(f"[{name}] {e}"); continue
        applied, snippet = designation_status(full)
        der = parse_derivatives(full, meta["assets2023"])
        types, units = note_profile(full)
        fv = der["fair_value"]
        tot_asset = sum(v[0] for v in fv.values())
        tot_liab  = sum(v[1] for v in fv.values())
        # parse quality: numeric FV is only comparable when the note is a clean
        # single-type KRW fair-value table (S-Oil style). Heterogeneous or non-KRW
        # notes yield non-comparable figures -> flagged, numbers not trusted.
        if types == ["fair_value"] and set(units) <= {"KRW_mn"}:
            quality = "validated_fair_value"
        elif not types:
            quality = "note_not_located"
        else:
            quality = "heterogeneous_not_comparable"
        rec.update(
            hedge_accounting_applied=applied,
            designation_snippet=snippet,
            note_table_types=types,
            note_units=units,
            parse_quality=quality,
            deriv_fv_asset_mnKRW=round(tot_asset, 1),
            deriv_fv_liab_mnKRW=round(tot_liab, 1),
            fv_trustworthy=(quality == "validated_fair_value"),
            fv_by_risk={k: [round(v[0], 1), round(v[1], 1)] for k, v in fv.items()},
            found_notional=der["found_notional"],
        )
        # DerivFVIntensity = (|asset|+|liab|) / total assets   (mn KRW / KRW)
        if meta["assets2023"]:
            rec["deriv_fv_intensity_bps"] = round(
                (tot_asset + tot_liab) * 1e6 / meta["assets2023"] * 1e4, 2)
        out.append(rec)
        print(f"[{name:14s}] HA={applied}  FV asset={tot_asset:,.0f}  liab={tot_liab:,.0f} mnKRW"
              f"  notional={'Y' if der['found_notional'] else 'n'}")
        time.sleep(0.3)

    json.dump(out, open("derivatives_parsed.json", "w"), ensure_ascii=False, indent=2)
    print("\nwrote derivatives_parsed.json")
    return out


if __name__ == "__main__":
    main()
