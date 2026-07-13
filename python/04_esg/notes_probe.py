"""
notes_probe.py
===================================================================
Demonstrate the NOTES-retrieval path that the structured account API
cannot reach: derivative notional and hedge-accounting designation live
in the filing document (사업보고서 주석), retrievable via list.json ->
document.xml. This probes a few firms, downloads the actual annual-report
document, and searches the raw text for derivative / hedge keywords to
prove the data is reachable -- and to measure, honestly, how clean it is.
"""
import json, io, zipfile, re, time
import urllib.request, urllib.parse

KEY = "e5d6c99ddded63ed330b7c230e7c333f276014fa"
BASE = "https://opendart.fss.or.kr/api"

PROBE = {           # corp_code : (name, filing-search window for FY2023 annual report)
    "00138279": ("S-Oil",   ("20240101", "20240415")),
    "00164645": ("HMM",     ("20240101", "20240415")),
    "00155319": ("POSCO홀딩스", ("20240101", "20240415")),
}
KW = ["파생상품", "위험회피", "현금흐름위험회피", "명목금액", "통화선도", "스왑", "파생상품자산"]


def get(url):
    with urllib.request.urlopen(url, timeout=40) as r:
        return r.read()


def find_annual_report(corp_code, bgn, end):
    q = urllib.parse.urlencode(dict(
        crtfc_key=KEY, corp_code=corp_code, bgn_de=bgn, end_de=end,
        pblntf_detail_ty="A001", page_count="10"))
    js = json.loads(get(f"{BASE}/list.json?{q}").decode("utf-8"))
    if js.get("status") != "000":
        return None, js.get("message")
    for it in js.get("list", []):
        if "사업보고서" in it.get("report_nm", ""):
            return it, None
    return (js["list"][0] if js.get("list") else None), None


def fetch_document_text(rcept_no):
    data = get(f"{BASE}/document.xml?crtfc_key={KEY}&rcept_no={rcept_no}")
    # document.xml returns a zip of the filing (XML/HTML fragments)
    texts = []
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        for nm in z.namelist():
            raw = z.read(nm)
            for enc in ("utf-8", "cp949", "euc-kr"):
                try:
                    texts.append(raw.decode(enc)); break
                except UnicodeDecodeError:
                    continue
    except zipfile.BadZipFile:
        for enc in ("utf-8", "cp949", "euc-kr"):
            try:
                texts.append(data.decode(enc)); break
            except UnicodeDecodeError:
                continue
    return "\n".join(texts)


def strip_tags(s):
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", s)


def main():
    out = {}
    for cc, (name, (bgn, end)) in PROBE.items():
        rec = dict(name=name)
        it, err = find_annual_report(cc, bgn, end)
        if not it:
            rec["error"] = err or "no filing found"; out[cc] = rec
            print(f"[{name}] no annual report ({err})"); continue
        rcept = it["rcept_no"]
        rec["report_nm"] = it["report_nm"].strip()
        rec["rcept_no"] = rcept
        rec["rcept_dt"] = it.get("rcept_dt")
        print(f"[{name}] {rec['report_nm']}  rcept_no={rcept}  ({rec['rcept_dt']})")
        try:
            txt = fetch_document_text(rcept)
        except Exception as e:
            rec["error"] = f"document fetch: {e}"; out[cc] = rec; continue
        plain = strip_tags(txt)
        rec["document_chars"] = len(plain)
        rec["keyword_hits"] = {k: plain.count(k) for k in KW}
        # pull a couple of derivative-notional context snippets as evidence
        snaps = []
        for m in re.finditer("명목금액", plain):
            seg = plain[max(0, m.start()-60): m.start()+80]
            snaps.append(seg.strip())
            if len(snaps) >= 2:
                break
        rec["notional_context"] = snaps
        out[cc] = rec
        print(f"    chars={rec['document_chars']:,}  hits={rec['keyword_hits']}")
        for s in snaps:
            print(f"    …{s}…")
        time.sleep(0.4)

    with open("notes_probe.json", "w") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print("\nwrote notes_probe.json")


if __name__ == "__main__":
    main()
