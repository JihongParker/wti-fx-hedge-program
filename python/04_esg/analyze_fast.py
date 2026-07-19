"""
analyze_fast.py — matrix-vectorized re-implementation of analyze_panel_v2.
Identical estimands and output schema; the panel is reshaped to a
[firms x years] matrix so each group-time cell and each bootstrap
replicate is pure numpy indexing (~1000x faster than the dict loops).
"""
import json
import numpy as np
from esg_hedge_engine import twfe_att, event_study
from analyze_panel import load_panel

RNG = np.random.default_rng(20260714)
MIN_CELL = 3
YEARS = list(range(2016, 2025))
YIDX = {y: i for i, y in enumerate(YEARS)}
NEXT = {2019.0: 2022.0, 2022.0: 2024.0, 2024.0: np.inf}
CUT = {2019.0: 2e12, 2022.0: 1e12, 2024.0: 5e11}


def to_matrix(rows, ykey):
    firms = sorted({r["corp_code"] for r in rows if r.get(ykey) is not None})
    fi = {f: i for i, f in enumerate(firms)}
    M = np.full((len(firms), len(YEARS)), np.nan)
    coh = np.full(len(firms), np.inf)
    x = np.full(len(firms), np.nan)
    for r in rows:
        v = r.get(ykey)
        if v is None or r["corp_code"] not in fi:
            continue
        i = fi[r["corp_code"]]
        M[i, YIDX[r["year"]]] = float(v)
        coh[i] = r["g"]
        x[i] = r.get("assets2018") or np.nan
    return M, coh, x, firms


def cs_matrix(M, coh, x, rule="notyet"):
    """weighted group-time ATT over the matrix panel."""
    num = den = 0.0
    any_cell = False
    for g in (2019.0, 2022.0, 2024.0):
        b = YIDX.get(int(g)-1)
        if b is None:
            continue
        gmask = coh == g
        if rule == "local":
            gmask &= (x >= CUT[g]) & (x <= 1.5*CUT[g])
        for t in YEARS:
            if t < g:
                continue
            ti = YIDX[t]
            if rule == "adjacent":
                nxt = NEXT[g]
                if nxt != np.inf and t >= nxt:
                    continue
                cmask = coh == nxt
            elif rule == "local":
                nxt = NEXT[g]
                cmask = (coh == nxt) & (x >= CUT[g]/1.5) & (x < CUT[g])
                if nxt != np.inf:
                    cmask &= (t < nxt)
            else:
                cmask = coh > t
            dg = M[gmask, ti] - M[gmask, b]
            dc = M[cmask, ti] - M[cmask, b]
            dg = dg[~np.isnan(dg)]
            dc = dc[~np.isnan(dc)]
            if len(dg) < MIN_CELL or len(dc) < MIN_CELL:
                continue
            any_cell = True
            num += (dg.mean() - dc.mean()) * len(dg)
            den += len(dg)
    return num/den if any_cell else float("nan")


def boot_matrix(M, coh, x, rule, nboot=999):
    n = M.shape[0]
    out = []
    for _ in range(nboot):
        idx = RNG.integers(0, n, n)
        s = cs_matrix(M[idx], coh[idx], x[idx], rule)
        if np.isfinite(s):
            out.append(s)
    out = np.array(out)
    if len(out) < 20:
        return float("nan"), (float("nan"), float("nan")), len(out)
    return float(out.std(ddof=1)), (float(np.percentile(out, 2.5)),
                                    float(np.percentile(out, 97.5))), len(out)


def flat_arrays(rows, ykey):
    R = [r for r in rows if r.get(ykey) is not None]
    Y = np.array([float(r[ykey]) for r in R])
    fid = np.array([r["corp_code"] for r in R])
    tid = np.array([r["year"] for r in R])
    coh = np.array([r["g"] for r in R], dtype=float)
    evt = np.array([r["evt"] for r in R])
    return Y, fid, tid, coh, evt


def run(rows, ykey, label, rule="notyet", with_es=True, nboot=999):
    M, coh, x, firms = to_matrix(rows, ykey)
    att = cs_matrix(M, coh, x, rule)
    se, ci, nb = boot_matrix(M, coh, x, rule, nboot)
    Y, fid, tid, cohf, evt = flat_arrays(rows, ykey)
    beta = twfe_att(Y, fid, tid, ((evt >= 0) & (cohf < np.inf)).astype(float))
    out = dict(outcome=label, n_obs=int(np.isfinite(M).sum()), n_firms=len(firms),
               n_treated_firms=int((coh < np.inf).sum()), mean_dep=float(np.nanmean(M)),
               att_cs=att, se_boot=se, ci95=ci, nboot=nb, beta_twfe=float(beta), rule=rule)
    if with_es:
        es = event_study(Y, fid, tid, evt, leads=3, lags=4)
        out["event_study"] = {str(k): v for k, v in es.items()}
        out["pretrend_max_abs"] = float(max(abs(v) for k, v in es.items() if k < -1))
        # bootstrap bands
        firms_u = np.unique(fid)
        idx_by = {f: np.where(fid == f)[0] for f in firms_u}
        draws = {str(k): [] for k in es}
        for _ in range(499):
            pick = RNG.choice(firms_u, size=len(firms_u), replace=True)
            idx = np.concatenate([idx_by[f] for f in pick])
            rel = np.concatenate([[f"{f}#{i}"]*len(idx_by[f]) for i, f in enumerate(pick)])
            try:
                esb = event_study(Y[idx], rel, tid[idx], evt[idx], leads=3, lags=4)
                for k, v in esb.items():
                    draws[str(k)].append(v)
            except Exception:
                pass
        out["event_bands"] = {k: [float(np.percentile(v, 2.5)), float(np.percentile(v, 97.5))]
                              for k, v in draws.items() if v}
    print(f"== {label:32s} rule={rule:8s} N={out['n_obs']} firms={out['n_firms']} "
          f"ATT={att:+.4f} CI [{ci[0]:+.4f},{ci[1]:+.4f}]")
    return out


def main():
    import time
    t0 = time.time()
    rows = load_panel()
    for r in rows:
        r["lfx"] = None if r.get("fxfwd") is None else float(np.log1p(r["fxfwd"]))
    res = {"panel_size": len(rows), "n_firms": len({r["corp_code"] for r in rows}),
           "parse_errors": sum(1 for r in rows if r.get("error"))}
    print(f"panel: {res['panel_size']} rows, {res['n_firms']} firms, errors {res['parse_errors']}")

    for ykey, nm in (("ha_bin", "HA"), ("deriv_use", "DerivUse"), ("lfx", "lnFXfwd")):
        res[nm] = run(rows, ykey, nm, "notyet", with_es=(nm != "lnFXfwd"))
        res[f"{nm}_adjacent"] = run(rows, ykey, nm, "adjacent", with_es=False)
        res[f"{nm}_local"] = run(rows, ykey, nm, "local", with_es=False)

    # placebo: fake g=2019 on the 2024 cohort, years <= 2023
    print("-- placebo --")
    sub = [dict(r) for r in rows if r["g"] in (2024.0, np.inf) and r["year"] <= 2023]
    for r in sub:
        r["g"] = 2019.0 if r["g"] == 2024.0 else np.inf
    res["placebo"] = {}
    for ykey, nm in (("ha_bin", "HA"), ("deriv_use", "DerivUse"), ("lfx", "lnFXfwd")):
        M, coh, x, firms = to_matrix(sub, ykey)
        a = cs_matrix(M, coh, x, "notyet")
        se, ci, nb = boot_matrix(M, coh, x, "notyet", 999)
        res["placebo"][nm] = dict(att=a, se=se, ci95=ci, n=int(np.isfinite(M).sum()))
        print(f"   {nm:9s}: {a:+.4f} CI [{ci[0]:+.3f},{ci[1]:+.3f}]")

    # E-pillar: env-info mandate 2022 on the >=2tn cohort, base 2021
    print("-- E-pillar (env-info 2022, >=2tn) --")
    sub = [dict(r) for r in rows if r["g"] in (2019.0, np.inf) and r["year"] >= 2019]
    for r in sub:
        r["g"] = 2022.0 if r["g"] == 2019.0 else np.inf
    res["epillar"] = {}
    for ykey, nm in (("ha_bin", "HA"), ("deriv_use", "DerivUse"), ("lfx", "lnFXfwd")):
        M, coh, x, firms = to_matrix(sub, ykey)
        a = cs_matrix(M, coh, x, "notyet")
        se, ci, nb = boot_matrix(M, coh, x, "notyet", 999)
        res["epillar"][nm] = dict(att=a, se=se, ci95=ci, n=int(np.isfinite(M).sum()))
        print(f"   {nm:9s}: {a:+.4f} CI [{ci[0]:+.3f},{ci[1]:+.3f}]")

    # H3 carbon split
    print("-- H3 --")
    res["H3"] = {}
    for flag, nm in ((1, "carbon"), (0, "noncarbon")):
        sub = [r for r in rows if r["carbon"] == flag]
        M, coh, x, firms = to_matrix(sub, "ha_bin")
        a = cs_matrix(M, coh, x, "notyet")
        se, ci, nb = boot_matrix(M, coh, x, "notyet", 999)
        res["H3"][nm] = dict(att=a, se=se, ci95=ci, n=int(np.isfinite(M).sum()))
        print(f"   {nm:10s}: {a:+.4f} CI [{ci[0]:+.3f},{ci[1]:+.3f}]")

    # adoption descriptives
    desc = {}
    for y in YEARS:
        tr = [r["ha_bin"] for r in rows if r["year"] == y and r["g"] != np.inf
              and r["year"] >= r["g"] and r["ha_bin"] is not None]
        ct = [r["ha_bin"] for r in rows if r["year"] == y and
              (r["g"] == np.inf or r["year"] < r["g"]) and r["ha_bin"] is not None]
        desc[str(y)] = dict(treated=(float(np.mean(tr)) if tr else None, len(tr)),
                            control=(float(np.mean(ct)) if ct else None, len(ct)))
    res["adoption_by_year"] = desc

    json.dump(res, open("results_real.json", "w"), indent=1)
    print(f"\nwrote results_real.json  ({time.time()-t0:.1f}s)")


if __name__ == "__main__":
    main()
