"""
analyze_panel_v2.py — upgraded estimation for the expanded panel.

Upgrades over v1:
  1. SIZE-TREND-ADJUSTED group-time ATT (outcome-regression flavour of
     Sant'Anna-Zhao): within each (g,t) cell, fit dY ~ ln(assets2018) on
     the not-yet-treated controls, subtract the predicted trend from each
     treated firm's dY. The v1 placebo revealed size-correlated trends;
     this converts that caveat into a correction. The placebo is re-run
     under the same adjustment (target: ~0).
  2. E-PILLAR LAYERED TEST: environmental-information disclosure mandate
     (Environmental Technology & Industry Support Act) binding KOSPI
     issuers with assets >= KRW 2tn from 2022 -- an exposure-adjacent
     (E-pillar) mandate. Treated: the 2019 cohort (>=2tn) post-2022,
     base 2021; controls: never-treated. Identifies the E-mandate
     increment on top of their standing G-mandate.
  3. Firm-cluster bootstrap CIs for every reported number, incl. H3 and
     pointwise event-study bands.
Outputs: results_real.json (superset of v1 schema).
"""
import json
import numpy as np
from esg_hedge_engine import twfe_att, event_study
from analyze_panel import load_panel, CARBON_KSIC

RNG = np.random.default_rng(20260714)
MIN_CELL = 3


NEXT_COHORT = {2019.0: 2022.0, 2022.0: 2024.0, 2024.0: np.inf}


def cs_or(Y, fid, tid, coh, x, times, rule="notyet"):
    """
    Group-time ATT. Comparison set per `rule`:
      notyet   : all firms not yet treated at t (Callaway-Sant'Anna default)
      adjacent : only the NEXT cohort (closest in size; assets ratio <~2x),
                 usable while that cohort is itself still untreated.
      local    : threshold-local -- treated restricted to assets2018 within
                 [cut, 1.5*cut] and controls within [cut/1.5, cut) of the
                 cohort's own cutoff (RD-flavoured DiD; 2019 cut = 2tn etc.)
    Assignment is by size threshold with NO covariate overlap across bands,
    so regression adjustment on assets would be pure extrapolation --
    comparability is achieved by comparison-set design instead.
    """
    CUT = {2019.0: 2e12, 2022.0: 1e12, 2024.0: 5e11}
    firms = np.unique(fid)
    val, xv, fcoh = {}, {}, {}
    for f, t, y, c, xx in zip(fid, tid, Y, coh, x):
        val[(f, t)] = y
        xv[f] = xx
        fcoh[f] = c
    cohorts = sorted({c for c in fcoh.values() if c < np.inf})
    att_gt, w_gt = [], []
    for g in cohorts:
        base = g-1
        if base < times.min():
            continue
        gf = [f for f in firms if fcoh[f] == g]
        if rule == "local":
            cut = CUT[g]
            gf = [f for f in gf if cut <= xv[f] <= 1.5*cut]
        for t in times:
            if t < g:
                continue
            if rule == "adjacent":
                nxt = NEXT_COHORT[g]
                if nxt != np.inf and t >= nxt:      # neighbour now treated
                    continue
                cf = [f for f in firms if fcoh[f] == nxt]
            elif rule == "local":
                cut = CUT[g]
                nxt = NEXT_COHORT[g]
                cf = [f for f in firms if fcoh[f] == nxt and cut/1.5 <= xv[f] < cut
                      and (nxt == np.inf or t < nxt)]
            else:
                cf = [f for f in firms if fcoh[f] > t]
            dg = [val[(f, t)]-val[(f, base)] for f in gf
                  if (f, t) in val and (f, base) in val]
            dc = [val[(f, t)]-val[(f, base)] for f in cf
                  if (f, t) in val and (f, base) in val]
            if len(dg) < MIN_CELL or len(dc) < MIN_CELL:
                continue
            att_gt.append(np.mean(dg) - np.mean(dc))
            w_gt.append(len(dg))
    if not att_gt:
        return float("nan")
    att_gt = np.array(att_gt); w_gt = np.array(w_gt, float)
    return float(np.sum(att_gt*w_gt)/np.sum(w_gt))


def boot(fun, fid, nboot=399):
    """pairs-cluster bootstrap over firms; fun(idx array)->stat"""
    firms = np.unique(fid)
    idx_by = {f: np.where(fid == f)[0] for f in firms}
    out = []
    for _ in range(nboot):
        pick = RNG.choice(firms, size=len(firms), replace=True)
        idx = np.concatenate([idx_by[f] for f in pick])
        rel = np.concatenate([[f"{f}#{i}"]*len(idx_by[f]) for i, f in enumerate(pick)])
        s = fun(idx, rel)
        if np.isfinite(s):
            out.append(s)
    out = np.array(out)
    if len(out) < 20:               # estimator infeasible on most resamples
        return float("nan"), (float("nan"), float("nan")), len(out)
    return float(out.std(ddof=1)), (float(np.percentile(out, 2.5)),
                                    float(np.percentile(out, 97.5))), len(out)


def prep(rows, ykey):
    R = [r for r in rows if r.get(ykey) is not None and r.get("assets2018")]
    Y = np.array([float(r[ykey]) for r in R])
    fid = np.array([r["corp_code"] for r in R])
    tid = np.array([r["year"] for r in R])
    coh = np.array([r["g"] for r in R], dtype=float)
    x = np.array([np.log(r["assets2018"]) for r in R])
    evt = np.array([r["evt"] for r in R])
    return R, Y, fid, tid, coh, x, evt


def run(rows, ykey, label, rule="notyet", with_es=True):
    R, Y, fid, tid, coh, x, evt = prep(rows, ykey)
    times = np.sort(np.unique(tid))
    att = cs_or(Y, fid, tid, coh, x, times, rule=rule)
    se, ci, nb = boot(lambda idx, rel: cs_or(Y[idx], rel, tid[idx], coh[idx], x[idx], times, rule=rule), fid)
    beta = twfe_att(Y, fid, tid, (evt >= 0).astype(float)*(coh < np.inf))
    es = event_study(Y, fid, tid, evt, leads=3, lags=4) if with_es else {}
    ntr = len({f for f, c in zip(fid, coh) if c < np.inf})
    print(f"== {label:38s} N={len(R)} ({len(np.unique(fid))} firms, {ntr} treated)")
    print(f"   ATT={att:+.4f}  SE {se:.4f}  CI [{ci[0]:+.4f},{ci[1]:+.4f}]  TWFE {beta:+.4f}")
    out = dict(outcome=label, n_obs=len(R), n_firms=int(len(np.unique(fid))),
               n_treated_firms=ntr, mean_dep=float(Y.mean()), att_cs=att,
               se_boot=se, ci95=ci, nboot=nb, beta_twfe=float(beta), rule=rule)
    if with_es:
        out["event_study"] = {str(k): v for k, v in es.items()}
        out["pretrend_max_abs"] = float(max(abs(v) for k, v in es.items() if k < -1))
    return out


def main():
    rows = load_panel()
    res = {"panel_size": len(rows),
           "parse_errors": sum(1 for r in rows if r.get("error"))}
    print(f"panel: {len(rows)} firm-years, {len({r['corp_code'] for r in rows})} firms")

    # -- primary outcomes under three comparison-set rules -------------------
    for r in rows:
        r["lfx"] = None if r.get("fxfwd") is None else float(np.log1p(r["fxfwd"]))
    for ykey, nm in (("ha_bin", "HA"), ("deriv_use", "DerivUse"), ("lfx", "lnFXfwd")):
        res[nm] = run(rows, ykey, f"{nm} (not-yet controls)", rule="notyet")
        res[f"{nm}_adjacent"] = run(rows, ykey, f"{nm} (adjacent cohort)", rule="adjacent", with_es=False)
        res[f"{nm}_local"] = run(rows, ykey, f"{nm} (threshold-local)", rule="local", with_es=False)

    # -- placebo under notyet & adjacent rules --------------------------------
    print("\n-- placebo: fake g=2019 on the 2024 cohort, years<=2023 --")
    sub = [dict(r) for r in rows if r["g"] in (2024.0, np.inf) and r["year"] <= 2023]
    for r in sub:
        r["g"] = 2019.0 if r["g"] == 2024.0 else np.inf
        r["lfx"] = None if r.get("fxfwd") is None else float(np.log1p(r["fxfwd"]))
    res["placebo"] = {}
    for ykey, nm in (("ha_bin", "HA"), ("deriv_use", "DerivUse"), ("lfx", "lnFXfwd")):
        R, Y, fid, tid, coh, x, evt = prep(sub, ykey)
        times = np.sort(np.unique(tid))
        a = cs_or(Y, fid, tid, coh, x, times, rule="notyet")
        se, ci, nb = boot(lambda idx, rel: cs_or(Y[idx], rel, tid[idx], coh[idx], x[idx], times, rule="notyet"), fid, nboot=199)
        res["placebo"][nm] = dict(att=a, se=se, ci95=ci, n=len(R))
        print(f"   {nm:9s}: {a:+.4f}  CI [{ci[0]:+.3f},{ci[1]:+.3f}]")

    # -- E-pillar layered test (env-info mandate 2022, >=2tn cohort) ---------
    print("\n-- E-pillar: env-info mandate 2022 on the >=2tn cohort --")
    sub = [dict(r) for r in rows if r["g"] in (2019.0, np.inf) and r["year"] >= 2019]
    for r in sub:
        r["g"] = 2022.0 if r["g"] == 2019.0 else np.inf   # E-mandate event
        r["lfx"] = None if r.get("fxfwd") is None else float(np.log1p(r["fxfwd"]))
    res["epillar"] = {}
    for ykey, nm in (("ha_bin", "HA"), ("deriv_use", "DerivUse"), ("lfx", "lnFXfwd")):
        R, Y, fid, tid, coh, x, evt = prep(sub, ykey)
        times = np.sort(np.unique(tid))
        a = cs_or(Y, fid, tid, coh, x, times, rule="notyet")
        se, ci, nb = boot(lambda idx, rel: cs_or(Y[idx], rel, tid[idx], coh[idx], x[idx], times, rule="notyet"), fid, nboot=199)
        res["epillar"][nm] = dict(att=a, se=se, ci95=ci, n=len(R))
        print(f"   {nm:9s}: ATT={a:+.4f}  CI [{ci[0]:+.3f},{ci[1]:+.3f}]  N={len(R)}")

    # -- H3 with bootstrap CI --------------------------------------------------
    print("\n-- H3 carbon split (not-yet controls) --")
    res["H3"] = {}
    for flag, nm in ((1, "carbon"), (0, "noncarbon")):
        sub = [r for r in rows if r["carbon"] == flag]
        R, Y, fid, tid, coh, x, evt = prep(sub, "ha_bin")
        times = np.sort(np.unique(tid))
        a = cs_or(Y, fid, tid, coh, x, times, rule="notyet")
        se, ci, nb = boot(lambda idx, rel: cs_or(Y[idx], rel, tid[idx], coh[idx], x[idx], times, rule="notyet"), fid, nboot=199)
        res["H3"][nm] = dict(att=a, se=se, ci95=ci, n=len(R))
        print(f"   {nm:10s}: ATT={a:+.4f}  CI [{ci[0]:+.3f},{ci[1]:+.3f}]  N={len(R)}")

    # -- event-study bands (adjusted primary outcomes) -----------------------
    for key, ykey in (("HA", "ha_bin"), ("DerivUse", "deriv_use")):
        R, Y, fid, tid, coh, x, evt = prep(rows, ykey)
        bands = {}
        firms = np.unique(fid)
        idx_by = {f: np.where(fid == f)[0] for f in firms}
        draws = {k: [] for k in res[key]["event_study"]}
        for _ in range(199):
            pick = RNG.choice(firms, size=len(firms), replace=True)
            idx = np.concatenate([idx_by[f] for f in pick])
            rel = np.concatenate([[f"{f}#{i}"]*len(idx_by[f]) for i, f in enumerate(pick)])
            try:
                esb = event_study(Y[idx], rel, tid[idx], evt[idx], leads=3, lags=4)
                for k, v in esb.items():
                    if str(k) in draws:
                        draws[str(k)].append(v)
            except Exception:
                pass
        for k, v in draws.items():
            if v:
                bands[k] = [float(np.percentile(v, 2.5)), float(np.percentile(v, 97.5))]
        res[key]["event_bands"] = bands

    # descriptives
    desc = {}
    for y in sorted({r["year"] for r in rows}):
        tr = [r["ha_bin"] for r in rows if r["year"] == y and r["evt"] is not None
              and r["g"] != np.inf and r["year"] >= r["g"] and r["ha_bin"] is not None]
        ct = [r["ha_bin"] for r in rows if r["year"] == y and
              (r["g"] == np.inf or r["year"] < r["g"]) and r["ha_bin"] is not None]
        desc[str(y)] = dict(treated=(float(np.mean(tr)) if tr else None, len(tr)),
                            control=(float(np.mean(ct)) if ct else None, len(ct)))
    res["adoption_by_year"] = desc

    json.dump(res, open("results_real.json", "w"), indent=1)
    print("\nwrote results_real.json")


if __name__ == "__main__":
    main()
