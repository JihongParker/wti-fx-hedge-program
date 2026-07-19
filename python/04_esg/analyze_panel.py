"""
analyze_panel.py — Stage 3: the real estimations.
Treatment: Korea's staggered mandatory corporate-governance-report
disclosure (the G pillar of ESG): cohorts g=2019 (assets>=2tn, FY2018),
g=2022 (>=1tn), g=2024 (>=0.5tn); assets<0.5tn not yet treated in-sample.
Assignment by FY2018 assets (intention-to-treat). D_it = 1{year >= g}.

Outcomes:
  HA        hedge-accounting designation dummy (note-sentence parser)
  DerivUse  any-derivative-instrument dummy

Estimators (identical code paths to the validated engine):
  * Callaway-Sant'Anna group-time ATT, not-yet-treated controls (primary)
  * event study (leads/lags, ref = -1)
  * TWFE benchmark
Inference: pairs-cluster bootstrap over FIRMS (399 reps) for the CS ATT.
Heterogeneity: carbon-intensive KSIC industries (H3).
Output: results_real.json + console table.
"""
import json
import numpy as np
from esg_hedge_engine import twfe_att, event_study, callaway_santanna, twoway_within

RNG = np.random.default_rng(20260713)
CARBON_KSIC = {"19", "20", "23", "24", "35", "50", "51"}   # refining, chem, minerals, metals, power, shipping, air


def load_panel():
    rows = [json.loads(l) for l in open("panel.jsonl")]
    # deduplicate (resume runs may duplicate): keep last per (firm, year)
    dd = {}
    for r in rows:
        dd[(r["corp_code"], r["year"])] = r
    rows = list(dd.values())
    for r in rows:
        g = r["cohort"]
        r["g"] = np.inf if g == "inf" else int(g)
        r["D"] = int(r["g"] != np.inf and r["year"] >= r["g"])
        r["evt"] = (r["year"] - r["g"]) if r["g"] != np.inf else -99
        r["carbon"] = int((r.get("induty") or "")[:2] in CARBON_KSIC)
        # Outcome coding (stated in the paper):
        #   HA=1  designation affirmatively found in the note sentence
        #   HA=0  operative negative found, OR no derivative instrument
        #         mentioned at all (no instrument -> no designation possible)
        #   HA=missing  derivatives present but designation language ambiguous
        if r.get("error"):
            r["ha_bin"] = None
        elif r.get("ha") is True:
            r["ha_bin"] = 1
        elif r.get("ha") is False or r.get("deriv_use") == 0:
            r["ha_bin"] = 0
        else:
            r["ha_bin"] = None
    return rows


def arrs(rows, ykey):
    R = [r for r in rows if r.get(ykey) is not None]
    fid = np.array([r["corp_code"] for r in R])
    tid = np.array([r["year"] for r in R])
    Y = np.array([float(r[ykey]) for r in R])
    coh = np.array([r["g"] for r in R], dtype=float)
    evt = np.array([r["evt"] for r in R])
    return R, Y, fid, tid, coh, evt


def cs_with_bootstrap(Y, fid, tid, coh, nboot=399):
    times = np.sort(np.unique(tid))
    att, _ = callaway_santanna(Y, fid, tid, coh, times)
    firms = np.unique(fid)
    idx_by_firm = {f: np.where(fid == f)[0] for f in firms}
    boots = []
    for _ in range(nboot):
        pick = RNG.choice(firms, size=len(firms), replace=True)
        idx = np.concatenate([idx_by_firm[f] for f in pick])
        # relabel resampled duplicate firms uniquely
        fid_b = np.concatenate([[f"{f}_{i}"]*len(idx_by_firm[f]) for i, f in enumerate(pick)])
        try:
            a, _ = callaway_santanna(Y[idx], fid_b, tid[idx], coh[idx], times)
            if np.isfinite(a):
                boots.append(a)
        except Exception:
            pass
    boots = np.array(boots)
    return att, float(boots.std(ddof=1)), (float(np.percentile(boots, 2.5)),
                                           float(np.percentile(boots, 97.5))), len(boots)


def run_outcome(rows, ykey, label):
    R, Y, fid, tid, coh, evt = arrs(rows, ykey)
    n_treat = len({r['corp_code'] for r in R if r['g'] != np.inf})
    att, se, ci, nb = cs_with_bootstrap(Y, fid, tid, coh)
    beta = twfe_att(Y, fid, tid, np.array([r["D"] for r in R]))
    es = event_study(Y, fid, tid, evt, leads=3, lags=4)
    pre = {k: v for k, v in es.items() if k < -1}
    out = dict(outcome=label, n_obs=len(R), n_firms=len(np.unique(fid)),
               n_treated_firms=n_treat, mean_dep=float(Y.mean()),
               att_cs=float(att), se_boot=se, ci95=ci, nboot=nb,
               beta_twfe=float(beta), event_study={str(k): v for k, v in es.items()},
               pretrend_max_abs=float(max(abs(v) for v in pre.values())))
    print(f"\n== {label} ==  N={out['n_obs']} firm-years, {out['n_firms']} firms "
          f"(treated {n_treat}), mean={out['mean_dep']:.3f}")
    print(f"  CS ATT   = {att:+.4f}  (boot SE {se:.4f}, 95% CI [{ci[0]:+.4f},{ci[1]:+.4f}], B={nb})")
    print(f"  TWFE     = {beta:+.4f}")
    print(f"  max|pre| = {out['pretrend_max_abs']:.4f}")
    return out


def main():
    rows = load_panel()
    ok = [r for r in rows if "error" not in r or r.get("ha_bin") is not None]
    n_err = sum(1 for r in rows if r.get("error"))
    print(f"panel: {len(rows)} firm-years, parse errors/missing report: {n_err}")
    res = {"panel_size": len(rows), "parse_errors": n_err}

    res["HA"] = run_outcome(rows, "ha_bin", "HedgeAccountingDummy")
    res["DerivUse"] = run_outcome(rows, "deriv_use", "DerivUse")

    # H3: carbon-intensive split (CS ATT within subsample)
    for flag, nm in ((1, "carbon"), (0, "noncarbon")):
        sub = [r for r in rows if r["carbon"] == flag]
        R, Y, fid, tid, coh, _ = arrs(sub, "ha_bin")
        try:
            a, _ = callaway_santanna(Y, fid, tid, coh, np.sort(np.unique(tid)))
        except Exception:
            a = float("nan")
        res[f"HA_{nm}"] = dict(att=float(a), n=len(R))
        print(f"  H3 {nm:10s}: ATT={a:+.4f} (N={len(R)})")

    # intensive margin: textual salience of FX-forward disclosure
    for r in rows:
        r["lfx"] = None if r.get("fxfwd") is None else float(np.log1p(r["fxfwd"]))
    res["lnFXfwd"] = run_outcome(rows, "lfx", "ln(1+FXforward mentions)")

    # placebo: assign fake g=2019 to the 2024 cohort, never-treated controls,
    # years <= 2023 (their real mandate starts 2024, so true effect = 0).
    # A nonzero placebo measures size-correlated differential trends.
    sub = [dict(r) for r in rows if r["g"] in (2024.0, np.inf) and r["year"] <= 2023]
    for r in sub:
        r["g"] = 2019.0 if r["g"] == 2024.0 else np.inf
    res["placebo"] = {}
    for ykey, nm in (("deriv_use", "DerivUse"), ("ha_bin", "HA"), ("lfx", "lnFXfwd")):
        R = [r for r in sub if r.get(ykey) is not None]
        Y = np.array([float(r[ykey]) for r in R])
        fid = np.array([r["corp_code"] for r in R]); tid = np.array([r["year"] for r in R])
        coh = np.array([r["g"] for r in R], dtype=float)
        att, se, ci, nb = cs_with_bootstrap(Y, fid, tid, coh, nboot=199)
        res["placebo"][nm] = dict(att=float(att), se=se, ci95=ci, n=len(R))
        print(f"  placebo {nm:9s}: ATT={att:+.4f} (SE {se:.4f}, CI [{ci[0]:+.3f},{ci[1]:+.3f}], N={len(R)})")

    # descriptive: HA adoption rate by year, treated vs not-yet
    desc = {}
    for y in range(2017, 2025):
        tr = [r["ha_bin"] for r in rows if r["year"] == y and r["D"] == 1 and r["ha_bin"] is not None]
        ct = [r["ha_bin"] for r in rows if r["year"] == y and r["D"] == 0 and r["ha_bin"] is not None]
        desc[y] = dict(treated=(float(np.mean(tr)) if tr else None, len(tr)),
                       control=(float(np.mean(ct)) if ct else None, len(ct)))
    res["adoption_by_year"] = {str(k): v for k, v in desc.items()}

    json.dump(res, open("results_real.json", "w"), indent=1)
    print("\nwrote results_real.json")


if __name__ == "__main__":
    main()
