"""
esg_hedge_engine.py
===================================================================
Engine for "Does Mandatory ESG Disclosure Change Corporate Hedging?"
(Park, 2026d -- the fourth paper of the program).

This engine does NOT touch OpenDART/XBRL (that is the paper's stated
"remaining 30%", which needs FSS API credentials and per-filing footnote
parsing). What it DOES do, in the verification spirit of the companion
trilogy -- where every optimum is certified by two independent solvers --
is validate the INTERNAL LOGIC of the research design on model-generated
data:

  PART 1  Structural solver.
          Solve the firm's joint hedging/disclosure program (Eq. 8) two
          independent ways -- the closed-form fixed point (Eqs. 14-16)
          and a direct multi-start numerical minimization -- and require
          them to agree. Then verify every comparative-statics SIGN in
          Table 2 (dh_f/dl>0, dd/dl>0, dh_f/da>0, dh_f/drho per Eq.20)
          by re-solving the full equilibrium under parameter shocks.

  PART 2  Design Monte Carlo.
          Generate a synthetic staggered-adoption panel whose data-
          generating process IS the structural model (a mandate raises
          lambda for treated firms at cohort-specific dates), then run
          the paper's estimators -- TWFE, event study, and a Callaway-
          Sant'Anna group-time ATT with not-yet-treated controls -- and
          check that they recover the model's predicted signs H1-H4 and
          the H(sub) substitution result, including the capability-
          confound attenuation the paper flags for H(sub).

Outputs: results JSON, a plain-text results log, and three PDF/PNG
figures under ./figures.

Self-contained: numpy/scipy/pandas/matplotlib only (no statsmodels);
the DiD estimators are implemented directly for transparency.
"""

import json
import numpy as np
from dataclasses import dataclass, asdict
from scipy.optimize import brentq, minimize
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

RNG = np.random.default_rng(20260713)

# =====================================================================
# PART 1 : STRUCTURAL MODEL
# =====================================================================

@dataclass
class Params:
    sigma_f: float = 0.30     # financial-exposure vol (WTI-ish diffusive)
    sigma_c: float = 0.15     # climate-exposure vol
    rho:     float = 0.20     # financial/climate correlation (positive)
    p_f:     float = 0.030    # per-unit financial hedge premium
    p_c:     float = 0.020    # per-unit climate hedge premium
    phi:     float = 0.15     # baseline residual-risk price (distress cost)
    lam:     float = 0.60     # regulatory-penalty stringency  (lambda)
    k:       float = 1.00     # disclosure-attenuation curvature
    a:       float = 0.05     # disclosure-cost curvature

    def Sigma(self):
        s = np.array([[self.sigma_f**2, self.rho*self.sigma_f*self.sigma_c],
                      [self.rho*self.sigma_f*self.sigma_c, self.sigma_c**2]])
        return s

    def p(self):
        return np.array([self.p_f, self.p_c])


def Lambda(pr: Params, d: float) -> float:
    """Effective marginal price of residual risk, Eq. (10)."""
    return pr.phi + pr.lam*np.exp(-pr.k*d)


def hedges_given_d(pr: Params, d: float):
    """Closed-form hedge ratios, Eqs. (14)-(15): h = 1 - (1/2Lambda) Sigma^{-1} p."""
    Sinv = np.linalg.inv(pr.Sigma())
    u = Sinv.dot(pr.p()) / (2.0*Lambda(pr, d))     # unhedged fractions
    h = 1.0 - u
    return h, u


def kappa(pr: Params) -> float:
    """kappa = p' Sigma^{-1} p  (Eq. 17)."""
    Sinv = np.linalg.inv(pr.Sigma())
    return float(pr.p().dot(Sinv).dot(pr.p()))


def disclosure_focus(pr: Params, d: float) -> float:
    """
    Fixed-point residual for d*, Eq. (16):
        8 a d (phi + lam e^{-kd})^2 - k lam kappa e^{-kd} = 0.
    """
    L = Lambda(pr, d)
    return 8.0*pr.a*d*L*L - pr.k*pr.lam*kappa(pr)*np.exp(-pr.k*d)


def solve_closed_form(pr: Params):
    """
    METHOD 1 -- closed form + scalar fixed point.
    Root-bracket d* on [0, d_hi]; LHS increasing/unbounded, RHS positive/
    decreasing, so a unique interior root exists whenever lam>0 (Sec. 3.5).
    """
    f0 = disclosure_focus(pr, 0.0)          # = -k lam kappa < 0
    d_hi = 1.0
    while disclosure_focus(pr, d_hi) < 0 and d_hi < 1e6:
        d_hi *= 2.0
    d_star = brentq(lambda d: disclosure_focus(pr, d), 0.0, d_hi, xtol=1e-14, rtol=1e-15)
    h_star, u_star = hedges_given_d(pr, d_star)
    R_star = float(u_star.dot(pr.Sigma()).dot(u_star))
    return dict(method="closed_form", d=d_star, h_f=h_star[0], h_c=h_star[1],
                R=R_star, Lambda=Lambda(pr, d_star))


def objective(pr: Params, x):
    """Total cost Pi(h_f,h_c,d), Eq. (8)."""
    h_f, h_c, d = x
    u = np.array([1.0-h_f, 1.0-h_c])
    R = float(u.dot(pr.Sigma()).dot(u))
    return pr.p_f*h_f + pr.p_c*h_c + pr.a*d*d + Lambda(pr, d)*R


def solve_numeric(pr: Params, n_starts=48):
    """
    METHOD 2 -- direct multi-start box-constrained minimization of Pi.
    Independent of the closed form; used to certify Method 1.
    """
    best = None
    bounds = [(0.0, 1.0), (0.0, 1.0), (0.0, 20.0)]
    for _ in range(n_starts):
        x0 = np.array([RNG.uniform(0, 1), RNG.uniform(0, 1), RNG.uniform(0, 5)])
        res = minimize(lambda x: objective(pr, x), x0, method="L-BFGS-B",
                       bounds=bounds, options=dict(ftol=1e-15, gtol=1e-12, maxiter=5000))
        if best is None or res.fun < best.fun:
            best = res
    h_f, h_c, d = best.x
    u = np.array([1.0-h_f, 1.0-h_c])
    R = float(u.dot(pr.Sigma()).dot(u))
    return dict(method="numeric", d=d, h_f=h_f, h_c=h_c, R=R,
                Lambda=Lambda(pr, d), Pi=best.fun)


def comparative_statics(pr: Params, eps=1e-4):
    """
    Verify the SIGNS in Table 2 by re-solving the full equilibrium
    (closed-form incl. the d* feedback) under +/- shocks. Returns the
    total derivatives d(h_f*)/d(theta) and d(d*)/d(theta).
    """
    def bump(**kw):
        p2 = Params(**{**asdict(pr), **kw})
        return solve_closed_form(p2)

    out = {}
    for name, attr, base in [("lambda", "lam", pr.lam),
                             ("a", "a", pr.a),
                             ("rho", "rho", pr.rho),
                             ("p_f", "p_f", pr.p_f),
                             ("sigma_f", "sigma_f", pr.sigma_f)]:
        hp = bump(**{attr: base*(1+eps) if base != 0 else eps})
        hm = bump(**{attr: base*(1-eps) if base != 0 else -eps})
        step = (base*(1+eps) - base*(1-eps)) if base != 0 else 2*eps
        out[name] = dict(
            dhf=(hp["h_f"]-hm["h_f"])/step,
            dd =(hp["d"]  -hm["d"]  )/step,
        )
    # analytic sign for dh_f/drho, Eq. (20): >0 iff 2 rho p_f/sf^2 < (1+rho^2) p_c/(sf sc)
    lhs = 2*pr.rho*pr.p_f/pr.sigma_f**2
    rhs = (1+pr.rho**2)*pr.p_c/(pr.sigma_f*pr.sigma_c)
    out["rho"]["analytic_sign_positive"] = bool(lhs < rhs)
    # derived boundary for the disclosure/stringency total derivative:
    #   sign(dd*/dlambda) = sign(phi - lambda e^{-k d*})   (Appendix, corrected)
    dstar = solve_closed_form(pr)["d"]
    marg_penalty = pr.lam*np.exp(-pr.k*dstar)
    out["lambda"]["dd_condition_phi_minus_lam_emkd"] = float(pr.phi - marg_penalty)
    out["lambda"]["dd_condition_predicts_positive"] = bool(pr.phi > marg_penalty)
    return out


def run_part1():
    pr = Params()
    m1 = solve_closed_form(pr)
    m2 = solve_numeric(pr)
    agree = dict(
        d   =abs(m1["d"]  -m2["d"]),
        h_f =abs(m1["h_f"]-m2["h_f"]),
        h_c =abs(m1["h_c"]-m2["h_c"]),
    )
    cs = comparative_statics(pr)
    # sign checks vs the paper's (corrected) Table 2 claims
    sign_checks = {
        "dhf_dlambda>0": cs["lambda"]["dhf"] > 0,                 # H1, robust
        "dd_dlambda_matches_boundary":                            # sign = sign(phi - lam e^{-kd})
            (cs["lambda"]["dd"] > 0) == cs["lambda"]["dd_condition_predicts_positive"],
        "dhf_da>0":      cs["a"]["dhf"]      > 0,   # H(sub) substitution (level)
        "dd_da<0":       cs["a"]["dd"]       < 0,   # costlier disclosure -> less disclosure
        "dhf_dpf<0":     cs["p_f"]["dhf"]    < 0,
        "dhf_dsigmaf>0": cs["sigma_f"]["dhf"] > 0,
        "dhf_drho_matches_analytic":
            (cs["rho"]["dhf"] > 0) == cs["rho"]["analytic_sign_positive"],
    }
    return dict(params=asdict(pr), method1=m1, method2=m2,
                solver_agreement=agree, comparative_statics=cs,
                sign_checks=sign_checks,
                all_signs_pass=all(sign_checks.values()),
                solvers_agree=max(agree.values()) < 1e-6)


# =====================================================================
# PART 2 : STAGGERED-ADOPTION DiD MONTE CARLO
# =====================================================================

def h_f_star(pr: Params) -> float:
    return solve_closed_form(pr)["h_f"]

def d_star(pr: Params) -> float:
    return solve_closed_form(pr)["d"]


def twoway_within(y, fid, tid):
    """Two-way (firm & time) within transform: y - ybar_i - ybar_t + ybar."""
    y = np.asarray(y, float)
    yb = y.mean()
    fi = {f: y[fid == f].mean() for f in np.unique(fid)}
    ti = {t: y[tid == t].mean() for t in np.unique(tid)}
    return y - np.array([fi[f] for f in fid]) - np.array([ti[t] for t in tid]) + yb


def twfe_att(Y, fid, tid, D):
    """Static TWFE DiD: regress within-Y on within-D."""
    yt = twoway_within(Y, fid, tid)
    dt = twoway_within(D.astype(float), fid, tid)
    beta = float(dt.dot(yt) / dt.dot(dt))
    return beta


def event_study(Y, fid, tid, evt, leads=3, lags=4):
    """Two-way FE event study; returns {event_time: coef}, normalized at -1."""
    ev = np.clip(evt, -leads, lags)
    cols, names = [], []
    for l in range(-leads, lags+1):
        if l == -1:      # reference
            continue
        cols.append((ev == l).astype(float))
        names.append(l)
    X = np.column_stack([twoway_within(c, fid, tid) for c in cols])
    yt = twoway_within(Y, fid, tid)
    coef, *_ = np.linalg.lstsq(X, yt, rcond=None)
    return dict(zip(names, coef.tolist()))


def callaway_santanna(Y, fid, tid, cohort, times):
    """
    Group-time ATT with NOT-YET-TREATED controls (Callaway-Sant'Anna 2021),
    aggregated to an overall post-treatment ATT weighted by cohort size.
      ATT(g,t) = [Ybar_{g,t} - Ybar_{g,g-1}]  -  [Ybar_{C,t} - Ybar_{C,g-1}]
    with control set C = firms not yet treated by t (cohort > t).
    """
    firms = np.unique(fid)
    # per firm-time value lookup
    val = {(f, t): Y[(fid == f) & (tid == t)][0] for f in firms for t in times
           if ((fid == f) & (tid == t)).any()}
    fcoh = {f: cohort[fid == f][0] for f in firms}
    treated_cohorts = sorted({c for c in fcoh.values() if c < np.inf})
    att_gt, weights, gt = [], [], []
    for g in treated_cohorts:
        base = g-1
        if base < times.min():
            continue
        g_firms = [f for f in firms if fcoh[f] == g]
        for t in times:
            if t < g:
                continue
            ctrl = [f for f in firms if fcoh[f] > t]     # not yet treated by t
            if not ctrl or not g_firms:
                continue
            try:
                dY_g = np.mean([val[(f, t)]-val[(f, base)] for f in g_firms])
                dY_c = np.mean([val[(f, t)]-val[(f, base)] for f in ctrl])
            except KeyError:
                continue
            att_gt.append(dY_g - dY_c)
            weights.append(len(g_firms))
            gt.append((g, t))
    att_gt = np.array(att_gt); weights = np.array(weights, float)
    overall = float(np.sum(att_gt*weights)/np.sum(weights))
    return overall, dict(zip([f"{g}_{t}" for g, t in gt], att_gt.tolist()))


RAMP = 3   # years over which the hedging response phases in (treatment dynamics)

def simulate_panel(N=900, T=10, confound=False, seed=1):
    """
    Build a staggered panel from the structural DGP.
      cohorts g in {4,6, inf}. A mandate raises stringency from lambda0 to
      lambda1 for treated firms after g; carbon-intensive firms experience a
      LARGER effective jump (their disclosed climate risk is more salient),
      which is the H3 channel. The hedging response phases in over RAMP years
      (treatment dynamics), so with staggered timing TWFE is biased and the
      heterogeneity-robust group-time ATT is the benchmark.
        Y_it = h_f-level(t) + firm FE + time FE + noise
        A_it (hedge-accounting adoption) is driven by the HEDGING LEVEL
             ("tail wags the dog", Glaum-Klocker), not by disclosure d.
      a_i (disclosure cost) is exogenous; the H(sub) test is cross-sectional
      (Sec. run_part2), where the capability confound -- high-a firms are also
      weak-risk-management firms -- biases the naive level estimate.
    """
    rng = np.random.default_rng(seed)
    base = Params()
    coh_choice = rng.choice([4, 6, np.inf], size=N, p=[0.30, 0.30, 0.40])
    lam0, lam1 = base.lam, 1.00

    carbon = rng.random(N) < 0.5
    # latent (inverse) sophistication s drives BOTH disclosure cost a and weak
    # risk-management capability, but they are correlated, not identical: a keeps
    # an independent component so the structural a->hedging channel is identified
    # once capability is controlled for.
    s = rng.normal(0, 1, N)                                   # high s = less sophisticated
    a_i = np.clip(base.a + 0.020*s + 0.015*rng.normal(0, 1, N), 0.02, 0.13)
    capability = s                                            # weak-capability index
    firmFE = rng.normal(0, 0.03, N)
    timeFE = rng.normal(0, 0.015, T)

    rows = []
    true_effects = []
    for i in range(N):
        g = coh_choice[i]
        lam1_i = lam1 + (0.30 if carbon[i] else 0.0)   # H3: bigger jump for carbon
        p0 = Params(**{**asdict(base), "a": a_i[i], "lam": lam0})
        p1 = Params(**{**asdict(base), "a": a_i[i], "lam": lam1_i})
        hf0, hf1 = h_f_star(p0), h_f_star(p1)
        cap_pen = (0.05*capability[i]) if confound else 0.0  # weak-capability firms hedge less
        for t in range(T):
            if g < np.inf and t >= g:
                phase = min(1.0, (t - g + 1)/RAMP)       # ramp-in of the effect
                eff = (hf1 - hf0)*phase
                true_effects.append(eff)
            else:
                eff = 0.0
            hf = hf0 + eff
            y = hf + firmFE[i] + timeFE[t] - cap_pen + rng.normal(0, 0.008)
            adopt_latent = 6.0*(hf - hf0) + firmFE[i] + rng.normal(0, 0.05)  # driven by hedging level
            evt = (t - g) if g < np.inf else -99
            rows.append((i, t, y, adopt_latent, int(g < np.inf and t >= g),
                         g if g < np.inf else np.inf, evt,
                         int(carbon[i]), a_i[i], capability[i]))
    import pandas as pd
    df = pd.DataFrame(rows, columns=["fid", "tid", "Y", "Alat", "D", "cohort",
                                     "evt", "carbon", "a", "cap"])
    df["A"] = (df["Alat"] > df["Alat"].median()).astype(int)
    true_att = float(np.mean(true_effects))
    return df, true_att


def ols(y, X):
    """Plain OLS with intercept; returns coefficient vector (const first)."""
    Xd = np.column_stack([np.ones(len(y)), X])
    beta, *_ = np.linalg.lstsq(Xd, y, rcond=None)
    return beta


def run_part2():
    out = {}

    # --- main run: treatment dynamics, no confound ----------------------
    df, true_att = simulate_panel(confound=False, seed=7)
    times = np.sort(df["tid"].unique())
    beta_twfe = twfe_att(df["Y"].values, df["fid"].values, df["tid"].values, df["D"].values)
    att_cs, _ = callaway_santanna(df["Y"].values, df["fid"].values, df["tid"].values,
                                  df["cohort"].values, times)
    es = event_study(df["Y"].values, df["fid"].values, df["tid"].values, df["evt"].values)

    # H2: DiD on hedge-accounting adoption A (driven by hedging level)
    beta_A = twfe_att(df["A"].values, df["fid"].values, df["tid"].values, df["D"].values)

    # H3: ATT within carbon vs non-carbon (heterogeneity-robust, per subgroup)
    def att_sub(mask):
        d = df[mask]
        a, _ = callaway_santanna(d["Y"].values, d["fid"].values, d["tid"].values,
                                 d["cohort"].values, np.sort(d["tid"].unique()))
        return a
    att_carbon = att_sub(df["carbon"] == 1)
    att_noncarbon = att_sub(df["carbon"] == 0)

    out["true_att"] = true_att
    out["beta_twfe"] = beta_twfe
    out["att_callaway_santanna"] = att_cs
    out["cs_minus_true"] = att_cs - true_att
    out["twfe_minus_true"] = beta_twfe - true_att
    out["twfe_bias_vs_cs"] = beta_twfe - att_cs
    out["event_study"] = es
    out["pretrend_max_abs"] = max(abs(v) for k, v in es.items() if k < 0)
    out["H1_hedging_att_positive"] = att_cs > 0
    out["H2_adoption_att_positive"] = beta_A > 0
    out["beta_adoption"] = beta_A
    out["H3_carbon_stronger"] = att_carbon > att_noncarbon
    out["att_carbon"] = att_carbon
    out["att_noncarbon"] = att_noncarbon

    # --- H(sub): CROSS-SECTIONAL test on treated firms' post-period hedging.
    #     The structural prediction dh_f/da>0 is a LEVEL effect (firm-FE-absorbed
    #     in a DiD), so it is tested cross-sectionally. The capability confound
    #     (high-a firms are weak-risk-management firms) biases the naive slope;
    #     controlling for capability recovers the structural sign.
    def hsub_cross_section(confound):
        d, _ = simulate_panel(confound=confound, seed=11)
        tr = d[(d["cohort"] < np.inf) & (d["tid"] >= d["cohort"])]
        firm = tr.groupby("fid").agg(Y=("Y", "mean"), a=("a", "first"),
                                     cap=("cap", "first")).reset_index()
        y = firm["Y"].values
        a = (firm["a"].values - firm["a"].values.mean())
        cap = firm["cap"].values
        b_naive = ols(y, a[:, None])[1]                      # a only
        b_ctrl  = ols(y, np.column_stack([a, cap]))[1]       # a controlling for capability
        return float(b_naive), float(b_ctrl)
    b_naive_clean, b_ctrl_clean = hsub_cross_section(confound=False)
    b_naive_conf,  b_ctrl_conf  = hsub_cross_section(confound=True)
    out["Hsub"] = dict(
        exogenous_naive=b_naive_clean, exogenous_controlled=b_ctrl_clean,
        confounded_naive=b_naive_conf, confounded_controlled=b_ctrl_conf,
        note=("Structural dh_f/da>0 shows up as a positive cross-sectional slope "
              "when a is exogenous. The capability confound turns the naive slope "
              "negative; controlling for (residualizing on) capability restores +."),
        exogenous_positive=b_naive_clean > 0,
        confound_flips_naive=b_naive_conf < 0,
        control_restores_positive=b_ctrl_conf > 0,
    )
    out["_df_main"] = df   # for figures
    return out


# =====================================================================
# FIGURES
# =====================================================================

def fig_model_verification(part1, path):
    pr = Params()
    fig, ax = plt.subplots(1, 3, figsize=(13, 3.9))
    # (a) h_f* vs lambda
    lams = np.linspace(0.1, 1.4, 60)
    hf = [solve_closed_form(Params(**{**asdict(pr), "lam": L}))["h_f"] for L in lams]
    ax[0].plot(lams, hf, color="#1f4e79", lw=2)
    ax[0].scatter([pr.lam], [part1["method1"]["h_f"]], color="#c0392b", zorder=5,
                  label="equilibrium")
    ax[0].scatter([pr.lam], [part1["method2"]["h_f"]], marker="x", s=90,
                  color="black", zorder=6, label="numeric solver")
    ax[0].set_xlabel(r"stringency  $\lambda$"); ax[0].set_ylabel(r"$h_f^\star$")
    ax[0].set_title(r"(a) $\partial h_f^\star/\partial\lambda>0$"); ax[0].legend(fontsize=8)
    # (b) h_f* vs a  (H(sub): positive)
    ays = np.linspace(0.02, 0.12, 60)
    hfa = [solve_closed_form(Params(**{**asdict(pr), "a": A}))["h_f"] for A in ays]
    dsa = [solve_closed_form(Params(**{**asdict(pr), "a": A}))["d"] for A in ays]
    ax[1].plot(ays, hfa, color="#1f4e79", lw=2, label=r"$h_f^\star$")
    ax[1].set_xlabel(r"disclosure cost  $a$"); ax[1].set_ylabel(r"$h_f^\star$")
    ax[1].set_title(r"(b) $\partial h_f^\star/\partial a>0$ (H$_{\rm sub}$)")
    axb = ax[1].twinx(); axb.plot(ays, dsa, color="#7f8c8d", lw=1.5, ls="--",
                                  label=r"$d^\star$")
    axb.set_ylabel(r"$d^\star$", color="#7f8c8d")
    # (c) h_f* vs rho
    rhos = np.linspace(-0.4, 0.7, 60)
    hfr = [solve_closed_form(Params(**{**asdict(pr), "rho": R}))["h_f"] for R in rhos]
    ax[2].plot(rhos, hfr, color="#1f4e79", lw=2)
    ax[2].axvline(pr.rho, color="#c0392b", ls=":", lw=1)
    ax[2].set_xlabel(r"correlation  $\rho$"); ax[2].set_ylabel(r"$h_f^\star$")
    ax[2].set_title(r"(c) sign of $\partial h_f^\star/\partial\rho$ per Eq.(20)")
    fig.suptitle("Structural solver: closed form (line) certified by numeric optimizer (x), "
                 "with verified comparative-statics signs", fontsize=10)
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(path+".pdf"); fig.savefig(path+".png", dpi=150); plt.close(fig)


def fig_did_recovery(part2, path):
    es = part2["event_study"]
    xs = sorted(es.keys())
    ys = [es[x] for x in xs]
    fig, ax = plt.subplots(1, 2, figsize=(11, 4.0))
    ax[0].axhline(0, color="gray", lw=0.8)
    ax[0].axvline(-0.5, color="#c0392b", ls=":", lw=1)
    ax[0].plot(xs, ys, "o-", color="#1f4e79")
    ax[0].set_xlabel("event time  (years since phase-in)")
    ax[0].set_ylabel(r"$\theta_\ell$  (effect on $h_f$)")
    ax[0].set_title("(a) Event study: flat pre-trend, positive post")
    labels = ["True ATT", "Callaway--\nSant'Anna", "TWFE"]
    vals = [part2["true_att"], part2["att_callaway_santanna"], part2["beta_twfe"]]
    cols = ["#27ae60", "#1f4e79", "#e67e22"]
    ax[1].bar(labels, vals, color=cols)
    for i, v in enumerate(vals):
        ax[1].text(i, v, f"{v:.4f}", ha="center", va="bottom", fontsize=9)
    ax[1].set_ylabel("ATT on hedge ratio")
    ax[1].set_title("(b) CS recovers true ATT; TWFE biased under\nheterogeneous timing")
    fig.tight_layout()
    fig.savefig(path+".pdf"); fig.savefig(path+".png", dpi=150); plt.close(fig)


def fig_hsub(part2, path):
    h = part2["Hsub"]
    fig, ax = plt.subplots(figsize=(7.6, 4.3))
    groups = ["exogenous\n(naive)", "confounded\n(naive)", "confounded\n(capability-ctrl)"]
    vals = [h["exogenous_naive"], h["confounded_naive"], h["confounded_controlled"]]
    cols = ["#27ae60", "#c0392b", "#1f4e79"]
    ax.axhline(0, color="gray", lw=0.9)
    ax.bar(groups, vals, color=cols)
    for i, v in enumerate(vals):
        ax.text(i, v, f"{v:+.4f}", ha="center",
                va="bottom" if v >= 0 else "top", fontsize=10)
    ax.set_ylabel(r"cross-sectional slope of $h_f$ on disclosure cost $a$")
    ax.set_title("H(sub): structural slope is positive; the capability confound\n"
                 "flips the naive estimate; controlling for capability restores it")
    fig.tight_layout()
    fig.savefig(path+".pdf"); fig.savefig(path+".png", dpi=150); plt.close(fig)


# =====================================================================
# MAIN
# =====================================================================

def main():
    print("="*66)
    print("PART 1  Structural solver + comparative-statics verification")
    print("="*66)
    p1 = run_part1()
    m1, m2 = p1["method1"], p1["method2"]
    print(f"  closed form : h_f={m1['h_f']:.6f}  h_c={m1['h_c']:.6f}  "
          f"d*={m1['d']:.6f}  R*={m1['R']:.6e}")
    print(f"  numeric opt : h_f={m2['h_f']:.6f}  h_c={m2['h_c']:.6f}  d*={m2['d']:.6f}")
    print(f"  max solver disagreement = {max(p1['solver_agreement'].values()):.2e}"
          f"   -> agree: {p1['solvers_agree']}")
    print("  comparative-statics signs:")
    for k, v in p1["sign_checks"].items():
        print(f"     {k:32s} {v}")
    print(f"  ALL SIGN CHECKS PASS: {p1['all_signs_pass']}")

    print("\n"+"="*66)
    print("PART 2  Staggered-adoption DiD Monte Carlo")
    print("="*66)
    p2 = run_part2()
    print(f"  true ATT (model)              = {p2['true_att']:.5f}")
    print(f"  Callaway-Sant'Anna ATT        = {p2['att_callaway_santanna']:.5f}"
          f"   (err {p2['cs_minus_true']:+.5f})")
    print(f"  TWFE ATT                      = {p2['beta_twfe']:.5f}"
          f"   (err {p2['twfe_minus_true']:+.5f})")
    print(f"  TWFE bias vs CS               = {p2['twfe_bias_vs_cs']:+.5f}")
    print(f"  max |pre-trend coef|          = {p2['pretrend_max_abs']:.5f}")
    print(f"  H1 hedging ATT>0              : {p2['H1_hedging_att_positive']}")
    print(f"  H2 adoption ATT>0            : {p2['H2_adoption_att_positive']}  "
          f"(beta_A={p2['beta_adoption']:.5f})")
    print(f"  H3 carbon>non-carbon         : {p2['H3_carbon_stronger']}  "
          f"({p2['att_carbon']:.5f} vs {p2['att_noncarbon']:.5f})")
    hs = p2["Hsub"]
    print("  H(sub) (cross-sectional slope of treated hedging on disclosure cost a):")
    print(f"     exogenous, naive            = {hs['exogenous_naive']:+.5f}")
    print(f"     confounded, naive           = {hs['confounded_naive']:+.5f}")
    print(f"     confounded, capability-ctrl = {hs['confounded_controlled']:+.5f}")
    print(f"     exogenous slope positive    : {hs['exogenous_positive']}")
    print(f"     confound flips naive to <0  : {hs['confound_flips_naive']}")
    print(f"     control restores positive   : {hs['control_restores_positive']}")

    # figures
    # Figures are rendered by make_figures_bw.py (canonical black-and-white
    # versions) from the JSON results, to avoid a colour/B&W regression.
    print("\n  (run make_figures_bw.py to render the black-and-white figures)")

    # persist results (drop the DataFrame)
    p2_save = {k: v for k, v in p2.items() if not k.startswith("_df")}
    results = dict(part1=p1, part2=p2_save)
    with open("engine_results.json", "w") as f:
        json.dump(results, f, indent=2, default=str)
    print("  results written: engine_results.json")
    return results


if __name__ == "__main__":
    main()
