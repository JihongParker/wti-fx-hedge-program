"""
precompute_site_data.py
===================================================================
Offline 200k-path Monte Carlo -> compact response surfaces for the
interactive site. THE computational strategy: the browser never
simulates; it evaluates closed forms and interpolates these surfaces.

Surfaces produced (all validated against paper anchors):
  1. p_KO(S0)   : stress-conditional knock-out probability curve under
                  the PHYSICAL measure (mu=0.139, Merton compensator),
                  symmetrized jumps (lambda=6.7846, theta=-0.0299,
                  delta=0.0844), sigma_diff=0.32419, barriers 120/50,
                  T=0.833, 217 steps, Brownian-bridge crossing.
                  Anchor: p_KO(113) ~ 0.8925 (Paper 1 Eq. 12).
  2. V_KO(S0), delta_KO(S0) : WTI-leg KO call value & delta under Q
                  (r_US=0.04, compensated drift), European exercise at T
                  (the FX factor separates by the homogeneity theorem of
                  Paper 3 Sec 7.4, so the WTI delta shape is exact up to
                  the linear S2 factor). Common random numbers across
                  grid nodes -> smooth differentiable curve.
Output: site_data.json  (grids + calibration block + embedded paper
anchors for the other, fully closed-form modules).
"""
import json, time
import numpy as np

# ------- canonical calibration (docs/supplementary_calibration_table.md)
SIG_DIFF = 0.3241851967428078
LAM, THJ, DLJ = 6.784615384615385, -0.029895127380394192, 0.08442936299651074
MU_P = 0.139
R_US, R_KRW = 0.04, 0.035
U_BAR, L_BAR = 120.0, 50.0
T, NSTEP = 0.833, 217
K_STRIKE = 78.94
NPATH = 200_000
SEED = 20260713

KAPPA = np.exp(THJ + 0.5*DLJ*DLJ) - 1.0          # E[e^Y - 1]
DT = T / NSTEP
SQDT = np.sqrt(DT)


def touch_and_payoff(S0, drift, rng, want_payoff):
    """
    Stream one 200k-path simulation from S0.
    Returns (p_touch, disc_payoff_mean or None).
    Log-space Euler with Poisson jumps at step end + Brownian-bridge
    crossing probability on the diffusive sub-interval (both barriers).
    """
    mu_step = (drift - LAM*KAPPA - 0.5*SIG_DIFF**2) * DT
    logU, logL = np.log(U_BAR), np.log(L_BAR)
    x = np.full(NPATH, np.log(S0))
    alive = np.ones(NPATH, dtype=bool)
    # bridge survival probability accumulates multiplicatively per path
    u_rand_needed = True
    surv = np.ones(NPATH)                          # P(no bridge crossing so far)
    for _ in range(NSTEP):
        z = rng.standard_normal(NPATH)
        x_diff = x + mu_step + SIG_DIFF*SQDT*z     # diffusion endpoint
        # Brownian bridge two-barrier (upper & lower, independent approx)
        with np.errstate(over='ignore'):
            pu = np.exp(-2.0*(logU - x)*(logU - x_diff)/(SIG_DIFF**2*DT))
            pl = np.exp(-2.0*(x - logL)*(x_diff - logL)/(SIG_DIFF**2*DT))
        pu = np.where((x < logU) & (x_diff < logU), np.minimum(pu, 1.0), 1.0)
        pl = np.where((x > logL) & (x_diff > logL), np.minimum(pl, 1.0), 1.0)
        surv = np.where(alive, surv*(1.0-pu)*(1.0-pl), surv)
        # jumps at step end
        nj = rng.poisson(LAM*DT, NPATH)
        jmp = np.where(nj > 0, THJ*nj + DLJ*np.sqrt(np.maximum(nj, 1))*rng.standard_normal(NPATH), 0.0)
        x = x_diff + jmp
        # direct level check (captures jump-through and endpoint touches)
        hit = (x >= logU) | (x <= logL)
        surv = np.where(alive & hit, 0.0, surv)
        alive = alive & ~hit
    p_touch = 1.0 - surv.mean()
    if not want_payoff:
        return p_touch, None
    payoff = np.maximum(np.exp(x) - K_STRIKE, 0.0) * surv     # survival-weighted
    return p_touch, np.exp(-R_US*T) * payoff.mean()


def main():
    t0 = time.time()
    out = {}

    # ---- 1. p_KO(S0) under P ------------------------------------------
    grid_p = sorted(set(list(np.arange(60, 120, 2.5)) + [78.94, 113.0, 118.0, 119.0]))
    pko = []
    rng = np.random.default_rng(SEED)
    for S0 in grid_p:
        p, _ = touch_and_payoff(S0, MU_P, rng, want_payoff=False)
        pko.append(round(float(p), 5))
        print(f"  pKO({S0:7.2f}) = {p:.4f}")
    out["pko_grid"] = {"S0": [round(float(s), 2) for s in grid_p], "pko": pko}
    print(f"  anchors: pKO(113)={pko[grid_p.index(113.0)]:.4f} (paper 0.8925), "
          f"pKO(78.94)={pko[grid_p.index(78.94)]:.4f} (paper raw-touch 0.451)")

    # ---- 2. V_KO(S0), delta via CRN central differences under Q -------
    grid_q = [float(s) for s in np.arange(55, 120, 2.0)]
    V = []
    for S0 in grid_q:
        rngq = np.random.default_rng(SEED + 7)     # CRN: same shocks every node
        _, v = touch_and_payoff(S0, R_US, rngq, want_payoff=True)
        V.append(round(float(v), 5))
        print(f"  V_KO({S0:6.1f}) = {v:.4f}")
    # central-difference delta on the grid
    delta = []
    for i in range(len(grid_q)):
        if i == 0:
            d = (V[1]-V[0])/(grid_q[1]-grid_q[0])
        elif i == len(grid_q)-1:
            d = (V[-1]-V[-2])/(grid_q[-1]-grid_q[-2])
        else:
            d = (V[i+1]-V[i-1])/(grid_q[i+1]-grid_q[i-1])
        delta.append(round(float(d), 5))
    out["ko_value_grid"] = {"S0": grid_q, "V": V, "delta": delta}

    # ---- 3. exact c* parabola from the paper's sweep (affine identity) -
    # Var(c) = VarA + 2c Cov + c^2 VarB ; fit exactly through 3 sweep points
    cs  = [0.5, 1.0, 1.5]
    sds = [32.187496866, 59.369149552, 89.793641566]     # bn KRW (Table fxsweep)
    A_ = np.array([[1, 2*c, c*c] for c in cs])
    b_ = np.array([s*s for s in sds])
    varA, cov, varB = np.linalg.solve(A_, b_)
    cstar_fit = -cov/varB
    out["cstar"] = {
        "varA": round(float(varA), 4), "cov": round(float(cov), 4),
        "varB": round(float(varB), 4),
        "cstar_production_fit": round(float(cstar_fit), 4),
        "cstar_exact_engine": -0.548,
        "sd_at_1_exact": 51.90, "sd_at_cstar_exact": 48.76,
        "sweep": {"c": [0.5, 0.75, 1.0, 1.25, 1.5],
                  "mean_bn": [-43.617, -49.682, -55.747, -61.813, -67.878],
                  "sd_bn":   [32.187, 44.991, 59.369, 74.415, 89.794]},
    }
    print(f"  c* production-fit = {cstar_fit:.4f} (exact engine -0.548)")

    # ---- 4. calibration + paper-anchor block (closed-form modules) -----
    out["calib"] = {
        "S1_0": 78.94, "S2_0": 1540.64, "K": K_STRIKE,
        "sig1_raw": 0.39455, "sig1_diff": SIG_DIFF, "sig2": 0.09258,
        "rho": 0.08763, "r_US": R_US, "r_KRW": R_KRW, "r_w": 0.07,
        "lam": LAM, "thJ": THJ, "dlJ": DLJ, "mu1_P": MU_P, "mu2_P": 0.0239,
        "U": U_BAR, "L": L_BAR, "T1": 0.833, "T2": 0.5, "nstep": NSTEP,
        "Q_oil": 2_000_000, "Q_usd": 157_880_000, "B": 45e9,
        "stress_WTI": 113.0, "stress_KRW": 1550.0,
        "P_B76": 12.6524, "P_GK": 84.667,
        "P_Sh_WTI": 15093.75, "P_Sh_FX": 2038.72, "P_K_standalone": 17377.11,
        "base_quanto_premium": 17132.47,
        "pko_stress_paper": 0.8925, "pstar_shapley": 0.10913,
        "pstar_standalone": 0.08304, "pbar": 0.06161,
        "A_FV0": 34_264_941_589,
    }

    # ---- 5. ESG engine results (event study, Hsub) ----------------------
    eng = json.load(open("engine_results.json"))
    out["esg"] = {
        "event_study": eng["part2"]["event_study"],
        "true_att": eng["part2"]["true_att"],
        "att_cs": eng["part2"]["att_callaway_santanna"],
        "att_twfe": eng["part2"]["beta_twfe"],
        "hsub": {k: v for k, v in eng["part2"]["Hsub"].items() if isinstance(v, (int, float))},
        "params": eng["part1"]["params"],
        "equilibrium": {k: eng["part1"]["method1"][k] for k in ("h_f", "h_c", "d", "R")},
    }

    out["_meta"] = {"npaths": NPATH, "nsteps": NSTEP, "seed": SEED,
                    "runtime_s": round(time.time()-t0, 1),
                    "note": "offline 200k-path surfaces; browser interpolates"}
    json.dump(out, open("site_data.json", "w"))
    print(f"\nwrote site_data.json  ({time.time()-t0:.0f}s)")


if __name__ == "__main__":
    main()
