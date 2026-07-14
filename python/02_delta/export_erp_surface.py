"""Export price/delta/KO-probability surfaces of the European double-KO quanto
for the HongERP Exotic Desk.

Model & calibration are the paper's (Park_quanto, production configuration):
  WTI  : Merton jump-diffusion under Q — drift r_US, diffusive sigma1=0.32419,
         jumps lambda=6.7846/yr, J ~ N(-0.0299, 0.0844^2), compensator drift.
  FX   : GBM under Q — drift r_KRW - r_US, sigma2=0.09258, corr rho=0.0876.
  KO   : S1 >= 120 or S1 <= 50, continuous monitoring approximated by
         260 steps/yr discrete checks + log-space Brownian-bridge crossing
         correction on the diffusive component (jumps checked discretely).
  Payoff (European variant): max(S1_T - K, 0) * S2_T * 1{no KO}, K = 78.94,
         discounted at r_KRW. Reported per quanto unit (KRW).

Surface trick: S1(t) = S1_0 * exp(X_t) with X independent of S1_0, so ONE
path bank per maturity is reused across the whole spot grid (common random
numbers across grid nodes — the paper's CRN discipline).

American early exercise is deliberately NOT modeled here (the ERP labels the
surface as European); the paper's American/European KO-rate gap (0.2309 vs
0.4369) is an early-exercise pre-emption artifact, per the paper §KO-rate.

Run:  python3 export_erp_surface.py out.json
"""
import json
import sys
import time

import numpy as np

# --- paper calibration (docs/supplementary_calibration_table.md, A1-A3) ---
SIGMA1 = 0.32419
LAMBDA = 6.7846
THETA_J = -0.0299
DELTA_J = 0.0844
SIGMA2 = 0.09258
RHO = 0.0876
R_US = 0.040
R_KRW = 0.035
S1_0_BASE = 78.94
S2_0 = 1540.64
K = 78.94
U, L = 120.0, 50.0
STEPS_PER_YEAR = 260

N_PATHS = 30_000
SEED = 20260714

T_GRID = [0.833, 0.667, 0.500, 0.333, 0.167, 0.083]
S_GRID = np.round(np.arange(52.0, 118.0 + 1e-9, 2.0), 2)  # 34 nodes inside barriers

kappa = np.exp(THETA_J + 0.5 * DELTA_J ** 2) - 1.0  # jump compensator


def simulate_bank(T: float, rng: np.random.Generator):
    """One CRN path bank: X (log-return paths, n_paths x (steps+1)) and the
    per-step diffusive variance for the bridge correction."""
    steps = max(20, int(round(STEPS_PER_YEAR * T)))
    dt = T / steps
    drift = (R_US - 0.5 * SIGMA1 ** 2 - LAMBDA * kappa) * dt
    z1 = rng.standard_normal((N_PATHS, steps))
    nj = rng.poisson(LAMBDA * dt, size=(N_PATHS, steps))
    jumps = THETA_J * nj + DELTA_J * np.sqrt(nj) * rng.standard_normal((N_PATHS, steps))
    dX = drift + SIGMA1 * np.sqrt(dt) * z1 + jumps
    X = np.concatenate([np.zeros((N_PATHS, 1)), np.cumsum(dX, axis=1)], axis=1)
    # FX terminal (only S2_T enters the payoff)
    z2 = rng.standard_normal((N_PATHS, steps))
    dW2 = RHO * z1 + np.sqrt(1 - RHO ** 2) * z2
    X2_T = (R_KRW - R_US - 0.5 * SIGMA2 ** 2) * T + SIGMA2 * np.sqrt(dt) * dW2.sum(axis=1)
    S2_T = S2_0 * np.exp(X2_T)
    return X, S2_T, dt, steps


def node_values(X, S2_T, dt, T, s1_0):
    """Survival-weighted discounted payoff and KO prob for one spot node,
    fully vectorized over the shared path bank."""
    bU = np.log(U / s1_0)
    bL = np.log(L / s1_0)
    Xi, Xj = X[:, :-1], X[:, 1:]
    inside_i = (Xi < bU) & (Xi > bL)
    inside_j = (Xj < bU) & (Xj > bL)
    inside = inside_i & inside_j
    var = SIGMA1 ** 2 * dt
    with np.errstate(over="ignore"):
        p_up = np.where(inside, np.exp(-2.0 * (bU - Xi) * (bU - Xj) / var), 0.0)
        p_dn = np.where(inside, np.exp(-2.0 * (Xi - bL) * (Xj - bL) / var), 0.0)
    step_surv = np.where(inside, (1.0 - p_up) * (1.0 - p_dn), 0.0)
    survival = np.prod(step_surv, axis=1)  # 0 on any discrete breach
    ko_prob = 1.0 - survival.mean()
    payoff = np.maximum(s1_0 * np.exp(X[:, -1]) - K, 0.0) * S2_T * survival
    price = np.exp(-R_KRW * T) * payoff.mean()
    return price, ko_prob


def main(out_path: str):
    t0 = time.time()
    price = np.zeros((len(T_GRID), len(S_GRID)))
    ko = np.zeros_like(price)
    for ti, T in enumerate(T_GRID):
        rng = np.random.default_rng(SEED + ti)  # CRN within a maturity, fresh across
        X, S2_T, dt, steps = simulate_bank(T, rng)
        for si, s1_0 in enumerate(S_GRID):
            price[ti, si], ko[ti, si] = node_values(X, S2_T, dt, T, s1_0)
        print(f"T={T:.3f} ({steps} steps) done  {time.time()-t0:5.1f}s", file=sys.stderr)

    # WTI delta via central differences on the spot grid (one-sided at edges)
    delta = np.gradient(price, S_GRID, axis=1)

    # anchors vs paper
    si0 = int(np.argmin(np.abs(S_GRID - S1_0_BASE)))
    anchor = {
        "s1": float(S_GRID[si0]),
        "T": T_GRID[0],
        "price_per_unit_KRW": float(price[0, si0]),
        "ko_prob_Q": float(ko[0, si0]),
        "paper_european_ko_rate": 0.4369,
        "paper_american_base_premium": 17132.47,
        "note": "European surface: price <= American 17,132.47 expected; KO prob "
                "compares to paper European engine 0.4369 (baseline spot differs "
                "from grid node by <=1 USD).",
    }
    out = {
        "meta": {
            "engine": "European double-KO quanto, jump-diffusion Q-measure MC, "
                      "260 steps/yr + log-space Brownian-bridge KO correction "
                      "(diffusive component; jumps checked discretely)",
            "paths": N_PATHS,
            "seed": SEED,
            "calibration": {
                "sigma1": SIGMA1, "lambda": LAMBDA, "thetaJ": THETA_J,
                "deltaJ": DELTA_J, "sigma2": SIGMA2, "rho": RHO,
                "rUS": R_US, "rKRW": R_KRW, "K": K, "U": U, "L": L,
                "S2_0": S2_0,
            },
            "units": "price: KRW per quanto unit at S2_0=1540.64 "
                     "(scales linearly in S2 by homogeneity, paper thm V/S2)",
            "anchor": anchor,
            "generated": time.strftime("%Y-%m-%d %H:%M:%S"),
        },
        "sGrid": S_GRID.tolist(),
        "tGrid": T_GRID,
        "price": np.round(price, 2).tolist(),
        "deltaWTI": np.round(delta, 2).tolist(),
        "koProb": np.round(ko, 4).tolist(),
    }
    with open(out_path, "w") as f:
        json.dump(out, f)
    print(json.dumps(anchor, indent=2))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "erp_surface.json")
