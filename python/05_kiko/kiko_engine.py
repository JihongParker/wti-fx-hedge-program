"""KIKO through the program — numerical engine.

Reconstructs a stylized 1-put : 2-call KIKO on USD/KRW under the research
program's own FX calibration (sigma2 = 0.09258, r_USD = 4.0%, r_KRW = 3.5%,
S0 = 1540.64) and prices it with the program's own method: risk-neutral GBM
Monte Carlo at 260 steps/yr with log-space Brownian-bridge first-passage
correction on every barrier (the P2/P3 machinery, FX leg has no jumps in the
program's calibration).

Structure (stylized single-settlement KIKO, T = 1y, continuous monitoring):
  long  1x  put  strike K, knocked OUT if S ever <= L   (protection leg)
  short 2x  call strike K, knocked IN  if S ever >= U   (financing leg)
K is solved so the net premium is zero given (L, U) = (0.95, 1.07) x S0.

Outputs (JSON + three b/w figures):
  1. zero-cost decomposition: K* vs forward, leg premia, the "sweetener"
  2. stress ladder: package + firm P&L at S_T in {+10%, +25%, +50%}, m = 1 / 2.5
  3. protection-leg mortality: P(put knocked out | S_T <= x), breakeven vs measured
  4. package delta profile Delta(S0): the quiet band and the cliff

Run:  python3 kiko_engine.py out_dir
"""
import json
import sys
import time

import numpy as np

# program calibration (docs/supplementary_calibration_table.md)
SIGMA = 0.09258
R_USD = 0.040
R_KRW = 0.035
S0 = 1540.64
T = 1.0
STEPS = 260
N_PATHS = 200_000
SEED = 20080915  # Lehman day, for the occasion

L_FRAC, U_FRAC = 0.95, 1.07
LEV = 2.0  # calls per put

rng_global = np.random.default_rng(SEED)


def simulate(s0: float, n_paths: int = N_PATHS, seed: int = SEED):
    """GBM under Q for USD/KRW (domestic KRW measure): drift r_KRW - r_USD."""
    rng = np.random.default_rng(seed)
    dt = T / STEPS
    z = rng.standard_normal((n_paths, STEPS))
    dX = (R_KRW - R_USD - 0.5 * SIGMA**2) * dt + SIGMA * np.sqrt(dt) * z
    X = np.concatenate([np.zeros((n_paths, 1)), np.cumsum(dX, axis=1)], axis=1)
    S = s0 * np.exp(X)
    return S, dt


def bridge_hit_prob(S, level: float, dt: float, side: str):
    """P(path crossed `level` between grid points) via log-space Brownian
    bridge; returns per-path survival probability of NOT hitting."""
    a = np.log(S[:, :-1] / level)
    b = np.log(S[:, 1:] / level)
    var = SIGMA**2 * dt
    if side == "down":  # hit if S <= level
        inside = (a > 0) & (b > 0)
        with np.errstate(over="ignore"):
            p = np.where(inside, np.exp(-2.0 * a * b / var), 1.0)
    else:  # hit if S >= level
        inside = (a < 0) & (b < 0)
        with np.errstate(over="ignore"):
            p = np.where(inside, np.exp(-2.0 * a * b / var), 1.0)
    surv = np.prod(1.0 - p, axis=1)
    return surv  # probability of never hitting, path by path


def price_legs(K: float, s0: float = S0, seed: int = SEED, n_paths: int = N_PATHS):
    """Present value (KRW per USD of notional) of the KO put and KI call."""
    S, dt = simulate(s0, n_paths, seed)
    disc = np.exp(-R_KRW * T)
    surv_L = bridge_hit_prob(S, L_FRAC * S0, dt, "down")  # put alive prob
    surv_U = bridge_hit_prob(S, U_FRAC * S0, dt, "up")
    hit_U = 1.0 - surv_U  # call knocked-in prob
    ST = S[:, -1]
    put_ko = disc * np.mean(np.maximum(K - ST, 0.0) * surv_L)
    call_ki = disc * np.mean(np.maximum(ST - K, 0.0) * hit_U)
    call_vanilla = disc * np.mean(np.maximum(ST - K, 0.0))
    put_vanilla = disc * np.mean(np.maximum(K - ST, 0.0))
    return dict(put_ko=put_ko, call_ki=call_ki, put_vanilla=put_vanilla,
                call_vanilla=call_vanilla)


def solve_zero_cost_K():
    """K* such that put_ko(K) - LEV * call_ki(K) = 0 (CRN across K)."""
    lo, hi = 0.9 * S0, 1.15 * S0
    for _ in range(40):
        mid = 0.5 * (lo + hi)
        legs = price_legs(mid)
        net = legs["put_ko"] - LEV * legs["call_ki"]
        # net premium rises with K (put dearer, call cheaper) -> bisect
        if net < 0:
            lo = mid
        else:
            hi = mid
    K = 0.5 * (lo + hi)
    return K, price_legs(K)


def mortality_curve(x_grid):
    """P(protection knocked out | S_T <= x) for x in grid."""
    S, dt = simulate(S0)
    surv_L = bridge_hit_prob(S, L_FRAC * S0, dt, "down")
    ST = S[:, -1]
    out = []
    for x in x_grid:
        sel = ST <= x
        if sel.sum() < 200:
            out.append(np.nan)
        else:
            out.append(float(1.0 - surv_L[sel].mean()))
    return out


def package_value(s0: float, K: float, seed: int = SEED):
    legs = price_legs(K, s0=s0, seed=seed, n_paths=60_000)
    return legs["put_ko"] - LEV * legs["call_ki"]


def delta_profile(K: float):
    grid = np.round(np.linspace(0.88 * S0, 1.16 * S0, 29), 1)
    vals = np.array([package_value(s, K) for s in grid])
    delta = np.gradient(vals, grid)
    return grid.tolist(), vals.tolist(), delta.tolist()


def main(out_dir: str):
    t0 = time.time()
    K, legs = solve_zero_cost_K()
    F = S0 * np.exp((R_KRW - R_USD) * T)
    sweetener_pct = (K - F) / F * 100

    # protection mortality at "protection actually needed" thresholds
    x_grid = (S0 * np.array([0.90, 0.92, 0.94, 0.95, 0.96, 0.98])).tolist()
    mort = mortality_curve(x_grid)
    # headline: conditional on ending at/below the KO barrier level region
    mort_at_945 = mort[3]  # S_T <= 0.95 S0

    # KO discount on the protection leg and breakeven mortality (P1 §7 logic):
    # discount d = 1 - putKO/putVanilla; the discount is worth taking only if
    # mortality when needed < d (you save d of premium, you lose protection
    # with prob p exactly when it pays) -> breakeven p_bar = d.
    disc_put = 1.0 - legs["put_ko"] / legs["put_vanilla"]

    # stress ladder (KI hit certain on any path reaching these levels)
    ladder = []
    for pct in (0.10, 0.25, 0.50):
        ST = S0 * (1 + pct)
        pkg = -LEV * max(ST - K, 0.0)  # per USD of hedged notional
        oper = ST - S0  # exporter revenue gain per USD of flow
        for m in (1.0, 2.5):
            net = oper + m * pkg
            ladder.append(dict(stress=pct, m=m, per_usd=net,
                               loss_ratio=(-m * pkg) / oper if oper > 0 else np.nan))

    grid, vals, delta = delta_profile(K)

    out = dict(
        meta=dict(sigma=SIGMA, r_usd=R_USD, r_krw=R_KRW, s0=S0, T=T,
                  steps=STEPS, paths=N_PATHS, L=L_FRAC * S0, U=U_FRAC * S0,
                  leverage=LEV, seed=SEED),
        zero_cost=dict(K=K, forward=F, sweetener_pct=sweetener_pct,
                       put_ko=legs["put_ko"], call_ki=legs["call_ki"],
                       put_vanilla=legs["put_vanilla"],
                       call_vanilla=legs["call_vanilla"],
                       net=legs["put_ko"] - LEV * legs["call_ki"],
                       put_ko_discount=disc_put),
        mortality=dict(x=x_grid, p=mort, headline=mort_at_945,
                       breakeven=disc_put),
        ladder=ladder,
        delta=dict(grid=grid, value=vals, delta=delta),
        runtime_s=time.time() - t0,
    )
    with open(f"{out_dir}/kiko_results.json", "w") as f:
        json.dump(out, f, indent=1)
    print(json.dumps(dict(K=round(K, 2), forward=round(F, 2),
                          sweetener_pct=round(sweetener_pct, 2),
                          put_ko=round(legs["put_ko"], 2),
                          call_ki=round(legs["call_ki"], 2),
                          put_ko_discount=round(disc_put, 4),
                          mortality_at_0p95=round(mort_at_945, 4)), indent=1))
    print(f"runtime {time.time()-t0:.1f}s")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
