"""Independent adjudication: is the MV coupling c* < 0 (Park_quanto) or is c = 1
optimal (build_ultimate_pipeline)? Settle it on the frozen engine with the
CORRECT leg decomposition.

Structure: V(S1,S2) = U(S1) * S2, where U is the USD value of the double-KO WTI
call (knock-out on S1). Exact deltas: dV/dS1 = U'(S1)*S2 (WTI), dV/dS2 = U = V/S2
(FX, the paper's homogeneity result). Practitioner family: delta_FX = c*delta_WTI.

Sell the quanto, delta-hedge WTI with delta1 = U'*S2, and the FX leg with
c*delta1 units of a USD/KRW forward. Per-path replication error:
  HE_i(c) = V_T,i - V0 - sum delta1*dS1  -  c*(sum delta1*dS2)
          = A_i + c*B_i,     A_i = V_T,i - V0 - sum delta1*dS1,  B_i = -sum delta1*dS2
  Var(HE) = VarA + 2c CovAB + c^2 VarB,   c* = -CovAB/VarB.

No predetermined answer: we compute A_i, B_i honestly and read the sign of c*.
"""
import sys, time
import numpy as np

SIGMA1, LAMBDA, THETA_J, DELTA_J = 0.32419, 6.7846, -0.0299, 0.0844
SIGMA2, RHO, R_US, R_KRW = 0.09258, 0.0876, 0.040, 0.035
S1_0, S2_0, K, U_, L = 78.94, 1540.64, 78.94, 120.0, 50.0
STEPS_PER_YEAR = 260
T = 0.833
N_PRICE = 30_000
N_HEDGE = 40_000
SEED = 20260716
kappa = np.exp(THETA_J + 0.5 * DELTA_J ** 2) - 1.0


def usd_call_grid(tau, s1_grid, rng):
    """U(tau, s1): USD value of the double-KO WTI call (per unit, no S2)."""
    steps = max(10, int(round(STEPS_PER_YEAR * tau)))
    dt = tau / steps
    drift = (R_US - 0.5 * SIGMA1 ** 2 - LAMBDA * kappa) * dt
    z1 = rng.standard_normal((N_PRICE, steps))
    nj = rng.poisson(LAMBDA * dt, size=(N_PRICE, steps))
    jm = THETA_J * nj + DELTA_J * np.sqrt(nj) * rng.standard_normal((N_PRICE, steps))
    dX = drift + SIGMA1 * np.sqrt(dt) * z1 + jm
    X = np.concatenate([np.zeros((N_PRICE, 1)), np.cumsum(dX, axis=1)], axis=1)
    var = SIGMA1 ** 2 * dt
    Xi, Xj = X[:, :-1], X[:, 1:]
    out = np.empty(len(s1_grid))
    for k, s0 in enumerate(s1_grid):
        bU, bL = np.log(U_ / s0), np.log(L / s0)
        inside = (Xi < bU) & (Xi > bL) & (Xj < bU) & (Xj > bL)
        with np.errstate(over="ignore"):
            pu = np.where(inside, np.exp(-2 * (bU - Xi) * (bU - Xj) / var), 0.0)
            pd = np.where(inside, np.exp(-2 * (Xi - bL) * (Xj - bL) / var), 0.0)
        surv = np.prod(np.where(inside, (1 - pu) * (1 - pd), 0.0), axis=1)
        out[k] = np.exp(-R_US * tau) * (np.maximum(s0 * np.exp(X[:, -1]) - K, 0.0) * surv).mean()
    return out


def main():
    t0 = time.time()
    s1_grid = np.round(np.arange(50.5, 119.5 + 1e-9, 1.5), 2)
    taus = np.round(np.linspace(T, 0.02, 20), 4)
    Ugrid = np.zeros((len(taus), len(s1_grid)))
    for i, tau in enumerate(taus):
        Ugrid[i] = usd_call_grid(tau, s1_grid, np.random.default_rng(SEED + i))
        print(f"  U(tau={tau:.3f})", file=sys.stderr)
    dU = np.gradient(Ugrid, s1_grid, axis=1)           # U'(tau, s1)
    tau_asc = taus[::-1]; U_asc = Ugrid[::-1]; dU_asc = dU[::-1]

    # hedge paths
    steps = max(24, int(round(STEPS_PER_YEAR * T)))
    dt = T / steps
    rng = np.random.default_rng(SEED + 777)
    drift = (R_US - 0.5 * SIGMA1 ** 2 - LAMBDA * kappa) * dt
    z1 = rng.standard_normal((N_HEDGE, steps))
    nj = rng.poisson(LAMBDA * dt, size=(N_HEDGE, steps))
    jm = THETA_J * nj + DELTA_J * np.sqrt(nj) * rng.standard_normal((N_HEDGE, steps))
    S1 = S1_0 * np.exp(np.concatenate([np.zeros((N_HEDGE, 1)),
         np.cumsum(drift + SIGMA1 * np.sqrt(dt) * z1 + jm, axis=1)], axis=1))
    z2 = rng.standard_normal((N_HEDGE, steps))
    dW2 = RHO * z1 + np.sqrt(1 - RHO ** 2) * z2
    S2 = S2_0 * np.exp(np.concatenate([np.zeros((N_HEDGE, 1)),
         np.cumsum((R_KRW - R_US - 0.5 * SIGMA2 ** 2) * dt + SIGMA2 * np.sqrt(dt) * dW2, axis=1)], axis=1))
    alive = ~(((S1 >= U_) | (S1 <= L)).any(axis=1))

    V0 = float(np.interp(S1_0, s1_grid, U_asc[np.searchsorted(tau_asc, T) - 1]) * S2_0)
    A = np.zeros(N_HEDGE); B = np.zeros(N_HEDGE); dead = np.zeros(N_HEDGE, bool)
    for s in range(steps):
        tau = T - s * dt
        # U'(tau, .) row via per-node interp in tau, then interp over S1 per path
        row = np.array([np.interp(tau, tau_asc, dU_asc[:, j]) for j in range(len(s1_grid))])
        up = np.interp(S1[:, s], s1_grid, row)          # U'(S1_s)
        d1 = up * S2[:, s]                               # delta_WTI = U'*S2  (KRW per USD/bbl)
        live = ~dead
        A[live] += d1[live] * (S1[live, s + 1] - S1[live, s])   # WTI hedge P&L
        B[live] += d1[live] * (S2[live, s + 1] - S2[live, s])   # FX-coupling P&L (per unit c)
        dead |= (S1[:, s + 1] >= U_) | (S1[:, s + 1] <= L)
    VT = np.maximum(S1[:, -1] - K, 0.0) * S2[:, -1] * alive
    A_cost = VT - V0 - A          # replication error, WTI-hedged (c=0)
    B_cost = -B                   # FX-coupling term
    VarA = float(np.var(A_cost, ddof=1)); VarB = float(np.var(B_cost, ddof=1))
    CovAB = float(np.cov(A_cost, B_cost, ddof=1)[0, 1])
    c_star = -CovAB / VarB
    var_at = lambda c: VarA + 2 * c * CovAB + c * c * VarB
    print(f"\nKO rate {1-alive.mean():.3f}   V0={V0:,.0f} KRW/unit")
    print(f"VarA={VarA:.3e}  VarB={VarB:.3e}  CovAB={CovAB:+.3e}")
    print(f"c* = -Cov/Var = {c_star:+.3f}    (Park_quanto: -0.548 | build_ultimate: +1.0)")
    print(f"Var(c=1)/Var(c*) = {var_at(1)/var_at(c_star):.2f}x   Var(c=0)/Var(c*) = {var_at(0)/var_at(c_star):.2f}x")
    print(f"verdict: {'c* < 0  -> Park_quanto' if c_star < -0.05 else ('c ~ 1 -> build_ultimate' if c_star>0.5 else 'ambiguous/near-zero')}")
    print(f"done {time.time()-t0:.1f}s")


if __name__ == "__main__":
    main()
