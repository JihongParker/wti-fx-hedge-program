"""Longstaff-Schwartz valuation of the double knock-out quanto WTI call.

Python port of the workbook's Calc_LSMC_Price / Run_LSMC_Engine (VBA module
DeltaHedging_revised).  Same dynamics, same basis, same exercise rule, same
Shapley split, so the paper's American-regime premiums can be reproduced from
source instead of read out of a spreadsheet.

  S1: WTI, Merton jump-diffusion, risk-neutral drift r_US, compensated jumps
  S2: USD/KRW, GBM, drift r_KRW - r_US, diffusive correlation rho with S1
  payoff (KRW): alive ? max(S1_T - K, 0) * S2_T : 0
  barrier: continuous double knock-out on S1 at [KOlower, KOupper], monitored
           at each step and, on jump-free steps, by the Brownian-bridge
           crossing probability
  basis: 1, v1, v2, v1^2, v2^2, v1*v2   with v1 = S1/K, v2 = S2/S2_0
"""
import json
import numpy as np

CAL = dict(S1_0=78.94, S2_0=1540.64, vol1=0.3241851967428078,
           vol2=0.09257650229436072, corr=0.08763086561435483,
           lam=6.784615384615385, thJ=-0.029895127380394192,
           dlJ=0.08442936299651074, r_US=0.04, r_KRW=0.035,
           KOupper=120.0, KOlower=50.0, T=0.833, steps=217)


def _paths(c, npaths, rng):
    dt = c['T']/c['steps']
    sqdt = np.sqrt(dt)
    kappa = np.exp(c['thJ'] + 0.5*c['dlJ']**2) - 1.0
    m1 = c['r_US'] - 0.5*c['vol1']**2 - c['lam']*kappa
    m2 = (c['r_KRW'] - c['r_US']) - 0.5*c['vol2']**2
    lnVar = c['vol1']**2*dt
    U, L = c['KOupper'], c['KOlower']

    n, S = npaths, c['steps']
    S1 = np.empty((n, S+1), np.float32); S2 = np.empty((n, S+1), np.float32)
    AL = np.zeros((n, S+1), bool)
    s1 = np.full(n, c['S1_0']); s2 = np.full(n, c['S2_0'])
    alive = np.ones(n, bool)
    S1[:, 0], S2[:, 0], AL[:, 0] = s1, s2, alive

    for i in range(1, S+1):
        prev = s1.copy()
        z1 = rng.standard_normal(n); z2 = rng.standard_normal(n)
        e2 = c['corr']*z1 + np.sqrt(1-c['corr']**2)*z2
        nJ = rng.poisson(c['lam']*dt, n)
        js = np.where(nJ > 0, rng.normal(c['thJ']*nJ, c['dlJ']*np.sqrt(np.maximum(nJ, 1))), 0.0)
        s1 = np.where(alive, s1*np.exp(m1*dt + c['vol1']*sqdt*z1 + js), s1)
        s2 = np.where(alive, s2*np.exp(m2*dt + c['vol2']*sqdt*e2), s2)
        hard = alive & ((s1 >= U) | (s1 <= L))
        alive &= ~hard
        if lnVar > 0:
            br = alive & (nJ == 0)
            pu = np.where(br, np.exp(-2*np.log(U/prev)*np.log(U/np.maximum(s1, 1e-12))/lnVar), 0.0)
            pl = np.where(br, np.exp(-2*np.log(prev/L)*np.log(np.maximum(s1, 1e-12)/L)/lnVar), 0.0)
            alive &= ~(br & (rng.random(n) < 1-(1-pu)*(1-pl)))
        S1[:, i], S2[:, i], AL[:, i] = s1, s2, alive
    return S1, S2, AL


def price(K=None, npaths=50_000, seed=12345, **over):
    c = dict(CAL); c.update(over)
    K = c['S1_0'] if K is None else K
    rng = np.random.default_rng(seed)
    S1, S2, AL = _paths(c, npaths, rng)
    S, disc = c['steps'], np.exp(-c['r_US']*c['T']/c['steps'])
    CF = np.where(AL[:, S], np.maximum(S1[:, S].astype(float)-K, 0)*S2[:, S], 0.0)
    for i in range(S-1, 0, -1):
        CF *= disc
        m = AL[:, i] & (S1[:, i] > K)
        if m.sum() > 10:
            v1 = S1[m, i].astype(float)/K
            v2 = S2[m, i].astype(float)/c['S2_0']
            X = np.column_stack([v1, v2, v1*v1, v2*v2, v1*v2, np.ones_like(v1)])
            X = np.nan_to_num(X, posinf=0.0, neginf=0.0)
            beta, *_ = np.linalg.lstsq(X, CF[m], rcond=1e-10)
            cont = X @ beta
            intr = (S1[m, i].astype(float)-K)*S2[m, i]
            ex = np.isfinite(cont) & (intr > cont)
            CF[np.where(m)[0][ex]] = intr[ex]
    v = (CF*disc).mean()
    return v, (CF*disc).std(ddof=1)/np.sqrt(npaths)


def shapley(npaths=50_000, seed=12345):
    """Workbook's split: phi_i = 1/2 v({i}) + 1/2 (v(N) - v(N\\{i}))."""
    base, se = price(npaths=npaths, seed=seed)
    wti, se_w = price(npaths=npaths, seed=seed, vol2=1e-5)          # FX frozen
    fx,  se_f = price(npaths=npaths, seed=seed, vol1=1e-5, lam=0.0)  # WTI frozen
    phi_w = 0.5*wti + 0.5*(base - fx)
    phi_f = 0.5*fx + 0.5*(base - wti)
    tot = phi_w + phi_f
    return dict(base=base, base_se=se, v_wti_only=wti, v_fx_only=fx,
                phi_wti=phi_w, phi_fx=phi_f, ratio_wti=phi_w/tot, ratio_fx=phi_f/tot,
                P_Sh_WTI=tot*phi_w/tot, P_Sh_FX=tot*phi_f/tot)


if __name__ == '__main__':
    import sys
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 50_000
    r = shapley(npaths=n)
    print(f"paths={n}")
    print(f"  joint quanto premium   = {r['base']:12,.2f} +/- {r['base_se']:.2f}  KRW/bbl   (workbook 17,132.47)")
    print(f"  v(WTI only)            = {r['v_wti_only']:12,.2f}   <- the standalone single-asset KO leg")
    print(f"  v(FX only)             = {r['v_fx_only']:12,.2f}")
    print(f"  Shapley WTI            = {r['phi_wti']:12,.2f}  ratio {r['ratio_wti']:.6f}   (workbook 15,093.75 / 0.881003)")
    print(f"  Shapley FX             = {r['phi_fx']:12,.2f}  ratio {r['ratio_fx']:.6f}   (workbook  2,038.72 / 0.118997)")
