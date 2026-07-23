"""Orthogonal Girsanov validation on the knock-out structure's own payoff.

Two independent checks of the same construction:

  density check   E[L] = 1 and sd[L] = sqrt(exp(theta^2 T) - 1), where
                  L = exp(theta W_T - theta^2 T / 2), theta = (mu_P - r_US)/sigma_1.
  reweighting     E^Q[X L] against E^P[X] on the barrier-monitored payoff X,
                  run on common random numbers so the two arms share every
                  Brownian increment, jump and bridge draw and differ only in
                  the drift.
"""
import os, sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lsmc_quanto import CAL                                            # noqa: E402

MU_P = 0.139


def _arm(c, npaths, seed, drift1):
    """One measure's arm; the seed fixes every draw, so arms are coupled."""
    n, T = c['steps'], c['T']
    dt = T/n; sq = np.sqrt(dt)
    kap = np.exp(c['thJ'] + 0.5*c['dlJ']**2) - 1.0
    m1 = drift1 - 0.5*c['vol1']**2 - c['lam']*kap
    m2 = (c['r_KRW'] - c['r_US']) - 0.5*c['vol2']**2
    lnV = c['vol1']**2*dt
    U, L = c['KOupper'], c['KOlower']
    rng = np.random.default_rng(seed)
    s1 = np.full(npaths, c['S1_0']); s2 = np.full(npaths, c['S2_0'])
    alive = np.ones(npaths, bool); W = np.zeros(npaths)
    for _ in range(n):
        prev = s1.copy()
        z1 = rng.standard_normal(npaths); z2 = rng.standard_normal(npaths)
        e2 = c['corr']*z1 + np.sqrt(1-c['corr']**2)*z2
        nJ = rng.poisson(c['lam']*dt, npaths)
        js = np.where(nJ > 0, rng.normal(c['thJ']*nJ, c['dlJ']*np.sqrt(np.maximum(nJ, 1))), 0.0)
        u = rng.random(npaths)
        W += sq*z1
        s1 = np.where(alive, s1*np.exp(m1*dt + c['vol1']*sq*z1 + js), s1)
        s2 = np.where(alive, s2*np.exp(m2*dt + c['vol2']*sq*e2), s2)
        alive &= ~((s1 >= U) | (s1 <= L))
        br = alive & (nJ == 0)
        pu = np.where(br, np.exp(-2*np.log(U/prev)*np.log(U/np.maximum(s1, 1e-9))/lnV), 0.)
        pl = np.where(br, np.exp(-2*np.log(prev/L)*np.log(np.maximum(s1, 1e-9)/L)/lnV), 0.)
        alive &= ~(br & (u < 1-(1-pu)*(1-pl)))
    X = np.where(alive, np.maximum(s1-c['S1_0'], 0)*s2, 0.0)
    return X, W


def validate(npaths=200_000, reps=30, seed0=4001, **over):
    c = dict(CAL); c.update(over)
    th = (MU_P - c['r_US'])/c['vol1']
    dm, rw, res, EL, sdL = [], [], [], [], []
    for k in range(reps):
        s = seed0 + k
        XQ, W = _arm(c, npaths, s, c['r_US'])
        XP, _ = _arm(c, npaths, s, MU_P)
        Lr = np.exp(th*W - 0.5*th*th*c['T'])
        a, b = float((XQ*Lr).mean()), float(XP.mean())
        dm.append(b); rw.append(a); res.append((a-b)/b)
        EL.append(float(Lr.mean())); sdL.append(float(Lr.std(ddof=1)))
    r = np.array(res)*100
    return dict(npaths=npaths, reps=reps, theta=th,
                direct_mean=float(np.mean(dm)), direct_sd=float(np.std(dm, ddof=1)),
                rw_mean=float(np.mean(rw)), rw_sd=float(np.std(rw, ddof=1)),
                resid_pct=float(r.mean()), resid_se=float(r.std(ddof=1)/np.sqrt(reps)),
                resid_sd=float(r.std(ddof=1)), resid_min=float(r.min()), resid_max=float(r.max()),
                EL=float(np.mean(EL)), sdL=float(np.mean(sdL)),
                sdL_theory=float(np.sqrt(np.exp(th*th*c['T'])-1)))


if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 200_000
    k = int(sys.argv[2]) if len(sys.argv) > 2 else 30
    v = validate(npaths=n, reps=k)
    print(f"theta = {v['theta']:.6f}   {n:,} paths x {k} replications, common random numbers")
    print(f"  density     E[L] = {v['EL']:.6f}   sd[L] = {v['sdL']:.6f}   theory {v['sdL_theory']:.6f}")
    print(f"  direct P    mean {v['direct_mean']:12,.2f}  across-rep sd {v['direct_sd']:9,.2f}")
    print(f"  reweighted  mean {v['rw_mean']:12,.2f}  across-rep sd {v['rw_sd']:9,.2f}")
    print(f"  residual    {v['resid_pct']:+.4f}% +/- {v['resid_se']:.4f}   "
          f"(rep sd {v['resid_sd']:.4f}, min {v['resid_min']:+.3f}, max {v['resid_max']:+.3f})")
