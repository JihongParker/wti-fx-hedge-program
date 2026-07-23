"""Delta hedging of the double knock-out quanto call: fitted surface vs closed-form proxy.

The LSMC backward induction of lsmc_quanto.py leaves a regression surface at
every step,

    V(S1,S2) ~ b0 + b1 v1 + b2 v2 + b3 v1^2 + b4 v2^2 + b5 v1 v2,
    v1 = S1/K,  v2 = S2/S2(0),

which is the object a desk actually hedges against: its gradient gives the two
deltas of the barrier structure directly,

    dV/dS1 = (b1 + 2 b3 v1 + b5 v2)/K,     dV/dS2 = (b2 + 2 b4 v2 + b5 v1)/S2(0).

The alternative a desk reaches for when it has no barrier model is the Black-76
delta of a vanilla call on the same strike, which is exactly correct for a
different option.  This module runs both hedges over the same path bank and
reports what the mismatch costs.
"""
import os, sys
import numpy as np
from math import log, sqrt, exp, erf

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lsmc_quanto import CAL, _paths                                    # noqa: E402

try:
    from scipy.special import ndtr as Ncdf
except ImportError:               # vectorised standard normal CDF
    _erf = np.vectorize(erf)
    Ncdf = lambda x: 0.5*(1+_erf(np.asarray(x, float)/sqrt(2)))


def fit_surface(K=None, npaths=200_000, seed=12345, **over):
    """Price by LSMC and keep the per-step regression coefficients."""
    c = dict(CAL); c.update(over)
    K = c['S1_0'] if K is None else K
    rng = np.random.default_rng(seed)
    S1, S2, AL = _paths(c, npaths, rng)
    n, disc = c['steps'], np.exp(-c['r_US']*c['T']/c['steps'])
    beta = np.zeros((n+1, 6))
    CF = np.where(AL[:, n], np.maximum(S1[:, n].astype(float)-K, 0)*S2[:, n], 0.0)
    for i in range(n-1, 0, -1):
        CF *= disc
        m = AL[:, i] & (S1[:, i] > K)
        if m.sum() > 10:
            v1 = S1[m, i].astype(float)/K
            v2 = S2[m, i].astype(float)/c['S2_0']
            X = np.column_stack([np.ones_like(v1), v1, v2, v1*v1, v2*v2, v1*v2])
            X = np.nan_to_num(X, posinf=0.0, neginf=0.0)
            b, *_ = np.linalg.lstsq(X, CF[m], rcond=1e-10)
            beta[i] = b
            cont = X @ b
            intr = (S1[m, i].astype(float)-K)*S2[m, i]
            ex = np.isfinite(cont) & (intr > cont)
            CF[np.where(m)[0][ex]] = intr[ex]
        else:
            beta[i] = beta[i+1]
    beta[0] = beta[1]; beta[n] = beta[n-1]
    return dict(price=float((CF*disc).mean()), beta=beta, K=K, cal=c)


def surface_delta(beta_i, S1, S2, K, S20):
    v1, v2 = S1/K, S2/S20
    d1 = (beta_i[1] + 2*beta_i[3]*v1 + beta_i[5]*v2)/K
    d2 = (beta_i[2] + 2*beta_i[4]*v2 + beta_i[5]*v1)/S20
    return d1, d2


def b76_delta(S1, K, sig, tau, r):
    tau = np.maximum(tau, 1e-8)
    d1 = (np.log(np.maximum(S1, 1e-9)/K)+0.5*sig*sig*tau)/(sig*np.sqrt(tau))
    return np.exp(-r*tau)*Ncdf(d1)


def run(fit, npaths=200_000, seed=777, scale=1.0, proxy=False, Qo=2_000_000, r_w=0.07):
    """Daily-rebalanced delta hedge of a short KO quanto position on Qo barrels.

    Hedge instruments: WTI futures (h1 barrels) and a USD forward (h2 dollars).
    Returns the hedge cost per path in KRW, defined as the terminal payoff owed
    less the premium carried forward and less what the hedge portfolio earned.
    """
    c, K, beta = fit['cal'], fit['K'], fit['beta']
    n = c['steps']; dt = c['T']/n; S20 = c['S2_0']; gr = exp(r_w*dt)
    rng = np.random.default_rng(seed)
    S1, S2, AL = _paths(c, npaths, rng)
    S1 = S1.astype(float); S2 = S2.astype(float)

    cash = np.zeros(npaths); h1 = np.zeros(npaths); h2 = np.zeros(npaths)
    for i in range(n):
        if proxy:
            g1 = Qo*b76_delta(S1[:, i], K, c['vol1'], c['T']-i*dt, c['r_US'])
            g2 = np.zeros(npaths)                     # a vanilla proxy has no FX leg
        else:
            dV1, dV2 = surface_delta(beta[i], S1[:, i], S2[:, i], K, S20)
            # the surface is fitted on in-the-money survivors only, so its
            # gradient is used only where it was fitted; elsewhere the barrier
            # call is out of the money and carries no first-order exposure
            dom = AL[:, i] & (S1[:, i] > K)
            g1 = Qo*np.where(dom, np.clip(dV1/S2[:, i], 0.0, 1.0), 0.0)
            g2 = Qo*np.where(dom, np.clip(dV2, -20.0, 20.0), 0.0)
        alive = AL[:, i]
        g1 = np.where(alive, scale*g1, 0.0); g2 = np.where(alive, scale*g2, 0.0)
        cash -= (g1-h1)*S1[:, i]*S2[:, i] + (g2-h2)*S2[:, i]
        h1, h2 = g1, g2
        cash *= gr
    cash += h1*S1[:, n]*S2[:, n] + h2*S2[:, n]
    payoff = np.where(AL[:, n], np.maximum(S1[:, n]-K, 0)*S2[:, n], 0.0)*Qo
    return payoff - fit['price']*Qo*exp(r_w*c['T']) - cash


if __name__ == '__main__':
    npaths = int(sys.argv[1]) if len(sys.argv) > 1 else 200_000
    fit = fit_surface(npaths=npaths)
    print(f"LSMC price {fit['price']:,.2f} KRW/bbl on {npaths:,} paths")
    for lbl, pr in (('fitted surface delta', False), ('Black-76 proxy delta', True)):
        h = run(fit, npaths=npaths, proxy=pr)
        print(f"  {lbl:22s} hedge cost mean {h.mean()/1e9:8.2f} bn   sd {h.std(ddof=1)/1e9:8.2f} bn")
    cs = np.linspace(-1.5, 1.5, 25)
    sds = [run(fit, npaths=50_000, scale=cc).std(ddof=1) for cc in cs]
    j = int(np.argmin(sds))
    print(f"  variance-minimising hedge scale c* = {cs[j]:+.3f}  (sd {sds[j]/1e9:.2f} bn)")
