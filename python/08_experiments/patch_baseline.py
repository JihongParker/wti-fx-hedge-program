"""Paper 2 headline numbers, as-shipped vs patched.

Three defects were live in the engine and each moves a headline:
  K default    lsmc_quanto.price / delta_hedge.fit_surface let K follow an
               overridden S1_0, so a finite difference measured a joint
               spot-and-strike translation rather than a delta.
  no exercise  run() paid the terminal payoff while fit['price'] is an American
               LSMC value, charging a European ledger an American premium.
  no FX net    the futures leg carries g1*S1 dollars; the FX leg was set to
               Qo*dV2 on top of it, roughly a 5x over-hedge.

Run: python3 python/08_experiments/patch_baseline.py [npaths]
"""
import sys, os, numpy as np
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '00_engine'))
from delta_hedge import fit_surface, run
from lsmc_quanto import CAL, price

N = int(sys.argv[1]) if len(sys.argv) > 1 else 60_000
fit = fit_surface(npaths=200_000, seed=12345)
bn = lambda x: x/1e9
print(f"surface fitted on 200,000 paths, price {fit['price']:,.2f} KRW/bbl; hedge bank {N:,}\n")

CONFIG = [("as shipped",        dict(net_fx=False, exercise=False)),
          ("+ exercise",        dict(net_fx=False, exercise=True)),
          ("+ exercise + FX net", dict(net_fx=True,  exercise=True))]

print(f"{'configuration':22} {'surface sd':>11} {'B76 sd':>9} {'winner':>10}")
for lbl, kw in CONFIG:
    s = run(fit, npaths=N, seed=777, **kw).std(ddof=1)
    b = run(fit, npaths=N, seed=777, proxy=True,
            exercise=kw['exercise'], net_fx=False).std(ddof=1)
    print(f"{lbl:22} {bn(s):11,.2f} {bn(b):9,.2f} {'surface' if s<b else 'Black-76':>10}")

print(f"\n{'configuration':22} {'c*':>8} {'sd(c=1)':>9} {'sd(c*)':>9} {'reduction':>10}")
# the paper's c multiplies the FX LEG ONLY: delta_FX = c * delta_WTI.
# run(scale=) multiplies both legs, so it is a different object; sweep c directly.
import delta_hedge as dh
from lsmc_quanto import _paths as _p
from math import exp as _exp
def fx_sweep(net_fx, exercise, cc):
    c, K, beta = fit['cal'], fit['K'], fit['beta']
    n = c['steps']; dt = c['T']/n; S20 = c['S2_0']; gr = _exp(0.07*dt); Qo = 2_000_000
    rng = np.random.default_rng(777); S1, S2, AL = _p(c, N, rng)
    S1 = S1.astype(float); S2 = S2.astype(float)
    cash = np.zeros(N); h1 = np.zeros(N); h2 = np.zeros(N)
    live = np.ones(N, bool); booked = np.zeros(N)
    for i in range(n):
        if exercise and i > 0:
            v1 = S1[:, i]/K; v2 = S2[:, i]/S20
            X = np.column_stack([np.ones(N), v1, v2, v1*v1, v2*v2, v1*v2])
            cont = np.nan_to_num(X) @ beta[i]; intr = np.maximum(S1[:, i]-K, 0)*S2[:, i]
            stop = live & AL[:, i] & (S1[:, i] > K) & np.isfinite(cont) & (intr > cont)
            if stop.any():
                cash[stop] += h1[stop]*S1[stop, i]*S2[stop, i] + h2[stop]*S2[stop, i]
                booked[stop] = intr[stop]*Qo; h1[stop] = 0.0; h2[stop] = 0.0; live[stop] = False
        dV1, dV2 = dh.surface_delta(beta[i], S1[:, i], S2[:, i], K, S20)
        dom = AL[:, i] & (S1[:, i] > K) & live
        dW = np.clip(dV1/S2[:, i], 0.0, 1.0)
        g1 = Qo*np.where(dom, dW, 0.0)
        fx = cc*Qo*dW*S1[:, i]                       # coupled family: delta_FX = c*delta_WTI
        if net_fx: fx = fx - g1*S1[:, i]
        g2 = np.where(dom, fx, 0.0)
        cash -= (g1-h1)*S1[:, i]*S2[:, i] + (g2-h2)*S2[:, i]; h1, h2 = g1, g2; cash *= gr
    cash += h1*S1[:, n]*S2[:, n] + h2*S2[:, n]
    booked[live] = np.where(AL[live, n], np.maximum(S1[live, n]-K, 0)*S2[live, n], 0.0)*Qo
    return booked - fit['price']*Qo*_exp(0.07*c['T']) - cash
for lbl, kw in CONFIG:
    cs = np.linspace(-5.0, 2.0, 29)
    sds = [fx_sweep(kw['net_fx'], kw['exercise'], x).std(ddof=1) for x in cs]
    j = int(np.argmin(sds)); k = max(1, min(len(cs)-2, j))
    a, b_, _ = np.polyfit(cs[k-1:k+2], np.array(sds[k-1:k+2])**2, 2)
    one = fx_sweep(kw['net_fx'], kw['exercise'], 1.0).std(ddof=1)
    print(f"{lbl:22} {-b_/(2*a):8.3f} {bn(one):9,.2f} {bn(sds[j]):9,.2f} {100*(1-sds[j]/one):9.2f}%")

print("\nK guard: price(S1_0=...) without K now raises")
try:
    price(npaths=1000, S1_0=CAL['S1_0']*1.05); print("  NOT RAISED - guard failed")
except ValueError as e:
    print(f"  raised: {str(e)[:60]}...")
