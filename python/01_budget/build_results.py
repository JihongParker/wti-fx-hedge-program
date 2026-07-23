"""Solve both allocation programs and write data/results/corrected_results.json.

sigma_res takes the raw historical WTI volatility in both regimes: it measures
the exposure left uncovered, and that exposure carries the jumps whichever
instrument was bought.  The jump-stripped diffusive volatility is a pricing
input and appears only inside the American premium constants.
"""
import json, os
import numpy as np


def _find(name):
    here = os.path.dirname(os.path.abspath(__file__))
    for c in (os.path.join(here, '..', '..', 'data', 'results', name), name):
        if os.path.exists(c):
            return c
    raise FileNotFoundError(name)


I = json.load(open(_find('opt_inputs.json')))['inputs']
S1, S2, RHO = I['sigma_WTI_hist'], I['sigma_FX'], I['rho']
Qo, Qu, B = I['Monthly_Oil_Need'], I['Monthly_USD_Need'], I['Max_Budget']
rw, T1, T2 = I['WACC'], I['Maturity_Oil'], I['Maturity_FX']
M1 = Qo*max(0, I['Stress_WTI']-I['WTI_spot'])*I['Stress_KRW']
M2 = Qu*max(0, I['Stress_KRW']-I['KRW_spot'])
K1E = Qo*I['Black76_call']*I['KRW_spot']*(1+rw*T1)
K2E = Qu*I['GK_call']*(1+rw*T2)
# both regimes carry cash premiums on the same simple-interest factor (1 + r_w T);
# the American legs previously used exp(r_w T), a 0.16% difference on K1A
K1A = I['LSMC_WTI_shapley']*Qo*(1+rw*T1)
K2A = I['LSMC_FX_shapley']*Qo*(1+rw*T2)

gm = lambda w1, w2: np.sqrt((1-w1)**2*S1*S1 + (1-w2)**2*S2*S2 + 2*(1-w1)*(1-w2)*S1*S2*RHO)
cE = lambda w1, w2: K1E*w1 + K2E*w2 + M1*(1-w1) + M2*(1-w2)
cA = lambda w1, w2: K1A*w1 + K2A*w2 + M1*(1-w1) + M2*(1-w2)
WL = (S1**2 - RHO*S1*S2)/(S1**2 + S2**2 - 2*RHO*S1*S2)


def solve(cost, budget):
    """min sigma_res over 0<=w<=1, w1+w2<=1, cost<=budget."""
    if cost(WL, 1-WL) <= budget:
        w = WL
    else:
        lo, hi = WL, 1.0
        for _ in range(300):
            m = (lo+hi)/2
            if cost(m, 1-m) > budget: lo = m
            else: hi = m
        w = hi
        if cost(w, 1-w) > budget + 1e-3:
            return None
    return dict(w1=w, w2=1-w, gmvp=float(gm(w, 1-w)), cost=float(cost(w, 1-w)))


def block(cost):
    rm = solve(cost, B)
    nb = solve(cost, np.inf)
    eps = 1e9
    up = solve(cost, B+eps)
    shadow = (rm['gmvp']-up['gmvp'])/1.0 if up else 0.0
    cm = dict(w1=1.0, w2=0.0, gmvp=float(gm(1, 0)), cost=float(cost(1, 0)))
    sweep = []
    for b in np.linspace(0.75*B, 1.30*B, 50):
        s = solve(cost, b)
        sweep.append(dict(B=float(b), gmvp=s['gmvp'] if s else None,
                          w1=s['w1'] if s else None, w2=s['w2'] if s else None))
    g0 = gm(0, 0)
    grad = ((-(S1*S1 + S1*S2*RHO)/g0), (-(S2*S2 + S1*S2*RHO)/g0))
    return dict(risk_min=rm, no_budget=nb, shadow_per_bn=float(shadow), cost_min=cm,
                cost_min_cap=rm, line_gmvp=dict(w1=WL, w2=1-WL, gmvp=float(gm(WL, 1-WL))),
                sweep=sweep, grad00=list(map(float, grad)), gmvp00=float(g0))


pm = 0.8925
out = dict(EU=block(cE), AM=block(cA),
           pko=dict(p_star=float((B-(K1A+M2))/M1),
                    min_cost_at={str(round(p, 5)): float(K1A + p*M1 + M2)
                                 for p in (0.0, 0.10913, 0.2309, 0.4369, pm)},
                    sadj_at_riskmin={str(round(p, 5)): float(
                        K1A*WL + K2A*(1-WL) + (1-WL*(1-p))*M1 + (1-(1-WL)*(1-p))*M2)
                        for p in (0.0, 0.05, 0.10913, 0.15, 0.2309, 0.3, 0.4369, pm)}),
           coeffs=dict(K1E=K1E, K2E=K2E, K1A=K1A, K2A=K2A, M1=M1, M2=M2))
json.dump(out, open(_find('corrected_results.json'), 'w'), indent=1)

for k in ('EU', 'AM'):
    r = out[k]['risk_min']
    print(f"{k}: w=({r['w1']:.6f}, {r['w2']:.6f})  cost {r['cost']:,.0f}  sigma {r['gmvp']:.6f}"
          f"   no-budget cost {out[k]['no_budget']['cost']:,.0f}")
print(f"line GMVP w1 {WL:.6f}   sigma(0,0) {out['EU']['gmvp00']:.4f}")
print(f"p*_shapley {out['pko']['p_star']:.5f}   sadj at pm {out['pko']['sadj_at_riskmin'][str(pm)]/1e9:.2f}bn")
