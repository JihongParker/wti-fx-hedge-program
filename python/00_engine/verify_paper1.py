"""Reproduce every headline number of Paper 1 from source and report the gap.

Run:  python3 python/00_engine/verify_paper1.py [npaths]
No spreadsheet is read.  Closed forms are computed here; the American premiums
come from lsmc_quanto.py and the stress mortality from jump_barrier.py.
"""
import json, os, sys
from math import log, sqrt, exp, erf
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from lsmc_quanto import shapley, price as ko_price          # noqa: E402
from jump_barrier import touch_prob                          # noqa: E402

I = json.load(open(os.path.join(HERE, '..', '..', 'data', 'results', 'opt_inputs.json')))['inputs']
N = lambda x: 0.5*(1+erf(x/sqrt(2)))


def black76(S, K, s, T, r):
    d1 = (log(S/K)+0.5*s*s*T)/(s*sqrt(T)); d2 = d1-s*sqrt(T)
    return exp(-r*T)*(S*N(d1)-K*N(d2))


def gk(S, K, s, T, rd, rf):
    d1 = (log(S/K)+(rd-rf+0.5*s*s)*T)/(s*sqrt(T)); d2 = d1-s*sqrt(T)
    return S*exp(-rf*T)*N(d1)-K*exp(-rd*T)*N(d2)


def sigma_res(w1, w2, s1, s2, rho):
    u, v = 1-w1, 1-w2
    return sqrt(u*u*s1*s1+v*v*s2*s2+2*u*v*s1*s2*rho)


def main(npaths=200_000):
    S1, S2 = I['WTI_spot'], I['KRW_spot']
    s1E, s1A, s2, rho = (I['sigma_WTI_hist'], I['sigma_WTI_diff'],
                         I['sigma_FX'], I['rho'])
    rUS, rKR, rw = I['r_US'], I['r_KRW'], I['WACC']
    T1, T2 = I['Maturity_Oil'], I['Maturity_FX']
    Qo, Qu, B = I['Monthly_Oil_Need'], I['Monthly_USD_Need'], I['Max_Budget']
    M1 = Qo*max(0, I['Stress_WTI']-S1)*I['Stress_KRW']
    M2 = Qu*max(0, I['Stress_KRW']-S2)

    rows = []
    add = rows.append
    P_B76 = black76(S1, 0.95*S1, s1E, T1, rUS)
    P_GK = gk(S2, 0.95*S2, s2, T2, rKR, rUS)
    add(('P_B76  (K = 0.95 S, USD/bbl)', 12.6524, P_B76, 'closed form'))
    add(('P_GK   (K = 0.95 S, KRW/USD)', 84.667, P_GK, 'closed form'))

    sh = shapley(npaths=npaths)
    add(('joint quanto premium (KRW/bbl)', 17132.47, sh['base'], f'LSMC {npaths:,}p'))
    add(('Shapley WTI share (ratio)', 0.881003, sh['ratio_wti'], f'LSMC {npaths:,}p'))
    add(('Shapley FX share (ratio)', 0.118997, sh['ratio_fx'], f'LSMC {npaths:,}p'))
    add(('P_K standalone KO (KRW/bbl)', 17377.11, sh['v_wti_only'], f'LSMC {npaths:,}p'))

    p, se = touch_prob(I['Stress_WTI'], npaths=npaths)
    add(('p_KO | stress spot 113', 0.8925, p, f'first passage {npaths:,}p'))

    K1E = Qo*P_B76*S2*(1+rw*T1); K2E = Qu*P_GK*(1+rw*T2)
    add(('European WTI full cover (KRW)', 41258998746, K1E, 'from P_B76'))
    add(('European FX full cover (KRW)', 13835108654, K2E, 'from P_GK'))

    # sigma_res takes the raw historical volatility in both regimes
    wl = (s1E**2-rho*s1E*s2)/(s1E**2+s2**2-2*rho*s1E*s2)
    add(('American risk-min w1 (line GMVP)', 0.965980, wl, 'closed form'))
    add(('American sigma_res', 0.091585, sigma_res(wl, 1-wl, s1E, s2, rho), 'closed form'))

    # on w1+w2=1 the European ledger is strictly decreasing in w1, so the
    # unconstrained line minimiser is priced out and the optimum slides up to
    # the point where the ledger meets the cap
    cE = lambda w1: w1*K1E+(1-w1)*K2E+(1-w1)*M1+w1*M2
    wlE = (s1E**2-rho*s1E*s2)/(s1E**2+s2**2-2*rho*s1E*s2)
    if cE(wlE) <= B:
        wE = wlE
    else:
        lo, hi = wlE, 1.0
        for _ in range(200):
            m = (lo+hi)/2
            if cE(m) > B: lo = m
            else: hi = m
        wE = hi
    add(('European risk-min w1 (vertex)', 0.970486, wE, 'budget x simplex'))
    add(('European sigma_res', 0.0916021, sigma_res(wE, 1-wE, s1E, s2, rho), 'closed form'))
    add(('European ledger at the vertex', 45000000000.0, cE(wE), 'budget binds')) 

    PK = sh['v_wti_only']*Qo*(1+rw*T1)      # same funding factor as P_V
    add(('pbar  (P_V - P_K)/UL_WTI', 0.042413, (K1E-PK)/M1, 'from engine P_K'))
    add(('p*    (B - P_K - UL_FX)/UL_WTI', 0.063848, (B-PK-M2)/M1, 'from engine P_K'))

    w = max(len(r[0]) for r in rows)
    print(f"{'quantity'.ljust(w)}  {'paper':>18}  {'from source':>18}   gap      source")
    print('-'*(w+66))
    for name, paper, got, src in rows:
        g = (got-paper)/paper if paper else 0.0
        flag = 'ok ' if abs(g) < 0.005 else ('~  ' if abs(g) < 0.03 else 'X  ')
        print(f"{name.ljust(w)}  {paper:18,.6f}  {got:18,.6f}  {flag}{g:+7.3%}  {src}")


if __name__ == '__main__':
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 200_000)
