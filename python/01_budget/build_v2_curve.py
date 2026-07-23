"""Rebuild v2_curve.json: the Section 8 mixed vanilla/KO program.

Solves eq. mixprog/mixledger directly from the frozen coefficients in
corrected_results.json, so every threshold the paper quotes for Section 8
(pfloor, pbar, p*) is derived here rather than carried as a literal.
"""
import json, os
import numpy as np

os.chdir('/Users/elijahjasper/modeling/Ratio_Optimization')
I = json.load(open('opt_scraped.json'))['inputs']
R = json.load(open('corrected_results.json'))
C = R['coeffs']

Qo, Qu = I['Monthly_Oil_Need'], I['Monthly_USD_Need']
B, rw, T1, T2 = I['Max_Budget'], I['WACC'], I['Maturity_Oil'], I['Maturity_FX']
s1A, s2, rho = I['sigma_WTI_diff'], I['sigma_FX'], I['rho']

PV, PGK = C['K1E'], C['K2E']        # vanilla WTI / FX full cover (eq. costeu)
PK      = 17377.11*Qo               # standalone LSMC KO premium, paper Sec.8
M1, M2  = C['M1'], C['M2']          # UL_WTI, UL_FX
K1A, K2A = C['K1A'], C['K2A']       # Shapley-attributed American coefficients
print(f"PV={PV:,.0f}  PK={PK:,.0f}  PGK={PGK:,.0f}  M1={M1:,.0f}  M2={M2:,.0f}")
print(f"K1A={K1A:,.0f}  K2A={K2A:,.0f}")

def sig(w1, w2):
    u, v = 1-w1, 1-w2
    return np.sqrt(u*u*s1A*s1A + v*v*s2*s2 + 2*u*v*s1A*s2*rho)

def solve(Pw, p):
    """min sigma s.t. w1*Pw + w2*PGK + (1-w1*(1-p))*M1 + (1-w2)*M2 <= B, w1+w2<=1."""
    def cost(w1, w2):
        return w1*Pw + w2*PGK + (1-w1*(1-p))*M1 + (1-w2)*M2
    best = None
    for w2 in np.linspace(0, 1, 20001):
        hi = 1-w2
        if cost(hi, w2) <= B:
            w1 = hi
        else:
            lo = 0.0
            if cost(lo, w2) > B:
                continue
            for _ in range(80):
                mid = (lo+hi)/2
                if cost(mid, w2) <= B: lo = mid
                else: hi = mid
            w1 = lo
        s = sig(w1, w2)
        if best is None or s < best[2]:
            best = (w1, w2, s)
    return best

# branch B (all-vanilla): p-independent
bV = solve(PV, 0.0)
sigV = bV[2]
print(f"vanilla branch: w=({bV[0]:.6f}, {bV[1]:.6f})  sigma={sigV:.7f}")

# closed-form thresholds
pcross = (PV - PK)/M1                       # instrument indifference  (eq. pbar)
pstar  = (B - PK - M2)/M1                   # pure-KO (standalone) infeasible
pstar_shapley = (B - K1A - M2)/M1           # pure-KO (Shapley-priced, Sec.7 eq. pstar)
wl1 = (s1A**2 - rho*s1A*s2)/(s1A**2 + s2**2 - 2*rho*s1A*s2); wl2 = 1-wl1
pfloor = (B - (wl1*PK + wl2*PGK + (1-wl1)*M1 + (1-wl2)*M2)) / (wl1*M1)
print(f"pfloor={pfloor:.6f}  pcross={pcross:.6f}  pstar(standalone)={pstar:.6f}  pstar(Shapley,S7)={pstar_shapley:.6f}")
print(f"line-GMVP w=({wl1:.6f},{wl2:.6f}) sigma={sig(wl1,wl2):.7f}")

curve = []
for p in np.linspace(0, 0.15, 101):
    try:
        b = solve(PK, float(p))
        curve.append([float(p), None if b is None else float(b[2])])
    except Exception:
        curve.append([float(p), None])

PM = 0.8925
out = {'sigV': float(sigV), 'pfloor': float(pfloor), 'pcross': float(pcross),
       'pstar': float(pstar), 'pstar_shapley': float(pstar_shapley),
       'curve': curve, 'pm': PM,
       'minatt_at_pm': float(K1A + PM*M1 + M2),
       'adopted_adj_at_pm': float(K1A*0.945202 + K2A*0.054798
                                  + (1-0.945202*(1-PM))*M1 + (1-0.054798*(1-PM))*M2)}
json.dump(out, open('v2_curve.json','w'), indent=1)
print(f"minatt@pm={out['minatt_at_pm']/1e9:.2f}bn  adopted@pm={out['adopted_adj_at_pm']/1e9:.2f}bn")
print("curve last feasible p:", max(p for p,s in curve if s is not None))
