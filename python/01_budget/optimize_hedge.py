"""Python reimplementation of the workbook's two static hedge-ratio programs.

Reproduces, from the verbatim workbook formulas pinned in opt_scraped.json:
  (1) risk-min:  min gmvp(w1,w2)  s.t. box, w1+w2<=1, TotalCost<=Budget
      -- cross-validated against the Excel Evolutionary-solver solutions
         stored at Encoding!C21:C24 and LSMC!J13:J15.
  (2) budget relaxation sweep (is the budget constraint binding?).
  (3) cost-min:  min TotalCost(w1,w2)  s.t. box, w1+w2<=1
      variant A: no risk constraint; variant B: gmvp <= gmvp_Excel.
  (4) analytic mechanism: line-restricted GMVP weights, gradients.

All outputs go to opt_results.json; figure grids go to opt_grids.npz.
"""
import json
import numpy as np
from scipy.optimize import minimize

with open('opt_scraped.json', encoding='utf-8') as f:
    D = json.load(f)
I = D['inputs']

BUDGET = I['Max_Budget']

# ----------------------------------------------------------------------
# model functions -- verbatim transcriptions of the stored formulas
# ----------------------------------------------------------------------
S1H, S2, RHO = I['sigma_WTI_hist'], I['sigma_FX'], I['rho']
S1D = I['sigma_WTI_diff']

def gmvp(w, s1, s2=S2, rho=RHO):
    w1, w2 = w
    return np.sqrt((1-w1)**2*s1**2 + (1-w2)**2*s2**2
                   + 2*(1-w1)*(1-w2)*s1*s2*rho)

def gmvp_eu(w): return gmvp(w, S1H)
def gmvp_am(w): return gmvp(w, S1D)

# European cost (Pricing!C1:C9)
K1_EU = I['Monthly_Oil_Need']*I['Black76_call']*I['KRW_spot']*(1+I['WACC']*I['Maturity_Oil'])
K2_EU = I['Monthly_USD_Need']*I['GK_call']*(1+I['WACC']*I['Maturity_FX'])
M1_EU = I['Monthly_Oil_Need']*max(0.0, I['Stress_WTI']-I['WTI_spot'])*I['Stress_KRW']
M2_EU = I['Monthly_USD_Need']*max(0.0, I['Stress_KRW']-I['KRW_spot'])

def cost_eu(w):
    w1, w2 = w
    return K1_EU*w1 + K2_EU*w2 + M1_EU*(1-w1) + M2_EU*(1-w2)

# American cost (Pricing!C11:C19) -- C15 transcribed exactly as stored,
# i.e. 1 - w1*Need*max(...), NOT (1-w1)*Need*max(...)
K1_AM = I['LSMC_WTI_shapley']*I['Monthly_Oil_Need']*np.exp(I['WACC']*I['Maturity_Oil'])
K2_AM = I['LSMC_FX_shapley']*I['Monthly_Oil_Need']*np.exp(I['Maturity_FX']*I['WACC'])
FXGAP = max(0.0, I['Stress_KRW']-I['KRW_spot'])

def cost_am(w):
    w1, w2 = w
    # eq:costam — both regimes share the stress ledger (M1_EU is the WTI
    # spike loss, M2_EU the FX-gap loss). Earlier code had c15 = 1.0 - w1*...
    # (an Excel-cell transcription slip using the FX gap for the WTI leg);
    # the paper's stress term is (1-w1)*M1_EU, restored here.
    c11 = K1_AM*w1
    c12 = K2_AM*w2
    c15 = (1-w1)*M1_EU
    c16 = (1-w2)*M2_EU
    return c11 + c12 + c15 + c16

ENGINES = {
    'european': dict(gm=gmvp_eu, cost=cost_eu, s1=S1H,
                     excel=D['european']),
    'american': dict(gm=gmvp_am, cost=cost_am, s1=S1D,
                     excel=D['american']),
}

# ----------------------------------------------------------------------
# generic constrained solve (multi-start SLSQP)
# ----------------------------------------------------------------------
def solve(objective, constraints, n_starts=64, seed=0):
    rng = np.random.default_rng(seed)
    best = None
    starts = [np.array([a, b]) for a in (0.05, 0.5, 0.95) for b in (0.05, 0.5, 0.95)]
    starts += list(rng.uniform(0, 1, size=(n_starts, 2)))
    for x0 in starts:
        r = minimize(objective, x0, method='SLSQP',
                     bounds=[(0, 1), (0, 1)], constraints=constraints,
                     options={'maxiter': 500, 'ftol': 1e-14})
        if not r.success:
            continue
        feas = all(c['fun'](r.x) >= -1e-7 for c in constraints)
        if feas and (best is None or r.fun < best.fun):
            best = r
    return best

def cons_sum():
    return {'type': 'ineq', 'fun': lambda w: 1.0 - (w[0]+w[1])}

def cons_budget(costf, budget):
    # scaled so SLSQP sees O(1) numbers
    return {'type': 'ineq', 'fun': lambda w: (budget - costf(w))/1e9}

def cons_risk(gm, cap):
    return {'type': 'ineq', 'fun': lambda w: cap - gm(w)}

R = {'inputs_summary': {
        'K1_EU': K1_EU, 'K2_EU': K2_EU, 'M1_EU': M1_EU, 'M2_EU': M2_EU,
        'K1_AM': K1_AM, 'K2_AM': K2_AM, 'FXGAP_notional_FX': I['Monthly_USD_Need']*FXGAP,
    }}

print('=== cost-function coefficients (bn KRW per unit ratio) ===')
print(f'EU: dCost/dw1 = {(K1_EU-M1_EU)/1e9:+.3f}   dCost/dw2 = {(K2_EU-M2_EU)/1e9:+.3f}')
print(f'AM: dCost/dw1 = {(K1_AM - I["Monthly_Oil_Need"]*FXGAP)/1e9:+.3f}   dCost/dw2 = {(K2_AM - I["Monthly_USD_Need"]*FXGAP)/1e9:+.3f}')

# ----------------------------------------------------------------------
# (1) risk-min replication + (2) budget analysis + (3) cost-min variants
# ----------------------------------------------------------------------
for name, E in ENGINES.items():
    gm, costf, xl = E['gm'], E['cost'], E['excel']
    out = {}

    # --- (1) replicate the Excel risk-min program exactly -------------
    cons = [cons_sum(), cons_budget(costf, BUDGET)]
    r = solve(gm, cons)
    w = r.x
    out['risk_min'] = {
        'w1': w[0], 'w2': w[1], 'w_sum': w[0]+w[1],
        'gmvp': gm(w), 'cost': costf(w),
        'excel_w1': xl['w1'], 'excel_w2': xl['w2'],
        'excel_gmvp': xl['gmvp'], 'excel_cost': xl['total_cost'],
        'dw1': w[0]-xl['w1'], 'dw2': w[1]-xl['w2'],
        'gmvp_gap': gm(w)-xl['gmvp'],
        'gmvp_at_excel_point': gm([xl['w1'], xl['w2']]),
    }
    # which constraints are active at the Python optimum?
    out['risk_min']['active'] = {
        'sum': abs(1-(w[0]+w[1])) < 1e-6,
        'budget': abs(BUDGET-costf(w)) < 1e-3*BUDGET,
    }

    # --- (2) budget relaxed: is the budget binding? --------------------
    r_free = solve(gm, [cons_sum()])
    wf = r_free.x
    out['risk_min_no_budget'] = {
        'w1': wf[0], 'w2': wf[1], 'gmvp': gm(wf), 'cost': costf(wf),
        'budget_needed': costf(wf),
        'gmvp_improvement': out['risk_min']['gmvp'] - gm(wf),
    }
    # budget sweep for the shadow-price curve
    sweep = []
    for B in np.arange(40e9, 52.0000001e9, 0.25e9):
        rb = solve(gm, [cons_sum(), cons_budget(costf, B)], n_starts=24, seed=1)
        if rb is None:
            sweep.append({'budget': B, 'feasible': False})
            continue
        sweep.append({'budget': B, 'feasible': True,
                      'w1': rb.x[0], 'w2': rb.x[1],
                      'gmvp': gm(rb.x), 'cost': costf(rb.x)})
    out['budget_sweep'] = sweep
    # shadow price at the stored budget (finite difference, +-0.1bn)
    eps = 0.1e9
    g_hi = solve(gm, [cons_sum(), cons_budget(costf, BUDGET+eps)], n_starts=24, seed=2)
    g_lo = solve(gm, [cons_sum(), cons_budget(costf, BUDGET-eps)], n_starts=24, seed=3)
    out['shadow_price_per_bnKRW'] = (gm(g_lo.x)-gm(g_hi.x))/(2*eps/1e9)

    # --- (3) cost-min variants -----------------------------------------
    rA = solve(lambda w: costf(w)/1e9, [cons_sum()])
    out['cost_min_A'] = {'w1': rA.x[0], 'w2': rA.x[1],
                         'cost': costf(rA.x), 'gmvp': gm(rA.x)}
    # corner enumeration double-check (cost is linear => corner optimum)
    corners = [(0,0),(1,0),(0,1),(0.5,0.5)]
    out['cost_corners'] = {f'({a},{b})': costf([a,b]) for a,b in corners}

    cap = xl['gmvp']
    rB = solve(lambda w: costf(w)/1e9, [cons_sum(), cons_risk(gm, cap)])
    out['cost_min_B'] = {'w1': rB.x[0], 'w2': rB.x[1],
                         'cost': costf(rB.x), 'gmvp': gm(rB.x),
                         'risk_cap': cap,
                         'saving_vs_excel_cost': xl['total_cost']-costf(rB.x)}

    # --- (4) analytic line-GMVP (mechanism for question b) -------------
    s1, s2, rho = E['s1'], S2, RHO
    # on w1+w2=1 with h=1-w: h2 = w1, h1 = 1-w1; classic 2-asset GMVP
    w1_star = (s1**2 - s1*s2*rho)/(s1**2 + s2**2 - 2*s1*s2*rho)
    out['line_gmvp'] = {
        'w1_star': w1_star, 'w2_star': 1-w1_star,
        'gmvp': gm([w1_star, 1-w1_star]),
        'cost': costf([w1_star, 1-w1_star]),
        'exceeds_budget': costf([w1_star, 1-w1_star]) > BUDGET,
    }
    # gradient of gmvp at full-unhedged (0,0) and at the optimum
    def grad_gm(w):
        h = 1e-7
        return [(gm([w[0]+h, w[1]])-gm([w[0]-h, w[1]]))/(2*h),
                (gm([w[0], w[1]+h])-gm([w[0], w[1]-h]))/(2*h)]
    out['grad_at_00'] = grad_gm([0.0, 0.0])
    out['grad_at_opt'] = grad_gm(w)

    R[name] = out
    print(f"\n=== {name} ===")
    print(f"risk-min  python w=({w[0]:.6f},{w[1]:.6f}) gmvp={gm(w):.8f} cost={costf(w):,.0f}")
    print(f"          excel  w=({xl['w1']:.6f},{xl['w2']:.6f}) gmvp={xl['gmvp']:.8f} cost={xl['total_cost']:,.0f}")
    print(f"          active constraints: {out['risk_min']['active']}")
    print(f"no-budget w=({wf[0]:.6f},{wf[1]:.6f}) gmvp={gm(wf):.8f} cost={costf(wf):,.0f}  (needs {costf(wf)/1e9:.2f}bn)")
    print(f"shadow price: {out['shadow_price_per_bnKRW']:.3e} gmvp per bn KRW")
    print(f"cost-min A w=({rA.x[0]:.4f},{rA.x[1]:.4f}) cost={costf(rA.x):,.0f} gmvp={gm(rA.x):.6f}")
    print(f"cost-min B w=({rB.x[0]:.6f},{rB.x[1]:.6f}) cost={costf(rB.x):,.0f} gmvp={gm(rB.x):.8f} (cap {cap:.8f})")
    print(f"line GMVP w1*={w1_star:.6f} cost={out['line_gmvp']['cost']:,.0f} > budget: {out['line_gmvp']['exceeds_budget']}")

# ----------------------------------------------------------------------
# grids for the 3-D figure (full [0,1]^2, both engines)
# ----------------------------------------------------------------------
n = 201
w1g = np.linspace(0, 1, n)
w2g = np.linspace(0, 1, n)
W1, W2 = np.meshgrid(w1g, w2g)

def grid(gm_s1, costf):
    H1, H2 = 1-W1, 1-W2
    G = np.sqrt(H1**2*gm_s1**2 + H2**2*S2**2 + 2*H1*H2*gm_s1*S2*RHO)
    C = np.array([[costf([a, b]) for a in w1g] for b in w2g])
    F = (W1+W2 <= 1+1e-12) & (C <= BUDGET)
    return G, C, F

G_EU, C_EU, F_EU = grid(S1H, cost_eu)
G_AM, C_AM, F_AM = grid(S1D, cost_am)
np.savez('opt_grids.npz', W1=W1, W2=W2,
         G_EU=G_EU, C_EU=C_EU, F_EU=F_EU,
         G_AM=G_AM, C_AM=C_AM, F_AM=F_AM)

with open('opt_results.json', 'w', encoding='utf-8') as f:
    json.dump(R, f, indent=2, default=float)
print('\nopt_results.json and opt_grids.npz written.')
