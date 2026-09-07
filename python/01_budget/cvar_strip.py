"""Strip re-formulation and CVaR extension of the static hedge-ratio program.

Replaces the single-tranche, single-stress-point program with:
  (1) a 12-slice maturity-matched strip (slice m covers month m's procurement
      with options expiring at T_m = m/12), so the monthly spending authority
      aggregates to an annual strip authority 12*B with no tranche overlap;
  (2) a CVaR_alpha objective on the simulated P-measure loss distribution of
      the whole strip, in place of the (113, 1550) point ledger, solved by
      Rockafellar-Uryasev SLSQP and cross-checked on a deterministic grid;
  (3) an FX-leg maturity design study: co-termed slices vs a 0.5y cap, priced
      under a maturity-widening bid-ask spread curve s(T) = s_1M * (12T)^gamma
      (vol points), which quantifies whether the production configuration's
      T_FX = 0.5 < T_WTI = 0.833 asymmetry is justified.

Inputs: data/results/opt_inputs.json (pinned calibration).
Outputs: data/results/cvar_strip_results.json and the paper figures
  fig_strip_premium.pdf, fig_cvar_contour.pdf, fig_budget_sweep_strip.pdf,
  fig_fx_maturity.pdf under papers/01_budget_allocation/figures/.
"""
import json, os
import numpy as np
from scipy.stats import norm
from scipy.optimize import minimize

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, '..', '..'))
FIGDIR = os.path.join(ROOT, 'papers', '01_budget_allocation', 'figures')
OUT = os.path.join(ROOT, 'data', 'results', 'cvar_strip_results.json')

I = json.load(open(os.path.join(ROOT, 'data', 'results', 'opt_inputs.json')))['inputs']

S1, S2 = I['WTI_spot'], I['KRW_spot']
SIG1H, SIG2, RHO = I['sigma_WTI_hist'], I['sigma_FX'], I['rho']
SIG1D = I['sigma_WTI_diff']
RUS, RKR, RW = I['r_US'], I['r_KRW'], I['WACC']
Q, QUSD = I['Monthly_Oil_Need'], I['Monthly_USD_Need']
B_TR = I['Max_Budget']                      # per-tranche authority (KRW 45bn)
B_YR = 12*B_TR                              # annual strip authority
K1, K2 = 0.95*S1, 0.95*S2
ST1, ST2 = I['Stress_WTI'], I['Stress_KRW']
LAM, THJ, DLJ = 6.7846, -0.0299, 0.08443    # symmetrized Merton triple (paper)
MU1 = 0.139                                  # physical WTI drift (paper, Sec. 5)
MU2 = 0.0                                    # baseline physical FX drift
ALPHA = 0.95
NPATH = 200_000
SEED = 20260824

M = np.arange(1, 13)/12.0                    # slice maturities

# ----------------------------------------------------------------------
# closed-form premiums per slice (paper pricing conventions per maturity)
# ----------------------------------------------------------------------
def b76(F, K, sig, T, r):
    sq = sig*np.sqrt(T)
    d1 = (np.log(F/K) + 0.5*sig*sig*T)/sq
    return np.exp(-r*T)*(F*norm.cdf(d1) - K*norm.cdf(d1-sq))

def gk(S, K, sig, T, rd, rf):
    sq = sig*np.sqrt(T)
    d1 = (np.log(S/K) + (rd-rf+0.5*sig*sig)*T)/sq
    return S*np.exp(-rf*T)*norm.cdf(d1) - K*np.exp(-rd*T)*norm.cdf(d1-sq)

# verify against the pinned single-tranche premiums
chk_b76 = b76(S1, K1, SIG1H, 0.833, RUS)
chk_gk = gk(S2, K2, SIG2, 0.5, RKR, RUS)
assert abs(chk_b76 - 12.6524) < 0.02, chk_b76
assert abs(chk_gk - 84.667) < 0.15, chk_gk

P1m = np.array([b76(S1, K1, SIG1H, t, RUS) for t in M])       # USD/bbl
P2m = np.array([gk(S2, K2, SIG2, t, RKR, RUS) for t in M])    # KRW/USD
carry = 1 + RW*M
K1S = float((Q*P1m*S2*carry).sum())          # full WTI strip premium, KRW
K2S = float((QUSD*P2m*carry).sum())          # full FX strip premium, KRW

def premium(w):
    return K1S*w[0] + K2S*w[1]

# annualized single-point stress ledger (strip form of the old cost function)
UL1 = 12*Q*max(0.0, ST1-S1)*ST2
UL2 = 12*QUSD*max(0.0, ST2-S2)

def stress_cost(w):
    return premium(w) + (1-w[0])*UL1 + (1-w[1])*UL2

# ----------------------------------------------------------------------
# P-measure path bank on the monthly grid
# ----------------------------------------------------------------------
def path_bank(rho=RHO, mu2=MU2, seed=SEED, n=NPATH):
    rng = np.random.default_rng(seed)
    dt = 1.0/12
    z1 = rng.standard_normal((n, 12))
    zx = rng.standard_normal((n, 12))
    z2 = rho*z1 + np.sqrt(1-rho*rho)*zx
    comp = LAM*(np.exp(THJ + 0.5*DLJ*DLJ) - 1)
    nj = rng.poisson(LAM*dt, (n, 12))
    jump = THJ*nj + DLJ*np.sqrt(nj)*rng.standard_normal((n, 12))
    d1 = (MU1 - 0.5*SIG1D*SIG1D - comp)*dt + SIG1D*np.sqrt(dt)*z1 + jump
    d2 = (mu2 - 0.5*SIG2*SIG2)*dt + SIG2*np.sqrt(dt)*z2
    s1p = S1*np.exp(np.cumsum(d1, axis=1))
    s2p = S2*np.exp(np.cumsum(d2, axis=1))
    return s1p, s2p

def loss_parts(s1p, s2p):
    """Per-path affine loss decomposition L(w) = A - w1*Bw - w2*Cc."""
    A = Q*(s1p*s2p - S1*S2).sum(axis=1)
    Bw = Q*(np.maximum(s1p-K1, 0)*s2p).sum(axis=1)
    Cc = QUSD*np.maximum(s2p-K2, 0).sum(axis=1)
    # FX leg capped at T=0.5: slices 7-12 hold the 0.5y option instead
    Ccap = QUSD*(np.maximum(s2p[:, :6]-K2, 0).sum(axis=1)
                 + 6*np.maximum(s2p[:, 5]-K2, 0))
    return A, Bw, Cc, Ccap

def cvar(L, alpha=ALPHA):
    k = int(np.ceil((1-alpha)*len(L)))
    return float(np.partition(L, -k)[-k:].mean())

def cvar_of(w, A, Bw, Cc, alpha=ALPHA):
    return cvar(A - w[0]*Bw - w[1]*Cc, alpha)

# ----------------------------------------------------------------------
# Rockafellar-Uryasev SLSQP + deterministic grid cross-check
# ----------------------------------------------------------------------
def solve_cvar(A, Bw, Cc, budget=B_YR, alpha=ALPHA, kappa=1.0):
    inv = 1.0/((1-alpha)*len(A))
    scale = 1e9
    def obj(x):
        L = A - x[0]*Bw - x[1]*Cc
        return (x[2] + inv*np.maximum(L - x[2], 0).sum())/scale
    cons = [{'type': 'ineq', 'fun': lambda x: (budget - premium(x))/scale},
            {'type': 'ineq', 'fun': lambda x: kappa - x[0] - x[1]}]
    best = None
    for w0 in [(0.9, 0.05), (0.5, 0.5), (0.97, 0.03), (0.7, 0.3), (0.2, 0.1)]:
        t0 = np.quantile(A - w0[0]*Bw - w0[1]*Cc, alpha)
        r = minimize(obj, [w0[0], w0[1], t0], method='SLSQP',
                     bounds=[(0, 1), (0, 1), (None, None)], constraints=cons,
                     options={'maxiter': 400, 'ftol': 1e-12})
        if r.success and (best is None or r.fun < best.fun):
            best = r
    w = best.x[:2]
    return w, cvar_of(w, A, Bw, Cc, alpha)

def grid_check(A, Bw, Cc, budget=B_YR, alpha=ALPHA, n=41, kappa=1.0):
    g = np.linspace(0, 1, n)
    best = (None, np.inf)
    for a in g:
        for b in g:
            if a + b > kappa + 1e-12 or premium([a, b]) > budget:
                continue
            v = cvar_of([a, b], A, Bw, Cc, alpha)
            if v < best[1]:
                best = ((a, b), v)
    return best

# ----------------------------------------------------------------------
# run
# ----------------------------------------------------------------------
R = {'strip': {
        'maturities': M.tolist(), 'P_B76_slice': P1m.tolist(),
        'P_GK_slice': P2m.tolist(), 'K1S': K1S, 'K2S': K2S,
        'full_both_premium': K1S + K2S, 'B_year': B_YR,
        'UL1_year': UL1, 'UL2_year': UL2,
        'single_tranche_b76_check': float(chk_b76),
        'single_tranche_gk_check': float(chk_gk)}}

s1p, s2p = path_bank()
A, Bw, Cc, Ccap = loss_parts(s1p, s2p)

sig_res = lambda w: np.sqrt((1-w[0])**2*SIG1H**2 + (1-w[1])**2*SIG2**2
                            + 2*(1-w[0])*(1-w[1])*SIG1H*SIG2*RHO)

# closed-form variance line solution (unchanged by the strip)
w1_line = (SIG1H**2 - RHO*SIG1H*SIG2)/(SIG1H**2 + SIG2**2 - 2*RHO*SIG1H*SIG2)
w_var = (w1_line, 1-w1_line)

# stress-ledger program (strip form): risk-min s.t. stress_cost <= B_YR
def solve_stress():
    cons = [{'type': 'ineq', 'fun': lambda w: (B_YR - stress_cost(w))/1e9},
            {'type': 'ineq', 'fun': lambda w: 1 - w[0] - w[1]}]
    best = None
    for w0 in [(0.9, 0.05), (0.97, 0.03), (0.5, 0.5)]:
        r = minimize(sig_res, w0, method='SLSQP', bounds=[(0, 1), (0, 1)],
                     constraints=cons, options={'maxiter': 400, 'ftol': 1e-14})
        if r.success and (best is None or r.fun < best.fun):
            best = r
    return best.x

w_stress = solve_stress()
w_c95, c95 = solve_cvar(A, Bw, Cc, alpha=0.95)
w_c99, c99 = solve_cvar(A, Bw, Cc, alpha=0.99)
w_c95k2, c95k2 = solve_cvar(A, Bw, Cc, alpha=0.95, kappa=2.0)
w_c99k2, _ = solve_cvar(A, Bw, Cc, alpha=0.99, kappa=2.0)
(gw, gv) = grid_check(A, Bw, Cc, alpha=0.95)
(gwk2, gvk2) = grid_check(A, Bw, Cc, alpha=0.95, kappa=2.0)
w_corner = (1.0, 0.0)

def summarize(w):
    w = list(map(float, w))
    return {'w1': w[0], 'w2': w[1], 'premium': premium(w),
            'stress_cost': stress_cost(w), 'sigma_res': float(sig_res(w)),
            'cvar95': cvar_of(w, A, Bw, Cc, 0.95),
            'cvar99': cvar_of(w, A, Bw, Cc, 0.99),
            'mean_loss': float((A - w[0]*Bw - w[1]*Cc).mean())}

R['solutions'] = {
    'variance_line': summarize(w_var),
    'stress_ledger': summarize(w_stress),
    'cvar95': summarize(w_c95),
    'cvar99': summarize(w_c99),
    'cost_corner_10': summarize(w_corner),
    'unhedged': summarize((0.0, 0.0)),
    'cvar95_kappa2': summarize(w_c95k2),
    'cvar99_kappa2': summarize(w_c99k2),
    'grid_check_cvar95': {'w1': gw[0], 'w2': gw[1], 'cvar95': gv,
                          'slsqp_gap': gv - c95},
    'grid_check_cvar95_kappa2': {'w1': gwk2[0], 'w2': gwk2[1], 'cvar95': gvk2,
                                 'slsqp_gap': gvk2 - c95k2},
}

# budget bindingness of the strip
R['budget'] = {
    'full_wti_strip': K1S, 'full_fx_strip': K2S,
    'full_both': K1S + K2S,
    'binds_at_cvar95': premium(w_c95) > B_YR - 1e6,
    'premium_at_cvar95': premium(w_c95),
    'slack_at_cvar95': B_YR - premium(w_c95),
}

# correlation sensitivity of the CVaR split (joint-tail channel), both caps
rho_sweep = {}
for r_ in (0.0876, 0.20, 0.30, 0.50):
    s1r, s2r = path_bank(rho=r_)
    Ar, Br_, Cr, _ = loss_parts(s1r, s2r)
    wr, vr = solve_cvar(Ar, Br_, Cr)
    wr2, vr2 = solve_cvar(Ar, Br_, Cr, kappa=2.0)
    rho_sweep[f'{r_:.4f}'] = {'w1': float(wr[0]), 'w2': float(wr[1]),
                              'cvar95': vr,
                              'w1_k2': float(wr2[0]), 'w2_k2': float(wr2[1]),
                              'cvar95_k2': vr2,
                              'premium_k2': premium(wr2)}
R['rho_sweep'] = rho_sweep

# drift robustness (UIP FX drift)
s1u, s2u = path_bank(mu2=RKR-RUS)
Au, Bu, Cu, _ = loss_parts(s1u, s2u)
wu, vu = solve_cvar(Au, Bu, Cu)
wu2, vu2 = solve_cvar(Au, Bu, Cu, kappa=2.0)
R['drift_uip'] = {'w1': float(wu[0]), 'w2': float(wu[1]), 'cvar95': vu,
                  'w1_k2': float(wu2[0]), 'w2_k2': float(wu2[1]),
                  'cvar95_k2': vu2}

# ----------------------------------------------------------------------
# FX maturity design: co-termed vs 0.5y-capped under the spread curve
# ----------------------------------------------------------------------
def fx_premium_spread(cap=False, s1m=0.4, gamma=0.5):
    """Full-coverage FX strip premium paid at ask, KRW (w2=1)."""
    tot = 0.0
    for m, t in enumerate(M, start=1):
        te = min(t, 0.5) if cap else t
        half = 0.5*s1m*(12*te)**gamma/100.0           # half-spread in vol
        tot += QUSD*gk(S2, K2, SIG2 + half, te, RKR, RUS)*(1 + RW*te)
    return tot

fx = {'mid_co': float((QUSD*P2m*carry).sum()),
      'mid_cap': float(sum(QUSD*gk(S2, K2, SIG2, min(t, 0.5), RKR, RUS)
                           * (1+RW*min(t, 0.5)) for t in M))}
# the FX leg is priced at the kappa=2 tail optimum, where it is actually held
w1c, w2c = float(w_c95k2[0]), float(w_c95k2[1])
dcvar = (cvar(A - w1c*Bw - w2c*Ccap) - cvar(A - w1c*Bw - w2c*Cc))
fx['dcvar95_w2eq1'] = cvar(A - w1c*Bw - Ccap) - cvar(A - w1c*Bw - Cc)
fx['dcvar95_at_k2_optimum'] = float(dcvar)
grid_g = [0.25, 0.5, 0.75, 1.0]
grid_s = [0.2, 0.4, 0.8]
rows = []
for s1m in grid_s:
    for g_ in grid_g:
        ask_co = fx_premium_spread(False, s1m, g_)
        ask_cap = fx_premium_spread(True, s1m, g_)
        rows.append({'s1M_volpts': s1m, 'gamma': g_,
                     'ask_co': ask_co, 'ask_cap': ask_cap,
                     'saving_full_cover': ask_co - ask_cap,
                     'saving_at_w2': w2c*(ask_co - ask_cap)})
fx['spread_grid'] = rows
fx['w1_w2_at_k2'] = [w1c, w2c]
R['fx_maturity'] = fx

# ----------------------------------------------------------------------
# figures (grayscale, matching the paper's style)
# ----------------------------------------------------------------------
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams.update({
    'font.family': 'serif', 'font.size': 10, 'axes.titlesize': 11,
    'axes.labelsize': 10, 'legend.fontsize': 8.5, 'figure.dpi': 150,
    'axes.prop_cycle': plt.cycler(color=['black']),
})
os.makedirs(FIGDIR, exist_ok=True)

# (1) per-slice premiums
fig, ax = plt.subplots(1, 2, figsize=(9, 3.2))
ax[0].plot(M*12, Q*P1m*S2*carry/1e9, 'o-', ms=3)
ax[0].set_xlabel('slice maturity (months)'); ax[0].set_ylabel('KRW bn')
ax[0].set_title('WTI leg premium per slice (full cover)')
for s1m, ls in zip(grid_s, ('--', '-', ':')):
    a = [QUSD*gk(S2, K2, SIG2+0.5*s1m*(12*t)**0.5/100, t, RKR, RUS)*(1+RW*t)/1e9
         for t in M]
    ax[1].plot(M*12, a, ls, lw=1, label=f'ask, $s_{{1M}}$={s1m} vp, $\\gamma$=0.5')
ax[1].plot(M*12, QUSD*P2m*carry/1e9, 'o-', ms=3, label='mid')
ax[1].set_xlabel('slice maturity (months)'); ax[1].set_ylabel('KRW bn')
ax[1].set_title('FX leg premium per slice (full cover)')
ax[1].legend(frameon=False)
fig.tight_layout(); fig.savefig(os.path.join(FIGDIR, 'fig_strip_premium.pdf'),
                                bbox_inches='tight'); plt.close(fig)

# (2) CVaR95 contours with premium boundary and solutions (kappa=2 window)
n = 81
g1 = np.linspace(0.6, 1.0, n); g2 = np.linspace(0.0, 1.0, n)
sub = slice(0, 50_000)
As, Bs, Cs = A[sub], Bw[sub], Cc[sub]
Z = np.empty((n, n))
for j, b in enumerate(g2):
    Lb = As - b*Cs
    for i, a in enumerate(g1):
        Z[j, i] = cvar(Lb - a*Bs, 0.95)
fig, ax = plt.subplots(figsize=(6.4, 5))
cs = ax.contour(g1, g2, Z/1e9, levels=14, colors='gray', linewidths=0.7)
ax.clabel(cs, fmt='%.0f', fontsize=7)
gg1, gg2 = np.meshgrid(g1, g2)
feas = (K1S*gg1 + K2S*gg2 <= B_YR)
ax.contourf(g1, g2, feas, levels=[0.5, 1.5], colors=['0.92'])
ax.plot(g1, 1-g1, 'k--', lw=1, label='$w_1+w_2=1$')
bb = (B_YR - K1S*g1)/K2S
ax.plot(g1, np.clip(bb, 0, 1), 'k-', lw=1.2, label='strip premium $=12B$')
for w_, mk, lb in [(w_c95, 'o', 'CVaR$_{95}$ min ($\\kappa$=1)'),
                   (w_c95k2, 'D', 'CVaR$_{95}$ min ($\\kappa$=2)'),
                   (w_var, 's', 'variance line'),
                   (w_corner, 'v', 'cost corner')]:
    ax.plot(w_[0], w_[1], mk, color='black', mfc='white', ms=7, label=lb)
ax.set_xlim(0.6, 1.0); ax.set_ylim(0, 1.0)
ax.set_xlabel('$w_1$ (WTI strip coverage)'); ax.set_ylabel('$w_2$ (FX strip coverage)')
ax.legend(frameon=False, loc='upper left')
ax.set_title('CVaR$_{95}$ of the strip loss (KRW bn), affordable set shaded')
fig.tight_layout(); fig.savefig(os.path.join(FIGDIR, 'fig_cvar_contour.pdf'),
                                bbox_inches='tight'); plt.close(fig)

# (3) authority sweep
Bs_ = np.linspace(300e9, 700e9, 33)
sweep = []
for b_ in Bs_:
    try:
        w1s, v1s = solve_cvar(A, Bw, Cc, budget=b_, kappa=1.0)
        w2s, v2s = solve_cvar(A, Bw, Cc, budget=b_, kappa=2.0)
        sweep.append((b_, v1s, v2s, float(w2s[1])))
    except Exception:
        sweep.append((b_, np.nan, np.nan, np.nan))
sw = np.array(sweep)
R['authority_sweep'] = [{'B': float(a), 'cvar95_k1': float(b),
                         'cvar95_k2': float(c), 'w2_k2': float(d)}
                        for a, b, c, d in sweep]
fig, ax = plt.subplots(figsize=(6.4, 3.6))
ax.plot(sw[:, 0]/1e9, sw[:, 1]/1e9, 'k--', lw=1.1, label='$\\kappa=1$')
ax.plot(sw[:, 0]/1e9, sw[:, 2]/1e9, 'k-', lw=1.3, label='$\\kappa=2$')
ax.axvline(B_YR/1e9, color='black', ls=':', lw=1)
ax.legend(frameon=False)
ax.set_xlabel('annual strip authority (KRW bn)')
ax.set_ylabel('optimal CVaR$_{95}$ (KRW bn)')
fig.tight_layout(); fig.savefig(os.path.join(FIGDIR, 'fig_budget_sweep_strip.pdf'),
                                bbox_inches='tight'); plt.close(fig)

# (4) FX maturity design: spread saving vs tail cost
fig, ax = plt.subplots(figsize=(6.4, 3.6))
gam = np.linspace(0.1, 1.2, 45)
for s1m, ls in zip(grid_s, ('--', '-', ':')):
    sv = [w2c*(fx_premium_spread(False, s1m, g_) -
               fx_premium_spread(True, s1m, g_))/1e9 for g_ in gam]
    ax.plot(gam, sv, ls, lw=1.1, label=f'spread+mid saving, $s_{{1M}}$={s1m} vp')
ax.axhline(dcvar/1e9, color='black', lw=1.4)
ax.annotate('CVaR$_{95}$ cost of the 0.5y cap', (0.15, dcvar/1e9), xytext=(0.15, dcvar/1e9*1.15), fontsize=8.5)
ax.set_xlabel('spread steepness $\\gamma$')
ax.set_ylabel('KRW bn (at optimal $w_2$)')
ax.legend(frameon=False)
fig.tight_layout(); fig.savefig(os.path.join(FIGDIR, 'fig_fx_maturity.pdf'),
                                bbox_inches='tight'); plt.close(fig)

json.dump(R, open(OUT, 'w'), indent=1, default=float)
print(json.dumps({k: R[k] for k in ('strip', 'solutions', 'budget', 'rho_sweep',
                                    'drift_uip')}, indent=1, default=float))
print('fx_maturity:', json.dumps(R['fx_maturity'], indent=1, default=float)[:2000])
print('written', OUT)
