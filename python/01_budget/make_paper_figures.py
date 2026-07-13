"""Figures for the static hedge-ratio allocation paper.

All numbers come from opt_scraped.json (workbook inputs) and
corrected_results.json (solved optima) only.  Grayscale throughout.
"""
import json, os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import cm
from mpl_toolkits.mplot3d import proj3d  # noqa

plt.rcParams.update({
    'font.family': 'serif', 'font.size': 10, 'axes.titlesize': 11,
    'axes.labelsize': 10, 'legend.fontsize': 8.5, 'figure.dpi': 150,
    'image.cmap': 'gray', 'axes.prop_cycle': plt.cycler(color=['black']),
})

D = json.load(open('opt_scraped.json')); I = D['inputs']
R = json.load(open('corrected_results.json'))
B = I['Max_Budget']; S2, RHO = I['sigma_FX'], I['rho']
C = R['coeffs']
K1E, K2E, K1A, K2A, M1, M2 = C['K1E'], C['K2E'], C['K1A'], C['K2A'], C['M1'], C['M2']

FIGDIR = 'figures'
os.makedirs(FIGDIR, exist_ok=True)

def savefig(fig, name):
    fig.savefig(os.path.join(FIGDIR, name), bbox_inches='tight')
    plt.close(fig)
    print('wrote', os.path.join(FIGDIR, name))

def gm(W1, W2, s1):
    H1, H2 = 1-W1, 1-W2
    return np.sqrt(H1**2*s1**2 + H2**2*S2**2 + 2*H1*H2*s1*S2*RHO)

def cost_eu(W1, W2): return K1E*W1 + K2E*W2 + M1*(1-W1) + M2*(1-W2)
def cost_am(W1, W2): return K1A*W1 + K2A*W2 + M1*(1-W1) + M2*(1-W2)

ENG = [
    ('European (Black-76 / GK premiums)', I['sigma_WTI_hist'], cost_eu, R['EU']),
    ('American (LSMC Shapley premiums)', I['sigma_WTI_diff'], cost_am, R['AM']),
]

# =====================================================================
# Figure 1: 3-D residual-volatility surface, feasible region shaded
# =====================================================================
def feas_boundary_paths(costf, n=801):
    g = np.linspace(0, 1, n)
    A, Bm = np.meshgrid(g, g)
    F = ((A + Bm <= 1) & (costf(A, Bm) <= B)).astype(float)
    figt, axt = plt.subplots()
    cs = axt.contour(A, Bm, F, levels=[0.5])
    paths = [p for p in cs.allsegs[0]]
    plt.close(figt)
    return paths

def draw_surface_panel(ax, s1, costf, res, win, coarse_n, fine_n, zoom):
    """One 3-D panel over window win = (w1lo, w1hi, w2lo, w2hi)."""
    w1lo, w1hi, w2lo, w2hi = win
    gw1 = np.linspace(w1lo, w1hi, coarse_n)
    gw2 = np.linspace(w2lo, w2hi, coarse_n)
    Ww1, Ww2 = np.meshgrid(gw1, gw2)
    ax.plot_wireframe(Ww1, Ww2, gm(Ww1, Ww2, s1), rstride=1, cstride=1,
                      color='0.72', linewidth=0.4)
    gf1 = np.linspace(w1lo, w1hi, fine_n)
    gf2 = np.linspace(w2lo, w2hi, fine_n)
    W1, W2 = np.meshgrid(gf1, gf2)
    G = gm(W1, W2, s1)
    Cg = costf(W1, W2)
    F = (W1 + W2 <= 1 + 1e-12) & (Cg <= B)
    Ff = F[:-1, :-1] & F[1:, :-1] & F[:-1, 1:] & F[1:, 1:]
    Gf = 0.25*(G[:-1, :-1] + G[1:, :-1] + G[:-1, 1:] + G[1:, 1:])
    gmin, gmax = G[F].min(), G[F].max()
    zspan = G.max() - G.min()
    norm = (Gf - gmin) / (gmax - gmin)
    shade = cm.gray(np.clip(0.10 + 0.70*norm, 0, 1))
    colors = np.zeros(Ff.shape + (4,))
    colors[...] = (0, 0, 0, 0)
    colors[Ff] = shade[Ff]
    ax.plot_surface(W1, W2, G + 0.004*zspan, facecolors=colors, rstride=1,
                    cstride=1, linewidth=0, antialiased=False, shade=False)
    for p in feas_boundary_paths(costf):
        m = ((p[:, 0] >= w1lo) & (p[:, 0] <= w1hi)
             & (p[:, 1] >= w2lo) & (p[:, 1] <= w2hi))
        if m.any():
            ax.plot(p[m, 0], p[m, 1], gm(p[m, 0], p[m, 1], s1) + 0.008*zspan,
                    color='black', lw=1.4)
    zfloor = G.min()
    rm, cmn = res['risk_min'], res['cost_min']
    for pt, mk, fc, sz, lab in [
            (rm, 'o', 'black', 130 if zoom else 90, 'risk minimum $(w_1^*, w_2^*)$'),
            (cmn, '^', 'white', 130 if zoom else 90, 'cost minimum')]:
        z = gm(pt['w1'], pt['w2'], s1)
        ax.plot([pt['w1']]*2, [pt['w2']]*2, [zfloor, z], color='black',
                lw=0.9, ls=':')
        ax.scatter([pt['w1']], [pt['w2']], [z + 0.012*zspan], marker=mk, s=sz,
                   facecolor=fc, edgecolor='black', linewidth=1.3,
                   depthshade=False, label=lab, zorder=10)
    ax.set_xlabel('$w_1$ (WTI hedge ratio)', labelpad=6, fontsize=9)
    ax.set_ylabel('$w_2$ (FX hedge ratio)', labelpad=6, fontsize=9)
    ax.set_zlabel(r'$\sigma_{res}$', labelpad=4, fontsize=9)
    ax.set_xlim(w1lo, w1hi); ax.set_ylim(w2lo, w2hi)
    ax.view_init(elev=24, azim=-55)
    return rm

fig = plt.figure(figsize=(11.5, 10.6))
ZOOMS = [(0.94, 1.0, 0.0, 0.06), (0.82, 1.0, 0.0, 0.05)]
for k, (title, s1, costf, res) in enumerate(ENG):
    # top row: full unit square
    ax = fig.add_subplot(2, 2, k+1, projection='3d')
    rm = draw_surface_panel(ax, s1, costf, res, (0, 1, 0, 1), 26, 201, zoom=False)
    ax.set_title(title, fontsize=10)
    ax.legend(loc='upper right', frameon=False, fontsize=7.5)
    xa, ya, _ = proj3d.proj_transform(rm['w1'], rm['w2'],
                                      gm(rm['w1'], rm['w2'], s1), ax.get_proj())
    ax.annotate('feasible region\n(shaded, outlined)', xy=(xa, ya),
                xycoords='data', xytext=(0.30, 0.16), textcoords='axes fraction',
                fontsize=8, ha='center', color='0.15',
                arrowprops=dict(arrowstyle='->', color='0.15', lw=0.9, shrinkB=6))
    # bottom row: 3-D zoom onto the feasible neighborhood
    axz = fig.add_subplot(2, 2, k+3, projection='3d')
    draw_surface_panel(axz, s1, costf, res, ZOOMS[k], 19, 161, zoom=True)
    axz.set_title('zoom onto the feasible region', fontsize=9)
fig.suptitle('Residual-volatility surface $\\sigma_{res}(w_1,w_2)$: feasible region (all four constraints) shaded,\n'
             'dark = low risk; infeasible region wireframe only. Top: full unit square. Bottom: 3-D zoom.',
             fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.94])
savefig(fig, 'fig_opt_surface3d.pdf')

# =====================================================================
# Figure 2: contour view, full domain + zoom
# =====================================================================
n2 = 601
g2 = np.linspace(0, 1, n2)
V1, V2 = np.meshgrid(g2, g2)
fig, axes = plt.subplots(2, 2, figsize=(10.4, 9.0))
for k, (title, s1, costf, res) in enumerate(ENG):
    G = gm(V1, V2, s1)
    Cg = costf(V1, V2)
    F = (V1 + V2 <= 1) & (Cg <= B)
    for row in (0, 1):
        ax = axes[row][k]
        ax.contourf(V1, V2, F.astype(float), levels=[0.5, 1.5], colors=['0.82'])
        cs = ax.contour(V1, V2, G, levels=10, colors='0.45', linewidths=0.7)
        ax.clabel(cs, inline=True, fontsize=6.5, fmt='%.2f')
        ax.contour(V1, V2, Cg, levels=[B], colors='black', linewidths=1.4)
        ax.plot([0, 1], [1, 0], color='black', ls='--', lw=1.0)
        rm = res['risk_min']
        ax.plot(rm['w1'], rm['w2'], marker='o', ms=7, color='black', ls='none',
                label='risk minimum')
        cmn = res['cost_min']
        ax.plot(cmn['w1'], cmn['w2'], marker='^', ms=8, mfc='white', mec='black',
                ls='none', label='cost minimum')
        cb = res['cost_min_cap']
        ax.plot(cb['w1'], cb['w2'], marker='x', ms=8, color='black', ls='none',
                mew=1.4, label='cost minimum under risk cap')
        if row == 0:
            ax.set_xlim(0, 1); ax.set_ylim(0, 1)
            ax.set_title(title, fontsize=10)
        else:
            if k == 0:
                ax.set_xlim(0.94, 1.002); ax.set_ylim(-0.002, 0.06)
            else:
                ax.set_xlim(0.82, 1.002); ax.set_ylim(-0.002, 0.05)
            ax.set_title('zoom near the optima', fontsize=9)
        ax.set_xlabel('$w_1$ (WTI hedge ratio)')
        ax.set_ylabel('$w_2$ (FX hedge ratio)')
axes[0][0].legend(loc='upper right', frameon=False)
fig.suptitle('Iso-risk contours of $\\sigma_{res}$, budget boundary (solid), $w_1+w_2=1$ (dashed);\n'
             'feasible region shaded gray. Top: full domain. Bottom: zoom.',
             fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.93])
savefig(fig, 'fig_opt_contour.pdf')

# =====================================================================
# Figure 3: budget relaxation sweep
# =====================================================================
fig, axes = plt.subplots(1, 2, figsize=(10.0, 3.8))
for k, (title, s1, costf, res) in enumerate(ENG):
    ax = axes[k]
    sw = res['sweep']
    Bx = [s['B']/1e9 for s in sw]
    Gy = [s['gmvp'] for s in sw]
    ax.plot(Bx, Gy, 'o-', color='black', ms=3, lw=1.0)
    ax.axvline(B/1e9, color='0.45', ls='--', lw=1.0)
    nb = res['no_budget']
    ax.axhline(nb['gmvp'], color='0.45', ls=':', lw=1.0)
    ax.text(0.97, 0.90, 'budget cap = 45 (dashed)\n'
            f"unconstrained floor $\\sigma_{{res}}$ = {nb['gmvp']:.5f} (dotted),\n"
            f"reached at budget = {nb['cost']/1e9:.2f} bn",
            transform=ax.transAxes, ha='right', va='top', fontsize=7.5, color='0.15')
    ax.set_xlabel('Budget $B$ (bn KRW)')
    ax.set_ylabel(r'optimal $\sigma_{res}(B)$')
    ax.set_title(title, fontsize=10)
    ax.ticklabel_format(axis='y', useOffset=False)
fig.suptitle('Budget relaxation: the risk-minimizing residual volatility as a function of the budget cap',
             fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.92])
savefig(fig, 'fig_budget_sweep.pdf')

# =====================================================================
# Figure 4: knock-out survival haircut (American structure)
# =====================================================================
V2 = json.load(open('v2_curve.json'))
PM = V2['pm']                       # stress-conditional pKO (MC measured)
PSTAR = V2['pstar']
PCROSS = V2['pcross']
PFLOOR = V2['pfloor']
SIGV = V2['sigV']

rmA = R['AM']['risk_min']
p = np.linspace(0, 0.9, 721)
min_cost = (K1A + p*M1 + M2) / 1e9                       # cheapest attainable, pure KO
adopted = (K1A*rmA['w1'] + K2A*rmA['w2']
           + (1-rmA['w1']*(1-p))*M1 + (1-rmA['w2']*(1-p))*M2) / 1e9
# mixed-program optimum ledger: floor-point cost while unconstrained, then
# budget-pinned at exactly B
w1f, w2f = 0.945202, 0.054798
mixed = np.where(p < PFLOOR,
                 (K1A*w1f + K2E*w2f + (1-w1f*(1-p))*M1 + (1-w2f)*M2)/1e9,
                 B/1e9)

fig, ax = plt.subplots(figsize=(7.2, 4.2))
ax.plot(p, adopted, color='black', lw=1.3,
        label='pure-KO book, adopted allocation')
ax.plot(p, min_cost, color='0.45', lw=1.3, ls='--',
        label='pure-KO book, minimum attainable')
ax.plot(p, mixed, color='black', lw=2.0, ls='-.',
        label='mixed vanilla/KO program optimum')
ax.axhline(B/1e9, color='black', lw=0.8, ls=':')
ax.text(0.012, B/1e9 + 2.5, 'budget cap = 45 bn', fontsize=8, color='0.15')
ax.axvline(PSTAR, color='0.45', lw=0.9, ls='-.')
ax.text(PSTAR + 0.008, 100, f'$p^*={PSTAR:.3f}$\n(pure KO dies)', fontsize=7.5, color='0.15')
for pk, lab, dy in [(0.2309, 'unconditional KO rate\n(net of early exercise)', 16),
                    (0.4369, 'unconditional\nbarrier-touch rate', 16)]:
    ax.plot([pk], [np.interp(pk, p, adopted)], marker='o', ms=5, color='black')
    ax.annotate(lab, xy=(pk, np.interp(pk, p, adopted)),
                xytext=(pk + 0.02, np.interp(pk, p, adopted) - dy),
                fontsize=7.5, color='0.15',
                arrowprops=dict(arrowstyle='->', color='0.15', lw=0.8))
yA = np.interp(PM, p, adopted)
ax.plot([PM], [yA], marker='*', ms=14, color='black')
ax.annotate(f'stress-conditional MC estimate\n$p_{{KO}}={PM:.4f}$: ledger {yA:.1f} bn',
            xy=(PM, yA), xytext=(PM - 0.34, yA - 32), fontsize=8, color='black',
            arrowprops=dict(arrowstyle='->', color='black', lw=0.9))
ax.set_xlabel('$p_{KO}$ (probability the KO structure is dead under stress)')
ax.set_ylabel('stress-adjusted total cost (bn KRW)')
ax.set_title('Knock-out survival haircut: pure-KO book vs. the mixed program')
ax.legend(frameon=False, loc='upper left', fontsize=8)
fig.tight_layout()
savefig(fig, 'fig_pko_haircut.pdf')

# =====================================================================
# Figure 5: mixed-program switch structure (sigma* vs pKO)
# =====================================================================
fig, ax = plt.subplots(figsize=(7.0, 4.0))
cv = [(a, b) for a, b in V2['curve'] if b is not None]
pk = [a for a, b in cv]
sk = [b for a, b in cv]
ax.plot(pk, sk, color='0.45', lw=1.4, ls='--', label='all-KO branch (budget-feasible part)')
ax.axhline(SIGV, color='0.45', lw=1.4, ls=':',
           label='all-vanilla branch $\\sigma=%.7f$' % SIGV)
# mixed = lower envelope
pe = np.linspace(0, 0.15, 601)
env = []
for q in pe:
    if q <= pk[-1]:
        env.append(min(np.interp(q, pk, sk), SIGV))
    else:
        env.append(SIGV)
ax.plot(pe, env, color='black', lw=2.2, label='mixed program optimum (envelope)')
for x, lab, tx, ty in [
        (PFLOOR, 'floor regime ends\n$p=%.4f$' % PFLOOR, PFLOOR - 0.024, 0.09165),
        (PCROSS, 'KO$\\to$vanilla switch\n$p=\\bar p=%.4f$' % PCROSS, PCROSS + 0.0015, 0.09145),
        (PSTAR, 'pure KO infeasible\n$p^*=%.4f$' % PSTAR, PSTAR + 0.0015, 0.09165)]:
    ax.axvline(x, color='0.6', lw=0.8, ls='-.')
    ax.text(tx, ty, lab, fontsize=7, color='0.15', va='top')
ax.annotate('measured stress-conditional $p_{KO}=0.8405$\n'
            '$\\Rightarrow$ deep in the all-vanilla regime $\\rightarrow$',
            xy=(0.149, SIGV), xytext=(0.085, 0.09085), fontsize=8, color='black')
ax.set_xlim(0, 0.15)
ax.set_ylim(0.0905, 0.0920)
ax.set_xlabel('$p_{KO}$ (probability the KO structure is dead under stress)')
ax.set_ylabel(r'optimal $\sigma_{res}$ of the mixed program')
ax.set_title('Instrument choice inside the mixed program: the three regimes')
ax.legend(frameon=False, loc='upper left', fontsize=8)
fig.tight_layout()
savefig(fig, 'fig_mixed_switch.pdf')

print('DONE')
