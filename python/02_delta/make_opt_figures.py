"""Figures for the static hedge-ratio optimization paper.

Everything is generated from opt_scraped.json / opt_results.json only
(Data-is-Truth), grayscale throughout, serif fonts, PDF output.
"""
import json, os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import cm
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

plt.rcParams.update({
    'font.family': 'serif',
    'font.size': 10,
    'axes.titlesize': 11,
    'axes.labelsize': 10,
    'legend.fontsize': 8.5,
    'figure.dpi': 150,
    'image.cmap': 'gray',
    'axes.prop_cycle': plt.cycler(color=['black']),
})

with open('opt_scraped.json', encoding='utf-8') as f:
    D = json.load(f)
with open('opt_results.json', encoding='utf-8') as f:
    R = json.load(f)
I = D['inputs']
BUDGET = I['Max_Budget']
S2, RHO = I['sigma_FX'], I['rho']

FIGDIR = 'figures'
os.makedirs(FIGDIR, exist_ok=True)

def savefig(fig, name):
    path = os.path.join(FIGDIR, name)
    fig.savefig(path, bbox_inches='tight')
    plt.close(fig)
    print('wrote', path)

# ---------------------------------------------------------------------
# model functions (identical transcriptions as in optimize_hedge.py)
# ---------------------------------------------------------------------
K1_EU = I['Monthly_Oil_Need']*I['Black76_call']*I['KRW_spot']*(1+I['WACC']*I['Maturity_Oil'])
K2_EU = I['Monthly_USD_Need']*I['GK_call']*(1+I['WACC']*I['Maturity_FX'])
M1_EU = I['Monthly_Oil_Need']*max(0.0, I['Stress_WTI']-I['WTI_spot'])*I['Stress_KRW']
M2_EU = I['Monthly_USD_Need']*max(0.0, I['Stress_KRW']-I['KRW_spot'])
K1_AM = I['LSMC_WTI_shapley']*I['Monthly_Oil_Need']*np.exp(I['WACC']*I['Maturity_Oil'])
K2_AM = I['LSMC_FX_shapley']*I['Monthly_Oil_Need']*np.exp(I['Maturity_FX']*I['WACC'])
FXGAP = max(0.0, I['Stress_KRW']-I['KRW_spot'])

def cost_eu(W1, W2):
    return K1_EU*W1 + K2_EU*W2 + M1_EU*(1-W1) + M2_EU*(1-W2)

def cost_am(W1, W2):
    return (K1_AM*W1 + K2_AM*W2
            + 1.0 - W1*I['Monthly_Oil_Need']*FXGAP
            + (1-W2)*I['Monthly_USD_Need']*FXGAP)

def gm(W1, W2, s1):
    H1, H2 = 1-W1, 1-W2
    return np.sqrt(H1**2*s1**2 + H2**2*S2**2 + 2*H1*H2*s1*S2*RHO)

ENG = [
    ('European (Black-76 / GK premiums)', I['sigma_WTI_hist'], cost_eu,
     D['european'], R['european']),
    ('American (LSMC Shapley premiums)', I['sigma_WTI_diff'], cost_am,
     D['american'], R['american']),
]

# =====================================================================
# Figure 1: 3-D gmvp surface over the full unit square.
# Two layers: coarse wireframe everywhere (constraint-violating region
# shows as bare skeleton), plus a fine surface whose faces are shaded
# (dark = low risk) only where all four constraints hold.
# =====================================================================
def feas_boundary_paths(costf, n=801):
    """Boundary polylines of the feasible set, via a throwaway 2-D contour."""
    g = np.linspace(0, 1, n)
    A, B = np.meshgrid(g, g)
    F = ((A + B <= 1) & (costf(A, B) <= BUDGET)).astype(float)
    figt, axt = plt.subplots()
    cs = axt.contour(A, B, F, levels=[0.5])
    paths = [p for p in cs.allsegs[0]]
    plt.close(figt)
    return paths

fig = plt.figure(figsize=(11.5, 5.6))
for k, (title, s1, costf, xl, res) in enumerate(ENG):
    ax = fig.add_subplot(1, 2, k+1, projection='3d')

    # layer 1: coarse wireframe over the whole square
    nw = 26
    gw = np.linspace(0, 1, nw)
    Ww1, Ww2 = np.meshgrid(gw, gw)
    ax.plot_wireframe(Ww1, Ww2, gm(Ww1, Ww2, s1), rstride=1, cstride=1,
                      color='0.72', linewidth=0.4)

    # layer 2: fine shaded surface, feasible faces only
    nf = 201
    gf = np.linspace(0, 1, nf)
    W1, W2 = np.meshgrid(gf, gf)
    G = gm(W1, W2, s1)
    C = costf(W1, W2)
    F = (W1 + W2 <= 1 + 1e-12) & (C <= BUDGET)
    Ff = F[:-1, :-1] & F[1:, :-1] & F[:-1, 1:] & F[1:, 1:]
    Gf = 0.25*(G[:-1, :-1] + G[1:, :-1] + G[:-1, 1:] + G[1:, 1:])
    gmin, gmax = G[F].min(), G[F].max()
    norm = (Gf - gmin) / (gmax - gmin)
    shade = cm.gray(np.clip(0.10 + 0.70*norm, 0, 1))
    colors = np.zeros(Ff.shape + (4,))
    colors[...] = (0, 0, 0, 0)          # infeasible faces: fully transparent
    colors[Ff] = shade[Ff]
    ax.plot_surface(W1, W2, G + 0.0015, facecolors=colors, rstride=1,
                    cstride=1, linewidth=0, antialiased=False, shade=False)

    # black outline of the feasible set, drawn on the surface
    for p in feas_boundary_paths(costf):
        ax.plot(p[:, 0], p[:, 1], gm(p[:, 0], p[:, 1], s1) + 0.003,
                color='black', lw=1.2)

    # Excel stored risk-min optimum
    zx = gm(xl['w1'], xl['w2'], s1)
    ax.scatter([xl['w1']], [xl['w2']], [zx + 0.004], marker='o', s=45,
               color='black', depthshade=False,
               label='Excel risk-min $(w_1^*, w_2^*)$')
    # Python cost-min (variant A)
    ca = res['cost_min_A']
    za = gm(ca['w1'], ca['w2'], s1)
    ax.scatter([ca['w1']], [ca['w2']], [za + 0.004], marker='^', s=55,
               facecolor='white', edgecolor='black', linewidth=1.0,
               depthshade=False, label='Python cost-min (no risk cap)')

    ax.set_xlabel('$w_1$ (WTI hedge ratio)', labelpad=6)
    ax.set_ylabel('$w_2$ (FX hedge ratio)', labelpad=6)
    ax.set_zlabel(r'$\sigma_{res}(w_1,w_2)$', labelpad=4)
    ax.set_title(title, fontsize=10)
    ax.view_init(elev=22, azim=-58)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.legend(loc='upper right', frameon=False, fontsize=7.5)

    # 2-D annotation arrow pointing at the (small) feasible region
    from mpl_toolkits.mplot3d import proj3d
    xa, ya, _ = proj3d.proj_transform(xl['w1'], xl['w2'],
                                      gm(xl['w1'], xl['w2'], s1), ax.get_proj())
    ax.annotate('feasible region\n(shaded, outlined)', xy=(xa, ya),
                xycoords='data', xytext=(0.30, 0.16), textcoords='axes fraction',
                fontsize=8, ha='center', color='0.15',
                arrowprops=dict(arrowstyle='->', color='0.15', lw=0.9,
                                shrinkB=6))

fig.suptitle('Residual-volatility surface $\\sigma_{res}(w_1,w_2)$ over the full unit square:\n'
             'feasible region (all four constraints) shaded, dark = low risk; infeasible region wireframe only',
             fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.90])
savefig(fig, 'fig_opt_surface3d.pdf')

# =====================================================================
# Figure 2: contour view, full domain (top) + zoom near optima (bottom)
# =====================================================================
n2 = 601
g2 = np.linspace(0, 1, n2)
V1, V2 = np.meshgrid(g2, g2)

fig, axes = plt.subplots(2, 2, figsize=(10.4, 9.0))
for k, (title, s1, costf, xl, res) in enumerate(ENG):
    G = gm(V1, V2, s1)
    C = costf(V1, V2)
    F = (V1 + V2 <= 1) & (C <= BUDGET)

    for row in (0, 1):
        ax = axes[row][k]
        # feasible region: light gray fill
        ax.contourf(V1, V2, F.astype(float), levels=[0.5, 1.5], colors=['0.82'])
        cs = ax.contour(V1, V2, G, levels=10, colors='0.45', linewidths=0.7)
        ax.clabel(cs, inline=True, fontsize=6.5, fmt='%.2f')
        # budget boundary and sum line
        ax.contour(V1, V2, C, levels=[BUDGET], colors='black', linewidths=1.4)
        ax.plot([0, 1], [1, 0], color='black', ls='--', lw=1.0)

        ax.plot(xl['w1'], xl['w2'], marker='o', ms=7, color='black', ls='none',
                label='Excel risk-min')
        rm = res['risk_min']
        ax.plot(rm['w1'], rm['w2'], marker='x', ms=8, color='black', ls='none',
                mew=1.6, label='Python risk-min')
        ca = res['cost_min_A']
        ax.plot(ca['w1'], ca['w2'], marker='^', ms=8, mfc='white', mec='black',
                ls='none', label='Python cost-min (A)')
        cb = res['cost_min_B']
        ax.plot(cb['w1'], cb['w2'], marker='s', ms=6, mfc='white', mec='black',
                ls='none', label='Python cost-min (B, same-risk)')

        if row == 0:
            ax.set_xlim(0, 1); ax.set_ylim(0, 1)
            ax.set_title(title, fontsize=10)
        else:
            if k == 0:
                ax.set_xlim(0.94, 1.001); ax.set_ylim(-0.002, 0.06)
            else:
                ax.set_xlim(0.90, 1.001); ax.set_ylim(-0.002, 0.10)
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
# Figure 3: budget relaxation sweep -- gmvp* as a function of the budget
# =====================================================================
fig, axes = plt.subplots(1, 2, figsize=(10.0, 3.8))
for k, (title, s1, costf, xl, res) in enumerate(ENG):
    ax = axes[k]
    sw = [s for s in res['budget_sweep'] if s.get('feasible')]
    B = [s['budget']/1e9 for s in sw]
    G = [s['gmvp'] for s in sw]
    ax.plot(B, G, 'o-', color='black', ms=3, lw=1.0)
    ax.axvline(BUDGET/1e9, color='0.45', ls='--', lw=1.0)
    nb = res['risk_min_no_budget']
    ax.axhline(nb['gmvp'], color='0.45', ls=':', lw=1.0)
    ax.text(0.97, 0.90, 'workbook budget = 45 (dashed)\n'
            f"unconstrained floor $\\sigma_{{res}}$ = {nb['gmvp']:.5f} (dotted),\n"
            f"reached at budget = {nb['budget_needed']/1e9:.2f} bn",
            transform=ax.transAxes, ha='right', va='top',
            fontsize=7.5, color='0.15')
    ax.set_xlabel('Budget $B$ (bn KRW)')
    ax.set_ylabel(r'optimal $\sigma_{res}(B)$')
    ax.set_title(title, fontsize=10)
    ax.ticklabel_format(axis='y', useOffset=False)
fig.suptitle('Budget relaxation: the risk-minimizing residual volatility as a function of the budget cap',
             fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.92])
savefig(fig, 'fig_budget_sweep.pdf')

print('DONE')
