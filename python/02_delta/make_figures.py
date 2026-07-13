import json, os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime

plt.rcParams.update({
    'font.family': 'serif',
    'font.size': 10,
    'axes.titlesize': 11,
    'axes.labelsize': 10,
    'legend.fontsize': 9,
    'figure.dpi': 150,
    'image.cmap': 'gray',
    'axes.prop_cycle': plt.cycler(color=['black']),
    'axes.unicode_minus': False,
})

with open('scraped_data.json', encoding='utf-8') as f:
    d = json.load(f)

pl = np.load('pl_paths.npz')

FIGDIR = 'figures'
if os.path.exists(FIGDIR):
    for fn in os.listdir(FIGDIR):
        if fn.startswith('fig_'):
            os.remove(os.path.join(FIGDIR, fn))
os.makedirs(FIGDIR, exist_ok=True)

def savefig(fig, name):
    path = os.path.join(FIGDIR, name)
    fig.savefig(path, bbox_inches='tight')
    plt.close(fig)
    print('wrote', path)

def moments(x):
    m = x.mean()
    sd = x.std(ddof=1)
    sd0 = x.std(ddof=0)
    skew = ((x - m) ** 3).mean() / sd0 ** 3
    kurt = ((x - m) ** 4).mean() / sd0 ** 4 - 3
    return m, sd, skew, kurt

BN = 1e9  # KRW billions divisor

# grayscale shades
K1 = '0.05'   # near-black
K2 = '0.35'
K3 = '0.60'
K4 = '0.80'

# ---------------------------------------------------------------
# Figure 1: historical WTI / KRW timeseries (real, motivating data)
# ---------------------------------------------------------------
ts = d['raw_timeseries']
dates = [datetime.fromisoformat(x) for x in ts['dates']]
fig, ax1 = plt.subplots(figsize=(6.2, 3.4))
ax1.plot(dates, ts['wti'], color='black', lw=0.9, ls='-', label='WTI (USD/bbl)')
ax1.set_ylabel('WTI spot (USD/bbl)')
ax2 = ax1.twinx()
ax2.plot(dates, ts['krw'], color='0.55', lw=0.9, ls='--', label='USD/KRW')
ax2.set_ylabel('USD/KRW (gray, dashed)')
ax1.xaxis.set_major_locator(mdates.YearLocator())
ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
ax1.set_title(f'Historical WTI and USD/KRW, {dates[0].date()} to {dates[-1].date()} (n={len(dates)})')
fig.tight_layout()
savefig(fig, 'fig_historical_timeseries.pdf')

# ---------------------------------------------------------------
# Figure 2: alive-and-ITM probability decay by maturity
# ---------------------------------------------------------------
fig, ax = plt.subplots(figsize=(6.2, 3.6))
itm = d['itm_diag']
mats = sorted(itm.items(), key=lambda x: float(x[0]))
shades = np.linspace(0.05, 0.75, len(mats))
styles = ['-', '--', '-.', ':', '-', '--']
for (mat, rows), gshade, ls in zip(mats, shades, styles):
    tot = [r[0] for r in rows]
    p = [r[1] for r in rows]
    ax.plot(tot, p, lw=1.1, color=str(gshade), ls=ls, label=f'T={float(mat):.1f}y')
ax.set_xlabel(r'$t/T$ (elapsed fraction of maturity)')
ax.set_ylabel(r'$P(\mathrm{alive} \cap \mathrm{ITM})$')
ax.set_title('Alive-and-ITM path probability decay, by maturity')
ax.legend(ncol=2, frameon=False)
fig.tight_layout()
savefig(fig, 'fig_itm_alive_decay.pdf')

# ---------------------------------------------------------------
# Figure 3: hedge cost distribution, full population (AM vs EU vs AS-TW-proxy)
# ---------------------------------------------------------------
pf = d['parity']['full']
labels = ['American\n(LSMC regression)', 'European\n(Black-76 proxy)', 'Asian\n(TW proxy)']
means = [pf['am_mean']/BN, pf['eu_mean']/BN, pf['as_mean']/BN]
stds  = [pf['am_std']/BN, pf['eu_std']/BN, pf['as_std']/BN]
fig, ax = plt.subplots(figsize=(6.2, 3.6))
x = range(len(labels))
ax.bar(x, means, yerr=stds, capsize=5, color=[K1, K3, K4], edgecolor='black',
       hatch=['', '//', 'xx'])
ax.set_xticks(list(x)); ax.set_xticklabels(labels)
ax.set_ylabel('Mean hedge cost, bn KRW (error bar = 1 StdDev)')
ax.axhline(0, color='black', lw=0.6)
ax.set_title(f"Hedge-cost distribution, full population (n={d['parity']['n']})")
fig.tight_layout()
savefig(fig, 'fig_hedge_cost_distribution.pdf')

# ---------------------------------------------------------------
# Figure 4: clean subset (non-KO, non-exercise) AM vs EU
# ---------------------------------------------------------------
pc = d['parity']['clean']
labels = ['American (LSMC)', 'European (Black-76)']
means = [pc['am_mean']/BN, pc['eu_mean']/BN]
stds  = [pc['am_std']/BN, pc['eu_std']/BN]
fig, ax = plt.subplots(figsize=(5.6, 3.6))
ax.bar(range(2), means, yerr=stds, capsize=5, color=[K1, K3], edgecolor='black', hatch=['', '//'])
ax.set_xticks(range(2)); ax.set_xticklabels(labels)
ax.set_ylabel('Mean hedge cost, bn KRW (error bar = 1 StdDev)')
ax.set_title(f"Clean subset (non-KO, non-exercise), n={pc['n_clean']}")
fig.tight_layout()
savefig(fig, 'fig_hedge_cost_clean_subset.pdf')

# ---------------------------------------------------------------
# Figure 5: FD delta vs regression delta across t/T
# ---------------------------------------------------------------
dc = d['deltacheck']
t_over_T = [r[0] for r in dc]
fd = [r[2] for r in dc]
reg = [r[3] for r in dc]
fig, ax = plt.subplots(figsize=(5.6, 3.6))
ax.plot(t_over_T, fd, 'o-', color='black', label='Finite-difference delta')
ax.plot(t_over_T, reg, 's--', color='0.55', label='LSMC regression delta')
ax.set_xlabel(r'$t/T$ (elapsed fraction of maturity)')
ax.set_ylabel(r'$\delta_{WTI}$')
ax.set_title('Finite-difference vs. regression delta')
ax.legend(frameon=False)
fig.tight_layout()
savefig(fig, 'fig_delta_check.pdf')

# ---------------------------------------------------------------
# Figure 6: FX hedge multiplier sweep (c-factor)
# ---------------------------------------------------------------
sw = d['deltafxsweep']
c = [r[0] for r in sw]
mean_pl = [r[1]/BN for r in sw]
std_pl = [r[2]/BN for r in sw]
fig, ax1 = plt.subplots(figsize=(6.0, 3.6))
ax1.plot(c, mean_pl, 'o-', color='black', label='Mean total P&L')
ax1.set_xlabel(r'FX hedge multiplier $c$ ($\delta_{FX} = c \cdot \delta_{WTI}$)')
ax1.set_ylabel('Mean total P&L, bn KRW (black, solid)')
ax1.axvline(1.0, color='0.6', lw=0.8, ls=':')
ax2 = ax1.twinx()
ax2.plot(c, std_pl, 's--', color='0.55', label='Std total P&L')
ax2.set_ylabel('Std total P&L, bn KRW (gray, dashed)')
ax1.set_title('Cost of fixed equal-delta FX coupling: multiplier sweep')
fig.tight_layout()
savefig(fig, 'fig_delta_fx_sweep.pdf')

# ---------------------------------------------------------------
# Figure 7: jump risk-premium sensitivity (Lambda_Q factor)
# ---------------------------------------------------------------
jp = d['jumppremiumsens']
lf = [r[0] for r in jp]
base = [r[2] for r in jp]
stress = [r[3] for r in jp]
fig, ax = plt.subplots(figsize=(5.6, 3.6))
ax.plot(lf, base, 'o-', color='black', label='Base premium')
ax.plot(lf, stress, 's--', color='0.55', label='Stress premium (WTI=113)')
ax.set_xlabel(r'$\lambda_Q$ scaling factor')
ax.set_ylabel('Premium, KRW per unit')
ax.set_title('Jump-intensity re-pricing sensitivity')
ax.legend(frameon=False)
fig.tight_layout()
savefig(fig, 'fig_jump_premium_sensitivity.pdf')

# ---------------------------------------------------------------
# Figure 8: k-sigma threshold robustness
# ---------------------------------------------------------------
lr = d['lambdarobustness']
k = [r[0] for r in lr]
ko = [r[7] for r in lr]
mean_pl = [r[8]/BN for r in lr]
fig, ax1 = plt.subplots(figsize=(5.6, 3.6))
ax1.plot(k, ko, 'o-', color='black')
ax1.set_xlabel(r'jump-threshold $k$ (multiples of $\sigma\sqrt{dt}$)')
ax1.set_ylabel('KO rate (black, solid)')
ax2 = ax1.twinx()
ax2.plot(k, mean_pl, 's--', color='0.55')
ax2.set_ylabel('Mean total P&L, bn KRW (gray, dashed)')
ax1.set_title(r'$k$-sigma calibration threshold robustness')
fig.tight_layout()
savefig(fig, 'fig_lambda_robustness.pdf')

# ---------------------------------------------------------------
# Figure 9: Girsanov reweighting validation + exact drift-bias residual
# ---------------------------------------------------------------
g = d['girsanov']
db = d['driftbias']
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(8.6, 3.4))
labels = ['Reweighted\n($\\mathbb{Q}$-paths $\\times$ Girsanov LR)', 'Direct\n(P-measure simulation)']
means = [g['reweighted_mean'], g['direct_mean']]
stds  = [g['reweighted_std'], g['direct_std']]
ax1.bar(range(2), means, yerr=stds, capsize=5, color=[K1, K3], edgecolor='black', hatch=['', '//'])
ax1.set_xticks(range(2)); ax1.set_xticklabels(labels)
ax1.set_ylabel('Mean payoff, KRW per unit')
ax1.set_title(f"Girsanov reweighting vs. direct P-sim\n(residual = {g['residual_pct']*100:.2f}% of direct mean)")

Tvals = [r[0] for r in db]
resid = [r[3] for r in db]
ax2.plot(Tvals, resid, 'o-', color='black')
ax2.set_xlabel('Maturity $T$ (years)')
ax2.set_ylabel('Simulated $-$ theoretical drift bias')
ax2.set_title('Closed-form P/Q drift-bias check\n(machine-precision agreement)')
ax2.ticklabel_format(axis='y', style='sci', scilimits=(0,0))
fig.tight_layout()
savefig(fig, 'fig_girsanov_validation.pdf')

# ---------------------------------------------------------------
# Figure 10 (NEW): empirical total-P&L distribution vs fitted normal,
# per engine, from the actual n=10,000 per-path ledger (real data).
# ---------------------------------------------------------------
fig, axes = plt.subplots(3, 1, figsize=(7.6, 8.6), sharey=False)
engines = [('American (LSMC regression)', pl['am_tot']), ('European (Black-76 proxy)', pl['eu_tot']), ('Asian (TW proxy)', pl['as_tot'])]
for ax, (label, series) in zip(axes, engines):
    x = series / BN
    m, sd, skew, kurt = moments(x)
    counts, bins, _ = ax.hist(x, bins=70, density=True, color='0.75', edgecolor='0.3', linewidth=0.3)
    grid = np.linspace(x.min(), x.max(), 400)
    normal_pdf = (1.0/(sd*np.sqrt(2*np.pi))) * np.exp(-0.5*((grid-m)/sd)**2)
    ax.plot(grid, normal_pdf, color='black', lw=1.8, ls='--', label='Fitted normal')
    ax.axvline(m, color='black', lw=1.0, ls=':')
    ax.set_title(label, fontsize=13)
    ax.set_xlabel('Total P&L, bn KRW', fontsize=11)
    ax.set_ylabel('Density', fontsize=11)
    ax.tick_params(labelsize=10)
    ax.text(0.03, 0.90, f"skew={skew:.2f}   kurt={kurt:.2f}", transform=ax.transAxes,
            va='top', ha='left', fontsize=10)
axes[0].legend(frameon=False, fontsize=10, loc='upper right')
fig.suptitle('Empirical total-P&L distribution vs. fitted normal density (n=10,000 paths per engine)', fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.94])
savefig(fig, 'fig_pl_distribution.pdf')

print("DONE: 10 figures written to", FIGDIR)
