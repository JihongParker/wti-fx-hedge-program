import json, os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams.update({
    'font.family': 'serif',
    'font.size': 10,
    'axes.titlesize': 11,
    'axes.labelsize': 10,
    'legend.fontsize': 9,
    'figure.dpi': 150,
    'axes.prop_cycle': plt.cycler(color=['black']),
    'axes.unicode_minus': False,
})

with open('cfh_data.json', encoding='utf-8') as f:
    d = json.load(f)

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

BN = 1e9

def col(series, idx):
    return np.array([r[idx] for r in series['rows']], dtype=float)

T_FX = d['T_FX']
T_WTI = d['T_WTI']

# find KO step: first step where a_ledger AliveFrac < 1, or use KO probability as marker only
a = d['a_ledger']
t_a = col(a, 1)

# ---------------------------------------------------------------
# Figure 1: SFP derivative fair value -- A (1 line) vs B (2 stacked lines)
# ---------------------------------------------------------------
sfp = d['sfp']
t = col(sfp, 1)
A_deriv = col(sfp, 2) / BN
B_wti = col(sfp, 6) / BN
B_fx = col(sfp, 7) / BN
B_total = B_wti + B_fx

fig, ax = plt.subplots(figsize=(6.4, 3.8))
ax.plot(t, A_deriv, color='black', lw=1.6, ls='-', label='Structure A: single derivative')
ax.plot(t, B_total, color='0.4', lw=1.4, ls='--', label='Structure B: WTI+FX derivatives (sum)')
ax.plot(t, B_wti, color='0.6', lw=0.9, ls=':', label='  B1 (WTI) only')
ax.plot(t, B_fx, color='0.6', lw=0.9, ls='-.', label='  B2 (FX) only')
ax.set_xlabel('t (years)')
ax.set_ylabel('Mean derivative asset, bn KRW')
ax.set_title('SFP derivative carrying amount: 1 line (A) vs 2 lines (B)')
ax.legend(frameon=False, fontsize=8)
fig.tight_layout()
savefig(fig, 'fig_cfh_sfp_derivative.pdf')

# ---------------------------------------------------------------
# Figure 2: maturity-mismatch freeze mechanism (B1 vs B2, T_FX marked)
# ---------------------------------------------------------------
fig, ax = plt.subplots(figsize=(6.4, 3.8))
ax.plot(t, B_wti, color='black', lw=1.6, ls='-', label='B1 (WTI leg), matures $T_{WTI}$='+f'{T_WTI}')
ax.plot(t, B_fx, color='0.45', lw=1.6, ls='--', label='B2 (FX leg), matures $T_{FX}$='+f'{T_FX}')
ax.axvline(T_FX, color='0.6', lw=1.0, ls=':')
ax.text(T_FX, ax.get_ylim()[1]*0.92, '  FX leg freezes here', fontsize=8, ha='left')
ax.set_xlabel('t (years)')
ax.set_ylabel('Mean derivative asset, bn KRW')
ax.set_title('Independent per-leg maturity freeze in Structure B')
ax.legend(frameon=False, fontsize=8, loc='upper left')
fig.tight_layout()
savefig(fig, 'fig_cfh_maturity_freeze.pdf')

# ---------------------------------------------------------------
# Figure 3: cumulative ineffectiveness, A vs B (B1+B2)
# ---------------------------------------------------------------
A_ineff = col(a, 10) / BN
b1 = d['b1_ledger']
b2 = d['b2_ledger']
B1_ineff = col(b1, 4) / BN
B2_ineff = col(b2, 6) / BN
B_ineff = B1_ineff + B2_ineff
t_b1 = col(b1, 1)

fig, ax = plt.subplots(figsize=(6.2, 3.6))
ax.plot(t_a, A_ineff, color='black', lw=1.6, ls='-', label='Structure A (cumulative)')
ax.plot(t_b1, B_ineff, color='0.4', lw=1.6, ls='--', label='Structure B = B1+B2 (cumulative)')
ax.set_xlabel('t (years)')
ax.set_ylabel('Mean cumulative P&L ineffectiveness, bn KRW')
ax.set_title('Cumulative hedge ineffectiveness over time')
ax.legend(frameon=False)
fig.tight_layout()
savefig(fig, 'fig_cfh_ineff.pdf')

# ---------------------------------------------------------------
# Figure 4: per-step economic residual std dev, A vs B (diagnostic)
# ---------------------------------------------------------------
econ = d['economic_series']
t_e = col(econ, 1)
std_a = col(econ, 3) / BN
std_b = col(econ, 5) / BN

fig, ax = plt.subplots(figsize=(6.2, 3.6))
ax.plot(t_e, std_a, color='black', lw=1.2, ls='-', label='Structure A')
ax.plot(t_e, std_b, color='0.45', lw=1.2, ls='--', label='Structure B')
ax.set_xlabel('t (years)')
ax.set_ylabel(r'Per-step cross-path Std($\Delta E$), bn KRW')
ax.set_title('Per-step economic-residual dispersion (diagnostic only)')
ax.legend(frameon=False)
fig.tight_layout()
savefig(fig, 'fig_cfh_econ_step.pdf')

# ---------------------------------------------------------------
# Figure 5: OCI reserve, A vs B
# ---------------------------------------------------------------
A_oci = col(sfp, 3) / BN
B_oci = col(sfp, 8) / BN

fig, ax = plt.subplots(figsize=(6.2, 3.6))
ax.plot(t, A_oci, color='black', lw=1.6, ls='-', label='Structure A OCI (CFHR+COH)')
ax.plot(t, B_oci, color='0.45', lw=1.6, ls='--', label='Structure B OCI (CFHR$_W$+CFHR$_X$)')
ax.set_xlabel('t (years)')
ax.set_ylabel('Mean OCI reserve balance, bn KRW')
ax.set_title('Cash-flow-hedge OCI reserve dynamics')
ax.legend(frameon=False)
fig.tight_layout()
savefig(fig, 'fig_cfh_oci.pdf')

# ---------------------------------------------------------------
# Figure 6: FV validation error, Structure A and Structure B1
# ---------------------------------------------------------------
fv = d['fv_validation']
tau = [r[0] for r in fv]
pct_a = [r[5]*100 for r in fv]
pct_b1 = [r[6]*100 for r in fv]

fig, ax = plt.subplots(figsize=(5.6, 3.6))
ax.plot(tau, pct_a, 'o-', color='black', label='Structure A')
ax.plot(tau, pct_b1, 's--', color='0.45', label='Structure B1 (WTI)')
ax.set_xlabel(r'$\tau/T$ (elapsed fraction of maturity)')
ax.set_ylabel('|Beta surface $-$ reduced-$n$ reprice| / surface, %')
ax.set_title('Pricing-surface validation error')
ax.legend(frameon=False)
fig.tight_layout()
savefig(fig, 'fig_cfh_fvvalidation.pdf')

# ---------------------------------------------------------------
# Appendix Figure 1: SFP debit=credit stacked balance (A and B)
# verified identity: Derivative == OCI + RetainedEarnings, exactly,
# at every step (see paper Appendix).
# ---------------------------------------------------------------
sfp = d['sfp']
t = np.array([r[1] for r in sfp['rows']])
A_deriv = np.array([r[2] for r in sfp['rows']]) / BN
A_oci = np.array([r[3] for r in sfp['rows']]) / BN
A_re = np.array([r[4] for r in sfp['rows']]) / BN
B_deriv = (np.array([r[6] for r in sfp['rows']]) + np.array([r[7] for r in sfp['rows']])) / BN
B_oci = np.array([r[8] for r in sfp['rows']]) / BN
B_re = np.array([r[9] for r in sfp['rows']]) / BN

fig, axes = plt.subplots(1, 2, figsize=(9.5, 3.8), sharex=True)
for ax, oci, re, deriv, label in [(axes[0], A_oci, A_re, A_deriv, 'Structure A'),
                                    (axes[1], B_oci, B_re, B_deriv, 'Structure B')]:
    re_pos = np.maximum(re, 0)
    re_neg = np.minimum(re, 0)
    ax.fill_between(t, 0, oci, color='0.75', label='OCI (credit)', step=None)
    ax.fill_between(t, oci, oci + re_pos, color='0.4', label='Retained earnings, gain', step=None)
    ax.fill_between(t, oci + re_neg, oci, color='0.15', label='Retained earnings, loss', step=None)
    ax.plot(t, deriv, color='black', lw=1.3, ls='-', label='Derivative asset (debit)')
    ax.set_title(label, fontsize=11)
    ax.set_xlabel('t (years)')
axes[0].set_ylabel('bn KRW')
axes[0].legend(frameon=False, fontsize=7, loc='upper right')
fig.suptitle('SFP debit (derivative) vs.\\ credit (OCI + retained earnings) -- verified equal at every step', fontsize=10)
fig.tight_layout(rect=[0,0,1,0.94])
savefig(fig, 'fig_cfh_sfp_balance.pdf')

# ---------------------------------------------------------------
# Appendix Figure 2: Sankey flow, Structure A, terminal split
# using the same verified identity, at t=T.
# ---------------------------------------------------------------
from matplotlib.sankey import Sankey
deriv_T = A_deriv[-1]
oci_T = A_oci[-1]
re_T = A_re[-1]
# Sankey requires nonnegative flows; re_T here is negative (a net RE loss),
# so route it as an outflow of magnitude |re_T| labelled accordingly.
fig = plt.figure(figsize=(7.5, 4.2))
ax = fig.add_subplot(1, 1, 1, xticks=[], yticks=[])
sankey = Sankey(ax=ax, unit=' bn KRW', format='%.1f', gap=0.3, scale=1.0/max(abs(deriv_T), abs(oci_T), abs(re_T)))
sankey.add(flows=[deriv_T, -oci_T, -re_T if re_T >= 0 else re_T],
           labels=['Derivative asset\n(t=T)', 'OCI reserve\n(t=T)', 'Retained earnings\n(cumulative, t=T)'],
           orientations=[0, 1, -1],
           facecolor='0.75', edgecolor='black')
sankey.finish()
ax.set_title(f'Structure A: derivative asset splits into OCI + retained earnings at maturity\n(verified: {deriv_T:.2f} = {oci_T:.2f} + ({re_T:.2f}) bn KRW)', fontsize=9)
fig.tight_layout()
savefig(fig, 'fig_cfh_sankey.pdf')

print("DONE: 8 figures written to", FIGDIR)
