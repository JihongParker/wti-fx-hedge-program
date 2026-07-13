import os
import numpy as np
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
XLSM = os.path.join(HERE, "Hedge.xlsm")
OUT = os.path.join(HERE, "figures")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family": "serif",
    "font.size": 10,
    "axes.edgecolor": "black",
    "axes.linewidth": 1.0,
    "axes.grid": True,
    "grid.color": "0.9",
    "grid.linewidth": 0.5,
    "figure.dpi": 300,
    "savefig.dpi": 300,
    "savefig.bbox": "tight",
})

print("Loading workbook...")
wb = openpyxl.load_workbook(XLSM, data_only=True, read_only=True)

def read_block(ws, r0, r1, c0, c1):
    rows = ws.iter_rows(min_row=r0, max_row=r1, min_col=c0, max_col=c1, values_only=True)
    return list(rows)

def col_values(ws, col, r0, r1):
    out = []
    for (v,) in read_block(ws, r0, r1, col, col):
        if isinstance(v, (int, float)):
            out.append(float(v))
    return np.array(out)

# Figure 1
print("Figure 1...")
ws_rob = wb["LambdaRobustness"]
k_sig = col_values(ws_rob, 1, 2, 6)
vol1 = col_values(ws_rob, 2, 2, 6)
lam = col_values(ws_rob, 3, 2, 6)

fig, ax1 = plt.subplots(figsize=(7, 4))
ax1.plot(k_sig, lam, 'ro-', linewidth=2, label=r'Jump Intensity $\lambda$')
ax1.set_xlabel(r'$k$-Sigma Threshold')
ax1.set_ylabel(r'Jump Intensity $\lambda$', color='red')
ax1.tick_params(axis='y', labelcolor='red')

ax2 = ax1.twinx()
ax2.plot(k_sig, vol1, 'bs-', linewidth=2, label=r'Diffusion Volatility $\sigma_1$')
ax2.set_ylabel(r'Diffusion Volatility $\sigma_1$', color='blue')
ax2.tick_params(axis='y', labelcolor='blue')
ax2.set_ylim(0, max(vol1)*1.5)

fig.suptitle('Robustness Sweep Plot (k-Sigma)', fontsize=12)
fig.savefig(os.path.join(OUT, "Figure_1.pdf"))
fig.savefig(os.path.join(OUT, "Figure_1.eps"))
plt.close(fig)

# Figure 2
print("Figure 2...")
am_vals = col_values(wb["American_Delta"], 13, 31, 10030) / 1e9
eu_vals = col_values(wb["European_Delta"], 13, 31, 10030) / 1e9
as_vals = col_values(wb["Asian_Delta"], 14, 31, 10030) / 1e9

fig, ax = plt.subplots(figsize=(8, 4))
ax.hist(am_vals, bins=80, alpha=0.7, color='green', label='American LSMC', density=True)
ax.hist(eu_vals, bins=80, alpha=0.5, color='orange', label='European Black-76', density=True)
ax.hist(as_vals, bins=80, alpha=0.5, color='purple', label='Asian Turnbull-Wakeman', density=True)

# Highlight red area for -300 bn KRW
ax.axvspan(-350, -250, color='red', alpha=0.2, label='Catastrophic Tail Risk')

ax.set_title('Terminal P&L Density Histogram', fontsize=12)
ax.set_xlabel('Hedged P&L (bn KRW)')
ax.set_ylabel('Density')
ax.legend()
fig.savefig(os.path.join(OUT, "Figure_2.pdf"))
fig.savefig(os.path.join(OUT, "Figure_2.eps"))
plt.close(fig)

# Figure 3
print("Figure 3...")
fig, ax = plt.subplots(figsize=(7, 4))
S1 = np.linspace(30, 120, 200)
K = 80
posFX = np.minimum(S1, K)

ax.plot(S1, posFX, 'k-', linewidth=2.5, label=r'$pos_{FX} = \min(S_1, K)$')
ax.axvline(K, color='gray', linestyle='--', label=r'Strike $K$')
ax.set_xlabel('WTI Price ($S_1$)')
ax.set_ylabel(r'Quanto Notional ($pos_{FX}$)')
ax.set_title('Quanto Notional Capping Mechanism', fontsize=12)
ax.legend()
fig.savefig(os.path.join(OUT, "Figure_3.pdf"))
fig.savefig(os.path.join(OUT, "Figure_3.eps"))
plt.close(fig)

print("Done.")
