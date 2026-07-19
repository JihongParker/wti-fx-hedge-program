"""
make_theory_ladder_figs.py
===================================================================
Two additional paper figures, black-and-white, manuscript-font
(serif text, Computer Modern math), matching make_figures_bw.py:

  fig_theory : the model in one row --
               (a) the residual-risk price Lambda(d)
               (b) hedge ratios h(d) with the corner and the
                   crowding-out of a binding floor
               (c) the disclosure fixed point (LHS/RHS of the
                   scalar FOC), existence & uniqueness visualized
  fig_ladder : the exposure ladder --
               (a) executed ATTs (95% CI) on the two realized
                   mandates vs the model's predicted-positive
                   climate rung
               (b) the finalized climate regime as a calendar:
                   content floor, then liability & assurance,
                   i.e. the dated path prediction

Calibration comes from the certified engine (esg_hedge_engine.Params);
ladder numbers are the published estimates of Section 7.
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from dataclasses import asdict
from esg_hedge_engine import Params

plt.rcParams.update({
    # match the manuscript: Computer Modern math, serif text (no sans)
    "font.family": "serif", "mathtext.fontset": "cm",
    "font.size": 10, "axes.titlesize": 10, "axes.labelsize": 10,
    "axes.edgecolor": "black", "axes.linewidth": 0.8,
    "figure.facecolor": "white", "savefig.facecolor": "white",
    "axes.grid": False, "legend.frameon": False,
})
K = "black"
G = "#6e6e6e"
LG = "#c9c9c9"


def _model_pieces(p: Params):
    Sig = p.Sigma()
    pv = np.array([p.p_f, p.p_c])
    Sinv_p = np.linalg.solve(Sig, pv)
    kappa = float(pv @ Sinv_p)

    def Lam(d):
        return p.phi + p.lam * np.exp(-p.k * d)

    def hedges(d):
        u = Sinv_p / (2.0 * Lam(d))
        return 1.0 - u  # unclipped; corners handled by caller

    # voluntary d*: root of 8 a d Lam(d)^2 = k lam kappa e^{-kd}
    ds = np.linspace(1e-6, 3.0, 20000)
    lhs = 8 * p.a * ds * Lam(ds) ** 2
    rhs = p.k * p.lam * kappa * np.exp(-p.k * ds)
    dstar = float(ds[np.argmin(np.abs(lhs - rhs))])
    return Lam, hedges, kappa, dstar


def fig_theory():
    p = Params()
    Lam, hedges, kappa, dstar = _model_pieces(p)
    d = np.linspace(0, 2.2, 400)

    fig, ax = plt.subplots(1, 3, figsize=(13, 3.9))

    # (a) the price of residual risk
    ax[0].plot(d, Lam(d), color=K, lw=1.8)
    ax[0].axhline(p.phi, color=G, lw=1.0, ls=":")
    ax[0].annotate(r"$\phi$ (distress floor)", xy=(1.55, p.phi),
                   xytext=(1.05, p.phi + 0.085), color=G, fontsize=9,
                   arrowprops=dict(arrowstyle="-", color=G, lw=0.7))
    ax[0].annotate(r"$\lambda$", xy=(0.02, Lam(0.0) - 0.02),
                   xytext=(0.22, Lam(0.0) - 0.05), fontsize=10,
                   arrowprops=dict(arrowstyle="-", color=K, lw=0.7))
    ax[0].set_xlabel(r"disclosure intensity  $d$")
    ax[0].set_ylabel(r"$\Lambda(d)=\phi+\lambda e^{-kd}$")
    ax[0].set_title(r"(a) the residual-risk price")
    ax[0].set_ylim(0, Lam(0.0) * 1.12)

    # (b) hedge ratios vs disclosure; corner; crowding-out of a floor
    hf = np.array([hedges(x)[0] for x in d])
    hc = np.array([hedges(x)[1] for x in d])
    hc_clip = np.clip(hc, 0.0, 1.0)
    ax[1].plot(d, hf, color=K, lw=1.8, label=r"$h_f^\star(d)$  financial")
    ax[1].plot(d, hc_clip, color=G, lw=1.8, ls="--",
               label=r"$h_c^\star(d)$  climate")
    d_corner = d[np.argmax(hc <= 0)] if np.any(hc <= 0) else None
    if d_corner is not None:
        ax[1].axvline(d_corner, color=LG, lw=1.0)
        ax[1].text(d_corner + 0.05, 0.30, "corner " r"$h_c=0$",
                   fontsize=8.5, color=G)
    dbar = dstar + 0.8
    ax[1].axvline(dstar, color=K, lw=0.9, ls=":")
    ax[1].axvline(dbar, color=K, lw=0.9, ls="-.")
    ax[1].text(dstar - 0.055, 0.028, r"$d^\star_v$", fontsize=10)
    ax[1].text(dbar + 0.03, 0.028, "d̲ (binding)", fontsize=9,
               style="italic")
    hf_at = np.interp([dstar, dbar], d, hf)
    ax[1].annotate("", xy=(dbar, hf_at[1]), xytext=(dstar, hf_at[0]),
                   arrowprops=dict(arrowstyle="->", color=K, lw=1.2))
    ax[1].text((dstar + dbar) / 2 - 0.12, hf_at[0] + 0.045,
               "crowding out", fontsize=9)
    ax[1].set_xlabel(r"disclosure intensity  $d$")
    ax[1].set_ylabel(r"hedge ratio")
    ax[1].set_ylim(0, 1.0)
    ax[1].set_title(r"(b) disclosure substitutes for hedging")
    ax[1].legend(fontsize=8, loc="upper right")

    # (c) the scalar fixed point: LHS increasing, RHS decreasing
    dd = np.linspace(0, 0.6, 400)
    LamV = Lam(dd)
    lhs = 8 * p.a * dd * LamV ** 2
    rhs = p.k * p.lam * kappa * np.exp(-p.k * dd)
    ax[2].plot(dd, lhs, color=K, lw=1.8,
               label=r"$8ad\,\Lambda(d)^2$  (marginal cost)")
    ax[2].plot(dd, rhs, color=G, lw=1.8, ls="--",
               label=r"$k\lambda\kappa e^{-kd}$  (penalty relief)")
    ax[2].plot([dstar], [np.interp(dstar, dd, rhs)], marker="o", ms=7,
               mfc=K, mec=K, ls="")
    ax[2].axvline(dstar, color=LG, lw=1.0)
    ax[2].text(dstar + 0.012, 0.001, r"$d^\star$", fontsize=10)
    ax[2].set_xlabel(r"disclosure intensity  $d$")
    ax[2].set_ylabel("marginal value")
    ax[2].set_title(r"(c) the disclosure fixed point")
    ax[2].legend(fontsize=8, loc="upper right")

    fig.tight_layout()
    fig.savefig("figures/fig_theory.pdf")
    fig.savefig("figures/fig_theory.png", dpi=150)
    plt.close(fig)


def fig_ladder():
    # published estimates (Section 7): ATT [lo, hi]
    rows = [
        ("Governance report\n(2019/22/24, G pillar)",
         [("hedge-acct.", -0.005, -0.068, 0.060),
          ("deriv. use",  -0.050, -0.113, 0.013)]),
        ("Environmental info.\n(2022, physical quantities)",
         [("hedge-acct.", 0.010, -0.048, 0.063),
          ("deriv. use", -0.009, -0.069, 0.049)]),
    ]

    fig, ax = plt.subplots(1, 2, figsize=(11.5, 4.1),
                           gridspec_kw={"width_ratios": [1.05, 1.0]})

    # (a) ladder of ATTs
    a0 = ax[0]
    y = 0
    ypos, ylab, group_mid = [], [], []
    for rung, ests in rows:
        y0 = y
        for name, att, lo, hi in ests:
            a0.plot([lo, hi], [y, y], color=K, lw=1.4)
            for xend in (lo, hi):
                a0.plot([xend, xend], [y - 0.12, y + 0.12], color=K, lw=1.4)
            mk = "o" if name == "hedge-acct." else "s"
            a0.plot([att], [y], marker=mk, ms=6.5, mfc=K, mec=K, ls="")
            ypos.append(y)
            ylab.append(f"{name}")
            y += 1
        group_mid.append((rung, (y0 + y - 1) / 2))
        y += 0.8
    # climate rung: predicted, not yet estimable
    a0.axhspan(y - 0.35, y + 0.75, color=LG, alpha=0.45)
    a0.annotate("", xy=(0.085, y + 0.2), xytext=(0.005, y + 0.2),
                arrowprops=dict(arrowstyle="->", color=K, lw=1.6))
    a0.text(0.093, y + 0.2, "model predicts $+$ (H1--H2)", fontsize=9,
            va="center")
    ypos.append(y + 0.2)
    ylab.append("climate (KSSB)\nfirst covered FY2027")
    a0.axvline(0, color=G, lw=0.9, ls=":")
    a0.set_yticks(ypos)
    a0.set_yticklabels(ylab, fontsize=8.5)
    a0.set_xlabel("group-time ATT (95% CI)")
    a0.set_title("(a) the exposure ladder: nulls below, prediction on top")
    a0.set_xlim(-0.14, 0.24)
    for rung, mid in group_mid:
        a0.text(0.232, mid, rung.replace("\n", " "), fontsize=8, color=G,
                ha="right", va="center", style="italic")

    # (b) the finalized regime as a calendar (dated path prediction)
    a1 = ax[1]
    # (year, label, label height) -- staggered to avoid collisions
    events = [
        (2026.5, "plan finalized\n(Jul 2026)", 0.14),
        (2027.0, "first covered\nfiscal year", 0.52),
        (2028.0, "first filings\nKOSPI $\\geq$ 10tn", 0.14),
        (2029.0, "$\\geq$ 5tn tier", 0.52),
        (2030.0, "assurance $+$\nliability bind", 0.14),
        (2031.0, "Scope 3\nphases in", 0.52),
    ]
    a1.axhline(0, color=K, lw=1.2)
    for x, lab, h in events:
        a1.plot([x], [0], marker="o", ms=5, mfc="white", mec=K, mew=1.2)
        a1.plot([x, x], [0.03, h - 0.02], color=LG, lw=0.8)
        a1.text(x, h, lab, fontsize=8, ha="center", va="bottom")
    # schematic lambda/k step: low during safe harbor, up at 2030
    xs = np.linspace(2026.4, 2032.6, 500)
    step = np.where(xs < 2028.0, -0.80, np.where(xs < 2030.0, -0.55, -0.25))
    a1.plot(xs, step, color=G, lw=1.8, ls="--")
    a1.text(2026.55, -0.47, r"penalty $\lambda$, informativeness $k$"
            "\n(schematic: safe harbor,\nthen enforcement)",
            fontsize=8, color=G, va="top")
    a1.axvspan(2030.0, 2031.6, color=LG, alpha=0.45)
    a1.text(2030.8, 0.92, "predicted response\nwindow (2030--31)",
            fontsize=8.5, ha="center", va="bottom")
    a1.set_ylim(-1.0, 1.35)
    a1.set_xlim(2026.3, 2032.7)
    a1.set_yticks([])
    a1.set_xticks([2027, 2028, 2029, 2030, 2031, 2032])
    a1.set_xticklabels(["2027", "2028", "2029", "2030", "2031", "2032"])
    a1.set_title("(b) the finalized climate regime: a dated path prediction")
    for s in ("left", "right", "top"):
        a1.spines[s].set_visible(False)

    fig.tight_layout()
    fig.savefig("figures/fig_ladder.pdf")
    fig.savefig("figures/fig_ladder.png", dpi=150)
    plt.close(fig)


if __name__ == "__main__":
    import os
    os.makedirs("figures", exist_ok=True)
    fig_theory()
    fig_ladder()
    print("wrote figures/fig_theory.pdf, figures/fig_ladder.pdf")
