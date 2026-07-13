"""
make_figures_bw.py
===================================================================
Regenerate ALL paper figures in clean BLACK-AND-WHITE (grayscale +
hatching + distinct linestyles/markers -- no colour), from the saved
result artifacts. Four figures:

  fig_model_verification : structural solver + comparative statics
  fig_did_recovery       : event study + estimator recovery (DiD MC)
  fig_hsub               : H(sub) identification fragility
  fig_real_data          : REAL OpenDART extraction findings
                           (designation status; note-table heterogeneity;
                            parse-quality summary)

Reads engine_results.json (Part 1/2) and derivatives_parsed.json /
real_panel.json (live extraction). Curves for the model panels are
recomputed from the certified solver in esg_hedge_engine.
"""
import json
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from dataclasses import asdict
from esg_hedge_engine import Params, solve_closed_form

# ---- global black-and-white style -----------------------------------
plt.rcParams.update({
    "font.size": 10, "axes.titlesize": 10, "axes.labelsize": 10,
    "axes.edgecolor": "black", "axes.linewidth": 0.8,
    "figure.facecolor": "white", "savefig.facecolor": "white",
    "axes.grid": False, "legend.frameon": False,
})
K = "black"
GRAYS = ["#2b2b2b", "#6e6e6e", "#a9a9a9", "#d0d0d0"]   # dark -> light
HATCH = ["", "///", "xxx", "..."]


def _barlabels(ax, bars, vals, fmt="{:.4f}"):
    for b, v in zip(bars, vals):
        ax.text(b.get_x()+b.get_width()/2, v,
                fmt.format(v), ha="center",
                va="bottom" if v >= 0 else "top", fontsize=9)


# =====================================================================
def fig_model_verification(p1):
    pr = Params()
    fig, ax = plt.subplots(1, 3, figsize=(13, 3.9))
    lams = np.linspace(0.1, 1.4, 60)
    hf = [solve_closed_form(Params(**{**asdict(pr), "lam": L}))["h_f"] for L in lams]
    ax[0].plot(lams, hf, color=K, lw=1.8)
    ax[0].plot([pr.lam], [p1["method1"]["h_f"]], marker="o", ms=8, mfc=K, mec=K, ls="",
               label="closed form")
    ax[0].plot([pr.lam], [p1["method2"]["h_f"]], marker="s", ms=12, mfc="none", mec=K,
               mew=1.5, ls="", label="numeric solver")
    ax[0].set_xlabel(r"stringency  $\lambda$"); ax[0].set_ylabel(r"$h_f^\star$")
    ax[0].set_title(r"(a) $\partial h_f^\star/\partial\lambda>0$  (H1)"); ax[0].legend(fontsize=8)

    ays = np.linspace(0.02, 0.12, 60)
    hfa = [solve_closed_form(Params(**{**asdict(pr), "a": A}))["h_f"] for A in ays]
    dsa = [solve_closed_form(Params(**{**asdict(pr), "a": A}))["d"] for A in ays]
    ax[1].plot(ays, hfa, color=K, lw=1.8, ls="-", label=r"$h_f^\star$ (left)")
    ax[1].set_xlabel(r"disclosure cost  $a$"); ax[1].set_ylabel(r"$h_f^\star$")
    ax[1].set_title(r"(b) $\partial h_f^\star/\partial a>0$, $\partial d^\star/\partial a<0$  (H$_{\rm sub}$)")
    axb = ax[1].twinx()
    axb.plot(ays, dsa, color="#6e6e6e", lw=1.6, ls="--", label=r"$d^\star$ (right)")
    axb.set_ylabel(r"$d^\star$")
    l1, la1 = ax[1].get_legend_handles_labels(); l2, la2 = axb.get_legend_handles_labels()
    ax[1].legend(l1+l2, la1+la2, fontsize=8, loc="center right")

    rhos = np.linspace(-0.4, 0.7, 60)
    hfr = [solve_closed_form(Params(**{**asdict(pr), "rho": R}))["h_f"] for R in rhos]
    ax[2].plot(rhos, hfr, color=K, lw=1.8)
    ax[2].axvline(pr.rho, color="#6e6e6e", ls=":", lw=1.2)
    ax[2].set_xlabel(r"correlation  $\rho$"); ax[2].set_ylabel(r"$h_f^\star$")
    ax[2].set_title(r"(c) sign of $\partial h_f^\star/\partial\rho$ per Eq.(20)")
    fig.suptitle("Structural solver: closed form (line) certified by the numeric optimizer "
                 r"(open square), agreement $\sim10^{-8}$", fontsize=10)
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    _save(fig, "fig_model_verification")


def fig_did_recovery(p2):
    es = p2["event_study"]; xs = sorted(int(k) for k in es.keys()); ys = [es[str(x)] for x in xs]
    fig, ax = plt.subplots(1, 2, figsize=(11, 4.0))
    ax[0].axhline(0, color="#999999", lw=0.8)
    ax[0].axvline(-0.5, color=K, ls=":", lw=1.2)
    ax[0].plot(xs, ys, marker="o", ms=5, color=K, mfc=K, lw=1.6)
    ax[0].set_xlabel("event time  (years since phase-in)")
    ax[0].set_ylabel(r"$\theta_\ell$  (effect on $h_f$)")
    ax[0].set_title("(a) Event study: flat pre-trend, effect ramps in")
    ax[0].annotate("pre-trend $\\approx$ 0", xy=(-2, 0.002), xytext=(-3, 0.03),
                   fontsize=8, arrowprops=dict(arrowstyle="->", color=K))

    labels = ["True ATT", "Callaway--\nSant'Anna", "TWFE"]
    vals = [p2["true_att"], p2["att_callaway_santanna"], p2["beta_twfe"]]
    bars = ax[1].bar(labels, vals, color=[GRAYS[0], GRAYS[1], "white"],
                     edgecolor=K, hatch=["", "", "xxx"], lw=1.0)
    _barlabels(ax[1], bars, vals)
    ax[1].set_ylabel("ATT on hedge ratio")
    ax[1].set_title("(b) CS recovers the truth; TWFE biased\nunder dynamic, staggered timing")
    ax[1].set_ylim(0, max(vals)*1.18)
    fig.tight_layout()
    _save(fig, "fig_did_recovery")


def fig_hsub(p2):
    h = p2["Hsub"]
    fig, ax = plt.subplots(figsize=(7.6, 4.3))
    groups = ["exogenous\n(naive)", "confounded\n(naive)", "confounded\n(capability-\ncontrolled)"]
    vals = [h["exogenous_naive"], h["confounded_naive"], h["confounded_controlled"]]
    bars = ax.bar(groups, vals, color=[GRAYS[0], "white", GRAYS[1]],
                  edgecolor=K, hatch=["", "xxx", "///"], lw=1.0)
    ax.axhline(0, color=K, lw=1.0)
    _barlabels(ax, bars, vals)
    ax.set_ylabel(r"cross-sectional slope of $h_f$ on disclosure cost $a$")
    ax.set_title("H(sub): structural slope is positive; a capability confound\n"
                 "flips the naive estimate; controlling for capability restores it")
    fig.tight_layout()
    _save(fig, "fig_hsub")


# =====================================================================
def fig_real_data(deriv, panel):
    # order firms by 2023 assets (desc); use latin display names
    disp = {"제주항공": "Jeju Air", "S-Oil": "S-Oil", "POSCO홀딩스": "POSCO",
            "SK이노베이션": "SK Innov.", "현대차": "Hyundai Mtr", "KG스틸": "KG Steel",
            "LG화학": "LG Chem", "삼성전자": "Samsung", "대한항공": "Korean Air",
            "팬오션": "Pan Ocean", "HMM": "HMM", "현대제철": "Hyundai Stl",
            "대한해운": "Korea Line", "기아": "Kia"}
    recs = [r for r in deriv if "error" not in r]
    recs = sorted(recs, key=lambda r: -(r.get("assets2023") or 0))
    names = [disp.get(r["name"], r["name"]) for r in recs]

    fig, ax = plt.subplots(1, 3, figsize=(14, 4.6),
                           gridspec_kw=dict(width_ratios=[1.05, 1.35, 0.9]))

    # (a) designation status -- the one robustly extractable variable
    y = np.arange(len(names))
    for i, r in enumerate(recs):
        s = r.get("hedge_accounting_applied")
        if s is True:
            ax[0].plot(1, i, marker="s", ms=11, mfc=K, mec=K)
        elif s is False:
            ax[0].plot(1, i, marker="s", ms=11, mfc="white", mec=K, mew=1.4)
        else:
            ax[0].plot(1, i, marker="x", ms=9, mec=K, mew=1.6)
    ax[0].set_yticks(y); ax[0].set_yticklabels(names, fontsize=8)
    ax[0].set_xlim(0.5, 1.5); ax[0].set_xticks([])
    ax[0].set_title("(a) Hedge-accounting designation\n(read from note sentence)")
    ax[0].invert_yaxis()
    from matplotlib.lines import Line2D
    leg = [Line2D([0], [0], marker="s", mfc=K, mec=K, ls="", ms=10, label="applied"),
           Line2D([0], [0], marker="s", mfc="white", mec=K, ls="", ms=10, label="not applied"),
           Line2D([0], [0], marker="x", mec=K, ls="", ms=9, label="ambiguous")]
    ax[0].legend(handles=leg, fontsize=8, loc="lower center", ncol=1,
                 bbox_to_anchor=(0.5, -0.28))

    # (b) note-table-type heterogeneity matrix -> why generic parsing fails
    types = ["fair_value", "pl_impact", "notional", "hedge_holdings", "other"]
    tlab = ["fair\nvalue", "P&L\nimpact", "notional", "hedge\nholdings", "other"]
    M = np.zeros((len(recs), len(types)))
    for i, r in enumerate(recs):
        for j, t in enumerate(types):
            if t in r.get("note_table_types", []):
                M[i, j] = 1
    ax[1].imshow(1 - M, cmap="gray", vmin=0, vmax=1, aspect="auto")
    ax[1].set_xticks(range(len(types))); ax[1].set_xticklabels(tlab, fontsize=8)
    ax[1].set_yticks(range(len(names))); ax[1].set_yticklabels(names, fontsize=8)
    for i in range(len(recs)):
        for j in range(len(types)):
            if M[i, j]:
                ax[1].text(j, i, "$\\blacksquare$", ha="center", va="center",
                           color="white", fontsize=7)
    ax[1].set_title("(b) Which note-table TYPES each firm files\n"
                    "(filled = present) — most mix types")
    n_multi = int(sum(M.sum(1) > 1))
    ax[1].set_xlabel(f"{n_multi} of {len(recs)} firms span $\\geq$2 table types")

    # (c) parse-quality summary
    from collections import Counter
    q = Counter(r.get("parse_quality", "n/a") for r in recs)
    order = ["validated_fair_value", "heterogeneous_not_comparable", "note_not_located"]
    qlab = ["clean, comparable\n(1 type, KRW)", "heterogeneous\n(not comparable)", "note not\nlocated"]
    vals = [q.get(k, 0) for k in order]
    bars = ax[2].bar(qlab, vals, color=[GRAYS[0], GRAYS[2], "white"],
                     edgecolor=K, hatch=["", "///", "xxx"], lw=1.0)
    for b, v in zip(bars, vals):
        ax[2].text(b.get_x()+b.get_width()/2, v, str(v), ha="center", va="bottom", fontsize=10)
    ax[2].set_ylabel("firms"); ax[2].set_ylim(0, max(vals)+1.4)
    ax[2].set_title("(c) Numeric comparability of\nthe derivative note")
    ax[2].tick_params(axis="x", labelsize=8)

    fig.suptitle("Live OpenDART extraction (FY2023 pilot): the hedge-accounting DESIGNATION is "
                 "reliably readable; the NUMERIC derivative tables are not cross-firm comparable",
                 fontsize=10)
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    _save(fig, "fig_real_data")


def _save(fig, name):
    fig.savefig(f"figures/{name}.pdf")
    fig.savefig(f"figures/{name}.png", dpi=150)
    plt.close(fig)
    print("wrote", name)


def main():
    res = json.load(open("engine_results.json"))
    deriv = json.load(open("derivatives_parsed.json"))
    panel = json.load(open("real_panel.json"))
    fig_model_verification(res["part1"])
    fig_did_recovery(res["part2"])
    fig_hsub(res["part2"])
    fig_real_data(deriv, panel)
    print("all figures regenerated in black-and-white.")


if __name__ == "__main__":
    main()
