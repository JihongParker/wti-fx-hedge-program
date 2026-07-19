"""
make_results_figures.py — B&W figures for the executed study (v2 schema).
  fig_real_eventstudy : event-study paths with bootstrap bands (HA, DerivUse)
  fig_real_rules      : ATT under three comparison-set rules + placebo + E-pillar
  fig_real_adoption   : adoption paths by year
Grayscale only.
"""
import json
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams.update({
    # match the manuscript: Computer Modern math, serif text (no sans)
    "font.family": "serif", "mathtext.fontset": "cm",
    "font.size": 10, "axes.titlesize": 10, "axes.labelsize": 10,
    "axes.edgecolor": "black", "axes.linewidth": 0.8,
    "figure.facecolor": "white", "savefig.facecolor": "white",
    "legend.frameon": False,
})
K = "black"; G = "#6e6e6e"; LG = "#c9c9c9"


def main():
    res = json.load(open("results_real.json"))

    # ---- event studies with bands --------------------------------------
    fig, ax = plt.subplots(1, 2, figsize=(11, 4.0), sharey=True)
    for a, key, ttl in ((ax[0], "DerivUse", "(a) Derivative use"),
                        (ax[1], "HA", "(b) Hedge-accounting designation")):
        es = {int(k): v for k, v in res[key]["event_study"].items()}
        xs = sorted(es); ys = [es[x] for x in xs]
        bands = res[key].get("event_bands", {})
        lo = [bands.get(str(x), [np.nan, np.nan])[0] for x in xs]
        hi = [bands.get(str(x), [np.nan, np.nan])[1] for x in xs]
        a.axhline(0, color="#999999", lw=0.8)
        a.axvline(-0.5, color=K, ls=":", lw=1.2)
        a.fill_between(xs, lo, hi, color=LG, alpha=0.55, lw=0)
        a.plot(xs, ys, marker="o", ms=5, color=K, lw=1.6)
        att = res[key]["att_cs"]; ci = res[key]["ci95"]
        a.set_title(f"{ttl}\nATT = {att:+.3f}  [{ci[0]:+.3f}, {ci[1]:+.3f}]")
        a.set_xlabel("event time (years since cohort phase-in)")
    ax[0].set_ylabel(r"$\theta_\ell$")
    n = res["HA"]["n_firms"]
    fig.suptitle(f"Event studies, {n}-firm OpenDART panel (governance-report mandate; "
                 "shaded: 95% bootstrap bands)", fontsize=10)
    fig.tight_layout(rect=[0, 0, 1, 0.92])
    _save(fig, "fig_real_eventstudy")

    # ---- ATT across comparison rules + placebo ---------------------------
    outs = [("HA", "Hedge-acct.\ndesignation"), ("DerivUse", "Derivative\nuse"),
            ("lnFXfwd", "ln(1+FX-fwd\nmentions)")]
    specs = [("main", "", "#2b2b2b", ""), ("adj", "_adjacent", "white", "///"),
             ("loc", "_local", "#9a9a9a", ""), ("placebo", None, "white", "xxx")]
    labels = ["not-yet controls", "adjacent cohort", "threshold-local", "placebo (fake mandate)"]
    fig, a = plt.subplots(figsize=(9.6, 4.4))
    xw = np.arange(len(outs)); w = 0.19
    for j, (tag, suf, col, hat) in enumerate(specs):
        vals, los, his = [], [], []
        for key, _ in outs:
            if tag == "placebo":
                d = res["placebo"][key]
            else:
                d = res[key + suf] if suf else res[key]
            vals.append(d.get("att", d.get("att_cs")))
            ci = d.get("ci95", (np.nan, np.nan))
            los.append(ci[0]); his.append(ci[1])
        pos = xw + (j-1.5)*w
        a.bar(pos, vals, w, color=col, edgecolor=K, hatch=hat, lw=0.9, label=labels[j])
        for p, v, l, h in zip(pos, vals, los, his):
            if v is not None and np.isfinite(v) and np.isfinite(l):
                a.errorbar(p, v, yerr=[[v-l], [h-v]], color=G, capsize=3, lw=1.1)
    a.axhline(0, color=K, lw=1.0)
    a.set_xticks(xw); a.set_xticklabels([nm for _, nm in outs], fontsize=9)
    a.set_ylabel("ATT")
    a.set_title("Mandate ATT under three comparison-set rules, against the placebo")
    a.legend(fontsize=8, ncol=2)
    fig.tight_layout()
    _save(fig, "fig_real_rules")

    # ---- adoption paths --------------------------------------------------
    ad = {int(k): v for k, v in res["adoption_by_year"].items()}
    yrs = sorted(ad)
    tr = [ad[y]["treated"][0] for y in yrs]
    ct = [ad[y]["control"][0] for y in yrs]
    fig, a = plt.subplots(figsize=(7.6, 4.0))
    a.plot(yrs, [np.nan if v is None else v for v in tr], marker="s", ms=6,
           color=K, lw=1.8, label="treated (post-mandate firm-years)")
    a.plot(yrs, [np.nan if v is None else v for v in ct], marker="o", ms=6, mfc="white",
           color=G, lw=1.6, ls="--", label="not-yet-treated")
    for g in (2019, 2022, 2024):
        a.axvline(g-0.5, color=K, ls=":", lw=1.0)
    a.set_xlabel("fiscal year"); a.set_ylabel("hedge-accounting adoption rate")
    a.set_title("Hedge-accounting adoption by fiscal year")
    a.legend(fontsize=9)
    fig.tight_layout()
    _save(fig, "fig_real_adoption")


def _save(fig, name):
    fig.savefig(f"figures/{name}.pdf")
    fig.savefig(f"figures/{name}.png", dpi=150)
    plt.close(fig)
    print("wrote", name)


if __name__ == "__main__":
    main()
