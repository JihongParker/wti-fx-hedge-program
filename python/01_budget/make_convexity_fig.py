#!/usr/bin/env python3
"""
GMVP residual-risk convexity surface over the full (w1,w2) in [0,1]^2 field.

  GMVP(w1,w2) = sqrt( ((1-w1) s1)^2 + ((1-w2) s2)^2 + 2(1-w1)(1-w2) s1 s2 rho )

This is a convex bowl whose global minimum is the *unconstrained* point
(1,1) -> 0 (hedge everything). The hedge-budget constraint w1+w2 <= 1 cuts the
square into the feasible triangle with vertices (0,0),(0,1),(1,0); the bowl's
bottom corner (1,1) is removed, so the constrained optimum lies on the
hypotenuse w1+w2=1.

Rendering rule (user spec): colour ONLY the feasible triangle; show the
infeasible region (w1+w2>1) as a transparent wireframe grid.
"""
import os, math
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "opt_figures")
os.makedirs(OUT, exist_ok=True)

wb = openpyxl.load_workbook(os.path.join(HERE, "Simulation.xlsm"),
                            data_only=True, read_only=True)
g = lambda s, c: wb[s][c].value

ENGINES = {
    "American": dict(s1=g("LSMC", "B10"), s2=g("LSMC", "B11"),
                     rho=g("LSMC", "B12"), wopt=g("LSMC", "J13"),
                     obj="LSMC!J15"),
    "European": dict(s1=g("Raw_Timeseries", "H2"), s2=g("Raw_Timeseries", "I2"),
                     rho=g("Raw_Timeseries", "J2"), wopt=g("Encoding", "C21"),
                     obj="Encoding!C24"),
}


def gmvp(w1, w2, s1, s2, rho):
    return np.sqrt(((1 - w1) * s1) ** 2 + ((1 - w2) * s2) ** 2
                   + 2 * (1 - w1) * (1 - w2) * s1 * s2 * rho)


def edge_min(s1, s2, rho):
    """argmin over the hypotenuse w2 = 1 - w1."""
    w1 = (s1 ** 2 - s1 * s2 * rho) / (s1 ** 2 + s2 ** 2 - 2 * s1 * s2 * rho)
    return w1, 1 - w1


def make_panel(ax, p, title):
    s1, s2, rho = p["s1"], p["s2"], p["rho"]
    N = 81
    w = np.linspace(0, 1, N)
    W1, W2 = np.meshgrid(w, w)
    Z = gmvp(W1, W2, s1, s2, rho)

    feas = (W1 + W2) <= 1.0 + 1e-9
    Zf = np.where(feas, Z, np.nan)          # coloured feasible triangle
    Zi = np.where(~feas, Z, np.nan)         # transparent wireframe infeasible

    # FULL field is rendered as a continuous surface (the whole convex bowl):
    #   - feasible triangle  -> solid viridis colour (height = colour)
    #   - infeasible region  -> faint transparent grey surface + grid overlay
    ax.plot_surface(W1, W2, Zi, color="0.72", rstride=1, cstride=1,
                    linewidth=0, antialiased=True, alpha=0.18)
    ax.plot_wireframe(W1, W2, Zi, color="0.45", rstride=3, cstride=3,
                      linewidth=0.45)
    ax.plot_surface(W1, W2, Zf, cmap="viridis", rstride=1, cstride=1,
                    linewidth=0, antialiased=True, alpha=1.0)

    # hypotenuse (budget edge) drawn on the surface
    t = np.linspace(0, 1, 240)
    ax.plot(t, 1 - t, gmvp(t, 1 - t, s1, s2, rho), color="black",
            linewidth=2.2, label="budget edge $w_1+w_2=1$")

    # constrained optimum on the edge
    e1, e2 = edge_min(s1, s2, rho)
    ax.scatter([e1], [e2], [gmvp(e1, e2, s1, s2, rho)], color="red", s=55,
               depthshade=False, zorder=10,
               label=f"constrained opt ($w_1${'='}{e1:.3f})")
    # stored Solver point
    so1 = p["wopt"]
    ax.scatter([so1], [1 - so1], [gmvp(so1, 1 - so1, s1, s2, rho)],
               color="white", edgecolor="black", s=55, depthshade=False,
               zorder=10, label=f"stored Solver ($w_1${'='}{so1:.3f})")
    # unconstrained min (1,1) -> 0
    ax.scatter([1], [1], [0], color="black", marker="x", s=55, depthshade=False,
               zorder=10, label="unconstrained min $(1,1)\\to 0$")

    ax.set_xlabel("$w_1$  (WTI hedge ratio)", fontsize=8, labelpad=4)
    ax.set_ylabel("$w_2$  (FX hedge ratio)", fontsize=8, labelpad=4)
    ax.set_zlabel("GMVP residual vol", fontsize=8, labelpad=4)
    ax.set_title(title, fontsize=10)
    ax.view_init(elev=30, azim=-58)
    ax.set_box_aspect((1, 1, 0.62))
    ax.tick_params(labelsize=6.5)
    ax.legend(fontsize=6.4, loc="upper right", framealpha=0.9)


def main():
    plt.rcParams.update({"font.family": "serif"})
    # combined two-panel
    fig = plt.figure(figsize=(12, 5.2))
    for k, (name, p) in enumerate(ENGINES.items(), 1):
        ax = fig.add_subplot(1, 2, k, projection="3d")
        make_panel(ax, p, f"{name}  ({p['obj']})")
    fig.suptitle("Residual-risk (GMVP) convexity over $[0,1]^2$ — feasible "
                 "triangle coloured, infeasible region transparent grid",
                 fontsize=11)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(OUT, f"gmvp_convexity.{ext}"), dpi=150,
                    bbox_inches="tight")
    print("wrote", os.path.join(OUT, "gmvp_convexity.png"))

    # individual panels too
    for name, p in ENGINES.items():
        f = plt.figure(figsize=(6.5, 5.5))
        ax = f.add_subplot(111, projection="3d")
        make_panel(ax, p, f"{name}  ({p['obj']})")
        f.savefig(os.path.join(OUT, f"gmvp_convexity_{name.lower()}.png"),
                  dpi=150, bbox_inches="tight")
        plt.close(f)
        print("wrote", f"gmvp_convexity_{name.lower()}.png")


if __name__ == "__main__":
    main()
