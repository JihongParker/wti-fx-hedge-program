#!/usr/bin/env python3
"""
Bell-shaped (interior-optimum) hedge surface over the feasible triangle.

Objective (mean-variance):
    Z(w1,w2) = Var(w1,w2) + gamma * ( c1*w1 + c2*w2 )

    Var(w1,w2) = (1-w1)^2 s1^2 + (1-w2)^2 s2^2 + 2(1-w1)(1-w2) s1 s2 rho   (convex,
                 pulls toward FULL hedge (1,1))
    c1*w1 + c2*w2  = per-notional hedge premium (linear, pulls toward NO hedge (0,0))

The two opposing pulls give an INTERIOR minimum -> a genuine bell. With the
per-notional premium rates from the workbook and gamma=0.30 the optimum lands
strictly inside the budget triangle (0,0)-(1,0)-(0,1).

Render rule: full [0,1]^2 as one continuous surface; feasible triangle coloured,
infeasible region (w1+w2>1) transparent grey grid.
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

s1, s2, rho = g("LSMC", "B10"), g("LSMC", "B11"), g("LSMC", "B12")
c1 = g("Black76", "B11") / g("Encoding", "B2")   # WTI premium rate ~0.16
c2 = g("GK", "B12") / g("Encoding", "B3")         # FX  premium rate ~0.055
GAMMA = 0.30

Sig = np.array([[s1 ** 2, s1 * s2 * rho], [s1 * s2 * rho, s2 ** 2]])
b = np.array([c1, c2])


def Z(W1, W2):
    var = ((1 - W1) ** 2 * s1 ** 2 + (1 - W2) ** 2 * s2 ** 2
           + 2 * (1 - W1) * (1 - W2) * s1 * s2 * rho)
    return var + GAMMA * (c1 * W1 + c2 * W2)


def interior_opt():
    r = np.linalg.inv(Sig) @ b * GAMMA / 2.0     # = (1-w*)
    return 1 - r


def main():
    plt.rcParams.update({"font.family": "serif"})
    N = 91
    w = np.linspace(0, 1, N)
    W1, W2 = np.meshgrid(w, w)
    ZZ = Z(W1, W2)
    feas = (W1 + W2) <= 1.0 + 1e-9
    Zf = np.where(feas, ZZ, np.nan)
    Zi = np.where(~feas, ZZ, np.nan)

    fig = plt.figure(figsize=(7.5, 6.2))
    ax = fig.add_subplot(111, projection="3d")

    ax.plot_surface(W1, W2, Zi, color="0.72", rstride=1, cstride=1,
                    linewidth=0, antialiased=True, alpha=0.16)
    ax.plot_wireframe(W1, W2, Zi, color="0.45", rstride=4, cstride=4,
                      linewidth=0.45)
    ax.plot_surface(W1, W2, Zf, cmap="viridis", rstride=1, cstride=1,
                    linewidth=0, antialiased=True, alpha=1.0)

    # budget edge
    t = np.linspace(0, 1, 200)
    ax.plot(t, 1 - t, Z(t, 1 - t), color="black", lw=2.0,
            label="budget edge $w_1+w_2=1$")

    # interior optimum
    wo = interior_opt()
    ax.scatter([wo[0]], [wo[1]], [Z(wo[0], wo[1])], color="red", s=70,
               depthshade=False, zorder=10,
               label=f"interior optimum ({wo[0]:.2f}, {wo[1]:.2f})")

    ax.set_xlabel("$w_1$  (WTI hedge ratio)", fontsize=9, labelpad=4)
    ax.set_ylabel("$w_2$  (FX hedge ratio)", fontsize=9, labelpad=4)
    ax.set_zlabel("$Z = \\mathrm{Var} + \\gamma\\,$cost", fontsize=9, labelpad=4)
    ax.set_title("Mean-variance hedge objective — bell with interior optimum "
                 f"inside the triangle ($\\gamma$={GAMMA})", fontsize=10)
    ax.view_init(elev=32, azim=-122)
    ax.set_box_aspect((1, 1, 0.62))
    ax.tick_params(labelsize=7)
    ax.legend(fontsize=7.5, loc="upper right", framealpha=0.9)

    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(OUT, f"gmvp_bell.{ext}"), dpi=150,
                    bbox_inches="tight")
    print("interior optimum:", interior_opt())
    print("wrote", os.path.join(OUT, "gmvp_bell.png"))


if __name__ == "__main__":
    main()
