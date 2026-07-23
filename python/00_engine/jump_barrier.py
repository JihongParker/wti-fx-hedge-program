"""Jump-adapted exact first-passage simulation of the calibrated WTI jump-diffusion.

Poisson jump times are drawn exactly inside each daily step; on every jump-free
sub-interval the Brownian-bridge crossing probability is applied to each barrier;
the barrier is checked directly at each jump instant.  Nothing is monitored on a
grid, so the reported touch probability carries no time-discretisation bias.
"""
import numpy as np

CAL = dict(sig=0.3241851967428078, lam=6.784615384615385,
           thJ=-0.029895127380394192, dlJ=0.08442936299651074,
           mu_P=0.139, U=120.0, L=50.0, T=0.833, nstep=217)


def touch_prob(S0, npaths=200_000, seed=20260713, cal=None, measure='P', r=None,
               antithetic=True, return_paths_alive=False):
    """P(the double barrier [L,U] is touched before T), started at S0."""
    c = dict(CAL); c.update(cal or {})
    sig, lam, thJ, dlJ = c['sig'], c['lam'], c['thJ'], c['dlJ']
    U, L, T, n = c['U'], c['L'], c['T'], c['nstep']
    dt = T / n
    kappa = np.exp(thJ + 0.5*dlJ**2) - 1.0            # Merton compensator
    mu = (c['mu_P'] if measure == 'P' else r) - 0.5*sig**2 - lam*kappa

    rng = np.random.default_rng(seed)
    x = np.full(npaths, np.log(S0))
    alive = np.ones(npaths, dtype=bool)
    u, l = np.log(U), np.log(L)

    for _ in range(n):
        N = rng.poisson(lam*dt, npaths)
        K = int(N.max())
        if K:
            ut = rng.random((npaths, K))*dt
            ut[np.arange(K)[None, :] >= N[:, None]] = dt      # unused slots pin to dt
            ut.sort(axis=1)
        bounds = np.concatenate([np.zeros((npaths, 1)),
                                 ut if K else np.zeros((npaths, 0)),
                                 np.full((npaths, 1), dt)], axis=1)
        seg = np.diff(bounds, axis=1)                          # (npaths, K+1) >= 0

        for j in range(seg.shape[1]):
            d = seg[:, j]
            pos = d > 0
            if not pos.any():
                continue
            z = rng.standard_normal(npaths)
            if antithetic:
                z[npaths//2:] = -z[:npaths//2]
            xn = np.where(pos, x + mu*d + sig*np.sqrt(np.maximum(d, 0))*z, x)
            # exact Brownian-bridge crossing probability on the jump-free segment
            v = np.where(pos, sig*sig*d, 1.0)
            pu = np.where(pos & (x < u) & (xn < u), np.exp(-2*(u-x)*(u-xn)/v), 1.0)
            pl = np.where(pos & (x > l) & (xn > l), np.exp(-2*(x-l)*(xn-l)/v), 1.0)
            hit = rng.random(npaths) < (1 - (1-pu)*(1-pl))
            alive &= ~(hit & pos)
            x = xn
            if j < K:                                          # a jump lands here
                has = N > j
                if has.any():
                    J = rng.normal(thJ, dlJ, npaths)
                    x = np.where(has, x + J, x)
                    alive &= ~(has & ((x >= u) | (x <= l)))     # checked at the instant
    p = 1.0 - alive.mean()
    se = np.sqrt(p*(1-p)/npaths)
    return (p, se, alive) if return_paths_alive else (p, se)


if __name__ == '__main__':
    for lbl, S0 in [('stress spot 113', 113.0), ('today 78.94', 78.94)]:
        p, se = touch_prob(S0)
        print(f'{lbl:16s} p_touch = {p:.4f} +/- {se:.4f}')
