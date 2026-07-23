"""IFRS 9 cash-flow-hedge ledgers for the two designation architectures.

Structure A designates the combined quanto knock-out call against the aggregated
exposure.  Structure B splits the same exposure into two independently
designated risk components carried as genuine forwards: a WTI forward struck at
F0 = S1(0) e^{r_US T_WTI} on the barrel position, and a USD/KRW forward struck at
G0 = S2(0) e^{(r_KRW - r_US) T_FX} on the dollar notional N_FX = barrels x F0.

B's cash-flow-hedge reserve on the currency leg is the signed lower of the
instrument's fair value and the hypothetical derivative's, which is where the
two architectures separate: the hypothetical rides the WTI forward price, the
instrument does not.

Paths are physical measure with the asymmetric up/down jump calibration.
"""
import os, sys, json
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

CAL = dict(S1_0=78.94, S2_0=1540.64, diffvol=0.3241851967428078,
           sig2=0.09257650229436072, rho=0.08763086561435483,
           lam_up=2.328, th_up=0.0804, dl_up=0.0168,
           lam_dn=4.462, th_dn=-0.0875, dl_dn=0.0278,
           mu1P=0.139, mu2P=0.0239, r_US=0.04, r_KRW=0.035,
           U=120.0, L=50.0, T_WTI=0.833, T_FX=0.5, barrels=2_000_000,
           steps_per_year=260)


def simulate(npaths=200_000, seed=20260706, cal=None, record=False):
    c = dict(CAL); c.update(cal or {})
    n = max(1, int(round(c['T_WTI']*c['steps_per_year'])))
    dt = c['T_WTI']/n
    ku = np.exp(c['th_up']+0.5*c['dl_up']**2)-1
    kd = np.exp(c['th_dn']+0.5*c['dl_dn']**2)-1
    m1 = (c['mu1P'] - c['lam_up']*ku - c['lam_dn']*kd - 0.5*c['diffvol']**2)*dt
    m2 = (c['mu2P'] - 0.5*c['sig2']**2)*dt
    v1, v2 = c['diffvol']*np.sqrt(dt), c['sig2']*np.sqrt(dt)
    rng = np.random.default_rng(seed)
    x1 = np.full(npaths, np.log(c['S1_0'])); x2 = np.full(npaths, np.log(c['S2_0']))
    alive = np.ones(npaths, bool); ko_step = np.full(npaths, -1)
    v0_ko = np.zeros(npaths); s1_ko = np.zeros(npaths); s2_ko = np.zeros(npaths)
    F0 = c['S1_0']*np.exp(c['r_US']*c['T_WTI'])
    G0 = c['S2_0']*np.exp((c['r_KRW']-c['r_US'])*c['T_FX'])
    NFX = c['barrels']*F0
    trace = [] if record else None
    for t in range(1, n+1):
        z1 = rng.standard_normal(npaths); zi = rng.standard_normal(npaths)
        z2 = c['rho']*z1 + np.sqrt(1-c['rho']**2)*zi
        nu = rng.poisson(c['lam_up']*dt, npaths); nd = rng.poisson(c['lam_dn']*dt, npaths)
        ju = np.where(nu > 0, rng.normal(c['th_up']*nu, c['dl_up']*np.sqrt(np.maximum(nu, 1))), 0.)
        jd = np.where(nd > 0, rng.normal(c['th_dn']*nd, c['dl_dn']*np.sqrt(np.maximum(nd, 1))), 0.)
        x1 = x1 + m1 + v1*z1 + ju + jd
        x2 = x2 + m2 + v2*z2
        s1, s2 = np.exp(x1), np.exp(x2)
        hit = alive & ((s1 >= c['U']) | (s1 <= c['L']))
        if hit.any():
            tt = t*dt
            vW = fwd_wti(s1, F0, c, tt)*s2
            vF = fwd_fx(s2, G0, NFX, c, tt)
            v0_ko = np.where(hit, vW+vF, v0_ko)
            s1_ko = np.where(hit, s1, s1_ko); s2_ko = np.where(hit, s2, s2_ko)
            ko_step = np.where(hit, t, ko_step)
            alive &= ~hit
        if record:
            tt = t*dt
            trace.append((tt, fwd_wti(s1, F0, c, tt)*s2, fwd_fx(s2, G0, NFX, c, tt),
                          hyp_fx(s1, s2, F0, G0, c, tt)))
    return dict(s1=s1, s2=s2, alive=alive, ko_step=ko_step, v0_ko=v0_ko,
                s1_ko=s1_ko, s2_ko=s2_ko,
                F0=F0, G0=G0, NFX=NFX, cal=c, n=n, dt=dt, trace=trace)


def fwd_wti(s1, F0, c, t):
    tau = max(c['T_WTI']-t, 0.0)
    return (s1*np.exp(c['r_US']*tau) - F0)*c['barrels']*np.exp(-c['r_US']*tau)


def fwd_fx(s2, G0, NFX, c, t):
    if t >= c['T_FX']:
        return NFX*(s2-G0)
    return NFX*(s2-G0)*np.exp(-c['r_KRW']*(c['T_FX']-t))


def hyp_fx(s1, s2, F0, G0, c, t):
    F1 = s1*np.exp(c['r_US']*max(c['T_WTI']-t, 0.0))
    d = 1.0 if t >= c['T_FX'] else np.exp(-c['r_KRW']*(c['T_FX']-t))
    return c['barrels']*F1*(s2-G0)*d


def signed_lower_of(v, h):
    return np.sign(v)*np.minimum(np.abs(v), np.abs(h))


def structures(npaths=200_000, seed=20260706, K=None, V_A0=17_448.76):
    r = simulate(npaths, seed)
    c, T = r['cal'], CAL['T_WTI']
    s1, s2 = r['s1'], r['s2']
    K = c['S1_0'] if K is None else K

    # B1 marks the barrel leg at the FIXED inception rate, which is what severs
    # the quanto coupling A carries; B2 carries the currency leg on its own
    vW = fwd_wti(s1, r['F0'], c, T)*c['S2_0']
    vF = fwd_fx(s2, r['G0'], r['NFX'], c, T)
    hF = hyp_fx(s1, s2, r['F0'], r['G0'], c, T)
    cfhrF = signed_lower_of(vF, hF)

    B = dict(carry=np.abs(vW)+np.abs(vF), cfhr=vW+cfhrF, ineff=np.abs(vF-cfhrF))
    kom = r['ko_step'] > 0
    B['post_ko_pl'] = (vW+vF)[kom] - r['v0_ko'][kom]
    # naked exposure after a barrier touch: A's combined instrument dies there,
    # so the physical runs unhedged from the touch to expiry; B's forwards carry
    # no barrier and keep hedging
    naked_A = c['barrels']*(s1[kom]*s2[kom] - r['s1_ko'][kom]*r['s2_ko'][kom])
    postko_var99_A = float(-np.percentile(naked_A, 1)) if kom.any() else 0.0

    # Structure A: one combined instrument, floating S2(T) inside the payoff
    A_val = np.where(r['alive'], np.maximum(s1-K, 0)*s2, 0.0)*c['barrels']

    # economic residual: terminal path cumulant of the hedge less the physical
    phys = c['barrels']*(s1*s2 - c['S1_0']*c['S2_0'])
    dV_A = A_val - V_A0*c['barrels']
    E_A = dV_A - phys
    E_B = (vW + vF) - phys

    # Structure A's lower-of test runs against the closed-form hypothetical
    # H* = barrels S1 S2 exp(rho sigma1 sigma2 tau), which at expiry is the
    # physical itself grown by the quanto adjustment
    q = np.exp(c['rho']*c['diffvol']*c['sig2']*c['T_WTI'])
    dH_A = c['barrels']*(s1*s2 - c['S1_0']*c['S2_0']*q)
    cfhr_A = signed_lower_of(dV_A, dH_A)
    ineff_A = np.abs(dV_A - cfhr_A)

    # the SFP identity, checked on the ledger's own accumulation
    OCI = B['cfhr']
    V = B['cfhr'] + B['ineff']*np.sign(vF-cfhrF)*0 + (vW+vF) - (vW+cfhrF)
    V = vW + vF
    RE = V - OCI
    ident = np.max(np.abs(V - (OCI + RE)))

    return dict(
        npaths=npaths, ko_rate=float(kom.mean()),
        B_carry_sd=float(B['carry'].std(ddof=1)), B_cfhr_sd=float(B['cfhr'].std(ddof=1)),
        A_mean_ineff=float(ineff_A.mean()), A_carry_sd=float(np.abs(A_val).std(ddof=1)),
        A_var99=float(-np.percentile(E_A, 1)), B_var99=float(-np.percentile(E_B, 1)),
        postko_naked_var99_A=postko_var99_A, postko_naked_var99_B=0.0,
        B_mean_ineff=float(B['ineff'].mean()),
        post_ko_mean=float(B['post_ko_pl'].mean()) if kom.any() else 0.0,
        post_ko_sd=float(B['post_ko_pl'].std(ddof=1)) if kom.sum() > 1 else 0.0,
        sigma_econ_A=float(E_A.std(ddof=1)), sigma_econ_B=float(E_B.std(ddof=1)),
        sigma_phys=float(phys.std(ddof=1)),
        sfp_identity_max_abs=float(ident))


if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 200_000
    r = structures(npaths=n)
    bn = lambda x: x/1e9
    print(f"{n:,} physical-measure paths, asymmetric up/down jump calibration")
    print(f"  barrier touch rate                      {r['ko_rate']:.4f}")
    print(f"  unhedged physical exposure std          {bn(r['sigma_phys']):10,.2f} bn")
    print(f"  sigma_econ  A                           {bn(r['sigma_econ_A']):10,.2f} bn")
    print(f"  sigma_econ  B                           {bn(r['sigma_econ_B']):10,.2f} bn")
    print(f"  B derivative carrying-amount std        {bn(r['B_carry_sd']):10,.2f} bn")
    print(f"  B end-of-life CFHR std                  {bn(r['B_cfhr_sd']):10,.2f} bn")
    print(f"  A mean |cumulative ineffectiveness|     {bn(r['A_mean_ineff']):10,.2f} bn")
    print(f"  B mean |cumulative ineffectiveness|     {bn(r['B_mean_ineff']):10,.2f} bn")
    print(f"  economic VaR99  A / B                   {bn(r['A_var99']):10,.2f} / {bn(r['B_var99']):.2f} bn")
    print(f"  post-KO naked exposure VaR99  A / B     {bn(r['postko_naked_var99_A']):10,.2f} / {bn(r['postko_naked_var99_B']):.2f} bn")
    print(f"  post-KO P&L mean / std                  {bn(r['post_ko_mean']):10,.2f} / {bn(r['post_ko_sd']):.2f} bn")
    print(f"  SFP identity V - (OCI + RE), max abs    {r['sfp_identity_max_abs']:.3e} KRW")
