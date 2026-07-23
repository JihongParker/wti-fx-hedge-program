"""Pull WTI and USD/KRW from FRED and calibrate every risk input from source.

Series: DCOILWTICO (Cushing WTI spot, USD/bbl), DEXKOUS (USD/KRW noon rate).
Outputs spot levels, annualised volatilities, the WTI-FX correlation, and the
Merton jump/diffusion split obtained by EM on the WTI log returns.
"""
import json, os, sys, urllib.request, urllib.parse
import numpy as np

FRED = 'https://api.stlouisfed.org/fred/series/observations'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..',
                   'data', 'results', 'market_calibration.json')


def _key():
    k = os.environ.get('FRED_API_KEY')
    if k:
        return k
    for p in (os.path.expanduser('~/trading-desk/.env'),):
        if os.path.exists(p):
            for line in open(p):
                if line.startswith('FRED_API_KEY='):
                    return line.split('=', 1)[1].strip()
    raise SystemExit('FRED_API_KEY not found')


def series(sid, start='2015-01-01'):
    q = urllib.parse.urlencode(dict(series_id=sid, api_key=_key(), file_type='json',
                                    observation_start=start))
    with urllib.request.urlopen(f'{FRED}?{q}', timeout=60) as r:
        obs = json.load(r)['observations']
    d = [(o['date'], float(o['value'])) for o in obs if o['value'] not in ('.', '')]
    return dict(d)


def em_jump_split(r, dt, iters=400, tol=1e-12):
    """Two-component Gaussian mixture on daily log returns: diffusion vs jump.

    r_t ~ (1-p) N(mu_d, s_d^2) + p N(mu_d + thJ, s_d^2 + dlJ^2).
    Returns the annualised diffusive vol and the Merton triple (lam, thJ, dlJ)
    with lam = p / dt.
    """
    n = len(r)
    mu = r.mean()
    sd = r.std(ddof=1)
    p, th, dl = 0.05, -0.5*sd, 1.5*sd
    prev = None
    for _ in range(iters):
        v0, v1 = sd*sd, sd*sd + dl*dl
        f0 = (1-p)*np.exp(-0.5*(r-mu)**2/v0)/np.sqrt(2*np.pi*v0)
        f1 = p*np.exp(-0.5*(r-mu-th)**2/v1)/np.sqrt(2*np.pi*v1)
        w = f1/np.maximum(f0+f1, 1e-300)
        p = max(w.mean(), 1e-6)
        mu = ((1-w)*r).sum()/max((1-w).sum(), 1e-9)
        th = (w*(r-mu)).sum()/max(w.sum(), 1e-9)
        sd = np.sqrt(max(((1-w)*(r-mu)**2).sum()/max((1-w).sum(), 1e-9), 1e-12))
        dl = np.sqrt(max((w*((r-mu-th)**2 - sd*sd)).sum()/max(w.sum(), 1e-9), 1e-12))
        ll = np.log(np.maximum(f0+f1, 1e-300)).sum()
        if prev is not None and abs(ll-prev) < tol:
            break
        prev = ll
    return dict(sigma_diff=sd/np.sqrt(dt), lam=p/dt, thJ=th, dlJ=dl,
                jump_frac=float(p), loglik=float(ll))


def calibrate(start='2015-01-01', window=1299, steps_per_year=260):
    wti, fx = series('DCOILWTICO', start), series('DEXKOUS', start)
    days = sorted(set(wti) & set(fx))
    days = days[-(window+1):]
    s1 = np.array([wti[d] for d in days]); s2 = np.array([fx[d] for d in days])
    r1 = np.diff(np.log(s1)); r2 = np.diff(np.log(s2))
    dt = 1.0/steps_per_year
    em = em_jump_split(r1, dt)
    out = dict(
        source='FRED DCOILWTICO / DEXKOUS', as_of=days[-1], n_returns=int(len(r1)),
        first_day=days[0], steps_per_year=steps_per_year,
        WTI_spot=float(s1[-1]), KRW_spot=float(s2[-1]),
        sigma_WTI_hist=float(r1.std(ddof=1)/np.sqrt(dt)),
        sigma_FX=float(r2.std(ddof=1)/np.sqrt(dt)),
        rho=float(np.corrcoef(r1, r2)[0, 1]),
        sigma_WTI_diff=float(em['sigma_diff']), lam=float(em['lam']),
        thJ=float(em['thJ']), dlJ=float(em['dlJ']),
        mu_WTI_P=float(r1.mean()/dt), mu_FX_P=float(r2.mean()/dt))
    return out


if __name__ == '__main__':
    c = calibrate()
    json.dump(c, open(OUT, 'w'), indent=1)
    paper = dict(WTI_spot=78.94, KRW_spot=1540.64, sigma_WTI_hist=0.39455,
                 sigma_FX=0.09258, rho=0.08763, sigma_WTI_diff=0.32419,
                 lam=6.7846, thJ=-0.029895, dlJ=0.084429, mu_WTI_P=0.139)
    w = max(len(k) for k in paper)
    print(f"FRED {c['source']}   {c['first_day']} .. {c['as_of']}   n={c['n_returns']}")
    print(f"{'input'.ljust(w)}  {'paper':>12}  {'FRED now':>12}")
    for k, v in paper.items():
        print(f"{k.ljust(w)}  {v:12.5f}  {c[k]:12.5f}")
    print(f"\nwritten to {os.path.normpath(OUT)}")
