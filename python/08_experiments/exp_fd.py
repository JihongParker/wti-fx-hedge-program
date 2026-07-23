import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from lsmc_quanto import price, CAL
from delta_hedge import fit_surface, surface_delta

S=CAL['S1_0']; S20=CAL['S2_0']
for eps in (0.01,0.05):
    est=[]
    for sd in (11,22,33,44,55):
        up,_=price(npaths=100_000, seed=sd, S1_0=S*(1+eps))
        dn,_=price(npaths=100_000, seed=sd, S1_0=S*(1-eps))
        est.append((up-dn)/(2*S*eps)/S20)
    e=np.array(est)
    print(f"FD delta eps={eps}: mean {e.mean():.4f}  sd across seeds {e.std(ddof=1):.4f}  range [{e.min():.4f},{e.max():.4f}]")
print()
reg=[]
for sd in (11,22,33,44,55):
    f=fit_surface(npaths=100_000, seed=sd)
    c,K,b=f['cal'],f['K'],f['beta']
    d1,_=surface_delta(b[1], np.array([c['S1_0']]), np.array([c['S2_0']]), K, c['S2_0'])
    reg.append(float(d1[0]/c['S2_0']))
r=np.array(reg)
print(f"regression delta(0) across seeds: mean {r.mean():.4f}  sd {r.std(ddof=1):.4f}  range [{r.min():.4f},{r.max():.4f}]")
