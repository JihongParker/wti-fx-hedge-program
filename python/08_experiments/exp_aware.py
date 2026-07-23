import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface, surface_delta, b76_delta
from lsmc_quanto import _paths
N=200_000
f=fit_surface(npaths=N, seed=12345)
c,K,beta=f['cal'],f['K'],f['beta']; S20=c['S2_0']; n=c['steps']; dt=c['T']/n
rng=np.random.default_rng(777); S1,S2,AL=_paths(c,N,rng); S1=S1.astype(float); S2=S2.astype(float)
print(f"{'t/T':>5} {'nodes':>9} {'surf mean':>10} {'B76 mean':>9} {'corr':>7} {'surf|S1>110':>12} {'B76|S1>110':>11}")
for tf in (0.0,0.25,0.5,0.75,0.95):
    i=max(1,int(tf*n))
    d1,_=surface_delta(beta[i],S1[:,i],S2[:,i],K,S20)
    surf=d1/S2[:,i]
    b76=b76_delta(S1[:,i],K,c['vol1'],c['T']-i*dt,c['r_US'])
    dom=AL[:,i]&(S1[:,i]>K)
    if dom.sum()<50: continue
    hi=dom&(S1[:,i]>110)
    cr=np.corrcoef(surf[dom],b76[dom])[0,1]
    sh=surf[hi].mean() if hi.sum()>20 else float('nan')
    bh=b76[hi].mean() if hi.sum()>20 else float('nan')
    print(f"{tf:5.2f} {dom.sum():9,} {surf[dom].mean():10.4f} {b76[dom].mean():9.4f} {cr:7.4f} {sh:12.4f} {bh:11.4f}")
print("\nA barrier-aware delta must FALL toward zero (and turn negative) as S1 -> U=120.")
