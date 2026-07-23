import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface, run, surface_delta
from lsmc_quanto import CAL, _paths

N=200_000
fit=fit_surface(npaths=N, seed=12345)
print(f"surface fitted on {N:,} paths, price {fit['price']:,.2f}")

# ---- how often does the raw delta go negative, and where? ----
c,K,beta=fit['cal'],fit['K'],fit['beta']
rng=np.random.default_rng(777)
S1,S2,AL=_paths(c,N,rng); S1=S1.astype(float); S2=S2.astype(float)
n=c['steps']; U=c['KOupper']
tot=neg=0; neg_near=0; near=0
for i in range(n):
    dV1,_=surface_delta(beta[i],S1[:,i],S2[:,i],K,c['S2_0'])
    dom=AL[:,i]&(S1[:,i]>K)
    raw=dV1/S2[:,i]
    tot+=dom.sum(); neg+=(dom&(raw<0)).sum()
    nb=dom&(S1[:,i]>0.9*U)          # within 10% of the upper barrier
    near+=nb.sum(); neg_near+=(nb&(raw<0)).sum()
print(f"raw delta < 0 on {neg:,}/{tot:,} hedged nodes = {100*neg/tot:.2f}%")
print(f"  within 10% of the upper barrier: {neg_near:,}/{near:,} = {100*neg_near/max(near,1):.2f}%")

# ---- hedge cost with and without the clip ----
for lbl,cl in (('clipped [0,1]  (production)',True),('unclipped',False)):
    h=run(fit,npaths=N,seed=777,clip=cl)
    print(f"{lbl:28s} mean {h.mean()/1e9:8.2f}bn   sd {h.std(ddof=1)/1e9:8.2f}bn")
