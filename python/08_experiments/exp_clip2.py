import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface, surface_delta
from lsmc_quanto import CAL, _paths
N=200_000
fit=fit_surface(npaths=N, seed=12345)
c,K,beta=fit['cal'],fit['K'],fit['beta']
rng=np.random.default_rng(777); S1,S2,AL=_paths(c,N,rng)
S1=S1.astype(float); S2=S2.astype(float); n=c['steps']; U=c['KOupper']
tot=hi=0; mins=[]; near_min=[]
for i in range(n):
    dV1,_=surface_delta(beta[i],S1[:,i],S2[:,i],K,c['S2_0'])
    dom=AL[:,i]&(S1[:,i]>K)
    if dom.sum()==0: continue
    raw=(dV1/S2[:,i])[dom]
    tot+=raw.size; hi+=(raw>1.0).sum(); mins.append(raw.min())
    nb=AL[:,i]&(S1[:,i]>0.9*U)
    if nb.sum(): near_min.append(((dV1/S2[:,i])[nb]).min())
mins=np.array(mins)
print(f"hedged nodes {tot:,}")
print(f"  raw > 1 (upper clip binds): {hi:,} = {100*hi/tot:.2f}%")
print(f"  global min raw delta over all steps: {mins.min():.4f}")
print(f"  min raw delta within 10% of upper barrier: {min(near_min):.4f}")
# is the basis even capable of a negative delta on the fitted domain?
print()
print("beta_1, 2*beta_3, beta_5 at a few steps (delta = (b1 + 2 b3 v1 + b5 v2)/K):")
for i in (1, 54, 108, 162, 210):
    b=beta[i]; print(f"   step {i:3d}: b1={b[1]:+.4e}  2b3={2*b[3]:+.4e}  b5={b[5]:+.4e}")
