import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface
from lsmc_quanto import _paths, CAL

# How often, and by how much, does the intrinsic floor bind on the fitted
# surface?  The floor is max(surface, intrinsic); it is a valid no-arbitrage
# bound for a barrier-free call but not for an up-and-out.
N=200_000
f=fit_surface(npaths=N, seed=12345)
c,K,beta=f['cal'],f['K'],f['beta']; S20=c['S2_0']; n=c['steps']; U=c['KOupper']
rng=np.random.default_rng(4242); S1,S2,AL=_paths(c,N,rng)
S1=S1.astype(float); S2=S2.astype(float)
tot=binds=0; near=near_binds=0; excess=[]
for i in range(1,n):
    v1=S1[:,i]/K; v2=S2[:,i]/S20
    b=beta[i]
    surf=b[0]+b[1]*v1+b[2]*v2+b[3]*v1*v1+b[4]*v2*v2+b[5]*v1*v2
    intr=np.maximum(S1[:,i]-K,0)*S2[:,i]
    dom=AL[:,i]&(S1[:,i]>K)
    if dom.sum()==0: continue
    bd=dom&(intr>surf)
    tot+=dom.sum(); binds+=bd.sum()
    if bd.sum(): excess.append(float(np.mean((intr-surf)[bd]/np.maximum(intr[bd],1))))
    nb=dom&(S1[:,i]>0.9*U); near+=nb.sum(); near_binds+=(nb&(intr>surf)).sum()
print(f"nodes {tot:,}")
print(f"  intrinsic floor binds on {binds:,} = {100*binds/tot:.2f}% of hedged nodes")
print(f"  within 10% of upper barrier: {near_binds:,}/{near:,} = {100*near_binds/max(near,1):.2f}%")
if excess: print(f"  mean relative lift where it binds: {100*np.mean(excess):.2f}% of intrinsic")
