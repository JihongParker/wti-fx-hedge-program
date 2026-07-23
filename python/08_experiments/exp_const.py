import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface, surface_delta, b76_delta
from lsmc_quanto import _paths
from math import exp
N=100_000
fit=fit_surface(npaths=N, seed=12345)
c,K,beta=fit['cal'],fit['K'],fit['beta']
n=c['steps']; dt=c['T']/n; S20=c['S2_0']; gr=exp(0.07*dt); Qo=2_000_000

def hedge(mode, const=None, npaths=N, seed=777):
    rng=np.random.default_rng(seed); S1,S2,AL=_paths(c,npaths,rng)
    S1=S1.astype(float); S2=S2.astype(float)
    cash=np.zeros(npaths); h1=np.zeros(npaths); h2=np.zeros(npaths)
    for i in range(n):
        dom=AL[:,i]&(S1[:,i]>K)
        if mode=='surface':
            dV1,dV2=surface_delta(beta[i],S1[:,i],S2[:,i],K,S20)
            g1=Qo*np.where(dom,np.clip(dV1/S2[:,i],0,1),0.0)
            g2=Qo*np.where(dom,np.clip(dV2,-20,20),0.0)
        elif mode=='b76':
            g1=Qo*b76_delta(S1[:,i],K,c['vol1'],c['T']-i*dt,c['r_US']); g2=np.zeros(npaths)
        elif mode=='const':                       # constant delta, no state dependence
            g1=Qo*np.where(dom,const,0.0); g2=np.zeros(npaths)
        elif mode=='const+fx':                    # constant WTI delta + the surface FX leg
            dV1,dV2=surface_delta(beta[i],S1[:,i],S2[:,i],K,S20)
            g1=Qo*np.where(dom,const,0.0); g2=Qo*np.where(dom,np.clip(dV2,-20,20),0.0)
        alive=AL[:,i]
        g1=np.where(alive,g1,0.0); g2=np.where(alive,g2,0.0)
        cash-=(g1-h1)*S1[:,i]*S2[:,i]+(g2-h2)*S2[:,i]; h1,h2=g1,g2; cash*=gr
    cash+=h1*S1[:,n]*S2[:,n]+h2*S2[:,n]
    pay=np.where(AL[:,n],np.maximum(S1[:,n]-K,0)*S2[:,n],0.0)*Qo
    return pay-fit['price']*Qo*exp(0.07*c['T'])-cash

res={}
for lbl,kw in [('surface (state-dependent)',dict(mode='surface')),
               ('Black-76 (state-dependent)',dict(mode='b76')),
               ('CONSTANT delta 0.70',dict(mode='const',const=0.70)),
               ('CONSTANT 0.70 + surface FX leg',dict(mode='const+fx',const=0.70))]:
    h=hedge(**kw); res[lbl]=(h.mean()/1e9,h.std(ddof=1)/1e9)
    print(f"  {lbl:32s} mean {res[lbl][0]:8.2f}bn  sd {res[lbl][1]:7.2f}bn")
sd_s=res['surface (state-dependent)'][1]; sd_b=res['Black-76 (state-dependent)'][1]
sd_c=res['CONSTANT delta 0.70'][1]; sd_cf=res['CONSTANT 0.70 + surface FX leg'][1]
print(f"\n  b76/surface       = {sd_b/sd_s:.4f}")
print(f"  b76/constant      = {sd_b/sd_c:.4f}   <- if ~= b76/surface, the gain is NOT barrier awareness")
print(f"  b76/constant+fx   = {sd_b/sd_cf:.4f}  <- isolates the FX leg's contribution")
