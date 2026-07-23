import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface, surface_delta
from lsmc_quanto import price, CAL
S0=CAL['S1_0']; S20=CAL['S2_0']; N=200_000

def fd_at(tfrac, seeds=(11,22,33,44,55), eps=0.05, npaths=100_000):
    T=CAL['T']*(1-tfrac); est=[]
    for sd in seeds:
        up,_=price(npaths=npaths,seed=sd,S1_0=S0*(1+eps),T=T)
        dn,_=price(npaths=npaths,seed=sd,S1_0=S0*(1-eps),T=T)
        est.append((up-dn)/(2*S0*eps)/S20)
    e=np.array(est); return e.mean(), e.std(ddof=1)

def reg_at(tfrac, seeds=(11,22,33,44,55), npaths=100_000):
    est=[]
    for sd in seeds:
        f=fit_surface(npaths=npaths,seed=sd)
        c,K,b=f['cal'],f['K'],f['beta']
        i=max(1,int(round(tfrac*c['steps'])))
        d1,_=surface_delta(b[i],np.array([c['S1_0']]),np.array([c['S2_0']]),K,c['S2_0'])
        est.append(float(d1[0]/c['S2_0']))
    e=np.array(est); return e.mean(), e.std(ddof=1)

print(f"{'t/T':>5} {'FD delta':>18} {'regression delta':>20} {'ratio':>7}")
for tf in (0.00,0.50,0.90):
    fm,fs=fd_at(tf); rm,rs=reg_at(tf)
    print(f"{tf:5.2f} {fm:9.4f} +/-{fs:6.4f} {rm:12.4f} +/-{rs:6.4f} {rm/fm:7.2f}x")
print("\npaper Table 5: 0.00 FD 0.6679 reg 0.7921 | 0.50 FD 0.5019 reg 0.5483 | 0.90 FD 0.5660 reg 0.5979")
