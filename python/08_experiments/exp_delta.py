import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
from delta_hedge import fit_surface, surface_delta
from lsmc_quanto import price, CAL

def reg_delta_at(tfrac, npaths, seed=12345):
    f=fit_surface(npaths=npaths, seed=seed)
    c,K,beta=f['cal'],f['K'],f['beta']
    i=max(1,int(round(tfrac*c['steps'])))
    d1,_=surface_delta(beta[i], np.array([c['S1_0']]), np.array([c['S2_0']]), K, c['S2_0'])
    return f['price'], float(d1[0]/c['S2_0'])

def fd_delta(npaths, seed=12345, eps=0.01):
    S=CAL['S1_0']
    up,_=price(npaths=npaths, seed=seed, S1_0=S*(1+eps))
    dn,_=price(npaths=npaths, seed=seed, S1_0=S*(1-eps))
    return (up-dn)/(2*S*eps)/CAL['S2_0']

print(f"{'paths':>9} {'price':>12} {'reg delta(0)':>13} {'FD delta(0)':>12} {'gap %':>8}")
for n in (10_000, 50_000, 100_000, 200_000, 500_000):
    p,rd = reg_delta_at(0.0, n)
    fd = fd_delta(n)
    print(f"{n:>9,} {p:>12,.2f} {rd:>13.4f} {fd:>12.4f} {100*abs(rd-fd)/abs(fd):>7.2f}%")
print()
print("paper Table 5 (production bank): FD 0.6679  reg 0.7921  gap 18.60%")
print("paper Table 11 (exact engine)  : inception delta 0.692 @10k -> 0.543-0.545 from 50k")
