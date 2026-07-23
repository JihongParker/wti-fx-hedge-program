import sys, numpy as np
sys.path.insert(0,'/Users/elijahjasper/Modeling/python/00_engine')
import cfh_ledger as L
from cfh_ledger import simulate, signed_lower_of, hyp_fx, CAL

def paired(npaths=200_000, seed=20260706, quanto_adj=True):
    r=simulate(npaths,seed); c,T=r['cal'],CAL['T_WTI']
    s1,s2=r['s1'],r['s2']; K=c['S1_0']
    vW=np.where(r['alive'],np.maximum(s1-K,0),0.0)*c['barrels']*c['S2_0']
    vF=np.maximum(s2-c['S2_0'],0)*r['NFX']
    hF=hyp_fx(s1,s2,r['F0'],r['G0'],c,T); cfhrF=signed_lower_of(vF,hF)
    A_val=np.where(r['alive'],np.maximum(s1-K,0)*s2,0.0)*c['barrels']
    phys=c['barrels']*(s1*s2-c['S1_0']*c['S2_0'])
    V_A0=17_448.76
    dV_A=A_val-V_A0*c['barrels']
    E_A=dV_A-phys; E_B=(vW+vF)-phys
    q=np.exp(c['rho']*c['diffvol']*c['sig2']*T) if quanto_adj else 1.0
    dH_A=c['barrels']*(s1*s2-c['S1_0']*c['S2_0']*q)
    ineff_A=np.abs(dV_A-signed_lower_of(dV_A,dH_A))
    ineff_B=np.abs(vF-cfhrF)
    return E_A,E_B,ineff_A,ineff_B

N=200_000
EA,EB,IA,IB=paired(N)
print(f"n = {N:,} paired paths (shared bank)")
print(f"  sigma_econ A = {EA.std(ddof=1)/1e9:8.2f} bn")
print(f"  sigma_econ B = {EB.std(ddof=1)/1e9:8.2f} bn")
# paired test on the PER-PATH difference of squared residuals (variance equality)
d = EA**2 - EB**2
se = d.std(ddof=1)/np.sqrt(N)
print(f"\n  paired test of equal variance: mean(E_A^2 - E_B^2) = {d.mean():.4e}")
print(f"     se = {se:.4e}   t = {d.mean()/se:8.3f}")
# also paired difference in |E|
d2=np.abs(EA)-np.abs(EB); se2=d2.std(ddof=1)/np.sqrt(N)
print(f"  paired mean(|E_A| - |E_B|) = {d2.mean()/1e9:8.4f} bn  se {se2/1e9:.4f} bn  t = {d2.mean()/se2:7.3f}")
# quanto adjustment on / off
_,_,IAq,_ = paired(N, quanto_adj=True)
_,_,IA0,_ = paired(N, quanto_adj=False)
print(f"\n  A mean|ineff| with quanto adj q : {IAq.mean()/1e9:8.3f} bn")
print(f"  A mean|ineff| with q = 1 (corrected drift): {IA0.mean()/1e9:8.3f} bn")
print(f"  change: {100*(IA0.mean()/IAq.mean()-1):+.3f}%")
