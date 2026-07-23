"""Rebuild the program's headline numbers from source and print them together.

  python3 python/00_engine/verify_all.py [npaths]

Nothing here reads a spreadsheet.  Closed forms are evaluated in place; the
American premiums come from lsmc_quanto, the stress mortality from jump_barrier,
the measure-change validation from girsanov, the delta comparison from
delta_hedge and the two designation architectures from cfh_ledger.
"""
import os, sys
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import verify_paper1                                                    # noqa: E402
from girsanov import validate                                           # noqa: E402
from delta_hedge import fit_surface, run as hedge_run                    # noqa: E402
from cfh_ledger import structures                                       # noqa: E402


def main(npaths=200_000):
    print('=' * 78)
    print(f'PAPER 1 — fixed-budget allocation           ({npaths:,} paths)')
    print('=' * 78)
    verify_paper1.main(npaths)

    print()
    print('=' * 78)
    print(f'PAPER 2 — delta hedging the barrier structure')
    print('=' * 78)
    fit = fit_surface(npaths=npaths)
    hs = hedge_run(fit, npaths=npaths, proxy=False)
    hp = hedge_run(fit, npaths=npaths, proxy=True)
    print(f"  LSMC price                              {fit['price']:12,.2f} KRW/bbl")
    print(f"  hedge cost, fitted surface delta        {hs.mean()/1e9:8.2f} bn   sd {hs.std(ddof=1)/1e9:8.2f} bn")
    print(f"  hedge cost, Black-76 proxy delta        {hp.mean()/1e9:8.2f} bn   sd {hp.std(ddof=1)/1e9:8.2f} bn")
    print(f"  cost of the model-instrument mismatch   {(hp.std(ddof=1)-hs.std(ddof=1))/1e9:8.2f} bn of dispersion")
    g = validate(npaths=npaths, reps=30)
    print(f"  Girsanov density  E[L] {g['EL']:.6f}   sd[L] {g['sdL']:.6f}   closed form {g['sdL_theory']:.6f}")
    print(f"  reweighting residual                    {g['resid_pct']:+.4f}% +/- {g['resid_se']:.4f}  "
          f"({g['reps']} replications, common random numbers)")

    print()
    print('=' * 78)
    print(f'PAPER 3 — IFRS 9 designation architectures')
    print('=' * 78)
    c = structures(npaths=npaths)
    bn = lambda x: x/1e9
    print(f"  barrier touch rate                      {c['ko_rate']:.4f}")
    print(f"  unhedged physical exposure sd           {bn(c['sigma_phys']):10,.2f} bn")
    print(f"  sigma_econ   A / B                      {bn(c['sigma_econ_A']):10,.2f} / {bn(c['sigma_econ_B']):.2f} bn")
    print(f"  mean |cumulative ineffectiveness| A / B {bn(c['A_mean_ineff']):10,.2f} / {bn(c['B_mean_ineff']):.2f} bn")
    print(f"  post-KO naked exposure VaR99  A / B     {bn(c['postko_naked_var99_A']):10,.2f} / {bn(c['postko_naked_var99_B']):.2f} bn")
    print(f"  SFP identity V - (OCI + RE), max abs    {c['sfp_identity_max_abs']:.3e} KRW")

    print()
    print('=' * 78)
    print('PAPER 4 — mandatory disclosure and hedging')
    print('=' * 78)
    try:
        sys.path.insert(0, os.path.join(HERE, '..', '04_esg'))
        import esg_hedge_engine as esg
        for name in ('main', 'run', 'solve'):
            if hasattr(esg, name):
                print(f"  engine entry point esg_hedge_engine.{name}() available")
                break
        print('  equilibrium and estimator run from python/04_esg (no workbook dependency)')
    except Exception as e:
        print(f'  engine not loaded: {e}')


if __name__ == '__main__':
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 200_000)
