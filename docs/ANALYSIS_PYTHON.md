# ANALYSIS_PYTHON.md

Every Python script run in this analysis pass, in execution order, with
the exact console output captured. This documents the numbers that feed
the three rewritten papers (`Delta_Simulation/Park_quanto.tex`,
`Hedge_Simulation/Park_QuantoCFH.tex`, `Ratio_Optimization/Park_hedge_optimization.tex`)
and the four "v3/v4" `.bas` files placed alongside each paper (none of
which touch `Modeling/Bas/`, which is left as the historical record of
the originally reviewed code).

## Why this pass exists

Direct inspection of the VBA actually embedded in the three workbooks
(via `oletools.olevba`, not inference) surfaced three defects that no
amount of re-reading the `.tex` files alone would have caught, because
the papers' prose describes procedures the code does not implement:

1. **Shapley-premium invariance error.** The American KO leg's cost
   coefficient (`LSMC!J9`) is a Shapley-value cost *attribution* of a
   jointly-priced quanto structure, reused as if it were the market price
   of a standalone WTI KO call. It is not — the mixed vanilla/KO program
   needs an independently-priced KO leg.
2. **The "asymmetric up/down EM decomposition"** described in
   `Park_quanto.tex` §4.1 does not exist in `Calibration_revised.bas`.
   The only jump-fitting routine there (`DecomposeJumpDiffusion`) pools
   all jump-flagged days — regardless of sign — into one regime. No
   up-regime, no down-regime, no moment-matching step is coded anywhere.
3. **Structure B in the CFH accounting engine is not forwards.** Despite
   `Park_QuantoCFH.tex` §4.2 stating "Structure B uses costless forwards"
   with linear forward-mark equations, the actual VBA (`CFH_Accounting_revised.bas`,
   and the "v3.1" refactor found live in `Hedge_Simulation/Hedge 복사본.xlsm`)
   prices Structure B's WTI leg as a standalone knock-out American call
   (`WTI_KO_Call_BetaFV` against a `Beta_Mat_WTI` continuation surface,
   sharing the same `koStep = AP_KO(sim)` as the joint quanto) and its FX
   leg as a Garman-Kohlhagen call option (`GKCall`) — verified by grep,
   not inferred.

All three are fixed here: fresh Python re-derivations for defects 1-2,
and a from-scratch forward re-simulation for defect 3, each cross-checked
against the corresponding new `.bas` module before the papers were
rewritten around the corrected numbers.

---

## 1. VBA extraction (verifying what the workbooks actually run)

```bash
python3 -m pip install --user oletools   # already installed
```

```python
# extract_vba.py
import sys, os
from oletools.olevba import VBA_Parser

path, outdir = sys.argv[1], sys.argv[2]
os.makedirs(outdir, exist_ok=True)
vp = VBA_Parser(path)
for (filename, stream_path, vba_filename, vba_code) in vp.extract_macros():
    if vba_filename.endswith('.bas'):
        with open(os.path.join(outdir, vba_filename), 'w', encoding='utf-8') as f:
            f.write(vba_code)
        print('wrote', vba_filename, len(vba_code), 'bytes')
vp.close()
```

Run against all three workbooks:

```
Delta_Simulation/Hedge.xlsm            -> Calibration_revised.bas (15,510B), CFH_Accounting_revised.bas (30,949B),
                                           PaperVerification_revised.bas (98,578B), DeltaHedging_revised.bas (73,113B)
Hedge_Simulation/Hedge 복사본.xlsm      -> same four modules, EXCEPT CFH_Accounting_revised.bas is 46,003B
                                           (a newer "PRODUCTION v3.1" refactor, header comment:
                                           "Refactored to resolve structural timeline defects and matrix symmetry")
Ratio_Optimization/Ratio_Optimization.xlsm -> same four + Ratio_Optimization.bas (28,681B, the "v2 corrected model")
```

`diff` confirmed `Calibration_revised.bas`, `PaperVerification_revised.bas`,
`DeltaHedging_revised.bas` are byte-identical across all three workbooks
(same shared library). `CFH_Accounting_revised.bas` differs: the
Hedge_Simulation workbook (the accounting paper's own workbook) carries
the newer v3.1 refactor. Grepping v3.1 for Structure B's leg formulas:

```
766: ws.Range("A1").Value = "Structure B Layer 1 -- WTI KO Option CFH (FX-fixed at S2_0)"
780: ws2.Range("A1").Value = "Structure B Layer 2 -- FX GK Option CFH (FIXED Notional Mismatch)"
352: wtiFinalBeta = WTI_KO_Call_BetaFV(S1, S1_0, Beta_Mat_WTI_Steps, barrels, koStep)
346: cumHIx = GKCall(s2, G0_FX, gVol2, tau_FX, r_US, r_KRW, N_FX)
```

Confirms defect 3 persists in the v3.1 refactor — the "refactor" fixed
something else ("timeline defects and matrix symmetry"), not the
option-vs-forward mismatch.

---

## 2. Genuine asymmetric up/down jump EM (fixes defect 2)

```python
# asymmetric_em.py
import numpy as np
import openpyxl
import json

wb = openpyxl.load_workbook(
    '/Users/elijahjasper/Desktop/Modeling/Delta_Simulation/Hedge.xlsm',
    read_only=True, data_only=True)
ws = wb['Raw_Timeseries']
ret = []
for row in ws.iter_rows(min_row=3, max_col=6, values_only=True):
    if row[5] is None:
        break
    ret.append(row[5])
ret = np.array(ret)
n = len(ret)

# Stage 1: identical robust k-sigma classification to the production
# DecomposeJumpDiffusion (MAD*1.4826 seed, iterative k=3 reclassification)
def mad_scale(x):
    med = np.median(x)
    return np.median(np.abs(x - med))

sigma = mad_scale(ret) * 1.4826
is_jump = np.zeros(n, dtype=bool)
prev_count = -1
for _ in range(10):
    normal = ret[~is_jump]
    mean_normal = normal.mean() if len(normal) else 0.0
    is_jump = np.abs(ret - mean_normal) > 3.0 * sigma
    cnt = is_jump.sum()
    if cnt == prev_count:
        break
    prev_count = cnt
    sigma = ret[~is_jump].std(ddof=1)

normal, jump = ret[~is_jump], ret[is_jump]
vol1_prod = normal.std(ddof=1) * np.sqrt(252)
lambda_prod = is_jump.sum() / n * 252
theta_prod, delta_prod = jump.mean(), jump.std(ddof=1)

# Stage 2 (NEW): split the classified jump days by SIGN -- this split does
# not exist anywhere in Calibration_revised.bas
up_jumps, dn_jumps = jump[jump > 0], jump[jump < 0]
n_up, n_dn = len(up_jumps), len(dn_jumps)
lambda_up, lambda_dn = n_up / n * 252, n_dn / n * 252
theta_up, delta_up = up_jumps.mean(), up_jumps.std(ddof=1)
theta_dn, delta_dn = dn_jumps.mean(), dn_jumps.std(ddof=1)

# Stage 4: moment-match to one symmetrized regime (Park_quanto.tex Sec 4.1's
# own stated formulas -- now actually evaluated on a genuine asymmetric fit)
lam_sym = lambda_up + lambda_dn
th_sym = (lambda_up * theta_up + lambda_dn * theta_dn) / lam_sym
EY2 = (lambda_up * (theta_up**2 + delta_up**2) + lambda_dn * (theta_dn**2 + delta_dn**2)) / lam_sym
dl_sym = np.sqrt(max(EY2 - th_sym**2, 0.0))
```

**Output:**

```
n WTI daily log returns = 1299

--- Stage 1 (reproduces production DecomposeJumpDiffusion, pooled) ---
jump days = 35 / 1299
vol1 (diffusive) = 0.324313  (production LSMC!B10 = 0.324185)
lambda (pooled)  = 6.789838  (production LSMC!B1  = 6.784615)
theta_J (pooled) = -0.029895  (production LSMC!B2  = -0.029895)
delta_J (pooled) = 0.084429  (production LSMC!B3  = 0.084429)

--- Stage 2 (NEW: asymmetric up/down split of the jump set) ---
up-jump days   = 12  (mean=0.080447, std=0.016751)
down-jump days = 23  (mean=-0.087465, std=0.027750)
lambda_up = 2.327945 /yr   theta_up = 0.080447   delta_up = 0.016751
lambda_dn = 4.461894 /yr   theta_dn = -0.087465   delta_dn = 0.027750

--- Stage 4 (moment-matched symmetrization of the GENUINE asymmetric fit) ---
lambda (symmetrized)  = 6.789838   vs production 6.789838
theta_J (symmetrized) = -0.029895   vs production -0.029895
delta_J (symmetrized) = 0.083395   vs production 0.084429
```

**Reading:** down-jumps occur 1.9x more often, are 1.9x larger in
magnitude, and 66% more volatile than up-jumps — the left-skew the
paper's prose claims is real, and is now actually computed rather than
asserted. The moment-matched symmetrization reproduces the production
`LSMC!B1:B3` values almost exactly (λ and θ_J match to the displayed
digit; δ_J differs by ~1.2% relative, a small-sample Bessel-correction
artifact of recombining two sub-variances vs. one direct pooled STDEV).
The production pooled parameters are therefore validated as a faithful
symmetrization of a genuinely asymmetric process — they were just never
derived that way in code before now.

---

## 3. Standalone WTI KO American call + asymmetric stress p_KO (fixes defects 1 and, partially, 2)

```python
# standalone_wti_ko.py  (uses asymmetric_jump_fit.json from step 2)
import numpy as np, json, warnings
warnings.filterwarnings('ignore', category=RuntimeWarning)

S1_0, K, U, L, r_US, T_oil = 78.94, 78.94, 120.0, 50.0, 0.04, 0.833
STEPS_PER_YEAR = 52
lam_up, th_up, dl_up = FIT['lambda_up'], FIT['theta_up'], FIT['delta_up']
lam_dn, th_dn, dl_dn = FIT['lambda_dn'], FIT['theta_dn'], FIT['delta_dn']
lam_sym, th_sym, dl_sym = FIT['lambda_sym'], FIT['theta_sym'], FIT['delta_sym']
diff_vol = FIT['vol1']

def simulate_paths(s0, T, n_paths, steps_per_year, lam_up, th_up, dl_up,
                    lam_dn, th_dn, dl_dn, drift_mode, rng):
    """Two independent compound-Poisson regimes (up/down) + diffusion."""
    n_steps = max(1, round(T * steps_per_year)); dt = T / n_steps
    kappa_up = np.exp(th_up + 0.5*dl_up**2) - 1
    kappa_dn = np.exp(th_dn + 0.5*dl_dn**2) - 1
    mu = 0.139 if drift_mode == 'physical' else r_US   # LSMC!B4 physical drift
    drift = (mu - lam_up*kappa_up - lam_dn*kappa_dn - 0.5*diff_vol**2) * dt
    vol_dt = diff_vol * np.sqrt(dt)
    x = np.full(n_paths, np.log(s0)); alive = np.ones(n_paths, bool)
    touched = np.zeros(n_paths, bool)
    for t in range(n_steps):
        z = rng.standard_normal(n_paths)
        n_up = rng.poisson(lam_up*dt, n_paths); n_dn = rng.poisson(lam_dn*dt, n_paths)
        jump = np.zeros(n_paths)
        for j in range(1, n_up.max()+1):
            idx = n_up >= j; jump[idx] += th_up + dl_up*rng.standard_normal(idx.sum())
        for j in range(1, n_dn.max()+1):
            idx = n_dn >= j; jump[idx] += th_dn + dl_dn*rng.standard_normal(idx.sum())
        x = np.where(alive, x + drift + vol_dt*z + jump, x)
        s = np.exp(x)
        hit = alive & ((s >= U) | (s <= L))
        touched |= hit; alive &= ~hit
    return touched

# (A) stress-conditional p_KO, genuine asymmetric regimes, S0=113
rng = np.random.default_rng(20260706)
touched_asym = simulate_paths(113.0, T_oil, 500_000, STEPS_PER_YEAR,
                               lam_up, th_up, dl_up, lam_dn, th_dn, dl_dn, 'physical', rng)
pko_asym = touched_asym.mean()

# re-derived pooled/symmetric case for exact comparability (splitting the
# symmetrized regime evenly with IDENTICAL up/down params reproduces a
# single pooled regime exactly)
rng2 = np.random.default_rng(20260706)
touched_sym = simulate_paths(113.0, T_oil, 500_000, STEPS_PER_YEAR,
                              lam_sym/2, th_sym, dl_sym, lam_sym/2, th_sym, dl_sym, 'physical', rng2)
pko_sym = touched_sym.mean()

# (B) standalone WTI KO American call, LSMC, Q-measure, no FX/Shapley
def price_standalone_wti_ko(s0, K, T, n_paths, steps_per_year, rng):
    n_steps = max(1, round(T*steps_per_year)); dt = T/n_steps
    kappa_up = np.exp(th_up + 0.5*dl_up**2) - 1
    kappa_dn = np.exp(th_dn + 0.5*dl_dn**2) - 1
    drift = (r_US - lam_up*kappa_up - lam_dn*kappa_dn - 0.5*diff_vol**2) * dt
    vol_dt = diff_vol * np.sqrt(dt)
    x = np.full(n_paths, np.log(s0)); alive = np.ones(n_paths, bool)
    S_path = np.zeros((n_steps+1, n_paths)); S_path[0] = s0
    alive_path = np.ones((n_steps+1, n_paths), bool)
    for t in range(1, n_steps+1):
        z = rng.standard_normal(n_paths)
        n_up = rng.poisson(lam_up*dt, n_paths); n_dn = rng.poisson(lam_dn*dt, n_paths)
        jump = np.zeros(n_paths)
        for j in range(1, n_up.max()+1):
            idx = n_up >= j; jump[idx] += th_up + dl_up*rng.standard_normal(idx.sum())
        for j in range(1, n_dn.max()+1):
            idx = n_dn >= j; jump[idx] += th_dn + dl_dn*rng.standard_normal(idx.sum())
        x = np.where(alive, x + drift + vol_dt*z + jump, x)
        s = np.exp(x); hit = alive & ((s >= U) | (s <= L)); alive &= ~hit
        S_path[t] = s; alive_path[t] = alive
    disc = np.exp(-r_US*dt)
    cashflow = np.where(alive_path[-1], np.maximum(S_path[-1]-K, 0.0), 0.0)
    for t in range(n_steps-1, 0, -1):
        cashflow *= disc
        itm = alive_path[t] & (S_path[t] > K)
        if itm.sum() > 50:
            X = S_path[t, itm]; Xn = X/K - 1.0; Y = cashflow[itm]
            basis = np.vstack([np.ones_like(Xn), Xn, Xn**2]).T
            gram = basis.T@basis + 1e-6*np.trace(basis.T@basis)/3*np.eye(3)
            coef = np.nan_to_num(np.linalg.solve(gram, basis.T@Y))
            cont = np.nan_to_num(basis@coef, nan=-1.0, posinf=1e18, neginf=-1.0)
            exercise = np.maximum(X-K, 0.0)
            do_ex = exercise > cont
            idx_itm = np.where(itm)[0]
            cashflow[idx_itm[do_ex]] = exercise[do_ex]
    cashflow *= disc
    return cashflow.mean(), cashflow.std(ddof=1)/np.sqrt(n_paths), 1-alive_path[-1].mean()

rng3 = np.random.default_rng(20260706)
p_standalone, se_p, touch_rate = price_standalone_wti_ko(S1_0, K, T_oil, 150_000, STEPS_PER_YEAR, rng3)
```

**Output:**

```
Asymmetric two-regime p_KO(stress, S0=113) = 0.8405 +- 0.0010  (500000 paths, weekly)
Pooled/symmetrized single-regime p_KO(stress, S0=113) = 0.8407 +- 0.0010  (re-derived here for exact comparability)
Previously reported (bas Monte Carlo, same style) p_KO = 0.8413 +- 0.0010

Standalone WTI KO American call (LSMC, independent of FX/Shapley):
  price = 11.3239 +- 0.0654 USD/bbl   (KO/exercise rate = 0.3886)
  vs. Shapley-attributed joint-quanto WTI share: 15,093.75 KRW/bbl
  standalone price in KRW/bbl at spot (x1540.64): 17,446.12
```

**Reading, part (A):** the asymmetric two-regime simulation (84.05%) and
the pooled single-regime simulation (84.07%), computed on the *same* MC
code with the *same* seed, differ by only 0.02 percentage points — well
inside the ±0.10pp sampling band at 500,000 paths. **This refutes the
directional hypothesis raised in the prior discussion** (that fixing the
symmetrization would materially move p_KO): over a 0.833-year horizon
with 32.4% diffusive volatility and only 6.79 jumps/year total, the
continuous diffusion term dominates the barrier-touch probability from a
stress spot only 6.2% below the upper barrier; the jump distribution's
sign-asymmetry is a second-order effect on this specific question. The
number is essentially unchanged (84.05% vs. the previously-reported
84.13%) — a robustness finding worth reporting precisely because it
contradicts a plausible-sounding prior guess.

**Reading, part (B):** the independent, standalone WTI KO American call
prices at 17,446 KRW/bbl — **15.6% above** the Shapley-attributed joint
premium share (15,094 KRW/bbl). Reusing a cooperative-game marginal
contribution as a stand-alone market price understated the true KO
premium; every downstream mixed-program threshold that depends on it
must be recomputed.

---

## 4. Mixed vanilla/KO program re-solved with the corrected P_K

```python
import numpy as np, json, math
from scipy.optimize import minimize, brentq

Qoil, Qusd = 2_000_000, 157_880_000.0
WACC, T_oil, T_fx = 0.07, 0.833, 0.5
K1E = 12.652449707035759*Qoil*1540.64*(1+WACC*T_oil)     # vanilla Black-76, UNCHANGED
K2E = 13835108653.715132                                  # vanilla GK FX, UNCHANGED
M1, M2, B = 105586000000.0, 1477756799.9999843, 45e9
s1a, s2, rho = 0.3241851967428078, 0.09257650229436072, 0.08763086561435483
K1A = 17446.12 * Qoil * np.exp(WACC*T_oil)                # CORRECTED standalone KO price

def gm(w1, w2):
    h1, h2 = 1-w1, 1-w2
    return math.sqrt(h1*h1*s1a*s1a + h2*h2*s2*s2 + 2*h1*h2*s1a*s2*rho)

def cost_KO(w1, w2, p): return K1A*w1 + K2E*w2 + (1-w1*(1-p))*M1 + (1-w2)*M2
def cost_V(w1, w2):     return K1E*w1 + K2E*w2 + (1-w1)*M1     + (1-w2)*M2

pbar  = (K1E - K1A) / M1
pstar = (B - (K1A + M2)) / M1
w1f, w2f = 0.9452021998978689, 0.05479779791049507   # unconstrained-floor allocation (unaffected by K1A)
pfloor = brentq(lambda p: cost_KO(w1f, w2f, p) - B, 0, 0.3)
```

**Output:**

```
K1A_new (KO leg premium coeff) = 36,987,301,705
K1E (vanilla WTI premium coeff, unchanged) = 41,258,998,746
p_bar (new, corrected P_K) = 0.040457  = 4.0457%     [was 0.08769 = 8.769% with the Shapley P_K]
p_floor (new) = 0.021029                              [was 0.068263]
p*_KO (infeasibility threshold, corrected P_K) = 0.061892 = 6.1892%   [was 0.10913 = 10.913%]
all-vanilla optimum: w=(0.970486, 0.029514) sigma=0.0911822 cost=45.0000bn   [unchanged: doesn't depend on K1A]
```

At the measured stress mortality (p_KO = 0.8405, from step 3):
cheapest attainable pure-KO stress-adjusted cost = **127.2bn KRW**
(vs. 45bn budget), adopted pure-KO book's ledger = **136.1bn KRW**.
Both thresholds moved down (the corrected, more expensive KO leg is
worth less mortality risk than previously computed), so the mixed
program switches to all-vanilla *even sooner* than the Shapley-based
calculation suggested — the qualitative conclusion is unchanged and, if
anything, strengthened.

---

## 5. CFH Structure B re-simulated as genuine linear forwards (fixes defect 3)

```python
# cfh_structureB_forwards.py
import numpy as np, json, warnings
warnings.filterwarnings('ignore', category=RuntimeWarning)

S1_0, S2_0 = 78.94, 1540.64
r_US, r_KRW = 0.04, 0.035
mu1_P, mu2_P = 0.139, 0.0239          # LSMC!B4, B5 (physical drifts)
rho, sig2 = 0.08763086561435483, 0.09257650229436072
U, L = 120.0, 50.0
barrels = 2_000_000.0
T_WTI, T_FX = 0.833, 0.5
STEPS_PER_YEAR = 52
n_paths = 40_000

# ... simulate correlated (S1, S2) under P-measure, S1 via the corrected
# asymmetric two-regime jump-diffusion (non-stopping past any barrier
# touch), S2 via correlated GBM ...

# Structure B leg marks -- the paper's OWN Section 4.2 closed-form equations,
# not the option-based formulas the actual bas code computes:
F0 = S1_0 * np.exp(r_US * T_WTI)
F1 = S1_path * np.exp(r_US * np.maximum(T_WTI - tgrid, 0.0))[:, None]
V_B_WTI = (F1 - F0) * barrels * np.exp(-r_US*np.maximum(T_WTI-tgrid,0.0))[:,None] * S2_path
HD_WTI  = V_B_WTI.copy()                      # matched notional -> ~0 ineffectiveness

G0 = S2_0 * np.exp((r_KRW - r_US) * T_FX)
N_FX = barrels * F0
G1 = S2_path[:fx_idx_max+1] * np.exp((r_KRW-r_US)*np.maximum(T_FX-tgrid_fx,0.0))[:,None]
V_B_FX = N_FX * (G1 - G0) * np.exp(-r_KRW*np.maximum(T_FX-tgrid_fx,0.0))[:,None]     # fixed notional
HD_FX  = barrels * F1[:fx_idx_max+1] * (G1 - G0) * np.exp(-r_KRW*np.maximum(T_FX-tgrid_fx,0.0))[:,None]  # floating

CFHR_FX  = np.sign(V_B_FX) * np.minimum(np.abs(V_B_FX), np.abs(HD_FX))   # IFRS 9 lower-of test
Ineff_FX = V_B_FX - CFHR_FX
```

**Output:**

```
KO rate (barrier touch, non-stopping, n=40000) = 0.4040

--- Structure B (genuine forwards) terminal statistics, n=40000 ---
Derivative carrying-amount std (gross, gross of netting): 73.3954 bn KRW
End-of-life CFHR std: 107.3342 bn KRW
Mean |cumulative ineffectiveness| (FX leg, terminal): 1.2024 bn KRW
5th/95th pctile of cumulative ineffectiveness (fan): [-3.5116, 4.4564] bn KRW
Post-KO P&L (KO paths only, n=16159): mean=11.6035 bn, std=80.7827 bn KRW
```

**Reading:** the directional claims survive — Structure A carries zero
ineffectiveness by construction (unchanged, not re-simulated here) while
Structure B's genuine-forward ineffectiveness is nonzero and fans out
over time (H2); post-KO P&L injects large two-sided volatility once
Structure B's surviving forwards are marked FVTPL (H3, mean +11.6bn,
std 80.8bn). What changes is *magnitude*: a linear forward cannot amplify
the notional mismatch the way the (mislabeled) option-based leg did — the
mean absolute ineffectiveness (1.2bn) is much smaller than what the
option-based code was generating. Correcting Structure B to what the
paper actually describes weakens the paper's headline numbers but not
its qualitative claim.

---

## Summary of all corrected values used in the rewritten papers

| Quantity | Old (Shapley/pooled/option-based) | New (corrected) |
|---|---|---|
| WTI KO leg premium (KRW/bbl) | 15,093.75 (Shapley split) | 17,446.12 (standalone LSMC) |
| Stress-conditional p_KO | 0.8413 (pooled single-regime) | 0.8405 (genuine asymmetric two-regime) — statistically indistinguishable |
| Mixed-program switch point p̄ | 8.769% | 4.046% |
| Mixed-program infeasibility p* | 10.913% | 6.189% |
| Mixed-program floor-regime end | 6.826% | 2.103% |
| CFH Structure B derivative carry std | (option-based, not independently re-derived here) | 73.40bn KRW (genuine forwards) |
| CFH Structure B mean |ineffectiveness| | (option-based) | 1.20bn KRW (genuine forwards) |
| CFH Structure B post-KO P&L mean/std | (option-based) | +11.60bn / 80.78bn KRW (genuine forwards) |
| Asymmetric jump regimes (new) | not computed anywhere | λ_up=2.328/yr θ_up=+8.04% δ_up=1.68% ; λ_dn=4.462/yr θ_dn=-8.75% δ_dn=2.78% |

All scripts above were executed in this session; JSON artifacts
(`asymmetric_jump_fit.json`, `standalone_and_pko_results.json`,
`cfh_structureB_v2.json`, `mixed_program_v2_full.json`) back every number
quoted in the three rewritten papers.

---

## 6. Representative-path "zombie FVTPL" figure and a vanilla-quanto benchmark

Two follow-up questions on the CFH paper (`Hedge_Simulation/Park_CFH.tex`):
does a single representative knock-out path exist anywhere to illustrate
Structure A's clean extinguishment against Structure B2's surviving,
FVTPL-bound FX leg; and does the paper compare either structure against a
**vanilla (barrier-free) quanto** — i.e. does it ever test whether the
knock-out feature, not the single-vs-split designation choice, is what
drives Structure A's elevated ineffectiveness?

Neither existed. Checked directly:

```bash
grep -n "detS_A\|detK_A" CFH_Accounting_revised.bas   # per-path detail arrays
grep -n "WriteLedgerA\|WriteLedgerB" CFH_Accounting_revised.bas  # what actually gets written to sheets
```

`detS_A`/`detK_A` (a "surviving" and a "knocked-out" representative path,
by name) are populated during the simulation loop but **never passed to
any `Write*` routine** — no sheet, anywhere, persists single-path detail;
every `CFH_*_Ledger` sheet is a cross-path mean/std trajectory. And the
paper's only comparison is Structure A (single quanto KO) vs. Structure B
(WTI KO leg + FX leg) — no vanilla benchmark is ever priced. Both had to
be built from scratch.

### 6.1 One LSMC-style regression fit, evaluated three ways

```python
# zombie_and_vanilla.py
import numpy as np

# calibration: same as the rest of the trilogy (S1_0=78.94, S2_0=1540.64,
# K=78.94 ATM, U/L=120/50, r_US=0.04, r_KRW=0.035, lambda=6.7846,
# theta_J=-0.0299, delta_J=0.0844, diff_vol1=0.3242, sig2=0.0926,
# rho=0.0876, barrels=2,000,000, T=0.833, weekly steps)

def fit_backward(payoff_fn, barrier=True, floor_zero=True):
    """5-term poly regression (beta0..beta5 on v1,v2,v1^2,v2^2,v1*v2),
    fit once per backward step on the full N=20,000-path bank; `barrier`
    kills paths at the shared KO event, `floor_zero` applies max(fitted,
    intrinsic,0) for a call vs. leaving the raw fitted value for a forward."""
    ...

VA, aliveA = fit_backward(payoff_call,     barrier=True,  floor_zero=True)   # Structure A (KO quanto)
VH, _      = fit_backward(payoff_forward,  barrier=False, floor_zero=False)  # HStar hypothetical (barrier-free forward)
VV, _      = fit_backward(payoff_call,     barrier=False, floor_zero=True)   # vanilla quanto (no barrier)
```

**Output:**

```
KO rate (n=20000): 0.3905
Structure A day-0 price:      42.39 bn KRW
Vanilla quanto day-0 price:   49.12 bn KRW
HStar day-0 (fwd hypothetical): 8.89 bn KRW

Mean |terminal cumulative ineffectiveness|:
  Structure A (KO quanto):      4.74 bn KRW   (production engine reports 23.36bn -- see caveat below)
  Vanilla quanto (no barrier):  6.55 bn KRW
  Structure B (from paper):     6.40 bn KRW

Full-horizon economic residual std:
  Structure A (KO):     94.25 bn KRW  (production engine reports 104.52bn -- same order of magnitude)
  Vanilla quanto:       45.12 bn KRW
```

**Caveat, stated plainly:** this is a *simplified* re-derivation (weekly
monitoring vs. the production engine's finer grid, a plain 5-term OLS fit
vs. whatever numerical safeguards the production `Beta_Mat` carries) built
solely to compare the *relative* effect of removing the barrier under
identical machinery. Its absolute ineffectiveness/residual numbers are
not a replacement for the production Structure A/B figures already in the
paper (23.36bn / 6.40bn / 104.52bn / 103.18bn) and are not cited as if
they were; they are close enough in order of magnitude (94bn vs.\ 104bn
for the economic residual) to trust the *comparison* between the three
structures computed on the same engine.

**The finding does not match the naive prior.** Going in, the expectation
was "remove the barrier, Structure A's ineffectiveness drops toward zero
along with the barrier's extinguishment mismatch." That is not what
happens: the vanilla quanto's ineffectiveness (6.55bn) is *not* smaller
than the KO structure's (4.74bn) in this comparison -- if anything,
slightly larger. The mechanism: a barrier-free call stays alive to expiry
on every path, so its own convex fair value keeps diverging from the
*linear* forward hypothetical (HStar) for the full horizon; a KO
structure's paths are frequently cut short by the barrier before that
convexity mismatch has as much time to accumulate. What removing the
barrier *does* unambiguously buy is economic: full-horizon residual risk
drops from ~94-104bn to ~45bn, because a vanilla structure never loses
all coverage the way a knocked-out one does. **The two axes (combined vs.
split; barriered vs. not) are not as cleanly separable as hypothesized in
the prior discussion** -- removing the barrier buys real economic
protection but does not mechanically buy accounting effectiveness, because
effectiveness here is driven by option-convexity-vs-linear-forward
mismatch, a feature of *any* option-based single-instrument designation,
barrier or not.

### 6.2 Representative knock-out path (the "zombie FVTPL" figure)

From the same N=20,000 path bank used above, one path knocking out near
mid-horizon (t≈0.41yr) was selected for illustration. Structure A and
Structure B1 (WTI KO leg, FX fixed at S2(0), sharing the same barrier
event) both cliff to zero at the KO step. Structure B2 (Garman-Kohlhagen
FX call, no barrier) is priced along the same path via the closed-form GK
formula and continues marking to fair value after the shared
discontinuation flag trips -- its subsequent changes are what flow to
FVTPL rather than OCI.

```python
# make_zombie_fig.py -- two-panel figure: WTI vs. barriers (top),
# Structure A / B1 / B2 fair value with the KO event marked (bottom)
```

Saved as `Hedge_Simulation/figures/fig_cfh_zombie.pdf` (and copied to
`~/Desktop/CFH_LinkedIn_Images/fig_cfh_zombie.png` for the companion
LinkedIn post).

### 6.3 The missing fourth cell -- and a validation failure it exposes

The comparison above tests one corner of a 2x2 design (combined-vs-split
instrument x barriered-vs-vanilla payoff) by adding "combined + vanilla."
A direct question: does "split + vanilla" (Structure B1 with its
knock-out feature removed, FX still fixed at inception) complete the
picture, and does it change the conclusion?

First, the production ledger was read directly rather than assumed:

```python
import openpyxl
wb = openpyxl.load_workbook('Hedge 복사본.xlsm', read_only=True, data_only=True)
ws1 = wb['CFH_B_Ledger_WTI']   # B1
ws2 = wb['CFH_B_Ledger_FX']    # B2
# terminal (step 217) CumIneff column of each
```

**Output:** B1's terminal cumulative ineffectiveness is **-KRW 6.57
billion**; B2's is **+KRW 0.18 billion** (frozen at its own T_FX=0.5
maturity, matching Section 8.6's freeze mechanism). Signed sum ≈ -6.40bn,
matching the paper's headline |ineffectiveness| figure. **Essentially all
of Structure B's reported ineffectiveness comes from B1, not B2** -- a
fact the paper's headline "6.40bn" number had not previously decomposed.

This matters because B1 carries the *same knock-out barrier* as Structure
A, yet in production B1's ineffectiveness (6.57bn) is only about a
quarter of Structure A's (23.36bn). If the barrier alone drove Structure
A's elevated ineffectiveness (Section 9.1's original claim), a barriered
split leg should show a comparable order of magnitude to a barriered
combined instrument -- it does not, by a factor of 3.5x, exactly
Section 8.2's headline ratio.

The missing fourth cell was built on the same simplified engine as
Section 6.1, to see whether it reproduces this known production gap
before trusting its answer for the "split + vanilla" cell:

```python
# fourth_cell.py -- same path bank, same fit_backward() as zombie_and_vanilla.py
def payoff_call_fixedFX(s1, s2):     return barrels * np.maximum(s1 - K, 0.0) * S2_0
def payoff_forward_fixedFX(s1, s2):  return barrels * (s1 - K) * S2_0   # HD, no floor

V_splitKO      = fit_backward(payoff_call_fixedFX, barrier=True,  floor_zero=True)
V_splitVanilla = fit_backward(payoff_call_fixedFX, barrier=False, floor_zero=True)
HD_fixedFX     = fit_backward(payoff_forward_fixedFX, barrier=False, floor_zero=False)
```

**Output:**

```
split+KO (B1 analog):       mean |terminal ineff| = 4.71 bn KRW   (production B1: 6.57bn)
split+vanilla:               mean |terminal ineff| = 6.52 bn KRW

Full 2x2, same simplified engine, mean |terminal ineffectiveness|:
                  KO / barriered      vanilla / barrier-free
  combined (A)    4.74 bn             6.55 bn
  split (B1)      4.71 bn             6.52 bn
```

**This is a validation failure, reported as one.** The simplified engine
puts combined-KO and split-KO within 1% of each other (4.74 vs 4.71bn),
but production shows them 3.5x apart (23.36 vs 6.57bn) under the
identical barrier. The engine that produced Section 6.1's "the barrier,
not the combination, drives the gap" conclusion does not reproduce the
one production comparison available to check it against, on the very
axis (combined vs. split) that conclusion depends on. The claim was
withdrawn as unconfirmed pending a higher-fidelity re-run -- see Section
7 below, where that re-run was done and the withdrawal is superseded by
a validated, opposite finding.

---

## 7. Production-fidelity re-implementation (the fix, and the real answer)

Direct extraction of the exact VBA mechanics (`CFH_Accounting_revised.bas`
v3.1, the live engine) identified precisely what the simplified engine of
Sections 6.1-6.3 got wrong:

```vba
' the hypothetical derivative is a CLOSED FORM, never a regression fit:
Private Function HStar(S1, s2, tau, barrels, rho, sigma1, sigma2) As Double
    HStar = barrels * S1 * s2 * Exp(rho * sigma1 * sigma2 * tau)
End Function

' Structure A's lower-of test compares INTRINSIC value changes (not fair
' value changes) against a DYNAMICALLY SCALED HStar:
rawH = ((Beta_Mat(stp,1) + 2*Beta_Mat(stp,3)*v1n + Beta_Mat(stp,5)*v2n) / S1_0) / s2
h_t = clip(rawH, 0, 1)
aCFHR = SignedLowerOf(aIntr - aIntr_0, h_t * (HStar_t - gHStar0))

' Structure B1's lower-of test uses NO such scaling -- a flat 1:1 forward:
bCFHRw = SignedLowerOf(Intr_WTI - Intr_WTI_0, barrels * (S1 - S1_0) * S2_0)

' at the KO step, remaining unamortized cost-of-hedging is swept directly
' into that step's ineffectiveness -- not preserved for later recycling:
a_residual_TV = Max(TV0 - aCOHamort, 0)
aIneff = (aIntr - aIntr_0) - aCFHR - a_residual_TV
```

None of this was in the simplified engine: it used fair-value changes, no
h_t scaling (implicitly h_t=1 always), a regression-fit HD instead of the
closed form, weekly instead of daily (217-step) monitoring, and no KO-step
residual-COH sweep. A faithful re-implementation (`production_engine.py` +
`production_engine_stage2.py`), two-stage exactly as production (fit
`Beta_Mat`/`Beta_Mat_WTI` on one 20,000-path bank, walk a *separate* fresh
20,000-path accounting bank forward evaluating the frozen surface, daily
217 steps), validates directly against the real ledger:

```
                          this engine    production     agreement
Structure A ineffectiveness   24.78bn      23.36bn        within 6%
Structure B1 ineffectiveness  -7.61bn      -6.57bn        within 16%
Economic residual std (A)    107.29bn     104.52bn        within 3%
Economic residual std (B)    104.43bn     103.18bn        within 1%
A/B1 ratio                     3.26x         3.56x        matches the paper's own "3.5x" headline
KO rate                        42.2%         45.1%         within 3pp
```

This is not the same engine as Sections 6.1-6.3 reporting the same
numbers again -- it is what those sections' numbers should have been,
had the lower-of test been implemented correctly the first time.

### 7.1 The validated 2x2

With the engine now reproducing the one comparison available to check it
against, the two vanilla (barrier-free) cells were re-derived on the same
mechanics (`production_2x2.py`):

```
                 KO / barriered      vanilla / barrier-free
  combined (A)      24.78bn               55.77bn
  split (B1)          7.70bn                0.00bn
```

**Split + vanilla is mechanically exact zero.** B1's lower-of test compares
option intrinsic value ($\mathrm{barrels}\cdot\max(S_1-K,0)\cdot S_2(0)$)
against a fixed-notional WTI forward ($\mathrm{barrels}\cdot(S_1-S_1(0))\cdot
S_2(0)$). These are the *same function* whenever the option is in the
money (both reduce to $\mathrm{barrels}\cdot(S_1-K)\cdot S_2(0)$), and the
`SignedLowerOf` zero-crossing rule handles the out-of-the-money case as an
exact zero as well. A single-asset option's intrinsic value cannot help
but track a forward on the same asset -- with or without a barrier.

**Combined + vanilla is the worst cell, not the best.** Structure A's
hedged item is a *product* of two floating quantities
($\max(S_1-K,0)\cdot S_2$); no forward-style HD, scaled or not, is that
same function of $(S_1,S_2)$, so the mismatch never mechanically
vanishes the way B1's does, and without a barrier to end a path early it
compounds for the full 217-day horizon (55.77bn) instead of being cut
short partway through (24.78bn with the barrier -- still large, but
smaller, because roughly 42% of paths stop accumulating the mismatch
partway through in exchange for a single concentrated KO-day charge).

**The corrected conclusion, replacing Section 6.3's withdrawal:** it is
not the barrier that drives Structure A's ineffectiveness -- it is
combining two floating exposures into one multiplicative payoff, which
no linear (or delta-scaled-linear) hypothetical derivative can track
exactly. The barrier is, if anything, a partial *mitigant* for the
combined structure (it truncates the accumulation), while for the split
structure the barrier is very nearly the *entire* source of its
otherwise-negligible ineffectiveness. Economic risk moves the other way
entirely: combined+vanilla's residual std (42.3bn) is far *below*
combined+KO's (107.3bn), since removing the barrier removes the
"total loss of coverage on KO" mechanism -- so accounting ineffectiveness
and economic residual risk respond to the barrier in *opposite*
directions for the combined structure, a genuinely new finding this
re-derivation makes possible.

---

## 8. Hypothetical-derivative closed-form derivation (numerical check)

The paper's $H^\star(t)=\mathrm{barrels}\cdot S_1(t)S_2(t)\exp(\rho\sigma_1\sigma_2(T-t))$
formula was previously stated without derivation. The derivation (Itô on
the product process, compensated-jump MGF, KRW discounting) predicts
$\mathbb E^{\mathbb Q}[R_1R_2]=\exp[(r_{KRW}+\rho\sigma_1\sigma_2)\tau]$
where $R_1=S_1(T)/S_1(t)$, $R_2=S_2(T)/S_2(t)$. Checked directly against
$4\times10^6$ simulated path pairs at the production calibration:

```python
closed = np.exp((r_KRW + rho*sig1*sig2) * tau)      # 1.031842
mc     = (R1 * R2).mean()                             # 1.031760 (N=4e6)
# relative gap: 8.0e-5, within MC noise at this sample size
```

Confirms the closed form to better than 1bp; used as the basis for the
Appendix~13.1 derivation added to Park_CFH.tex.

## 9. FX maturity-freeze impact on the post-KO FVTPL headline (Section 8.6)

The production ledger reports (`CFH_Comparison` sheet, H4 block) "B mean
post-KO FVTPL P&L (KO paths)" = KRW 1,509,724,097.59 and its std =
KRW 71,946,732,843.60 -- confirmed as the exact source of the paper's
"KRW 1.51bn / 71.95bn" headline (a genuine per-path population statistic,
distinct from `CFH_B_Ledger_FX`'s `CumPostKO_FVTPL` column, which is a
cross-path running mean used only for the charted trajectory).

Hypothesis, read directly from the VBA (`WalkPath`): B2's freeze fires
purely on calendar time (`tau_FX<=0`), independent of `koStep`. Once
`bDiscont=True`, `bFVTPL` accrues `(fvWTI-fvWprev)+(cumHIx-fxXprev)`
every step; `fvWTI` is pinned at 0 post-KO, so the only driver is
`cumHIx`'s change -- which itself stops the moment `cumHIx` freezes.
For any path with $\tau_{KO}\ge T_{FX}=0.5$, B2 is *already* frozen
before KO happens, so `bFVTPL` should accrue exactly zero for the
entire post-KO period on that path.

```python
# fx_freeze_impact.py -- extends the validated production-fidelity
# engine (Section 7) with per-path bFVTPL tracking and a KO-timing tag
early = ever_ko & (t_ko < T_FX)   # KO before FX leg's own maturity
late  = ever_ko & (t_ko >= T_FX)  # KO at/after FX leg's own maturity
```

**Output** (n=20,000 accounting paths, daily monitoring, validated
mechanics):

```
KO paths: 8440/20000 (42.2%)
  KO before T_FX=0.5: 4312 (51.1% of KO paths)
  KO at/after T_FX=0.5: 4128 (48.9% of KO paths)

Post-KO FVTPL, ALL KO paths:     mean=0.26bn  std=5.32bn
Post-KO FVTPL, EARLY KO (t<0.5): mean=0.52bn  std=7.43bn
Post-KO FVTPL, LATE  KO (t>=0.5): mean=0.0000bn  std=0.0000bn   <-- exact zero
```

The late-KO subgroup's exact zero is confirmed on every one of the 4,128
paths in that bucket, not just on average -- a direct arithmetic
consequence of the freeze (a frozen fair value has no further changes to
contribute to FVTPL), not a modeling approximation. The near-even 51/49
population split means this isn't an edge case affecting a small tail;
essentially half the knocked-out population contributes nothing at all
to the "zombie FVTPL" mechanism the paper describes.

**Caveat, reported rather than resolved:** the split engine's absolute
FVTPL magnitudes (0.26bn/5.32bn pooled) are 6-13x smaller than the
production headline (1.51bn/71.95bn) -- a real gap, not yet isolated to
a specific cause (candidates: the accounting-bank seed/sample, or a
subtler mismatch in how `cumHIx`'s pre-freeze volatility compounds
relative to production's exact FX vol/rate handling). Given this,
Park_CFH.tex's new Section 8.6 paragraph relies only on the population
split fraction and the exact-zero result on the late-KO subgroup -- both
deterministic consequences of the freeze mechanism, independent of the
engine's calibration -- and explicitly does not claim the early-KO
subgroup's simulated mean/std reproduce the pooled production figures.

---

## 10. Optimal WTI–FX delta coupling (Paper 1, sec:cstar / sec:structdelta)

Implements FRAMEWORK_Optimal_FX_Delta.md against the American (early-
exercise) delta-hedging engine of `DeltaHedging_revised.bas`. Ground truth
read directly from `Delta_Simulation/Hedge.xlsm`: WACC=0.07, base premium
J2=17132.47, SIM_RUNS=10000, KO=0.2309, exercise=0.4474, and hedge cost =
`opt_profit - tot_profit` = -34.498bn - (-38.741bn) = **+4.243bn**, std
**54.41bn** (matching the paper's Table 6 exactly, confirming hedge cost =
`cumul_cost`).

### 10.1 The load-bearing structural fact (read from the VBA, not assumed)

`cumul_cost = cumul_FX - margin_WTI`. The WTI leg (`margin_WTI`) and every
stopping decision (KO via path bank, early exercise via `intrinsic >
V_base`) are computed with **zero dependence on `c` (`delta_ratio`)**; only
`delta_FX = c*delta_WTI` and the FX position/cost scale with `c`, linearly.
Therefore per-path hedge cost is EXACTLY affine in c:

    HC_p(c) = A_p + c*B_p,   A_p = HC_p(0),  B_p = HC_p(1)-HC_p(0)

so Var[HC(c)] is an exact upward parabola with closed-form minimizer
`c* = -Cov(A,B)/Var(B)`. This is an identity forced by the engine, not a fit.

### 10.2 Engine replica and validation (optimal_fx_delta_v2.py / _v3.py)

VBA-exact port: per-barrel `(S1-K)*S2` LSMC surface on 50k-100k Q-paths
(plain OLS, VBA basis order), bridge KO only on non-jump steps (VBA line
383), physical drifts (Drift1=0.139, Drift2=0.0239 from LSMC!B4/B5) for the
hedge path bank, WACC=0.07, American exercise `intrinsic > Vb`.

```
                          replica        production      status
hedge-cost std (c=1)      ~53.1bn        54.41bn         within 2.5%  [PASS]
total-P&L shape           unimodal       unimodal        skew 0.15 vs 0.131,
                                                          kurt 1.60 vs 1.333
early-exercise rate        0.541          0.4474         +9pp  [residual gap]
KO (American partition)    0.167          0.2309         (exercise pre-empts
                                                          more in replica)
G3 linearity |HC(0.5)-(A+0.5B)|/std       3e-16          machine precision [PASS]
```

The exercise-rate gap (systematic, ~20sigma, not RNG) shifts the mean hedge
cost (replica -9.9bn vs production +4.24bn) but leaves the std -- which is
what c* depends on -- within 2.5%. Per Framework Section I, absolute levels
are scoped as replica-derived; the affine identity and closed forms are
structure-driven and stated unconditionally.

### 10.3 Results (optimal_fx_delta_v3.py, 100k surface, 3 hedge seeds)

```
c* = -0.54   (stable across seeds, range [-0.557, -0.518])  <-- NEGATIVE
variance reduction:  std 53.1bn (c=1) -> 49.9bn (c*)  = 6.0% lower  <-- second-order
structural over-hedge (reconciled to FX-contract count):
    convention pos_FX (c=1) median 1109.6 contracts
    structural V/S2         median  241.8 contracts   -> 4.6x over-hedge
homogeneity diagnostic |dV/dS2 - V/S2|/(V/S2) median = 8.8%
```

**c\* is negative** because Cov(A,B)>0: the conventionally-signed FX leg
adds cost-variance that co-moves with the WTI leg. The swept range
[0.5,1.5] is the right arm of the parabola whose vertex is c*=-0.54
(fig_cstar_parabola.pdf).

**Structural FX delta (homogeneity theorem, exact):** the quanto payoff,
barrier, and jumps act on S1 only; S2 enters multiplicatively, so
V(t)=S2(t)*g(S1,path,t) is exactly linear in S2, giving
`delta_FX^struct = dV/dS2 = V/S2` -- the option's own USD value. The
equal-delta convention sizes off the WTI hedge notional instead, a 4.6x
over-hedge.

**Sign divergence (the interesting finding):** c*=-0.54 (short FX,
variance-minimizing) vs V/S2=+242 contracts (long FX, value-hedging)
disagree in SIGN -- they answer different questions, which the single
scalar c conflates.

### 10.4 Trilogy convergence

The dynamic optimum (FX coupling is a second-order lever, ~6% variance) and
Paper 3's static optimum (w2* ~= 3% FX coverage) reach the same verdict --
the FX leg deserves little direct hedging -- from independent engines and
objectives. Cross-links: the homogeneity/quanto argument here reuses the
CFH paper's Appendix 13.1 H* derivation on the S2 axis; the "over-sized FX
notional" theme echoes CFH's B2 fixed-notional mismatch.

---

## 11. 200k jump-adapted EXACT engine (Tier-S rigor upgrade, all three papers)

`upgraded_engine_200k.py` + `fixup_200k.py`. All samples n=200,000 per the
upgrade directive. Backups of all three papers' pdf/tex in
`Modeling/_backups_20260710/` before any text change.

### 11.1 Jump-adapted exact simulation (replaces Euler + step-level bridge)

Poisson jump TIMES drawn exactly within each step; between jumps the GBM
increment is exact (lognormal); the Brownian-bridge crossing probability is
applied on each jump-free SUB-interval and the barrier is checked directly
at each jump instant. This is exact first passage under the model: the
Euler discretization bias and the bridge's jump-blindness (the E_jump term
formalized in both papers' Limitations) are both eliminated by construction.
A production-style KO flag (endpoint + whole-step bridge only when nJ=0,
exactly the VBA rule) is computed ON THE SAME PATHS, so E_jump is measured
path-matched:

```
Q-bank (200k): exact KO 0.4391 vs production-style 0.4377  -> E_jump = +0.14pp
P-bank (200k): exact KO 0.4512 vs production-style 0.4494  -> E_jump = +0.18pp
```

The formal O(dt) claim now carries a measured value: the production scheme
understates continuous-time KO frequency by ~0.15-0.18pp at the daily grid.

### 11.2 Pricing at 200k, basis robustness, variance-reduction autopsy

```
LSMC price/unit (poly-6):   17,442.62 +- 44.16   (production workbook: 17,132.47, +1.81%)
LSMC price/unit (cubic-10): 17,589.53 +- 46.39   (+0.84% vs poly-6)
```
NOTE: a degree-2 tensor Laguerre basis spans the SAME 6-dim space as the
polynomial basis (identical fit by construction) -- the honest basis test
is a strictly larger space, hence cubic-10.

Variance reduction is structurally DEFEATED by the barrier:
```
control variate (Merton-series quanto-European closed form, analytic mean
  19,316.55 vs empirical 19,314.12 -- formula verified):  SE x1.01 only,
  beta_CV = 0.054: KO'd paths have zero option payoff exactly where the
  European CV payoff is largest (S1 through 120), destroying correlation.
antithetic pairing: SE x1.00 -- the KO payoff is non-monotone in the shocks.
```
Effective SE reduction must come from N itself (embarrassingly parallel),
which is why this pass runs 200k.

### 11.3 Convergence

```
in N:      10k 17,330+-194 | 25k 17,486+-125 | 50k 17,367+-88 | 100k 17,408+-62 | 200k 17,443+-44
in steps (50k): 109: 17,141 KO=0.4403 | 217: 17,574 KO=0.4398 | 434: 17,702 KO=0.4395
```
KO invariant across grids (exact first passage verified empirically); price
varies with step count only through EXERCISE-DATE density (a Bermudan
contract-definition effect, ~+0.7% per grid doubling), not discretization.

### 11.4 Esscher-consistent jump risk-premium sweep (fixes the old
intensity-only limitation)

Tilt h applied jointly: lambda_h = lambda*exp(h*theta+h^2*delta^2/2),
theta_h = theta + h*delta^2, martingale compensator recomputed:
```
h=-0.5: lam 6.893 th -3.35%  price 17,475  KO 0.4451
h= 0.0: lam 6.785 th -2.99%  price 17,312  KO 0.4403
h=+0.5: lam 6.690 th -2.63%  price 17,282  KO 0.4308
h=+1.0: lam 6.608 th -2.28%  price 17,291  KO 0.4280
```
Price impact of a full self-consistent jump-risk-premium band: ~1.1%.

### 11.5 Paper-3 feeds (200k, daily, exact first passage)

```
stress-conditional p_KO from (113, 1550), physical drift:
    symmetrized pool: 0.8937      genuine two-regime: 0.8925
    (the prior weekly-monitored figure 0.8405 understated touch frequency
     by ~5pp -- weekly gaps miss intra-week touches; daily exact monitoring
     is strictly more accurate for a continuously-monitored barrier)
standalone single-asset American KO P_K (200k): 11.2792 USD/bbl
    = 17,377.11 KRW/bbl -> total 34.754bn  (P_V total 41.259bn)
p_bar = (41.259-34.754)/105.586 = 6.16%
```

### 11.6 Fixup pass (fixup_200k.py, paper3_numbers.py)

**AB dual, tightened** (V-hat = max(intrinsic, continuation) on both the
outer martingale increments and the inner conditional means; 2,000 outer x
64 inner, exercise dates every 10 steps + maturity):
```
lower (LSMC) 17,442.62 +- 44   upper 17,402.79 +- 95   gap -0.23% ~ 0
```
The duality gap is statistically indistinguishable from zero: the fitted
exercise policy is near-optimal and the price is pinned from both sides.
(The first-pass 80% gap came from using the bare continuation surface
without the max -- a diagnostic bug, not a property of the price.)

**Faithful v2 hedge engine on the exact banks (200k):**
```
exercise 0.5305  KO 0.1712   HC(c=1): mean -8.99bn std 51.90bn
   (production: mean +4.24bn std 54.41bn -> std within 4.6%; mean gap is
    the known exercise-rate mismatch, scoped as replica-derived)
c* = -0.5482   std(c*) 48.76bn (6.1% reduction)   linearity 6.1e-16
```
A hastily rewritten hedge loop inside upgraded_engine_200k.py gave std
82bn (position-seeding + delta-formula deviations from the VBA); it was
DISCARDED in favor of the validated v2 loop -- Phase F's JSON is
superseded by phaseF_fixed.json. The B76-proxy replica (mean -10.05bn vs
production 66.91bn) does NOT reproduce the production European engine
(unported sizing branch: pos_FX uses S0_WTI always at line ~1310) and is
NOT used in any paper; the American-vs-B76 headline remains
production-sourced.

**Paper-3 final numbers (paper3_final.json):**
```
P_K standalone (200k exact): 17,377.11 KRW/bbl -> 34,754,224,635 total
saving P_V - P_K = 6.505bn      p_bar  = 0.06161
p_star(standalone, S8) = 0.06243     p_floor = 0.04218
stress p_KO = 0.8925 +- 0.0007 (two-regime) / 0.8937 (symmetrized)
S7 Shapley ledgers @ p=0.8925: adopted 136.60bn, min attainable 127.72bn
S8 standalone adopted ledger @ 0.8925: 130.16bn
multiples: p_KO/p_bar = 14.5x, p_KO/p_star7(0.109) = 8.2x
```
STRUCTURAL SHIFT in paper 3: with the exact-first-passage P_K (cheaper
than the prior weekly-monitored derivation) the switch point p_bar (6.16%)
and the pure-KO death point p_star (6.24%) nearly coincide -- the
optimizer now rides the KO leg almost to the brink of the pure-KO book's
own infeasibility, where previously (p_bar 4.05% << p_star 6.19%) it
abandoned KO well before. Verdict unchanged: measured mortality 89% is
14x past the switch; all-vanilla.

### 11.7 Paper integration

- Paper 1 (Park_quanto.tex): new Section 8 "Numerical Rigor" (E_jump
  measured; error decomposition with orders; N/steps convergence; dual
  bound; Esscher-consistent sweep; 200k re-derivations; variance-reduction
  autopsy); contributions block (i)-(vi) added to the Introduction;
  literature-gap paragraph ("the convention lives in the seam between
  three literatures") in the Literature Review; industry-implications
  paragraph (treasury / commodity desk / structured desk) in Discussion;
  c*/parabola numbers updated to 200k exact (-0.548, 51.90->48.76bn);
  Limitations items on the bridge and on Esscher rewritten around the new
  measured/implemented results. 31pp, 0 errors; new overfulls 0.45/2.98pt
  (invisible).
- Paper 2 (Park_CFH.tex): eq:jumpgap now carries the measured +0.14/+0.18pp
  value; Limitations discretization bullet updated + exact-engine
  directional re-confirmation of the combined>>split ordering. 27pp.
- Paper 3 (Park_hedge_optimization.tex): every stress/threshold number
  updated (see 11.6); measurement paragraph rewritten to jump-adapted
  exact first passage at 200k; figures fig_pko_haircut / fig_mixed_switch
  regenerated with the ORIGINAL ledger formulas (Shapley K1A/K2A curves)
  and new thresholds; internal inconsistency fixed (the old text's KO-leg
  total 36.99bn did not equal its own per-bbl price x barrels; all totals
  now derive from one P_K). 21pp, 0 overfull.
- fig_cstar_parabola regenerated at 200k exact (Delta_Simulation/figures).
- Backups of all three pre-revision pdf/tex: Modeling/_backups_20260710/.

### 11.8 Reviewer-response pass (convergence_tables.py)

Reviewer asks: shorter abstract, less defensive prose (Secs 3/5 -20~30%),
EM detail -> appendix, REAL basis alternatives, explicit convergence tables.

**T1 paths (nested subsamples of ONE 500k exact bank):**
```
N       price       se     delta0   KO
10k     17,298.5   196.6   0.692    0.4340   <- delta noisy at 10k!
50k     17,414.9    88.4   0.547    0.4395
100k    17,394.0    62.5   0.544    0.4391
200k    17,468.5    44.4   0.543    0.4397
500k    17,442.7    28.0   0.545    0.4397
```
**T2 steps (@50k):** 109/217/434/868 -> price 17,141/17,574/17,702/17,812
(Bermudan exercise-date density, shrinking increments), KO flat
0.4403/0.4398/0.4395/0.4395 (exact first passage: grid-invariant).

**T3 basis (@200k):** poly6 17,468.5 | hinge9 17,468.5 (EXACT no-op --
max(v1-1,0) is linear on the ITM-only fitted region, so kink-aware bases
are redundant BY CONSTRUCTION under LSMC's ITM-only regression -- a
methodological finding) | Laguerre-functions-6 (e^{-x/2}-weighted,
genuinely non-polynomial) 17,502.5 (+0.19%) | cubic10 17,638.7 (+0.97%).
Basis risk bounded at ~1%.

**Text surgery on Park_quanto.tex:** abstract 5,244 -> 1,635 chars, nearly
all digits removed; EM mixture mechanics moved to new Appendix A (main
text keeps hazard statement + moment-matching eqs + table); Sec-3 quanto
paragraph cut ~60%; three tables inserted in Sec 8 replacing inline number
lists; Sec 5.2 now points to the measured basis study. 33pp (from 35
despite +3 tables +1 appendix), 0 errors, no new visible overfulls.

### 11.9 Dual-as-pricing-model + second compression pass (dual_production.py)

Per directive, the pricing model is now DUAL-BOUNDED LSMC: price reported
as an interval, not a point.

Production-resolution AB dual (5,000 outer x 128 inner, 44-date skeleton
= every 5th day + maturity, Vhat = max(intrinsic, continuation) throughout):
```
V0 in [17,443 +- 44,  17,885 +- 59]  KRW/unit    gap +2.54%
```
The earlier coarse dual (2,000x64, every 10th day) showed gap ~0 because a
sparse skeleton TRUNCATES the pathwise max -- it was not a valid upper
bound for the daily contract. The production-resolution figure is the
honest one, and its 2.5% splits into (i) skeleton truncation (downward)
and (ii) inner-noise convexity bias (upward); the coarse-skeleton collapse
to ~0 suggests most of the 2.5% is estimator bias, not genuine policy
suboptimality. Paper text (sec 8.3) rewritten to state exactly this, with
the interval as the headline pricing object.

Second compression pass (reviewer: "too kind a paper"):
- Girsanov 3.3: mechanical derivation collapsed to inline; the praised
  theta_{2,indep} double-counting argument kept intact. ~55% shorter.
- Bridge 5.4 intuition paragraph: ~60% shorter (formulas untouched).
- 6.3 trimodality paragraph: ~40% shorter. 6.4 alive-ITM: ~60% shorter.
- 7.1 FD-vs-regression explanation: halved as requested.
- 7.3 re-implementation caveat: 1 sentence version.
- 7.2-7.4 (the novel claims) untouched, per reviewer instruction.
Compile: 33pp, 0 errors, 5 overfulls (4 pre-existing + 0.45pt invisible).

Next-paper note (NOT started, per instruction): FFT-based PIDE vs LSMC for
two-asset American exercise -- both assets American -- comparison study.

### 11.10 Trilogy-wide exact-engine unification pass (cfh_exact_200k.py)

Delta paper: third prose-compression batch (rigor intro/e-jump/Esscher/
rigor-200k/Discussion-tension/Conclusion all halved; equations untouched);
32pp.

CFH on the exact engine at 200k (production accounting mechanics: ridge +
intrinsic-floor fit, closed-form HStar, intrinsic lower-of, h_t, KO-day
COH sweeps):
```
combined KO 13.72bn | combined vanilla 6.43bn | split KO 1.99bn | split vanilla 0.00
ratio (KO) 6.9x     freeze split: early 52.6% / late 47.4% (of KO paths)
```
CRITICAL HONESTY FINDING: across engine variants (Euler-20k daily w/
safeguards: 24.78/55.77; exact-200k w/o safeguards: 34.19/6.42; exact-200k
w/ safeguards: 13.72/6.43) the combined-VANILLA cell's level -- and even
its ordering vs combined-KO -- is implementation-sensitive. The "barrier
is a partial mitigant / removing it more than doubles ineffectiveness
(55.8bn)" claim is NOT robust and was REMOVED from Park_CFH.tex (abstract,
sec 9.1, Discussion). What IS robust in every variant: split-vanilla = 0
(mechanical), split-KO small & sweep-driven, combined >> split in both
barrier states (3.3x-17x). The paper now claims exactly that and no more.
Freeze-split numbers updated to the exact-200k measurement (52.6/47.4).

Paper 3 already runs on the new value engine (P_K, p_KO at 200k exact,
sec 11.6); its Shapley coefficients (K1A/K2A) remain production-derived --
they are a cost-reporting convention the paper's own sec 8 deprecates for
decisions, and no decision-relevant quantity depends on them.

### 11.11 Typo audit + non-inherent-limitation resolution pass

**Typo audit**: dictionary-based spell-check (macOS /usr/share/dict/words +
domain vocabulary) across all three .tex files, including \caption/\section/
\item/\footnote prose (not just body text). Zero genuine typos found;
"barres" does not appear anywhere in Modeling/. Three false positives from
the checker's own text-unwrapping (missing separators between adjacent
LaTeX commands) verified against source and confirmed non-issues.

**k-sigma threshold robustness (Park_quanto sec:lambdarob)** -- textbook
case of the requested fix: previously ran on "a separate, smaller,
independently-seeded path bank," explicitly flagged as not comparable to
baseline. Re-run on the SAME jump-adapted exact engine, SAME n=200,000,
for k in {2.5, 3.0, 3.5}:
```
k    lambda    price      KO      totP&L mean    totP&L std
2.5  15.714   17,610.40  0.1744   -54.97bn        64.30bn
3.0   6.790   17,481.98  0.1727   -56.16bn        66.18bn
3.5   3.686   17,455.32  0.1711   -54.95bn        64.80bn
```
KO rate now directly comparable to Section 6.1 baseline (0.4391 exact /
0.4512 physical) since same engine+N. Table, figure (fig_lambda_robustness,
regenerated), and the "not directly comparable" limitation bullet updated/
removed accordingly.

**CFH FVTPL-split magnitude gap** -- re-tested at 200k exact with the SAME
ridge+intrinsic-floor safeguards that closed the analogous ineffectiveness-
ratio gap (17.1x -> 6.9x): result unchanged (mean 0.28bn/std 5.14bn vs.
production 1.51bn/71.95bn, ~6x gap persists). This RULES OUT sample size
and fit-instability as the cause -- the gap is genuine and not an artifact
of the kind the user asked to eliminate. Limitations bullet rewritten to
state this precisely (confirmed non-artifact) rather than leave it as an
open "not yet isolated" question.

**Abstracts**: Park_CFH.tex and Park_hedge_optimization.tex abstracts
rewritten in Park_quanto.tex's compact style -- headline mechanism and
qualitative findings only, specific figures removed except where load-
bearing for the core claim. CFH: 3,359 -> 2,026 chars. Ratio: 5,387 ->
2,081 chars.

Compile status: Park_quanto.tex 32pp/4 overfull (pre-existing, invisible),
Park_CFH.tex 26pp/1 overfull (pre-existing), Park_hedge_optimization.tex
unchanged from sec 11.10. Revision-log language scan: 0 hits, all three.

### 11.12 Root-cause resolution of the two remaining open discrepancies

**CFH FVTPL gap (was: "6x smaller, cause unresolved") -- ROOT CAUSE FOUND.**
Fresh VBA re-extraction from the live workbook (diff against the original
extraction: identical, no code drift) plus a direct, formula-aware read of
`CFH_Comparison` (openpyxl, `data_only=False`) established that A21/B21
("B mean post-KO FVTPL") and A22/B22 ("B std post-KO FVTPL") are **static
values, not formulas**, sitting at rows the live `WriteComparison` routine
neither clears (`Range("A1:G20").ClearContents`) nor writes to (the sub
ends at row 20, "Number of derivative lines"). `tB_postKO()` is populated
per-path (line 544) but never aggregated into a mean/std anywhere in the
current 46,003-byte module. **These two cells are orphaned output from a
version of the routine that no longer exists in the workbook** -- not
reproducible from the current engine by construction, regardless of
sample size, safeguards, or path law. (Every OTHER headline statistic --
H1-H3, all of them -- was verified to trace to live-populated arrays
aggregated by code still present in the module; the staleness is confined
to this one H4 pair.) This fully explains the unexplained "6-13x, then 6x"
gap from sections 11.6/11.11: no amount of re-implementation fidelity
could have closed it, because the target was never a reproducible number.

Resolution: Park_CFH.tex now reports its own n=200,000 exact-engine
measurement (mean KRW 0.28bn, std KRW 5.14bn on knocked-out paths; early-
KO subgroup alone: mean 0.53bn, std 7.08bn) as the paper's primary,
authoritative H4 figure, with the stale-cell finding stated as a direct
methodological result (Section 8.4) rather than an unresolved caveat.
Every downstream reference (zombie-path figure caption, freeze-mechanism
section, Discussion) updated to the new number.

**Delta paper exercise-rate gap (was: "0.541 vs 0.4474, cause unknown")
-- NARROWED, not fully closed.** `exercise_rate_isolation.py`: ran the
IDENTICAL LSMC-fit + exercise-decision logic on (A) a faithful
Euler-Maruyama + whole-step-bridge-on-jump-free-steps port of
`Build_Path_Bank` (matching production's exact path law) and (B) the
jump-adapted exact engine, both at n=100,000, same regression basis:
```
                  KO      exercise   raw barrier-touch (no exercise)
(A) Euler+bridge  0.1716  0.5209     0.4484
(B) exact         0.1721  0.5291     0.4533
production        0.2309  0.4474     0.4369
```
(A) and (B) are nearly identical to EACH OTHER and both diverge from
production by a similar amount -- this rules out "exact vs.\ Euler path
law" as the cause (a live, testable hypothesis the reviewer implicitly
raised) and localizes it to the fitted continuation-value surface's
exercise boundary, common to both re-implementations, not to path
discretization. Full closure would require bit-for-bit reproduction of
Excel's `LinEst` on its own RNG stream, judged out of proportion to the
value it would add; the paper now states the narrowed, evidenced finding
rather than the vaguer "differs enough" caveat.

**Abstracts** shortened further per request: Delta 1,635->1,225 chars,
CFH 2,026->1,479, Ratio 2,081->1,531.

Final compile status: Park_quanto.tex 31pp (4 pre-existing invisible
overfulls), Park_CFH.tex 25pp (1 pre-existing), Park_hedge_optimization.tex
20pp (0 overfull). Revision-log scan: 0 paper-revision-history hits in all
three (one CFH hit is legitimate software-version-history prose, verified
and retained).
