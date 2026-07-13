# Paper ↔ VBA audit (strict)

**Date:** 2026-07-13  
**Question 1:** Is `vba/from_xlsm/**` a faithful extract of the workbooks?  
**Question 2:** Does that embedded VBA implement what the *rewritten papers* claim?

---

## 1. Extraction integrity (Q1) — **PASS**

Every `.bas` under `vba/from_xlsm/<workbook>/` was re-checked against a live `oletools.olevba` extract of the corresponding `.xlsm`. **All modules: byte-for-byte match.**

| Workbook | Modules checked | Result |
|----------|-----------------|--------|
| `Ratio_Optimization.xlsm` | 5 bas | PASS |
| `Delta_Simulation/Hedge.xlsm` | 4 bas | PASS |
| `Hedge_Simulation/Hedge 복사본.xlsm` | 4 bas | PASS |

So the GitHub VBA is **not** a lazy copy of loose `Bas/*.bas`.  
(For contrast: `Bas/DeltaHedging_revised.bas` is **79,850 B** vs `Hedge.xlsm` extract **73,113 B** — different files.)

---

## 2. Which workbook is the production truth for each paper?

| Paper | Primary VBA truth | Notes |
|-------|-------------------|--------|
| **01 Budget** | `from_xlsm/Ratio_Optimization/Ratio_Optimization.bas` | Risk/cost min under budget |
| **02 Delta** | `from_xlsm/Hedge/DeltaHedging_revised.bas` (+ path bank / LSMC) | Am/Eu/As engines; `delta_ratio=1` hard-coded |
| **02 Calib** | `from_xlsm/*/Calibration_revised.bas` | Pooled jump EM only |
| **03 CFH** | **`from_xlsm/Hedge_복사본/CFH_Accounting_revised.bas`** (v3.1, ~46 KB) | Structure B = **WTI KO call + GK FX call** |
| **03 CFH (other)** | `Hedge.xlsm` / `Ratio_Optimization.xlsm` CFH (~31 KB) | **Different engine**: Structure B uses **linear forwards** helpers |

**Important:** There are **two incompatible CFH implementations** across workbooks. The rewritten CFH paper (B1 = WTI KO, B2 = GK call, closed-form \(H^\star\)) matches **복사본 v3.1**, not the forward-style CFH in `Hedge.xlsm` / `Ratio_Optimization.xlsm`.

---

## 3. Paper claim vs xlsm implementation (Q2)

Legend: **MATCH** = paper prose is implemented in that xlsm module · **PARTIAL** = core engine present, paper’s headline extension is elsewhere · **MISMATCH** = paper describes something the xlsm module does not do · **PYTHON** = number/result comes from the Python audit pass, not VBA

### Paper 1 — Budget allocation (`Park_hedge_optimization`)

| Paper claim | xlsm `Ratio_Optimization.bas` | Verdict |
|-------------|------------------------------|---------|
| Convex risk min / cost min under budget \(B=45\)bn, \(w_1+w_2\le 1\) | `OPT_RiskMin_*`, `OPT_CostMin_*`, grid/Solver | **MATCH** |
| EU premiums Black-76 / GK | Cost coeffs from model sheet (B76/GK links) | **MATCH** |
| AM premiums from LSMC Shapley J9/J10 | `=LSMC!J9`, `=LSMC!J10` hard-wired | **MATCH** (and this is the known *conceptual* issue the paper later *corrects* for mixed program) |
| Stress \(p_{KO}\) haircut analysis | `p_KO` input exists; **report-only** stress adjustment in comments | **PARTIAL** — parameter present; not a full Monte Carlo stress engine inside this module |
| Mixed vanilla/KO program; **standalone** KO price (not Shapley); \(\bar p\), all-vanilla at stress | **No** `OPT_PriceStandaloneWTIKO` / mixed envelope solver in xlsm | **MISMATCH vs xlsm** |
| Standalone KO + \(p_{KO}^{stress}=0.8925\), switch at \(\bar p\approx 0.0616\) | Implemented in **Python** audit + `paper_companion/Ratio_Optimization_v3.bas` | **PYTHON / companion**, not production xlsm |

**Bottom line P1:** Production xlsm implements the **static allocation with Shapley AM premiums**. The paper’s **instrument-mix / standalone KO** results are **not** in the embedded `Ratio_Optimization.bas`.

### Paper 2 — Delta / calibration (`Park_quanto`)

| Paper claim | xlsm | Verdict |
|-------------|------|---------|
| Correlated jump-diffusion path bank, LSMC American engine | `Build_Path_Bank`, `Run_LSMC_Engine`, `Run_American_DeltaHedge` | **MATCH** |
| European Black-76 proxy hedge | `Run_European_DeltaHedge` | **MATCH** |
| Asian Turnbull–Wakeman-style proxy hedge | `Run_Asian_DeltaHedge` | **MATCH** |
| FX delta = \(c\cdot\delta_{WTI}\) coupling | `delta_ratio As Double: delta_ratio = 1#` then `delta_FX = delta_ratio * delta_WTI` | **PARTIAL** — coupling **exists**, but **hard-coded \(c=1\)**; no sweep of \(\{0.5,\ldots,1.5\}\) in the bas |
| Barrier / bridge touch | KO checks + bridge-style `Exp(-2*Log…)` patterns in related engines | **MATCH** (bridge logic present in CFH path gen; delta engines use KO flags on path bank) |
| Closed-form \(c^\star=-\mathrm{Cov}(A,B)/\mathrm{Var}(B)\approx -0.55\) | **No** Cov/Var minimizer, no \(c^\star\) in `DeltaHedging_revised.bas` | **MISMATCH vs xlsm** → **PYTHON** (`docs/ANALYSIS_PYTHON.md` §10; figures `fig_cstar_parabola`) |
| Structural \(\delta_{FX}=V/S_2\), 4.6× over-hedge | **No** homogeneity / \(V/S_2\) notional calc in bas | **MISMATCH vs xlsm** → **PYTHON** |
| Asymmetric up/down jump EM (\(\lambda_{up},\lambda_{dn},\ldots\)) | `DecomposeJumpDiffusion` pools **all** jump-flagged days into **one** \((\lambda,\theta_J,\delta_J)\) | **MISMATCH vs xlsm** → companion `Calibration_asymmetric_v3.bas` + Python |

**Bottom line P2:** xlsm is a real multi-engine delta hedge simulator with **fixed \(c=1\)**. The paper’s **covariance-aware \(c^\star\)** and **\(V/S_2\)** results are **post-engine analysis (Python)**, not embedded VBA.

### Paper 3 — IFRS 9 CFH (`Park_CFH`)

| Paper claim | `Hedge_복사본` CFH v3.1 | `Hedge.xlsm` / `Ratio` CFH | Verdict |
|-------------|-------------------------|----------------------------|---------|
| Structure A: single quanto KO CFH | Yes | Yes (variant) | **MATCH** (복사본) |
| Structure B1: **WTI KO call**, fixed \(S_2(0)\) | `WTI_KO_Call_BetaFV`, LSMC WTI KO price | **No** — uses `ForwardWTI` | **MATCH only 복사본** |
| Structure B2: **Garman–Kohlhagen FX call**, fixed notional | `GKCall` | **No** — `ForwardFX` | **MATCH only 복사본** |
| HD \(H^\star=\mathrm{barrels}\,S_1 S_2 e^{\rho\sigma_1\sigma_2\tau}\) | `HStar = barrels * S1 * s2 * Exp(rho*sigma1*sigma2*tau)` | No `HStar` in forward CFH | **MATCH only 복사본** |
| Signed lower-of, COH amort, KO sweep | `SignedLowerOf`, COH arrays, residual TV to ineff on KO | Partial / different structure | **MATCH 복사본** |
| Paper numbers (3.5× ineff, econ σ ~104bn, etc.) | Produced by running this engine + Python re-impl fidelity checks | N/A | **Engine MATCH; figures from engine/Python** |

**Bottom line P3:** The **rewritten paper aligns with `Hedge 복사본.xlsm` CFH v3.1**.  
Shipping only the older forward CFH from `Hedge.xlsm` would **misrepresent** the paper. Repo correctly includes **both**, separated by workbook folder.

### Paper 4 — ESG

No production VBA engine; Python (`esg_hedge_engine.py`, OpenDART pipeline). VBA audit N/A.

---

## 4. Role of `vba/paper_companion/` (not xlsm)

| File | Why it exists | Relation to paper |
|------|---------------|-------------------|
| `Ratio_Optimization_v3.bas` | Standalone KO + mixed program hooks | Aligns P1 §§7–8 (Shapley not used as KO market price) |
| `Calibration_asymmetric_v3.bas` | True up/down EM | Aligns P2 / P1 stress jump language |
| `CFH_Accounting_forwards_v4.bas` | Structure B as **forwards** | **Not** the main rewritten P3 design (P3 is option-based B); historical / alternative branch |

These must **never** be labeled “production extract.”

---

## 5. Overall verdict

### Extraction accuracy
**Accurate.** `vba/from_xlsm/**` is a true embed extract. Not a scrape of loose `Bas/`.

### Paper ↔ production VBA “perfection”
**Not a 1:1 identity.** The rewritten papers are a **hybrid**:

```
xlsm engines (path bank, LSMC, CFH ledgers, static OPT)
        +
Python audit (standalone KO, p_KO stress, c*, V/S2, asymmetric EM, mixed program)
        +
paper_companion bas (demo of fixes)
        =
paper claims and tables
```

| Paper | Can you reproduce the *full paper* from xlsm alone? |
|-------|------------------------------------------------------|
| 01 Budget (risk min tables) | Largely **yes** (Shapley AM premiums) |
| 01 Budget (mixed / stress KO 89%) | **No** — need Python / v3 |
| 02 Delta (Am vs proxy hedge cost) | **Yes** (engines exist) |
| 02 Delta (\(c^\star\), \(V/S_2\)) | **No** — need Python |
| 02 Calib asymmetric EM | **No** in production calib — need companion/Python |
| 03 CFH | **Yes** if you use **복사본 v3.1**, not Hedge/Ratio CFH |

### Integrity statement for GitHub readers

> Production VBA in this repo is extracted from Excel workbooks and verified against those workbooks.  
> Several headline results in the PDFs are computed in **Python** on top of those engines (or with corrected modules), as documented in `docs/ANALYSIS_PYTHON.md`.  
> Do not assume every equation in the paper has a one-line counterpart inside the embedded `.bas`.

---

## 6. Recommended mapping when reading code

1. **CFH paper numbers** → `vba/from_xlsm/Hedge_복사본/CFH_Accounting_revised.bas`  
2. **Delta engines** → `vba/from_xlsm/Hedge/DeltaHedging_revised.bas`  
3. **Budget static OPT** → `vba/from_xlsm/Ratio_Optimization/Ratio_Optimization.bas`  
4. **\(c^\star\), \(V/S_2\), \(p_{KO}^{stress}\), mixed program** → `docs/ANALYSIS_PYTHON.md` + `python/` + figures under `papers/*/figures/`  
5. **Do not use** `Bas/` or unlabelled companion files as “what the xlsm ran.”
