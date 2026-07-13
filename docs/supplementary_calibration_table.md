# Shared Calibration Master Table (Part A)

Cross-referenced 2026-07-13 directly against the three current `.tex` sources plus the underlying engine logs (`Delta_Simulation/scraped_data.json`, `Hedge_Simulation/cfh_data.json`, `Ratio_Optimization/opt_results.json`, `Ratio_Optimization/corrected_results.json`) and the production VBA source (`Delta_Simulation/Calibration_asymmetric_v3.bas`). Paper labels follow the titles in the current instruction: **P1** = IFRS 9 (`Hedge_Simulation/Park_CFH.tex`), **P2** = WTI–FX allocation (`Ratio_Optimization/Park_hedge_optimization.tex`), **P3** = Delta hedging (`Delta_Simulation/Park_quanto.tex`).

## A1. Asset dynamics parameters

| Parameter | P1 (CFH) | P2 (Ratio) | P3 (Delta) | Status |
|---|---|---|---|---|
| σ1 raw historical (WTI) | — (not stated) | 0.39455 | 0.39455 | Consistent |
| σ1 diffusive (jump-stripped) | 0.32419 | 0.32419 | 0.32419 | Consistent |
| σ2 (FX vol) | 0.09258 | 0.09258 | 0.09258 | Consistent |
| ρ (WTI–FX corr) | 0.0876 | 0.08763 / 0.0876 (both appear) | 0.0876 | Consistent (P2 shows both a 5-dp table value and a 4-dp rounded prose value — not a contradiction) |
| λ (symmetrized jump intensity) | 6.7846 | 6.785 (rounded) | 6.7846 / 6.79 (rounded in a sensitivity table) | Consistent |
| θJ (symmetrized) | -0.0299 *(fixed this pass — was -2.99% before)* | not directly stated as a standalone symmetrized figure | -0.0299 *(fixed this pass — was -2.99% before)* | Now consistent in decimal form |
| δJ (symmetrized) | 0.0844 *(fixed this pass)* | not directly stated | 0.0844 *(fixed this pass)* | Now consistent |
| λup / θup / δup | not used | 2.328/yr, +8.04%, 1.68% | 2.328/yr, +8.04%, 1.68% | Consistent — P2 and P3 both use the asymmetric fit for the stress-conditional KO probability specifically; it does **not** feed the affine risk objective's σ1 in P2 (confirmed: P2 §6.2 uses the symmetrized diffusive σ1=0.32419 for the American risk objective, the asymmetric params are used only in §7's stress-conditional simulation) |
| λdn / θdn / δdn | not used | 4.462/yr, -8.75%, 2.78% | 4.462/yr, -8.75%, 2.78% | Consistent |
| rUS | — | not explicit in P2 (implicit in premiums) | 4.0% | — |
| rKRW | — | not explicit | 3.5% | — |
| rw (WACC/funding rate) | — | 0.07 (Table 1) | — | Only in P2 |

**A1 note (VBA-verified):** the up/down asymmetric fit is confirmed genuine, not illustrative, by `Calibration_asymmetric_v3.bas`'s own reference-result comment block: independently re-derived in Python against the same `Raw_Timeseries!F` column and matched to the production LSMC!B1:B3 cells to within Bessel-correction-scale rounding (δJ differs ~1.2% relative, explained in the VBA comment as an E[Y²]-recombination vs.\ direct-STDEV.S small-sample effect, not a bug).

## A2. Market data

| Item | P1 | P2 | P3 | Status |
|---|---|---|---|---|
| S1(0) | 78.94 | 78.94 | 78.94 | Consistent |
| S2(0) | — (not restated numerically in the checked excerpt) | 1540.64 | 1540.64 | Consistent |
| Historical window (raw price observations) | — | not stated | 1,301 daily observations (Fig. 1 caption) | Only P3 states this |
| Historical window (return series used for calibration) | — | **1,299** *(fixed this pass — was "1,300")* | 1,299 (two places) | **Corrected.** VBA source `Calibration_asymmetric_v3.bas` line 51 states explicitly: "Raw\_Timeseries!F column, n=1299 returns, k=3" — this is the ground truth. P2 previously said 1,300 daily observations; there is no source evidence for 1,300, only for 1,299. Fixed in P2 (two occurrences, §Data and §Limitations). |
| Stress WTI | 113 (implied via companion refs) | 113 USD/bbl | 113 (via companion refs) | Consistent |
| Stress USD/KRW | 1550 (implied) | 1550 | 1550 | Consistent |

**Open question (not fabricated, flagged honestly):** P3's "1,301 daily observations" (raw price series, Figure 1) vs. the VBA-confirmed "1,299 returns" used in the actual EM fit differ by 2, which is *consistent* with 1,301 prices → 1,300 possible return pairs → 2 further trimmed for an unstated reason (e.g., edge-effect trimming). This is not necessarily an error, but the paper does not explain the 2-observation gap. Logged in `OPEN_ISSUES.md` rather than silently resolved, since I have no source confirming the reason.

## A3. Contract specification

| Item | P1 | P2 | P3 | Status |
|---|---|---|---|---|
| Barrier U | 120 | 120 | 120 | Consistent |
| Barrier L | 50 | 50 | 50 | Consistent |
| T_WTI | 0.833 | 0.833 (T1) | 0.833 (single T) | Consistent |
| T_FX | 0.5 | 0.5 (T2) | not modeled separately | P3 deliberately studies a single-maturity quanto KO (its own scope statement, §1: "the FX leg" in P3 is not a separately-maturing leg the way Structure B's FX leg is in P1 — P3 studies the *combined* Structure-A-type instrument only, so a single T is the correct, intentional scope, not an oversight; this matches P1's own Structure A, which also uses a single combined maturity) |

## A4. Premium / pricing figures — the highest-priority check

| Quantity | Value | Source | Status |
|---|---|---|---|
| Base premium (per quanto unit, KRW/bbl) | 17,132.47 | P1 §6 (implied), P3 §6.1, `scraped_data.json: american.base_premium=17132.47079434821` | Consistent across all sources |
| **Total inception premium (2,000,000 bbl)** | **KRW 34,264,941,589** | P1 §8.1 states this exactly; independently confirmed two ways: (1) `cfh_data.json: comparison.premium_outflow_check = 34264941588.69642` and `A_FV0_mean = 34264941588.701107`; (2) direct arithmetic `17132.47079434821 × 2,000,000 = 34,264,941,588.696` | **P3 previously stated KRW 34,264,941.59 (§6.1) — a 1000× scale error. Fixed this pass.** Root cause: `Delta_Simulation/scraped_data.json`'s own `total_premium` field (34264941.58869642) itself carries the same 1000× bug, and P3's prose was drawn directly from that buggy field rather than cross-checked against `base_premium × barrels`. **Note for the author:** the underlying JSON log's `total_premium`/`opt_max` fields are still wrong in the data file itself; only the paper's prose has been corrected here. If that JSON is reused downstream, the field should be regenerated. |
| Shapley WTI premium | 15,093.75 KRW/bbl | P1 (§8-adjacent), P2 §3.2 and Table 1, P3 §6.1 | Consistent, same units (KRW/bbl) in all three |
| Shapley FX premium | 2,038.72 KRW/USD | P2 §3.2 and Table 1, P3 §6.1 | Consistent, same units (KRW/USD) |
| Standalone KO price P_K (P2 §8) | 17,377.11 KRW/bbl | P2 §8: "an independent single-asset LSMC American KO call, priced with no FX coupling... comes to P_K = KRW 17,377.11/bbl — 15.1% above the Shapley share" | P2's own text already explains *why* this differs from the 15,093.75 Shapley figure — it is a genuinely different pricing object (standalone-tradable vs. cooperative-game attribution of the jointly-priced structure), not a discrepancy requiring reconciliation. Confirmed the explanation is present and adequate. |
| Base premium at other sections (P3 §7.x sensitivity, §8.3 exact engine) | 17,468.5 (200k exact engine), and other sweep-specific values | P3 §8 (now folded into the merged "Model Validation and Robustness" section) | These are *intentionally different* numbers from *different engines* (production 217-step Euler vs. 200,000-path jump-adapted exact), and P3's text already states this distinction at each point of use (confirmed while relocating these tables in the prior restructuring pass). The specific figure "17,188.10" cited in the review instructions does **not** appear anywhere in the current P3 source — logged as not-applicable in `OPEN_ISSUES.md` rather than invented a fix for a number that isn't there. |

## A5. KO rate / exercise rate — disambiguation table

| Figure | Value | What it actually measures | Papers using it |
|---|---|---|---|
| Unconditional KO rate (net of early exercise) | 23.09% | American-style engine, full unconditional path population, 10,000/production paths | P2 §7 |
| Unconditional raw barrier-touch rate | 45.1% (P1) / 43.69% (P3, "European/Asian 0.4369") | Raw first-passage touch frequency (no competing exercise event) | P1 (shared path bank, 45.1%), P2 §7 (45.1%), P3 §6.1 (0.4369 = 43.69% — **note the 45.1% vs 43.69% figures are NOT the same quantity**: P1/P2's 45.1% is measured on the *shared/accounting* path bank; P3's 0.4369 is the *pricing* path bank's European/Asian barrier-touch rate. These are two different Monte Carlo banks by design (different path counts/seeds), both legitimately "raw barrier touch," and the ~1.4pp gap between them is ordinary sampling noise between independently-seeded banks, not an error. This distinction should be, and after this pass now is, traceable via this table rather than left for a reader to reconcile unaided.) |
| American KO rate + exercise rate | 23.09% + 44.74% = 67.83%; alive-to-maturity 32.17% | Full three-way partition of the American path population | P1, P3 §6.1 (sums to 100.00% exactly: 23.09+44.74+32.17) |
| Exact-engine KO rates (200k jump-adapted) | 0.4391 (Q-bank) / 0.4512 (P-bank) | First-passage-exact touch rates on the *exact* engine, Q- and P-measure banks respectively | P3 §Model Validation (formerly §8.3.1) |
| Stress-conditional KO probability | 89.25% ± 0.07pp (200,000 paths) | Re-simulated from the stress state itself (WTI=113), not the unconditional average | P2 §7 |

**A5 conclusion:** these five-plus figures are not five measurements of one quantity — they are legitimately different quantities (unconditional vs. stress-conditional; net-of-exercise vs. raw-touch; production-bank vs. exact-engine-bank) that happen to share the vocabulary "KO rate." None of the differences found are errors; the risk is reader confusion, which this table is designed to close. Recommend each paper add a one-line footnote at first use of any KO-rate figure pointing to this table.

## A6. Regression basis terminology

The continuation-value surface is `β0 + β1·v1 + β2·v2 + β3·v1² + β4·v2² + β5·v1·v2` — five regressor terms plus a constant, six coefficients total (β0 through β5) — identical object in P1 and P3. P1 calls it "5-term" (one occurrence). P3 calls it "5-term" in three places and "6-term" in its basis-robustness table (counting the constant, for consistency with that table's other rows: "9-term" hinge-augmented, "10-term" cubic, which do count their own constants). **Fixed this pass:** added one clarifying parenthetical at P3's formal basis definition (§Numerical Model Formulation) reconciling both conventions explicitly, rather than renaming either — renaming the table would have broken its internal consistency with the Hinge-9/Cubic-10 rows, which also count constants.

## Items flagged in the original review instructions but NOT found in current source (see `OPEN_ISSUES.md` for the full list)

- P1 TOC "15 The Hypothetical Derivative" — current source correctly shows "5 The Hypothetical Derivative"; no fix needed.
- P3 §7.5 "17,188.10" — string does not appear anywhere in `Park_quanto.tex`.
- P1 Appendix B "verified equal... within 10⁻⁶–10⁻⁸ KRW absolute" — the current appendix does state a `$10^{-6}$--$10^{-8}$` KRW absolute range (confirmed present, contrary to my own first-pass search which used the wrong dash character); no fix needed, this one is real and correct as written.
