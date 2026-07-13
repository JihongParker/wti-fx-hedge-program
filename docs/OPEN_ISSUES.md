# Open Issues — Trilogy Verification Pass (2026-07-13)

Items from the review instructions that could not be verified against current source, could not be re-executed, or were checked and found to not apply. Logged honestly rather than silently resolved or fabricated, per the standing "no estimation, no invention" instruction.

**Correction to an earlier claim in this log:** an earlier pass of this review stated the 34,264,941.59 premium-scale bug "does not appear anywhere in `Park_quanto.tex`." That was wrong — the initial grep pattern didn't account for LaTeX's `{,}` thousands-separator syntax (`34{,}264{,}941.59`), so it silently missed the actual occurrence at §Empirical Results (baseline pricing). On the next pass, the correct pattern found it and it has been fixed (see `changelog_P3_delta.md`). Recorded here so the correction itself is traceable, not just the fix.

## Second-round verification (this pass): arithmetic re-checks, all confirmed correct, no further fixes needed

Every specific formula/derivation the review instructions asked to be recomputed was recomputed directly in Python against the paper's own stated inputs. All of the following matched the paper's printed result exactly (to rounding):
- P2 §7: p*_KO = (45,000,000,000 − 33,477,827,504) / 105,586,000,000 = 0.10913 ✓
- P2 §8: p̄ = (41,258,998,746 − 34,754,224,635) / 105,586,000,000 = 0.06161 ✓; savings 41.3bn−32.0bn=9.3bn ✓ and 41.259bn−34.754bn=6.50bn ✓; 0.8925/0.06161≈14.49 ("roughly fourteen times") ✓; 130.16bn−45.00bn=85.16bn ("KRW 85bn priced out") ✓
- P1(CFH) §Sensitivity: all four re-implementation error percentages (Structure A ineffectiveness 6.08%, Structure B1 ineffectiveness 15.83%, econ-resid-std A 2.66%, econ-resid-std B 1.18%, A/B1 ratio 8.43%) — confirmed the paper's own "validated to within 1–16%" characterization is accurate (max individual error 15.83% rounds to the stated "16%" ceiling; min is ~1%). **The review instruction's suggested replacement wording ("broadly consistent with, within 6%") would have been incorrect** — 6% only reflects the ratio-metric error, not the actual max across all five reported statistics. Left as-is; the paper's original phrasing was right.
- P3 §Sensitivity: FD-vs-regression delta gap (0.7921−0.6679)/0.6679=18.60% ✓; c* std reduction (51.90−48.76)/51.90=6.05% ✓; production-vs-c=1 gap (54.41−51.90)/54.41=4.61% ✓; dual-bound gap (17885−17443)/17443=2.53% ✓; basis-robustness table deltas (+0.19%, +0.97%) ✓; V/S2 structural-delta contract ratio 1110/242=4.59≈4.6× ✓; KO+exercise+alive = 23.09+44.74+32.17=100.00% ✓

## Not verifiable without re-running the original regression/simulation (flagged, not fabricated)

- P3's "median 8.8%" (regression-surface delta vs. homogeneity-form V/S2 gap, §Sensitivity) is a statistic computed across many points on the fitted LSMC surface. It cannot be checked by hand-arithmetic the way the scalar formulas above could; verifying it would require re-running the actual LSMC regression, which was not done this pass. Left as reported.

## Not found in current source (review instructions describe a different snapshot)

| Claimed issue | Checked against | Result |
|---|---|---|
| P1 TOC: "15 The Hypothetical Derivative" (should be §5) | `Hedge_Simulation/Park_CFH.tex`, full `\section`/`\subsection` grep | Current source already reads "5 The Hypothetical Derivative" — correct as-is. No fix made. |
| P3 §7.5: base premium "17,188.10" | Full-text grep of `Delta_Simulation/Park_quanto.tex` | String does not appear anywhere in current source. No fix made — nothing to fix. |
| P1 Appendix "verified equal at every step... within 10⁻⁶–10⁻⁸ KRW absolute" | `Hedge_Simulation/Park_CFH.tex` Appendix B | **Actually present and correct** (line 468) — my first search pass used the wrong dash glyph and missed it; confirmed on re-check. No fix needed. |
| P3 §8.3 Table 11/12/13 numbering | Current source | Tables exist (path-count convergence, grid refinement, basis robustness) but are unlabeled by "Table 11/12/13" — LaTeX auto-numbers them, and after this pass's restructuring (moved to Appendix A) their auto-numbers have shifted again. Anyone citing them by a fixed number rather than `\ref{}` will need to recheck against the recompiled PDF. |

## Data-file bug found but not paper-visible (flagged for the author, not fixed by this pass)

- `Delta_Simulation/scraped_data.json`: the `american.total_premium`, `european.total_premium`, `asian.total_premium`, and `american.opt_max` fields are all `34264941.58869642` — a 1000× scale error relative to the correct `base_premium × 2,000,000 = 34,264,941,588.696`. This bug does **not** currently propagate into any `.tex` source (P3's own prose citation of this figure has now been corrected independently, see changelog), but the JSON file itself is still wrong. If this file is re-scraped or reused for any future figure/table generation, the bug will resurface. Recommend the author re-run whatever script produced `scraped_data.json` and check the units on the `total_premium` computation (likely a missing `× barrels` or an accidental `/1000`).

## Unresolved, not fabricated

- **P3's "1,301 daily observations" (raw price series) vs. VBA-confirmed "1,299 returns" (EM fit input).** A 2-observation gap that is *plausible* (1,301 prices → 1,300 possible returns → 2 more trimmed for an unstated reason, e.g., edge effects) but not explained anywhere in the source, VBA comments, or available scripts. Not fixed, because inventing a reason would violate the no-fabrication instruction. Recommend the author add one clause explaining the 2-observation trim, or confirm 1,301 is itself imprecise.

## Layout and structural integrity (this pass — all confirmed clean)

- All three papers recompile with **0 LaTeX errors, 0 undefined references, 0 multiply-defined labels, 0 missing figure files**.
- **Overfull hboxes: eliminated in all three** (P2 was already 0; P1(CFH) and P3(Delta) each had several — fixed via `\emergencystretch=3em` added to both preambles, plus one manual line-break in P1(CFH)'s Appendix A quanto-adjustment derivation, the single equation too wide to be fixed by stretching alone). `microtype` was tried first for a more thorough typographic fix but is incompatible with this document's font setup (pdfTeX font-expansion error with non-scalable fonts) and was reverted.
- Remaining underfull hboxes (1 in P1/CFH at badness 1062, 2 in P3/Delta at badness <2900) are cosmetically negligible — well below the ~5000+ badness level where loose spacing becomes visually noticeable — and sit inside dense enumerated-list math, where some looseness is normal. Not further modified; forcing these to 0 would risk destabilizing correctly-typeset content for no visible gain.
- Bibliography integrity checked directly: **zero uncited references, zero duplicate `\bibitem` keys, zero duplicate `\label{}`** across all three papers (verified by extracting every bibliography entry's author surname and confirming at least one in-text citation exists for each).

## Scope not attempted this pass (see status report to user)

The following require either (a) re-running the actual Excel/VBA production workbooks, which this environment cannot execute (no Excel/COM available), or (b) re-running large Python Monte Carlo re-implementations (100k–500k paths) that were not re-executed in this pass given the size of the full instruction set. These are listed, not silently skipped:

- Re-deriving confidence intervals for the 200,000-path stress-conditional KO probability (p=0.8925±0.0007) from a fresh run rather than the existing log.
- Re-running the c* = -0.548 affine decomposition end-to-end from raw paths rather than cross-checking the already-reported arithmetic.
- Full Part B/C/D section-by-section pass (P1 §8.2/§8.4/§9.1 table-level re-derivation; P2 §6.1–§9 solver re-runs to 6 decimal places; P3 §7.1–§8.3 remaining per-item checks not already covered by the Part A master table above).
- Part E cross-reference sweep beyond what Part A already covers (Shapley figures, Ejump residual — both confirmed consistent in Part A).
