# VBA

## `production_from_xlsm/`

Extracted from live workbooks via `oletools.olevba` (2026-07).

| File | Source workbook | Notes |
|------|-----------------|--------|
| `Calibration_revised.bas` | shared across workbooks | Jump/diffusion calibration used in production |
| `DeltaHedging_revised.bas` | shared | LSMC delta engine |
| `PaperVerification_revised.bas` | shared | Verification harness |
| `Ratio_Optimization.bas` | `Ratio_Optimization.xlsm` | Static allocation as shipped in workbook |
| `CFH_Accounting_revised_v3.1.bas` | `Hedge 복사본.xlsm` | **Newest** CFH engine (PRODUCTION v3.1) |
| `CFH_Accounting_revised_workbook.bas` | `Hedge.xlsm` | Older CFH in main Hedge workbook |

`Bas/` at the Modeling root is **historical** and is not shipped in this repo.

## `paper_companion/`

Modules written so the **rewritten paper text** matches code (see `docs/ANALYSIS_PYTHON.md`):

| File | Fixes |
|------|--------|
| `Ratio_Optimization_v3.bas` | Standalone KO premium; mixed vanilla/KO program |
| `CFH_Accounting_forwards_v4.bas` | Structure B as linear forwards (not KO call + GK call) |
| `Calibration_asymmetric_v3.bas` | Up/down asymmetric EM for WTI jumps |

Import into a copy of the production workbook for demo; do not overwrite production without review.
