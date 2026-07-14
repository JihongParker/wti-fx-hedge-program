# WTI–FX Hedge Program

Four companion working papers on one economic position: a Korean crude-oil importer’s joint **WTI × USD/KRW** exposure.

| # | Paper | Folder |
|---|--------|--------|
| 1 | Optimal WTI–FX hedge ratios under a fixed budget | [`papers/01_budget_allocation/`](papers/01_budget_allocation/) |
| 2 | Covariance-aware delta hedging of a quanto knock-out | [`papers/02_delta_hedging/`](papers/02_delta_hedging/) |
| 3 | IFRS 9 cash-flow-hedge accounting (combined vs split) | [`papers/03_ifrs9_cfh/`](papers/03_ifrs9_cfh/) |
| 4 | ESG disclosure, hedging & DiD design (Korea / KSSB) | [`papers/04_esg_disclosure/`](papers/04_esg_disclosure/) |
| 5 | KIKO through the program — applied note on the 2008 knock-in knock-out episode | [`papers/05_kiko_note/`](papers/05_kiko_note/) |

Shared calibration snapshot: [`docs/supplementary_calibration_table.md`](docs/supplementary_calibration_table.md).

---

## Repository layout

```
papers/          PDF + TeX + figures used in each paper
python/          Reproduction / figure / engine scripts (paper-driving pass)
vba/             Excel VBA for demo & audit
  from_xlsm/              modules extracted from live .xlsm (oletools) — workbook truth
  paper_companion/        corrected modules aligned to rewritten papers (NOT xlsm)
data/results/    Small JSON result dumps (not full path banks)
docs/            Calibration table, analysis notes, open issues
```

Legacy Excel workbooks, intermediate `.npz` banks, and unorganized working folders are **not** in git (see `.gitignore`). They remain on the local machine if you still have the full Modeling tree.

---

## Python

Requires a normal scientific stack, e.g.:

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install numpy scipy pandas matplotlib
# ESG OpenDART extras as needed: requests, lxml, ...
```

| Area | Entry points |
|------|----------------|
| Budget | `python/01_budget/optimize_hedge.py`, `make_paper_figures.py` |
| Delta | `python/02_delta/make_figures.py`, `make_ultimate_figs.py`, `build_ultimate_pipeline.py` |
| CFH | `python/03_cfh/make_figures.py`, `extract_data.py` |
| ESG | `python/04_esg/esg_hedge_engine.py`, `opendart_pipeline.py` |
| KIKO | `python/05_kiko/kiko_engine.py` |

Numbers that rewrote the papers after the VBA audit are documented in [`docs/ANALYSIS_PYTHON.md`](docs/ANALYSIS_PYTHON.md).

---

## VBA

**Loose `Bas/*.bas` is not production.** Those files often **differ a lot** from what is embedded in `.xlsm`. This repo only ships:

1. **`vba/from_xlsm/<Workbook>/`** — `oletools.olevba` extract of each workbook’s `vbaProject.bin`, **byte-checked** against a fresh extract.
2. **`vba/paper_companion/`** — intentionally different modules written to match **rewritten paper prose** (not the shipped engines).

| Folder | Source workbook | Flagship |
|--------|-----------------|----------|
| `from_xlsm/Ratio_Optimization/` | `Ratio_Optimization.xlsm` | `Ratio_Optimization.bas` |
| `from_xlsm/Hedge/` | `Delta_Simulation/Hedge.xlsm` | `DeltaHedging_revised.bas`, older CFH |
| `from_xlsm/Hedge_복사본/` | `Hedge 복사본.xlsm` | **CFH PRODUCTION v3.1** (~46 KB) |

Shared module *names* are not always byte-identical across workbooks — use the folder for the workbook you mean.

Paper companions (`Ratio_Optimization_v3`, `CFH_Accounting_forwards_v4`, `Calibration_asymmetric_v3`) document fixes in `docs/ANALYSIS_PYTHON.md` (e.g. production Structure B is still options, not forwards).

Details: [`vba/README.md`](vba/README.md) · [`vba/from_xlsm/MANIFEST.md`](vba/from_xlsm/MANIFEST.md)
---

## Papers only

```bash
open papers/01_budget_allocation/Park_hedge_optimization.pdf
open papers/02_delta_hedging/Park_Quanto.pdf   # or Park_quanto.pdf
open papers/03_ifrs9_cfh/Park_CFH.pdf
open papers/04_esg_disclosure/Park_ESG_disclosure.pdf
```

TeX sources sit next to each PDF. Figure paths inside TeX may still say `figures/...` relative to the original paper folder — figures are copied under each paper’s `figures/`.

---

## License / disclaimer

Working papers for research discussion. Simulation results only — not investment advice. Respect IFRS / data-vendor terms if you extend the ESG pipeline with live OpenDART keys (do not commit API keys).

Author: Jihong Park · Pusan National University · 2026
