# WTI–FX Hedge Program

Four companion working papers on one economic position: a Korean crude-oil importer’s joint **WTI × USD/KRW** exposure.

| # | Paper | Folder |
|---|--------|--------|
| 1 | Optimal WTI–FX hedge ratios under a fixed budget | [`papers/01_budget_allocation/`](papers/01_budget_allocation/) |
| 2 | Covariance-aware delta hedging of a quanto knock-out | [`papers/02_delta_hedging/`](papers/02_delta_hedging/) |
| 3 | IFRS 9 cash-flow-hedge accounting (combined vs split) | [`papers/03_ifrs9_cfh/`](papers/03_ifrs9_cfh/) |
| 4 | ESG disclosure, hedging & DiD design (Korea / KSSB) | [`papers/04_esg_disclosure/`](papers/04_esg_disclosure/) |

Shared calibration snapshot: [`docs/supplementary_calibration_table.md`](docs/supplementary_calibration_table.md).

---

## Repository layout

```
papers/          PDF + TeX + figures used in each paper
python/          Reproduction / figure / engine scripts (paper-driving pass)
vba/             Excel VBA for demo & audit
  production_from_xlsm/   modules extracted from live .xlsm (oletools)
  paper_companion/        corrected companion modules aligned to rewritten papers
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

Numbers that rewrote the papers after the VBA audit are documented in [`docs/ANALYSIS_PYTHON.md`](docs/ANALYSIS_PYTHON.md).

---

## VBA

**Do not use the historical `Bas/` folder** (left out of this repo). It is not the latest.

| Location | What it is |
|----------|------------|
| `vba/production_from_xlsm/` | Extracted with `oletools.olevba` from the production workbooks |
| `vba/paper_companion/` | Drop-in / corrected modules that match the **rewritten** paper text |

### Production extraction (verified)

| Workbook (local) | Notable modules |
|------------------|-----------------|
| `Ratio_Optimization.xlsm` | `Ratio_Optimization.bas`, shared `Calibration_revised`, `DeltaHedging_revised`, `PaperVerification_revised`, older CFH |
| `Delta_Simulation/Hedge.xlsm` | Same shared library (byte-identical calib/delta/verification) |
| `Hedge_Simulation/Hedge 복사본.xlsm` | **Latest CFH**: `CFH_Accounting_revised` PRODUCTION **v3.1** (~1068 lines) |

Stored here as:

- `vba/production_from_xlsm/CFH_Accounting_revised_v3.1.bas` ← from 복사본 (newest)
- `vba/production_from_xlsm/CFH_Accounting_revised_workbook.bas` ← from `Hedge.xlsm` (older)
- `vba/production_from_xlsm/Ratio_Optimization.bas` ← as embedded in workbook
- `vba/production_from_xlsm/DeltaHedging_revised.bas`
- `vba/production_from_xlsm/Calibration_revised.bas`
- `vba/production_from_xlsm/PaperVerification_revised.bas`

### Paper companion (aligned to prose)

| Module | Role |
|--------|------|
| `Ratio_Optimization_v3.bas` | Standalone KO premium + mixed vanilla/KO (fixes Shapley misuse) |
| `CFH_Accounting_forwards_v4.bas` | Structure B as **genuine forwards** (production still priced options) |
| `Calibration_asymmetric_v3.bas` | Asymmetric up/down jump EM (paper § calibration) |

Re-extract from a local `.xlsm` if needed:

```bash
python3 - <<'PY'
from oletools.olevba import VBA_Parser
import sys, os
path, outdir = sys.argv[1], sys.argv[2]
os.makedirs(outdir, exist_ok=True)
vp = VBA_Parser(path)
for _, _, name, code in vp.extract_macros():
    if name and name.endswith('.bas'):
        open(os.path.join(outdir, name), 'w', encoding='utf-8').write(code)
        print('wrote', name)
vp.close()
PY
# usage: python3 extract.py /path/to/Hedge.xlsm ./out_vba
```

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
