# VBA

## Critical distinction

| Location | What it is |
|----------|------------|
| **`from_xlsm/`** | Modules **extracted from live `.xlsm`** via `oletools.olevba` (what Excel actually runs) |
| **`paper_companion/`** | Corrected / rewritten modules aligned to **paper prose** after the Python audit — **not** what the workbooks shipped |
| ~~`Bas/`~~ | Historical loose dumps — **not in this repo**, and **not identical** to xlsm |

The loose files under the old `Bas/` folder and the `*_v3.bas` / `*_v4.bas` companions **differ** from the embedded workbook code (size, headers, and logic). Do not treat them as “the production engine.”

---

## `from_xlsm/` — production (workbook-truth)

Each subfolder is one workbook. Every `.bas` was pulled from that file’s `vbaProject.bin` and **byte-checked** against a fresh extract.

| Folder | Source workbook | Notes |
|--------|-----------------|--------|
| `Ratio_Optimization/` | `Ratio_Optimization.xlsm` | Includes `Ratio_Optimization.bas` + shared library |
| `Hedge/` | `Delta_Simulation/Hedge.xlsm` | Shared calib / delta / CFH (older CFH) |
| `Hedge_복사본/` | `Hedge_Simulation/Hedge 복사본.xlsm` | **Newest CFH** = `CFH_Accounting_revised.bas` (~46 KB, PRODUCTION v3.1) |

See `from_xlsm/MANIFEST.md` and each folder’s `SOURCE.txt`.

### Cross-workbook note

Shared module names are **not** always byte-identical across workbooks (e.g. `DeltaHedging_revised.bas` differs slightly between `Ratio_Optimization.xlsm` and `Hedge.xlsm` — mostly identifier casing). Always open the folder for the workbook you care about.

### Re-extract

```bash
python3 - <<'PY'
from oletools.olevba import VBA_Parser
import sys, os
path, outdir = sys.argv[1], sys.argv[2]
os.makedirs(outdir, exist_ok=True)
vp = VBA_Parser(path)
for _, _, name, code in vp.extract_macros():
    if name and name.endswith('.bas'):
        open(os.path.join(outdir, name), 'w', encoding='utf-8').write(code or '')
        print('wrote', name, len(code or ''))
vp.close()
PY
# python3 extract.py /path/to/Hedge.xlsm ./out
```

---

## `paper_companion/` — paper-aligned (intentionally different)

These are **not** scraped from xlsm. They exist because the rewritten papers describe procedures the production VBA did not implement (see `docs/ANALYSIS_PYTHON.md`):

| File | Intent |
|------|--------|
| `Ratio_Optimization_v3.bas` | Standalone KO premium + mixed vanilla/KO (not Shapley attribution as market price) |
| `CFH_Accounting_forwards_v4.bas` | Structure B as **linear forwards** (production still prices KO call + GK option) |
| `Calibration_asymmetric_v3.bas` | Asymmetric up/down jump EM (production pools jumps) |

Use for demo / audit of the paper narrative. Do not overwrite production without review.
