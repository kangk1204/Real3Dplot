# 3D Plot Dashboard Pipeline

This pipeline generates a browser-ready interactive 3D plot dashboard (`.html`) from Excel (`.xlsx`, `.xls`), TSV, and CSV files.

Core goals:

- Fast loading: CSV/TSV processing is built on `polars`
- Lightweight output: single self-contained HTML, with automatic sampling for large datasets
- Memory efficiency: numeric downcasting, categorical dictionary encoding, and streaming-style Excel sampling
- Cross-platform behavior: consistent on Ubuntu, macOS, and Windows 11

## Features

- Input formats: `.csv`, `.tsv`, `.txt`, `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`
- 3D axis mapping: choose `x`, `y`, `z` columns
- Style mapping: color (`--color`), size (`--size`), label/search (`--label`)
- Interactions: rotate, zoom, hover, and point-click detail view
- Exploration tools: axis-range filters, category filters, text search
- Export: export only currently visible points as CSV
- Figure-grade output: high-resolution 3D PNG export and 2D quadrant PNG export (2x2 panel)
- Themes/styles: Nebula, Paper Figure, Cartoon, Molstar-like + 3 point styles
- Figure presets: Default/Nature/Cell/NeurIPS-style presets
- Unsupervised clustering: browser-side K-means (k=2..12) with cluster colors/legend
- Large-cluster acceleration: sampled training + full-point assignment + centroid visualization
- Cluster analytics: train count, inertia, silhouette (sampled), and cluster sizes
- Undo/Redo + snapshot save/restore
- Linked 2D brushing: brush in 2D and filter linked 3D points
- Session preset export/import (JSON)
- Camera state persistence via browser localStorage

If the dataset has fewer than three numeric columns, the pipeline auto-generates helper numeric columns such as `__row_index` and `__code_<column>` so 3D exploration remains available.

## Installation

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -e .
```

Windows PowerShell:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e .
```

## Usage

Basic:

```bash
3dplot-dashboard data.csv
```

Specify output file:

```bash
3dplot-dashboard data.tsv -o out/dashboard.html
```

Set initial axis/style mappings:

```bash
3dplot-dashboard sales.xlsx \
  --sheet 0 \
  --x longitude \
  --y latitude \
  --z revenue \
  --color region \
  --size profit \
  --label customer_name
```

Adjust sampling cap:

```bash
3dplot-dashboard huge.csv --max-points 150000
```

Open the generated dashboard automatically:

```bash
3dplot-dashboard data.csv --open
```

## CLI Options

```text
3dplot-dashboard INPUT
  -o, --output PATH
  --sheet SHEET
  --x COLUMN
  --y COLUMN
  --z COLUMN
  --color COLUMN
  --size COLUMN
  --label COLUMN
  --delimiter DELIMITER
  --max-points N
  --seed N
  --title TITLE
  --plotly-url URL
  --open
```

## Performance Strategy

- CSV/TSV: read with `polars.scan_csv()`; for large row counts, apply systematic sampling before collect
- XLSX/XLS: read in read-only/on-demand style with reservoir sampling
- Downcast floats to `Float32` and integers to the smallest safe dtype
- Apply dictionary encoding to low-cardinality string columns to shrink HTML payload size
- Run clustering only on currently visible points to cap browser-side compute
- Normalize coordinates before clustering to reduce axis-scale bias; speed up large datasets via sampled training
- Reuse visible-index caches when filters/ranges/search are unchanged for better interaction latency

## Offline Use

By default, generated HTML references the Plotly CDN URL. For fully offline usage, provide a local `plotly.min.js` path with `--plotly-url`.

Example:

```bash
3dplot-dashboard data.csv --plotly-url ./vendor/plotly-2.35.2.min.js
```

## Generate Sample Data

```bash
.venv/bin/python scripts/generate_sample_data.py
```

Output files:

- `examples/cluster_demo.csv`
- `examples/cluster_demo.tsv`
- `examples/cluster_demo.xlsx`

## Tests

```bash
.venv/bin/pip install -e .[dev]
.venv/bin/pytest
```
