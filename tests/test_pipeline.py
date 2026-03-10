from __future__ import annotations

from argparse import Namespace

import polars as pl

from three_d_plot_dashboard.html_builder import render_dashboard_html
from three_d_plot_dashboard.pipeline import PreparedFrame, _normalize_headers, build_dashboard_payload


def make_args(**overrides: object) -> Namespace:
    base = {
        "x": None,
        "y": None,
        "z": None,
        "color": None,
        "size": None,
        "label": None,
        "title": "Test Dashboard",
    }
    base.update(overrides)
    return Namespace(**base)


def test_normalize_headers_deduplicates_and_fills_blanks() -> None:
    headers = _normalize_headers(["value", "", "value", None])
    assert headers == ["value", "column_2", "value_2", "column_4"]


def test_build_dashboard_payload_adds_derived_numeric_columns() -> None:
    frame = pl.DataFrame(
        {
            "segment": ["a", "b", "a"],
            "score": [10, 20, 15],
            "label": ["alpha", "beta", "gamma"],
        }
    )
    prepared = PreparedFrame(
        frame=frame,
        total_rows=3,
        sampled_rows=3,
        sampling_note=None,
        source_name="demo.csv",
    )

    payload = build_dashboard_payload(prepared, make_args())

    assert payload["defaults"]["x"] == "score"
    assert payload["defaults"]["y"].startswith("__")
    assert payload["defaults"]["z"].startswith("__")
    assert any(item["name"] == "__row_index" for item in payload["columnMeta"])
    assert payload["columns"]["segment"]["encoding"] == "dictionary"


def test_render_dashboard_html_injects_data() -> None:
    payload = {
        "title": "Smoke",
        "sourceName": "demo.csv",
        "sheetName": None,
        "rowCount": 1,
        "sampledRowCount": 1,
        "samplingNote": None,
        "generatedAt": "2026-03-10T10:00:00+09:00",
        "columnOrder": ["x", "y", "z"],
        "columnMeta": [
            {"name": "x", "dtype": "Float32", "kind": "numeric", "derived": False, "nullCount": 0, "min": 0, "max": 1},
            {"name": "y", "dtype": "Float32", "kind": "numeric", "derived": False, "nullCount": 0, "min": 0, "max": 1},
            {"name": "z", "dtype": "Float32", "kind": "numeric", "derived": False, "nullCount": 0, "min": 0, "max": 1},
        ],
        "columns": {
            "x": {"encoding": "number", "data": [0]},
            "y": {"encoding": "number", "data": [0]},
            "z": {"encoding": "number", "data": [1]},
        },
        "defaults": {"x": "x", "y": "y", "z": "z", "color": None, "size": None, "label": None, "filterColumn": None},
    }

    html = render_dashboard_html(payload, "Smoke")

    assert "__APP_DATA__" not in html
    assert "Smoke" in html
    assert "demo.csv" in html
    assert 'id="themeSelect"' in html
    assert 'id="runClusterBtn"' in html
    assert 'id="download3dPngBtn"' in html
    assert 'id="undoBtn"' in html
    assert 'id="linked2dPlot"' in html
    assert 'id="sideResizeHandle"' in html
