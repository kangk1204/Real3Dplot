from __future__ import annotations

import html
import json
from importlib.resources import files
from typing import Any

PLOTLY_CDN_URL = "https://cdn.plot.ly/plotly-2.35.2.min.js"


def render_dashboard_html(payload: dict[str, Any], title: str, plotly_url: str = PLOTLY_CDN_URL) -> str:
    template = _load_template()
    payload_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    return (
        template.replace("__PAGE_TITLE__", html.escape(title))
        .replace("__PLOTLY_SOURCE__", html.escape(plotly_url, quote=True))
        .replace("__APP_DATA__", payload_json)
    )


def _load_template() -> str:
    template_path = files("three_d_plot_dashboard").joinpath("templates/dashboard.html")
    return template_path.read_text(encoding="utf-8")
