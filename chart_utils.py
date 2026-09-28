# -*- coding: utf-8 -*-
"""Small, dependency-light helpers shared by the Streamlit charts.

The research pipeline receives data from several sources (SQLite, AkShare and
LLM-produced JSON).  A single ``None``, ``NaN`` or mismatched series can make a
Plotly chart fail during a Streamlit rerun.  These helpers keep chart rendering
deterministic without changing the underlying financial data.
"""
from __future__ import annotations

import math
from numbers import Real
from typing import Any, Iterable, List, Sequence, Tuple


def finite_number(value: Any, default: float = 0.0) -> float:
    """Return a finite float suitable for Plotly, or ``default``."""
    try:
        number = float(value)
        return number if math.isfinite(number) else float(default)
    except (TypeError, ValueError, OverflowError):
        return float(default)


def finite_series(values: Any, default: float = 0.0) -> List[float]:
    """Convert an arbitrary iterable into a finite numeric list."""
    if values is None or isinstance(values, (str, bytes)):
        return []
    try:
        return [finite_number(v, default) for v in values]
    except TypeError:
        return []


def aligned_series(labels: Any, values: Any, default: float = 0.0) -> Tuple[List[str], List[float]]:
    """Align x labels and y values, truncating safely to the shorter length."""
    xs = [] if labels is None else [str(x) for x in labels]
    ys = finite_series(values, default)
    n = min(len(xs), len(ys))
    return xs[:n], ys[:n]


def sanitize_json(value: Any) -> Any:
    """Recursively replace non-finite JSON numbers with zero.

    ``json.loads`` accepts ``NaN``/``Infinity`` by default, while Plotly does
    not.  Sanitising at the report boundary prevents one malformed field from
    breaking every downstream chart and export.
    """
    if isinstance(value, dict):
        return {str(k): sanitize_json(v) for k, v in value.items()}
    if isinstance(value, list):
        return [sanitize_json(v) for v in value]
    if isinstance(value, tuple):
        return [sanitize_json(v) for v in value]
    if isinstance(value, Real) and not isinstance(value, bool):
        return finite_number(value)
    return value


def apply_plotly_theme(fig: Any) -> Any:
    """Apply a compact, consistent theme while preserving chart-specific layout."""
    try:
        fig.update_layout(
            template="plotly_white",
            font=dict(family="Arial, Microsoft YaHei, sans-serif", color="#334155", size=11),
            hoverlabel=dict(bgcolor="#ffffff", bordercolor="#cbd5e1", font_size=11),
            legend=dict(itemclick="toggle", itemdoubleclick="toggleothers"),
        )
    except Exception:
        # A chart should still render if a future Plotly version rejects one
        # optional layout attribute.
        pass
    return fig


PLOTLY_CONFIG = {
    "displaylogo": False,
    "responsive": True,
    "toImageButtonOptions": {"format": "png", "filename": "research_chart", "scale": 2},
}
