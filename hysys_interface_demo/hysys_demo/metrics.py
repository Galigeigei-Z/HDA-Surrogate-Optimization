from __future__ import annotations

from collections import defaultdict

import pandas as pd

from .session import HysysContext, ensure_context


UTILITY_COST_USD_PER_KJ = {
    "FH": 4.249e-6,
    "CW": 2.125e-7,
    "Elec": 1.580e-5,
    "HP": 2.50e-6,
    "MP": 2.20e-6,
    "Ref": 3.30e-6,
}


def _match_utility_category(stream_name: str) -> str:
    suffix = stream_name.split()[-1]
    for key in UTILITY_COST_USD_PER_KJ:
        if key.lower() in suffix.lower():
            return key
    return "FH"


def collect_energy_table(*, context: HysysContext | None = None) -> pd.DataFrame:
    ctx = ensure_context(context)
    rows: list[dict[str, float | str]] = []
    for name in ctx.flowsheet.EnergyStreams.Names:
        energy_stream = ctx.flowsheet.EnergyStreams.Item(name)
        heat_flow_kw = float(energy_stream.HeatFlow.Value)
        rows.append(
            {
                "name": name,
                "heat_flow_kw": heat_flow_kw,
                "heat_flow_kj_h": heat_flow_kw * 3600.0,
                "utility_category": _match_utility_category(name),
            }
        )
    return pd.DataFrame(rows)


def calculate_utility_cost_per_hour(*, context: HysysContext | None = None) -> float:
    energy_df = collect_energy_table(context=context)
    grouped_kj_h: defaultdict[str, float] = defaultdict(float)
    for row in energy_df.itertuples(index=False):
        grouped_kj_h[str(row.utility_category)] += float(row.heat_flow_kj_h)

    total = 0.0
    for category, heat_kj_h in grouped_kj_h.items():
        total += heat_kj_h * UTILITY_COST_USD_PER_KJ[category]
    return total
