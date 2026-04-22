from __future__ import annotations

from dataclasses import dataclass

from .session import HysysContext, ensure_context


@dataclass(frozen=True)
class HdaFlowsheetMap:
    hydrogen_stream: str = "Fresh H2 Feed"
    feed_stream: str = "S5"
    linked_feed_stream: str | None = "S16"
    reactor_outlet_stream: str = "S6"
    condenser_stream: str = "S23"
    benzene_column_stream: str = "S26"
    toluene_column_stream: str = "S30"
    reactor_name: str = "Reactor1"
    purge_splitter: str = "Sep1"
    recycle_splitter: str = "Sep2"
    benzene_column: str = "T2"
    toluene_column: str = "T3"


DEFAULT_HDA_SAMPLE = {
    "H2_flow_kmol_h": 190.0,
    "feed_pressure_kpa": 3200.0,
    "feed_temperature_c": 630.0,
    "operation_mode": "adiabatic",
    "reactor_volume_m3": 160.0,
    "recycle_split_fraction": 0.80,
    "purge_split_fraction": 0.05,
    "condenser_pressure_kpa": 930.0,
    "benzene_column_pressure_kpa": 110.0,
    "toluene_column_pressure_kpa": 180.0,
    "benzene_column_trays": 41,
    "toluene_column_trays": 25,
    "benzene_feed_stage_ratio": 0.30,
    "toluene_feed_stage_ratio": 0.40,
}


def set_stream_conditions(
    stream_name: str,
    *,
    molar_flow_kmol_h: float | None = None,
    temperature_c: float | None = None,
    pressure_kpa: float | None = None,
    linked_pressure_stream: str | None = None,
    linked_pressure_offset_kpa: float = 0.0,
    context: HysysContext | None = None,
) -> None:
    ctx = ensure_context(context)
    stream = ctx.flowsheet.MaterialStreams.Item(stream_name)

    if molar_flow_kmol_h is not None:
        stream.MolarFlow.Value = molar_flow_kmol_h / 3600.0
    if temperature_c is not None:
        stream.Temperature.Value = temperature_c
    if pressure_kpa is not None:
        stream.Pressure.Value = pressure_kpa
        if linked_pressure_stream is not None:
            linked = ctx.flowsheet.MaterialStreams.Item(linked_pressure_stream)
            linked.Pressure.Value = pressure_kpa + linked_pressure_offset_kpa


def set_operation_mode(
    mode: str,
    *,
    inlet_stream: str,
    outlet_stream: str,
    adiabatic_delta_t_c: float = 40.0,
    context: HysysContext | None = None,
) -> None:
    ctx = ensure_context(context)
    inlet = ctx.flowsheet.MaterialStreams.Item(inlet_stream)
    outlet = ctx.flowsheet.MaterialStreams.Item(outlet_stream)
    inlet_temp_c = float(inlet.Temperature.Value)

    normalized = mode.strip().lower()
    if normalized == "adiabatic":
        outlet.Temperature.Value = inlet_temp_c + adiabatic_delta_t_c
        return
    if normalized == "isothermal":
        outlet.Temperature.Value = inlet_temp_c
        return
    raise ValueError("mode must be 'adiabatic' or 'isothermal'")


def set_splitter_fraction(
    operation_name: str,
    split_fraction: float,
    *,
    context: HysysContext | None = None,
) -> None:
    ctx = ensure_context(context)
    operation = ctx.case.Flowsheet.Operations.Item(operation_name)
    operation.SplitsValue = (split_fraction, 1.0 - split_fraction)


def set_reactor_volume(
    reactor_name: str,
    volume_m3: float,
    *,
    context: HysysContext | None = None,
) -> None:
    ctx = ensure_context(context)
    reactor = ctx.flowsheet.Operations.Item(reactor_name)
    reactor.TotalVolumeValue = volume_m3


def set_column_trays(
    column_name: str,
    tray_count: int,
    *,
    context: HysysContext | None = None,
) -> None:
    ctx = ensure_context(context)
    column = ctx.flowsheet.Operations.Item(column_name)
    tray_section = None
    for operation in column.ColumnFlowsheet.Operations:
        if operation.TypeName.lower() == "traysection":
            tray_section = operation
            break
    if tray_section is None:
        raise RuntimeError(f"TraySection not found in {column_name}")
    tray_section.NumberOfTrays = tray_count
    column.ColumnFlowsheet.Run()


def set_column_feed_stage_ratio(
    column_name: str,
    stage_ratio: float,
    *,
    context: HysysContext | None = None,
) -> None:
    ctx = ensure_context(context)
    column = ctx.flowsheet.Operations.Item(column_name)
    cfs = column.ColumnFlowsheet

    feed_stream = None
    for name in cfs.FeedStreams.Names:
        candidate = cfs.FeedStreams.Item(name)
        if candidate.TypeName.lower() == "materialstream":
            feed_stream = candidate
            break
    if feed_stream is None:
        raise RuntimeError(f"No material feed stream found for {column_name}")

    tray_section = None
    for operation in cfs.Operations:
        if operation.TypeName.lower() == "traysection":
            tray_section = operation
            break
    if tray_section is None:
        raise RuntimeError(f"TraySection not found in {column_name}")

    tray_count = int(tray_section.NumberOfTrays)
    tray_number = max(1, min(int(stage_ratio * tray_count), tray_count))
    tray_section.SpecifyFeedLocation(feed_stream, tray_number)
    cfs.Run()


def apply_hda_demo_sample(
    sample: dict[str, float | int | str],
    *,
    mapping: HdaFlowsheetMap | None = None,
    context: HysysContext | None = None,
) -> dict[str, float | int | str]:
    ctx = ensure_context(context)
    case_map = mapping or HdaFlowsheetMap()

    set_operation_mode(
        str(sample["operation_mode"]),
        inlet_stream=case_map.feed_stream,
        outlet_stream=case_map.reactor_outlet_stream,
        context=ctx,
    )
    set_stream_conditions(
        case_map.hydrogen_stream,
        molar_flow_kmol_h=float(sample["H2_flow_kmol_h"]),
        context=ctx,
    )
    set_stream_conditions(
        case_map.feed_stream,
        temperature_c=float(sample["feed_temperature_c"]),
        pressure_kpa=float(sample["feed_pressure_kpa"]),
        linked_pressure_stream=case_map.linked_feed_stream,
        linked_pressure_offset_kpa=50.0,
        context=ctx,
    )
    set_stream_conditions(
        case_map.condenser_stream,
        pressure_kpa=float(sample["condenser_pressure_kpa"]),
        context=ctx,
    )
    set_stream_conditions(
        case_map.benzene_column_stream,
        pressure_kpa=float(sample["benzene_column_pressure_kpa"]),
        context=ctx,
    )
    set_stream_conditions(
        case_map.toluene_column_stream,
        pressure_kpa=float(sample["toluene_column_pressure_kpa"]),
        context=ctx,
    )
    set_reactor_volume(case_map.reactor_name, float(sample["reactor_volume_m3"]), context=ctx)
    set_splitter_fraction(
        case_map.purge_splitter,
        float(sample["purge_split_fraction"]),
        context=ctx,
    )
    set_splitter_fraction(
        case_map.recycle_splitter,
        float(sample["recycle_split_fraction"]),
        context=ctx,
    )
    set_column_trays(
        case_map.benzene_column,
        int(sample["benzene_column_trays"]),
        context=ctx,
    )
    set_column_trays(
        case_map.toluene_column,
        int(sample["toluene_column_trays"]),
        context=ctx,
    )
    set_column_feed_stage_ratio(
        case_map.benzene_column,
        float(sample["benzene_feed_stage_ratio"]),
        context=ctx,
    )
    set_column_feed_stage_ratio(
        case_map.toluene_column,
        float(sample["toluene_feed_stage_ratio"]),
        context=ctx,
    )
    return dict(sample)
