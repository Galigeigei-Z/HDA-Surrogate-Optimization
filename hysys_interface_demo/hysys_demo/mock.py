from __future__ import annotations

from dataclasses import dataclass

from .session import HysysContext


class ValueHolder:
    def __init__(self, value: float):
        self.Value = value


class FakeMaterialStream:
    TypeName = "materialstream"

    def __init__(self, name: str, *, molar_flow: float = 0.01, temperature: float = 25.0, pressure: float = 101.325):
        self.Name = name
        self.MolarFlow = ValueHolder(molar_flow)
        self.Temperature = ValueHolder(temperature)
        self.Pressure = ValueHolder(pressure)


class FakeEnergyStream:
    def __init__(self, name: str, heat_flow_kw: float):
        self.Name = name
        self.HeatFlow = ValueHolder(heat_flow_kw)


class NamedContainer:
    def __init__(self, mapping: dict[str, object]):
        self._mapping = mapping

    @property
    def Names(self) -> list[str]:
        return list(self._mapping.keys())

    def Item(self, name: str) -> object:
        return self._mapping[name]

    def __iter__(self):
        return iter(self._mapping.values())


class FakeTraySection:
    TypeName = "traysection"

    def __init__(self, name: str, number_of_trays: int):
        self.Name = name
        self.NumberOfTrays = number_of_trays
        self.last_feed_location: tuple[str, int] | None = None

    def SpecifyFeedLocation(self, stream: FakeMaterialStream, tray_number: int) -> None:
        self.last_feed_location = (stream.Name, tray_number)


class FakeColumnFlowsheet:
    def __init__(self, feed_stream_name: str, tray_section: FakeTraySection):
        self.FeedStreams = NamedContainer({feed_stream_name: FakeMaterialStream(feed_stream_name)})
        self.Operations = [tray_section]
        self.run_count = 0

    def Run(self) -> None:
        self.run_count += 1


class FakeColumnOperation:
    def __init__(self, name: str, tray_count: int, feed_stream_name: str):
        self.Name = name
        self.ColumnFlowsheet = FakeColumnFlowsheet(feed_stream_name, FakeTraySection(f"{name}_tray_section", tray_count))


class FakeReactorOperation:
    def __init__(self, total_volume_value: float):
        self.TotalVolumeValue = total_volume_value


class FakeSplitterOperation:
    def __init__(self, split_first: float):
        self.SplitsValue = (split_first, 1.0 - split_first)


class FakeOperations:
    def __init__(self, mapping: dict[str, object]):
        self._mapping = mapping

    def Item(self, name: str) -> object:
        return self._mapping[name]


@dataclass
class FakeFlowsheet:
    MaterialStreams: NamedContainer
    EnergyStreams: NamedContainer
    Operations: FakeOperations


@dataclass
class FakeCase:
    Title: str
    Flowsheet: FakeFlowsheet


def build_mock_context() -> HysysContext:
    material_streams = NamedContainer(
        {
            "Fresh H2 Feed": FakeMaterialStream("Fresh H2 Feed", molar_flow=190.0 / 3600.0),
            "S5": FakeMaterialStream("S5", temperature=630.0, pressure=3200.0),
            "S6": FakeMaterialStream("S6", temperature=670.0, pressure=3150.0),
            "S16": FakeMaterialStream("S16", pressure=3250.0),
            "S23": FakeMaterialStream("S23", pressure=930.0),
            "S26": FakeMaterialStream("S26", pressure=110.0),
            "S30": FakeMaterialStream("S30", pressure=180.0),
            "T2_feed": FakeMaterialStream("T2_feed"),
            "T3_feed": FakeMaterialStream("T3_feed"),
        }
    )
    energy_streams = NamedContainer(
        {
            "Heater1 HP": FakeEnergyStream("Heater1 HP", 120.0),
            "Heater2 FH": FakeEnergyStream("Heater2 FH", 240.0),
            "Cooler1 CW": FakeEnergyStream("Cooler1 CW", -60.0),
            "Column Condenser Ref": FakeEnergyStream("Column Condenser Ref", -30.0),
            "Compressor Elec": FakeEnergyStream("Compressor Elec", -18.0),
        }
    )
    operations = FakeOperations(
        {
            "Reactor1": FakeReactorOperation(160.0),
            "Sep1": FakeSplitterOperation(0.05),
            "Sep2": FakeSplitterOperation(0.80),
            "T2": FakeColumnOperation("T2", tray_count=41, feed_stream_name="T2_feed"),
            "T3": FakeColumnOperation("T3", tray_count=25, feed_stream_name="T3_feed"),
        }
    )
    flowsheet = FakeFlowsheet(
        MaterialStreams=material_streams,
        EnergyStreams=energy_streams,
        Operations=operations,
    )
    case = FakeCase(Title="Mock HDA Case", Flowsheet=flowsheet)
    return HysysContext(hysys=None, case=case, flowsheet=flowsheet)
