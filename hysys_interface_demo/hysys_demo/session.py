from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class HysysContext:
    hysys: Any
    case: Any
    flowsheet: Any


def connect_to_active_case() -> HysysContext:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "win32com.client is required to connect to Aspen HYSYS."
        ) from exc

    hysys = win32com.client.GetObject(None, "HYSYS.Application")
    case = hysys.ActiveDocument
    flowsheet = case.Flowsheet
    return HysysContext(hysys=hysys, case=case, flowsheet=flowsheet)


def ensure_context(context: HysysContext | None = None) -> HysysContext:
    return context if context is not None else connect_to_active_case()
