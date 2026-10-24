"""Kleine Zuordnungsbausteine. Tatsächliche Adressen stehen in mapping.py."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

Kind = Literal["auto", "number", "text", "bool"]


# region L01: Absolute Einzelzelle
@dataclass(frozen=True)
class Cell:
    address: str
    kind: Kind = "auto"
    required: bool = False
# endregion


# region L02: Relative Zelle; Offsets beginnen bei 0
@dataclass(frozen=True)
class Offset:
    row: int
    column: int
    kind: Kind = "auto"
    required: bool = False
# endregion


# region L03: Wiederholte Blöcke mit festem Zeilenabstand
@dataclass(frozen=True)
class Blocks:
    sheet: str
    first_row: int
    first_column: int
    height: int
    reference: Offset = Offset(0, 0, "text", True)
    last_row: int | None = None  # None: bis zur letzten verwendeten Excel-Zeile.
    # Optional: feste Beschriftungen nur in tatsächlich belegten Blöcken prüfen.
    labels: dict[tuple[int, int], str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if min(self.first_row, self.first_column, self.height) < 1:
            raise ValueError("Startzeile, Startspalte und Blockhöhe müssen positiv sein.")
        if self.last_row is not None and self.last_row < self.first_row:
            raise ValueError("last_row muss mindestens first_row entsprechen.")
# endregion


# region L04: Materialblock mit Einzelwerten
@dataclass(frozen=True)
class MaterialsLayout:
    blocks: Blocks
    material: Offset
    parameters: dict[str, Offset]
# endregion


# region L05: Einstellungsblock mit einem oder mehreren Werten pro Schlüssel
@dataclass(frozen=True)
class SettingsLayout:
    blocks: Blocks
    name: Offset
    parameters: dict[str, tuple[Offset, ...]]
# endregion


# region L06: Datenrechteck innerhalb eines Maschinenblocks
@dataclass(frozen=True)
class Rectangle:
    row: int
    column: int
    rows: int
    columns: int
    kind: Kind = "number"

    def __post_init__(self) -> None:
        if min(self.row, self.column) < 0 or min(self.rows, self.columns) < 1:
            raise ValueError("Matrixoffsets müssen >= 0 und Dimensionen >= 1 sein.")


@dataclass(frozen=True)
class MachinesLayout:
    blocks: Blocks
    data: Rectangle
# endregion


# region L07: Eine Matrixauswahl innerhalb eines Step-Blocks
@dataclass(frozen=True)
class MatrixSlot:
    reference: Offset
    # Reihenfolge: Startzeile, Startspalte, Endzeile, Endspalte.
    bounds: tuple[Offset, Offset, Offset, Offset] | None = None
# endregion


# region L08: Step-Block mit direkten Parametern und beliebig vielen Referenzen
@dataclass(frozen=True)
class StepsLayout:
    blocks: Blocks
    name: Offset
    parameters: dict[str, Offset]
    material_references: Offset | None = None
    setting_references: Offset | None = None
    matrix_slots: tuple[MatrixSlot, ...] = ()
# endregion


# region L09: Gesamte Excel-Zuordnung
@dataclass(frozen=True)
class WorkbookLayout:
    metadata_sheet: str
    metadata: dict[str, Cell]
    materials: MaterialsLayout
    machines: MachinesLayout
    settings: SettingsLayout
    steps: StepsLayout
# endregion
