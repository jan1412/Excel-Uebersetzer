"""Fachliche Datenklassen. Hier kannst du eigene Parameter/Strukturen ergänzen."""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import TypeAlias

Scalar: TypeAlias = Decimal | str | bool | None
Matrix: TypeAlias = tuple[tuple[Scalar, ...], ...]


# region M01: Zellherkunft und verständliche Fehler
@dataclass(frozen=True)
class Source:
    sheet: str
    address: str

    def __str__(self) -> str:
        return f"{self.sheet}!{self.address}"


class TranslationError(ValueError):
    """Erwarteter Eingabe-/Zuordnungsfehler mit verständlicher Nachricht."""
# endregion


# region M02: Einzelwerte für den Header
@dataclass
class Metadata:
    parameters: dict[str, Scalar] = field(default_factory=dict)
# endregion


# region M03: Materialblock; jeder Parameter enthält genau einen Wert
@dataclass
class MaterialBlock:
    reference: str
    material: str
    parameters: dict[str, Scalar]
    source: Source | None = None
# endregion


# region M04: Einstellungsblock; jeder Parameter darf mehrere Werte besitzen
@dataclass
class SettingBlock:
    reference: str
    name: str
    parameters: dict[str, tuple[Scalar, ...]]
    source: Source | None = None
# endregion


# region M05: Maschinenmatrix; reference stammt aus lokaler Zelle (1, 1)
@dataclass
class MachineConfiguration:
    reference: str
    matrix: Matrix
    source: Source | None = None
# endregion


# region M06: Verweis auf einen Matrixabschnitt; Indizes 1-basiert, inklusive
@dataclass(frozen=True)
class MatrixReference:
    configuration: str
    # None bedeutet: gesamte Datenmatrix (ohne Beschriftungszeilen/-spalten).
    bounds: tuple[int, int, int, int] | None = None
    source: Source | None = None
# endregion


# region M07: Excel-Step vor dem Auflösen seiner Verweise
@dataclass
class Step:
    reference: str
    name: str
    parameters: dict[str, Scalar] = field(default_factory=dict)
    material_references: tuple[str, ...] = ()
    setting_references: tuple[str, ...] = ()
    matrix_references: tuple[MatrixReference, ...] = ()
    source: Source | None = None
# endregion


# region M08: Step nach dem Auflösen; Namensräume verhindern Überschreiben
@dataclass
class ResolvedMatrix:
    reference: MatrixReference
    values: Matrix


@dataclass
class ResolvedStep:
    step: Step
    materials: tuple[MaterialBlock, ...]
    settings: tuple[SettingBlock, ...]
    matrices: tuple[ResolvedMatrix, ...]
# endregion


# region M09: Vollständig eingelesenes Projekt
@dataclass
class Project:
    metadata: Metadata
    materials: dict[str, MaterialBlock]
    settings: dict[str, SettingBlock]
    machines: dict[str, MachineConfiguration]
    steps: list[Step]
# endregion
