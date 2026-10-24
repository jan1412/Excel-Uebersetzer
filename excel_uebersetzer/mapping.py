"""DEINE ANPASSUNGSSTELLE: Beispieladressen durch echte Zellbezüge ersetzen.

Absolute Excel-Zeilen/-Spalten beginnen bei 1, Offset-Werte bei 0.
Jeder BAUSTEIN ist unabhängig anpassbar. Siehe docs/CFG_BEISPIEL.md.
"""
from .layout import (
    Blocks, Cell, MachinesLayout, MaterialsLayout, MatrixSlot, Offset,
    Rectangle, SettingsLayout, StepsLayout, WorkbookLayout,
)


# region C01: Metadata -> Header
METADATA_SHEET = "Metadata"
METADATA = {
    "project_name": Cell("B2", "text", True),
    "author": Cell("B3", "text"),
    "version": Cell("B4", "text"),
}
# endregion


# region C02: Materialien; A2:C5, A6:C9, ... (Blockhöhe 4)
MATERIALS = MaterialsLayout(
    blocks=Blocks("Materialien", first_row=2, first_column=1, height=4),
    material=Offset(0, 1, "text", True),
    parameters={
        "density": Offset(1, 1, "number", True),
        "thickness": Offset(2, 1, "number"),
        "note": Offset(3, 1, "text"),
    },
)
# endregion


# region C03: Maschinen; A2:D5, A6:D9, ...; ID links oben = lokale (1, 1)
MACHINES = MachinesLayout(
    blocks=Blocks("Maschinenkonfiguration", 2, 1, 4),
    # Daten B3:D5 relativ zu A2; Überschriften gehören nicht zur Datenmatrix.
    data=Rectangle(row=1, column=1, rows=3, columns=3),
)
# endregion


# region C04: Einstellungen; A2:C5, A6:C9, ...
SETTINGS = SettingsLayout(
    blocks=Blocks("Einstellungen", 2, 1, 4),
    name=Offset(0, 1, "text", True),
    parameters={
        "speed_rpm": (Offset(1, 1, "number", True),),
        "duration_s": (Offset(2, 1, "number", True),),
        "limits": (Offset(3, 1, "number"), Offset(3, 2, "number")),
    },
)
# endregion


# region C05: Ablauferstellung; A2:E7, A8:E13, ... (Blockhöhe 6)
STEPS = StepsLayout(
    blocks=Blocks("Ablauferstellung", 2, 1, 6),
    name=Offset(0, 1, "text", True),
    material_references=Offset(1, 1, "text"),
    setting_references=Offset(2, 1, "text"),
    parameters={"feed": Offset(5, 1, "number")},
    matrix_slots=(
        MatrixSlot(
            reference=Offset(3, 1, "text"),
            bounds=(Offset(4, 1, "number"), Offset(4, 2, "number"),
                    Offset(4, 3, "number"), Offset(4, 4, "number")),
        ),
    ),
)
# endregion


# region C06: Einzelne Bausteine zur vollständigen Zuordnung verbinden
LAYOUT = WorkbookLayout(
    metadata_sheet=METADATA_SHEET,
    metadata=METADATA,
    materials=MATERIALS,
    machines=MACHINES,
    settings=SETTINGS,
    steps=STEPS,
)
# endregion
