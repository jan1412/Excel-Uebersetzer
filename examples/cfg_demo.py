"""Ohne Excel-Datei: Datenklassen, Verweise und beide Vorlagensets ausprobieren.

Start aus dem Projektordner: python -m examples.cfg_demo
Die Datei cfg_cells.json dokumentiert dazu die beispielhaften Excel-Zellen.
"""
from decimal import Decimal as D
from pathlib import Path

from excel_uebersetzer.models import (
    MachineConfiguration, MaterialBlock, MatrixReference, Metadata, Project,
    SettingBlock, Step,
)
from excel_uebersetzer.renderer import document_chunks, write_document
from excel_uebersetzer.resolver import resolve_steps, validate_project


# region B01: Beispielprojekt direkt aus Datenklassen zusammensetzen
def build_demo_project() -> Project:
    """Entspricht den Zellbeispielen in cfg_cells.json und CFG_BEISPIEL.md."""
    return Project(
        metadata=Metadata({"project_name": "CFG-Beispiel", "author": "Beispiel", "version": "1.0"}),
        materials={
            "MAT_STEEL_001": MaterialBlock("MAT_STEEL_001", "Stahl", {"density": D("7.85"), "thickness": D("1.0"), "note": None}),
            "MAT_STEEL_002": MaterialBlock("MAT_STEEL_002", "Stahl", {"density": D("7.85"), "thickness": D("2.0"), "note": "Zweiter Block desselben Materials"}),
        },
        settings={
            "CFG001": SettingBlock("CFG001", "Standard", {"speed_rpm": (D("1500"),), "duration_s": (D("10"),), "limits": (D("0"), D("1"))}),
            "CFG002": SettingBlock("CFG002", "Schonend", {"speed_rpm": (D("900"),), "duration_s": (D("20"),), "limits": (D("0.1"), D("0.5"))}),
        },
        machines={"MCH001": MachineConfiguration("MCH001", tuple(
            tuple(D(str(value)) for value in row) for row in ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        ))},
        steps=[
            Step("STEP001", "Bearbeiten", {"feed": D("0")}, ("MAT_STEEL_001",), ("CFG001",),
                 (MatrixReference("MCH001", (1, 2, 2, 3)),)),
            Step("STEP002", "Finish", {"feed": None}, ("MAT_STEEL_001", "MAT_STEEL_002"), ("CFG002",),
                 (MatrixReference("MCH001"),)),
        ],
    )
# endregion


# region B02: Beispielprojekt auflösen und beide Formate erzeugen
def main() -> None:
    """Die Demo ersetzt nur das Excel-Einlesen; alle Folgeschritte sind identisch."""
    project = build_demo_project()
    steps = resolve_steps(project)
    validate_project(project, steps)
    for format_name in ("txt", "py"):
        path = Path("output") / f"cfg_beispiel.{format_name}"
        write_document(path, document_chunks(project.metadata, steps, format_name), format_name)
        print(path.resolve())
# endregion


if __name__ == "__main__":
    main()
