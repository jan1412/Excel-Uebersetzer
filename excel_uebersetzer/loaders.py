"""Ein Baustein pro Sheet-Typ. Eigene Excel-Leser hier ergänzen/ersetzen."""
from __future__ import annotations

from typing import Any

from .excel_reader import add_unique, block_address, iter_block_rows, read_cell, read_offset, require_sheets
from .layout import Cell, MachinesLayout, MaterialsLayout, Offset, SettingsLayout, StepsLayout, WorkbookLayout
from .models import MaterialBlock, MachineConfiguration, MatrixReference, Metadata, Project, SettingBlock, Source, Step, TranslationError
from .values import positive_index, split_references


# region S01: Metadata als einzelne Headerparameter lesen
def read_metadata(workbook: Any, sheet_name: str, fields: dict[str, Cell]) -> Metadata:
    """Kein Blockraster; jedes Feld besitzt eine absolute Zelladresse."""
    return Metadata({name: read_cell(workbook[sheet_name], spec) for name, spec in fields.items()})
# endregion


# region S02: Wiederholte Materialblöcke lesen
def read_materials(workbook: Any, layout: MaterialsLayout) -> dict[str, MaterialBlock]:
    """Mehrere Blöcke dürfen dasselbe Material, aber keine doppelte Referenz haben."""
    sheet, blocks, result = workbook[layout.blocks.sheet], layout.blocks, {}
    for origin in iter_block_rows(sheet, blocks, [layout.material, *layout.parameters.values()]):
        source = Source(sheet.title, block_address(blocks, origin, blocks.reference))
        read = lambda spec: read_offset(sheet, blocks, origin, spec)
        item = MaterialBlock(read(blocks.reference), read(layout.material),
                             {key: read(spec) for key, spec in layout.parameters.items()}, source)
        add_unique(result, item)
    return result
# endregion


# region S03: Einstellungsgruppen mit mehrteiligen Parametern lesen
def read_settings(workbook: Any, layout: SettingsLayout) -> dict[str, SettingBlock]:
    """Wertepositionen erhalten: (0.00, None, 1.00) wird nicht verdichtet."""
    sheet, blocks, result = workbook[layout.blocks.sheet], layout.blocks, {}
    fields = [layout.name, *(spec for group in layout.parameters.values() for spec in group)]
    for origin in iter_block_rows(sheet, blocks, fields):
        source = Source(sheet.title, block_address(blocks, origin, blocks.reference))
        read = lambda spec: read_offset(sheet, blocks, origin, spec)
        item = SettingBlock(read(blocks.reference), read(layout.name),
                            {key: tuple(read(spec) for spec in group) for key, group in layout.parameters.items()}, source)
        add_unique(result, item)
    return result
# endregion


# region S04: Maschinenmatrizen lesen; Beschriftungsränder bleiben außen vor
def read_machines(workbook: Any, layout: MachinesLayout) -> dict[str, MachineConfiguration]:
    """Matrixgröße/-position wird ausschließlich durch layout.data definiert."""
    sheet, blocks, result = workbook[layout.blocks.sheet], layout.blocks, {}
    rect = layout.data
    rows = tuple(tuple(Offset(rect.row + r, rect.column + c, rect.kind)
                       for c in range(rect.columns)) for r in range(rect.rows))
    for origin in iter_block_rows(sheet, blocks, [cell for row in rows for cell in row]):
        source = Source(sheet.title, block_address(blocks, origin, blocks.reference))
        matrix = tuple(tuple(read_offset(sheet, blocks, origin, cell) for cell in row) for row in rows)
        item = MachineConfiguration(read_offset(sheet, blocks, origin, blocks.reference), matrix, source)
        add_unique(result, item)
    return result
# endregion


# region S05: Step-Liste mit direkten Parametern und ungelösten IDs lesen
def read_steps(workbook: Any, layout: StepsLayout) -> list[Step]:
    """Physische Blockreihenfolge = spätere Ausgabereihenfolge."""
    sheet, blocks, result = workbook[layout.blocks.sheet], layout.blocks, {}
    fields = [layout.name, *layout.parameters.values()]
    fields.extend(spec for spec in (layout.material_references, layout.setting_references) if spec)
    for slot in layout.matrix_slots:
        fields.append(slot.reference)
        fields.extend(slot.bounds or ())
    for origin in iter_block_rows(sheet, blocks, fields):
        source = Source(sheet.title, block_address(blocks, origin, blocks.reference))
        read = lambda spec: read_offset(sheet, blocks, origin, spec)
        matrices = []
        for slot in layout.matrix_slots:
            reference, bounds = read(slot.reference), None
            values = tuple(read(spec) for spec in slot.bounds) if slot.bounds else ()
            slot_source = Source(sheet.title, block_address(blocks, origin, slot.reference))
            if reference is None:
                if any(value is not None for value in values):
                    raise TranslationError(f"{slot_source}: Matrixgrenzen ohne Konfigurations-ID.")
                continue
            if values and any(value is not None for value in values):
                if any(value is None for value in values):
                    raise TranslationError(f"{slot_source}: Alle vier Matrixgrenzen angeben oder alle leer lassen.")
                bounds = tuple(positive_index(value, slot_source) for value in values)
            matrices.append(MatrixReference(reference, bounds, slot_source))
        item = Step(
            reference=read(blocks.reference), name=read(layout.name),
            parameters={key: read(spec) for key, spec in layout.parameters.items()},
            material_references=split_references(read(layout.material_references), source) if layout.material_references else (),
            setting_references=split_references(read(layout.setting_references), source) if layout.setting_references else (),
            matrix_references=tuple(matrices), source=source,
        )
        add_unique(result, item)
    return list(result.values())
# endregion


# region S06: Alle Sheet-Leser zu einem Projekt verbinden
def read_project(workbook: Any, layout: WorkbookLayout) -> Project:
    """Zentrale Austauschstelle, falls du einzelne Sheet-Leser ersetzen möchtest."""
    require_sheets(workbook, [layout.metadata_sheet, layout.materials.blocks.sheet,
                            layout.machines.blocks.sheet, layout.settings.blocks.sheet,
                            layout.steps.blocks.sheet])
    return Project(
        metadata=read_metadata(workbook, layout.metadata_sheet, layout.metadata),
        materials=read_materials(workbook, layout.materials),
        settings=read_settings(workbook, layout.settings),
        machines=read_machines(workbook, layout.machines),
        steps=read_steps(workbook, layout.steps),
    )
# endregion
