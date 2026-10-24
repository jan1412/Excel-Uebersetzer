"""Optionale Adapter für Zell-Dropdowns und verknüpfte Formularsteuerelemente.

Die normalen Sheet-Leser benötigen nur gespeicherte IDs. Diese Bausteine sind
für eigene Leser gedacht, wenn eine verknüpfte Zelle stattdessen einen Index
enthält. Die Dropdown-/ActiveX-Objekte selbst werden nicht untersucht.
"""
from __future__ import annotations

from typing import Any, Sequence

from openpyxl.utils.cell import range_boundaries, get_column_letter

from .excel_reader import read_cell
from .layout import Cell
from .models import Scalar, Source, TranslationError
from .values import positive_index


# region Q01: Gespeicherte ID oder verknüpften 1-basierten Listenindex auflösen
def selection_to_id(
    value: Scalar, options: Sequence[str | None] | None = None,
    source: Source | str = "Dropdown",
) -> str | None:
    """options=None: Wert ist ID. Mit options: Wert ist Listenposition."""
    if value is None:
        return None
    if options is None:
        if not isinstance(value, str):
            raise TranslationError(f"{source}: ID-Text erwartet; für einen Listenindex options angeben.")
        return value
    index = positive_index(value, source)
    if index > len(options) or options[index - 1] is None:
        raise TranslationError(f"{source}: Listenindex {index} hat keine gültige ID.")
    return options[index - 1]
# endregion


# region Q02: Verknüpfte Auswahlzelle und optional ihre Quellliste lesen
def read_dropdown_id(
    sheet: Any, address: str, options_sheet: Any | None = None,
    options_range: str | None = None,
) -> str | None:
    """Quellbereich muss genau eine Spalte haben; Leerplätze behalten ihren Index.

    Ein normales Zell-Dropdown braucht nur sheet und address. Für ein
    Formular-Kombinationsfeld gib dessen verknüpfte Zelle und Quellliste an.
    """
    value = read_cell(sheet, Cell(address))
    options = None
    if (options_sheet is None) != (options_range is None):
        raise TranslationError("options_sheet und options_range gemeinsam angeben.")
    if options_sheet is not None:
        try:
            first_column, first_row, last_column, last_row = range_boundaries(options_range)
        except ValueError as exc:
            raise TranslationError("Dropdown-Quelle als Zellbereich angeben, z.B. A2:A20.") from exc
        if first_column != last_column or None in (first_column, first_row, last_column, last_row):
            raise TranslationError("Die Dropdown-Quellliste muss genau eine Spalte umfassen.")
        column = get_column_letter(first_column)
        options = [read_cell(options_sheet, Cell(f"{column}{row}", "text"))
                   for row in range(first_row, last_row + 1)]
    return selection_to_id(value, options, Source(sheet.title, address))
# endregion
