"""Die Bausteine zu einem Ablauf verbinden; auch aus eigenen Skripten nutzbar."""
from __future__ import annotations

from pathlib import Path

from .excel_reader import open_excel
from .layout import WorkbookLayout
from .loaders import read_project
from .mapping import LAYOUT
from .models import Project, ResolvedStep, TranslationError
from .renderer import DEFAULT_TEMPLATE_DIR, document_chunks, write_document
from .resolver import resolve_steps, validate_project


# region P01: Excel einlesen, IDs auflösen, fachliche Prüfung aufrufen
def load_and_resolve(
    input_path: str | Path, layout: WorkbookLayout = LAYOUT,
) -> tuple[Project, list[ResolvedStep]]:
    """Nur Daten lesen/verarbeiten; erzeugt keine Ausgabe."""
    with open_excel(input_path) as workbook:
        project = read_project(workbook, layout)
    steps = resolve_steps(project)
    validate_project(project, steps)
    return project, steps
# endregion


# region P02: Vollständige Übersetzung ohne GUI
def translate(
    input_path: str | Path, output_path: str | Path, format_name: str = "txt",
    layout: WorkbookLayout = LAYOUT, template_dir: Path = DEFAULT_TEMPLATE_DIR,
) -> Path:
    """Ein einfacher Funktionsaufruf für eigene Programme."""
    if Path(input_path).resolve() == Path(output_path).resolve():
        raise TranslationError("Die Ausgabedatei darf nicht die Excel-Eingabedatei sein.")
    project, steps = load_and_resolve(input_path, layout)
    return write_document(output_path, document_chunks(project.metadata, steps, format_name, template_dir), format_name)
# endregion
