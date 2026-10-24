"""Textbausteine mit string.Template zusammensetzen, ohne weitere Bibliothek."""
from __future__ import annotations

from collections.abc import Iterator
from decimal import Decimal
from pathlib import Path
from string import Template
import os
import tempfile

from .models import Metadata, ResolvedStep, Scalar, TranslationError
from .values import format_number

DEFAULT_TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "templates"


# region T01: Einzelwert für lesbare Textausgabe formatieren
def text_value(value: Scalar) -> str:
    """None -> ka; echte Zahlen -> Punktformat; Text bleibt Text."""
    if value is None:
        return "ka"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (Decimal, int, float)):
        return format_number(value)
    return str(value)
# endregion


# region T02: Sichere Python-Literale für Werte und verschachtelte Strukturen
def python_literal(value: object) -> str:
    """Maskiert Texte mit repr(); führt niemals Excel-Inhalte aus.

    Decimal wird als Python-Zahlenliteral ausgegeben. Python speichert diese
    Zahlen im generierten Skript als float (siehe Dokumentation zur Präzision).
    """
    if value is None or isinstance(value, bool):
        return repr(value)
    if isinstance(value, (Decimal, int, float)):
        return format_number(value)
    if isinstance(value, str):
        return repr(value)
    if isinstance(value, dict):
        return "{" + ", ".join(f"{python_literal(k)}: {python_literal(v)}" for k, v in value.items()) + "}"
    if isinstance(value, (list, tuple)):
        return "[" + ", ".join(python_literal(item) for item in value) + "]"
    raise TranslationError(f"Nicht unterstützter Python-Ausgabewert: {type(value).__name__}.")
# endregion


# region T03: Einfache Parameterzeilen für .txt erzeugen
def parameter_lines(parameters: dict[str, Scalar], indent: str = "") -> str:
    """Die Reihenfolge entspricht der Einfügereihenfolge der Parameter."""
    return "\n".join(f"{indent}{key} = {text_value(value)}" for key, value in parameters.items())
# endregion


# region T04: Aufgelösten Step in einen neutralen Ausgabedatensatz umwandeln
def step_payload(item: ResolvedStep) -> dict:
    """HIER eigene Ausgabeparameter/Fachnamen ergänzen."""
    return {
        "id": item.step.reference,
        "name": item.step.name,
        "parameters": item.step.parameters,
        "materials": {block.reference: {"material": block.material, "parameters": block.parameters}
                      for block in item.materials},
        "settings": {block.reference: {"name": block.name, "parameters": block.parameters}
                     for block in item.settings},
        "matrices": [{"configuration": block.reference.configuration,
                      "bounds": block.reference.bounds, "values": block.values}
                     for block in item.matrices],
    }
# endregion


# region T05: Templatevariablen eines einzelnen Steps bereitstellen
def step_context(item: ResolvedStep, index: int) -> dict[str, str]:
    """Lesbare .txt-Variablen und maskierte *_literal-Variablen für Python."""
    payload = step_payload(item)
    material_parts, setting_parts, matrix_parts = [], [], []
    for block in item.materials:
        material_parts.append(f"  {block.reference} ({block.material})\n" + parameter_lines(block.parameters, "    "))
    for block in item.settings:
        rows = [f"    {key} = [" + ", ".join(text_value(value) for value in values) + "]"
                for key, values in block.parameters.items()]
        setting_parts.append(f"  {block.reference} ({block.name})\n" + "\n".join(rows))
    for block in item.matrices:
        rows = ["    " + "  ".join(text_value(value) for value in row) for row in block.values]
        matrix_parts.append(f"  {block.reference.configuration}\n" + "\n".join(rows))
    context = {
        "step_index": str(index), "step_id": item.step.reference, "step_name": item.step.name,
        "parameters_text": parameter_lines(item.step.parameters, "  "),
        "materials_text": "\n".join(material_parts) or "  ka",
        "settings_text": "\n".join(setting_parts) or "  ka",
        "matrices_text": "\n".join(matrix_parts) or "  ka",
        "step_id_literal": python_literal(item.step.reference),
        "step_name_literal": python_literal(item.step.name),
        "step_literal": python_literal(payload),
    }
    context.update({f"{key}_literal": python_literal(payload[key])
                    for key in ("parameters", "materials", "settings", "matrices")})
    return context
# endregion


# region T06: Einzelnen Textbaustein strikt einsetzen
def render_template(path: Path, context: dict[str, str]) -> str:
    """${name} ersetzen; fehlende Variablen/fehlerhafte $-Syntax melden."""
    try:
        return Template(path.read_text(encoding="utf-8")).substitute(context)
    except (KeyError, ValueError) as exc:
        raise TranslationError(f"Vorlage {path.name}: {exc}. Literalzeichen $ als $$ schreiben.") from exc
# endregion


# region T07: Header, alle Steps und Footer nacheinander erzeugen
def document_chunks(
    metadata: Metadata, steps: list[ResolvedStep], format_name: str = "txt",
    template_dir: Path = DEFAULT_TEMPLATE_DIR,
) -> Iterator[str]:
    """Eine Hauptvorlage pro Teil; hier zusätzliche Step-Typen verzweigen."""
    if format_name not in {"txt", "py"}:
        raise TranslationError(f"Unbekanntes Vorlagenset {format_name!r}; txt oder py verwenden.")
    template_dir = Path(template_dir)
    common = {"metadata_text": parameter_lines(metadata.parameters),
              "metadata_literal": python_literal(metadata.parameters), "step_count": str(len(steps))}
    yield render_template(template_dir / f"header.{format_name}.tpl", common)
    for index, item in enumerate(steps, start=1):
        yield render_template(template_dir / f"step.{format_name}.tpl", {**common, **step_context(item, index)})
    yield render_template(template_dir / f"footer.{format_name}.tpl", common)
# endregion


# region T08: Vollständige Datei schreiben; .py vor Übernahme syntaktisch prüfen
def write_document(output: str | Path, chunks: Iterator[str], format_name: str = "txt") -> Path:
    """UTF-8, LF, temporäre Datei; Fehler hinterlassen keine halbe Zieldatei.

    Bestehende Zieldateien werden nach erfolgreichem Erstellen ersetzt.
    Die Syntaxprüfung verwendet den Python-Interpreter des Generators.
    """
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", newline="\n",
                                         dir=output.parent, prefix=".excel-export-", delete=False) as handle:
            temporary = Path(handle.name)
            for chunk in chunks:
                handle.write(chunk)
        if format_name == "py":
            compile(temporary.read_text(encoding="utf-8"), str(output), "exec")
        os.replace(temporary, output)
    except SyntaxError as exc:
        raise TranslationError(f"Python-Ausgabe ungültig, Zeile {exc.lineno}: {exc.msg}") from exc
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return output
# endregion
