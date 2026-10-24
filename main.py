"""Start: python main.py. Ohne Eingabepfad öffnet sich die Dateiauswahl."""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import BadZipFile

from excel_uebersetzer.dialogs import choose_excel, choose_output, show_message
from excel_uebersetzer.models import TranslationError
from excel_uebersetzer.pipeline import load_and_resolve, translate
from excel_uebersetzer.renderer import DEFAULT_TEMPLATE_DIR


# region A01: Kommandozeilenoptionen vorbereiten
def parse_arguments(argv: list[str] | None = None) -> argparse.Namespace:
    """Dateipfade optional; --check arbeitet ohne Ausgabedatei."""
    parser = argparse.ArgumentParser(description="Excel-Übersetzer: anpassbares Grundgerüst")
    parser.add_argument("input", nargs="?", type=Path, help="Aktuelle, gespeicherte .xlsx/.xlsm")
    parser.add_argument("--output", "-o", type=Path, help="Zieldatei; ohne Angabe Speicherdialog")
    parser.add_argument("--format", choices=("txt", "py"), help="Vorlagenset; Standard anhand .py-Endung, sonst txt")
    parser.add_argument("--templates", type=Path, default=DEFAULT_TEMPLATE_DIR, help="Eigener Vorlagenordner")
    parser.add_argument("--check", action="store_true", help="Nur einlesen und Verweise prüfen")
    return parser.parse_args(argv)
# endregion


# region A02: Dialoge und Übersetzungsbausteine verbinden
def main(argv: list[str] | None = None) -> int:
    """Rückgabewert: 0 erfolgreich/abgebrochen, 1 Fehler."""
    args = parse_arguments(argv)
    used_gui = args.input is None or (args.output is None and not args.check)
    try:
        input_path = args.input or choose_excel()
        if input_path is None:
            return 0
        if args.check:
            project, steps = load_and_resolve(input_path)
            message = (f"Eingelesen: {len(project.materials)} Materialblöcke, "
                       f"{len(project.settings)} Einstellungsblöcke, "
                       f"{len(project.machines)} Maschinenkonfigurationen, {len(steps)} Steps.")
        else:
            format_name = args.format or ("py" if args.output and args.output.suffix.lower() == ".py" else "txt")
            output_path = args.output or choose_output(input_path, format_name)
            if output_path is None:
                return 0
            # Im Speicherdialog gewählte .py-Endung ebenfalls berücksichtigen.
            if args.format is None:
                format_name = "py" if output_path.suffix.lower() == ".py" else "txt"
            result = translate(input_path, output_path, format_name, template_dir=args.templates)
            message = f"Datei erstellt: {result.resolve()}"
        print(message)
        if used_gui:
            show_message("Excel-Übersetzer", message)
        return 0
    except (TranslationError, OSError, BadZipFile) as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        if used_gui:
            try:
                show_message("Excel-Übersetzer – Fehler", str(exc), error=True)
            except Exception:
                # Keine GUI verfügbar: Die ursprüngliche Konsolenmeldung bleibt erhalten.
                pass
        return 1
# endregion


if __name__ == "__main__":
    raise SystemExit(main())
