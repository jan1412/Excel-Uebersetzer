"""Native Dateidialoge. Tkinter wird erst bei tatsächlicher GUI-Nutzung geladen."""
from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

from .models import TranslationError


# region D00: Gemeinsames unsichtbares Tkinter-Elternfenster
@contextmanager
def dialog_root():
    """Fehlendes Tkinter/Display verständlich melden; Fenster immer schließen."""
    try:
        import tkinter as tk
    except ImportError as exc:
        raise TranslationError("Tkinter fehlt. Alternativ mit Eingabe- und Ausgabepfad in der Konsole starten.") from exc
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        raise TranslationError("Kein grafischer Dateidialog verfügbar. Alternativ: python main.py DATEI.xlsx --output AUSGABE.txt") from exc
    try:
        root.withdraw()
        yield root
    finally:
        root.destroy()
# endregion


# region D01: Excel-Datei im Explorer-/Dateiauswahlfenster auswählen
def choose_excel() -> Path | None:
    """Abbrechen liefert None. Keine Makroausführung, keine Excel-Automation."""
    with dialog_root() as root:
        from tkinter import filedialog
        filename = filedialog.askopenfilename(
            parent=root, title="Excel-Datei für den Übersetzer auswählen",
            filetypes=[("Excel-Arbeitsmappen", "*.xlsx *.xlsm")],
        )
        return Path(filename) if filename else None
# endregion


# region D02: Zielpfad und Dateiendung frei auswählen
def choose_output(input_path: Path, format_name: str = "txt") -> Path | None:
    """Vorlagenset und Dateiendung sind getrennt; alle Dateien sind erlaubt."""
    with dialog_root() as root:
        from tkinter import filedialog
        filename = filedialog.asksaveasfilename(
            parent=root, title="Erzeugte Datei speichern", initialdir=str(input_path.parent),
            initialfile=f"{input_path.stem}_export.{format_name}", defaultextension=f".{format_name}",
            filetypes=[("Textdatei", "*.txt"), ("Python-Skript", "*.py"), ("Alle Dateien", "*.*")],
        )
        return Path(filename) if filename else None
# endregion


# region D03: Ergebnis oder Fehler als kleines Infofenster anzeigen
def show_message(title: str, message: str, error: bool = False) -> None:
    """Nur die Anwendung benutzt diesen Baustein; die Fachmodule benötigen keine GUI."""
    with dialog_root() as root:
        from tkinter import messagebox
        (messagebox.showerror if error else messagebox.showinfo)(title, message, parent=root)
# endregion
