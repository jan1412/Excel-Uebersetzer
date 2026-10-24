"""Fehlende Werte, Zahlen und Referenzlisten an einer Stelle behandeln."""
from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
import re

from .layout import Kind
from .models import Scalar, Source, TranslationError


# region V01: Fehlend erkennen; 0 und False bleiben gültig
def is_missing(value: object) -> bool:
    """None, leere/weiße Texte und 'ka' (ohne Beachtung der Schreibung)."""
    return value is None or (
        isinstance(value, str) and value.strip().casefold() in {"", "ka"}
    )
# endregion


# region V02: Dezimalzahl lesen; keine lokale Komma-Konvertierung
def to_decimal(value: object, source: Source | str = "Wert") -> Decimal:
    """Excel-Zahl oder Zahlentext mit Punkt -> endliche Decimal-Zahl."""
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal, str)):
        raise TranslationError(f"{source}: Zahl erwartet, erhalten: {value!r}.")
    if isinstance(value, str):
        value = value.strip()
        if not re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", value):
            raise TranslationError(f"{source}: Zahl mit Dezimalpunkt erwartet: {value!r}.")
    try:
        number = Decimal(str(value))
    except InvalidOperation as exc:
        raise TranslationError(f"{source}: Ungültige Zahl {value!r}.") from exc
    if not number.is_finite():
        raise TranslationError(f"{source}: Zahl muss endlich sein.")
    return number
# endregion


# region V03: Allgemeiner Zellwert -> normalisierter interner Wert
def normalize_value(
    value: object, kind: Kind = "auto", required: bool = False,
    source: Source | str = "Wert",
) -> Scalar:
    """Typ explizit vorgeben oder echte Excel-Zahlen automatisch übernehmen.

    auto lässt Zahlentexte als Text stehen; number macht daraus Decimal.
    Texte werden außen getrimmt. Datums-/Zeitobjekte werden ISO-Texte.
    """
    if is_missing(value):
        if required:
            raise TranslationError(f"{source}: Pflichtwert fehlt.")
        return None
    if kind == "number":
        return to_decimal(value, source)
    if kind == "text":
        return value.isoformat() if isinstance(value, (date, datetime, time)) else str(value).strip()
    if kind == "bool":
        if isinstance(value, bool):
            return value
        if isinstance(value, str) and value.strip().casefold() in {"true", "false", "ja", "nein"}:
            return value.strip().casefold() in {"true", "ja"}
        raise TranslationError(f"{source}: Wahrheitswert erwartet: {value!r}.")
    if kind != "auto":
        raise TranslationError(f"{source}: Unbekannter Datentyp {kind!r}.")
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, Decimal)):
        return to_decimal(value, source)
    if isinstance(value, (date, datetime, time)):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    raise TranslationError(f"{source}: Nicht unterstützter Zellwert {value!r}.")
# endregion


# region V04: Dezimalzahl ausgeben; mindestens zwei Nachkommastellen, keine Rundung
def format_number(value: Decimal | int | float, minimum_places: int = 2) -> str:
    """1 -> '1.00'; 1.2345 -> '1.2345'; 1E+3 -> '1000.00'."""
    if minimum_places < 1:
        raise ValueError("minimum_places muss mindestens 1 sein.")
    text = format(to_decimal(value), "f")
    integer, dot, fraction = text.partition(".")
    return integer + "." + fraction.ljust(minimum_places, "0")
# endregion


# region V05: Positive Ganzzahl für Matrixindizes (kein Abrunden)
def positive_index(value: Scalar, source: Source | str) -> int:
    """1.00 ist Index 1; 1.50 und 0 sind ungültige Indizes."""
    number = to_decimal(value, source)
    if number < 1 or number != number.to_integral_value():
        raise TranslationError(f"{source}: Positiver ganzzahliger Index erwartet: {value!r}.")
    return int(number)
# endregion


# region V06: Mehrere IDs aus einer Zelle lesen
def split_references(value: Scalar, source: Source | str) -> tuple[str, ...]:
    """'CFG001; CFG002' -> ('CFG001', 'CFG002'); fehlend -> ()."""
    if value is None:
        return ()
    if not isinstance(value, str):
        raise TranslationError(f"{source}: Referenztext erwartet.")
    references = tuple(part.strip() for part in value.split(";") if part.strip())
    if len(references) != len(set(references)):
        raise TranslationError(f"{source}: Doppelte ID in Referenzliste {value!r}.")
    return references
# endregion
