"""IDs auflösen; Matrixbereiche auswählen; Platz für eigene Fachregeln."""
from __future__ import annotations

from typing import TypeVar

from .models import Matrix, MatrixReference, Project, ResolvedMatrix, ResolvedStep, Source, TranslationError

T = TypeVar("T")


# region R01: Einzelne Konfiguration nach ID nachschlagen
def lookup(index: dict[str, T], reference: str, kind: str, source: Source | None) -> T:
    """ID-Vergleich ist exakt und unterscheidet Groß-/Kleinschreibung."""
    try:
        return index[reference]
    except KeyError as exc:
        raise TranslationError(f"{source}: Unbekannte {kind}-Referenz {reference!r}.") from exc
# endregion


# region R02: Matrixabschnitt aus 1-basierten inklusiven Grenzen ausschneiden
def select_matrix(matrix: Matrix, reference: MatrixReference) -> Matrix:
    """Grenzen beziehen sich auf die Datenmatrix, nicht auf Excel-Adressen."""
    if not matrix or not matrix[0] or any(len(row) != len(matrix[0]) for row in matrix):
        raise TranslationError(f"{reference.source}: Datenmatrix muss rechteckig und nicht leer sein.")
    if reference.bounds is None:
        return matrix
    top, left, bottom, right = reference.bounds
    if any(isinstance(n, bool) or not isinstance(n, int) for n in reference.bounds):
        raise TranslationError(f"{reference.source}: Matrixgrenzen müssen ganzzahlig sein.")
    if not (1 <= top <= bottom <= len(matrix) and 1 <= left <= right <= len(matrix[0])):
        raise TranslationError(f"{reference.source}: Matrixabschnitt {reference.bounds} außerhalb der Datenmatrix {len(matrix)}×{len(matrix[0])}.")
    return tuple(tuple(row[left - 1:right]) for row in matrix[top - 1:bottom])
# endregion


# region R03: Alle Step-Verweise auflösen; keine automatische Parametervermischung
def resolve_steps(project: Project) -> list[ResolvedStep]:
    """Keine Vererbung/Override-Magie: jeder referenzierte Block bleibt separat."""
    result = []
    for step in project.steps:
        materials = tuple(lookup(project.materials, ref, "Material", step.source) for ref in step.material_references)
        settings = tuple(lookup(project.settings, ref, "Einstellungs", step.source) for ref in step.setting_references)
        matrices = []
        for ref in step.matrix_references:
            config = lookup(project.machines, ref.configuration, "Maschinen", ref.source or step.source)
            matrices.append(ResolvedMatrix(ref, select_matrix(config.matrix, ref)))
        result.append(ResolvedStep(step, materials, settings, tuple(matrices)))
    return result
# endregion


# region R04: Eigene fachliche Regeln ergänzen
def validate_project(project: Project, steps: list[ResolvedStep]) -> None:
    """ANPASSEN: z.B. Geschwindigkeit > 0 oder kompatible Material/Maschine-Paare.

    Keine solchen Fachregeln werden vorausgesetzt. Leere Abläufe sind erlaubt.
    Die vorhandene Prüfung schützt selbst angepasste Leser vor doppelten Steps.
    """
    ids = [item.step.reference for item in steps]
    if len(ids) != len(set(ids)):
        raise TranslationError("Die Step-Liste enthält doppelte IDs.")
# endregion
