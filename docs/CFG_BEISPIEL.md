# Durchgängiges CFG-Beispiel

Die hier beschriebenen Beispieladressen sind in `mapping.py` voreingestellt.
`examples/cfg_cells.json` enthält alle Beispielzellen. Die echte Excel-Datei
kann anders aussehen: Du passt dann die Zuordnung oder einzelne Leser an.

## 1. Metadata: Einzelparameter für den Header

| Zelle | Inhalt | Interner Schlüssel |
| --- | --- | --- |
| B2 | CFG-Beispiel | `project_name` |
| B3 | Beispiel | `author` |
| B4 | 1.0 als Text | `version` |

```python
METADATA = {
    "project_name": Cell("B2", "text", True),
    "author": Cell("B3", "text"),
    "version": Cell("B4", "text"),
}
```

`read_metadata()` liefert `Metadata(parameters={...})`. Der Header bekommt
`${metadata_text}` für lesbaren Text oder `${metadata_literal}` für Python.

## 2. Materialien: mehrere Blöcke je Material

| Zeile im Block | Spalte A | Spalte B |
| --- | --- | --- |
| 1 | MAT_STEEL_001 | Stahl |
| 2 | Dichte | 7.85 |
| 3 | Dicke | 1.00 |
| 4 | Notiz | ka |

Erster Block: `A2:B5`. Zweiter Block: `A6:B9`, beispielsweise mit
`MAT_STEEL_002`, demselben Material `Stahl` und Dicke `2.00`.

```python
MaterialBlock(
    reference="MAT_STEEL_001",
    material="Stahl",
    parameters={"density": Decimal("7.85"), "thickness": Decimal("1.00"), "note": None},
)
```

Jeder Parameter dieses Blocktyps besitzt genau einen Wert. Die Materialnamen
dürfen sich wiederholen. Die Referenz-ID muss pro Materialblock eindeutig sein.
Beim Kopieren `A2:B5` nach `A10:B13` also eine neue ID eintragen.

## 3. Maschinenkonfiguration: ID in lokaler Position (1, 1)

| Excel-Zeile | A | B | C | D |
| --- | --- | --- | --- | --- |
| 2 | MCH001 | X | Y | Z |
| 3 | R1 | 1.00 | 2.00 | 3.00 |
| 4 | R2 | 4.00 | 5.00 | 6.00 |
| 5 | R3 | 7.00 | 8.00 | 9.00 |

`A2` ist die lokale Position `(1, 1)` des gesamten Konfigurationsblocks.
Die eigentliche Datenmatrix ist `B3:D5`. Deshalb:

```python
MachinesLayout(
    blocks=Blocks("Maschinenkonfiguration", 2, 1, 4),
    data=Rectangle(row=1, column=1, rows=3, columns=3),
)
```

Die `Rectangle`-Offsets sind 0-basiert und beziehen sich auf `A2`.
Die spätere Step-Auswahl `(1, 2, 2, 3)` bezieht sich hingegen 1-basiert auf
die **Datenmatrix** und liefert:

| | Spalte 2 | Spalte 3 |
| --- | --- | --- |
| Zeile 1 | 2.00 | 3.00 |
| Zeile 2 | 5.00 | 6.00 |

Das ist Excel-Bereich `C3:D4`. Grenzen sind immer inklusive.
Ohne Grenzen wird die gesamte Datenmatrix verwendet.

## 4. Einstellungen: CFG001 als Parametergruppe

| Excel-Zeile | A | B | C |
| --- | --- | --- | --- |
| 2 | CFG001 | Standard | |
| 3 | Drehzahl | 1500.00 | |
| 4 | Dauer | 10.00 | |
| 5 | Grenzen | 0.00 | 1.00 |

`read_settings()` liefert:

```python
SettingBlock(
    reference="CFG001",
    name="Standard",
    parameters={
        "speed_rpm": (Decimal("1500.00"),),
        "duration_s": (Decimal("10.00"),),
        "limits": (Decimal("0.00"), Decimal("1.00")),
    },
)
```

Auch Parameter mit nur einem Wert werden hier als Tupel gespeichert. Das
vermeidet Sonderfälle zwischen einem und mehreren Werten. Ein fehlender
zweiter Wert bleibt als `None` an seiner Position erhalten.

Für `CFG002` den gesamten Block vier Zeilen nach unten kopieren, die ID
ändern und beispielsweise `900.00`, `20.00` sowie Grenzen `0.10`/`0.50` setzen.

## 5. Ablauferstellung: Verweise und direkte Step-Parameter

| Zelle | Inhalt | Bedeutung |
| --- | --- | --- |
| A2 | STEP001 | Eindeutige Step-ID |
| B2 | Bearbeiten | Name |
| B3 | MAT_STEEL_001 | Gewählter Materialblock |
| B4 | CFG001 | Gewählte Einstellungsgruppe |
| B5 | MCH001 | Gewählte Maschinenkonfiguration |
| B6 | 1.00 | Startzeile der Datenmatrix |
| C6 | 2.00 | Startspalte der Datenmatrix |
| D6 | 2.00 | Endzeile der Datenmatrix |
| E6 | 3.00 | Endspalte der Datenmatrix |
| B7 | 0.00 | Direkter Parameter `feed`; gültiger Nullwert |

`B3:B5` dürfen normale Dropdownzellen sein. Der Generator benötigt nur die
gespeicherten IDs. Er muss die Dropdowndefinitionen dafür nicht auswerten.
Formeln, die dieselben IDs zurückgeben, funktionieren unter der vereinbarten
Annahme einer aktuellen gespeicherten Excel-Datei ebenfalls.

Mehrere Material- oder Einstellungsreferenzen lassen sich mit Semikolon
angeben: `CFG001; CFG002`. Ein normaler Excel-Dropdown ist dabei weiterhin
eine Einzelauswahl; Mehrfachauswahl wird vom Generator nicht in Excel ergänzt.
Für getrennte Auswahlzellen kannst du die entsprechenden Leser erweitern.

Das ungelöste Datenobjekt lautet vereinfacht:

```python
Step(
    reference="STEP001",
    name="Bearbeiten",
    parameters={"feed": Decimal("0.00")},
    material_references=("MAT_STEEL_001",),
    setting_references=("CFG001",),
    matrix_references=(MatrixReference("MCH001", (1, 2, 2, 3)),),
)
```

Ein zweiter Step beginnt in Zeile 8. Im Beispiel nutzt er beide Materialblöcke,
`CFG002`, die vollständige Matrix und einen fehlenden direkten `feed`-Wert.
Eine leere Referenz bedeutet: diesen Referenztyp für den Step nicht verwenden.

## 6. Bausteine zu einem eigenen Ablauf verbinden

```python
from pathlib import Path
from excel_uebersetzer.dialogs import choose_excel
from excel_uebersetzer.excel_reader import open_excel
from excel_uebersetzer.loaders import read_project
from excel_uebersetzer.mapping import LAYOUT
from excel_uebersetzer.resolver import resolve_steps, validate_project
from excel_uebersetzer.renderer import document_chunks, write_document

source = choose_excel()
if source is not None:
    with open_excel(source) as workbook:
        project = read_project(workbook, LAYOUT)

    steps = resolve_steps(project)
    validate_project(project, steps)
    chunks = document_chunks(project.metadata, steps, "txt")
    write_document(Path("output/mein_ablauf.txt"), chunks)
```

Oder dieselben vorbereiteten Bausteine über eine einzige Funktion aufrufen:

```python
from excel_uebersetzer.pipeline import translate
translate("MeineDatei.xlsx", "output/mein_ablauf.py", format_name="py")
```

## 7. Einzelne aufgelöste Parameter in eigene Befehle einsetzen

```python
from excel_uebersetzer.values import format_number

item = steps[0]  # ResolvedStep für STEP001
config = next(block for block in item.settings if block.reference == "CFG001")
rpm = config.parameters["speed_rpm"][0]
duration = config.parameters["duration_s"][0]
command = f"run_step(speed_rpm={format_number(rpm)}, duration_s={format_number(duration)})"
```

Der fertige Text lautet:

```python
run_step(speed_rpm=1500.00, duration_s=10.00)
```

Für ein vollständiges Skript muss die Funktion `run_step` im Header importiert
oder definiert werden. Die ausgelieferte Python-Vorlage ist bereits vollständig
lauffähig und enthält stattdessen eine Datenliste und eine einfache `main()`.

## 8. Ohne Excel ausprobieren

```console
python -m examples.cfg_demo
```

Danach sind die generierten Beispiele unter `output/` verfügbar. Das Projekt
bleibt bis zur Anpassung deiner echten Zellzuordnungen bei diesen Beispieldaten.
