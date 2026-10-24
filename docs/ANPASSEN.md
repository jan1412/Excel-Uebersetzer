# Das Grundgerüst an deine Excel-Datei anpassen

## 1. Zunächst nur mapping.py bearbeiten

Eine Einzelzelle beschreibst du mit `Cell`:

```python
"project_name": Cell("F3", kind="text", required=True)
```

Ein wiederholter Bereich wird über `Blocks` definiert:

```python
Blocks(
    sheet="Einstellungen",
    first_row=10,
    first_column=3,  # C
    height=8,
    last_row=None,
)
```

Das ergibt Blockanfänge `C10`, `C18`, `C26`, ... . `height` enthält auch
eventuelle Trenn-/Leerzeilen. `last_row=None` berücksichtigt nach unten
kopierte Blöcke automatisch. Optional kann eine letzte Bereichszeile fest
vorgegeben werden; sie sollte das Ende eines vollständigen Blocks sein.

Eine Zelle innerhalb dieses Blocks wird mit 0-basierten Offsets adressiert:

```python
Offset(0, 0, "text", True)    # C10 im ersten Block
Offset(2, 1, "number", True)  # D12 im ersten Block, D20 im zweiten
```

Die Referenz liegt standardmäßig links oben. Eine andere Referenzzelle ist
möglich: `Blocks(..., reference=Offset(0, 2, "text", True))`.
Für Maschinenblöcke entspricht der voreingestellte Standard deiner Vorgabe
einer Referenz in lokaler Zelle `(1, 1)`.

### Feldtypen

| kind | Verhalten |
| --- | --- |
| `auto` | Echte Zahlen → Decimal; Text bleibt Text; bool bleibt bool |
| `number` | Zahl oder Zahlentext mit Punkt → Decimal |
| `text` | Wert → Text; außen liegende Leerzeichen entfernen |
| `bool` | Excel-bool oder Text `True`/`False`/`Ja`/`Nein` |

Bei allen Typen werden leer und `ka` zuerst als fehlend behandelt. `required`
legt fest, ob das erlaubt ist. Es existiert kein impliziter Ersatz durch 0.
Da `ka` als reservierte Fehlend-Kennung vereinbart ist, kann es nicht zugleich
eine reguläre Referenz-ID oder ein wörtlicher Parametertext sein.

`number` akzeptiert auch ganze Eingabezahlen und Exponentialschreibweise;
die Ausgabe verwendet immer die normale Dezimalschreibweise mit Punkt.
Die von Excel angezeigten Dezimalstellen oder lokalisierte Trennzeichen
werden nicht aus dem Zellformat übernommen.

### Feste Beschriftungen optional prüfen

```python
Blocks("Einstellungen", 2, 1, 4,
       labels={(1, 0): "Drehzahl", (2, 0): "Dauer"})
```

Das prüft A3/A4, A7/A8 usw., sobald der jeweilige Block Daten enthält.
So lassen sich einige verschobene oder unvollständig kopierte Bereiche
erkennen. Ohne `labels` werden sichtbare Beschriftungen nicht vorausgesetzt.
Zellfarben, Schriftarten und Rahmen sind keine Erkennungsmerkmale.

Ein vollständig leerer Datenblock wird ignoriert, auch wenn seine festen
Beschriftungen stehen bleiben. Ein belegter Block ohne Pflicht-ID ist ein Fehler.
Nur gemappte Datenfelder zählen dabei; zusätzliche nicht gemappte Notizen
werden nicht als Datensatz interpretiert.

## 2. Andere Blockformen zusammensetzen

Die Standardleser setzen pro Layout eine feste Blockhöhe und eine feste
Startspalte voraus. Bei zwei nebeneinanderliegenden Materialbereichen kannst
du dieselbe Funktion zweimal verwenden und die Ergebnisse eindeutig verbinden:

```python
from dataclasses import replace
from excel_uebersetzer.excel_reader import add_unique
from excel_uebersetzer.loaders import read_materials
from excel_uebersetzer.mapping import MATERIALS

left = MATERIALS
right = replace(MATERIALS, blocks=replace(MATERIALS.blocks, first_column=5))
materials = {}
for layout in (left, right):
    for block in read_materials(workbook, layout).values():
        add_unique(materials, block)
```

Entsprechend lassen sich getrennte Matrixbereiche mit verschiedenen Größen
über mehrere `MachinesLayout`-Definitionen lesen. Setze dabei `first_row` und
`last_row` passend, damit die einzelnen Leser nur ihren Bereich betrachten.

Bei variablen Blockgrößen ersetzt du beispielsweise nur `read_settings()`.
Solange das Ergebnis weiter `dict[str, SettingBlock]` ist, bleiben Resolver
und Renderer verwendbar. Für komplett individuelle Zellzuordnungen kannst du
`read_cell()`, `read_offset()` und `add_unique()` einzeln kombinieren.

## 3. Dropdowns und Formeln

Normale Zell-Dropdowns und Formeln mit einer gespeicherten ID brauchen keinen
besonderen Adapter. Die Standardleser verwenden den gespeicherten Zellwert.

Wenn ein Formular-Kombinationsfeld stattdessen seinen **Listenindex** in eine
verknüpfte Zelle schreibt, kannst du diesen optionalen Baustein nutzen:

```python
from excel_uebersetzer.dropdowns import read_dropdown_id

selected_id = read_dropdown_id(
    workbook["Ablauferstellung"], "G2",  # verknüpfte Zelle des Steuerelements
    options_sheet=workbook["Einstellungen"],
    options_range="H2:H20",              # exakt die zugehörige ID-Quellliste
)
```

Die Quellliste muss eine Spalte mit IDs enthalten. Leerplätze werden nicht
entfernt, weil sie die Positionszählung verändern würden. Der Index beginnt
bei 1. Bei einer normalen Textauswahl genügt
`read_dropdown_id(workbook["Ablauferstellung"], "B4")`.

Du baust den Adapter in deinen angepassten Step-Leser ein, etwa als
`setting_references=(selected_id,) if selected_id is not None else ()`.
Die voreingestellten Standardleser gehen von direkt gespeicherten IDs aus.
Bei kopierten Steuerelementen muss deren Excel-Zellverknüpfung pro Block
selbst stimmen; die Steuerelementobjekte werden nicht geprüft.

Es gibt weder eine Formel-Aktualitätsregel noch einen Excel-Prozess im
Hintergrund. Die einzige Formelannahme bleibt: Excel ist aktuell und gespeichert.
Die Inhalte fehlender Formelergebniszellen werden wie andere fehlende Werte
behandelt, nicht separat untersucht.

## 4. Mehrere Werte und mehrere Matrixabschnitte

Ein mehrteiliger Einstellungsparameter bekommt mehrere Offsets:

```python
"ramp": (
    Offset(1, 1, "number"),
    Offset(1, 2, "number"),
    Offset(1, 3, "number"),
)
```

Der interne Wert ist immer ein Tupel, z.B. `(Decimal("0"), None, Decimal("1"))`.
Eine Liste mit Zahlen besitzt also keine implizite Entfernung fehlender Slots.

Ein Step kann mehrere Maschinenabschnitte verwenden. Ergänze in
`STEPS.matrix_slots` weitere `MatrixSlot`-Einträge. Jeder Slot bekommt eine
Referenzzelle sowie optional vier Zellen für Startzeile, Startspalte, Endzeile
und Endspalte. Ohne diese vier Werte wird die gesamte Datenmatrix verwendet.

Eine Matrixgrenze darf `1.00` sein, aber nicht `1.50`. Das sind Positionen,
keine beliebigen Parameter. Die Matrixdaten selbst dürfen fehlende Werte
enthalten; entsprechende fachliche Pflichten ergänzt du selbst.

## 5. Fachliche Datenklassen anpassen

Die Grundklassen sind generisch und besitzen absichtlich keine vorweggenommenen
Maschinen-/Materialregeln. Du kannst Felder ergänzen oder eigene Klassen nutzen.

```python
@dataclass
class MySetting:
    reference: str
    speed_rpm: Decimal
    lower_limit: Decimal | None
    upper_limit: Decimal | None
```

Die dafür benötigten Importe sind `from dataclasses import dataclass` und
`from decimal import Decimal`. Passe dann den betreffenden Leser, den
Resolverzugriff und `step_payload()`/`step_context()` an.

Für geringe Änderungen reichen die vorhandenen Parameter-Dictionaries.
Beispielsweise ist `config.parameters["speed_rpm"][0]` bereits direkt nutzbar.

Die Standardstruktur hält direkte Step-Werte, Materialien, Einstellungen und
Maschinenmatrizen getrennt. Es gibt keine automatische Priorität oder
Vererbung. Wenn du später lokale Overrides möchtest, lege diese Regel in
einer eigenen Funktion ausdrücklich fest.

## 6. Textbausteine und vollständiges Zielskript

Die Vorlagen sind normale UTF-8-Textdateien. Platzhalter heißen `${name}`:

```text
Step ${step_index}: ${step_name}
${settings_text}
```

`$$` ergibt ein wörtliches Dollarzeichen. Es gibt keine besondere Template-
Programmiersprache und keine Schleifen in den Vorlagen. Die Schleife über
Steps steht in `document_chunks()` und ist damit als Python-Baustein anpassbar.

Der Ablauf schreibt Header einmal, den Step-Baustein für jeden aufgelösten
Step und den Footer einmal. Weitere wiederholte Bausteine kannst du vor oder
innerhalb der Step-Schleife mit `render_template()` ergänzen.

Für einen zusätzlichen eigenen Parameter:

1. In `mapping.py` ein Zellfeld hinzufügen.
2. Der generische Leser übernimmt es in das Parameter-Dictionary.
3. In `step_context()` eine neue Templatevariable bereitstellen.
4. Diese Variable in der passenden Vorlage verwenden.

Beispiel für einen direkt gemappten Vorschub:

```python
# In step_context(), nach Erstellen von context:
context["feed_literal"] = python_literal(item.step.parameters["feed"])
```

```python
# In step.py.tpl:
set_feed(${feed_literal})
```

Für Python-Werte immer `python_literal()` benutzen. Ein Text sollte nicht
einfach zwischen selbst gesetzte Anführungszeichen eingefügt werden.
Code, Imports, Funktionen und Einrückungen gehören ausdrücklich in die Vorlagen.
Der Generator führt die erzeugte Datei nicht aus.

Die vorhandene `.py`-Vorlage erwartet einen Step als Dateneintrag innerhalb
einer Liste. Wenn du stattdessen Befehle schreiben möchtest, müssen Header,
Step und Footer gemeinsam zu dieser neuen Zielstruktur passen.

Für ein drittes Textformat ergänzt du die Vorlagendateien und die zulässigen
Formate in `document_chunks()` und `parse_arguments()`. Die eigentliche
Dateiendung ist frei. Binärformate erfordern einen anderen Schreibbaustein.

## 7. Eigene Prüfungen ergänzen

`validate_project()` ist die vorgesehene Stelle für Fachregeln. Beispiel:

```python
for item in steps:
    for setting in item.settings:
        speed = setting.parameters["speed_rpm"][0]
        if speed is None or speed <= 0:
            raise TranslationError(f"{setting.source}: Drehzahl muss positiv sein.")
```

Das ist ein Beispiel und keine voreingestellte Regel. Nullwerte sind allgemein
gültig; ob ein einzelner Parameter null sein darf, entscheidest du fachlich.

Vorhanden sind die für das Zusammensetzen notwendigen Prüfungen: Pflichtfelder,
Datentypen, doppelte IDs, unbekannte Verweise, Matrixgrenzen und fehlende
Templatevariablen. Ein komplett vergessener, nicht referenzierter Block oder
ein plausibler, aber inhaltlich falscher Wert kann dadurch nicht erkannt werden.

`--check` prüft Excel-Einlesen und Auflösen. Es rendert keine Vorlagen und
garantiert deshalb nicht, dass eine selbst geänderte Vorlage gültig ist.
Die komplette Übersetzung prüft zusätzlich die Vorlagen und bei `py` die Syntax.
