# Katalog der vorbereiteten Bausteine

Jeder Baustein steht im jeweiligen Skript in einem benannten `# region`-Block.
Die Funktionsbeschreibungen nennen Aufgabe und Rückgabe; die Tabellen zeigen
die Stelle für eigene Anpassungen. Die Module werden importiert und laufen
gemeinsam in einem Python-Prozess.

## Anwendung und Dateidialoge

| Datei / Region | Baustein | Verwendung |
| --- | --- | --- |
| `main.py` A01 | `parse_arguments()` | CLI-Pfade und Optionen auslesen |
| `main.py` A02 | `main()` | Dateiauswahl, Formatwahl, Pipeline und Ergebnisfenster verbinden |
| `dialogs.py` D00 | `dialog_root()` | Gemeinsames unsichtbares Elternfenster mit Aufräumen |
| `dialogs.py` D01 | `choose_excel()` | Native Excel-Dateiauswahl, Abbruch → None |
| `dialogs.py` D02 | `choose_output()` | Nativer Speicherdialog, Abbruch → None |
| `dialogs.py` D03 | `show_message()` | Info- oder Fehlerfenster |

## Werte und Referenzadapter

| Datei / Region | Baustein | Verwendung |
| --- | --- | --- |
| `values.py` V01 | `is_missing()` | Leer/ka erkennen, 0 und False erhalten |
| `values.py` V02 | `to_decimal()` | Endliche Zahl mit Punkt lesen |
| `values.py` V03 | `normalize_value()` | Zelltyp, Fehlend und Pflichtwert behandeln |
| `values.py` V04 | `format_number()` | Dezimalzahl ohne zusätzliche Rundung schreiben |
| `values.py` V05 | `positive_index()` | Ganzzahlige Matrix-/Listenposition prüfen |
| `values.py` V06 | `split_references()` | Semikolongetrennte IDs in ein Tupel umwandeln |
| `dropdowns.py` Q01 | `selection_to_id()` | Direkte ID oder Formular-Listenindex auflösen |
| `dropdowns.py` Q02 | `read_dropdown_id()` | Auswahlzelle und optional eine ID-Quellliste lesen |

## Excel-Einlesegrundlagen

| Datei / Region | Baustein | Verwendung |
| --- | --- | --- |
| `excel_reader.py` E01 | `open_excel()` | Arbeitsmappe mit `data_only=True` öffnen/schließen |
| `excel_reader.py` E02 | `require_sheets()` | Benötigte Sheets prüfen |
| `excel_reader.py` E03 | `read_cell()` | Absolute Einzelzelle normalisieren |
| `excel_reader.py` E04 | `block_address()` | Relative Blockposition in Excel-Adresse übersetzen |
| `excel_reader.py` E05 | `read_offset()` | Relative Blockzelle lesen |
| `excel_reader.py` E06 | `iter_block_rows()` | Belegte Blockanfänge liefern; optionale Labels prüfen |
| `excel_reader.py` E07 | `add_unique()` | Block-ID ohne Überschreiben in einen Index aufnehmen |

## Leser für deine fünf Sheet-Typen

| Datei / Region | Baustein | Rückgabe |
| --- | --- | --- |
| `loaders.py` S01 | `read_metadata()` | `Metadata` mit Einzelwerten für den Header |
| `loaders.py` S02 | `read_materials()` | `dict[str, MaterialBlock]` |
| `loaders.py` S03 | `read_settings()` | `dict[str, SettingBlock]` |
| `loaders.py` S04 | `read_machines()` | `dict[str, MachineConfiguration]` |
| `loaders.py` S05 | `read_steps()` | Geordnete `list[Step]` mit noch ungelösten IDs |
| `loaders.py` S06 | `read_project()` | Alle fünf Leser als `Project` verbinden |

## Auflösen und fachliche Prüfungen

| Datei / Region | Baustein | Verwendung |
| --- | --- | --- |
| `resolver.py` R01 | `lookup()` | Eine ID im passenden Namensraum nachschlagen |
| `resolver.py` R02 | `select_matrix()` | 1-basierter, inklusiver Matrixausschnitt |
| `resolver.py` R03 | `resolve_steps()` | Konfigurationsobjekte an die Steps anbinden |
| `resolver.py` R04 | `validate_project()` | Eigene Fachregeln ergänzen |

## Textausgabe und Vorlagen

| Datei / Region | Baustein | Verwendung |
| --- | --- | --- |
| `renderer.py` T01 | `text_value()` | Einzelwert für lesbaren Text |
| `renderer.py` T02 | `python_literal()` | Sicher maskierter Python-Wert, auch verschachtelt |
| `renderer.py` T03 | `parameter_lines()` | Einfache Schlüssel-Wert-Zeilen |
| `renderer.py` T04 | `step_payload()` | Ausgabestruktur des aufgelösten Steps anpassen |
| `renderer.py` T05 | `step_context()` | Platzhalter für einen Step bereitstellen |
| `renderer.py` T06 | `render_template()` | Einen Textbaustein strikt einsetzen |
| `renderer.py` T07 | `document_chunks()` | Header, Steps und Footer nacheinander liefern |
| `renderer.py` T08 | `write_document()` | UTF-8 schreiben; Python prüfen; fertige Datei übernehmen |

## Zusammengesetzte Abläufe und Demo

| Datei / Region | Baustein | Verwendung |
| --- | --- | --- |
| `pipeline.py` P01 | `load_and_resolve()` | Nur einlesen, auflösen und fachlich prüfen |
| `pipeline.py` P02 | `translate()` | Vollständige Übersetzung ohne GUI |
| `examples/cfg_demo.py` B01 | `build_demo_project()` | CFG-Beispiel aus internen Datenklassen herstellen |
| `examples/cfg_demo.py` B02 | `main()` | Demo-Text und Demo-Python-Datei schreiben |

## Grundklassen und Zuordnungsblöcke

Diese Bausteine definieren Strukturen oder Konstanten und sind keine separat
auszuführenden Funktionen.

| Datei | Regionen | Inhalt |
| --- | --- | --- |
| `models.py` | M01–M09 | Quelle/Fehler, Metadata, MaterialBlock, SettingBlock, MachineConfiguration, MatrixReference, Step, ResolvedStep/ResolvedMatrix, Project |
| `layout.py` | L01–L09 | Cell, Offset, Blocks, MaterialsLayout, SettingsLayout, Rectangle/MachinesLayout, MatrixSlot, StepsLayout, WorkbookLayout |
| `mapping.py` | C01–C06 | Konkrete Zuordnung aller fünf Sheets und Verbindung zu LAYOUT |

## Vorhandene Templatevariablen

| Gültigkeit | Variablen |
| --- | --- |
| Alle drei Teile | `metadata_text`, `metadata_literal`, `step_count` |
| Nur Step | `step_index`, `step_id`, `step_name` |
| Nur Step, Text | `parameters_text`, `materials_text`, `settings_text`, `matrices_text` |
| Nur Step, Python | `step_id_literal`, `step_name_literal`, `step_literal`, `parameters_literal`, `materials_literal`, `settings_literal`, `matrices_literal` |

`step_literal` ist der vollständige Step-Datensatz. Die einzelnen
`*_literal`-Variablen erlauben eine freiere Python-Zielstruktur.
