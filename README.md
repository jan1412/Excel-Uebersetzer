# Excel-Übersetzer

Anpassbares Python-Grundgerüst für **Excel → Parameter/Datenklassen → Textbausteine → vollständige Datei**.
Die Zelladressen sind Beispiele. Du ersetzt sie durch deine tatsächlichen
Zuordnungen und passt die Fachklassen sowie den Zieltext selbst an.

## Start

Python 3.10 oder neuer mit Tkinter verwenden. Unter Windows gehört Tkinter
zur üblichen Python-Installation. Die Entwicklung wurde mit Python 3.12 geprüft.

```console
python -m pip install -r requirements.txt
python main.py
```

Ohne Pfade öffnen sich ein nativer Dialog zur Excel-Auswahl und anschließend
ein Speicherdialog. Unter Windows kannst du auch `start.bat` verwenden.

```console
python main.py "C:\Daten\Parameter.xlsx" --output "C:\Daten\Ergebnis.txt"
python main.py "C:\Daten\Parameter.xlsx" --output "C:\Daten\Ergebnis.py"
python main.py "C:\Daten\Parameter.xlsx" --check
```

Eine `.py`-Zielendung wählt automatisch das Python-Vorlagenset. Mit `--format txt`
oder `--format py` lässt sich das unabhängig von der Dateiendung festlegen.
`--templates ORDNER` verwendet eigene Vorlagen.

## Die ersten Anpassungen

1. In [`mapping.py`](excel_uebersetzer/mapping.py) Sheetnamen, Einzelzellen,
   Startpositionen, Blockhöhen, Parameter und Matrixgrößen ändern.
2. In [`models.py`](excel_uebersetzer/models.py) bei Bedarf eigene Fachklassen ergänzen.
3. In [`loaders.py`](excel_uebersetzer/loaders.py) nur die Leser ändern, deren
   Excel-Struktur nicht zum vorbereiteten Blockraster passt.
4. In [`renderer.py`](excel_uebersetzer/renderer.py) die Ausgabevariablen anpassen
   und unter [`templates/`](templates/) die gewünschten Textbausteine eintragen.

Jeder vorbereitete Funktionsbaustein besitzt einen benannten `# region`-Block
und eine Beschreibung. VS Code kann diese Bereiche einklappen. Du kannst die
Bausteine importieren, ersetzen oder zu einem eigenen Ablauf verbinden.

## Dokumentation

- [CFG-Beispiel über alle fünf Sheets](docs/CFG_BEISPIEL.md)
- [Eigene Zuordnungen, Datenklassen und Vorlagen](docs/ANPASSEN.md)
- [Katalog aller Funktionsbausteine](docs/BAUSTEINE.md)

## Festgelegte Behandlung der Eingaben

- Die Excel-Datei gilt als aktuell, vollständig berechnet und gespeichert.
  Es gibt keine Formel-Neuberechnung und keine Aktualitätskontrolle.
- `openpyxl` liest gespeicherte Ergebnisse mit `data_only=True`.
- Unterstützt werden `.xlsx` und `.xlsm`; Makros werden nicht ausgeführt.
- `None`, leere/weiße Zelltexte und `ka` (auch `KA`) werden intern `None`.
- `0` und `False` sind gültige Werte. Pflichtwerte werden ausdrücklich im Mapping markiert.
- Zahlen werden intern als `Decimal` verarbeitet. Zahlentexte in `number`-Feldern
  benötigen einen Punkt. Eine automatische Umdeutung von `1,25` findet nicht statt.
- Die Ausgabe verwendet mindestens zwei Nachkommastellen und keine zusätzliche
  Rundung: `1` → `1.00`, `1.2345` → `1.2345`. Excel-Anzeigeformate werden nicht kopiert.
- Textzellen bleiben Text. `auto` interpretiert echte Excel-Zahlen numerisch;
  der Text `"1.00"` bleibt dabei Text. Für Zahlentexte `kind="number"` verwenden.
- In Python-Ausgaben erscheinen fehlende Werte als `None`, in Textausgaben als `ka`.
- Material-, Einstellungs- und Maschinenreferenzen besitzen getrennte Namensräume.
  Direkte Step-Parameter überschreiben keine referenzierten Gruppen.

## Demo ohne Excel und Prüfung

```console
python -m examples.cfg_demo
python output/cfg_beispiel.py
python -m unittest discover -s tests -v
```

Die Demo erzeugt `output/cfg_beispiel.txt` und `output/cfg_beispiel.py` direkt
aus den Datenklassen. Sie ist unabhängig von einer fertigen Excel-Datei.
[`cfg_cells.json`](examples/cfg_cells.json) dokumentiert die dazu passenden
Beispielzellen; es ist keine zusätzliche Laufzeit-Konfigurationsdatei.

Die Tests prüfen Einleser mit repräsentativen Excel-Zellen im Speicher, kopierte
Blöcke, Referenzauflösung, Matrixgrenzen, fehlende Werte, Nullwerte, Vorlagen,
Python-Syntax und Dialogsteuerung. Eine echte Excel-Neuberechnung wird nicht
benötigt oder geprüft. Native Windows-Dialoge wurden in dieser Umgebung nicht
interaktiv geprüft.

## Grenzen des Grundgerüsts

Die Standardleser erkennen vertikal wiederholte Blöcke mit jeweils gleicher
Höhe. Unterschiedliche Bereiche, horizontale Anordnungen oder abweichende
Blocktypen kannst du mit mehreren Layouts oder eigenen Lesern abdecken; die
übrigen Bausteine bleiben dabei nutzbar. Zusammengeführte Zellen werden über
ihre linke obere Zelle adressiert.

Die Vorlagen erzeugen Textformate. Eine andere Dateiendung macht aus Text kein
PDF, DOCX, XLSX oder ausführbares Binärprogramm. Die Python-Vorlage liefert ein
vollständiges lauffähiges Beispielskript mit den aufgelösten Daten; fachliche
Maschinen-/Anwendungsbefehle ergänzt du in den Vorlagen.

Die `.py`-Vorlage schreibt gewöhnliche Zahlenliterale. Beim Ausführen speichert
Python diese als `float`; für höhere Rechengenauigkeit im erzeugten Skript
kannst du `Decimal("...")` ausgeben. Excel selbst hat bereits eine begrenzte
Zahlengenauigkeit; der Generator stellt verlorene Stellen nicht wieder her.

Quelldateien werden nicht gespeichert oder verändert. Bestehende Zieldateien
werden erst durch eine vollständig erzeugte, bei `.py` auch syntaktisch
geprüfte Ausgabe ersetzt.

## Technische Referenzen

- [openpyxl: gespeicherte Formelresultate / data_only](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.reader.excel.html)
- [Python: native Tkinter-Dateidialoge](https://docs.python.org/3/library/dialog.html)
- [Python: string.Template](https://docs.python.org/3/library/string.html#template-strings)
