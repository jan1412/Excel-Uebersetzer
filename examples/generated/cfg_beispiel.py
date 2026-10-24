# -*- coding: utf-8 -*-
"""Aus Excel erzeugte Beispieldaten. Fachliche Befehle in den Vorlagen ergänzen."""

METADATA = {'project_name': 'CFG-Beispiel', 'author': 'Beispiel', 'version': '1.0'}

STEPS = [
    {
        "id": 'STEP001',
        "name": 'Bearbeiten',
        "parameters": {'feed': 0.00},
        "materials": {'MAT_STEEL_001': {'material': 'Stahl', 'parameters': {'density': 7.85, 'thickness': 1.00, 'note': None}}},
        "settings": {'CFG001': {'name': 'Standard', 'parameters': {'speed_rpm': [1500.00], 'duration_s': [10.00], 'limits': [0.00, 1.00]}}},
        "matrices": [{'configuration': 'MCH001', 'bounds': [1.00, 2.00, 2.00, 3.00], 'values': [[2.00, 3.00], [5.00, 6.00]]}],
    },
    {
        "id": 'STEP002',
        "name": 'Finish',
        "parameters": {'feed': None},
        "materials": {'MAT_STEEL_001': {'material': 'Stahl', 'parameters': {'density': 7.85, 'thickness': 1.00, 'note': None}}, 'MAT_STEEL_002': {'material': 'Stahl', 'parameters': {'density': 7.85, 'thickness': 2.00, 'note': 'Zweiter Block desselben Materials'}}},
        "settings": {'CFG002': {'name': 'Schonend', 'parameters': {'speed_rpm': [900.00], 'duration_s': [20.00], 'limits': [0.10, 0.50]}}},
        "matrices": [{'configuration': 'MCH001', 'bounds': None, 'values': [[1.00, 2.00, 3.00], [4.00, 5.00, 6.00], [7.00, 8.00, 9.00]]}],
    },
]


def main():
    for step in STEPS:
        print(step["id"], step["name"])
        # Hier die späteren fachlichen Aktionen ergänzen.


if __name__ == "__main__":
    main()
