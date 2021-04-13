#!/bin/bash

# BUILD
python3 sdtables_cli.py build --schema examples --output test --filename-as-sheet
python3 sdtables_cli.py build --schema examples --output test2 --noseed
python3 sdtables_cli.py build --schema examples --output test3 --format json

# DISPLAY
python3 sdtables_cli.py display --input test.xlsx
python3 sdtables_cli.py display --input test.xlsx --format yaml
python3 sdtables_cli.py display --input test2.xlsx --format yaml
python3 sdtables_cli.py display --input test3.xlsx --format json

# VALIDATE
python3 sdtables_cli.py validate --schema examples/vehicles.yaml --input examples/example.xlsx
python3 sdtables_cli.py validate --schema examples --input examples/example.xlsx
python3 sdtables_cli.py validate --schema examples --input examples/example.xlsx --format json
