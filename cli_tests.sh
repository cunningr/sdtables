#!/bin/bash


# DISPLAY
python3 sdtables_cli.py display --input examples/example.xlsx
python3 sdtables_cli.py display --input examples/example.xlsx --format yaml
python3 sdtables_cli.py display --input examples/example.xlsx --format json

# VALIDATE
python3 sdtables_cli.py validate --schema examples/vehicles.yaml --input examples/example.xlsx
python3 sdtables_cli.py validate --schema examples --input examples/example.xlsx
python3 sdtables_cli.py validate --schema examples --input examples/example.xlsx --format json
