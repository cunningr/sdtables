# sdtables

WARNING: Version 2.0.0 introduces breaking changes.  If you depend on previous functions please pin your dependencies to 1.0.8 and plan to migrate to the new ```SdTables()``` class.

sdtables (schema defined tables) is a module providing convenient wrapper functions for working with and creating tabulated data from various sources including MS Excel.  We are using pythons jsonschema to build table structures and validate data.

Install with pip3:

```
pip3 install sdtables
```

Check out the ```./examples``` for usage.

TODO: Update docs
Documentation can be [found here](https://cunningr.github.io/sdtables/)

## SDtables CLI

For convenience, you can now use the ```sdtables``` script to parse an Excel file and dump the containing tables to CLI in various formats using pythons [tabulate](https://github.com/astanin/python-tabulate) module.

Example:

```
(general) CUNNINGR-M-X436:sdtables cunningr$ sdtables display --input examples/example.xlsx
            ____  ____  _        _     _              ____ _     ___
           / ___||  _ \| |_ __ _| |__ | | ___  ___   / ___| |   |_ _|
           \___ \| | | | __/ _` | '_ \| |/ _ \/ __| | |   | |    | |
            ___) | |_| | || (_| | |_) | |  __/\__ \ | |___| |___ | |
           |____/|____/ \__\__,_|_.__/|_|\___||___/  \____|_____|___|


Worksheet: Sheet1

Table Name: super.heros
+--------+-----------+-------+----------+---------------------+
| Name   | Surname   |   Age | Gender   | Date of Birth       |
+========+===========+=======+==========+=====================+
| Bat    | Man       |    34 | M        | 1954-04-04 00:00:00 |
+--------+-----------+-------+----------+---------------------+
| Super  | Man       |    32 | M        | 1954-04-05 00:00:00 |
+--------+-----------+-------+----------+---------------------+
| Wonder | Woman     |    26 | F        | 1954-04-06 00:00:00 |
+--------+-----------+-------+----------+---------------------+
| Super  | Woman     |    38 | F        | 1954-04-07 00:00:00 |
+--------+-----------+-------+----------+---------------------+

Table Name: vehicles
+---------+---------+------------+--------+
| Make    | Model   | Colour     | Type   |
+=========+=========+============+========+
| Fiat    | Panda   | Light Blue | Car    |
+---------+---------+------------+--------+
| Citroen | AX      | Grey       | Car    |
+---------+---------+------------+--------+
| Honda   | MT5     | Blue       | Bike   |
+---------+---------+------------+--------+
| Yamaha  | TZR     | Black      | Bike   |
+---------+---------+------------+--------+
```

Use the ```--output``` option to set the ```tablulate``` output format.

## Features Summary

### Excel (xlsx)

| Issue | Description | Status |
|:---:|:---|:---:|
| | Load data to dict from Excel tables | Complete |
| | Add data tables to excel using first row keys as headers | Complete |
| | Add data tables to excel using a schema to define headers | Complete |
| | Add Excel data validation to excel using a schema (enum\|tref\|default values) | Complete |
| | Add schema with Excel data validation to excel without data | Complete |
| | Update data in existing table | On Hold |
| | Delete table | On Hold |
| | Validate table data against jsonschema draft7_format_checker | Complete |
| | Can tref schema definition be improved? | New |
| | Update docs | New |
| | Create examples | Complete |
| | Cli for convenient testing and common operations | Complete |

Notes:

 * Update data in table -> create new git Issue
  * Without schema there is no data validation.  Can this be enhanced?
  * This breaks data validation since this is not moved when rows are inserted
 * Delete table -> create new git Issue
    This is broken as we don't remove the data or other formatting
 * Validate table data against schema (add to docs)
  * We are using [draft7_format_checker](https://python-jsonschema.readthedocs.io/en/latest/validate/#validating-formats) which means we can check common formats such as ipv4 and regex
  * Currently need to save the Excel file before we can read back table data.
    Need to store table data natively and use this for validation


