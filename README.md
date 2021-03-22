# sdtables

sdtables (schema defined tables) is a module providing convenient wrapper functions for working with and creating tabulated data from various sources including MS Excel.  We are using pythons jsonschema to build table structures and validate data

Install with pip3:

```
pip3 install sdtables
```

TODO: Update docs
Documentation can be [found here](https://cunningr.github.io/sdtables/)

## Features Summary

### Excel (xlsx)

| Issue | Description | Status |
|:---:|:---|:---:|
| | Load data to dict from Excel tables | Complete |
| | Add data tables to excel using first row keys as headers | Complete |
| | Add data tables to excel using a schema to define headers | Complete |
| | Add Excel data validation to excel using a schema (enum|tref|default values) | Complete |
| | Add schema with Excel data validation to excel without data | Complete |
| | Update data in existing table | On Hold |
| | Delete table | On Hold |
| | Validate table data against jsonschema draft7_format_checker | Complete |
| | Can tref schema definition be improved? | New |
| | Update docs | New |
| | Create examples | Complete |

Notes:

 * Update data in table -> file Issue
  * Without schema there is no data validation.  Can this be enhanced?
  * This breaks data validation since this is not moved when rows are inserted
 * Delete table -> file Issue
    This is broken as we don't remove the data or other formatting
 * Validate table data against schema (add to docs)
  * We are using [draft7_format_checker](https://python-jsonschema.readthedocs.io/en/latest/validate/#validating-formats) which means we can check common formats such as ipv4 and regex
  * Currently need to save the Excel file before we can read back table data.
    Need to store table data natively and use this for validation


