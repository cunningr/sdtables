from sdtables.sdtables import SdTables

tables = SdTables()
tables.load_xlsx_file('./example.xlsx')

# # Tables names by worksheet
print('\nTable Names:')
print(tables.table_names)

# Get a table as dictionary
table_name = 'super.heros'
print('\nTable: {}'.format(table_name))
print(tables.get_table_as_dict(table_name))

# Get all tables as a dictionary and dump to json
print('\nAll Tables:')
print(tables.get_all_tables_as_dict())