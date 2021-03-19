# coding: utf-8


"""
    xlTables - Load/generate table data with Excel
    from python dictionary structures

    cunningr - 2020

    Requires openpyxl >= 2.6.2, jsonschema


"""

import openpyxl
from openpyxl import Workbook
from sdtables import xlTables


class XlsxSchema:
    def __init__(self, wb=None):
        self.table_names = []
        # self._tables = {}
        # self.schema_names = []
        # self.schemas = {}
        # self._dict = {}
        if wb is not None:
            self.wb = wb
            self.sheetnames = self.wb.sheetnames
            self._get_table_data()
        else:
            self.wb = Workbook()
            ws = self.wb.active
            self.wb.remove(ws)
            self.sheetnames = self.wb.sheetnames

        self.workbook_dict = {}
        self.workbook_dict_flat = {}
        self.string_only = False
        self.fill_empty = False

    @classmethod
    def load_xlsx_schema(cls, xlsx, data_only=False):
        wb = openpyxl.load_workbook(filename=xlsx, data_only=data_only)
        return cls(wb)

    def _get_table_data(self):
        """
        Internal function used to index tables from loaded xlsx file during initialisation
        :return:
        """
        _table_names = []
        _tables = {}
        for sheet in self.sheetnames:
            for table in self.wb[sheet].tables.values():
                _table_names.append((sheet, table.name))
                # _tables.update({table.name: table})

        self.table_names = _table_names
        # self._tables = _tables

    def get_table_as_dict(self, worksheet_name, table_name):
        """
        Takes a worksheet name and table name and returns the data as list of dictionaries

        Args:
            worksheet_name: Openpyxl worksheet name
            table_name: Openpyxl table name

        Returns:
            A list of dictionaries (rows)

        """
        ws = self.wb[worksheet_name]
        return xlTables.build_dict_from_table(ws, table_name)

    def get_all_tables_as_dict(self, flatten=False, squash=False):
        """
        Returns all table data from the Openpyxl workbook object.  By default each table is nested in a dictionary
        using the worksheet names as keys E.g.

        { "worksheet_name":
            [
                { "table_name": [{"col1": "value", "col2": "value"}]}
            ]
        }

        Args:
            flatten: Removes the worksheet_name hierarchy from the returned dictionary
            squash: Replaces the table_name with the worksheet_name.
                    Only one table per worksheet allowed and ignores additional tables

        Returns:
            A list of dictionaries (rows)

        """

        _dict = {}
        for table in self.table_names:
            worksheet_name, table_name = table
            ws = self.wb[worksheet_name]
            table_dict = xlTables.build_dict_from_table(ws, table_name)
            print(table_dict)

            if flatten:
                if squash:
                    print('ERROR: Do not set flatten=True and squash=True together')
                    return
                _dict.update(table_dict)
            elif squash:
                _dict_key = list(table_dict.keys())[0]
                _dict.update({worksheet_name: table_dict[table_name]})
            else:
                if not _dict.get(worksheet_name):
                    _dict.update({worksheet_name: {}})
                _dict[worksheet_name].update(table_dict)

        return _dict
    #
    # def _build_dict_from_table(self, table, name=None, string_only=False, fill_empty=False):
    #     """
    #     Internal function for building a dictionary from an Excel table definition
    #
    #     Args:
    #         table: internal dict with table sheetname as key and Openpyxl table object as value
    #         name: The name to be used as the returned dictionary key (defaults to table.name)
    #
    #     Returns:
    #         The sheet name and the table data represented as a python dict with the table name as key
    #
    #     """
    #
    #     if len(table.keys()) != 1:
    #         print('raise exception')
    #         return
    #     else:
    #         sheetname = list(table.keys())[0]
    #         _table_def = table[sheetname]
    #         _worksheet = self.wb[sheetname]
    #
    #     if name is None:
    #         name = _table_def.name
    #
    #     # Get the cell range of the table
    #     _table_range = _table_def.ref
    #
    #     _keys = []
    #
    #     for _column in _table_def.tableColumns:
    #         _keys.append(_column.name)
    #
    #     _num_columns = len(_keys)
    #     _row_width = len(_worksheet[_table_range][0])
    #     if _num_columns != _row_width:
    #         print('ERROR: Key count and row elements are not equal' + _num_columns, _row_width)
    #
    #     _new_dict = {name: {}}
    #     _rows_list = []
    #
    #     for _row in _worksheet[_table_range]:
    #         _row_dict = {}
    #         for _cell, _key in zip(_row, _keys):
    #             if _cell.value == _key:
    #                 # Pass over headers where cell.value equal key
    #                 pass
    #             else:
    #                 if fill_empty == True and _cell.value == None:
    #                     _row_dict[_key] = ""
    #                 elif string_only == True:
    #                     _row_dict[_key] = str(_cell.value).lstrip().rstrip()
    #                 else:
    #                     _row_dict.update({_key: _cell.value})
    #
    #         if bool(_row_dict):
    #             _rows_list.append(_row_dict)
    #
    #     _new_dict[name] = _rows_list
    #
    #     return sheetname, _new_dict

    def create_table_from_data(self, name, data, sheetname='Sheet1', table_style='TableStyleMedium2', row_offset=2, col_offset=1):
        if type(name) is not str or type(data) is not list:
            print('ERROR: table name must be of type str and data of type list')
        if sheetname not in self.wb.sheetnames:
            _ws = self.wb.create_sheet(sheetname)
        else:
            _ws = self.wb[sheetname]

        schema = {name: {'properties': xlTables.build_schema_from_row(data[0])}}
        xlTables.add_schema_table_to_worksheet(_ws, name, schema[name], data=data, table_style=table_style, row_offset=row_offset, col_offset=col_offset)
        self._get_table_data()

    def create_table_from_schema(self, name, schema, sheetname='default', data=None, table_style='TableStyleMedium2', row_offset=2, col_offset=1):
        if type(name) is not str or type(schema) is not dict:
            print('ERROR: table name must be of type str and schema of type dict')
        if sheetname not in self.wb.sheetnames:
            _ws = self.wb.create_sheet(sheetname)
        else:
            _ws = self.wb[sheetname]

        return xlTables.add_schema_table_to_worksheet(_ws, name, schema, data=data, table_style='TableStyleMedium2', row_offset=2, col_offset=1)

    def save_xlsx(self, filename):
        xlsx_filename = '{}.xlsx'.format(filename)
        self.wb.save(xlsx_filename)

# Retrieve a list of schema names under a given worksheet
# list(filter(lambda item: "network_settings" in item.keys(), meme.schemanames))