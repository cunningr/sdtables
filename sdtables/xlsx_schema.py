# coding: utf-8


"""
    xlTables - Load/generate table data with Excel
    from python dictionary structures

    cunningr - 2020

    Requires openpyxl >= 2.6.2, jsonschema


"""

import os
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

    def create_table_from_data(self, table_name, data, worksheet_name='Sheet1', table_style='TableStyleMedium2', row_offset=2, col_offset=1):
        if type(table_name) is not str or type(data) is not list:
            print('ERROR: table name must be of type str and data of type list')
        if worksheet_name not in self.wb.sheetnames:
            _ws = self.wb.create_sheet(worksheet_name)
        else:
            _ws = self.wb[worksheet_name]

        schema = {table_name: {'properties': xlTables.build_schema_from_row(data[0])}}
        xlTables.add_schema_table_to_worksheet(_ws, table_name, schema[table_name], data=data, table_style=table_style, row_offset=row_offset, col_offset=col_offset)
        self._get_table_data()

        return

    # Does not work
    def update_table_data(self, worksheet_name, table_name, data, append=True):
        xlTables.update_table_data(self.wb, worksheet_name, table_name, data, append=append)
        return

    def create_table_from_schema(self, table_name, schema, worksheet_name='default', data=None, table_style='TableStyleMedium2', row_offset=2, col_offset=1):
        if type(table_name) is not str or type(schema) is not dict:
            print('ERROR: table name must be of type str and schema of type dict')
        if worksheet_name not in self.wb.sheetnames:
            _ws = self.wb.create_sheet(worksheet_name)
        else:
            _ws = self.wb[worksheet_name]

        return xlTables.add_schema_table_to_worksheet(_ws, table_name, schema, data=data, table_style=table_style, row_offset=row_offset, col_offset=col_offset)

    def delete_table(self, worksheet_name, table_name, row_offset=2, col_offset=1):
        xlTables.delete_table(self.wb, worksheet_name, table_name, row_offset=row_offset, col_offset=col_offset)
        return

    def save_xlsx(self, filename):
        xlsx_filename = '{}/{}.xlsx'.format(os.getcwd(), filename)
        self.wb.save(xlsx_filename)

        return

# Retrieve a list of schema names under a given worksheet
# list(filter(lambda item: "network_settings" in item.keys(), meme.schemanames))