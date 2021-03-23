import logging
from sdtables.sdtables import SdTables
from tabulate import tabulate
import sdtables_cli.common as common

logger = logging.getLogger('main.{}'.format(__name__))


class Display:
    def __init__(self, args):
        # Run setup tasks
        self.args = args

        # Run the command with args
        self.run()

    @staticmethod
    def add_args(_key, _subparsers):
        _args = _subparsers.add_parser(_key, help='use sdtables display -h for help')
        _args.add_argument('--output', default='grid', help='Pythons tabulate module output format (default=grid)')
        _args.add_argument('--input', required=True, help='Path to .xlsx file as input')
        return _args

    def run(self):
        tables = SdTables()
        tables.load_xlsx_file(self.args.input)
        _tables = tables.get_all_tables_as_dict()
        for _sheetname, _tables in _tables.items():
            print('Worksheet: {}'.format(_sheetname))
            for _tablename, _data in _tables.items():
                print('\nTable Name: {}'.format(_tablename))
                print(tabulate(_data, headers='keys', tablefmt=self.args.output))
