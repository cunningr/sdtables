import logging
from sdtables.sdtables import SdTables
from tabulate import tabulate
import yaml, json
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
        _args.add_argument('--format', default='grid', help='(yaml|json|Pythons tabulate module output format (default=grid))')
        _args.add_argument('--output', help='Output file name (default=stdout)')
        _args.add_argument('--input', required=True, help='Path to .xlsx file as input')
        return _args

    def run(self):
        tables = SdTables()
        tables.load_xlsx_file(self.args.input)
        _tables = tables.get_all_tables_as_dict()
        if self.args.format == 'yaml':
            if self.args.output:
                with open(self.args.output, 'w') as file:
                    yaml.dump(_tables, file)
            else:
                print(yaml.dump(_tables))
        elif self.args.format == 'json':
            if self.args.output:
                with open(self.args.output, 'w') as file:
                    json.dump(_tables, file, indent=4, default=str)
            else:
                print(json.dumps(_tables, indent=4, default=str))
        else:
            for _sheetname, _tables in _tables.items():
                if self.args.output:
                    with open(self.args.output, 'w') as file:
                        file.write('Worksheet: {}\n'.format(_sheetname))
                        for _tablename, _data in _tables.items():
                            file.write('\nTable Name: {}\n'.format(_tablename))
                            file.write(tabulate(_data, headers='keys', tablefmt=self.args.format))
                            file.write('\n')
                else:
                    print('Worksheet: {}'.format(_sheetname))
                    for _tablename, _data in _tables.items():
                        print('\nTable Name: {}'.format(_tablename))
                        print(tabulate(_data, headers='keys', tablefmt=self.args.format))
