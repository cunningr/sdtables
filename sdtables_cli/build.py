import logging
import os
from sdtables.sdtables import SdTables
import yaml, json
import sdtables_cli.common as common

logger = logging.getLogger('main.{}'.format(__name__))


class Build:
    def __init__(self, args):
        # Run setup tasks
        self.args = args
        self.schema = {}

        # Run the command with args
        self.run()

    @staticmethod
    def add_args(_key, _subparsers):
        _args = _subparsers.add_parser(_key, help='use sdtables display -h for help')
        _args.add_argument('--schema', required=True, help='Path to schema file(s).  Can be <file> or <dir>')
        _args.add_argument('--output', required=True, help='Path to output .xlsx file')
        _args.add_argument('--format', default='yaml', help='Schema file format (json|yaml) (default=yaml)')
        _args.add_argument('--noseed', action='store_true', help='Exclude seed data if present in schema file')
        _args.add_argument('--filename-as-sheet', action='store_true', help='Use the schema filename as the xlsx sheetname')
        return _args

    def run(self):
        if os.path.isfile(self.args.schema):
            self.schema.update(self._load_schema_from_file(self.args.schema))
        elif os.path.isdir(self.args.schema):
            if self.args.format == 'yaml':
                l_files = [x for x in os.listdir(self.args.schema) if x.endswith(".yaml") or x.endswith(".yml")]
            elif self.args.format == 'json':
                l_files = [x for x in os.listdir(self.args.schema) if x.endswith(".json")]
            for file in l_files:
                _file_path = '{}/{}'.format(self.args.schema, file)
                self.schema.update(self._load_schema_from_file(_file_path))

        tables = SdTables()

        for _name, schema in self.schema.items():
            if self.args.noseed:
                data = None
            else:
                data = schema.get('seed', None)

            if self.args.filename_as_sheet:
                worksheet_name = schema.get('file', None)
            else:
                worksheet_name = schema.get('worksheet', 'main')

            tables.add_xlsx_table_from_schema(_name, schema, data=data, worksheet_name=worksheet_name)

        # Save the xlsx workbook
        tables.save_xlsx(self.args.output)

    def _load_schema_from_file(self, path):
        with open(path, 'r') as file:
            if '.yaml' in path or '.yml' in path:
                _dict = yaml.load(file, Loader=yaml.FullLoader)
            elif '.json' in path:
                _dict = json.load(file)
            else:
                print(path)

        _file_name, _ = self._filename_from_path(path).split('.', 1)
        _dict.update({'file': _file_name})
        if _dict.get('name'):
            _schema_name = _dict['name']
            return {_dict['name']: _dict}
        else:
            return {_file_name: _dict}

    @staticmethod
    def _filename_from_path(path):
        return path.split('/')[-1]
