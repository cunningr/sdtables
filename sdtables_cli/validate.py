import logging
import os
from sdtables.sdtables import SdTables
import yaml, json
import sdtables_cli.common as common

logger = logging.getLogger('main.{}'.format(__name__))


class Validate:
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
        _args.add_argument('--input', required=True, help='Path to .xlsx file as input')
        _args.add_argument('--format', default='yaml', help='Schema file format (json|yaml) (default=yaml)')
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
        tables.load_xlsx_file(self.args.input)
        for _name, schema in self.schema.items():
            tables.add_schema(_name, schema)

        tables.validate_table_data(stdout=True)
        # print(json.dumps(self.schema, indent=4))

    @staticmethod
    def _load_schema_from_file(path):
        with open(path, 'r') as file:
            if '.yaml' in path or '.yml' in path:
                _dict = yaml.load(file, Loader=yaml.FullLoader)
            elif '.json' in path:
                _dict = json.load(file)
            else:
                print(path)

        if _dict.get('name'):
            _schema_name = _dict['name']
        else:
            _schema_name, _ = file.split('.', 2)

        return {_dict['name']: _dict}

