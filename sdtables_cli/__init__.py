from sdtables_cli.display import Display
from sdtables_cli.validate import Validate
from sdtables_cli.build import Build

__version__ = "2.0.2"

name = 'SDtables CLI'
description = 'ACLI wrapper for sdtables'
usage = 'sdtables <command> [<args>]'
model = {
    'display': Display,
    'validate': Validate,
    'build': Build
}
