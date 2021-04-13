from sdtables_cli.display import Display
from sdtables_cli.validate import Validate
from sdtables_cli.build import Build

name = 'SDtables CLI'
description = 'ACLI wrapper for sdtables'
usage = 'sdtables <command> [<args>]'
model = {
    'display': Display,
    'validate': Validate,
    'build': Build
}
