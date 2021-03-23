import sdtables_cli as prog
from pyfiglet import Figlet
import argparse
import logging
import os

logger = logging.getLogger('main')


def setup_logging():
    log_level = os.environ.get('LOGGING', 'INFO')
    level = logging.getLevelName(log_level.upper())
    logger.setLevel(level)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s')
    ch = logging.StreamHandler()
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    logger.propagate = False


setup_logging()


def main():
    f = Figlet(justify='center')
    print(f.renderText(prog.name))
    parser = argparse.ArgumentParser(
        prog=prog.name,
        description=prog.description,
        usage=prog.usage
    )
    parser.set_defaults(func=lambda args: parser.print_help())
    subparsers = parser.add_subparsers(help='Optional commands:')
    for key, value in prog.model.items():
        opt = value.add_args(key, subparsers)
        opt.set_defaults(func=value)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
