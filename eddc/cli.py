"""
Moduł CLI do obsługi aplikacji w wierszu poleceń.
"""

import sys
from eddc.utils import create_template_file
from eddc.file_processor import update_xlsx


def run_cli(args):
    """
    Główna funkcja CLI.

    :param args: Argumenty przekazane z main.py w formie Namespace.
    """

    if args.template:
        create_template_file(args.template)
        print(f"Szablon został zapisany w: {args.template}")
    elif args.file:
        output_file = args.output_file or args.file
        update_xlsx(args.file, output_file)
        print(f"Dane zostały zapisane w pliku {output_file}.")
    else:
        print("Brak wymaganych argumentów dla trybu CLI.\n"
              "Użyj --help, aby zobaczyć dostępne opcje.")