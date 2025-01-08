"""
Główny moduł aplikacji, który dynamicznie uruchamia CLI lub GUI na podstawie argumentów wiersza poleceń.
"""

import argparse
import sys
from eddc import cli, gui


def setup_argparse():
    """
    Konfiguruje argumenty wiersza poleceń.

    :return: Przeparsowane argumenty.
    """
    parser = argparse.ArgumentParser(
        description="Aplikacja do sprawdzania uprawnień kierowców.",
        epilog="Przykład użycia:\n"
               "  python main.py --cli --file dane.xlsx\n"
               "  python main.py --gui",
        formatter_class=argparse.RawTextHelpFormatter
    )

    parser.add_argument(
        "--cli",
        action="store_true",
        help="Uruchamia aplikację w trybie wiersza poleceń (CLI)."
    )

    parser.add_argument(
        "--gui",
        action="store_true",
        help="Uruchamia aplikację w trybie graficznym (GUI)."
    )

    parser.add_argument(
        "--file",
        type=str,
        help="Ścieżka do istniejącego pliku Excel z danymi do sprawdzenia (tylko w trybie CLI)."
    )

    parser.add_argument(
        "--output_file",
        type=str,
        help="Ścieżka do wyjściowego pliku Excel (tylko w trybie CLI)."
    )

    parser.add_argument(
        "--template",
        type=str,
        help="Ścieżka do nowego pliku Excel zawierającego pusty szablon danych (tylko w trybie CLI)."
    )

    return parser.parse_args()


def main():
    """
    Główna funkcja aplikacji.
    """
    args = setup_argparse()

    # Sprawdź, czy uruchomić CLI
    if args.cli:
        cli_args = argparse.Namespace(
            file=args.file,
            output_file=args.output_file,
            template=args.template
        )
        cli.run_cli(cli_args)

    # Sprawdź, czy uruchomić GUI
    elif args.gui or not any(vars(args).values()):  # Brak argumentów = domyślnie GUI
        gui.run_gui()

    else:
        print("Nie rozpoznano trybu uruchamiania. Użyj --help, aby zobaczyć dostępne opcje.")


if __name__ == "__main__":
    main()
