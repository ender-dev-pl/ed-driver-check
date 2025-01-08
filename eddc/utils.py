"""
Moduł narzędzi pomocniczych.
"""

import hashlib
from openpyxl import Workbook


def normalize_for_hash(first_name, last_name, document_number):
    """
    Normalizuje dane kierowcy i generuje hash MD5.

    :param first_name: Imię kierowcy.
    :param last_name: Nazwisko kierowcy.
    :param document_number: Numer dokumentu kierowcy.
    :return: Hash MD5 jako string.
    """
    combined = f"{first_name}{last_name}{document_number}".upper()
    normalized = combined.translate(str.maketrans("ĄĆĘŁŃÓŚŻŹ", "ACELNOSZZ"))
    return hashlib.md5(normalized.encode()).hexdigest().upper()


def create_template_file(output_path):
    """
    Tworzy pusty plik szablonu Excel.

    :param output_path: Ścieżka do pliku wyjściowego.
    """
    columns = ["imię", "nazwisko", "nr dokumentu", "wydający", "ważność", "status", "zmiana statusu"]

    # Tworzenie nowego skoroszytu
    wb = Workbook()
    ws = wb.active
    ws.title = "Szablon"

    # Dodanie nagłówków do arkusza
    for col_num, column_title in enumerate(columns, start=1):
        ws.cell(row=1, column=col_num, value=column_title)

    # Zapis do pliku
    wb.save(output_path)
