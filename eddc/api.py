"""
Moduł do obsługi komunikacji z API.
"""

import requests
from eddc.utils import normalize_for_hash

API_URL = "https://moj.gov.pl/nforms/api/UprawnieniaKierowcow/2.0.10/data/driver-permissions"


def process_driver_data(first_name, last_name, document_number):
    """
    Wysyła zapytanie do API w celu sprawdzenia uprawnień kierowcy.

    :param first_name: Imię kierowcy.
    :param last_name: Nazwisko kierowcy.
    :param document_number: Numer dokumentu kierowcy. (pole "nr blankietu")
    :return: JSON z odpowiedzią API lub słownik z kluczem "error" w przypadku błędu.
    """
    if not first_name or not last_name or not document_number:
        return {"error": "brak danych"}

    hash_value = normalize_for_hash(first_name, last_name, document_number)
    try:
        response = requests.get(f"{API_URL}?hashDanychWyszukiwania={hash_value}")
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        if response.status_code == 400:
            return {"error": "błąd danych"}
        return {"error": f"HTTP error: {http_err}"}
    except requests.exceptions.RequestException as req_err:
        return {"error": f"Błąd API: {req_err}"}
