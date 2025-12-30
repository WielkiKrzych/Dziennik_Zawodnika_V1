"""
Główny moduł Kombajnu Triathlonisty.

Punkt wejścia do generowania pliku Excel z dziennikiem treningowym.
"""

import argparse
import logging
import sys
import traceback
from pathlib import Path
from typing import Optional

from openpyxl import Workbook

from kombajn.config import SHEET_CONFIG
from kombajn.sheets import (
    SettingsSheet,
    LogSheet,
    DashboardSheet,
    CHOSourcesSheet,
)
from kombajn.utils import safe_save_workbook, setup_logging


def create_workbook() -> Workbook:
    """
    Tworzy kompletny skoroszyt z wszystkimi arkuszami.
    
    Returns:
        Gotowy skoroszyt Excel
    """
    logger = logging.getLogger("kombajn")
    
    logger.info("Rozpoczynam tworzenie skoroszytu...")
    wb = Workbook()
    
    # Tworzenie arkuszy
    logger.info("Tworzę zakładkę [Ustawienia i Cele]...")
    SettingsSheet(wb).create()
    
    logger.info("Tworzę zakładkę [Dziennik]...")
    LogSheet(wb).create()
    
    logger.info("Tworzę zakładkę [Dashboard]...")
    DashboardSheet(wb).create()
    
    logger.info("Tworzę zakładkę [Źródła CHO]...")
    CHOSourcesSheet(wb).create()
    
    # Wymuszenie pełnego przeliczenia formuł przy otwieraniu
    try:
        wb.calculation.calcMode = 'auto'
    except AttributeError:
        pass
    
    try:
        wb.calculation_properties.fullCalcOnLoad = True
    except (AttributeError, Exception):
        # Starsze wersje openpyxl mogą nie mieć tego atrybutu
        pass
    
    logger.info("Skoroszyt utworzony pomyślnie.")
    return wb


def main(
    output_filename: Optional[str] = None,
    output_dir: Optional[Path] = None
) -> int:
    """
    Główna funkcja programu.
    
    Args:
        output_filename: Opcjonalna nazwa pliku wyjściowego
        output_dir: Opcjonalny katalog wyjściowy
        
    Returns:
        Kod wyjścia (0 = sukces, 1 = błąd)
    """
    logger = setup_logging()
    
    print("Cześć. Zaczynam tworzyć Twój 'kombajn v2'...")
    
    try:
        # Tworzenie skoroszytu
        wb = create_workbook()
        
        # Zapis pliku
        filename = output_filename or SHEET_CONFIG.OUTPUT_FILENAME
        output_path = safe_save_workbook(wb, filename, output_dir, logger)
        
        # Komunikat sukcesu
        print("-" * 60)
        print("GOTOWE! 🚀")
        print(f"Plik '{output_path.name}' został stworzony.")
        print("-" * 60)
        print("\nJak zacząć:")
        print("1. Otwórz plik i idź do [Ustawienia i Cele].")
        print("2. Idź do [Dziennika]. ŻÓŁTE pola wypełniasz ręcznie.")
        print("3. SZARE pola liczą się same. Przeciągnij formuły z wiersza 2 w dół.")
        
        return 0
        
    except ImportError as e:
        logger.error(f"Brak wymaganej biblioteki: {e}")
        print("[BŁĄD] Nie znaleziono biblioteki 'openpyxl'.")
        print("Uruchom w terminalu: pip install openpyxl")
        return 1
        
    except PermissionError as e:
        logger.error(f"Błąd uprawnień: {e}")
        print(f"[BŁĄD] {e}")
        return 1
        
    except ValueError as e:
        logger.error(f"Błąd walidacji: {e}")
        print(f"[BŁĄD] {e}")
        return 1
        
    except Exception as e:
        logger.error(f"Nieoczekiwany błąd: {e}")
        logger.debug(traceback.format_exc())
        print(f"Wystąpił nieoczekiwany błąd: {e}")
        return 1


def cli() -> None:
    """Interfejs linii poleceń."""
    parser = argparse.ArgumentParser(
        description="Kombajn Triathlonisty - Generator dziennika treningowego Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Przykłady użycia:
  python -m kombajn.main
  python -m kombajn.main -o moj_dziennik.xlsx
  python -m kombajn.main -o dziennik.xlsx -d C:\\Dokumenty
        """
    )
    
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help=f"Nazwa pliku wyjściowego (domyślnie: {SHEET_CONFIG.OUTPUT_FILENAME})"
    )
    
    parser.add_argument(
        "-d", "--directory",
        type=Path,
        default=None,
        help="Katalog wyjściowy (domyślnie: bieżący katalog)"
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Tryb szczegółowy (więcej logów)"
    )
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger("kombajn").setLevel(logging.DEBUG)
    
    exit_code = main(args.output, args.directory)
    sys.exit(exit_code)


if __name__ == "__main__":
    cli()
