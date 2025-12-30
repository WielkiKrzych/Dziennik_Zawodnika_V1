"""
Główny moduł Dziennika Kolarza.

Punkt wejścia do generowania pliku Excel z dziennikiem treningowym
zgodnym z metrykami WKO5/INSCYD.
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
    PowerZonesSheet,
)
from kombajn.utils import safe_save_workbook, setup_logging


def create_workbook() -> Workbook:
    """
    Tworzy kompletny skoroszyt z wszystkimi arkuszami.
    
    Arkusze:
    - Ustawienia (profil mocy WKO5, profil metaboliczny INSCYD)
    - Dziennik (42 kolumny z metrykami power i PMC)
    - Dashboard (PMC Chart, podsumowania)
    - Strefy Mocy (7 stref Coggan)
    - Źródła CHO (baza produktów)
    
    Returns:
        Gotowy skoroszyt Excel
    """
    logger = logging.getLogger("kombajn")
    
    logger.info("Rozpoczynam tworzenie skoroszytu...")
    wb = Workbook()
    
    # Tworzenie arkuszy
    logger.info("Tworzę zakładkę [Ustawienia]...")
    SettingsSheet(wb).create()
    
    logger.info("Tworzę zakładkę [Dziennik]...")
    LogSheet(wb).create()
    
    logger.info("Tworzę zakładkę [Dashboard]...")
    DashboardSheet(wb).create()
    
    logger.info("Tworzę zakładkę [Strefy Mocy]...")
    PowerZonesSheet(wb).create()
    
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
    
    print("🚴 Dziennik Kolarza v3 - WKO5/INSCYD Edition")
    print("=" * 50)
    
    try:
        wb = create_workbook()
        
        filename = output_filename or SHEET_CONFIG.OUTPUT_FILENAME
        output_path = safe_save_workbook(wb, filename, output_dir, logger)
        
        print("-" * 50)
        print("GOTOWE! 🚀")
        print(f"Plik '{output_path.name}' został stworzony.")
        print("-" * 50)
        print("\n📖 Jak zacząć:")
        print("1. Otwórz plik i ustaw FTP w [Ustawienia]")
        print("2. Sprawdź [Strefy Mocy] - przeliczą się automatycznie")
        print("3. Wypełniaj [Dziennik] danymi z Garmina/Zwift")
        print("4. Śledź formę w [Dashboard] (CTL/ATL/TSB)")
        print("\n💡 Wskazówki:")
        print("• TSB +10 do +25 = gotowy na wyścig")
        print("• Tygodniowy TSS: 300-500 (amator), 500-800 (zaawansowany)")
        
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
        description="Dziennik Kolarza - Generator dziennika z metrykami WKO5/INSCYD",
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
