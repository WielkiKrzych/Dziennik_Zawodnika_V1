"""
Funkcje narzędziowe dla Kombajnu Triathlonisty.

Ten moduł zawiera funkcje pomocnicze do:
- Walidacji i bezpieczeństwa ścieżek plików
- Konfiguracji logowania
- Innych operacji wspólnych
"""

import logging
import re
import sys
from pathlib import Path
from typing import Optional

from openpyxl import Workbook


def setup_logging(
    level: int = logging.INFO,
    format_string: Optional[str] = None
) -> logging.Logger:
    """
    Konfiguruje logowanie dla aplikacji.
    
    Args:
        level: Poziom logowania (domyślnie INFO)
        format_string: Opcjonalny format wiadomości
        
    Returns:
        Skonfigurowany logger
    """
    if format_string is None:
        format_string = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    logging.basicConfig(
        level=level,
        format=format_string,
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger("kombajn")


def sanitize_filename(filename: str) -> str:
    """
    Czyści nazwę pliku z niebezpiecznych znaków.
    
    Args:
        filename: Oryginalna nazwa pliku
        
    Returns:
        Bezpieczna nazwa pliku
        
    Raises:
        ValueError: Gdy nazwa pliku jest pusta lub nieprawidłowa
    """
    if not filename:
        raise ValueError("Nazwa pliku nie może być pusta")
    
    # Usuń komponenty ścieżki (Path Traversal protection)
    safe_name = Path(filename).name
    
    # Usuń niebezpieczne znaki
    safe_name = re.sub(r'[<>:"/\\|?*]', '_', safe_name)
    
    # Usuń wiodące/końcowe spacje i kropki
    safe_name = safe_name.strip('. ')
    
    if not safe_name:
        raise ValueError("Nazwa pliku zawiera tylko niedozwolone znaki")
    
    return safe_name


def validate_output_path(
    filename: str,
    output_dir: Optional[Path] = None
) -> Path:
    """
    Waliduje i tworzy bezpieczną ścieżkę wyjściową.
    
    Args:
        filename: Nazwa pliku do zapisania
        output_dir: Katalog wyjściowy (domyślnie bieżący katalog)
        
    Returns:
        Zwalidowana ścieżka absolutna
        
    Raises:
        ValueError: Gdy ścieżka jest niebezpieczna
        PermissionError: Gdy brak uprawnień do zapisu
    """
    if output_dir is None:
        output_dir = Path.cwd()
    
    output_dir = output_dir.resolve()
    
    # Sanityzuj nazwę pliku
    safe_name = sanitize_filename(filename)
    
    # Upewnij się, że ma rozszerzenie .xlsx
    if not safe_name.lower().endswith('.xlsx'):
        safe_name += '.xlsx'
    
    # Utwórz pełną ścieżkę
    output_path = (output_dir / safe_name).resolve()
    
    # Sprawdź czy ścieżka nadal jest w dozwolonym katalogu (Path Traversal check)
    try:
        output_path.relative_to(output_dir)
    except ValueError:
        raise ValueError(
            f"Niedozwolona ścieżka pliku: próba wyjścia poza katalog {output_dir}"
        )
    
    # Sprawdź uprawnienia do zapisu
    if output_dir.exists() and not output_dir.is_dir():
        raise ValueError(f"Ścieżka {output_dir} nie jest katalogiem")
    
    if not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=True)
    
    return output_path


def safe_save_workbook(
    workbook: Workbook,
    filename: str,
    output_dir: Optional[Path] = None,
    logger: Optional[logging.Logger] = None
) -> Path:
    """
    Bezpiecznie zapisuje skoroszyt Excel.
    
    Args:
        workbook: Skoroszyt do zapisania
        filename: Nazwa pliku
        output_dir: Katalog wyjściowy
        logger: Logger do komunikatów
        
    Returns:
        Ścieżka do zapisanego pliku
        
    Raises:
        ValueError: Przy problemach z walidacją ścieżki
        PermissionError: Przy braku uprawnień
        Exception: Inne błędy zapisu
    """
    if logger is None:
        logger = logging.getLogger("kombajn")
    
    output_path = validate_output_path(filename, output_dir)
    
    logger.info(f"Zapisywanie do: {output_path}")
    
    try:
        workbook.save(output_path)
        logger.info(f"Plik zapisany pomyślnie: {output_path}")
        return output_path
    except PermissionError:
        logger.error(f"Brak uprawnień do zapisu: {output_path}")
        raise PermissionError(
            f"Nie można zapisać pliku '{output_path}'. "
            "Sprawdź czy plik nie jest otwarty w innym programie."
        )
    except Exception as e:
        logger.error(f"Błąd zapisu pliku: {e}")
        raise
