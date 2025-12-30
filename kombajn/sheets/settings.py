"""
Arkusz Ustawienia i Cele.

Zawiera dane użytkownika: BMR, TEF, NEAT, cele kaloryczne i makroskładnikowe.
"""

from typing import Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from kombajn.config import DEFAULTS
from kombajn.sheets.base import BaseSheet


class SettingsSheet(BaseSheet):
    """
    Arkusz z ustawieniami i celami użytkownika.
    
    Zawiera:
    - Stałe dane metaboliczne (BMR, TEF, NEAT)
    - Obliczane CPM (baza)
    - Cele kaloryczne i makroskładnikowe
    - Dane ciągnięte z dziennika
    """
    
    def __init__(self, workbook: Workbook) -> None:
        """Inicjalizuje arkusz Ustawienia i Cele."""
        super().__init__(workbook, "Ustawienia i Cele")
    
    def create(self) -> Worksheet:
        """
        Tworzy arkusz Ustawienia i Cele.
        
        Returns:
            Utworzony arkusz
        """
        ws = self._create_worksheet(use_active=True)
        
        self._add_section_headers(ws)
        self._add_labels(ws)
        self._add_formulas(ws)
        self._add_input_cells(ws)
        self._add_info_texts(ws)
        self._set_column_widths([30, 15, 25])
        
        return ws
    
    def _add_section_headers(self, ws: Worksheet) -> None:
        """Dodaje nagłówki sekcji."""
        section_headers: Dict[str, str] = {
            'A1': "TWOJE STAŁE DANE",
            'A8': "TWOJE CELE",
            'A14': "DANE ŚCIĄGANE Z DZIENNIKA",
        }
        
        for cell_ref, value in section_headers.items():
            ws[cell_ref] = value
            ws[cell_ref].font = Font(bold=True, size=14)
            # Scalanie komórek dla nagłówków sekcji
            row = cell_ref[1:]
            ws.merge_cells(f'{cell_ref}:{get_column_letter(3)}{row}')
    
    def _add_labels(self, ws: Worksheet) -> None:
        """Dodaje etykiety pól."""
        labels: Dict[str, str] = {
            'A2': "BMR (kcal)",
            'A3': "TEF (kcal)",
            'A4': "NEAT (kcal)",
            'A6': "CPM (Baza)",
            'A9': "Planowany Deficyt (np. 500)",
            'A11': "CEL: Białko (g / kgmc)",
            'A12': "CEL: Tłuszcze (% TDEE)",
            'A15': "Aktualna waga (ostatni wpis)",
        }
        
        for cell_ref, value in labels.items():
            ws[cell_ref] = value
            ws[cell_ref].font = Font(bold=True)
    
    def _add_formulas(self, ws: Worksheet) -> None:
        """Dodaje formuły obliczeniowe."""
        # CPM (Baza) = BMR + TEF + NEAT
        ws['B6'] = "=SUM(B2:B4)"
        self.styles.apply_formula_style(ws['B6'])
        
        # Aktualna waga - ostatni wpis z dziennika
        ws['B15'] = (
            '=IFERROR(LOOKUP(2,1/(\'Dziennik\'!C:C<>"Błędna formuła"),'
            '\'Dziennik\'!C:C), "Brak danych")'
        )
        self.styles.apply_formula_style(ws['B15'])
    
    def _add_input_cells(self, ws: Worksheet) -> None:
        """Dodaje komórki do wpisywania z wartościami domyślnymi."""
        input_values: Dict[str, Tuple[int | float, None]] = {
            'B2': (DEFAULTS.BMR, None),
            'B3': (DEFAULTS.TEF, None),
            'B4': (DEFAULTS.NEAT, None),
            'B9': (DEFAULTS.DEFICIT, None),
            'B11': (DEFAULTS.PROTEIN_RATIO, None),
            'B12': (DEFAULTS.FAT_RATIO, None),
        }
        
        for cell_ref, (value, _) in input_values.items():
            ws[cell_ref].fill = self.styles.input_fill
            ws[cell_ref].value = value
    
    def _add_info_texts(self, ws: Worksheet) -> None:
        """Dodaje teksty informacyjne."""
        ws['C12'] = "(Wpisz 0.25 dla 25%)"
        self.styles.apply_info_style(ws['C12'])
