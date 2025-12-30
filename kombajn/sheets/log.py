"""
Arkusz Dziennik.

Główny arkusz do codziennego logowania danych treningowych i zdrowotnych.
"""

from typing import Dict, List

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from kombajn.config import (
    LOG_HEADERS,
    LOG_INPUT_COLUMNS,
    LOG_SECTION_END_COLUMNS,
    LOG_COLUMN_WIDTHS,
    SHEET_CONFIG,
)
from kombajn.sheets.base import BaseSheet


class LogSheet(BaseSheet):
    """
    Arkusz dziennika treningowego.
    
    Zawiera kolumny dla:
    - Daty i tygodnia
    - Fizjologii (waga, RHR, HRV, sen)
    - Treningu (kalorie, czas, jakość)
    - Dolegliwości
    - Obliczeń kalorycznych (TDEE, cel kcal)
    - Spożycia i bilansu
    - Celów i spożycia makroskładników
    - Notatek dodatkowych
    """
    
    def __init__(self, workbook: Workbook) -> None:
        """Inicjalizuje arkusz Dziennik."""
        super().__init__(workbook, "Dziennik")
    
    def create(self) -> Worksheet:
        """
        Tworzy arkusz Dziennik.
        
        Returns:
            Utworzony arkusz
        """
        ws = self._create_worksheet()
        
        self._add_headers(ws)
        self._format_data_rows(ws)
        self._add_formulas(ws)
        self._add_date_column(ws)
        self._add_info_note(ws)
        self._set_column_widths(LOG_COLUMN_WIDTHS)
        
        # Zamrożenie pierwszego wiersza
        ws.freeze_panes = 'A2'
        
        return ws
    
    def _add_headers(self, ws: Worksheet) -> None:
        """Dodaje nagłówki kolumn."""
        for i, header in enumerate(LOG_HEADERS, 1):
            cell = ws.cell(row=1, column=i)
            cell.value = header
            
            # Styl nagłówka
            self.styles.apply_header_style(cell)
            
            # Gruba ramka dla końca sekcji
            if i in LOG_SECTION_END_COLUMNS:
                cell.border = self.styles.thick_right_border
    
    def _format_data_rows(self, ws: Worksheet) -> None:
        """Formatuje wiersze danych (żółte/szare tło, ramki)."""
        max_rows = SHEET_CONFIG.MAX_LOG_ROWS + 1  # +1 dla nagłówka
        
        for i, _ in enumerate(LOG_HEADERS, 1):
            is_input = i in LOG_INPUT_COLUMNS
            is_section_end = i in LOG_SECTION_END_COLUMNS
            border = self.styles.get_section_border(is_section_end)
            fill = self.styles.input_fill if is_input else self.styles.formula_fill
            
            for row in range(2, max_rows + 1):
                cell = ws.cell(row=row, column=i)
                cell.fill = fill
                cell.border = border
    
    def _add_formulas(self, ws: Worksheet) -> None:
        """Dodaje formuły do wiersza 2."""
        formulas: Dict[str, str] = {
            # Tydzień
            'B2': '=IF(ISNUMBER(A2), WEEKNUM(A2, 2), "")',
            
            # Waga średnia 7-dniowa
            'D2': '=IF(ISNUMBER(C2), AVERAGE(C2:INDEX(C:C, MAX(2, ROW()-6))), "")',
            
            # TDEE: baza + trening
            'O2': (
                "=IF(ISNUMBER(J2), 'Ustawienia i Cele'!$B$6 + J2, "
                "'Ustawienia i Cele'!$B$6)"
            ),
            
            # CEL Kcal: TDEE - deficyt
            'P2': "=O2 - 'Ustawienia i Cele'!$B$9",
            
            # Bilans: spożyte - cel
            'R2': '=IF(ISBLANK(Q2), "", Q2 - P2)',
            
            # CEL Białko: współczynnik * waga
            'S2': (
                '=IF(OR(C2="",C2=0), "", '
                "IFERROR('Ustawienia i Cele'!$B$11 * C2, 0))"
            ),
            
            # CEL Tłuszcze: % TDEE / 9 (kcal na gram)
            'T2': "=IFERROR((P2 * 'Ustawienia i Cele'!$B$12) / 9, 0)",
            
            # CEL Węgle: pozostałe kcal / 4
            'U2': '=IFERROR((P2 - (S2*4) - (T2*9)) / 4, 0)',
        }
        
        for cell_ref, formula in formulas.items():
            ws[cell_ref] = formula
            ws[cell_ref].font = self.styles.formula_font
    
    def _add_date_column(self, ws: Worksheet) -> None:
        """Dodaje kolumnę dat z automatycznym wypełnianiem."""
        # A2 - data startowa (do wpisania)
        ws['A2'] = ""
        ws['A2'].fill = self.styles.input_fill
        ws['A2'].number_format = 'yyyy-mm-dd'
        
        # Automatyczne wypełnianie kolejnych dat
        for row in range(3, SHEET_CONFIG.INITIAL_DAYS_COUNT + 2):
            ws[f'A{row}'] = '=IF(ISBLANK($A$2), "", $A$2 + (ROW()-2))'
            ws[f'A{row}'].number_format = 'yyyy-mm-dd'
    
    def _add_info_note(self, ws: Worksheet) -> None:
        """Dodaje notatkę informacyjną."""
        # Kolumna AC (29) - notatka
        ws['AC2'] = "Wypełnij żółte pola. Szare liczą się same."
        self.styles.apply_info_style(ws['AC2'])
        ws['AC2'].fill = self.styles.input_fill
