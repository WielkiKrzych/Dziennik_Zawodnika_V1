"""
Arkusz Dashboard.

Podsumowania tygodniowe z formułami agregującymi dane z dziennika.
"""

import datetime
from typing import Dict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

from kombajn.sheets.base import BaseSheet


class DashboardSheet(BaseSheet):
    """
    Arkusz dashboardu z podsumowaniami tygodniowymi.
    
    Zawiera:
    - Wybór numeru tygodnia
    - Średnie wartości: waga, kalorie, bilans
    - Suma czasu treningu
    - Średnie: jakość snu, samopoczucie
    - Instrukcję tworzenia wykresów
    """
    
    def __init__(self, workbook: Workbook) -> None:
        """Inicjalizuje arkusz Dashboard."""
        super().__init__(workbook, "Dashboard")
    
    def create(self) -> Worksheet:
        """
        Tworzy arkusz Dashboard.
        
        Returns:
            Utworzony arkusz
        """
        ws = self._create_worksheet()
        
        self._add_title(ws)
        self._add_week_selector(ws)
        self._add_summary_table(ws)
        self._add_chart_instructions(ws)
        self._set_column_widths([30, 20, 50])
        
        return ws
    
    def _add_title(self, ws: Worksheet) -> None:
        """Dodaje tytuł dashboardu."""
        ws['A1'] = "PODSUMOWANIE TYGODNIOWE"
        ws['A1'].font = Font(bold=True, size=16)
    
    def _add_week_selector(self, ws: Worksheet) -> None:
        """Dodaje selektor numeru tygodnia."""
        ws['A3'] = "Wpisz nr tygodnia:"
        ws['A3'].font = Font(bold=True)
        
        # Domyślnie aktualny tydzień
        current_week = datetime.date.today().isocalendar()[1]
        ws['B3'] = current_week
        ws['B3'].fill = self.styles.input_fill
        ws['B3'].font = Font(bold=True)
    
    def _add_summary_table(self, ws: Worksheet) -> None:
        """Dodaje tabelę podsumowań."""
        # Nagłówki tabeli
        headers: Dict[str, str] = {
            'A5': "Wskaźnik",
            'B5': "Średnia / Suma",
            'C5': "Komentarz",
        }
        
        for cell_ref, value in headers.items():
            ws[cell_ref] = value
            self.styles.apply_header_style(ws[cell_ref])
        
        # Etykiety wskaźników
        labels: Dict[str, str] = {
            'A6': "Średnia waga (kg)",
            'A7': "Śr. Kcal (Spożyte)",
            'A8': "Śr. Bilans Kcal (dnia)",
            'A9': "Łączny Czas Treningu (h)",
            'A10': "Śr. Jakość Snu (1-5)",
            'A11': "Śr. Samopoczucie (1-5)",
        }
        
        for cell_ref, value in labels.items():
            ws[cell_ref] = value
            ws[cell_ref].font = Font(bold=True)
        
        # Formuły agregujące
        formulas: Dict[str, str] = {
            # Średnia waga (kolumna C)
            'B6': "=IFERROR(AVERAGEIFS('Dziennik'!C:C, 'Dziennik'!B:B, $B$3), \"Brak danych\")",
            
            # Średnie spożyte kcal (kolumna Q)
            'B7': "=IFERROR(AVERAGEIFS('Dziennik'!Q:Q, 'Dziennik'!B:B, $B$3), \"Brak danych\")",
            
            # Średni bilans kcal (kolumna R)
            'B8': "=IFERROR(AVERAGEIFS('Dziennik'!R:R, 'Dziennik'!B:B, $B$3), \"Brak danych\")",
            
            # Suma czasu treningu (kolumna L) / 60 min = godziny
            'B9': "=IFERROR(SUMIFS('Dziennik'!L:L, 'Dziennik'!B:B, $B$3) / 60, \"Brak danych\")",
            
            # Średnia jakość snu (kolumna H)
            'B10': "=IFERROR(AVERAGEIFS('Dziennik'!H:H, 'Dziennik'!B:B, $B$3), \"Brak danych\")",
            
            # Średnie samopoczucie (kolumna I)
            'B11': "=IFERROR(AVERAGEIFS('Dziennik'!I:I, 'Dziennik'!B:B, $B$3), \"Brak danych\")",
        }
        
        for cell_ref, formula in formulas.items():
            ws[cell_ref] = formula
            self.styles.apply_formula_style(ws[cell_ref])
    
    def _add_chart_instructions(self, ws: Worksheet) -> None:
        """Dodaje instrukcję tworzenia wykresów."""
        ws['C16'] = "INSTRUKCJA DO WYKRESÓW"
        ws['C16'].font = Font(bold=True)
        
        instructions = (
            "1. Przejdź do zakładki 'Dziennik'.\n"
            "2. Zaznacz kolumny (np. 'Data' i 'Waga (śr. 7-dniowa)').\n"
            "3. Wybierz Wstawianie -> Wykres.\n"
            "4. Wytnij (Ctrl+X) i wklej (Ctrl+V) go tutaj."
        )
        
        ws['C17'] = instructions
        ws['C17'].alignment = Alignment(wrap_text=True, vertical="top")
        self.styles.apply_info_style(ws['C17'])
        ws.row_dimensions[17].height = 70
