"""
Arkusz Źródła CHO.

Baza produktów węglowodanowych z kalkulatorem porcji.
"""

from typing import List, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from kombajn.config import (
    CHO_HEADERS,
    CHO_COLUMN_WIDTHS,
    CHO_SAMPLE_DATA,
    SHEET_CONFIG,
)
from kombajn.sheets.base import BaseSheet


class CHOSourcesSheet(BaseSheet):
    """
    Arkusz źródeł węglowodanów.
    
    Zawiera bazę produktów z:
    - Nazwą produktu
    - Wielkością porcji
    - Zawartością CHO i kcal na 100g
    - Obliczeniami dla porcji
    - Typem i uwagami
    """
    
    def __init__(self, workbook: Workbook) -> None:
        """Inicjalizuje arkusz Źródła CHO."""
        super().__init__(workbook, "Źródła CHO")
    
    def create(self) -> Worksheet:
        """
        Tworzy arkusz Źródła CHO.
        
        Returns:
            Utworzony arkusz
        """
        ws = self._create_worksheet()
        
        self._add_headers(ws)
        self._format_input_columns(ws)
        self._add_sample_data(ws)
        self._set_column_widths(CHO_COLUMN_WIDTHS)
        
        # Zamrożenie pierwszego wiersza
        ws.freeze_panes = 'A2'
        
        return ws
    
    def _add_headers(self, ws: Worksheet) -> None:
        """Dodaje nagłówki kolumn."""
        for i, header in enumerate(CHO_HEADERS, 1):
            cell = ws.cell(row=1, column=i)
            cell.value = header
            self.styles.apply_header_style(cell)
    
    def _format_input_columns(self, ws: Worksheet) -> None:
        """Formatuje kolumny do wpisywania (żółte tło)."""
        # Kolumny 2, 3, 4 (Porcja, CHO/100g, kcal/100g)
        input_cols = [2, 3, 4]
        max_rows = SHEET_CONFIG.MAX_LOG_ROWS + 1
        
        for row in range(2, max_rows + 1):
            for col in input_cols:
                ws.cell(row=row, column=col).fill = self.styles.input_fill
    
    def _add_sample_data(self, ws: Worksheet) -> None:
        """Dodaje przykładowe dane produktów."""
        for row_idx, item in enumerate(CHO_SAMPLE_DATA, start=2):
            name, portion, cho_100g, kcal_100g, product_type, notes = item
            
            # Dane podstawowe
            ws.cell(row=row_idx, column=1).value = name
            ws.cell(row=row_idx, column=2).value = portion
            ws.cell(row=row_idx, column=3).value = cho_100g
            ws.cell(row=row_idx, column=4).value = kcal_100g
            ws.cell(row=row_idx, column=7).value = product_type
            ws.cell(row=row_idx, column=8).value = notes
            
            # Formuły dla obliczonych kolumn
            self._add_portion_formulas(ws, row_idx)
    
    def _add_portion_formulas(self, ws: Worksheet, row: int) -> None:
        """
        Dodaje formuły obliczające wartości dla porcji.
        
        Args:
            ws: Arkusz roboczy
            row: Numer wiersza
        """
        # CHO w porcji = (Porcja * CHO/100g) / 100
        cho_formula = f'=IF(OR(B{row}="",C{row}=""), "", B{row} * C{row} / 100)'
        ws.cell(row=row, column=5).value = cho_formula
        self.styles.apply_formula_style(ws.cell(row=row, column=5))
        
        # kcal w porcji = (Porcja * kcal/100g) / 100
        kcal_formula = f'=IF(OR(B{row}="",D{row}=""), "", B{row} * D{row} / 100)'
        ws.cell(row=row, column=6).value = kcal_formula
        self.styles.apply_formula_style(ws.cell(row=row, column=6))
