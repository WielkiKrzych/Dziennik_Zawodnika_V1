"""
Testy jednostkowe dla Kombajnu Triathlonisty.

Testuje tworzenie arkuszy, walidację ścieżek i konfigurację.
"""

import tempfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from kombajn.config import (
    DEFAULTS,
    SHEET_CONFIG,
    COLORS,
    LOG_HEADERS,
    CHO_HEADERS,
)
from kombajn.styles import ExcelStyles, DEFAULT_STYLES
from kombajn.utils import (
    sanitize_filename,
    validate_output_path,
    safe_save_workbook,
)
from kombajn.sheets import (
    SettingsSheet,
    LogSheet,
    DashboardSheet,
    CHOSourcesSheet,
)
from kombajn.main import create_workbook


class TestConfig:
    """Testy konfiguracji."""
    
    def test_defaults_values(self):
        """Sprawdza wartości domyślne."""
        assert DEFAULTS.BMR == 1800
        assert DEFAULTS.TEF == 200
        assert DEFAULTS.NEAT == 300
        assert DEFAULTS.DEFICIT == 500
        assert DEFAULTS.PROTEIN_RATIO == 2.0
        assert DEFAULTS.FAT_RATIO == 0.25
    
    def test_sheet_config(self):
        """Sprawdza konfigurację arkuszy."""
        assert SHEET_CONFIG.MAX_LOG_ROWS == 1000
        assert SHEET_CONFIG.INITIAL_DAYS_COUNT == 90
        assert SHEET_CONFIG.OUTPUT_FILENAME.endswith('.xlsx')
    
    def test_colors_defined(self):
        """Sprawdza czy kolory są zdefiniowane."""
        assert len(COLORS.HEADER_BG) == 6  # HEX bez #
        assert len(COLORS.HEADER_TEXT) == 6
        assert len(COLORS.INPUT_BG) == 6
        assert len(COLORS.FORMULA_BG) == 6
    
    def test_log_headers_count(self):
        """Sprawdza liczbę nagłówków dziennika."""
        assert len(LOG_HEADERS) == 27
    
    def test_cho_headers_count(self):
        """Sprawdza liczbę nagłówków CHO."""
        assert len(CHO_HEADERS) == 8


class TestStyles:
    """Testy stylów Excel."""
    
    def test_default_styles_exist(self):
        """Sprawdza istnienie domyślnych stylów."""
        assert DEFAULT_STYLES is not None
        assert isinstance(DEFAULT_STYLES, ExcelStyles)
    
    def test_styles_have_attributes(self):
        """Sprawdza atrybuty stylów."""
        styles = ExcelStyles()
        assert styles.header_font is not None
        assert styles.header_fill is not None
        assert styles.input_fill is not None
        assert styles.formula_fill is not None
        assert styles.thin_border is not None
        assert styles.thick_right_border is not None


class TestUtils:
    """Testy funkcji narzędziowych."""
    
    def test_sanitize_filename_simple(self):
        """Proste sanityzowanie nazwy pliku."""
        assert sanitize_filename("test.xlsx") == "test.xlsx"
        assert sanitize_filename("my file.xlsx") == "my file.xlsx"
    
    def test_sanitize_filename_removes_path_components(self):
        """Usuwa komponenty ścieżki."""
        assert sanitize_filename("../test.xlsx") == "test.xlsx"
        assert sanitize_filename("..\\..\\test.xlsx") == "test.xlsx"
        assert sanitize_filename("/etc/passwd") == "passwd"
    
    def test_sanitize_filename_removes_special_chars(self):
        """Usuwa niedozwolone znaki."""
        assert sanitize_filename("test<>:.xlsx") == "test___.xlsx"
    
    def test_sanitize_filename_empty_raises(self):
        """Pusta nazwa pliku rzuca wyjątek."""
        with pytest.raises(ValueError):
            sanitize_filename("")
    
    def test_validate_output_path_adds_extension(self):
        """Dodaje rozszerzenie .xlsx jeśli brakuje."""
        with tempfile.TemporaryDirectory() as tmpdir:
            path = validate_output_path("test", Path(tmpdir))
            assert path.suffix == ".xlsx"
    
    def test_validate_output_path_prevents_traversal(self):
        """Path traversal jest bezpiecznie obsługiwany - plik zostaje w dozwolonym katalogu."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Próba path traversal - powinna być bezpiecznie obsłużona
            path = validate_output_path("../../etc/passwd", Path(tmpdir))
            # Plik powinien pozostać w dozwolonym katalogu
            assert str(path).startswith(tmpdir)
            assert path.name == "passwd.xlsx"
    
    def test_safe_save_workbook(self):
        """Testuje bezpieczny zapis skoroszytu."""
        wb = Workbook()
        with tempfile.TemporaryDirectory() as tmpdir:
            path = safe_save_workbook(wb, "test.xlsx", Path(tmpdir))
            assert path.exists()
            assert path.suffix == ".xlsx"


class TestSheets:
    """Testy tworzenia arkuszy."""
    
    def test_settings_sheet_creation(self):
        """Testuje tworzenie arkusza Ustawienia."""
        wb = Workbook()
        sheet = SettingsSheet(wb).create()
        
        assert sheet.title == "Ustawienia i Cele"
        assert sheet['A1'].value == "TWOJE STAŁE DANE"
        assert sheet['B2'].value == DEFAULTS.BMR
    
    def test_log_sheet_creation(self):
        """Testuje tworzenie arkusza Dziennik."""
        wb = Workbook()
        # Najpierw musi być aktywny arkusz
        wb.active.title = "Temp"
        sheet = LogSheet(wb).create()
        
        assert sheet.title == "Dziennik"
        assert sheet.cell(row=1, column=1).value == "Data"
        assert sheet.cell(row=1, column=3).value == "Waga (kg)"
    
    def test_dashboard_sheet_creation(self):
        """Testuje tworzenie arkusza Dashboard."""
        wb = Workbook()
        wb.active.title = "Temp"
        sheet = DashboardSheet(wb).create()
        
        assert sheet.title == "Dashboard"
        assert sheet['A1'].value == "PODSUMOWANIE TYGODNIOWE"
    
    def test_cho_sources_sheet_creation(self):
        """Testuje tworzenie arkusza Źródła CHO."""
        wb = Workbook()
        wb.active.title = "Temp"
        sheet = CHOSourcesSheet(wb).create()
        
        assert sheet.title == "Źródła CHO"
        assert sheet.cell(row=1, column=1).value == "Nazwa produktu"


class TestMain:
    """Testy głównego przepływu."""
    
    def test_create_workbook_returns_workbook(self):
        """Testuje czy create_workbook zwraca skoroszyt."""
        wb = create_workbook()
        assert isinstance(wb, Workbook)
    
    def test_create_workbook_has_all_sheets(self):
        """Testuje czy wszystkie arkusze zostały utworzone."""
        wb = create_workbook()
        sheet_names = wb.sheetnames
        
        assert "Ustawienia i Cele" in sheet_names
        assert "Dziennik" in sheet_names
        assert "Dashboard" in sheet_names
        assert "Źródła CHO" in sheet_names
    
    def test_create_workbook_sheet_count(self):
        """Testuje liczbę utworzonych arkuszy."""
        wb = create_workbook()
        assert len(wb.sheetnames) == 4
    
    def test_full_workflow(self):
        """Testuje pełny przepływ: tworzenie i zapis."""
        wb = create_workbook()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            path = safe_save_workbook(wb, "full_test.xlsx", Path(tmpdir))
            
            assert path.exists()
            assert path.stat().st_size > 0
            
            # Sprawdź czy można otworzyć
            from openpyxl import load_workbook
            loaded = load_workbook(path)
            assert len(loaded.sheetnames) == 4


class TestFormulas:
    """Testy formuł Excel."""
    
    def test_log_sheet_has_formulas(self):
        """Sprawdza czy dziennik zawiera formuły."""
        wb = Workbook()
        wb.active.title = "Temp"
        sheet = LogSheet(wb).create()
        
        # Sprawdź formuły w wierszu 2
        assert sheet['B2'].value.startswith('=')  # Tydzień
        assert sheet['D2'].value.startswith('=')  # Średnia waga
        assert sheet['O2'].value.startswith('=')  # TDEE
    
    def test_settings_sheet_cpm_formula(self):
        """Sprawdza formułę CPM."""
        wb = Workbook()
        sheet = SettingsSheet(wb).create()
        
        assert sheet['B6'].value == "=SUM(B2:B4)"
    
    def test_dashboard_has_averageifs(self):
        """Sprawdza czy dashboard używa AVERAGEIFS."""
        wb = Workbook()
        wb.active.title = "Temp"
        sheet = DashboardSheet(wb).create()
        
        assert 'AVERAGEIFS' in sheet['B6'].value


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
