"""
Konfiguracja i stałe dla Kombajnu Triathlonisty.

Ten moduł zawiera wszystkie wartości domyślne, stałe konfiguracyjne
i parametry, które wcześniej były zakodowane na sztywno w kodzie.
"""

from dataclasses import dataclass
from typing import Dict, List, Tuple


# =============================================================================
# WARTOŚCI DOMYŚLNE DLA UŻYTKOWNIKA
# =============================================================================

@dataclass(frozen=True)
class DefaultValues:
    """Domyślne wartości metaboliczne i celów."""
    
    # Metabolizm bazowy
    BMR: int = 1800  # Basal Metabolic Rate (kcal)
    TEF: int = 200   # Thermic Effect of Food (kcal)
    NEAT: int = 300  # Non-Exercise Activity Thermogenesis (kcal)
    
    # Cele
    DEFICIT: int = 500           # Planowany deficyt kaloryczny
    PROTEIN_RATIO: float = 2.0   # Białko (g / kg masy ciała)
    FAT_RATIO: float = 0.25      # Tłuszcze (% TDEE)


DEFAULTS = DefaultValues()


# =============================================================================
# PARAMETRY ARKUSZY
# =============================================================================

@dataclass(frozen=True)
class SheetConfig:
    """Konfiguracja parametrów arkuszy."""
    
    # Liczba wierszy do przygotowania w dzienniku
    MAX_LOG_ROWS: int = 1000
    
    # Liczba dni do automatycznego wypełnienia dat
    INITIAL_DAYS_COUNT: int = 90
    
    # Nazwa pliku wyjściowego
    OUTPUT_FILENAME: str = "kombajn_triathlonisty_v2.xlsx"


SHEET_CONFIG = SheetConfig()


# =============================================================================
# KOLORY (HEX)
# =============================================================================

@dataclass(frozen=True)
class Colors:
    """Paleta kolorów używana w arkuszach."""
    
    # Nagłówki
    HEADER_BG: str = "4F81BD"      # Niebieski
    HEADER_TEXT: str = "FFFFFF"    # Biały
    
    # Komórki do wpisywania
    INPUT_BG: str = "FFFFCC"       # Jasny żółty
    
    # Komórki z formułami
    FORMULA_BG: str = "F2F2F2"     # Jasny szary
    
    # Tekst informacyjny
    INFO_TEXT: str = "808080"      # Szary


COLORS = Colors()


# =============================================================================
# NAGŁÓWKI DZIENNIKA
# =============================================================================

LOG_HEADERS: List[str] = [
    # Ogólne
    "Data", "Tydzień",
    # Fizjologia (Rano)
    "Waga (kg)", "Waga (śr. 7-dniowa)", "RHR", "HRV", "Sen (h)", "Jakość snu (1-5)", 
    "Samopoczucie (1-5)",
    # Trening (Wpisywane)
    "Trening (Kcal)", "Trening (Czas, min)", "Jakość Treningu (1-5)",
    # Dolegliwości (Fizjo)
    "Dolegliwości (Opis)", "Dolegliwości (Ból 1-5)",
    # Obliczenia Kaloryczne (Formuły)
    "TDEE (Szacowane)", "CEL Kcal (na dziś)",
    # Spożycie (Wpisywane)
    "Spożyte Kcal",
    # Bilans (Formuła)
    "Bilans Kcal (dnia)",
    # Obliczenia Makro (Formuły)
    "CEL Białko (g)", "CEL Tłuszcze (g)", "CEL Węgle (g)",
    # Spożycie Makro (Wpisywane)
    "Spożyte Białko (g)", "Spożyte Tłuszcze (g)", "Spożyte Węgle (g)",
    # Inne (Wpisywane)
    "Płyny Spożyte (L)", "Suplementy (Notatka)", "Notatki (Ogólne)"
]

# Kolumny do ręcznego wpisania (1-based index) - żółte tło
LOG_INPUT_COLUMNS: List[int] = [
    1, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
    18, 23, 24, 25, 26, 27, 28
]

# Kolumny kończące sekcje logiczne (gruba prawa krawędź)
LOG_SECTION_END_COLUMNS: List[int] = [
    2,   # Po Tydzień
    8,   # Po Jakość snu
    9,   # Po Samopoczucie
    12,  # Po Jakość Treningu
    14,  # Po Dolegliwości
    16,  # Po CEL Kcal
    17,  # Po Spożyte Kcal
    18,  # Po Bilans
    21,  # Po CEL Węgle
    24,  # Po Spożyte Węgle
    25,  # Po Płyny
    26,  # Po Suplementy
]

# Szerokości kolumn dziennika
LOG_COLUMN_WIDTHS: List[int] = [
    12, 8,   # Data, Tydzień
    10, 10, 8, 8, 8, 10,  # Waga, Waga śr, RHR, HRV, Sen, Jakość snu
    12,  # Samopoczucie
    12, 12, 12,  # Trening Kcal, Czas, Jakość
    25, 12,  # Dolegliwości Opis, Ból
    12, 12,  # TDEE, CEL Kcal
    12,  # Spożyte Kcal
    12,  # Bilans
    12, 12, 12,  # Cele Makro
    12, 12, 12,  # Spożyte Makro
    10, 20, 30   # Płyny, Suple, Notatki
]


# =============================================================================
# NAGŁÓWKI ŹRÓDEŁ CHO
# =============================================================================

CHO_HEADERS: List[str] = [
    "Nazwa produktu", "Porcja (g)", "CHO / 100g (g)", "kcal / 100g",
    "CHO w porcji (g)", "kcal w porcji", "Typ (żel/baton/napój/inne)", "Uwagi"
]

CHO_COLUMN_WIDTHS: List[int] = [30, 14, 14, 14, 16, 14, 20, 40]

# Przykładowe dane produktów CHO
CHO_SAMPLE_DATA: List[Tuple[str, int, float, int, str, str]] = [
    ("Banan", 100, 23, 89, "owoc", "Klasyczny wybór przed/po treningu"),
    ("Żel energetyczny", 40, 60, 240, "żel", "Szybka dawka CHO podczas wysiłku"),
    ("Rodzynki", 30, 79, 299, "suszone owoce", "Gęste źródło cukrów, wygodne w transporcie"),
    ("Daktyle (Medjool)", 24, 75, 277, "suszone owoce", "Szybkie i skoncentrowane źródło energii"),
    ("Miód", 15, 82, 304, "płynne", "Szybkie cukry, łatwe do dodania do napoju"),
    ("Chleb biały (kromka)", 30, 49, 265, "pieczywo", "Łatwo dostępne źródło węglowodanów"),
    ("Wafelek ryżowy", 9, 85, 387, "wafelek", "Niska gęstość energetyczna; szybkie węgle"),
    ("Baton energetyczny (np. Clif)", 68, 48, 400, "baton", 
     "Kombinacja węgli i tłuszczu — dłuższe uwalnianie energii"),
    ("Napój izotoniczny (Gatorade)", 250, 6.9, 26, "napój", 
     "Płynne źródło CHO i elektrolitów podczas treningu"),
    ("Żelki (gummy)", 40, 72.5, 475, "słodycze", "Szybkie cukry, dobre w krótkich wysiłkach"),
    ("Maltodekstryna", 30, 100, 400, "proszek", "Czyste węglowodany, szybko dostępne"),
    ("Fruktoza", 30, 100, 399, "cukier", "Wolniej metabolizowana niż glukoza; stosować z umiarem"),
    ("Mieszanka MD:FR (1:0.8)", 30, 100, 399, "mieszanka", 
     "Maltodekstryna:Fruktoza w stosunku 1:0.8 — szybkie uzupełnienie CHO"),
]
