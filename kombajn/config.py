"""
Konfiguracja i stałe dla Dziennika Kolarza.

Ten moduł zawiera wszystkie wartości domyślne, stałe konfiguracyjne
i parametry dla dziennika treningowego zgodnego z WKO5/INSCYD.
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
# POWER METRICS (WKO5)
# =============================================================================

@dataclass(frozen=True)
class PowerDefaults:
    """Domyślne wartości dla metryk mocy (WKO5)."""
    
    FTP: int = 250              # Functional Threshold Power (W)
    MAX_HR: int = 185           # Tętno maksymalne
    RESTING_HR: int = 50        # Tętno spoczynkowe
    WEIGHT_KG: float = 75.0     # Waga (kg) do obliczeń W/kg


POWER_DEFAULTS = PowerDefaults()


# =============================================================================
# METABOLIC PROFILING (INSCYD)
# =============================================================================

@dataclass(frozen=True)
class MetabolicDefaults:
    """Domyślne wartości dla profilu metabolicznego (INSCYD)."""
    
    VO2MAX: float = 55.0        # ml/kg/min - maksymalne zużycie tlenu
    VLAMAX: float = 0.50        # mmol/L/s - max produkcja mleczanu
    FATMAX_PERCENT: float = 0.55  # % FTP gdzie max spalanie tłuszczu
    

METABOLIC_DEFAULTS = MetabolicDefaults()


# =============================================================================
# STREFY MOCY (COGGAN)
# =============================================================================

@dataclass(frozen=True)
class PowerZone:
    """Definicja strefy mocy."""
    number: int
    name: str
    name_en: str
    min_pct: float
    max_pct: float
    description: str


POWER_ZONES: List[PowerZone] = [
    PowerZone(1, "Aktywna regeneracja", "Active Recovery", 0.00, 0.55,
              "Bardzo łatwa jazda, regeneracja po ciężkim treningu"),
    PowerZone(2, "Wytrzymałość", "Endurance", 0.55, 0.75,
              "Długie jazdy aerobowe, budowanie bazy tlenowej"),
    PowerZone(3, "Tempo", "Tempo", 0.75, 0.90,
              "Umiarkowanie ciężka praca, trening tempa"),
    PowerZone(4, "Próg (Sweet Spot)", "Threshold", 0.90, 1.05,
              "Praca na progu FTP, bardzo efektywny trening"),
    PowerZone(5, "VO2max", "VO2max", 1.05, 1.20,
              "Krótkie interwały 3-8 min, rozwój VO2max"),
    PowerZone(6, "Anaerobowa", "Anaerobic", 1.20, 1.50,
              "Bardzo krótkie wysiłki 30s-2min, praca glikolityczna"),
    PowerZone(7, "Nerwowo-mięśniowa", "Neuromuscular", 1.50, 3.00,
              "Sprinty <30s, maksymalna moc rekrutacji mięśni"),
]


# =============================================================================
# STREFY TĘTNA
# =============================================================================

@dataclass(frozen=True)
class HRZone:
    """Definicja strefy tętna."""
    number: int
    name: str
    min_pct: float  # % HRmax
    max_pct: float


HR_ZONES: List[HRZone] = [
    HRZone(1, "Z1 - Regeneracja", 0.50, 0.60),
    HRZone(2, "Z2 - Wytrzymałość", 0.60, 0.70),
    HRZone(3, "Z3 - Tempo", 0.70, 0.80),
    HRZone(4, "Z4 - Próg", 0.80, 0.90),
    HRZone(5, "Z5 - VO2max", 0.90, 1.00),
]


# =============================================================================
# PARAMETRY ARKUSZY
# =============================================================================

@dataclass(frozen=True)
class SheetConfig:
    """Konfiguracja parametrów arkuszy."""
    
    MAX_LOG_ROWS: int = 1000
    INITIAL_DAYS_COUNT: int = 90
    OUTPUT_FILENAME: str = "dziennik_kolarza_v3.xlsx"
    
    # Parametry PMC (Performance Management Chart)
    CTL_DAYS: int = 42   # Chronic Training Load - 42 dni
    ATL_DAYS: int = 7    # Acute Training Load - 7 dni


SHEET_CONFIG = SheetConfig()


# =============================================================================
# KOLORY (HEX)
# =============================================================================

@dataclass(frozen=True)
class Colors:
    """Paleta kolorów używana w arkuszach."""
    
    # Nagłówki
    HEADER_BG: str = "2E5090"      # Ciemnoniebieski
    HEADER_TEXT: str = "FFFFFF"    # Biały
    
    # Komórki do wpisywania
    INPUT_BG: str = "FFFFCC"       # Jasny żółty
    
    # Komórki z formułami
    FORMULA_BG: str = "E8E8E8"     # Jasny szary
    
    # Tekst informacyjny
    INFO_TEXT: str = "666666"      # Szary
    
    # Strefy mocy
    ZONE_1: str = "B4C6E7"  # Jasnoniebieski - regeneracja
    ZONE_2: str = "92D050"  # Zielony - wytrzymałość
    ZONE_3: str = "FFEB9C"  # Żółty - tempo
    ZONE_4: str = "FFC000"  # Pomarańczowy - próg
    ZONE_5: str = "FF6600"  # Ciemny pomarańcz - VO2max
    ZONE_6: str = "FF0000"  # Czerwony - anaerobowa
    ZONE_7: str = "7030A0"  # Fioletowy - nerwowo-mięśniowa
    
    # PMC
    CTL_COLOR: str = "0070C0"  # Niebieski - fitness
    ATL_COLOR: str = "FF6699"  # Różowy - zmęczenie
    TSB_COLOR: str = "00B050"  # Zielony - forma


COLORS = Colors()


# =============================================================================
# NAGŁÓWKI DZIENNIKA KOLARSKIEGO
# =============================================================================

LOG_HEADERS: List[str] = [
    # === SEKCJA 1: OGÓLNE ===
    "Data", "Tydzień", "Dzień tyg.",
    
    # === SEKCJA 2: FIZJOLOGIA (RANO) ===
    "Waga (kg)", "Waga śr. 7d", "RHR", "HRV (ms)", 
    "Sen (h)", "Jakość snu (1-5)", "Samopoczucie (1-5)",
    
    # === SEKCJA 3: DANE TRENINGU (Z GARMIN/ZWIFT) ===
    "Czas jazdy (min)", "Dystans (km)", "Przewyższenia (m)",
    "Avg Power (W)", "NP (W)", "Max Power (W)",
    "Avg Kadencja", "Avg HR", "Max HR",
    
    # === SEKCJA 4: METRYKI WKO5 (FORMUŁY) ===
    "IF", "TSS", "W/kg (NP)", "Strefa dom.",
    
    # === SEKCJA 5: PMC (FORMUŁY) ===
    "CTL", "ATL", "TSB",
    
    # === SEKCJA 6: KALORIE ===
    "Kcal treningu", "TDEE", "CEL Kcal", "Spożyte Kcal", "Bilans Kcal",
    
    # === SEKCJA 7: MAKROSKŁADNIKI ===
    "CEL B (g)", "CEL T (g)", "CEL W (g)",
    "Spoż. B (g)", "Spoż. T (g)", "Spoż. W (g)",
    
    # === SEKCJA 8: CHO PODCZAS TRENINGU ===
    "CHO/h (g)", "Nawodnienie (L)",
    
    # === SEKCJA 9: NOTATKI ===
    "Typ treningu", "RPE (1-10)", "Notatki"
]

# Kolumny do ręcznego wpisania (1-based index) - żółte tło
LOG_INPUT_COLUMNS: List[int] = [
    1,       # Data
    4,       # Waga (kg)
    6, 7,    # RHR, HRV
    8, 9, 10,# Sen, Jakość snu, Samopoczucie
    11, 12, 13,  # Czas, Dystans, Przewyższenia
    14, 15, 16,  # Avg/NP/Max Power
    17, 18, 19,  # Kadencja, HR
    30,      # Spożyte Kcal (kolumna 30, nie 31)
    35, 36, 37,  # Spożyte makro
    38, 39,  # CHO/h, Nawodnienie
    40, 41, 42   # Typ, RPE, Notatki
]

# Kolumny kończące sekcje logiczne (gruba prawa krawędź)
LOG_SECTION_END_COLUMNS: List[int] = [
    3,   # Po Dzień tyg.
    10,  # Po Samopoczucie
    19,  # Po Max HR
    23,  # Po Strefa dominująca
    26,  # Po TSB
    31,  # Po Bilans Kcal
    37,  # Po Spożyte Węgle
    39,  # Po Nawodnienie
]

# Szerokości kolumn dziennika
LOG_COLUMN_WIDTHS: List[int] = [
    12, 6, 6,    # Data, Tydzień, Dzień
    8, 8, 6, 8,  # Waga, Waga śr, RHR, HRV
    6, 10, 12,   # Sen, Jakość snu, Samopoczucie
    10, 10, 12,  # Czas, Dystans, Przewyższenia
    10, 10, 10,  # Avg/NP/Max Power
    10, 8, 8,    # Kadencja, HR
    6, 8, 8, 10, # IF, TSS, W/kg, Strefa
    8, 8, 8,     # CTL, ATL, TSB
    10, 10, 10, 10, 10,  # Kalorie
    8, 8, 8,     # Cele makro
    8, 8, 8,     # Spożyte makro
    8, 10,       # CHO/h, Nawodnienie
    15, 8, 30    # Typ, RPE, Notatki
]


# =============================================================================
# NAGŁÓWKI ŹRÓDEŁ CHO
# =============================================================================

CHO_HEADERS: List[str] = [
    "Nazwa produktu", "Porcja (g)", "CHO / 100g (g)", "kcal / 100g",
    "CHO w porcji (g)", "kcal w porcji", "Typ", "Szybkość wchłaniania", "Uwagi"
]

CHO_COLUMN_WIDTHS: List[int] = [30, 12, 12, 12, 14, 12, 15, 18, 35]

# Przykładowe dane produktów CHO - rozszerzone dla kolarstwa
CHO_SAMPLE_DATA: List[Tuple] = [
    ("Żel SiS GO", 60, 36.7, 138, "żel", "szybka", "Popularny żel bez kofeiny"),
    ("Żel Maurten 100", 40, 62.5, 250, "żel", "szybka", "Żel hydrożelowy, łagodny dla żołądka"),
    ("Żel z kofeiną", 40, 55, 200, "żel", "szybka", "Na ostatnie godziny wyścigu"),
    ("Baton Clif", 68, 66, 400, "baton", "średnia", "Dobre na długie jazdy w Z2"),
    ("Daktyle Medjool (3 szt)", 72, 75, 277, "naturalny", "średnia", "Naturalne źródło CHO"),
    ("Banan", 120, 23, 89, "naturalny", "średnia", "Klasyka, dobry na przerwę"),
    ("Napój Maurten 320", 500, 16, 64, "napój", "szybka", "80g CHO na bidon 500ml"),
    ("Napój SiS GO", 500, 7.2, 29, "napój", "szybka", "36g CHO na bidon"),
    ("Maltodekstryna", 30, 100, 400, "proszek", "bardzo szybka", "Do własnych miksów"),
    ("Fruktoza", 30, 100, 399, "proszek", "średnia", "Mieszać z MD 1:0.8"),
    ("Mix MD:Fruktoza 1:0.8", 54, 100, 399, "mieszanka", "szybka", "Optymalny stosunek 90g/h"),
    ("Rodzynki", 40, 79, 299, "naturalny", "średnia", "Wygodne w kieszeni"),
    ("Żelki Haribo", 50, 77, 343, "słodycze", "szybka", "Szybkie cukry na sprint"),
    ("Ryż biały (ugotowany)", 150, 28, 130, "posiłek", "średnia", "Carb-loading dzień przed"),
    ("Makaron (ugotowany)", 200, 25, 131, "posiłek", "średnia", "Bazowy posiłek kolarski"),
]


# =============================================================================
# TYPY TRENINGÓW
# =============================================================================

TRAINING_TYPES: List[str] = [
    "Z2 Wytrzymałość",
    "Z3 Tempo",
    "Sweet Spot",
    "FTP Interwały",
    "VO2max Interwały",
    "Sprinty",
    "Wyścig/Zawody",
    "Regeneracja",
    "Test FTP",
    "Grupowa jazda",
    "Commute",
    "Inne"
]


# =============================================================================
# WALIDACJA KONFIGURACJI
# =============================================================================

def _validate_config() -> None:
    """
    Waliduje spójność konfiguracji przy imporcie modułu.
    
    Raises:
        AssertionError: Gdy konfiguracja jest niespójna
    """
    # Sprawdź czy liczba nagłówków = liczba szerokości kolumn
    assert len(LOG_HEADERS) == len(LOG_COLUMN_WIDTHS), (
        f"Niezgodność: LOG_HEADERS ({len(LOG_HEADERS)}) != "
        f"LOG_COLUMN_WIDTHS ({len(LOG_COLUMN_WIDTHS)})"
    )
    
    # Sprawdź czy wszystkie indeksy input columns są w zakresie
    max_col = len(LOG_HEADERS)
    invalid_input_cols = [c for c in LOG_INPUT_COLUMNS if c < 1 or c > max_col]
    assert not invalid_input_cols, (
        f"LOG_INPUT_COLUMNS zawiera nieprawidłowe indeksy: {invalid_input_cols}. "
        f"Zakres: 1-{max_col}"
    )
    
    # Sprawdź czy wszystkie indeksy section end są w zakresie
    invalid_section_cols = [c for c in LOG_SECTION_END_COLUMNS if c < 1 or c > max_col]
    assert not invalid_section_cols, (
        f"LOG_SECTION_END_COLUMNS zawiera nieprawidłowe indeksy: {invalid_section_cols}. "
        f"Zakres: 1-{max_col}"
    )
    
    # Sprawdź liczba nagłówków CHO = liczba szerokości
    assert len(CHO_HEADERS) == len(CHO_COLUMN_WIDTHS), (
        f"Niezgodność: CHO_HEADERS ({len(CHO_HEADERS)}) != "
        f"CHO_COLUMN_WIDTHS ({len(CHO_COLUMN_WIDTHS)})"
    )


# Uruchom walidację przy imporcie
_validate_config()

