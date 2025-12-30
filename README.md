# Kombajn Triathlonisty 🏊‍♂️🚴‍♂️🏃‍♂️

Generator dziennika treningowego w formacie Excel dla triathlonistów i sportowców wytrzymałościowych.

## Funkcjonalności

| Zakładka | Opis |
|----------|------|
| **Ustawienia i Cele** | Konfiguracja metabolizmu (BMR, TEF, NEAT) i celów kalorycznych |
| **Dziennik** | Codzienny log: waga, sen, trening, kalorie, makroskładniki |
| **Dashboard** | Podsumowania tygodniowe z automatycznymi obliczeniami |
| **Źródła CHO** | Baza produktów węglowodanowych z kalkulatorem porcji |

## Instalacja

### Wymagania

- Python 3.9+
- openpyxl 3.1+

### Kroki instalacji

```bash
# Klonowanie repozytorium
git clone <repo-url>
cd Dziennik_Zawodnika_V1

# Instalacja zależności
pip install -r requirements.txt
```

## Użycie

### Podstawowe

```bash
python -m kombajn.main
```

Wygeneruje plik `kombajn_triathlonisty_v2.xlsx` w bieżącym katalogu.

### Opcje linii poleceń

```bash
# Własna nazwa pliku
python -m kombajn.main -o moj_dziennik.xlsx

# Własny katalog wyjściowy
python -m kombajn.main -d C:\Dokumenty\Treningi

# Tryb szczegółowy (debug)
python -m kombajn.main -v
```

### Pomoc

```bash
python -m kombajn.main --help
```

## Jak korzystać z wygenerowanego pliku

1. **Otwórz plik** w programie Excel lub LibreOffice Calc
2. **Przejdź do [Ustawienia i Cele]** i wprowadź swoje dane:
   - BMR, TEF, NEAT
   - Planowany deficyt
   - Cele makroskładników
3. **Codziennie wypełniaj [Dziennik]**:
   - Żółte komórki → wypełniasz ręcznie
   - Szare komórki → obliczają się automatycznie
4. **Sprawdzaj [Dashboard]** dla podsumowań tygodniowych

## Struktura projektu

```
Dziennik_Zawodnika_V1/
├── kombajn/
│   ├── __init__.py          # Eksporty pakietu
│   ├── main.py              # Punkt wejścia CLI
│   ├── config.py            # Stałe i konfiguracja
│   ├── styles.py            # Style Excel
│   ├── utils.py             # Funkcje pomocnicze
│   └── sheets/
│       ├── __init__.py
│       ├── base.py          # Klasa bazowa arkuszy
│       ├── settings.py      # Arkusz Ustawienia
│       ├── log.py           # Arkusz Dziennik
│       ├── dashboard.py     # Arkusz Dashboard
│       └── cho_sources.py   # Arkusz Źródła CHO
├── tests/
│   └── test_kombajn.py      # Testy jednostkowe
├── requirements.txt
└── README.md
```

## Rozwój

### Uruchamianie testów

```bash
python -m pytest tests/ -v
```

### Pokrycie kodu

```bash
python -m pytest tests/ --cov=kombajn --cov-report=html
```

## Licencja

MIT License

## Autor

Athlete Tools
