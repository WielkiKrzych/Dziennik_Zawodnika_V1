"""
Kombajn Triathlonisty - Generator dziennika treningowego Excel.

Ten pakiet umożliwia generowanie pliku Excel z:
- Ustawieniami i celami kalorycznymi
- Dziennikiem codziennych wpisów
- Dashboardem z podsumowaniami tygodniowymi
- Bazą źródeł węglowodanów
"""

from kombajn.main import create_workbook, main

__version__ = "2.0.0"
__author__ = "Athlete Tools"
__all__ = ["create_workbook", "main"]
