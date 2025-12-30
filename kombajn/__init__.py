"""
Dziennik Kolarza - Generator dziennika treningowego Excel.

Ten pakiet umożliwia generowanie pliku Excel z metrykami WKO5/INSCYD:
- Strefy mocy (7 stref Coggan)
- Metryki performance (TSS, IF, NP)
- Performance Management Chart (CTL, ATL, TSB)
- Profil metaboliczny (VO2max, VLaMax)
- Baza produktów węglowodanowych
"""

from kombajn.main import create_workbook, main

__version__ = "3.0.0"
__author__ = "Athlete Tools"
__all__ = ["create_workbook", "main", "__version__"]
