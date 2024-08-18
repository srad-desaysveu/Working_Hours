# necessary to run script from subdir:
import sys
import os
print (os.getcwd())
sys.path.append(os.getcwd())
import numpy as np

import pandas as pd
import numpy as np
import datetime
import holidays
import os
import matplotlib.pyplot as plt

from Working_hours import *

########################################################################################
#%%
########################################################################################
filename = r"D:\Desay\01_personal_data\01_Zeitablage\2023\23_09 - Kopie.xlsx"

########################################################################################
#%%
########################################################################################

def main(dateipfad):
    # code body here
    daten = lese_excel_datei(dateipfad)
    if daten is not None:
        dateiname = os.path.splitext(os.path.basename(dateipfad))[0]
        year, month = map(int, dateiname.split('-')[0].split('_'))

        ergebnisse = berechne_arbeitszeiten(daten)

        tmp_path = os.path.join(os.getcwd(),os.path.basename(dateipfad))
        erstelle_getrennte_arbeitszeit_und_ueberstunden_diagramme(ergebnisse, tmp_path)

        gesamte_arbeitszeit, gesamte_überstunden = berechne_monatliche_arbeitszeit_und_überstunden(ergebnisse)

        gesamte_sollarbeitszeit = berechne_sollarbeitszeit(dateipfad, sollarbeitszeit_pro_tag=8)

        print(f"Gesamte Arbeitszeit {2000+year} im Monat {month}: {gesamte_arbeitszeit} Stunden")
        print(f"Gesamte Überstudnen {2000+year} im Monat {month}: {gesamte_überstunden} Stunden")
        print(f"Gesamte Sollarbeitszeit {2000+year} im Monat {month}: {gesamte_sollarbeitszeit} Stunden")
        
        speichere_monatliche_werte_in_excel(tmp_path, gesamte_arbeitszeit, gesamte_überstunden)
    return

########################################################################################
#%%
########################################################################################

if __name__ == '__main__':
    print("This only executes when %s is executed rather than imported" % __file__)
    main(filename)