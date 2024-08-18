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
filename = r"C:\Users\sradzijewski\OneDrive - DesaySV Europe\Dokumente\01_personal_data\01_Zeitablage\2023\23_09.xlsx"
filename = r"C:\Users\sradzijewski\OneDrive - DesaySV Europe\Dokumente\01_personal_data\01_Zeitablage\2023\23_10.xlsx"
filename = r"C:\Users\sradzijewski\OneDrive - DesaySV Europe\Dokumente\01_personal_data\01_Zeitablage\2023\23_11.xlsx"
filename = r"C:\Users\sradzijewski\OneDrive - DesaySV Europe\Dokumente\01_personal_data\01_Zeitablage\2023\23_12.xlsx"

Urlaub_2023 = [('2023-11-06', '2023-11-10'), ('2023-12-27', '2023-12-29'), ('2023-12-21', '2023-12-22'),]
Urlaub_2024 = [('2024-03-25', '2023-03-29'), ('2024-05-27', '2024-05-31'), ('2024-08-01', '2024-08-11'), ('2024-08-29', '2024-09-08'),]
Krank_2024 = [('2024-01-10', '2024-01-10'),('2024-02-08', '2024-02-09'),('2024-02-12', '2024-02-14'),]

########################################################################################
#%%
########################################################################################

def main(dateipfad, Urlaub):
    # code body here
    daten = lese_excel_datei(dateipfad)
    if daten is not None:
        dateiname = os.path.splitext(os.path.basename(dateipfad))[0]
        year, month = map(int, dateiname.split('-')[0].split('_'))

        tmp_path = os.path.join(os.getcwd(),os.path.basename(dateipfad))
        
        monats_ergebnisse = berechne_arbeitszeiten(daten)
        erstelle_getrennte_arbeitszeit_und_ueberstunden_diagramme(monats_ergebnisse, tmp_path)

        gesamte_arbeitszeit, gesamte_überstunden = berechne_monatliche_arbeitszeit_und_überstunden(monats_ergebnisse)

        gesamte_sollarbeitszeit = berechne_sollarbeitszeit(dateipfad, sollarbeitszeit_pro_tag=8)

        gesamte_sollarbeitszeit = korrigiere_sollarbeitszeit_fuer_urlaub(gesamte_sollarbeitszeit, Urlaub, dateipfad)

        print(f"Gesamte Arbeitszeit {2000+year} im Monat {month}: {gesamte_arbeitszeit} Stunden")
        print(f"Gesamte Überstudnen {2000+year} im Monat {month}: {gesamte_überstunden} Stunden")
        print(f"Gesamte Sollarbeitszeit {2000+year} im Monat {month}: {gesamte_sollarbeitszeit} Stunden")

        speichere_monatliche_werte_in_excel(tmp_path, gesamte_arbeitszeit, gesamte_überstunden)

        # os.path.dirname(tmp_path)
        Gesamtüberstunden_Monatsübergreifend =  summiere_ueberstunden_ueber_monate([os.path.dirname(tmp_path)])

        print(f"Gesamte Überstunden {2000+year} : {Gesamtüberstunden_Monatsübergreifend} Stunden")
    return

########################################################################################
#%%
########################################################################################

if __name__ == '__main__':
    print("This only executes when %s is executed rather than imported" % __file__)
    main(filename, Urlaub_2023)