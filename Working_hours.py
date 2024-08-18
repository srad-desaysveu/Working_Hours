import pandas as pd
import numpy as np
import datetime
import holidays
import os
import matplotlib.pyplot as plt

def lese_excel_datei(dateipfad):
    """
    Diese Funktion liest die Excel-Datei ein und gibt die Daten als pandas DataFrame zurück.
    """
    try:
        daten = pd.read_excel(dateipfad)
        print("Datei erfolgreich eingelesen.")
        return daten
    except Exception as e:
        print(f"Fehler beim Einlesen der Datei: {e}")
        return None

def berechne_arbeitszeiten(daten, sollarbeitszeit=8):
    """
    Diese Funktion berechnet die gesamte effektive Arbeitszeit pro Tag, die auf 10 Stunden gekappte Arbeitszeit
    sowie die Überstunden.
    
    daten: DataFrame, das die eingelesenen Exceldaten enthält
    sollarbeitszeit: Sollarbeitszeit pro Tag in Stunden (Standard: 8 Stunden)
    
    Rückgabe: DataFrame mit den berechneten Werten für jeden Tag
    """
    
    # Konvertiere Start und Ende in datetime
    daten['Start'] = pd.to_datetime(daten['Start'])
    daten['Ende'] = pd.to_datetime(daten['Ende'])

    # Berechne die Dauer der Arbeitszeit (in Sekunden)
    daten['berechnete_dauer'] = (daten['Ende'] - daten['Start']).dt.total_seconds()

    # Konvertiere die Pause in Sekunden
    daten['Pause'] = daten['Pause'].astype(str)
    
    def convert_pause_to_seconds(pause):
        try:
            return pd.to_timedelta(pause).total_seconds()
        except:
            return 0

    daten['pause_in_sekunden'] = daten['Pause'].apply(convert_pause_to_seconds)

    # Berechne die effektive Arbeitszeit (Dauer - Pause) in Stunden
    daten['effektive_arbeitszeit_in_stunden'] = (daten['berechnete_dauer'] - daten['pause_in_sekunden']) / 3600

    # Gruppiere die Daten nach Tag und summiere die effektive Arbeitszeit pro Tag
    effektive_arbeitszeit_pro_tag = daten.groupby('Tag')['effektive_arbeitszeit_in_stunden'].sum()

    # Implementiere die 10-Stunden-Cap
    effektive_arbeitszeit_pro_tag_cap = effektive_arbeitszeit_pro_tag.copy()
    effektive_arbeitszeit_pro_tag_cap[effektive_arbeitszeit_pro_tag_cap > 10] = 10

    # Berechne die Überstunden basierend auf der gesamten effektiven Arbeitszeit
    überstunden = effektive_arbeitszeit_pro_tag_cap - sollarbeitszeit

    # Erstelle einen DataFrame mit den berechneten Werten
    ergebnisse = pd.DataFrame({
        'Effektive_Arbeitszeit': effektive_arbeitszeit_pro_tag,
        'Effektive_Arbeitszeit_10hCap': effektive_arbeitszeit_pro_tag_cap,
        'Überstunden': überstunden
    })

    return ergebnisse

def erstelle_getrennte_arbeitszeit_und_ueberstunden_diagramme(ergebnisse, dateipfad):
    """
    Diese Funktion erstellt zwei getrennte Grafiken: eine für die Effektive Arbeitszeit pro Tag und eine für die Überstunden pro Tag.
    Die Grafiken werden im gleichen Pfad wie die Excel-Datei gespeichert mit dem Suffix '_Arbeitszeitdiagramm.svg'.
    
    ergebnisse: DataFrame, das die berechneten Werte enthält
    dateipfad: Pfad zur zugrunde liegenden Excel-Datei
    """
    # Datei-Name und Pfad bestimmen
    dateiname = os.path.splitext(os.path.basename(dateipfad))[0]
    speicherpfad = os.path.join(os.path.dirname(dateipfad), f"{dateiname}_Arbeitszeitdiagramm.svg")

    # Grafik erstellen mit zwei Subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))

    # Effektive Arbeitszeit pro Tag im ersten Subplot
    ergebnisse['Effektive_Arbeitszeit'].plot(kind='bar', ax=ax1, color='skyblue', label='Effektive Arbeitszeit')
    ax1.axhline(y=8, color='red', linestyle='--', label='Sollarbeitszeit')
    ax1.axhline(y=10, color='green', linestyle='--', label='Maximale Arbeitszeit (10h)')
    ax1.set_title('Effektive Arbeitszeit pro Tag')
    ax1.set_ylabel('Arbeitszeit (Stunden)')
    ax1.set_xlabel('Tag')
    ax1.legend()
    ax1.grid(True)

    # Überstunden pro Tag im zweiten Subplot
    ergebnisse['Überstunden'].plot(kind='bar', ax=ax2, color='orange', label='Überstunden')
    ax2.axhline(y=0, color='red', linestyle='--', label='Nullpunkt')
    ax2.set_title('Überstunden pro Tag')
    ax2.set_ylabel('Überstunden (Stunden)')
    ax2.set_xlabel('Tag')
    ax2.legend()
    ax2.grid(True)

    # Layout anpassen und speichern
    plt.tight_layout()
    plt.savefig(speicherpfad)
    print(f"Diagramm gespeichert unter: {speicherpfad}")

def berechne_monatliche_arbeitszeit_und_überstunden(ergebnisse):
    """
    Diese Funktion berechnet die gesamte Arbeitszeit und die gesamten Überstunden eines Monats.
    
    ergebnisse: DataFrame, das die berechneten Werte enthält
    
    Rückgabe: Tuple mit der gesamten Arbeitszeit und den gesamten Überstunden
    """
    gesamte_arbeitszeit = ergebnisse['Effektive_Arbeitszeit'].sum()
    gesamte_überstunden = ergebnisse['Überstunden'].sum()
    
    return gesamte_arbeitszeit, gesamte_überstunden

def speichere_monatliche_werte_in_excel(dateipfad, gesamte_arbeitszeit, gesamte_überstunden):
    """
    Diese Funktion speichert die gesamten monatlichen Arbeitszeiten und Überstunden in eine neue Excel-Datei.
    
    dateipfad: Pfad zur ursprünglichen Excel-Datei, um den Speicherort und Dateinamen zu bestimmen
    gesamte_arbeitszeit: gesamte Arbeitszeit des Monats
    gesamte_überstunden: gesamte Überstunden des Monats
    """
    # Datei-Name und Pfad bestimmen
    dateiname = os.path.splitext(os.path.basename(dateipfad))[0]
    speicherpfad = os.path.join(os.path.dirname(dateipfad), f"{dateiname}_Monatswerte.xlsx")

    # Erstelle ein DataFrame für die Werte
    daten = pd.DataFrame({
        'Gesamte Arbeitszeit (Stunden)': [gesamte_arbeitszeit],
        'Gesamte Überstunden (Stunden)': [gesamte_überstunden]
    })

    # Speichere die Daten in eine neue Excel-Datei
    daten.to_excel(speicherpfad, index=False)
    print(f"Monatliche Werte gespeichert unter: {speicherpfad}")

def berechne_sollarbeitszeit(dateipfad, sollarbeitszeit_pro_tag=8):
    """
    Diese Funktion berechnet die Sollarbeitszeit eines Monats basierend auf der Anzahl der Werktage 
    und berücksichtigt dabei die Feiertage in Thüringen.
    
    dateipfad: Pfad zur Excel-Datei, um die Daten zu laden
    sollarbeitszeit_pro_tag: Anzahl der Sollarbeitsstunden pro Tag (Standard: 8 Stunden)
    
    Rückgabe: Die gesamte Sollarbeitszeit des Monats in Stunden.
    """
    # Extrahiere das Datum (YY_MM) aus dem Dateinamen
    dateiname = os.path.splitext(os.path.basename(dateipfad))[0]
    year, month = map(int, dateiname.split('-')[0].split('_'))

    # Erstelle ein Datumsbereich für den Monat
    start_date = datetime.date(2000 + year, month, 1)
    if month == 12:  # Handle year change for December
        end_date = datetime.date(2000 + year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        end_date = datetime.date(2000 + year, month + 1, 1) - datetime.timedelta(days=1)
    
    # Generiere eine Liste aller Tage im Monat
    all_days = pd.date_range(start=start_date, end=end_date, freq='D')

    # Identifiziere alle Feiertage in Thüringen für das angegebene Jahr
    feiertage = holidays.Germany(years=2000 + year, state='TH')

    # Zähle die Anzahl der Werktage (Montag bis Freitag) unter Berücksichtigung der Feiertage
    werktage = np.isin(all_days.weekday, [0, 1, 2, 3, 4]).sum()
    feiertage_im_monat = sum(1 for day in all_days if day in feiertage and day.weekday() < 5)

    # Berechne die Sollarbeitszeit unter Berücksichtigung der Feiertage
    gesamt_sollarbeitszeit = (werktage - feiertage_im_monat) * sollarbeitszeit_pro_tag

    return gesamt_sollarbeitszeit

# Beispielver
