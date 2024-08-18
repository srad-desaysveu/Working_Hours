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

        # Suchen nach dem Index der Zeile, die text enthält
        text = "Diesen Block in jedes neue Excel Rüberkopieren"
        # idx = daten[daten.eq(text).any()].index
        
        # if not idx.empty:
        #     # Daten abschneiden, alles nach der gefundenden Zeile ignorieren
        #     daten = daten.iloc[:idx[0]]

        idx = daten.apply(lambda row: row.astype(str).str.contains(text).any(), axis=1)
        if idx.any():
            daten = daten.loc[:idx.idxmax()-1]

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
    überstunden_pro_tag = effektive_arbeitszeit_pro_tag_cap - sollarbeitszeit

    # Erstelle einen DataFrame mit den berechneten Werten
    ergebnisse = pd.DataFrame({
        'Effektive_Arbeitszeit': effektive_arbeitszeit_pro_tag,
        'Effektive_Arbeitszeit_10hCap': effektive_arbeitszeit_pro_tag_cap,
        'Überstunden': überstunden_pro_tag
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

def berechne_gesamtueberstunden(gesamte_arbeitszeit, gesamte_sollarbeitszeit):
    """
    Diese Funktion berechnet die Gesamtüberstunden eines Monats basierend auf der gesamten 
    effektiven Arbeitszeit und der Sollarbeitszeit.
    
    gesamte_arbeitszeit: Die gesamte effektive Arbeitszeit im Monat in Stunden.
    gesamte_sollarbeitszeit: Die gesamte Sollarbeitszeit im Monat in Stunden.
    
    Rückgabe: Die Gesamtüberstunden im Monat in Stunden.
    """
    gesamtueberstunden = gesamte_arbeitszeit - gesamte_sollarbeitszeit
    
    # Gesamtüberstunden dürfen nicht negativ sein
    if gesamtueberstunden < 0:
        gesamtueberstunden = 0
    
    return gesamtueberstunden

def korrigiere_sollarbeitszeit_fuer_urlaub(gesamte_sollarbeitszeit, urlaubszeiträume, dateipfad):
    """
    Diese Funktion korrigiert die monatliche Sollarbeitszeit basierend auf den angegebenen Urlaubstagen.
    
    gesamte_sollarbeitszeit: Die ursprüngliche monatliche Sollarbeitszeit in Stunden.
    urlaubszeiträume: Eine Liste von Tuple mit Urlaubstagen (z.B. [('2023-09-10', '2023-09-15'), ('2023-09-20', '2023-09-22')])
    dateipfad: Der Pfad zur Excel-Datei, um den Monat und das Jahr zu extrahieren.
    
    Rückgabe: Die angepasste Sollarbeitszeit unter Berücksichtigung der Urlaubstage.
    """
    # Extrahiere das Jahr und den Monat aus dem Dateinamen
    dateiname = os.path.splitext(os.path.basename(dateipfad))[0]
    year, month = map(int, dateiname.split('-')[0].split('_'))

    # Konvertiere die Urlaubsdaten in datetime-Objekte und filtere die für den aktuellen Monat
    urlaubstage = []
    for start, ende in urlaubszeiträume:
        start_date = datetime.datetime.strptime(start, '%Y-%m-%d').date()
        end_date = datetime.datetime.strptime(ende, '%Y-%m-%d').date()
        
        # Nur Urlaubstage im aktuellen Monat berücksichtigen
        if start_date.year == 2000 + year and start_date.month == month:
            urlaubstage.extend(pd.date_range(start=start_date, end=end_date).tolist())
        elif end_date.year == 2000 + year and end_date.month == month:
            urlaubstage.extend(pd.date_range(start=start_date, end=end_date).tolist())
    
    # Generiere eine Liste aller Arbeitstage im Monat
    start_date = datetime.date(2000 + year, month, 1)
    if month == 12:
        end_date = datetime.date(2000 + year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        end_date = datetime.date(2000 + year, month + 1, 1) - datetime.timedelta(days=1)
    
    all_days = pd.date_range(start=start_date, end=end_date, freq='D')
    werktage = [day for day in all_days if day.weekday() < 5]  # Montag bis Freitag

    # Zähle die Anzahl der Urlaubstage, die auf Werktage fallen
    urlaubstage_werktage = [day for day in urlaubstage if day in werktage]

    # Subtrahiere die Anzahl der Urlaubstage von den Sollarbeitstagen
    korrigierte_sollarbeitszeit = gesamte_sollarbeitszeit - len(urlaubstage_werktage) * 8  # 8 Stunden pro Urlaubstag

    return max(korrigierte_sollarbeitszeit, 0)

def summiere_ueberstunden_ueber_monate(dateipfade):
    """
    Diese Funktion liest die Überstunden aus mehreren Excel-Dateien mit "_Monatswerte" im Namen
    ein und summiert die Überstunden über alle Monate hinweg.
    
    dateipfade: Eine Liste von Pfaden zu den Excel-Dateien, die die monatlichen Überstunden enthalten.
    
    Rückgabe: Die gesamte Anzahl der Überstunden über alle angegebenen Monate hinweg.
    """

    def list_monatswerte_files(directory_path):
        # List all files in the given directory
        all_files = os.listdir(directory_path)
        
        # Filter files that end with "_Monatswerte.xlsx"
        monatswerte_files = [file for file in all_files if file.endswith('_Monatswerte.xlsx')]
        
        # Create the full path for each file
        monatswerte_files = [os.path.join(directory_path, file) for file in monatswerte_files]

        return monatswerte_files
    
    gesamtueberstunden = 0

    for pfad in dateipfade:
        filenames = list_monatswerte_files(pfad)
        for file in filenames:
            try:
                # Lese die Excel-Datei ein
                daten = pd.read_excel(file, sheet_name="Sheet1")
                # # Suche nach dem Blatt mit dem Namen, der "_Monatswerte" enthält
                # blattname = next(name for name in daten.keys() if "_Monatswerte" in name)
                # monatliche_werte = daten[blattname]
                
                # Extrahiere die Überstunden aus der entsprechenden Spalte (hier angenommen als "Gesamte Überstunden (Stunden)")
                ueberstunden = daten["Gesamte Überstunden (Stunden)"].sum()
                gesamtueberstunden += ueberstunden
            except Exception as e:
                print(f"Fehler beim Einlesen oder Verarbeiten der Datei {pfad}: {e}")

    return gesamtueberstunden

def korrigiere_gesamtueberstunden_manuell(bisherige_ueberstunden, korrekturwert):
    """
    Diese Funktion ermöglicht es, die Gesamtüberstunden manuell zu korrigieren, falls die berechneten Werte
    nicht korrekt erscheinen.
    
    bisherige_ueberstunden: Die bisherige Summe der Überstunden in Stunden.
    korrekturwert: Der Korrekturwert in Stunden, der zu den bisherigen Überstunden addiert oder subtrahiert wird.
    
    Rückgabe: Die korrigierte Anzahl der Überstunden.
    """
    korrigierte_ueberstunden = bisherige_ueberstunden + korrekturwert
    
    # Überstunden dürfen nicht negativ sein
    if korrigierte_ueberstunden < 0:
        korrigierte_ueberstunden = 0
    
    return korrigierte_ueberstunden
