import pandas as pd
from scipy.signal import find_peaks
import os
import numpy as np
from openpyxl import load_workbook
import traceback


def find_neighboring_peak(data, peak_index, direction="left", peak_type="max", prominence_threshold=50):
    # Funktion bleibt unverändert
    if direction == "left":
        search_range = range(peak_index - 1, -1, -1)  # Von peak_index-1 bis 0
    else:
        search_range = range(peak_index + 1, len(data))  # Von peak_index+1 bis zum Ende

    for idx in search_range:
        # Überprüfen, ob idx innerhalb der Grenzen von data liegt
        if idx < 0 or idx >= len(data):
            continue  # überspringen, wenn idx außerhalb der Grenzen ist

        if peak_type == "max":
            # Bedingungen für einen Hochpunkt
            if idx > 0 and idx < len(
                    data) - 1:  # Sicherstellen, dass der Zugriff auf data[idx - 1] und data[idx + 1] gültig ist
                if data[idx] > data[idx - 1] and data[idx] > data[idx + 1]:
                    # Prominenz prüfen
                    prominence = data[idx] - min(data[idx - 1], data[idx + 1])

                    if prominence >= prominence_threshold:
                        return idx, data[idx]

                # Zusätzliche Überprüfung für konstante Hochpunkte
                if data[idx] >= data[idx - 1] and data[idx] >= data[idx + 1]:
                    surrounding_values = data[max(0, idx - 2):min(len(data), idx + 3)]
                    if np.all(surrounding_values >= surrounding_values.max() - prominence_threshold):
                        return idx, data[idx]

        elif peak_type == "min":
            # Bedingungen für einen Tiefpunkt
            if idx > 0 and idx < len(
                    data) - 1:  # Sicherstellen, dass der Zugriff auf data[idx - 1] und data[idx + 1] gültig ist
                if data[idx] < data[idx - 1] and data[idx] < data[idx + 1]:
                    # Prominenz prüfen
                    prominence = max(data[idx - 1], data[idx + 1]) - data[idx]
                    if prominence >= prominence_threshold:
                        return idx, data[idx]

                # Zusätzliche Überprüfung für konstante Tiefpunkte
                if data[idx] <= data[idx - 1] and data[idx] <= data[idx + 1]:
                    surrounding_values = data[max(0, idx - 2):min(len(data), idx + 3)]
                    if np.all(surrounding_values <= surrounding_values.min() + prominence_threshold):
                        return idx, data[idx]

    return None  # Wenn kein Peak gefunden wurde


# Gesamtes Ergebnis DataFrame
results = []

try:
    # 1. Pfad zum Verzeichnis
    folder_path = r'K:\Team\Böhmer_Michael\ISO_3'  # Pfad eingeben für Ordner mit Dateien
    file_list = os.listdir(folder_path)

    # Nur Excel-Dateien filtern
    excel_files = [f for f in file_list if f.endswith('.xlsx')]

    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)

        # 2. Excel-Datei einlesen (spezifische Arbeitsblätter)
        with pd.ExcelFile(file_path) as xls:
            df_wiederholungen = pd.read_excel(xls, sheet_name='Wiederholungen')
            name = df_wiederholungen.iloc[1, 0]  # A2 ist der Name
            patient_id = df_wiederholungen.iloc[1, 1]  # B2 ist die ID

            if pd.isna(name):
                name = "n.a."
            if pd.isna(patient_id):
                patient_id = "n.a."

            # Daten für das linke Bein
            df_isokin_links = pd.read_excel(xls, sheet_name='Isokin_Kon_Kon_60_60_Links')

            # Daten für das rechte Bein
            df_isokin_rechts = pd.read_excel(xls, sheet_name='Isokin_Kon_Kon_60_60_Rechts')

        # Drehmoment und Winkel
        drehmoment_links = df_isokin_links.iloc[:, 2]  # Spalte C
        drehmoment_rechts = df_isokin_rechts.iloc[:, 2]  # Spalte C
        winkel_links = df_isokin_links.iloc[:, 1]  # Spalte B
        winkel_rechts = df_isokin_rechts.iloc[:, 1]  # Spalte B

        # Peaks finden
        prominence_threshold = 50  # Beispielwert
        peaks_links, _ = find_peaks(drehmoment_links, prominence=prominence_threshold)
        peaks_rechts, _ = find_peaks(drehmoment_rechts, prominence=prominence_threshold)

        # Initialisierung der Variablen
        max_extension_links, max_flexion_links = "nachbearbeiten", "nachbearbeiten"
        max_extension_rechts, max_flexion_rechts = "nachbearbeiten", "nachbearbeiten"
        index_extension_links, index_flexion_links = "n.a.", "n.a."
        index_extension_rechts, index_flexion_rechts = "n.a.", "n.a."

        # Überprüfung der gefundenen Peaks
        if len(peaks_links) != 8 or len(peaks_rechts) != 8:
            print(
                f"In Datei {file_name}: Die Hochpunkte sind nicht wie erwartet. Manuelle Nachbearbeitung erforderlich!")
            # Speichern der Ergebnisse in einem Dictionary
            result = {
                'Dateiname': file_name,
                'Name': name,
                'ID': patient_id,
                'Max Extension links': "nachbearbeiten",
                'Max Extension rechts': "nachbearbeiten",
                'Max Flexion links': "nachbearbeiten",
                'Max Flexion rechts': "nachbearbeiten",
                'Seitenunterschied Extension absolut': "nachbearbeiten",
                'Seitenunterschied Extension relativ': "nachbearbeiten",
                'Seitenunterschied Flexion absolut': "nachbearbeiten",
                'Seitenunterschied Flexion relativ': "nachbearbeiten",
                'Verhaeltnis Flexion Extension Links': "nachbearbeiten",
                'Verhaeltnis Flexion Extension rechts': "nachbearbeiten",
                'Unterschied Extension Flexion links': "nachbearbeiten",
                'Unterschied Extension Flexion rechts': "nachbearbeiten",
                'Winkel maximales Drehmoment links Extension': "nachbearbeiten",  # Winkel der Extension links
                'Winkel maximales Drehmoment rechts Extension': "nachbearbeiten",  # Winkel der Extension rechts
                'Winkel maximales Drehmoment links Flexion': "nachbearbeiten",  # Winkel der Flexion links
                'Winkel maximales Drehmoment rechts Flexion': "nachbearbeiten",  # Winkel der Flexion rechts
                'ROM Extension links': "nachbearbeiten",
                'ROM Extension rechts': "nachbearbeiten",
                'ROM Flexion links': "nachbearbeiten",
                'ROM Flexion rechts': "nachbearbeiten"
            }

            # Ergebnisse zur Gesamtliste hinzufügen
            results.append(result)

            continue
        else:
            # Maximalwerte und Indizes berechnen
            drehmoment_values_links = drehmoment_links.iloc[peaks_links].round(2).to_list()
            drehmoment_values_rechts = drehmoment_rechts.iloc[peaks_rechts].round(2).to_list()

            # Extension
            extension_values_links = [drehmoment_values_links[i] for i in [0, 2, 4, 6]]
            extension_indices_links = [peaks_links[i] for i in [0, 2, 4, 6]]
            extension_values_rechts = [drehmoment_values_rechts[i] for i in [0, 2, 4, 6]]
            extension_indices_rechts = [peaks_rechts[i] for i in [0, 2, 4, 6]]

            # Flexion
            flexion_values_links = [drehmoment_values_links[i] for i in [1, 3, 5, 7]]
            flexion_indices_links = [peaks_links[i] for i in [1, 3, 5, 7]]
            flexion_values_rechts = [drehmoment_values_rechts[i] for i in [1, 3, 5, 7]]
            flexion_indices_rechts = [peaks_rechts[i] for i in [1, 3, 5, 7]]

            # Maximalwerte finden
            max_extension_links = max(extension_values_links)
            max_flexion_links = max(flexion_values_links)
            max_extension_rechts = max(extension_values_rechts)
            max_flexion_rechts = max(flexion_values_rechts)

            # Indizes der maximalen Werte speichern
            index_extension_links = extension_indices_links[extension_values_links.index(max_extension_links)]
            index_flexion_links = flexion_indices_links[flexion_values_links.index(max_flexion_links)]
            index_extension_rechts = extension_indices_rechts[extension_values_rechts.index(max_extension_rechts)]
            index_flexion_rechts = flexion_indices_rechts[flexion_values_rechts.index(max_flexion_rechts)]

            # Winkel der maximalen Drehmomentwerte
            max_extension_links_winkel = df_isokin_links.iloc[index_extension_links, 1]
            max_flexion_links_winkel = df_isokin_links.iloc[index_flexion_links, 1]
            max_extension_rechts_winkel = df_isokin_rechts.iloc[index_extension_rechts, 1]
            max_flexion_rechts_winkel = df_isokin_rechts.iloc[index_flexion_rechts, 1]

        # 7. Berechnung der Seitenunterschiede (absolut und relativ) für Extension und Flexion

        # Berechnung für Extension
        if isinstance(max_extension_links, (int, float)) and isinstance(max_extension_rechts, (int, float)):
            seitenunterschied_extension_absolut = abs(max_extension_links - max_extension_rechts)

            # Sicherstellen, dass der kleinere Wert immer im Zähler steht
            min_extension = min(max_extension_links, max_extension_rechts)
            max_extension = max(max_extension_links, max_extension_rechts)

            # Relativer Seitenunterschied in Prozent
            seitenunterschied_extension_relativ = round((1 - (min_extension / max_extension)) * 100, 2)
        else:
            seitenunterschied_extension_absolut = "nachbearbeiten"
            seitenunterschied_extension_relativ = "nachbearbeiten"

        # Berechnung für Flexion
        if isinstance(max_flexion_links, (int, float)) and isinstance(max_flexion_rechts, (int, float)):
            seitenunterschied_flexion_absolut = abs(max_flexion_links - max_flexion_rechts)

            # Sicherstellen, dass der kleinere Wert immer im Zähler steht
            min_flexion = min(max_flexion_links, max_flexion_rechts)
            max_flexion = max(max_flexion_links, max_flexion_rechts)

            # Relativer Seitenunterschied in Prozent
            seitenunterschied_flexion_relativ = round((1 - (min_flexion / max_flexion)) * 100, 2)
        else:
            seitenunterschied_flexion_absolut = "nachbearbeiten"
            seitenunterschied_flexion_relativ = "nachbearbeiten"

        # 8. Berechnung der Verhältnisse für Extension/Flexion für dasselbe Bein

        # Berechnung für das linke Bein
        if isinstance(max_flexion_links, (int, float)) and max_flexion_links != 0:
            verhaeltnis_flexion_extension_links = round(max_flexion_links / max_extension_links, 2)
        else:
            verhaeltnis_flexion_extension_links = "n.a."

        # Berechnung für das rechte Bein
        if isinstance(max_flexion_rechts, (int, float)) and max_flexion_rechts != 0:
            verhaeltnis_flexion_extension_rechts = round(max_flexion_rechts / max_extension_rechts, 2)
        else:
            verhaeltnis_flexion_extension_rechts = "n.a."

        # 9. Berechnung der Unterschiede für Extension/Flexion für dasselbe Bein

        # Berechnung für linkes Bein
        if isinstance(max_extension_links, (int, float)) and isinstance(max_flexion_links, (int, float)):
            unterschied_extension_flexion_links = abs(max_extension_links - max_flexion_links)
        else:
            unterschied_extension_flexion_links = "n.a."

        # Berechnung für rechtes Bein
        if isinstance(max_extension_rechts, (int, float)) and isinstance(max_flexion_rechts, (int, float)):
            unterschied_extension_flexion_rechts = abs(max_extension_rechts - max_flexion_rechts)
        else:
            unterschied_extension_flexion_rechts = "n.a."

        # 10. ROM (Range of Motion) berechnen

        # Links: ROM für Extension
        winkelhochpunkt_extension_links = find_neighboring_peak(winkel_links, index_extension_links, direction="left",
                                                                peak_type="max", prominence_threshold=50)
        winkeltiefpunkt_extension_links = find_neighboring_peak(winkel_links, index_extension_links, direction="right",
                                                                peak_type="min", prominence_threshold=50)
        if winkelhochpunkt_extension_links is None or winkeltiefpunkt_extension_links is None:
            rom_extension_links = "nachbearbeiten"
        else:
            rom_extension_links = f"{str(round(winkeltiefpunkt_extension_links[1], 2)).replace('.', ',')} - {str(round(winkelhochpunkt_extension_links[1], 2)).replace('.', ',')}"

        # Rechts: ROM für Extension
        winkelhochpunkt_extension_rechts = find_neighboring_peak(winkel_rechts, index_extension_rechts,
                                                                 direction="left", peak_type="max",
                                                                 prominence_threshold=50)
        winkeltiefpunkt_extension_rechts = find_neighboring_peak(winkel_rechts, index_extension_rechts,
                                                                 direction="right", peak_type="min",
                                                                 prominence_threshold=50)
        if winkelhochpunkt_extension_rechts is None or winkeltiefpunkt_extension_rechts is None:
            rom_extension_rechts = "nachbearbeiten"
        else:
            rom_extension_rechts = f"{str(round(winkeltiefpunkt_extension_rechts[1], 2)).replace('.', ',')} - {str(round(winkelhochpunkt_extension_rechts[1], 2)).replace('.', ',')}"

        # Links: ROM für Flexion
        winkelhochpunkt_flexion_links = find_neighboring_peak(winkel_links, index_flexion_links, direction="right",
                                                              peak_type="max", prominence_threshold=50)
        winkeltiefpunkt_flexion_links = find_neighboring_peak(winkel_links, index_flexion_links, direction="left",
                                                              peak_type="min", prominence_threshold=50)
        if winkelhochpunkt_flexion_links is None or winkeltiefpunkt_flexion_links is None:
            rom_flexion_links = "nachbearbeiten"
        else:
            rom_flexion_links = f"{str(round(winkeltiefpunkt_flexion_links[1], 2)).replace('.', ',')} - {str(round(winkelhochpunkt_flexion_links[1], 2)).replace('.', ',')}"

        # Rechts: ROM für Flexion
        winkelhochpunkt_flexion_rechts = find_neighboring_peak(winkel_rechts, index_flexion_rechts, direction="right",
                                                               peak_type="max", prominence_threshold=50)
        winkeltiefpunkt_flexion_rechts = find_neighboring_peak(winkel_rechts, index_flexion_rechts, direction="left",
                                                               peak_type="min", prominence_threshold=50)
        if winkelhochpunkt_flexion_rechts is None or winkeltiefpunkt_flexion_rechts is None:
            rom_flexion_rechts = "nachbearbeiten"
        else:
            rom_flexion_rechts = f"{str(round(winkeltiefpunkt_flexion_rechts[1], 2)).replace('.', ',')} - {str(round(winkelhochpunkt_flexion_rechts[1], 2)).replace('.', ',')}"

        # Speichern der Ergebnisse in einem Dictionary
        result = {
            'Dateiname': file_name,
            'Name': name,
            'ID': patient_id,
            'Max Extension links': max_extension_links,
            'Max Extension rechts': max_extension_rechts,
            'Max Flexion links': max_flexion_links,
            # 'IndexExtensionLinks': index_extension_links,
            # 'IndexFlexionLinks': index_flexion_links,
            'Max Flexion rechts': max_flexion_rechts,
            # 'IndexExtensionRechts': index_extension_rechts,
            # 'IndexFlexionRechts': index_flexion_rechts,
            'Seitenunterschied Extension absolut': seitenunterschied_extension_absolut,
            'Seitenunterschied Extension relativ': seitenunterschied_extension_relativ,
            'Seitenunterschied Flexion absolut': seitenunterschied_flexion_absolut,
            'Seitenunterschied Flexion relativ': seitenunterschied_flexion_relativ,
            'Verhaeltnis Flexion Extension Links': verhaeltnis_flexion_extension_links,
            'Verhaeltnis Flexion Extension rechts': verhaeltnis_flexion_extension_rechts,
            'Unterschied Extension Flexion links': unterschied_extension_flexion_links,
            'Unterschied Extension Flexion rechts': unterschied_extension_flexion_rechts,
            'Winkel maximales Drehmoment links Extension': max_extension_links_winkel,  # Winkel der Extension links
            'Winkel maximales Drehmoment rechts Extension': max_extension_rechts_winkel,  # Winkel der Extension rechts
            'Winkel maximales Drehmoment links Flexion': max_flexion_links_winkel,  # Winkel der Flexion links
            'Winkel maximales Drehmoment rechts Flexion': max_flexion_rechts_winkel,  # Winkel der Flexion rechts
            'ROM Extension links': rom_extension_links,
            'ROM Extension rechts': rom_extension_rechts,
            'ROM Flexion links': rom_flexion_links,
            'ROM Flexion rechts': rom_flexion_rechts
        }

        # Ergebnisse zur Gesamtliste hinzufügen
        results.append(result)

    # 3. Ergebnisse zu einem DataFrame zusammenfassen
    results_df = pd.DataFrame(results)

    # 4. Ergebnisse in eine neue Excel-Datei speichern
    output_file_path = os.path.join(folder_path, 'auswertung_gesamt.xlsx')
    results_df.to_excel(output_file_path, index=False)

    # Spaltenbreite auf 20 setzen
    workbook = load_workbook(output_file_path)
    try:
        sheet = workbook.active
        for column in sheet.columns:
            adjusted_width = 20  # Breite auf 20 setzen
            sheet.column_dimensions[column[0].column_letter].width = adjusted_width

        # Arbeitsmappe speichern
        workbook.save(output_file_path)
    finally:
        workbook.close()  # Sicherstellen, dass die Arbeitsmappe immer geschlossen wird

    print(f"Die Ergebnisse wurden erfolgreich in {output_file_path} gespeichert.")


except Exception as e:
    print(f"Ein Fehler ist aufgetreten beim Verarbeiten der Datei {file_name}: {e}")
    print("Detaillierte Fehlermeldung:")
    traceback.print_exc()  # Gibt die vollständige Fehlermeldung und den Stack-Trace aus
