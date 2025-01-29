#Artem
import os
from matplotlib.figure import Figure
from typing import Optional

from src.data_export import insert_content, create_report_template
from src.data_preperation import prepare_data, prepare_plots
from src.dataloader import load_data


def prepare_report(data_path: Optional[str]):
    create_report_template()
    unproccessed_data = load_data(data_path)
    proccessed_data = prepare_data(unproccessed_data)
    figures: list[Figure] = prepare_plots(proccessed_data)

    def get_sign_dependency(value: float) -> str:
        if value > 0:
            return "gestiegen"
        elif value < 0:
            return "gesunken"
        else:
            return "unverändert"

    # Process each row except the first one (which is used as reference)
    for idx in range(1, len(proccessed_data)):
        # Extract dates
        first_week = proccessed_data["Datum"].iloc[idx-1].split("-")[0].strip()
        first_week_ending = proccessed_data["Datum"].iloc[idx-1].split("-")[1].strip()
        current_week = proccessed_data["Datum"].iloc[idx].strip()

        # Extract data for first_analysis
        verkaufen = proccessed_data["Anzahl der Verkäufe"].iloc[idx]
        gesamtverkaufe = proccessed_data["Gesamtverkäufe (€)"].iloc[idx]
        kosten = proccessed_data["Kosten (€)"].iloc[idx]
        ertag = proccessed_data["Ertrag (€)"].iloc[idx]    
        ruckgaben = proccessed_data["Rückgaben (€)"].iloc[idx]
        beschadigte = proccessed_data["Beschädigte Ware (€)"].iloc[idx]
        gewinn = proccessed_data["Gewinn (€)"].iloc[idx]

        first_analysis = f"In dieser Woche konnten wir mit {verkaufen:,} Verkäufen einen Gesamtgewinn von {gesamtverkaufe:,} Euro erzielen.\
Nach Abzug der Kosten in Höhe von {kosten:,} Euro ergibt sich ein Nettoertrag von {ertag:,} Euro. Berücksichtigt\
man Rückgaben im Wert von {ruckgaben:,} Euro sowie beschädigte Ware im Wert von {beschadigte:,} Euro, beläuft sich\
der endgültige Gewinn auf {gewinn:,} Euro."

        # Extract data for second_analysis
        Gesamtverkäufe_Prozent_Änderung = proccessed_data["Gesamtverkäufe_Prozent_Änderung"].iloc[idx]
        zeichen_GPÄ = get_sign_dependency(Gesamtverkäufe_Prozent_Änderung)

        Kosten_Prozent_Änderung = proccessed_data["Kosten_Prozent_Änderung"].iloc[idx]
        zeichen_KPÄ = get_sign_dependency(Kosten_Prozent_Änderung)

        Anzahl_Verkäufe_Prozent_Änderung = proccessed_data["Anzahl_Verkäufe_Prozent_Änderung"].iloc[idx]
        zeichen_AVPÄ = get_sign_dependency(Anzahl_Verkäufe_Prozent_Änderung)
        
        Rückgaben_Prozent_Änderung = proccessed_data["Rückgaben_Prozent_Änderung"].iloc[idx]
        zeichen_RPÄ = get_sign_dependency(Rückgaben_Prozent_Änderung)

        Beschädigte_Ware_Prozent_Änderung = proccessed_data["Beschädigte_Ware_Prozent_Änderung"].iloc[idx]
        zeichen_BWPÄ = get_sign_dependency(Beschädigte_Ware_Prozent_Änderung)

        second_analysis = f"Im Vergleich zur Woche vom {first_week}-{first_week_ending} ist der Umsatz um {Gesamtverkäufe_Prozent_Änderung}% {zeichen_GPÄ}.\
Die Kosten sind um ca. {Kosten_Prozent_Änderung}% {zeichen_KPÄ}, jedoch im Verhältnis zum Umsatz moderater.\
Die Anzahl der Verkäufe hat sich um {Anzahl_Verkäufe_Prozent_Änderung}% {zeichen_AVPÄ}, während die Rückgaben um ca. {Rückgaben_Prozent_Änderung}% {zeichen_RPÄ} sind.\
Gleichzeitig ist der Anteil der beschädigten Ware um {Beschädigte_Ware_Prozent_Änderung}% {zeichen_BWPÄ}."

        content_dict = {
            "week": f"{current_week}",
            "Wochenbericht": f"{first_analysis}",
            "Vergleich zur vorherigen Woche": f"{second_analysis}",
        }

        # Insert content for each week
        insert_content(content_dict=content_dict, image=figures[idx-1], idx=idx)