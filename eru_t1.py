import datetime
from typing import List
import pandas as pd
import json
import os
from pprint import pprint
from dataclasses import dataclass, field
from typing import List, Optional

from utils import get_date, months_by_quarter


class JsonHandler:
    def __init__(self, ):
        self.quarter = None
        self.region = [
            {
                "name": "Hlavní město Praha",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Středočeský",
                "paliv": 3,
                "paliva": [
                    "Zemni plyn",
                    "Biomasa - Brikety a pelety",
                    "Biomasa - Piliny. kura. stepky. drevni odpad",
                ]
            },
            {
                "name": "Jihočeský",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Plzeňský",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Karlovarský",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Ústecký",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Liberecký",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Královéhradecký",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Pardubický",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Vysočina",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Jihomoravský",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Olomoucký",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Zlínský",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
            {
                "name": "Moravskoslezský",
                "paliv": 1,
                "paliva": [
                    "Zemni plyn",
                ]
            },
        ]
        self.statements = []
        self.data_dir = os.path.join(os.path.dirname(__file__), "data-t1")
        self.data_df = None

    def read_file(self, excel_file_path: str, sheet_name: str):
        df = pd.read_excel(
            excel_file_path, sheet_name=sheet_name, skiprows=0, header=1)
        # Convert all numeric columns to float
        numeric_columns = df.select_dtypes(include=['int64', 'float64']).columns
        df[numeric_columns] = df[numeric_columns].astype(float)

        # round all files to 2 decimal places
        df = df.round(2)

        return df.fillna(0)

    def make_statements(self, df: pd.DataFrame, date: datetime.date):
        # Iterate through rows and create statements
        for index, row in df.iterrows():
            statement = self.make_statement(row, date)
            self.statements.append(statement)

    # Find row based on month, region, and typPaliva
    def filter_data(self, data, month, region, typPaliva=None):

        if typPaliva is None:
            return data[(data["month"] == month) & (data["region"] == region)]

        return data[(data["month"] == month) & (data["region"] == region) & (data["typPaliva"] == typPaliva)]

    def make_gen_sell_heat_part(self, region, month, palivo):
        filtered_data = self.filter_data(
            self.data_df, month, region["name"], palivo)
        return {
            "typPaliva": palivo,
            "ztraty": float(filtered_data["h_ztraty"].sum()),
            "bruttoVyroba": float(filtered_data["h_bruttoVyroba"].sum()),
            "bilancniRozdil": float(filtered_data["h_bilancniRozdil"].sum()),
            "primeDodavkyCizimSubjektum": float(filtered_data["h_primeDodavkyCizimSubjektum"].sum()),
            "technologickaVlastniSpotreba": float(filtered_data["h_technologickaVlastniSpotreba"].sum()),
            "dodavkyDoVlastnihoPodnikuNeboZarizeni": float(filtered_data["h_dodavkyDoVlastnihoPodnikuNeboZarizeni"].sum()),
        }

    def make_paliva_part(self, region, month, palivo):
        filtered_data = self.filter_data(
            self.data_df, month, region["name"], palivo)

        return {
            "typPaliva": palivo,
            "jednotkyPaliv": "MWh" if palivo == "Zemni plyn" else "t",
            "porizeniPaliv": float(filtered_data["pal_porizeniPaliv"].sum()),
            "spotrebaPaliva": float(filtered_data["pal_spotrebaPaliva"].sum()),
            "vyhrevnostHodnota": float(filtered_data["pal_vyhrevnostHodnota"].sum()),
            "vyhrevnostJednotky": "MJ/m3" if palivo == "Zemni plyn" else "GJ/t"
        }

    def make_region_part(self, region, month) -> dict:
        print(region)
        filtered_data = self.filter_data(self.data_df, month, region["name"])

        json = {
            "dataZaKraj": region["name"],
            "vykazaneHodnoty": {
                "paliva": {
                    "pocetPaliv": str(region["paliv"]),
                    "paliva": [self.make_paliva_part(region, month, palivo) for palivo in region["paliva"]],
                    "vyrobaADodavkaTepla": [self.make_gen_sell_heat_part(region, month, palivo) for palivo in region["paliva"]],
                },
                "celkovyInstalovanyVykon": float(filtered_data["celkovyInstalovanyVykon"].sum()),
                "bilanceDodavekAZdroju": {
                    "nakup": float(filtered_data["bil_nakup"].sum()),
                    "saldo": float(filtered_data["bil_saldo"].sum()),
                    "ztraty": float(filtered_data["bil_ztraty"].sum()),
                    "doprava": float(filtered_data["bil_doprava"].sum()),
                    "ostatni": float(filtered_data["bil_ostatni"].sum()),
                    "prumysl": float(filtered_data["bil_prumysl"].sum()),
                    "domacnosti": float(filtered_data["bil_domacnosti"].sum()),
                    "energetika": float(filtered_data["bil_energetika"].sum()),
                    "bruttoVyroba": float(filtered_data["bil_bruttoVyroba"].sum()),
                    "stavebnictvi": float(filtered_data["bil_stavebnictvi"].sum()),
                    "bilancniRozdil": float(filtered_data["bil_bilancniRozdil"].sum()),
                    "vlastniSpotreba": float(filtered_data["bil_vlastniSpotreba"].sum()),
                    "zemedelstviALesnictvi": float(filtered_data["bil_zemedelstviALesnictvi"].sum()),
                    "dodavkyObchodnimSubjektum": float(filtered_data["bil_dodavkyObchodnimSubjektum"].sum()),
                    "obchodSluzbySkolstviZdravotnictvi": float(filtered_data["bil_obchodSluzbySkolstviZdravotnictvi"].sum())
                }
            }
        }

        return json

    def make_month_part(self, month) -> dict:
        json = {
            "dataZaMesic": month,
            "kraje": {
                "kraje": [self.make_region_part(region, month) for region in self.region]
            },
        }
        return json

    def make_statement(self) -> dict:
        # Actual date-time string
        datestr = get_date()
        month_list = months_by_quarter(self.quarter)

        json = {
            "identifikacniUdajeVykazu": {
                "ico": "29060109",
                "typVykazu": "T1",
                "typPeriody": "QUARTER",
                "cisloLicence": [
                    "311018326",
                    "321018327"
                ],
                "vykazovanyRok": 2024,
                "datovaSchranka": "jakub.urban@hotmail.com",
                "drzitelLicence": "ČEZ Energo, s.r.o.",
                "kontaktniTelefon": "+420721055966",
                "vykazovanaPerioda": self.quarter,
                "odpovednyPracovnik": "Jakub Urban",
                "datumVytvoreniVykazu": datestr
            },
            "t1Komentar": {
                "komentar": "Some comment for T1"
            },
            "t1": {
                "mesice": [
                    self.make_month_part(month) for month in month_list
                ]
            },
        }

        return json

    def generate_statements(self) -> List[dict]:

        # Return the list of statements as JSON
        json_base = {
            "typVykazu": "T1",
            "vykazy": [self.make_statement()]
        }

        return json_base


# Usage example:
handler = JsonHandler()

# Local inFile path
# infile = os.path.join(handler.data_dir, 'XML Generování-latest.xlsm')

# File path to the Excel file stored in oneDrive
infile = '/mnt/c/Users/JakubUrban/OneDrive - ČEZ Energo, s.r.o/_Pracovní 2024/02_ERU/ERÚ-T1.xlsx'
handler.quarter = 4
handler.data_df = handler.read_file(infile, 'Export')

today = datetime.date.today()
# handler.make_statements(date=today, df=data)
# handler.make_statement()

json_output = handler.generate_statements()


# Save the JSON to a file

outfile_path = os.path.join(handler.data_dir, 'dataT1.json')
with open(outfile_path, 'w', encoding="utf-8") as outfile:
    json.dump(json_output, outfile, indent=4, ensure_ascii=False)
