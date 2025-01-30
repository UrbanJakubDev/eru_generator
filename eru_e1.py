import datetime
import pandas as pd
import json
import os


from typing import List
from utils import get_date


class JsonHandler:
    def __init__(self):
        self.statements = []
        self.data_dir = os.path.join(os.path.dirname(__file__), "data-e1")
        
    
  
    def read_file(self, excel_file_path: str, sheet_name: str):
        # Assuming your Excel file has a sheet named 'Sheet1'

        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, skiprows=0, header=1)
        # Assuming your Excel sheet has columns like 'typPaliva', 'jednotkaPaliva', etc.
        return df.fillna(0)

    def make_statements(self, df: pd.DataFrame, date: datetime.date):
        # Iterate through rows and create statements
        for index, row in df.iterrows():

            if row["nazevVyrobny"] == 0:
                continue
            statement = self.make_statement(row, date)
            self.statements.append(statement)

    def make_statement(self, row, date) -> dict:
        # Actual date-time string
        datestr = get_date()
        
        year = date.year
        month = date.month
        
        json = {
            "e1vykaz": {
                "e1Paliva": {
                    "paliva": [
                        {
                            "typPaliva": "Zemni plyn",
                            "jednotkaPaliva": "MWh",
                            "vyhrevnostHodnota": row["g-vyhrevnostHodnota"],
                            "vyhrevnostJednotka": "MJ/m3",
                            "porizeniPalivaCelkem": row["g-porizeniPalivaCelkem"],
                            "spotrebaPalivaVyrobaTepla": row["g-spotrebaPalivaVyrobaTepla"],
                            "spotrebaPalivaVyrobaElektriny": row["g-spotrebaPalivaVyrobaElektriny"],
                        }
                    ]
                },
                "pocetPaliv": "1",
                "e1Komentar": {
                    "komentar": ""
                },
                "e1BilanceKvet": {
                    "bilanceKvet": [
                        {
                            "palivoPouziteNaKvet": "Zemni plyn",
                            "vsazkaPalivaHodnota": row["vsazkaPalivaHodnota"],
                            "vyrobaElektrinyBrutto": row["vyrobaElektrinyBrutto"],
                            "dodavkaUzitecnehoTepla": row["dodavkaUzitecnehoTepla"],
                        }
                    ]
                },
                "e1TechnologieKvet": {
                    "technologieKvet": [
                        {
                            "technologieKvet": "spalovací pístový motor s rekuperací tepla",
                            "instalovanyTepelnyVykon": row["instalovanyTepelnyVykon"],
                            "instalovanyElektrickyVykon": row["instalovanyElektrickyVykon"],
                        }
                    ],
                    "pocetTechnologii": "1"
                },
                "e1BilanceDodavekAZdroju": {
                    "bilanceDodavekAZdrojuI": [
                        {
                            "saldo": 0,
                            "ztraty": row["bde1-ztraty"],
                            "nakupOdber": row["bde1-nakupOdber"],
                            "bilanceZdroju": "elektro",
                            "bilancniRozdil": 0,
                            "vlastniSpotrebaCelkem": row["bde1-vlastniSpotrebaCelkem"],
                            "spotrebaElektrinyNaPrecerpavaniVPve": row["bde1-spotrebaElektrinyNaPrecerpavaniVPve"]
                        },
                        {
                            "saldo": 0,
                            "ztraty": row["bdt1-ztraty"],
                            "nakupOdber": row["bdt1-nakupOdber"],
                            "bilanceZdroju": "teplo",
                            "bilancniRozdil": 0,
                            "vlastniSpotrebaCelkem": row["bdt1-vlastniSpotrebaCelkem"],
                            "spotrebaElektrinyNaPrecerpavaniVPve": 0
                        }
                    ],
                    "bilanceDodavekAZdrojuII": [
                        {
                            "doprava": row["bde2-doprava"],
                            "ostatni": row["bde2-ostatni"],
                            "prumysl": row["bde2-prumysl"],
                            "domacnosti": row["bde2-domacnosti"],
                            "energetika": row["bde2-energetika"],
                            "stavebnictvi": row["bde2-stavebnictvi"],
                            "bilanceDodavek": "elektro",
                            "zemedelstviALesnictvi": row["bde2-zemedelstviALesnictvi"],
                            "dodavkyObchodnimSubjektum": row["bde2-dodavkyObchodnimSubjektum"],
                            "obchodSluzbySkolstviZdravotnictvi": row["bde2-obchodSluzbySkolstviZdravotnictvi"]
                        }, {
                            "doprava": row["bdt2-doprava"],
                            "ostatni": row["bdt2-ostatni"],
                            "prumysl": row["bdt2-prumysl"],
                            "domacnosti": row["bdt2-domacnosti"],
                            "energetika": row["bdt2-energetika"],
                            "stavebnictvi": row["bdt2-stavebnictvi"],
                            "bilanceDodavek": "teplo",
                            "zemedelstviALesnictvi": row["bdt2-zemedelstviALesnictvi"],
                            "dodavkyObchodnimSubjektum": row["bdt2-dodavkyObchodnimSubjektum"],
                            "obchodSluzbySkolstviZdravotnictvi": row["bdt2-obchodSluzbySkolstviZdravotnictvi"]
                        }
                    ]
                },
                "e1VyrobaADodavkaElektrinyATepla": {
                    "vyrobaADodavkaElektrinyATepla": [
                         {
                            "typ": "elektro",
                            "ztraty": row["ve-ztraty"],
                            "bruttoVyroba": row["ve-bruttoVyroba"],
                            "pouzitePalivo": "Zemni plyn",
                            "bilancniRozdil": 0,
                            "primeDodavkyCizimSubjektum": row["ve-primeDodavkyCizimSubjektum"],
                            "dodavkyDoVlastnihoPodnikuNeboZarizeni": row["ve-dodavkyDoVlastnihoPodnikuNeboZarizeni"],
                            "technologickaVlastniSpotrebaNaVyrobuTepla": row["ve-technologickaVlastniSpotrebaNaVyrobuTepla"],
                            "technologickaVlastniSpotrebaNaVyrobuElektriny": row["ve-technologickaVlastniSpotrebaNaVyrobuElektriny"]
                        },
                        {
                            "typ": "teplo",
                            "ztraty": row["vt-ztraty"],
                            "bruttoVyroba": row["vt-bruttoVyroba"],
                            "pouzitePalivo": "Zemni plyn",
                            "bilancniRozdil": 0,
                            "primeDodavkyCizimSubjektum": row["vt-primeDodavkyCizimSubjektum"],
                            "dodavkyDoVlastnihoPodnikuNeboZarizeni": row["vt-dodavkyDoVlastnihoPodnikuNeboZarizeni"],
                            "technologickaVlastniSpotrebaNaVyrobuTepla": row["vt-technologickaVlastniSpotrebaNaVyrobuTepla"],
                            "technologickaVlastniSpotrebaNaVyrobuElektriny": row["vt-technologickaVlastniSpotrebaNaVyrobuElektriny"]
                        }
                    ]
                },
                "e1TechnologieAInstalovanyVykonVyrobny": {
                    "kraj": row["kraj"],
                    "idVyrobny": row["idVyrobny"],
                    "nazevVyrobny": row["nazevVyrobny"],
                    "pripojenoKPsDs": row["pripojenoKPsDs"],
                    "technologieVyrobny": "Plynová a spalovací (PSE)",
                    "celkovyInstalovanyTepelnyVykonMWt": row["instalovanyTepelnyVykon"],
                    "celkovyInstalovanyElektrickyVykonMWe": row["instalovanyElektrickyVykon"],
                    "licencovanyCelkovyInstalovanyTepelnyVykonMWt": row["instalovanyTepelnyVykon"],
                    "licencovanyCelkovyInstalovanyElektrickyVykonMWe": row["instalovanyElektrickyVykon"]
                }
            },
            "identifikacniUdajeVykazu": {
                "ico": "29060109",
                "typVykazu": "E1",
                "typPeriody": "MONTH",
                "cisloLicence": ["111018325", ],
                "vykazovanyRok": int(year),
                "datovaSchranka": "n9mpdz8",
                "drzitelLicence": "ČEZ Energo, s.r.o.",
                "kontaktniTelefon": "+420721055966",
                "vykazovanaPerioda": int(row["vykazovanaPerioda"]),
                "odpovednyPracovnik": "Jakub Urban",
                "datumVytvoreniVykazu": datestr,
            }
        }
        return json

    def generate_statements(self) -> List[dict]:

        # Return the list of statements as JSON
        json_base = {
            "typVykazu": "E1",
            "vykazy": self.statements
        }
        
        return json_base

# Usage example:
handler = JsonHandler()

#Local inFile path
# infile = os.path.join(handler.data_dir, 'XML Generování-latest.xlsm')

# File path to the Excel file stored in oneDrive
infile = '/mnt/c/Users/JakubUrban/OneDrive - ČEZ Energo, s.r.o/_Pracovní 2024/02_ERU/XML Generování 24.xlsm'


data = handler.read_file(infile, 'ExportERU')

today = datetime.date.today()
handler.make_statements(date=today, df=data)

json_output = handler.generate_statements()


# Save the JSON to a file

outfile_path = os.path.join(handler.data_dir, 'data.json')
with open(outfile_path, 'w', encoding="utf-8") as outfile:
    json.dump(json_output, outfile, indent=4, ensure_ascii=False)
