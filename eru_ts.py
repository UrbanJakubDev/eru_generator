
import datetime
from typing import List
import pandas as pd
import json
import os
from pprint import pprint
from dataclasses import dataclass, field
from typing import List, Optional
import xml.etree.ElementTree as ET


@dataclass
class XmlHeader:
    hlav_vyplnil_tel: str
    hlav_spol: str
    hlav_ic: str
    hlav_rok: str
    hlav_vyplnil: str
    hlav_datum_vyplnen: str


class XMLHandler:
    def __init__(self, xml_header: object):
        self.heating_systems = []
        self.xml_header = xml_header
        self.data_dir = os.path.join(os.path.dirname(__file__), "data-ts")

        self.xml = None

    def make_xml(self) -> dict:

        header_elements = {
            "hlav_vyplnil_tel": {"value": self.xml_header["hlav_vyplnil_tel"]},
            "hlav_spol": {"value": self.xml_header["hlav_spol"]},
            "hlav_ic": {"value": self.xml_header["hlav_ic"]},
            "hlav_rok": {"value": self.xml_header["hlav_rok"]},
            "hlav_vyplnil": {"value": self.xml_header["hlav_vyplnil"]},
            "hlav_datum_vyplnen": {"value": self.xml_header["hlav_datum_vyplnen"]},
            "hlav_ctvrtlet": {"value": None, "attributes": {"xsi:nil": "true"}},
            "hlav_mesic": {"value": None, "attributes": {"xsi:nil": "true"}},
            "hlav_schvalil": {"value": None, "attributes": {"xsi:nil": "true"}},
            "hlav_schvalil_tel": {"value": None, "attributes": {"xsi:nil": "true"}}
        }

        root = ET.Element("root")
        root.set("xmlns:ERU", "http://ddd.dd.dat")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

        for tag, data in header_elements.items():

            if 'attributes' not in data:
                element = ET.SubElement(root, tag)
            else:
                element = ET.SubElement(root, tag, data["attributes"])

            if data["value"] is not None:
                element.text = data["value"]

        pocet_soustav = ET.SubElement(root, "pocet_soustav")
        pocet_soustav.text = str(len(self.heating_systems))
        
        
        # Get data from xls file using pandas
        infile = os.path.join(self.data_dir, 'file-final.xlsx')
        df = pd.read_excel(infile, sheet_name='List1', skiprows=0, header=0)
        
        # Round the numbers to 2 decimal places
        df = df.round(0)
        df = df.fillna(0)
        
        # XML Body
        skupina_soustavy = ET.SubElement(root, "skupina_soustavy")
        for index, row in df.iterrows():
            soustavy = ET.SubElement(skupina_soustavy, "soustavy")
            nazev_soustavy = ET.SubElement(soustavy, "nazev_soustavy")
            nazev_soustavy.text = row['nazev_soustavy']
            kraj = ET.SubElement(soustavy, "kraj")
            kraj.text = row['kraj']
            
            # vst_skupina_vyrobny
            vst_skupina_vyrobny = ET.SubElement(soustavy, "vst_skupina_vyrobny")
            vst_vyrobny = ET.SubElement(vst_skupina_vyrobny, "vst_vyrobny")
            vst_vyrobny_nazev = ET.SubElement(vst_vyrobny, "vst_vyrobny_nazev")
            vst_vyrobny_nazev.text = row['vst_vyrobny_nazev']
            vst_vyroba = ET.SubElement(soustavy, "vst_vyroba")
            vst_vyroba.text = str(row['vst_vyroba'])
            vst_vyroba_kvet = ET.SubElement(soustavy, "vst_vyroba_kvet")
            vst_vyroba_kvet.text = str(row['vst_vyroba_kvet'])
            vst_skupina_nakup = ET.SubElement(soustavy, "vst_skupina_nakup")
            vst_nakup = ET.SubElement(vst_skupina_nakup, "vst_nakup")
            vst_dodavatel = ET.SubElement(vst_nakup, "vst_dodavatel")
            vst_dodavatel.set("xsi:nil", "true")
            vst_misto_predani = ET.SubElement(vst_nakup, "vst_misto_predani")
            vst_misto_predani.set("xsi:nil", "true")
            vst_mnozstvi = ET.SubElement(vst_nakup, "vst_mnozstvi")
            vst_mnozstvi.set("xsi:nil", "true")
            vyst_skupina_vystup = ET.SubElement(soustavy, "vyst_skupina_vystup")
            vyst_vystup = ET.SubElement(vyst_skupina_vystup, "vyst_vystup")
            vyst_odberatel = ET.SubElement(vyst_vystup, "vyst_odberatel")
            vyst_odberatel.text = row['vyst_odberatel']
            vyst_misto_predani = ET.SubElement(vyst_vystup, "vyst_misto_predani")
            vyst_misto_predani.text = row['vyst_misto_predani']
            vyst_mnozstvi = ET.SubElement(vyst_vystup, "vyst_mnozstvi")
            vyst_mnozstvi.text = str(row['vyst_mnozstvi'])
            vst_neobn_c_mn = ET.SubElement(soustavy, "vst_neobn_c_mn")
            vst_neobn_c_mn.text = str(row['vst_neobn_c_mn'])
            vst_obn_c_mn = ET.SubElement(soustavy, "vst_obn_c_mn")
            vst_obn_c_mn.set("xsi:nil", "true")
            vst_druh_c_mn = ET.SubElement(soustavy, "vst_druh_c_mn")
            vst_druh_c_mn.set("xsi:nil", "true")
            vst_neobn_kv_mn = ET.SubElement(soustavy, "vst_neobn_kv_mn")
            vst_neobn_kv_mn.text = str(row['vst_neobn_kv_mn'])
            vst_obn_kv_mn = ET.SubElement(soustavy, "vst_obn_kv_mn")
            vst_obn_kv_mn.set("xsi:nil", "true")
            cislo_soustavy = ET.SubElement(soustavy, "cislo_soustavy")
            cislo_soustavy.text = str(row['cislo_soustavy'])
            komentar = ET.SubElement(soustavy, "komentar")
            komentar.set("xsi:nil", "true")
            
            


        self.xml = root

    def read_file(self, excel_file_path: str, sheet_name: str = 'Export23'):
        # Assuming your Excel file has a sheet named 'Sheet1'

        df = pd.read_excel(
            excel_file_path, sheet_name=sheet_name, skiprows=0, header=1)
        # Assuming your Excel sheet has columns like 'typPaliva', 'jednotkaPaliva', etc.
        return df.fillna(0)

    def generate_heating_systems(self) -> List[dict]:

        datestr = datetime.datetime.now().strftime("%Y-%m-%d")
        # Return the list of heating_systems as JSON
        self.make_xml()


# Define the type of the xml header
# Object of xml header
header: XmlHeader = {
    "hlav_vyplnil_tel": "+420721055966",
    "hlav_spol": "18530_S",
    "hlav_ic": "29060109",
    "hlav_rok": "2023",
    "hlav_vyplnil": "Tomáš Janoušek",
    "hlav_datum_vyplnen": "2023-02-27"
}


handler = XMLHandler(xml_header=header)
# infile = os.path.join(handler.data_dir, 'XML Generování.xlsm')
# data = handler.read_file(infile)
data = None

today = datetime.date.today()
# handler.make_heating_systems(date=today, df=data)
handler.generate_heating_systems()


# Save the JSON to a file
outfile_path = os.path.join(handler.data_dir, 'data.xml')
tree = ET.ElementTree(handler.xml)
tree.write(outfile_path, encoding='utf-8', xml_declaration=True)
