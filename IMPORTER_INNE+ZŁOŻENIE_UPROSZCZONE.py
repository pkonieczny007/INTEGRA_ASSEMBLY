#!/usr/bin/env python3
import os
import sys
import urllib.request
import pandas as pd
import xml.etree.ElementTree as ET
import json
import xml.dom.minidom as minidom

GITHUB_MAPPING_URL = "https://raw.githubusercontent.com/pkonieczny007/INTEGRA_ASSEMBLY/main/mapping.json"
MAPPING_FILE = "mapping.json"
EXCEL_FILE = "wykaz.xlsx"

def fetch_mapping():
    """
    Pobiera plik mapping.json z GitHub, jeśli nie istnieje lokalnie.
    """
    if not os.path.exists(MAPPING_FILE):
        try:
            print(f"Pobieram {MAPPING_FILE} z GitHub...")
            urllib.request.urlretrieve(GITHUB_MAPPING_URL, MAPPING_FILE)
            print("Pobrano mapping.json.")
        except Exception as e:
            print(f"Błąd pobierania {MAPPING_FILE}: {e}")
            sys.exit(1)

def find_column_name(data_columns, possible_names):
    """
    Funkcja przeszukująca listę kolumn pod kątem występowania jednej z możliwych nazw.
    """
    for name in possible_names:
        if name in data_columns:
            return name
    return None

def main():
    # Upewnij się, że mamy mapping.json
    fetch_mapping()

    # Wczytaj dane z pliku Excel
    try:
        data = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"Błąd podczas wczytywania pliku Excel '{EXCEL_FILE}': {e}")
        sys.exit(1)
    
    # Wczytaj mapowanie z pliku mapping.json (kodowanie UTF-8)
    try:
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            mapping = json.load(f)
    except Exception as e:
        print(f"Błąd podczas wczytywania pliku '{MAPPING_FILE}': {e}")
        sys.exit(1)
    
    # Utwórz korzeń XML
    root = ET.Element("DATAEX")
    
    # Utwórz mapowanie nazw kolumn z pliku mapping.json
    column_mapping = {key: find_column_name(data.columns, names) for key, names in mapping.items()}
    
    # Szablony dla nowych produktów
    zlozenie_template = [
        {"FldRef": "PrdRef", "FldType": "20"},
        {"FldRef": "PrdName", "FldType": "20"},
        {"FldRef": "Assembly", "FldValue": "1", "FldType": "100"},
        {"FldRef": "PCATEGORY", "FldValue": "2", "FldType": "100"},
        {"FldRef": "ForSale", "FldValue": "1", "FldType": "30"}
    ]
    inne_template = [
        {"FldRef": "PrdRef", "FldType": "20"},
        {"FldRef": "PrdName", "FldType": "20"},
        {"FldRef": "StdCost", "FldValue": "0", "FldType": "10"},
        {"FldRef": "DIS_PClass", "FldValue": "1", "FldType": "100"},
        {"FldRef": "ForSale", "FldValue": "1", "FldType": "30"},
        {"FldRef": "Weight", "FldValue": "0", "FldType": "100"}
    ]
    
    # Nowe produkty
    root.append(ET.Comment("Nowe produkty: ZŁOŻENIE i INNE"))
    for _, row in data.iterrows():
        typ = row[column_mapping.get("TYP")] if column_mapping.get("TYP") else None
        if typ == "ZŁOŻENIE":
            cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in zlozenie_template:
                if "FldValue" in field:
                    val = field["FldValue"]
                else:
                    col = column_mapping.get(field["FldRef"])
                    val = str(row[col]) if col else ""
                ET.SubElement(cmd, "FIELD", FldRef=field["FldRef"], FldValue=val, FldType=field["FldType"])
        elif typ == "INNE":
            cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in inne_template:
                if "FldValue" in field:
                    val = field["FldValue"]
                else:
                    col = column_mapping.get(field["FldRef"])
                    val = str(row[col]) if col else ""
                ET.SubElement(cmd, "FIELD", FldRef=field["FldRef"], FldValue=val, FldType=field["FldType"])
    
    # Elementy w złożeniach
    root.append(ET.Comment("Wgrywanie elementów do złożeń"))
    for _, row in data.iterrows():
        cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefOrg", FldValue=str(row.get("ZŁOŻENIE", "")), FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefDst", FldValue=str(row.get("REFERENCJA_ELEMENTU", "")), FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PQUANT",    FldValue=str(row.get("Algorytm dla Anz", "")), FldType="100")
    
    # Zapis do pliku
    output_file = "output_combined.xml"
    xml_str = ET.tostring(root, encoding="utf-8")
    dom = minidom.parseString(xml_str)
    pretty_xml = dom.toprettyxml(indent="  ", newl="\n")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(pretty_xml)
    print(f"Plik XML zapisany jako: {output_file}")

if __name__ == "__main__":
    main()
