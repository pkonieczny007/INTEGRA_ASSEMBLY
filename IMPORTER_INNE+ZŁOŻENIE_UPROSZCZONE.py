#!/usr/bin/env python3
import pandas as pd
import xml.etree.ElementTree as ET
import json
import sys
import xml.dom.minidom as minidom

def find_column_name(data_columns, possible_names):
    """
    Funkcja przeszukująca listę kolumn pod kątem występowania jednej z możliwych nazw.
    """
    for name in possible_names:
        if name in data_columns:
            return name
    return None

def main():
    # Wczytaj dane z pliku Excel
    try:
        data = pd.read_excel("wykaz.xlsx")
    except Exception as e:
        print(f"Błąd podczas wczytywania pliku Excel: {e}")
        sys.exit(1)
    
    # Wczytaj mapowanie z pliku mapping.json (kodowanie UTF-8)
    try:
        with open("mapping.json", "r", encoding="utf-8") as file:
            mapping = json.load(file)
    except Exception as e:
        print(f"Błąd podczas wczytywania pliku mapping.json: {e}")
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
    
    # Dodaj komentarz dla sekcji tworzenia nowych produktów
    root.append(ET.Comment("Nowe produkty: ZŁOŻENIE i INNE"))
    
    # Przetwórz każdy wiersz danych – tworzenie nowych produktów
    for _, row in data.iterrows():
        typ = None
        if column_mapping.get("TYP"):
            typ = row[column_mapping["TYP"]]
        
        # Jeśli typ to "ZŁOŻENIE" – użyj szablonu złożenia
        if typ == "ZŁOŻENIE":
            command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in zlozenie_template:
                if "FldValue" not in field:
                    col_name = column_mapping.get(field["FldRef"])
                    fld_value = str(row[col_name]) if col_name is not None else ""
                else:
                    fld_value = field["FldValue"]
                ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])
        # Jeśli typ to "INNE" – użyj szablonu INNE
        elif typ == "INNE":
            command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in inne_template:
                if "FldValue" not in field:
                    col_name = column_mapping.get(field["FldRef"])
                    fld_value = str(row[col_name]) if col_name is not None else ""
                else:
                    fld_value = field["FldValue"]
                ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])
    
    # Dodaj komentarz dla sekcji wgrywania elementów do złożeń
    root.append(ET.Comment("Wgrywanie elementów do złożeń"))
    
    # Dla każdego wiersza danych dodaj polecenie wgrywania elementów do złożeń.
    # Zakładamy, że w danych istnieją kolumny: "ZŁOŻENIE", "REFERENCJA_ELEMENTU" oraz "Algorytm dla Anz"
    for _, row in data.iterrows():
        command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
        ET.SubElement(command, "FIELD", FldRef="PrdRefOrg", FldValue=str(row["ZŁOŻENIE"]), FldType="20")
        ET.SubElement(command, "FIELD", FldRef="PrdRefDst", FldValue=str(row["REFERENCJA_ELEMENTU"]), FldType="20")
        ET.SubElement(command, "FIELD", FldRef="PQUANT", FldValue=str(row["Algorytm dla Anz"]), FldType="100")
    
    # Zapisz wygenerowany XML do pliku
    output_file = "output_combined.xml"
    xml_str = ET.tostring(root, encoding="utf-8")
    dom = minidom.parseString(xml_str)
    pretty_xml = dom.toprettyxml(indent="  ", newl="\n")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(pretty_xml)
    print(f"Plik XML zapisany jako: {output_file}")

if __name__ == "__main__":
    main()
