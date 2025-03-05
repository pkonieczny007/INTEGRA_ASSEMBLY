#!/usr/bin/env python3
import pandas as pd
import xml.etree.ElementTree as ET
import json
import subprocess
import sys

def find_column_name(data_columns, possible_names):
    """
    Funkcja przeszukująca listę kolumn pod kątem występowania jednej z możliwych nazw.
    """
    for name in possible_names:
        if name in data_columns:
            return name
    return None

def generate_zlozenie_xml(data, mapping, output_file="output_zlozenie.xml"):
    """
    Generuje plik XML dla komend "ZŁOŻENIE" na podstawie danych z Excela i mapowania.
    """
    # Utwórz korzeń XML
    root = ET.Element("DATAEX")
    
    # Szablon XML dla typu "ZŁOŻENIE"
    assembly_template = [
        {"FldRef": "PrdRef", "FldType": "20"},
        {"FldRef": "PrdName", "FldType": "20"},
        {"FldRef": "Assembly", "FldValue": "1", "FldType": "100"},
        {"FldRef": "PCATEGORY", "FldValue": "2", "FldType": "100"},
        {"FldRef": "ForSale", "FldValue": "1", "FldType": "30"}
    ]
    
    # Mapowanie kolumn według mapping.json
    column_mapping = {key: find_column_name(data.columns, names) for key, names in mapping.items()}
    
    # Iteracja przez każdy wiersz danych
    for _, row in data.iterrows():
        # Sprawdzenie, czy wartość w kolumnie odpowiadającej "TYP" to "ZŁOŻENIE"
        typ_col = column_mapping.get("TYP")
        if typ_col is not None and row[typ_col] == "ZŁOŻENIE":
            command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in assembly_template:
                # Jeśli nie zdefiniowano FldValue na stałe, pobieramy wartość z odpowiedniej kolumny
                if "FldValue" not in field:
                    col_name = column_mapping.get(field["FldRef"])
                    fld_value = str(row[col_name]) if col_name is not None else ""
                else:
                    fld_value = field.get("FldValue", "")
                ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])
    
    # Zapisz wygenerowany XML do pliku
    tree = ET.ElementTree(root)
    tree.write(output_file, encoding="utf-8", xml_declaration=True)
    print(f"Plik XML dla ZŁOŻENIA zapisany jako: {output_file}")

def generate_z_bazy_xml(data, output_file="output_z_bazy_do_zlozen.xml"):
    """
    Generuje plik XML z komendami wgrywania elementów do złożenia.
    Uwaga: Zakładamy, że w pliku Excel występują kolumny "ZŁOŻENIE", "REFERENCJA_ELEMENTU"
    oraz "Algorytm dla Anz".
    """
    xml_lines = []
    xml_lines.append('<DATAEX>')
    xml_lines.append('<!-- 3. KOMENDA DO WGRYWANIA ELEMENTÓW Z BAZY DO ZŁOŻENIA -->')
    
    for _, row in data.iterrows():
        xml_lines.append('<COMMAND Name="Import" TblRef="PR_SSTT_00000100">')
        xml_lines.append(f'    <FIELD FldRef="PrdRefOrg" FldValue="{row["ZŁOŻENIE"]}" FldType="20"/>')
        xml_lines.append(f'    <FIELD FldRef="PrdRefDst" FldValue="{row["REFERENCJA_ELEMENTU"]}" FldType="20"/>')
        xml_lines.append(f'    <FIELD FldRef="PQUANT" FldValue="{row["Algorytm dla Anz"]}" FldType="100"/>')
        xml_lines.append('</COMMAND>')
    
    xml_lines.append('</DATAEX>')
    
    xml_content = "\n".join(xml_lines)
    with open(output_file, "w", encoding="utf-8") as xml_file:
        xml_file.write(xml_content)
    print(f"Plik XML dla wgrywania elementów zapisany jako: {output_file}")

def upload_file_via_powersync(file_path):
    """
    Funkcja wywołująca narzędzie powersync w celu wgrania pliku.
    Zakładamy, że powersync jest dostępny w PATH i akceptuje komendę "upload".
    """
    try:
        subprocess.run(["powersync", "upload", file_path], check=True)
        print(f"Plik {file_path} został wgrany przez powersync.")
    except subprocess.CalledProcessError as e:
        print(f"Błąd podczas wgrywania pliku {file_path} przez powersync: {e}")
    except FileNotFoundError:
        print("Narzędzie powersync nie zostało znalezione. Upewnij się, że jest zainstalowane i dostępne w PATH.")

def main():
    # Wczytaj dane z pliku Excel (domyślnie wykaz.xlsx)
    try:
        data = pd.read_excel("wykaz.xlsx")
    except Exception as e:
        print(f"Błąd podczas wczytywania pliku Excel: {e}")
        sys.exit(1)
    
    # Wczytaj mapowanie z pliku mapping.json z kodowaniem UTF-8
    try:
        with open("mapping.json", "r", encoding="utf-8") as file:
            mapping = json.load(file)
    except Exception as e:
        print(f"Błąd podczas wczytywania pliku mapping.json: {e}")
        sys.exit(1)
    
    # Generowanie XML dla ZŁOŻENIA
    xml_zlozenie = "output_zlozenie.xml"
    generate_zlozenie_xml(data, mapping, xml_zlozenie)
    
    # Generowanie XML dla wgrywania elementów do złożeń
    xml_z_bazy = "output_z_bazy_do_zlozen.xml"
    generate_z_bazy_xml(data, xml_z_bazy)
    
    # Wgrywanie plików XML przy użyciu powersync
    upload_file_via_powersync(xml_zlozenie)
    upload_file_via_powersync(xml_z_bazy)

if __name__ == "__main__":
    main()
