import pandas as pd
import xml.etree.ElementTree as ET
import json
import os
import requests
from io import BytesIO

# Ścieżki i link
mapping_file = "mapping.json"
github_url = "https://raw.githubusercontent.com/pkonieczny007/INTEGRA_ASSEMBLY/main/mapping.json"
excel_file = "wykaz.xlsx"

# Pobierz mapping.json jeśli nie istnieje
if not os.path.exists(mapping_file):
    print("Plik mapping.json nie istnieje, pobieram z GitHuba...")
    try:
        response = requests.get(github_url)
        response.raise_for_status()
        with open(mapping_file, "w", encoding="utf-8") as f:
            f.write(response.text)
        print("Pobrano mapping.json")
    except Exception as e:
        print(f"Błąd pobierania mapping.json: {e}")
        exit(1)

# Wczytaj dane z Excela
data = pd.read_excel(excel_file, dtype=str)

# Wczytaj mapowanie z mapping.json
with open(mapping_file, "r", encoding="utf-8") as file:
    mapping = json.load(file)

# Funkcja do znajdowania odpowiednich nazw kolumn
def find_column_name(data_columns, possible_names):
    for name in possible_names:
        if name in data_columns:
            return name
    return None

# Utwórz mapowanie kolumn na podstawie mapping.json
column_mapping = {
    key: find_column_name(data.columns, names)
    for key, names in mapping.items()
}

# Sprawdź czy klucz "TYP" jest zmapowany
if column_mapping.get("TYP") is None:
    print("Brakuje kolumny 'TYP' w danych wejściowych (mapping.json).")
    exit(1)

# Szablon XML dla typu ZŁOŻENIE
assembly_template = [
    {"FldRef": "PrdRef", "FldType": "20"},
    {"FldRef": "PrdName", "FldType": "20"},
    {"FldRef": "Assembly", "FldValue": "1", "FldType": "100"},
    {"FldRef": "PCATEGORY", "FldValue": "2", "FldType": "100"},
    {"FldRef": "ForSale", "FldValue": "1", "FldType": "30"}
]

# Tworzenie XML
root = ET.Element("DATAEX")
root.append(ET.Comment("KOMENDA IMPORTU DLA ELEMENTÓW TYPU ZŁOŻENIE"))

for _, row in data.iterrows():
    if row[column_mapping["TYP"]].strip().upper() == "ZŁOŻENIE":
        command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
        for field in assembly_template:
            ref = field["FldRef"]
            fld_value = (
                str(row[column_mapping[ref]]).strip()
                if "FldValue" not in field and ref in column_mapping and column_mapping[ref]
                else field.get("FldValue", "")
            )
            ET.SubElement(command, "FIELD", FldRef=ref, FldValue=fld_value, FldType=field["FldType"])

# Zapis XML do pliku z kodowaniem Windows-1250
buffer = BytesIO()
tree = ET.ElementTree(root)
tree.write(buffer, encoding="windows-1250", xml_declaration=True)

xml_content = buffer.getvalue().replace(
    b'encoding=\'utf-8\'', b'encoding="windows-1250"'
)

with open("output_zlozenie.xml", "wb") as f:
    f.write(xml_content)

print("Plik 'output_zlozenie.xml' został zapisany z kodowaniem Windows-1250.")
