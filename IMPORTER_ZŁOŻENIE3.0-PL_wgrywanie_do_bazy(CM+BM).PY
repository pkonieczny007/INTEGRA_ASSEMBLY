import pandas as pd
import xml.etree.ElementTree as ET
import json
import os
from io import BytesIO
import requests

# Ścieżki i pliki
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
data = pd.read_excel(excel_file, dtype=str).fillna("")

# Wczytaj mapowanie
with open(mapping_file, "r", encoding="utf-8") as file:
    mapping = json.load(file)

# Znajdź nazwę kolumny z Excela dla danego pola
def find_column_name(data_columns, possible_names):
    for name in possible_names:
        if name in data_columns:
            return name
    return None

# Tworzenie mapowania kolumn
column_mapping = {
    key: find_column_name(data.columns, names)
    for key, names in mapping.items()
}

# Szablon dla ZŁOŻENIA
assembly_template = [
    {"FldRef": "PrdRef", "FldType": "20"},
    {"FldRef": "PrdName", "FldType": "20"},
    {"FldRef": "Assembly", "FldValue": "1", "FldType": "100"},
    {"FldRef": "PCATEGORY", "FldValue": "2", "FldType": "100"},
    {"FldRef": "ForSale", "FldValue": "1", "FldType": "30"}
]

# Utwórz XML
root = ET.Element("DATAEX")
root.append(ET.Comment("KOMENDA IMPORTU DLA ELEMENTÓW TYPU ZŁOŻENIE I BM_ZŁOŻENIE"))

# 1. Przetwarzanie TYP == ZŁOŻENIE
if column_mapping.get("TYP"):
    for _, row in data.iterrows():
        if str(row.get(column_mapping["TYP"], "")).strip().upper() == "ZŁOŻENIE":
            command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in assembly_template:
                ref = field["FldRef"]
                if "FldValue" in field:
                    fld_value = field["FldValue"]
                else:
                    col = column_mapping.get(ref, "")
                    fld_value = str(row.get(col, "")).strip()
                ET.SubElement(command, "FIELD", FldRef=ref, FldValue=fld_value, FldType=field["FldType"])

# 2. Przetwarzanie BM_TYP == ZŁOŻENIE
if column_mapping.get("BM_TYP") and column_mapping.get("BM_REFERENCJA"):
    for _, row in data.iterrows():
        if str(row.get(column_mapping["BM_TYP"], "")).strip().upper() == "ZŁOŻENIE":
            command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
            for field in assembly_template:
                ref = field["FldRef"]
                if ref == "PrdRef":
                    fld_value = str(row.get(column_mapping["BM_REFERENCJA"], "")).strip()
                elif ref == "PrdName":
                    fld_value = str(row.get(column_mapping.get("BM_NAME", "BM_REFERENCJA"), "")).strip()
                else:
                    fld_value = field.get("FldValue", "")
                ET.SubElement(command, "FIELD", FldRef=ref, FldValue=fld_value, FldType=field["FldType"])

# 3. Zbieranie unikalnych złożeń
zlozenia = pd.Series(dtype=str)

if column_mapping.get("ZLOZENIE") and column_mapping["ZLOZENIE"] in data.columns:
    zlozenia = pd.concat([zlozenia, data[column_mapping["ZLOZENIE"]]])

if column_mapping.get("BM_ZŁOŻENIE") and column_mapping["BM_ZŁOŻENIE"] in data.columns:
    zlozenia = pd.concat([zlozenia, data[column_mapping["BM_ZŁOŻENIE"]]])

zlozenia = zlozenia.dropna().unique()



# 4. Dodanie unikalnych złożeń jako osobnych pozycji
for ref in zlozenia:
    ref = str(ref).strip()
    if ref == "":
        continue
    command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
    for field in assembly_template:
        fld_value = ref if field["FldRef"] in ["PrdRef", "PrdName"] else field.get("FldValue", "")
        ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])

# 5. Zapisz XML
buffer = BytesIO()
tree = ET.ElementTree(root)
tree.write(buffer, encoding="windows-1250", xml_declaration=True)

xml_content = buffer.getvalue().replace(
    b'encoding=\'utf-8\'', b'encoding="windows-1250"'
)

with open("output_zlozenie.xml", "wb") as f:
    f.write(xml_content)

print("Plik 'output_zlozenie.xml' został zapisany.")
