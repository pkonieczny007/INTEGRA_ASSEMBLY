import pandas as pd
import xml.etree.ElementTree as ET
import json
from io import BytesIO

# Wczytaj dane z Excela
data = pd.read_excel("wykaz.xlsx", dtype=str)  # Wymuszamy tekst

# Wczytaj mapowanie z pliku JSON
with open("mapping.json", "r", encoding="utf-8") as file:
    mapping = json.load(file)

# Funkcja do wyszukiwania właściwej nazwy kolumny
def find_column_name(data_columns, possible_names):
    for name in possible_names:
        if name in data_columns:
            return name
    return None

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

# Mapowanie nazw kolumn
column_mapping = {key: find_column_name(data.columns, names) for key, names in mapping.items()}

# Generuj wpisy XML
for _, row in data.iterrows():
    if row[column_mapping["TYP"]] == "ZŁOŻENIE":
        command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
        for field in assembly_template:
            fld_value = (
                str(row[column_mapping[field["FldRef"]]]).strip()
                if "FldValue" not in field and field["FldRef"] in column_mapping and column_mapping[field["FldRef"]] is not None
                else field.get("FldValue", "")
            )
            ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])

# Zapisz do bufora z kodowaniem Windows-1250
buffer = BytesIO()
tree = ET.ElementTree(root)
tree.write(buffer, encoding="windows-1250", xml_declaration=True)

# Zamień nagłówek XML na poprawny, jeśli trzeba
xml_content = buffer.getvalue().replace(
    b'encoding=\'utf-8\'', b'encoding="windows-1250"'
)

# Zapisz do pliku
with open("output_zlozenie.xml", "wb") as f:
    f.write(xml_content)

print("Plik output_zlozenie.xml został zapisany z kodowaniem Windows-1250.")
