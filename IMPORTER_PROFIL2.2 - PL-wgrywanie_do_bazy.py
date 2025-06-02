import pandas as pd
import xml.etree.ElementTree as ET
import json
from io import BytesIO

# Wczytaj dane z Excela
data = pd.read_excel("wykaz.xlsx", dtype=str)  # Wymuszenie typu str dla bezpieczeństwa

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

# Szablon XML dla typu "PROFIL"
profile_template = [
    {"FldRef": "PrdRef", "FldType": "20"},
    {"FldRef": "PrdName", "FldType": "20"},
    {"FldRef": "MatRef",  "FldValue": "[MS]", "FldType": "20"},
    {"FldRef": "Length", "FldType": "100"},
    {"FldRef": "Profile", "FldValue": "PROFIL", "FldType": "20"},
    {"FldRef": "PCategory", "FldValue": "56", "FldType": "100"},
    {"FldRef": "ForSale", "FldValue": "1", "FldType": "30"}
]

# Mapowanie kolumn
column_mapping = {key: find_column_name(data.columns, names) for key, names in mapping.items()}

# Generuj komendy XML dla typu "PROFIL"
for _, row in data.iterrows():
    if row[column_mapping["TYP"]] == "PROFIL":
        command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
        for field in profile_template:
            fld_value = (
                str(row[column_mapping[field["FldRef"]]]).strip()
                if "FldValue" not in field and field["FldRef"] in column_mapping and column_mapping[field["FldRef"]] is not None
                else field.get("FldValue", "")
            )
            ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])

# Zapisz XML do bufora z kodowaniem Windows-1250
buffer = BytesIO()
tree = ET.ElementTree(root)
tree.write(buffer, encoding="windows-1250", xml_declaration=True)

# Zamień nagłówek XML na poprawny (jeśli wpisany jest 'utf-8')
xml_content = buffer.getvalue().replace(
    b'encoding=\'utf-8\'', b'encoding="windows-1250"'
)

# Zapisz do pliku
with open("output_profil.xml", "wb") as f:
    f.write(xml_content)

print("Plik output_profil.xml został zapisany z kodowaniem Windows-1250.")
