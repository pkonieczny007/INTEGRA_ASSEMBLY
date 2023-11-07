import pandas as pd
import xml.etree.ElementTree as ET
import json

# Wczytaj dane z Excela
data = pd.read_excel("wykaz.xlsx")

# Wczytaj mapowanie z pliku JSON
"""with open("mapping.json", "r") as file:
    mapping = json.load(file)""" #zmiana na kodowanie utf-8
 
with open("mapping.json", "r", encoding='utf-8') as file:
    mapping = json.load(file)

# Funkcja do wyszukiwania właściwej nazwy kolumny w danych
def find_column_name(data_columns, possible_names):
    for name in possible_names:
        if name in data_columns:
            return name
    return None

# Utwórz korzeń XML
root = ET.Element("DATAEX")

# Szablon dla TYP "BLACHA_KOOP"
blacha_koop_template = [
    {"FldRef": "PrdRef", "FldType": "20"},
    {"FldRef": "PrdName", "FldType": "20"},
    {"FldRef": "MatRef", "FldType": "20"},
    {"FldRef": "Height",  "FldType": "100"},
    {"FldRef": "Priority", "FldValue": "0", "FldType": "100"},
    {"FldRef": "UData1", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData2", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData3", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData4", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData5", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData6", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData7", "FldValue": "", "FldType": "20"},
    {"FldRef": "UData8", "FldValue": "", "FldType": "20"},
]

# Mapuj kolumny z danych na wartości używane w skrypcie
column_mapping = {key: find_column_name(data.columns, names) for key, names in mapping.items()}

# Przeiteruj po każdym wierszu w danych
for _, row in data.iterrows():
    # Używaj mapowania, aby dostosować się do dynamicznych nazw kolumn
    if row[column_mapping["TYP"]] == "BLACHA_KOOP":
        command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PRODUCTS")
        for field in blacha_koop_template:
            fld_value = (
                str(row[column_mapping[field["FldRef"]]])
                if "FldValue" not in field and field["FldRef"] in column_mapping and column_mapping[field["FldRef"]] is not None
                else field.get("FldValue", "")
            )
            ET.SubElement(command, "FIELD", FldRef=field["FldRef"], FldValue=fld_value, FldType=field["FldType"])

# Zapisz XML do pliku
tree = ET.ElementTree(root)
tree.write("output_blachaKOOP.xml", encoding="utf-8", xml_declaration=True)
