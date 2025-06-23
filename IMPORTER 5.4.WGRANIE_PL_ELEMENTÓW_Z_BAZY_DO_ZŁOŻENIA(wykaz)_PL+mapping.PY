import pandas as pd
from io import BytesIO
import xml.etree.ElementTree as ET
import os
import requests
import json

# Ścieżki
mapping_file = "mapping.json"
github_url = "https://raw.githubusercontent.com/pkonieczny007/INTEGRA_ASSEMBLY/main/mapping.json"
excel_file = "wykaz.xlsx"

# 1. Pobierz mapping.json, jeśli nie istnieje
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

# 2. Wczytaj dane z Excela
df = pd.read_excel(excel_file)

# 3. Wczytaj mapowanie (jeśli potrzebne do rozszerzenia działania)
with open(mapping_file, "r", encoding="utf-8") as f:
    mapping = json.load(f)

# 4. Tworzenie XML
root = ET.Element("DATAEX")
comment = ET.Comment("3. KOMENDA DO WGRYWANIA ELEMENTÓW Z BAZY DO ZŁOŻENIA")
root.append(comment)

for _, row in df.iterrows():
    command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
    ET.SubElement(command, "FIELD", FldRef="PrdRefOrg", FldValue=str(row["ZŁOŻENIE"]), FldType="20")
    ET.SubElement(command, "FIELD", FldRef="PrdRefDst", FldValue=str(row["REFERENCJA_ELEMENTU"]), FldType="20")
    ET.SubElement(command, "FIELD", FldRef="PQUANT", FldValue=str(row["Algorytm dla Anz"]), FldType="100")

# 5. Zapis XML do pliku z kodowaniem windows-1250
buffer = BytesIO()
tree = ET.ElementTree(root)
tree.write(buffer, encoding="windows-1250", xml_declaration=True)

xml_content = buffer.getvalue().replace(
    b'encoding=\'utf-8\'', b'encoding="windows-1250"'
)

with open("output_z_bazy_do_zlozen.xml", "wb") as f:
    f.write(xml_content)

print("Plik XML został utworzony z kodowaniem Windows-1250.")
