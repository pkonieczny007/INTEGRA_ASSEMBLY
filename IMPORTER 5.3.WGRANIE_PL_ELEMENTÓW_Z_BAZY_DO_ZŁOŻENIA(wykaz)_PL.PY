import pandas as pd
from io import BytesIO
import xml.etree.ElementTree as ET

# Wczytanie danych z pliku Excel
df = pd.read_excel("wykaz.xlsx")

# Utwórz korzeń XML
root = ET.Element("DATAEX")
ET.Comment("3. KOMENDA DO WGRYWANIA ELEMENTÓW Z BAZY DO ZŁOŻENIA")

# Tworzenie komend XML
for _, row in df.iterrows():
    command = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
    ET.SubElement(command, "FIELD", FldRef="PrdRefOrg", FldValue=str(row["ZŁOŻENIE"]), FldType="20")
    ET.SubElement(command, "FIELD", FldRef="PrdRefDst", FldValue=str(row["REFERENCJA_ELEMENTU"]), FldType="20")
    ET.SubElement(command, "FIELD", FldRef="PQUANT", FldValue=str(row["Algorytm dla Anz"]), FldType="100")

# Zapis do bufora z kodowaniem windows-1250
buffer = BytesIO()
tree = ET.ElementTree(root)
tree.write(buffer, encoding="windows-1250", xml_declaration=True)

# Zamień nagłówek XML, jeśli trzeba
xml_content = buffer.getvalue().replace(
    b'encoding=\'utf-8\'', b'encoding="windows-1250"'
)

# Zapisz do pliku
with open("output_z_bazy_do_zlozen.xml", "wb") as f:
    f.write(xml_content)

print("Plik XML został utworzony z kodowaniem Windows-1250.")
