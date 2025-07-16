import pandas as pd
import xml.etree.ElementTree as ET
import os
import requests
import json
from io import BytesIO

# Ścieżki
mapping_file = "mapping.json"
github_url = "https://raw.githubusercontent.com/pkonieczny007/INTEGRA_ASSEMBLY/main/mapping.json"
excel_file = "wykaz.xlsx"

# 1. Pobierz mapping.json jeśli nie istnieje
if not os.path.exists(mapping_file):
    response = requests.get(github_url)
    response.raise_for_status()
    with open(mapping_file, "w", encoding="utf-8") as f:
        f.write(response.text)

# 2. Wczytaj dane z Excela
df = pd.read_excel(excel_file, dtype=str).fillna("")

# 3. Wczytaj mapping
with open(mapping_file, "r", encoding="utf-8") as f:
    mapping = json.load(f)

# 4. Lokalizator kolumn
def find_column_name(columns, candidates):
    for name in candidates:
        if name in columns:
            return name
    return None

column_mapping = { key: find_column_name(df.columns, names)
                   for key, names in mapping.items() }

# 5. Uporządkuj bez krzaków
column_mapping.update({
    "ZLOZENIE": "ZŁOŻENIE",
    "REFERENCJA_ELEMENTU": "REFERENCJA_ELEMENTU",
    "Anz": "Anz.",
    "BM_ZŁOŻENIE": "BM_ZŁOŻENIE",
    "BM_REFERENCJA": "BM_REFERENCJA",
    "TYP": column_mapping["TYP"],
    "BM_TYP": column_mapping["BM_TYP"]
})

# 6. Diagnostyka
print("Wszystkie wiersze:", len(df))
print("Czarny montaż (TYP==ZŁOŻENIE):",
      df[column_mapping["TYP"]].str.strip().str.upper().eq("ZŁOŻENIE").sum())
print("Biały montaż (BM_TYP==ZŁOŻENIE):",
      df[column_mapping["BM_TYP"]].str.strip().str.upper().eq("ZŁOŻENIE").sum())

# 7. Budowanie XML
root = ET.Element("DATAEX")
root.append(ET.Comment("ZŁOŻENIA — CZARNY i BIAŁY MONTAŻ (z filtrem pustych i typu)"))

# 7.1 Czarny montaż (TYP)
mask_czarny = df[column_mapping["TYP"]].str.strip().str.upper() == "ZŁOŻENIE"
for _, row in df[mask_czarny].iterrows():
    prd_ref_org = row[column_mapping["ZLOZENIE"]].strip()
    prd_ref_dst = row[column_mapping["REFERENCJA_ELEMENTU"]].strip()
    qty         = row[column_mapping["Anz"]].strip()
    if prd_ref_org and prd_ref_dst:
        cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
        cmd.append(ET.Comment("ZLOZENIE"))
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefOrg", FldValue=prd_ref_org, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefDst", FldValue=prd_ref_dst, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PQUANT",   FldValue=qty,         FldType="100")

# 7.2 Biały montaż (BM_TYP)
mask_bialy = df[column_mapping["BM_TYP"]].str.strip().str.upper() == "ZŁOŻENIE"
for _, row in df[mask_bialy].iterrows():
    prd_ref_org = row[column_mapping["BM_ZŁOŻENIE"]].strip()
    prd_ref_dst = row[column_mapping["BM_REFERENCJA"]].strip()
    qty         = row[column_mapping["Anz"]].strip()
    if prd_ref_org and prd_ref_dst:
        cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
        cmd.append(ET.Comment("BM_ZLOZENIE"))
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefOrg", FldValue=prd_ref_org, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefDst", FldValue=prd_ref_dst, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PQUANT",   FldValue=qty,         FldType="100")
# 7.3 Czarny montaż PROFIL
mask = df[column_mapping["TYP"]].str.strip().str.upper() == "PROFIL"
for _, row in df[mask].iterrows():
    prd_ref_org = row[column_mapping["ZLOZENIE"]].strip()
    prd_ref_dst = row[column_mapping["REFERENCJA_ELEMENTU"]].strip()
    qty         = row[column_mapping["Anz"]].strip()
    if prd_ref_org and prd_ref_dst:
        cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
        cmd.append(ET.Comment("PROFIL"))
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefOrg", FldValue=prd_ref_org, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefDst", FldValue=prd_ref_dst, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PQUANT",   FldValue=qty,         FldType="100")

# 7.4 Biały montaż PROFIL
mask = df[column_mapping["BM_TYP"]].str.strip().str.upper() == "PROFIL"
for _, row in df[mask].iterrows():
    prd_ref_org = row[column_mapping["BM_ZŁOŻENIE"]].strip()
    prd_ref_dst = row[column_mapping["BM_REFERENCJA"]].strip()
    qty         = row[column_mapping["Anz"]].strip()
    if prd_ref_org and prd_ref_dst:
        cmd = ET.SubElement(root, "COMMAND", Name="Import", TblRef="PR_SSTT_00000100")
        cmd.append(ET.Comment("BM_PROFIL"))
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefOrg", FldValue=prd_ref_org, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PrdRefDst", FldValue=prd_ref_dst, FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="PQUANT",   FldValue=qty,         FldType="100")
# 8. Zapis pliku
buf = BytesIO()
ET.ElementTree(root).write(buf, encoding="windows-1250", xml_declaration=True)
xml = buf.getvalue().replace(b"encoding='utf-8'", b"encoding=\"windows-1250\"")
with open("output_z_bazy_do_zlozen.xml", "wb") as f:
    f.write(xml)

print("Zapisano output_z_bazy_do_zlozen.xml")
