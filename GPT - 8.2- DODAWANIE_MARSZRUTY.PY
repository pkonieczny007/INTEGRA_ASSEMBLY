import pandas as pd
import xml.etree.ElementTree as ET

def generuj_xml_komende(prd_ref, wrk_ref, opr_ref, o_order, prev_opr):
    command = ET.Element("COMMAND", Name="Import", TblRef="PRODUCT OPERATIONS")
    
    field_prd_ref = ET.SubElement(command, "FIELD", FldRef="PrdRef", FldValue=prd_ref, FldType="20")
    field_o_order = ET.SubElement(command, "FIELD", FldRef="OOrder", FldValue=str(o_order), FldType="100")
    field_prev_opr = ET.SubElement(command, "FIELD", FldRef="PrevOpr", FldValue=str(prev_opr), FldType="100")
    field_wrk_ref = ET.SubElement(command, "FIELD", FldRef="WrkRef", FldValue=wrk_ref, FldType="20")
    field_opr_ref = ET.SubElement(command, "FIELD", FldRef="OprRef", FldValue=opr_ref, FldType="20")
    
    return command

def generuj_xml_z_excela(plik_excel):
    try:
        df = pd.read_excel(plik_excel)
    except Exception as e:
        print(f"Błąd odczytu pliku Excela: {e}")
        return ""

    root = ET.Element("DATAEX")
    
    max_idx = max(int(col[6:]) for col in df.columns if col.startswith(("WrkRef", "OprRef")) and col[6:].isdigit())

    for _, wiersz in df.iterrows():
        ilosc_komend = sum(not pd.isna(wiersz[f"WrkRef{i}"]) and not pd.isna(wiersz[f"OprRef{i}"]) for i in range(1, max_idx+1))

        if ilosc_komend == 0:
            continue

        prd_ref = wiersz["PrdRef"]

        for i in range(1, ilosc_komend + 1):
            wrk_ref = wiersz[f"WrkRef{i}"]
            opr_ref = wiersz[f"OprRef{i}"]
            o_order = i 
            prev_opr = i - 1

            komenda = generuj_xml_komende(prd_ref, wrk_ref, opr_ref, o_order, prev_opr)
            root.append(komenda)

    tree = ET.ElementTree(root)
    xml_dane = ET.tostring(root, encoding="unicode")

    return xml_dane

if __name__ == "__main__":
    plik_excel = "dane.xlsx"
    xml_dane = generuj_xml_z_excela(plik_excel)

    with open("marszruty.xml", "w") as plik:
        plik.write(xml_dane)

    print("Dane XML zostały wygenerowane i zapisane w pliku wynik.xml.")
