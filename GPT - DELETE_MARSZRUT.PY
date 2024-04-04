import pandas as pd
import xml.etree.ElementTree as ET

def usun_marszrute(prd_ref):
    command = ET.Element("COMMAND", Name="DELETE", TblRef="PRODUCT OPERATIONS")
    conditions = ET.SubElement(command, "CONDITIONS")
    condition = ET.SubElement(conditions, "CONDITION")
    ET.SubElement(condition, "FIELD", FldRef="PrdRef")
    ET.SubElement(condition, "EQUAL")
    ET.SubElement(condition, "CONSTANT", Value=prd_ref, FldType="20")
    return command

def generuj_xml_z_excela(plik_excel):
    try:
        df = pd.read_excel(plik_excel)
    except Exception as e:
        print(f"Błąd odczytu pliku Excela: {e}")
        return ""

    root = ET.Element("DATAEX")

    for _, wiersz in df.iterrows():
        prd_ref = wiersz["PrdRef"]
        
        # Generuj komendę do usunięcia marszruty dla danego PrdRef
        komenda_usuniecia = usun_marszrute(prd_ref)
        root.append(komenda_usuniecia)

    tree = ET.ElementTree(root)
    xml_dane = ET.tostring(root, encoding="unicode")

    return xml_dane

if __name__ == "__main__":
    plik_excel = "dane.xlsx"  # Upewnij się, że ten plik istnieje w bieżącym katalogu
    xml_dane = generuj_xml_z_excela(plik_excel)

    with open("usun_marszruty.xml", "w") as plik:  # Zapisz wygenerowane dane XML do pliku
        plik.write(xml_dane)

    print("Dane XML do usunięcia marszrut zostały wygenerowane i zapisane w pliku 'usun_marszruty.xml'.")
