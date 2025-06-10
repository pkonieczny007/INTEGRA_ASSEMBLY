import pandas as pd
import xml.etree.ElementTree as ET
from tkinter import Tk, filedialog

def main():
    # Okno dialogowe do wyboru pliku
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Wybierz plik Excel", 
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    # Wczytaj dane z Excela
    df = pd.read_excel(file_path)

    # Sprawdzenie wymaganych kolumn
    required_columns = ['REFERENCJA_ELEMENTU', 'Algorytm dla Anz', 'ZAMÓWIENIE']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Brakuje wymaganej kolumny: {col}")

    # Tworzenie struktury XML
    root_xml = ET.Element("DATAEX")
    for idx, row in df.iterrows():
        cmd = ET.SubElement(root_xml, "COMMAND", Name="Import", TblRef="SALEORDERLINES")
        ET.SubElement(cmd, "FIELD", FldRef="OrdRef", FldValue=str(row['ZAMÓWIENIE']), FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="ArtRef", FldValue=str(row['REFERENCJA_ELEMENTU']), FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="LineNum", FldValue=str(idx + 1), FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="Quantity", FldValue="1", FldType="20")  # Zakładamy ilość = 1

    # Zapisz plik XML
    output_file = file_path.replace(".xlsx", "_saleorderlines.xml").replace(".xls", "_saleorderlines.xml")
    tree = ET.ElementTree(root_xml)
    tree.write(output_file, encoding="utf-8", xml_declaration=True)
    print(f"Plik XML zapisany jako: {output_file}")

if __name__ == "__main__":
    main()
