import pandas as pd
import xml.etree.ElementTree as ET
from tkinter import Tk, filedialog
from io import BytesIO

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

    # Bufor dla zapisu w Windows-1250
    buffer = BytesIO()
    tree = ET.ElementTree(root_xml)
    tree.write(buffer, encoding="windows-1250", xml_declaration=True)

    # Poprawienie nagłówka XML, jeśli domyślnie ustawione na utf-8
    xml_content = buffer.getvalue().replace(
        b"encoding='utf-8'", b'encoding="windows-1250"'
    )

    # Zapisz plik XML
    output_file = file_path.replace(".xlsx", "_saleorderlines.xml").replace(".xls", "_saleorderlines.xml")
    with open(output_file, "wb") as f:
        f.write(xml_content)

    print(f"Plik XML zapisany jako: {output_file}")

if __name__ == "__main__":
    main()
