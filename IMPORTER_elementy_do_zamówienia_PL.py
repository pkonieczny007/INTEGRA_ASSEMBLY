import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
from tkinter import Tk, filedialog

# Funkcja główna
def main():
    # Wybór pliku Excel
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Wybierz plik Excel", 
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    # Wczytaj dane
    df = pd.read_excel(file_path, dtype=str)

    # Sprawdź wymagane kolumny
    required_columns = ['REFERENCJA_ELEMENTU', 'Algorytm dla Anz', 'ZAMÓWIENIE']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Brakuje kolumny: {col}")

    # Budowa XML
    root_xml = ET.Element("DATAEX")

    for idx, row in df.iterrows():
        cmd = ET.SubElement(root_xml, "COMMAND", Name="Import", TblRef="SALEORDERLINES")
        ET.SubElement(cmd, "FIELD", FldRef="OrdRef", FldValue=row['ZAMÓWIENIE'], FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="ArtRef", FldValue=row['REFERENCJA_ELEMENTU'], FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="LineNum", FldValue=str(idx + 1), FldType="20")
        ET.SubElement(cmd, "FIELD", FldRef="Quantity", FldValue="1", FldType="20")  # lub użyj np. row['ILOŚĆ'] jeśli jest

    # Zapis do bufora z kodowaniem windows-1250
    buffer = BytesIO()
    tree = ET.ElementTree(root_xml)
    tree.write(buffer, encoding="windows-1250", xml_declaration=True)

    # Naprawa nagłówka (opcjonalna)
    xml_bytes = buffer.getvalue().replace(
        b"encoding='utf-8'", b'encoding="windows-1250"'
    )

    # Zapisz do pliku
    output_file = file_path.replace(".xlsx", "_saleorderlines.xml").replace(".xls", "_saleorderlines.xml")
    with open(output_file, "wb") as f:
        f.write(xml_bytes)

    print(f"Zapisano plik: {output_file} z kodowaniem windows-1250.")

# Uruchom
if __name__ == "__main__":
    main()
