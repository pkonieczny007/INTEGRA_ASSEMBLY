#!/usr/bin/env python3
import os
import pandas as pd

# Ustawienia pandas - by wyświetlenie tabeli w terminalu było przejrzyste
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)

# Mapa: nazwa pliku -> szukany ciąg znaków
files_search = {
    "output_z_bazy_do_zlozen.xml": '<COMMAND Name="Import" TblRef="PR_SSTT_00000100">',
    "output_zlozenie.xml": '<COMMAND Name="Import" TblRef="PRODUCTS">',
    "output_profil.xml": '<COMMAND Name="Import" TblRef="PRODUCTS">',
    "output_INNE.xml": '<COMMAND Name="Import" TblRef="PRODUCTS">'
}

report_data = []

for filename, pattern in files_search.items():
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as file:
            content = file.read()
        count = content.count(pattern)
        count_value = str(count) if count > 0 else "brak"
    else:
        # Jeśli plik nie istnieje, informujemy o tym.
        count_value = "plik nie znaleziony"
    
    # Dodajemy dane do listy wynikowej
    report_data.append({
        "PLIK": filename,
        "ILOŚĆ": count_value,
        "SZCZEGÓŁY": pattern
    })

# Tworzymy DataFrame z raportem
df = pd.DataFrame(report_data)

# Wyświetlamy wynikową tabelkę w terminalu
print("\nRaport wygenerowany na podstawie przeszukania plików:")
print(df.to_string(index=False))

# Zapisujemy raport do plików Excel oraz CSV
excel_filename = "raport.xlsx"
#csv_filename = "raport.csv"
df.to_excel(excel_filename, index=False)
#df.to_csv(csv_filename, index=False)

print(f"\nRaport zapisano do plików: '{excel_filename}'")
