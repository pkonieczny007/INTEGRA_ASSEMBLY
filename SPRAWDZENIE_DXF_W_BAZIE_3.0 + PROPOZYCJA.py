import pandas as pd
import json
import glob
import re

# Wczytanie mapowania nazw kolumn z pliku JSON
with open("mapping.json", "r", encoding='utf-8') as f:
    mapping = json.load(f)

# Wczytanie danych z pliku "wykaz.xlsx"
for possible_name in mapping['PrdRef']:
    try:
        wykaz = pd.read_excel("wykaz.xlsx", usecols=[possible_name])
        break
    except:
        continue

# Wczytanie danych z kolumny "TYP"
for possible_name_typ in mapping['TYP']:
    try:
        wykaz_typ = pd.read_excel("wykaz.xlsx", usecols=[possible_name_typ])
        break
    except:
        continue

# Dołączenie kolumny "TYP" do wykazu
wykaz['TYP'] = wykaz_typ

# Wczytanie danych z pliku produktowego
product_file_path = glob.glob("Produkty*.xlsx")[0]
produkty = pd.read_excel(product_file_path)

# Funkcja do ekstrakcji numeru rysunku i pozycji
def extract_drawing_and_position(ref):
    parts = ref.split('_')
    drawing = parts[2]  # Rysunek jest po drugim znaku "_"
    position = re.search(r'p\d+', ref)  # Pozycja zaczyna się od "p" i ma cyfry
    position = position.group() if position else None
    return drawing, position

# Wyszukiwanie propozycji
def find_proposal(ref, produkty_df):
    drawing, position = extract_drawing_and_position(ref)
    drawing_regex = re.escape(drawing) + r'.*'
    position_regex = re.escape(position) + r'.*' if position else '.*'

    # Szukanie pasujących rysunków z i bez "SL"
    matches_drawing = produkty_df[produkty_df['Referencja'].str.contains(drawing_regex, regex=True)]
    matches_position = matches_drawing[matches_drawing['Referencja'].str.contains(position_regex, regex=True)]

    if not matches_position.empty:
        return matches_position['Referencja'].iloc[0]  # Zwróć pierwszą pasującą pozycję
    elif not matches_drawing.empty:
        return matches_drawing['Referencja'].iloc[0]  # Zwróć pierwszy pasujący rysunek
    else:
        # Spróbuj znaleźć bez "SL"
        drawing_without_sl = drawing.replace("SL", "", 1)
        matches_drawing_without_sl = produkty_df[produkty_df['Referencja'].str.contains(drawing_without_sl)]
        if not matches_drawing_without_sl.empty:
            return matches_drawing_without_sl['Referencja'].iloc[0]
    return 'Brak propozycji'

# Sprawdzenie, czy referencje z "wykaz" występują w "produkty"
wykaz['wystepuje_w_INTEGRA'] = wykaz[possible_name].isin(produkty['Referencja']).map({True: 'TAK', False: 'BRAK'})

# Propozycje dla brakujących referencji
wykaz['PROPOZYCJA'] = wykaz.apply(lambda row: find_proposal(row[possible_name], produkty) if row['wystepuje_w_INTEGRA'] == 'BRAK' else '', axis=1)

# Zapisanie wyników do nowego pliku Excel
wykaz.to_excel("output_sprawdzenie_dxf_z_bazy.xlsx", index=False)
