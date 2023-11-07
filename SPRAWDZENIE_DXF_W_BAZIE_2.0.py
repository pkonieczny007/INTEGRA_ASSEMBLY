import pandas as pd
import json
import glob

# Wczytanie mapowania nazw kolumn z pliku JSON
with open("mapping.json", "r", encoding='utf-8') as f:
    mapping = json.load(f)

# Wczytanie danych z pliku "wykaz.xlsx"
# Próba znalezienia prawidłowej nazwy kolumny z pliku mapping.json
for possible_name in mapping['PrdRef']:
    try:
        wykaz = pd.read_excel("wykaz.xlsx", usecols=[possible_name])
        break
    except:
        continue

# Próba wczytania danych z kolumny "TYP"
for possible_name_typ in mapping['TYP']:
    try:
        wykaz_typ = pd.read_excel("wykaz.xlsx", usecols=[possible_name_typ])
        break
    except:
        continue

# Dołączenie kolumny "TYP" do wykazu
wykaz['TYP'] = wykaz_typ

# Wczytanie danych z pliku, którego nazwa zaczyna się na "Produkty"
product_file_path = glob.glob("Produkty*.xlsx")[0]
produkty = pd.read_excel(product_file_path, usecols=["Referencja"])

# Sprawdzenie, czy referencje z "wykaz" występują w "produkty"
wykaz['wystepuje_w_INTEGRA'] = wykaz[possible_name].isin(produkty['Referencja']).map({True: 'TAK', False: 'NIE'})

# Dodanie nazwy pliku produktowego lub 'BRAK' w zależności od tego, czy referencja występuje w "produkty"
wykaz['plik_produktowy'] = wykaz['wystepuje_w_INTEGRA'].map({'TAK': product_file_path, 'NIE': 'BRAK'})

# Zmiana nazwy kolumny "possible_name" na "Referencja z wykazu"
wykaz.rename(columns={possible_name: 'Referencja z wykazu'}, inplace=True)

# Zapisanie wyników do nowego pliku Excel
wykaz.to_excel("output_sprawdzenie_dxf_z_bazy.xlsx", index=False)
