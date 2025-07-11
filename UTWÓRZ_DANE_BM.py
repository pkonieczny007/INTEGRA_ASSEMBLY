import pandas as pd

# 1. Wczytaj wszystkie dane jako stringi, brakujące wypełnij pustym tekstem
df = pd.read_excel('wykaz.xlsx', dtype=str).fillna('')

# 2. Przygotuj DataFrame dla czarnego montażu
df_czarny = df[['REFERENCJA_ELEMENTU', 'MARSZRUTA', 'TYP']].copy()
df_czarny.columns = ['PrdRef', 'MARSZRUTA', 'typ']

# 3. Przygotuj DataFrame dla białego montażu
df_bialy = df[['BM_REFERENCJA', 'MARSZRUTA', 'BM_TYP']].copy()
df_bialy.columns = ['PrdRef', 'MARSZRUTA', 'typ']

# 4. Połącz oba DataFrame’y
dane = pd.concat([df_czarny, df_bialy], ignore_index=True)

# 5. Usuń te wiersze, w których PrdRef jest pusty
dane = dane[dane['PrdRef'].str.strip() != '']

# 6. Dodaj puste kolumny WrkRef1…WrkRef7 i OprRef1…OprRef7
for i in range(1, 8):
    dane[f'WrkRef{i}'] = ''
    dane[f'OprRef{i}'] = ''

# 7. Zapisz do Excela
dane.to_excel('dane.xlsx', index=False)
print(f"Zapisano {len(dane)} wierszy do dane.xlsx")
