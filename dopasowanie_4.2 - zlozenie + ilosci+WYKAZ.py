import pandas as pd

def find_zlozenie(df):
    zlozenie_list = []
    for i, row in df.iterrows():
        stufe = row['Stufe']
        nazwa = row['REFERENCJA_ELEMENTU']
        zlozenie = ""
        for j in range(i - 1, -1, -1):
            prev_stufe = df.loc[j, 'Stufe']
            prev_nazwa = df.loc[j, 'REFERENCJA_ELEMENTU']
            if prev_stufe == stufe - 1:
                zlozenie = prev_nazwa
                break
        zlozenie_list.append(zlozenie)
    return zlozenie_list

# Wczytujemy dane z pliku Excela do DataFrame
excel_file = 'wykaz.xlsx'  # Podaj właściwą nazwę pliku Excela
df = pd.read_excel(excel_file)

# Znajdujemy wartości ZŁOŻENIE dla każdego poziomu
df['ZŁOŻENIE'] = find_zlozenie(df)

# Obliczamy wartości "Algorytm dla Anz" na podstawie kolumny "Anz" i "ZŁOŻENIE"
for i, row in df.iterrows():
    if i == 0 or row['Stufe'] == 1:
        df.at[i, 'Algorytm dla Anz'] = row['Anz']
    else:
        zlozenie = row['ZŁOŻENIE']
        prev_anz = df[df['REFERENCJA_ELEMENTU'] == zlozenie].iloc[0]['Anz']
        df.at[i, 'Algorytm dla Anz'] = row['Anz'] / prev_anz

# Nazwa pliku wynikowego
output_excel_file = 'wynik_dopasowanie.xlsx'

# Zapisujemy DataFrame z dodanymi kolumnami "ZŁOŻENIE" i "Algorytm dla Anz" do nowego pliku Excela
df.to_excel(output_excel_file, index=False)

print(f"Plik wynikowy został zapisany jako '{output_excel_file}'.")
