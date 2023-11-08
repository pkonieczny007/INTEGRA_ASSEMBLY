import pandas as pd
import os
from shutil import copyfile

# Kopiowanie oryginalnego pliku dane.xlsx do nowego pliku dane_org.xlsx
copyfile('dane.xlsx', 'dane_org.xlsx')

# Wczytanie danych z pliku dane_org.xlsx, który jest teraz kopią oryginalnych danych
dane_df = pd.read_excel('dane_org.xlsx')

# Wczytanie mapowania z pliku tabela_mapowania.xlsx
mapowanie_df = pd.read_excel('tabela_mapowania.xlsx')

# Funkcja do wyszukiwania i dopasowywania danych mapowania
def dopasuj_dane(dane, mapowanie):
    # Tworzenie pustego DataFrame do przechowywania wyników
    wyniki_df = dane.copy()
    # Dodanie kolumn dla WrkRef i OprRef
    for i in range(1, 8):
        wyniki_df[f'WrkRef{i}'] = None
        wyniki_df[f'OprRef{i}'] = None
    
    # Iterowanie przez wiersze DataFrame z danymi
    for index, row in dane.iterrows():
        # Pobranie odpowiedniego wiersza mapowania
        mapowanie_row = mapowanie[(mapowanie['typ'] == row['typ']) & (mapowanie['MARSZRUTA'] == row['MARSZRUTA'])]
        
        # Jeśli znaleziono dopasowanie, aktualizacja wynikowego DataFrame
        if not mapowanie_row.empty:
            for i in range(1, 8):
                if pd.notna(mapowanie_row.iloc[0][f'WrkRef{i}']):
                    wyniki_df.at[index, f'WrkRef{i}'] = mapowanie_row.iloc[0][f'WrkRef{i}']
                if pd.notna(mapowanie_row.iloc[0][f'OprRef{i}']):
                    wyniki_df.at[index, f'OprRef{i}'] = mapowanie_row.iloc[0][f'OprRef{i}']
    
    return wyniki_df

# Wykonanie dopasowania
wynikowe_dane_df = dopasuj_dane(dane_df, mapowanie_df)

# Zapisanie wyników do tego samego pliku Excel (dane.xlsx), nadpisując go
wynikowe_dane_df.to_excel('dane.xlsx', index=False)

print("Dopasowanie danych marszruty zostało zakończone i zapisane w pliku dane.xlsx. Oryginalne dane są dostępne w pliku dane_org.xlsx.")
