import os
import shutil

# Definicja listy plików
lista_plikow = """
GPT - DELETE_MARSZRUT.PY
GPT - 9.0- DODAWANIE_MARSZRUTY+DELETE.PY
GPT - 8.2- DODAWANIE_MARSZRUTY.PY
DOPASOWANIE_DANYCH_DO_MARSZRUTY.PY
STARYdane.xlsx
tabela_mapowania.xlsx
dane.xlsx
""".strip().split('\n')

# Tworzenie nowego folderu "MARSZRUTY", jeśli nie istnieje
folder_docelowy = "MARSZRUTY"
if not os.path.exists(folder_docelowy):
    os.makedirs(folder_docelowy)

# Kopiowanie plików z listy do nowego folderu
for plik in lista_plikow:
    sciezka_zrodlowa = plik
    sciezka_docelowa = os.path.join(folder_docelowy, plik)
    
    # Sprawdzanie, czy plik istnieje przed próbą skopiowania
    if os.path.exists(sciezka_zrodlowa):
        shutil.copy(sciezka_zrodlowa, sciezka_docelowa)
        print(f'Plik {plik} został skopiowany.')
    else:
        print(f'Plik {plik} nie istnieje i nie może zostać skopiowany.')

