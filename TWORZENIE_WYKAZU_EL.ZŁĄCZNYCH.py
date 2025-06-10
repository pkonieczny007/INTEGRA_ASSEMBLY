import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import unicodedata

# Funkcja do normalizacji nazw (usuwanie polskich znaków i zamiana na małe litery)
def normalize(name):
    return ''.join(
        c for c in unicodedata.normalize('NFKD', name)
        if not unicodedata.combining(c)
    ).lower()

def main():
    # Okno dialogowe do wyboru pliku
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Wybierz plik Excel",
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
    )

    if not file_path:
        print("Nie wybrano pliku.")
        return

    # Wczytanie arkuszy
    xls = pd.ExcelFile(file_path)
    sheet_map = {normalize(name): name for name in xls.sheet_names}

    # Szukanie arkusza "złączne" bez uwzględniania wielkości i znaków
    for norm_name, real_name in sheet_map.items():
        if "zlaczne" in norm_name:
            df = pd.read_excel(xls, sheet_name=real_name)
            df.to_excel("elementy_złączne.xlsx", index=False)
            print(f"Zapisano arkusz '{real_name}' do 'elementy_złączne.xlsx'.")
            return

    print("Nie znaleziono zakładki zawierającej 'ZŁĄCZNE'.")

if __name__ == "__main__":
    main()
