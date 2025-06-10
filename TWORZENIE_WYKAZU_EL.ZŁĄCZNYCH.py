import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# Lista dopuszczalnych nazw arkuszy
acceptable_sheet_names = ["ZŁĄCZNE", "złączne", "Złączne", "zlaczne", "Zlaczne", "ZLACZNE"]

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

    # Szukanie arkusza o dokładnej nazwie z listy
    for sheet_name in xls.sheet_names:
        if sheet_name in acceptable_sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.to_excel("elementy_złączne.xlsx", index=False)
            print(f"Zapisano arkusz '{sheet_name}' do 'elementy_złączne.xlsx'.")
            return

    print("Nie znaleziono zakładki o nazwie z listy dopuszczalnych nazw:" +
          ", ".join(acceptable_sheet_names))

if __name__ == "__main__":
    main()
