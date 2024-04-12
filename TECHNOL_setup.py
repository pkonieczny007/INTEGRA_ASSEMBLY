import os
import requests
import zipfile
import io
import shutil

# Definicja listy plików jako wielolinijkowego stringa, łatwego do edycji
lista_plikow = """
sprawdzanie_v.3.0.py
dane.xlsx
wynik.py
AKTUALIZUJ_wykaz.py

""".strip().split('\n')

# Funkcja do pobierania i rozpakowywania repozytorium
def download_and_extract(repo_url, target_folder):
    response = requests.get(repo_url)
    if response.status_code == 200:
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            temp_dir = 'temp_extracted'
            z.extractall(temp_dir)
            extracted_folder = os.path.join(temp_dir, os.listdir(temp_dir)[0])

            # Przenieś zawartość do docelowego folderu
            os.makedirs(target_folder, exist_ok=True)
            for item in os.listdir(extracted_folder):
                shutil.move(os.path.join(extracted_folder, item), target_folder)
            shutil.rmtree(temp_dir)
        print(f"Files downloaded and moved to {target_folder}.")
    else:
        print(f"Error downloading: {response.status_code}")

# Tworzenie folderu "TECHNOL"
technol_folder = "TECHNOL"
os.makedirs(technol_folder, exist_ok=True)

# Pobranie danych z GitHub
github_url = "https://github.com/pkonieczny007/sprawdzanie_technologii_zlozeniu/archive/refs/heads/main.zip"
download_and_extract(github_url, technol_folder)

# Porządkowanie danych
archiwum_folder = os.path.join(technol_folder, "archiwum")
os.makedirs(archiwum_folder, exist_ok=True)

# Przenoszenie niepotrzebnych plików do archiwum
for item in os.listdir(technol_folder):
    item_path = os.path.join(technol_folder, item)
    if item not in lista_plikow:
        shutil.move(item_path, archiwum_folder)
        print(f"Moved {item} to archive.")

print("File organization complete.")
