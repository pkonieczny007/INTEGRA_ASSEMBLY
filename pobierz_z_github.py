import os
import requests

# Adres do pliku z wykazem
wykaz_url = "https://raw.githubusercontent.com/pkonieczny007/INTEGRA_ASSEMBLY/main/wykaz_polskie.txt"
wykaz_file = "wykaz_polskie.txt"
repo_base_url = "https://raw.githubusercontent.com/pkonieczny007/INTEGRA_ASSEMBLY/main/"

# 1. Sprawdź czy istnieje plik wykaz_polskie.txt
if not os.path.exists(wykaz_file):
    print("Brak pliku 'wykaz_polskie.txt'. Pobieram z GitHuba...")
    try:
        response = requests.get(wykaz_url)
        response.raise_for_status()
        with open(wykaz_file, "w", encoding="utf-8") as f:
            f.write(response.text)
        print("Plik wykaz_polskie.txt został pobrany.")
        print("Zakończono działanie. Uruchom ponownie skrypt, aby pobrać pliki z wykazu.")
        exit(0)
    except Exception as e:
        print(f"Błąd podczas pobierania wykazu: {e}")
        exit(1)

# 2. Jeśli plik już istnieje, pobierz wszystkie pliki z listy
with open(wykaz_file, "r", encoding="utf-8") as f:
    pliki = [line.strip() for line in f if line.strip()]

print("Rozpoczynam pobieranie plików z GitHuba:")

for plik in pliki:
    # Zamień np. "IMPORTER_PROFIL2.2 - PL.py" → IMPORTER_PROFIL2.2%20-%20PL.py
    remote_name = plik.replace(" ", "%20")
    url = repo_base_url + remote_name
    local_name = os.path.basename(plik)

    try:
        print(f"Pobieram: {plik}...")
        r = requests.get(url)
        r.raise_for_status()
        with open(local_name, "wb") as f:
            f.write(r.content)
        print(f"Zapisano: {local_name}")
    except Exception as e:
        print(f"Błąd pobierania {plik}: {e}")

print("✅ Wszystkie dostępne pliki z listy zostały pobrane (jeśli były dostępne).")
