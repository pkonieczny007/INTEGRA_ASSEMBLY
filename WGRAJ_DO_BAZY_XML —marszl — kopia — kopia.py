import os
import subprocess

def run_powersync(xml_file):
    """
    Uruchamia XmlImporter.exe (powersync) z zadanymi parametrami.
    
    Parametry:
      xml_file: Pełna ścieżka do pliku XML do importu.
    """
    # Ścieżka do pliku XmlImporter.exe
    executable = r"C:\Lantek\System\Common\XmlImporter.exe"
    
    # Przygotowujemy listę argumentów:
    # -dff XMLImporter -src <xml_file> -user "SYSADM" -pass "" -hide 0
    command = [
        executable,
        "-dff", "XMLImporter",
        "-src", xml_file,
        "-user", "SYSADM",
        "-pass", "",
        "-hide", "0"
    ]
    
    try:
        subprocess.run(command, check=True)
        print("Powersync uruchomiony pomyślnie.")
    except subprocess.CalledProcessError as e:
        print(f"Błąd podczas uruchamiania powersync: {e}")
    except FileNotFoundError:
        print("Nie znaleziono XmlImporter.exe. Sprawdź podaną ścieżkę.")

if __name__ == "__main__":
    # Ustal katalog, w którym znajduje się skrypt
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Domyślnie wybieramy plik output_zlozenie.xml z katalogu skryptu
    xml_file = os.path.join(script_dir, "marszruty.xml")
    
    if not os.path.exists(xml_file):
        print(f"Plik XML '{xml_file}' nie został znaleziony.")
    else:
        run_powersync(xml_file)
