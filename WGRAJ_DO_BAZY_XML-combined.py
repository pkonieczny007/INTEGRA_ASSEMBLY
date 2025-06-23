import os
import subprocess

def run_powersync(importer_path, xml_file):
    """
    Uruchamia XMLImporter.exe (powersync) z zadanymi parametrami.
    
    Parametry:
      importer_path: Pełna ścieżka do pliku XMLImporter.exe (np. "C:\\Lantek\\System\\Common\\XMLImporter.exe")
      xml_file: Pełna ścieżka do pliku XML do importu
    """
    # Używamy bezpośrednio podanej ścieżki do pliku wykonywalnego
    executable = importer_path
    
    # Przygotowujemy listę argumentów (analogicznie do VBA):
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
        print("Nie znaleziono XMLImporter.exe. Sprawdź podaną ścieżkę.")

if __name__ == "__main__":
    # Ustal katalog, w którym znajduje się skrypt
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Domyślnie wybieramy plik output_combined.xml z katalogu skryptu
    xml_file = os.path.join(script_dir, "output_combined.xml")
    
    # Podaj pełną ścieżkę do pliku XMLImporter.exe
    importer_path = r"C:\Lantek\System\Common\XMLImporter.exe"
    
    if not os.path.exists(xml_file):
        print(f"Plik XML '{xml_file}' nie został znaleziony.")
    else:
        run_powersync(importer_path, xml_file)
