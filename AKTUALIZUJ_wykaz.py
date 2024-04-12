import os
import shutil

# Określenie ścieżki do folderu nadrzędnego i nazwy pliku
parent_folder = os.path.dirname(os.getcwd())
file_name = 'wykaz.xlsx'
source_file_path = os.path.join(parent_folder, file_name)

# Sprawdzenie, czy plik istnieje w folderze nadrzędnym
if os.path.exists(source_file_path):
    # Ścieżka docelowa, gdzie plik zostanie skopiowany
    destination_file_path = os.path.join(os.getcwd(), file_name)
    
    # Kopiowanie pliku
    shutil.copy2(source_file_path, destination_file_path)
    print(f"Plik '{file_name}' został skopiowany do bieżącego folderu.")
else:
    print(f"Plik '{file_name}' nie istnieje w folderze nadrzędnym.")
