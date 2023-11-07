import chardet

def detect_file_encoding(file_path):
    with open(file_path, 'rb') as file:
        result = chardet.detect(file.read())
    return result['encoding']

file_path = 'mapping.json'  # Upewnij się, że ścieżka do pliku jest poprawna
encoding = detect_file_encoding(file_path)

print(f'Kodowanie: {encoding}')
