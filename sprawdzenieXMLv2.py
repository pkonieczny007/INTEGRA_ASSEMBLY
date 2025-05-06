import os
import xml.etree.ElementTree as ET

def count_command_occurrences(filename, tblref_value):
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        return sum(1 for cmd in root.iter('COMMAND') if cmd.attrib.get('Name') == 'Import' and cmd.attrib.get('TblRef') == tblref_value)
    except Exception as e:
        log(f"Błąd podczas przetwarzania pliku {filename}: {e}")
        return None

def count_empty_commands_z_bazy(filename):
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        all_cmds = [
            cmd for cmd in root.iter('COMMAND')
            if cmd.attrib.get('Name') == 'Import' and cmd.attrib.get('TblRef') == 'PR_SSTT_00000100'
        ]

        empty_cmds = []
        for cmd in all_cmds:
            fields = list(cmd.findall('FIELD'))
            if len(fields) == 3:
                keys = {
                    (f.attrib.get("FldRef"), f.attrib.get("FldValue"), f.attrib.get("FldType"))
                    for f in fields
                }
                if keys == {
                    ("PrdRefOrg", "nan", "20"),
                    ("PrdRefDst", "nan", "20"),
                    ("PQUANT", "nan", "100"),
                }:
                    empty_cmds.append(cmd)

        return len(all_cmds), len(empty_cmds)

    except Exception as e:
        log(f"Błąd podczas analizy pustych elementów w {filename}: {e}")
        return None, None

def check_non_empty_commands_for_quantity(filename):
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        issues = 0

        for cmd in root.iter('COMMAND'):
            if cmd.attrib.get('Name') == 'Import' and cmd.attrib.get('TblRef') == 'PR_SSTT_00000100':
                fields = list(cmd.findall('FIELD'))
                field_set = {
                    (f.attrib.get("FldRef"), f.attrib.get("FldValue"), f.attrib.get("FldType"))
                    for f in fields
                }
                if field_set != {
                    ("PrdRefOrg", "nan", "20"),
                    ("PrdRefDst", "nan", "20"),
                    ("PQUANT", "nan", "100"),
                }:
                    for f in fields:
                        if f.attrib.get("FldRef") == "PQUANT":
                            if f.attrib.get("FldValue") == "nan":
                                issues += 1
        return issues
    except Exception as e:
        log(f"Błąd przy sprawdzaniu ilości w niepustych rekordach: {e}")
        return None

def check_length_field_for_comma(filename):
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        errors = []
        for field in root.iter('FIELD'):
            if field.attrib.get('FldRef') == 'Length':
                fld_value = field.attrib.get('FldValue', '')
                if ',' in fld_value:
                    errors.append(f"Znak ',' w FldValue='{fld_value}'")
        return errors
    except Exception as e:
        log(f"Błąd podczas sprawdzania wartości pola Length w pliku {filename}: {e}")
        return ["Błąd parsowania"]

# Bufor na wynik do zapisu do pliku
output_lines = []

def log(message):
    print(message)
    output_lines.append(message)

def main():
    folder = os.getcwd()
    log(f"📂 Sprawdzanie plików w folderze: {folder}\n")

    # Plik 1: output_zlozenie.xml
    file1 = "output_zlozenie.xml"
    count1 = count_command_occurrences(file1, "PRODUCTS")
    if count1 is not None:
        log(f"{file1}: <COMMAND Name='Import' TblRef='PRODUCTS'> wystąpień: {count1}")

    # Plik 2: output_z_bazy_do_zlozen.xml
    file2 = "output_z_bazy_do_zlozen.xml"
    count2, empty2 = count_empty_commands_z_bazy(file2)
    if count2 is not None:
        non_empty = count2 - empty2
        log(f"{file2}:")
        log(f"  - Wszystkich: {count2}")
        log(f"  - Pustych: {empty2}")
        log(f"  - Niepustych: {non_empty}")
        qty_errors = check_non_empty_commands_for_quantity(file2)
        if qty_errors is not None:
            log(f"  - Nieprawidłowa ilość (PQUANT='nan') w {qty_errors} rekordach ❌")

    # Plik 3: output_profil.xml
    file3 = "output_profil.xml"
    count3 = count_command_occurrences(file3, "PRODUCTS")
    if count3 is not None:
        log(f"{file3}: <COMMAND Name='Import' TblRef='PRODUCTS'> wystąpień: {count3}")
    comma_issues = check_length_field_for_comma(file3)
    if comma_issues:
        log(f"{file3}: Błędy w polach 'Length':")
        for issue in comma_issues:
            log(f"  - {issue}")
    else:
        log(f"{file3}: Pola 'Length' bez błędnych przecinków ✅")

    # Skrócone podsumowanie
    log("\n📌 Podsumowanie:")
    log(f"- {file1} – {count1} elementów")
    log(f"- {file2} – {non_empty} elementów, {empty2} pustych, błędne ilości: {qty_errors}")
    log(f"- {file3} – {count3} elementy {'bez błędnych przecinków' if not comma_issues else 'z błędami przecinków'}")

    # Zapisz wynik do pliku
    with open("sprawdzenieXML.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
        f.write("\n")

if __name__ == "__main__":
    main()
