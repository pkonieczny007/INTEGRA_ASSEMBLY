#!/usr/bin/env python3
import os
import xml.etree.ElementTree as ET

# Bufor na wynik do zapisu do pliku
output_lines = []

def log(message):
    """Dodaje komunikat do bufora i wyświetla go."""
    print(message)
    output_lines.append(message)

def count_command_occurrences(filename, tblref_value):
    """Zlicza komendy <COMMAND Name='Import' TblRef='{tblref_value}'> w pliku XML."""
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        return sum(
            1
            for cmd in root.iter('COMMAND')
            if cmd.attrib.get('Name') == 'Import' and cmd.attrib.get('TblRef') == tblref_value
        )
    except Exception as e:
        log(f"Błąd podczas przetwarzania pliku {filename}: {e}")
        return None

def count_empty_commands_z_bazy(filename):
    """
    Zlicza wszystkie komendy do PR_SSTT_00000100 i te, które mają wszystkie trzy pola z wartością 'nan'.
    Zwraca: (liczba wszystkich, liczba pustych)
    """
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        all_cmds = [
            cmd for cmd in root.iter('COMMAND')
            if cmd.attrib.get('Name') == 'Import' and cmd.attrib.get('TblRef') == 'PR_SSTT_00000100'
        ]
        empty_cmds = [
            cmd for cmd in all_cmds
            if len(cmd.findall('FIELD')) == 3 and all(f.attrib.get('FldValue') == 'nan' for f in cmd.findall('FIELD'))
        ]
        return len(all_cmds), len(empty_cmds)
    except Exception as e:
        log(f"Błąd podczas analizy pustych elementów w {filename}: {e}")
        return None, None

def check_non_empty_commands_for_quantity(filename):
    """
    Zlicza komendy PR_SSTT_00000100, które mają PQUANT='nan' ale PrdRefOrg i PrdRefDst != 'nan'.
    """
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        issues = 0
        for cmd in root.iter('COMMAND'):
            if cmd.attrib.get('Name') == 'Import' and cmd.attrib.get('TblRef') == 'PR_SSTT_00000100':
                fields = {f.attrib.get("FldRef"): f.attrib.get("FldValue") for f in cmd.findall('FIELD')}
                if (
                    fields.get("PQUANT") == "nan"
                    and fields.get("PrdRefOrg") not in (None, "nan")
                    and fields.get("PrdRefDst") not in (None, "nan")
                ):
                    issues += 1
        return issues
    except Exception as e:
        log(f"Błąd przy sprawdzaniu ilości w niepustych rekordach w {filename}: {e}")
        return None

def check_length_field_for_comma(filename):
    """
    Zwraca listę wartości FldValue z przecinkiem dla pola Length.
    """
    try:
        tree = ET.parse(filename)
        root = tree.getroot()
        return [
            f.attrib.get('FldValue')
            for f in root.iter('FIELD')
            if f.attrib.get('FldRef') == 'Length' and ',' in f.attrib.get('FldValue', '')
        ]
    except Exception as e:
        log(f"Błąd podczas sprawdzania pola Length w {filename}: {e}")
        return ["Błąd parsowania"]

def analyze_xml_file(filename, check_combined=False):
    """
    Analizuje podany plik XML:
    - dla PRODUCTS: zlicza Import do PRODUCTS
    - dla PR_SSTT_00000100: zlicza wszystkie i puste
    - sprawdza błędy PQUANT w niepustych
    - sprawdza przecinki w polu Length
    Jeśli check_combined=True, wykonuje wszystkie powyższe testy w jednym.
    """
    log(f"\n🔍 Analiza pliku: {filename}")
    if not os.path.exists(filename):
        log(f"❌ Plik {filename} nie istnieje.")
        return

    # PRODUCTS
    cnt_prod = count_command_occurrences(filename, "PRODUCTS")
    if cnt_prod is not None:
        log(f"- Import do PRODUCTS: {cnt_prod}")

    # PR_SSTT_00000100
    cnt_all, cnt_empty = count_empty_commands_z_bazy(filename)
    if cnt_all is not None:
        log(f"- Komendy do PR_SSTT_00000100: {cnt_all} (pustych: {cnt_empty})")
        cnt_qty_issues = check_non_empty_commands_for_quantity(filename)
        if cnt_qty_issues is not None:
            log(f"- Błędy ilości (PQUANT='nan'): {cnt_qty_issues}")

    # Length
    cnt_commas = check_length_field_for_comma(filename)
    # sprawdzamy Length tylko dla plików profil lub combined
    if filename.endswith("profil.xml") or check_combined:
        if cnt_commas:
            log(f"- Błędy przecinka w polu Length: {cnt_commas}")
        else:
            log("- Pola 'Length' bez błędnych przecinków ✅")

def main():
    folder = os.getcwd()
    log(f"📂 Sprawdzanie plików w folderze: {folder}")

    xml_files = [
        ("output_zlozenie.xml", False),
        ("output_z_bazy_do_zlozen.xml", False),
        ("output_profil.xml", False),
        ("output_combined.xml", True),  # combined -> wszystkie sprawdzenia
    ]

    for fname, is_combined in xml_files:
        analyze_xml_file(fname, check_combined=is_combined)

    log("\n📌 Podsumowanie:")
    for fname, _ in xml_files:
        log(f"- {fname}")

    report_name = "sprawdzenieXML_combined.txt"
    with open(report_name, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
    log(f"\n✅ Raport zapisano jako: {report_name}")

if __name__ == "__main__":
    main()

