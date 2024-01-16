import pandas as pd

# Wczytanie danych z pliku Excel
df = pd.read_excel("wykaz.xlsx")

# Tworzenie pliku XML
xml_content = '<DATAEX>\n'
xml_content += '<!-- 3. KOMENDA DO WGRYWANIA ELEMENTÓW Z BAZY DO ZŁOŻENIA -->\n'

for _, row in df.iterrows():
    xml_content += '<COMMAND Name="Import" TblRef="PR_SSTT_00000100">\n'
    xml_content += f'    <FIELD FldRef="PrdRefOrg" FldValue="{row["ZŁOŻENIE"]}" FldType="20"/>\n'
    xml_content += f'    <FIELD FldRef="PrdRefDst" FldValue="{row["REFERENCJA_ELEMENTU"]}" FldType="20"/>\n'
    xml_content += f'    <FIELD FldRef="PQUANT" FldValue="{row["Algorytm dla Anz"]}" FldType="100"/>\n'
    xml_content += '</COMMAND>\n'

xml_content += '</DATAEX>'

# Zapisanie danych do pliku XML
with open("output_z_bazy_do_zlozen.xml", "w", encoding='utf-8') as xml_file:
    xml_file.write(xml_content)

print("Plik XML został utworzony: plik_wyjsciowy -> xml")
