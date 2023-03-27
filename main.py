import pandas as pd
import glob
import openpyxl

Pfad = 'C:\\Users\\Daniel\\Desktop\\Progammieren\\Python\\Reading_Material_QA32\\input_files\\'
Pfad_Inhalt = []

XLSX_dateien = glob.glob(Pfad + "/*.XLSX")

for datei in XLSX_dateien:
    datei = datei[len(Pfad):]
    Pfad_Inhalt.append(datei)

Pfad_mit_Dateiname = Pfad + Pfad_Inhalt[0]

print(Pfad_mit_Dateiname)

excel_verarbeitet = pd.DataFrame()

excelinput = pd.read_excel(Pfad_mit_Dateiname)
excelinput = excelinput.iloc[:,3]

print(excelinput[0])
print(len(excelinput))

Material = []


for zeilen in excelinput:
    Material.append(zeilen)

Material_unique = sorted(list(set(Material)))

for DateiName in range(1, len(Pfad_Inhalt)):

    temp_Material = []
    temp_excelinput = pd.read_excel(Pfad + Pfad_Inhalt[DateiName])
    temp_excelinput = temp_excelinput.iloc[:,3]

    for zeilen in temp_excelinput:
        temp_Material.append(zeilen)

    for element in temp_Material:
        if element not in Material_unique:
            Material_unique.append(element)

sorted_Material_unique = sorted(Material_unique, reverse=False)

with open('C:\\Users\\Daniel\\Desktop\\Progammieren\\Python\\Reading_Material_QA32\\output_file\\Material_unique.txt', 'w') as file: # Öffnen der Textdatei im Schreibmodus
    for element in sorted_Material_unique:
        file.write(str(element) + '\n') # Schreiben jedes Elements in eine separate Zeile



print("xxx")