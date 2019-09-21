import docx2txt
from datetime import datetime
import re
from pathlib import Path
import csv
import xlwt 
from xlwt import Workbook 
import sys
from tkinter import filedialog


dirname = filedialog.askdirectory(initialdir="./",  title='Ordner mit Rechnungen auswählen')
texts = []
pathlist = Path(dirname).glob('**/*.docx')

for path in pathlist:
     # because path is object not string
     path_in_str = str(path)
     result = docx2txt.process(path_in_str)
     texts.append(result)
dates = []
amounts = []
taxes = []
comparableDate = []



#alle .docx Dateien durchsuchen
for t in texts:
    #Datum Finden
    match = re.search(r'geliefert am[:]* \d{2}.\d{2}.\d{4}', t)
    date = []
    date = datetime.strptime(match.group().split(" ")[-1], '%d.%m.%Y').date().strftime('%d.%m.%Y')
    dates.append(date)
    #Datum zum Vergleichen im Format MMDD speichern
    comparableDate.append(date.split(".")[1]+date.split(".")[0])

    #Rechnungsbeträge finden. Vorletzes Match sind die Steuern, das Match davor der Betrag
    match =re.findall('[0-9]*[.]*[0-9]*,[0-9]*',t)
    amounts.append(match[-3])
    taxes.append(match[-2])

#Anhand des Datums sortieren
sortedDates = [x for _,x in sorted(zip(comparableDate,dates))]
sortedAmounts = [x for _,x in sorted(zip(comparableDate,amounts))]
sortedTaxes = [x for _,x in sorted(zip(comparableDate,taxes))]
dateDate=comparableDate.sort()


# Workbook erstellen
wb = Workbook() 
  
# add_sheet: Seite erstellen

#Kopfzeilen für die Quartale erstellen
sheet1 = wb.add_sheet('Sheet 1') 

for i in range(4):
    j = 4*i
    sheet1.write(0,j,"Datum")
    sheet1.write(0,j+1,"Betrag")
    sheet1.write(0,j+2,"Steuern")

#Aktuelle Zeile der Quartale
linesFirst = 1
linesSecond = 1
linesThird = 1
linesFourth = 1

#Werte in die Quartale einordnen
for index,d in enumerate(sortedDates):
    if(int(comparableDate[index])<=331): #Januar-März
        sheet1.write(linesFirst,0,d)
        sheet1.write(linesFirst,1,sortedAmounts[index])
        sheet1.write(linesFirst,2,sortedTaxes[index])
        linesFirst +=1
    elif(int(comparableDate[index])<=630): #April-Juni
        sheet1.write(linesSecond,4,d)
        sheet1.write(linesSecond,5,sortedAmounts[index])
        sheet1.write(linesSecond,6,sortedTaxes[index])
        linesSecond += 1
    elif(int(comparableDate[index]) <= 930): #Juli-September
        sheet1.write(linesThird,8,d)
        sheet1.write(linesThird,9,sortedAmounts[index])
        sheet1.write(linesThird,10,sortedTaxes[index])
        linesThird +=1
    else: #Oktober-Dezember
        sheet1.write(linesFourth,12,d)
        sheet1.write(linesFourth,13,sortedAmounts[index])
        sheet1.write(linesFourth,14,sortedTaxes[index])
        linesFourth +=1

wb.save('Abrechnung.xls') 
print("Fertig! Das Ergebnis wurde in 'Abrechnug.xls' gespeichert")
input("Drücke einen beliebigen Knopf zum Beenden .")