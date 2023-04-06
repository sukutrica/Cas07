Zadatak 1

import openpyxl as openpyxl

# ime_prezime=["Veljko Krstic","Nikola Maric","Sava Milunovic"]
# plata=[70000,80000,150000]


l=[line.strip() for line in open("zaposleni.txt")]
z=[[x.split("|")[0],eval(x.split("|")[1])] for x in l]

wb=openpyxl.Workbook()
ws=wb.active

ws['A1'].value="Ime Prezime"
ws.cell(row=1,column=2).value="Plata"

for i in range(2,len(z)+2):
    ws.cell(row=i,column=1).value=z[i-2][0]
    ws.cell(row=i,column=2).value=z[i-2][1]

wb.save(filename='zaposleni.xlsx')

Zadatak 2

import openpyxl as openpyxl

wb=openpyxl.load_workbook('demo1.xlsx') ##Ucitava excel file

print(wb.sheetnames) ##vraca nazive sheetovaadatak 

Zadatak 2-1
Не можете да правите, мењате и отпремате фајлове … Нема довољно меморијског простора. Набавите још меморијског простора или уклоните фајлове са Диска, Google слика или Gmail-а.
import openpyxl as openpyxl

wb=openpyxl.load_workbook('zaposleni.xlsx')
ws=wb.active

# ##Citanje vrednosti iz odredjenih celija- Nacin 1
# ime_prezime=ws['A2'].value ##Uzima vrednost iz celije A2
# plata=ws['B2'].value ##Uzima vrednost iz celije B2

# ##Citanje vrednosti iz odredjenih celija- Nacin 2
# ime_prezime1=ws.cell(row=3,column=1).value
# plata1=ws.cell(row=3,column=2).value

# ##Citanje cele kolone A
# c=ws["A"]

# for i in c:
#     print(i.value)

# ##Citanje celog reda 2

# r=ws[2]

# for i in r:
#     print(i.value)

##Citanje vise celija odjednom
c=ws["A2":"B7"]

for i in c:
    print("Ime prezime:",i[0].value)
    print("Plata:",i[1].value)
    print("="*20)
    


# print("Ime prezime:",ime_prezime) 
# print("Plata:",plata)
# print("Ime prezime:",ime_prezime1) 
# print("Plata:",plata1)

Zadatak 3
Не можете да правите, мењате и отпремате фајлове … Нема довољно меморијског простора. Набавите још меморијског простора или уклоните фајлове са Диска, Google слика или Gmail-а.
import openpyxl as openpyxl

l=[line.strip() for line in open("proizvodi.txt")]
z=[[x.split("|")[0],int(x.split("|")[1]),eval(x.split("|")[2])] for x in l]

wb=openpyxl.Workbook()
ws=wb.active

ws['A1'].value="Naziv"
ws['B1'].value="Kolicina"
ws['C1'].value="Cena"

for i in range(2,len(z)+2):
    ws.cell(row=i,column=1).value=z[i-2][0]
    ws.cell(row=i,column=2).value=z[i-2][1]
    ws.cell(row=i,column=3).value=z[i-2][2]

wb.save(filename='proizvodi.xlsx')

Zadatak 3-1
import openpyxl as openpyxl

wb=openpyxl.load_workbook('proizvodi.xlsx')
ws=wb.active

c=ws["A2":"C6"]

for i in c:
    print("Naziv:",i[0].value)
    print("Kolicina:",i[1].value)
    print("Cena:",i[2].value)
    print("Vrednost:",i[1].value*i[2].value)
    print("="*25)
