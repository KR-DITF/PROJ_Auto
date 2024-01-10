import time
import openpyxl
from openpyxl import workbook, load_workbook

wb = load_workbook("saraksts_IMEI_konti.xlsx") #atveram galveno workbook, kurā atrodas visi konta nr. un to statusi
ws = wb.active
ws.title = "konti"
max_r = ws.max_row 

konti = []
for row1 in range (2,max_r+1): #ejam cauri visiem konta nr.
    status = ws["C"+str(row1)].value
    if status == None: #pārbaudam vai kontam ir vai nav statuss, ja nav, tad pievienojam kontu sarakstā
        k_nr = ws["A"+str(row1)].value
        konti.append(k_nr) #pievienojam kontu sarakstam
print(konti)

sheetn = []
sheets = wb.sheetnames
i = 1
for s_name in sheets: #excel dokuments mainās katru nedēļu, tapēc jāatrod cik daudz lapas ir dokumentā
    sheet = wb[s_name]
    sheetn.append(sheet) #pievienojam visu lapu nosaukumus sarakstam, lai tām var iet cauri vienai pēc otras
print(sheetn)
    
wb1 = load_workbook("aktualie.xlsx") #atver otro workbook, no kuras nolasa vai konts atrodas sarakstā un vai bilance <=10
ws1 = wb1.active
ws1.title = "Lapa1"
max_ro = ws1.max_row

for sk in konti: #sāk loop, lai izietu cauri kontu nr., kuri vēl nav apstrādāti
    for row2 in range(2,max_ro+1): #katram kontam iet cauri otrajam workbook
        ch = ws1["A"+str(row2)].value
        bal = ws1["E"+str(row2)].value
    
        



wb.close()
wb1.close()
