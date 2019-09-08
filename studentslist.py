#On import la library openPyXL
import openpyxl

#Ouvrir le fichier excel
excelDocument = openpyxl.load_workbook('Eleves.xlsx')
print(type(excelDocument))
#Acceder a une feuille
sheet = excelDocument.worksheets[0]
cell = sheet['B2'].value
print(type(cell))

#Ici nous allons creer des classeur
#wb = openpyxl.Workbook()

#sheet = wb.active
#sheet.title = 'data'
#wb.save('Moyenne.xlsx')
