import csv
import xlsxwriter
import glob

files = glob.glob('*.csv')
workbook = xlsxwriter.Workbook("Redes2.xlsx")

for file in files:
    with open (file, 'rb') as csvfile:
        reader = csv.reader(csvfile)
        worksheet = workbook.add_worksheet('Dados'+file[0:3])

        linha = 2
        coluna = 1

        

        for row in reader:
            worksheet.write(linha, coluna, row[1].upper())
            worksheet.write(linha, coluna+1, row[4][8:])
            worksheet.write(linha, coluna+2, row[0])
            linha += 1
        worksheet.add_table('B2:D'+str(linha-1),{'columns':[{'header':'BSSID'},
														   {'header':'Canal'},
														   {'header':'ESSID'}]})

        

workbook.close()
    
