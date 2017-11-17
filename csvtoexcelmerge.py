import csv
import xlsxwriter
import glob

files = glob.glob('*.csv')
workbook = xlsxwriter.Workbook("Redes2.xlsx")

worksheet = workbook.add_worksheet('3 andar')
linha = 2
coluna = 1
flag = 0

merge = []

for file in files:
	with open (file, 'r') as csvfile:
		reader = csv.reader(csvfile)
		

		linha = 2
		coluna = 1
		
		for row in reader:
			flag = 0
			for line in merge:
				if (line['BSSID'] == row[1].upper() and line['canal'] == row[4][8:]):
					flag = 1
			if (flag == 0):
				merge.append({'BSSID': row[1].upper(), 'canal': row[4][8:], 'ESSID': row[0]})
			flag = 0
			

for row in merge:
    print(row)
    worksheet.write(linha, coluna, row['BSSID'])
    worksheet.write(linha, coluna+1, row['canal'])
    worksheet.write(linha, coluna+2, row['ESSID'])
    linha += 1
	
worksheet.add_table('B2:D'+str(linha),{'columns':[{'header':'BSSID'},
													{'header':'Canal'},
													{'header':'ESSID'}]})

        

workbook.close()
    

