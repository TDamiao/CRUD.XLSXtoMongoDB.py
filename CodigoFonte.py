import xlsxwriter
import xlrd
from pymongo import MongoClient


client = MongoClient('localhost', 27017)
db = client['unifeg']
clima = db['clima']

book = xlrd.open_workbook("Amanda.xlsx")


sh = book.sheet_by_index(0)

linhas = int(sh.nrows)

print("\nO FUTURO CLIMÁTICO - GUAXUPE/MONTE_SANTO/ARCEBURGO \n")



print ('\nPROCESSANDO TABELA 01...')

for rx in range(linhas):
	if(rx > 0):
		clima_x = {
		"dia": int(sh.cell_value(rowx=rx, colx=0)), 
		"mes": int(sh.cell_value(rowx=rx, colx=1)), 
		"ano": int(sh.cell_value(rowx=rx, colx=2)), 
		"hora": int(sh.cell_value(rowx=rx, colx=3)), 
		"nome": (sh.cell_value(rowx=rx, colx=4)), 
		"cidade": (sh.cell_value(rowx=rx, colx=5)), 
		"estado": (sh.cell_value(rowx=rx, colx=6)), 
		"temperatura": (sh.cell_value(rowx=rx, colx=7)),
 		"umidade": int(sh.cell_value(rowx=rx, colx=8)),
 		"clima": (sh.cell_value(rowx=rx, colx=9)),
 		"chance_de_chuva": int(sh.cell_value(rowx=rx, colx=10))}
		mee_id = clima.insert_one(clima_x)

		
print ('PROCESSAMENTO CONCLUIDO!\n')


book = xlrd.open_workbook("JoaoGabriel.xlsx")


sh = book.sheet_by_index(0)

linhas = int(sh.nrows)


print ('\nPROCESSANDO TABELA 02...')

for rx in range(linhas):
	if(rx > 0):
		clima_x = {
		"dia": int(sh.cell_value(rowx=rx, colx=0)), 
		"mes": int(sh.cell_value(rowx=rx, colx=1)), 
		"ano": int(sh.cell_value(rowx=rx, colx=2)), 
		"hora": int(sh.cell_value(rowx=rx, colx=3)), 
		"nome": (sh.cell_value(rowx=rx, colx=4)), 
		"cidade": (sh.cell_value(rowx=rx, colx=5)), 
		"estado": (sh.cell_value(rowx=rx, colx=6)), 
		"temperatura": (sh.cell_value(rowx=rx, colx=7)),
 		"umidade": int(sh.cell_value(rowx=rx, colx=8)),
 		"clima": (sh.cell_value(rowx=rx, colx=9)),
 		"chance_de_chuva": int(sh.cell_value(rowx=rx, colx=10))}
		mee_id = clima.insert_one(clima_x)

		
print ('PROCESSAMENTO CONCLUIDO!\n')


book = xlrd.open_workbook("Thiago.xlsx")


sh = book.sheet_by_index(0)

linhas = int(sh.nrows)



print ('\nPROCESSANDO TABELA 03...')

for rx in range(linhas):
	if(rx > 0):
		clima_x = {
		"dia": int(sh.cell_value(rowx=rx, colx=0)), 
		"mes": int(sh.cell_value(rowx=rx, colx=1)), 
		"ano": int(sh.cell_value(rowx=rx, colx=2)), 
		"hora": int(sh.cell_value(rowx=rx, colx=3)), 
		"nome": (sh.cell_value(rowx=rx, colx=4)), 
		"cidade": (sh.cell_value(rowx=rx, colx=5)), 
		"estado": (sh.cell_value(rowx=rx, colx=6)), 
		"temperatura": (sh.cell_value(rowx=rx, colx=7)),
 		"umidade": int(sh.cell_value(rowx=rx, colx=8)),
 		"clima": (sh.cell_value(rowx=rx, colx=9)),
 		"chance_de_chuva": int(sh.cell_value(rowx=rx, colx=10))}
		mee_id = clima.insert_one(clima_x)

		
print ('PROCESSAMENTO CONCLUIDO!\n')



workbook = xlsxwriter.Workbook('ClimaRegiao.xlsx')
worksheet = workbook.add_worksheet()


row = 0
col = 0

worksheet.write(row,col, "Dia")
worksheet.write(row,col+1, "Mês")
worksheet.write(row,col+2, "Ano")
worksheet.write(row,col+3, "Hora")
worksheet.write(row,col+4, "Nome")
worksheet.write(row,col+5, "Cidade")
worksheet.write(row,col+6, "Estado")
worksheet.write(row,col+7, "Temperatura")
worksheet.write(row,col+8, "Umidade")
worksheet.write(row,col+9, "Clima")
worksheet.write(row,col+10, "Chance de Chuva")
	

for clima in clima.find().sort([('dia', 1), ('hora', 1)]):
    row +=1
    worksheet.write(row,col, clima['dia'])
    worksheet.write(row,col+1, clima['mes'])
    worksheet.write(row,col+2, clima['ano'])
    worksheet.write(row,col+3, clima['hora'])
    worksheet.write(row,col+4, clima['nome'])
    worksheet.write(row,col+5, clima['cidade'])
    worksheet.write(row,col+6, clima['estado'])
    worksheet.write(row,col+7, clima['temperatura'])
    worksheet.write(row,col+8, clima['umidade'])
    worksheet.write(row,col+9, clima['clima'])
    worksheet.write(row,col+10, clima['chance_de_chuva'])


print ('DADOS INSERIDOS NO BANCO DE DADOS!')	
print ('NOVA TABELA ClimaRegiao.xlsx CRIADA!\n')
	
workbook.close()