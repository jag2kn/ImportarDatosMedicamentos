#!/usr/bin/python

import MySQLdb
import sys
from openpyxl import load_workbook
from unidecode import unidecode


if len(sys.argv)<3:
	print "Error: faltan parametros"
	print "Uso: python saludar.py Listado.xlsx NombreTabla"
	exit(0)


db = MySQLdb.connect(host="localhost", # your host, usually localhost
                     user="root", # your username
                      passwd="123456", # your password
                      db="InOperaciones") # name of the data base


#print "Hola "+sys.argv[1]




wb = load_workbook(sys.argv[1], use_iterators = True)

print wb.get_sheet_names()
hojas = wb.get_sheet_names()


for hoja in hojas:
	print hoja
	ws = wb.get_sheet_by_name(name = hoja)

	highest_column = ws.get_highest_column()
	print highest_column

	highest_row = ws.get_highest_row()
	print highest_row


	print ' * Creando tabla'+sys.argv[2]
	db.query('DROP TABLE IF EXISTS '+sys.argv[2]+';')
	
	
	#for i in range(0, highest_column):
	celdas = ws.get_cells(7, 1, 7, highest_column)
	celdas5 = ws.get_cells(7, 1, 7, highest_column)
	
	listaNombres = []
	#if primera:
	creacion = 'CREATE TABLE IF NOT EXISTS `'+sys.argv[2]+'` ( \
		`id` int(11) NOT NULL AUTO_INCREMENT PRIMARY KEY '

	contador=0
	for x in celdas:
		valor = unidecode(x.value)
		valor = valor.replace(" ", "_")
		valor = valor.replace(".", "")
		print valor
		
		if contador>=7 and contador<=10:
			valor="CUM_"+valor
		
		creacion = creacion + ',\
		`'+valor+'` varchar(256) NOT NULL'
		
		listaNombres.append(valor)

		contador=contador+1

	creacion = creacion + '\
		) ENGINE=InnoDB DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ; '


	print ' * Creando tabla'+sys.argv[2]+' DONE'
	print "Debug: "
	print creacion
	print listaNombres

	db.query(creacion)


	print ' * Creando registros'
	cuenta = 100
	for j in range(8, highest_row):
		celdas = ws.get_cells(j, 1, j, highest_column)
		insert = "INSERT INTO "+sys.argv[2]+"` (`id`, `"+"`,`".join(listaNombres)+"`) VALUES (NULL"
		for x in celdas:
			#print type(x.value)
			if type(x.value)==unicode:
				valor = x.value
				valor = valor.encode('ascii','ignore')
				valor = valor.replace("'", "")

			else:
				valor = str(x.value)
			insert = insert + ", '"+valor+"'"
		insert = insert + ");"

		#print "\n\n"
		#print insert

		db.execute(insert)

		if (j-8) % cuenta == 0:
			print " * Insertados "+str(j-8)+" registros de "+str(highest_row-8)
			print insert
















	
