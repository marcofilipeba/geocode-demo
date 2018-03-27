# coding: latin-1
from xlrd import open_workbook
import locale
import sqlite3
import argparse
import csv
import geocoder

locale.setlocale(locale.LC_ALL,'portuguese')



# inicio
# gestão de parametros
parser = argparse.ArgumentParser(description='Gera ficheiro com as coordenadas (latitudes e longitudes) dos clientes')
parser.add_argument('fx_clientes',  help='IN -> o nome do ficheiro com os endereços dos clientes (XLSX)')
parser.add_argument('fx_coordenadas',  help='OUT -> o nome do ficheiro com as coordenadas dos clientes')
args = parser.parse_args()
#fxout = args.ficheiro
print args.fx_clientes
print args.fx_coordenadas


#teste do JSON

import json
import urllib2


#inicio do processo

def remove_non_ascii(text):
	return ''.join([i if ord(i) < 128 else ' ' for i in text])

def cria_tabelas(cursor):
	cursor.execute('''create table coordenadas (
		codcli integer,
		url text,
		lat real,
		lng real,
		result text,
		tentativas integer)''')

def loads(book,sheet,tabela,cursor):
	s = book.sheet_by_name(sheet)
	print 'Load:',s.name
	for row in range(1, s.nrows):
		values = []
		values.append(s.cell(row,0).value)

		
	
		print s.cell(row,3), s.cell(row,5), s.cell(row,6)


		txt = s.cell(row,3).value +", "+ s.cell(row,5).value +", "+ s.cell(row,6).value+", PORTUGAL"

		values.append(txt)




		# processo antigo com o Google (cuja chave está limitada a poucos calls)		
		#url = "https://maps.google.com/maps/api/geocode/json?address="+txt+"&sensor=false&key=zzzzzzzzzzzzzzzzzzzzz"
		#data = json.load(urllib2.urlopen(url, timeout = 2))
		
		# obtive a chave em https://www.bingmapsportal.com entrando com a minha conta MS

		#print txt

		data = geocoder.bing(txt, key = 'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz')
		

		#print data.json

		try:
			#print data[u'results'][0][u'geometry'][u'location'][u'lat']
			#print data[u'results'][0][u'geometry'][u'location'][u'lng']
			values.append(data.json['lat'])
			values.append(data.json['lng'])
			values.append('OK')
			values.append(1)
			
		except KeyError:
			print "falhou"
			values.append(0)
			values.append(0)
			values.append('KO')
			values.append(1)
			
		#print s.cell(row,0), s.cell(row,1)
		#for col in range(s.ncols):
		#	values.append(s.cell(row,col).value)
		#campos = '?,'*nrcampos
		cursor.execute('insert into '+tabela+' values (?,?,?,?,?,?)',values)
	db.commit()


def recursivo(cursor, nrmax):
	
	cursor.execute('''select codcli, url, lat, lng, result, tentativas from coordenadas where result = "KO" ''')
	linhas= cursor.fetchall()
	for row in linhas:
		#print row[0],row[1],row[2],row[3],row[4],row[5]
		
		values = []
		values.append(row[0])
		values.append(row[1])
		
		url = "https://maps.google.com/maps/api/geocode/json?address="+row[1]+"&sensor=false&key=zzzzzzzzzzzzzzzzzzzzzz"
		data = json.load(urllib2.urlopen(url, timeout = 2))
		
		try:
			#print data[u'results'][0][u'geometry'][u'location'][u'lat']
			#print data[u'results'][0][u'geometry'][u'location'][u'lng']
			values.append(data[u'results'][0][u'geometry'][u'location'][u'lat'])
			values.append(data[u'results'][0][u'geometry'][u'location'][u'lng'])
			values.append('OK')
			values.append(row[5]+1)
			cursor.execute('delete from coordenadas where codcli = '+str(row[0]))
			cursor.execute('insert into coordenadas values (?,?,?,?,?,?)',values)
			
		except IndexError:
			#print "falhou"
			cursor.execute('update coordenadas set tentativas ='+str(row[5]+1)+' where codcli = '+str(row[0]))


	cursor.execute('select count(1) from coordenadas where result = "KO" and tentativas < '+str(nrmax))
	registos = cursor.fetchall()
	for reg in registos:
		print "conta:"
		print reg[0]
		if reg[0] > 0:
			recursivo(cursor,nrmax)

def gerafich(cursor):
	cursor.execute('''select codcli, url, lat, lng, result, tentativas from coordenadas''')
	linhas= cursor.fetchall()
	with open(args.fx_coordenadas, 'wb') as csvfile:
		escritor = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
		for row in linhas:
			#print row[0], row[1], row[2]
			escritor.writerow(( row[0],row[1].encode('utf-8'),locale.format('%.15f',row[2]),locale.format('%.15f',row[3]),row[4],row[5]))
	print 'Escrita: '+args.fx_coordenadas


# leitura ficheiro
db = sqlite3.connect(':memory:');
cur = db.cursor()
wb = open_workbook(args.fx_clientes)
cria_tabelas(cur)

loads(wb,'clientes','coordenadas',cur)

#recursivo(cur,5)

gerafich(cur)
