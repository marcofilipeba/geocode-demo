# coding: latin-1
from xlrd import open_workbook
import locale
import sqlite3
import argparse
import csv
import geocoder

locale.setlocale(locale.LC_ALL,'portuguese')

#API Google  -> AIzaSyBpfZq_1seOTXcsJ-U7y7C0H9hdjaH1ZtA <-

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

#url = "http://maps.google.com/maps/api/geocode/json?address=avenida%C2%A0da%C2%A0boavista,%C2%A01681+4100-132+porto&sensor=false&key=AIzaSyBpfZq_1seOTXcsJ-U7y7C0H9hdjaH1ZtA"
#data = json.load(urllib2.urlopen(url))
#print data
#print data[u'results']

#print "---"

#print data[u'results'][0][u'geometry']

#print "---"

#print data[u'results'][0][u'geometry'][u'location']

#print data[u'results'][0][u'geometry'][u'location'][u'lat']
#print data[u'results'][0][u'geometry'][u'location'][u'lng']



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

		
		#txt = remove_non_ascii(s.cell(row,3).value).replace(" ","%C2%A0")+"+"+remove_non_ascii(s.cell(row,5).value).replace(" ","%C2%A0")+"+"+remove_non_ascii(s.cell(row,6).value).replace(" ","%C2%A0")
		
		print s.cell(row,3), s.cell(row,5), s.cell(row,6)

		#txt = remove_non_ascii(s.cell(row,3).value).replace(" "," ")+" ,"+remove_non_ascii(s.cell(row,5).value).replace(" "," ")+" ,"+remove_non_ascii(s.cell(row,6).value).replace(" "," ")

		#txt = s.cell(row,3).value.encode('utf-8') +" ,"+ s.cell(row,5).value.encode('utf-8') +" ,"+ s.cell(row,6).value.encode('utf-8')

		txt = s.cell(row,3).value +", "+ s.cell(row,5).value +", "+ s.cell(row,6).value+", PORTUGAL"

		values.append(txt)



# import geocoder
# g = geocoder.bing('rua Costa cabral 811, 4200, porto', key = 'Ar-sW38yqFrxA6D7OOMIvxLuD_Zx9WKOg33NJQ3NLl4pdFXy_S2PlR3jFqab8MWp')
# g.json
# {'status': 'OK', 'city': u'Porto', 'confidence': 9, 'neighborhood': u'Paranhos', 'encoding': 'utf-8', 'country': u'Portugal', 'provider': 'bing', 'location': 'rua Costa cabral 811, 4200, porto', 'state': u'Porto', 'street': u'Rua de Costa Cabral 811', 'bbox': {'northeast': [41.1681541, -8.5929867], 'southwest': [41.1659057, -8.5959733]}, 'status_code': 200, 'address': u'Rua de Costa Cabral 811, Porto, Porto 4200-224, Portugal', 'lat': 41.1670299, 'ok': True, 'lng': -8.59448, 'postal': u'4200-224', 'quality': u'Address', 'accuracy': u'Rooftop'}
# g.json['lat']
# 41.1670299
# g.json['lng']
# -8.59448

		# processo antigo com o Google (cuja chave está limitada a poucos calls)		
		#url = "https://maps.google.com/maps/api/geocode/json?address="+txt+"&sensor=false&key=AIzaSyBpfZq_1seOTXcsJ-U7y7C0H9hdjaH1ZtA"
		#data = json.load(urllib2.urlopen(url, timeout = 2))
		
		# obtive a chave em https://www.bingmapsportal.com entrando com a minha conta MS (marco.filipe@segafredo.pt)

		#print txt

		data = geocoder.bing(txt, key = 'Ar-sW38yqFrxA6D7OOMIvxLuD_Zx9WKOg33NJQ3NLl4pdFXy_S2PlR3jFqab8MWp')
		

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
		
		url = "https://maps.google.com/maps/api/geocode/json?address="+row[1]+"&sensor=false&key=AIzaSyBpfZq_1seOTXcsJ-U7y7C0H9hdjaH1ZtA"
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