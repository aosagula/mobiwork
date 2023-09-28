import pymssql
import pymsgbox
import sys
import xlsxwriter
import datetime





ppc = pymsgbox.prompt('Indicar Numero de PPC?', 'PPC', '')
if ( len(sys.argv) < 2):
	print (" uso: %s <Directorio del archivo> ", sys.argv[0])
	sys.exit()
	
if( len(ppc.strip()) == 0):
	pymsgbox.alert('DEBE INDICAR UN PPC')
	sys.exit()
	
if (  ppc.isdigit() == False ) :
	pymsgbox.alert("EL PPC a PROCESAR >>>>> %s <<<<< DEBE SER UN MUMERO" % ( ppc ))
	sys.exit()



print ("EL ARCHIVO SE GUARDARA EN : %s" % ( str(sys.argv[1])))
path = str( sys.argv[1])

salida = '%s\\%s.xlsx' % ( path, ppc)

workbook = xlsxwriter.Workbook(salida)
worksheet = workbook.add_worksheet()

server = '192.168.0.201'
db = 'VKM_Prod'
user = 'vaclog'
password  = 'hola$$123'
try:
	archivo = open( salida, 'w')



	##cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+db+';UID='+username+';PWD='+ password)
	conn = pymssql.connect(server, user, password, db)		
	print ('CONECTADO')
	cursor = conn.cursor(as_dict=True)
	cursor.execute( 

	"select S1.SEntDocNro as 'Armado', EntID as 'Entidad', EntEntIdC as 'ClienteDestino', \
		EntNombre as 'Nombre', \
		CONVERT(VARCHAR(30), SEntFecAlt, 103) as 'FechaAlta', \
		CONVERT(VARCHAR(30), S.SentFecFin, 103) as 'FechaFin', \
		S.SentPCCId as 'PCC', 		\
		CASE WHEN PCCFec IS NULL THEN 'Pendiente de Entrega'\
		ELSE CONVERT(VARCHAR(30), PCCFec, 103) 		END AS 'FechaPCC',\
		S.SentNroDsp as 'Delivery', 		S.LogEntId as 'ID.Cliente', \
		S.SEntEstEnt as 'Estado', 		S.SentObs as 'Observacion',\
		SUM(SEntIteCnt) as 'Pedidas', \
		SUM(SEntIteCon) as 'Preparadas'\
		from SENT1 (nolock) S1 \
		inner join SENT S (nolock) on S1.SEntDocNro = S.SEntDocNro \
		left join ENT (nolock) on SEntEntId = EntID \
		left join LOPCC on SentPCCId = PCCId \
		where SentTVar = 'S' and SentTipUso = 'S' and  SEntEstEnt <> 'ANU' \
		AND SEntFecAlt >= DATEADD(Day, -35, GETDATE()) \
		AND S.SEntEstEnt = 'CAR' \
		AND S.SentPCCId = %s \
		group by S.SentPCCId, s1.SEntDocNro, EntID, EntEntIdC, EntNombre, SEntFecAlt, \
		S.SentFecFin, S.SentPCCId, PCCFec, S.SentNroDsp, S.LogEntId, S.SEntEstEnt, S.SentObs \
		ORDER BY S.SentPCCId, SEntFecAlt ASC", ppc )
	i = 0
	
	worksheet.write(i, 0, 'WorkOrder_Export')
	i = i + 1

	
	
	worksheet.write(i, 0, 'Generado el %s' % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
	i = i + 1
	worksheet.write(i , 0 , 'Versión')
	worksheet.write(i, 1, '10.0.119')

	i = i + 2
	titulo = [

		'Tipo de OT',
		'Armado',
		'Entidad',
		'Cliente de Destino',
		'Cliente',
		'Fecha Alta',
		'Fecha Fin',
		'PCC',
		'Fecha PCC',
		'Delivery',
		'ID Cliente',
		'Estado',
		'Observaciones',
		'Ud.Pedidas',
		'Ud.Preparadas'
	]

	

	
	for col_num, col_data in enumerate(titulo):
			worksheet.write(i, col_num, col_data)
	
	
	for row_data in cursor:
		i = i + 1
		print(['Entrega',
				row_data['Armado'],
				row_data['Entidad'],
				row_data['ClienteDestino'].strip(),
				row_data['Nombre'].strip(),
				row_data['FechaAlta'],
				row_data['FechaFin'],
				row_data['PCC'],
				row_data['FechaPCC'],
				row_data['Delivery'].strip(),
				row_data['ID.Cliente'],
				row_data['Estado'],
				row_data['Observacion'].strip(),
				row_data['Pedidas'],
				row_data['Preparadas']
		])
		line = [
				'Entrega',
				row_data['Armado'],
				row_data['Entidad'],
				row_data['ClienteDestino'].strip(),
				row_data['Nombre'].strip(),
				row_data['FechaAlta'],
				row_data['FechaFin'],
				row_data['PCC'],
				row_data['FechaPCC'],
				row_data['Delivery'].strip(),
				row_data['ID.Cliente'],
				row_data['Estado'],
				row_data['Observacion'].strip(),
				row_data['Pedidas'],
				row_data['Preparadas']
				]
		for col_num, col_data in enumerate(line):
			worksheet.write(i, col_num, col_data)
		

	
			

	workbook.close()
		

	pymsgbox.alert("EL ARCHIVO %s SE HA GENERADO CON EXITO" % ( salida ))
	sys.exit()
except Exception as e:
    print(e)

