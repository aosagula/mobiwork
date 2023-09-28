import pymssql
import pymsgbox

import sys
import codecs
import datetime






ppc = pymsgbox.prompt('Indicar Numero de PPC?', 'PPC', '')
if ( len(sys.argv) < 2):
    print (" uso: %s <Directorio del archivo> ", sys.argv[0])
    sys.exit()
    
if( len(ppc.strip()) == 0):
    pymsgbox.alert('DEBE INDICAR UN PPC')
    sys.exit()
    
# if (  ppc.isdigit() == False ) :
# 	pymsgbox.alert("EL PPC a PROCESAR >>>>> %s <<<<< DEBE SER UN MUMERO" % ( ppc ))
# 	sys.exit()



print ("EL ARCHIVO SE GUARDARA EN : %s" % ( str(sys.argv[1])))
path = str( sys.argv[1])

# salida = '%s\\%s.csv' % ( path, ppc)






server = '192.168.0.201'
db = 'VKM_Prod'
user = 'vaclog'
password  = 'hola$$123'
try:
    # archivo = codecs.open( salida, 'w', "utf-8")

    print( ppc.split(','))
    listPpc=""
    for i, ppcitem in enumerate(ppc.split(',')):
        if (i != len(ppc.split(','))-1):
            listPpc += '\'' + ppcitem + '\', '
        else:
            listPpc += '\'' + ppcitem + '\''

    print(listPpc)
    ##cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+db+';UID='+username+';PWD='+ password)
    conn = pymssql.connect(server, user, password, db)		
    print ('CONECTADO')
    cursor = conn.cursor(as_dict=True)

    sentence = "select S1.SEntDocNro as 'Armado',ENT.EntID as 'Entidad', EntEntIdC as 'ClienteDestino', \
        EntNombre as 'Nombre', \
        CONVERT(VARCHAR(30), SEntFecAlt, 103) as 'FechaAlta', \
        CONVERT(VARCHAR(30), S.SentFecFin, 103) as 'FechaFin', \
        S.SentPCCId as 'PCC', 		\
        CASE WHEN PCCFec IS NULL THEN 'Pendiente de Entrega'\
        ELSE CONVERT(VARCHAR(30), PCCFec, 103) 		END AS 'FechaPCC',\
        S.SentNroDsp as 'Delivery', 		S.LogEntId as 'ID.Cliente', \
        S.SEntEstEnt as 'Estado', 		S.SentObs as 'Observacion',\
        SUM(SEntIteCnt) as 'Pedidas', \
        SUM(SEntIteCon) as 'Preparadas',\
        E6.EntLogId as cuenta_id \
        from SENT1 (nolock) S1 \
        inner join SENT S (nolock) on S1.SEntDocNro = S.SEntDocNro \
        left join ENT (nolock) on SEntEntId =ENT.EntID \
        left join LOPCC on SentPCCId = PCCId \
        left join ENT6 E6 on E6.EntID = ENT.EntID \
        where SentTVar = 'S' and SentTipUso = 'S' and  SEntEstEnt <> 'ANU' \
        AND SEntFecAlt >= DATEADD(Day, -35, GETDATE()) \
        AND S.SEntEstEnt = 'CAR' \
        AND S.SentPCCId IN ( %s ) \
        group by S.SentPCCId, s1.SEntDocNro,ENT.EntID, EntEntIdC, EntNombre, SEntFecAlt, \
        S.SentFecFin, S.SentPCCId, PCCFec, S.SentNroDsp, S.LogEntId, S.SEntEstEnt, S.SentObs, E6.EntLogId \
        ORDER BY S.SentPCCId, SEntFecAlt ASC" % listPpc 
    cursor.execute(     sentence )
    i = 0
    
    salida = '%s\\AGD_%s_V2.csv' % ( path, ppc.replace(',','_'))
    with codecs.open (salida, 'w', 'ansi') as csv_file:
        
        row = 'C'
        csv_file.write( '%s\n' % row)

        i = i + 1
        date = datetime.datetime.now().strftime("%d/%m/%Y")
        hour = datetime.datetime.now().strftime("%H:%M")

        nro_vuelta = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

        titulo = "%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;\n" %  (
            'command',
            'active',
            'activitiesOrigin',
            'ScheduleType',
            'date',
            'hour',
            'Agent',
            'CF_Armado',
            'CF_Entidad',
            'CF_Cliente_Destino',
            'serviceLocal',
            'CF_Fecha_Alta',
            'CF_Fecha_Fin',
            'CF_PCC',
            'CF_Fecha_PCC',
            'CF_Delivery',
            'CF_ID.Cliente',
            'CF_Estado',
            'CF_Observación',
            'CF_UD._Pedidas',
            'CF_UD._Preparadas',
            'CF_Nro_de_Vuelta'
        )
        csv_file.write(titulo)
        for row_data in cursor:
           
            i = i + 1
            if row_data['cuenta_id'] == 12716:
                entrega = 'Entrega DMD'
            if row_data['cuenta_id'] == 12716:
                entrega = 'Entrega BEEPURE'
            else:
                entrega = 'Entrega.'
            row = "%s;%d;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%s;%d;%d;%s;\n" % ('I',
                    1,
                    7,
                    entrega,
                    date,
                    hour,
                    'adminvaclog1',
                    row_data['Armado'],
                    row_data['Entidad'],
                    row_data['ClienteDestino'].strip(),
                    row_data['ClienteDestino'].strip(),
                    #row_data['Nombre'].strip(),
                    row_data['FechaAlta'],
                    row_data['FechaFin'],
                    row_data['PCC'],
                    row_data['FechaPCC'],
                    row_data['Delivery'].strip(),
                    row_data['ID.Cliente'],
                    row_data['Estado'],
                    row_data['Observacion'].strip(),
                    row_data['Pedidas'],
                    row_data['Preparadas'],
                    nro_vuelta)
            csv_file.write(row)
    csv_file.close()
    pymsgbox.alert("EL ARCHIVO %s SE HA GENERADO CON EXITO" % ( salida ))
    sys.exit()
except Exception as e:
    print(e)

