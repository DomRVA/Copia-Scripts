import smb
from smb.SMBConnection import SMBConnection
import socket
import pyodbc
import pandas as pd
#from pandas import DataFrame
import datetime
from datetime import datetime
# import os

share = "hqshare01"
shareUID = 'orchadmin'
sharePWD ='%TRE5tyu'
shareDomain = 'HDL'
sharePath = 'hqshare01/lab share/Incomplete Orders'
IP = socket.gethostbyname(share)
print(IP)
conn = SMBConnection(shareUID,sharePWD,shareUID,sharePath, domain=shareDomain, use_ntlm_v2=True, is_direct_tcp=True)
if conn.connect(IP, 445):
    print('Conncection to {} Established'.format(sharePath))
else:
    print('Connection to {} could not be established'.format(sharePath))
# outpath = sharePath
outpath = '/run/user/1000/gvfs/smb-share:server=hqshare01,share=lab%20share/Incomplete Orders'
# if os.path.isdir(outpath):
# 	print("found the Directory")
# else:
# 	print("Can't find that directory")


filename = datetime.now().strftime("%Y-%m-%d")
server = 'HIP-COPIADBVM01'
database = 'copia'
username = 'copia'
password = 'copia'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

# cursor = cnxn.cursor()
query="""SELECT DISTINCT copia.Requisition.requisitionNumber as [Order ID], copia.SubSpecimen.ID as [SID], copia.Location.name as [Ordering Location], 
	CONVERT(varchar(20),dateadd(hour,-4,dateadd(s,copia.Requisition.receivedStamp/1000, '01/01/1970 00:00:00')),121) as [Delivery Date],	
	Replace(Replace(Replace(Replace(copia.Requisition.requisitionStatus, 15, 'Partial Results'), 10, 'No Results'), 5, 'Partially Collected'), 0, 'Not Collected') as [Order Status]
FROM copia.Requisition
	Left Join copia.Specimen on copia.Specimen.requisitionKey = copia.Requisition.requisitionKey
	left join copia.SubSpecimen on copia.Specimen.specimenKey = copia.SubSpecimen.specimenKey
	JOIN copia.Location on copia.Requisition.orderingLocationKey = copia.Location.locationKey
WHERE copia.Requisition.requisitionStatus != 20 --not in ('20', '999')
	AND copia.Requisition.canceledStamp < 0
	AND DATEDIFF(s, '1970-01-01 00:00:00', GETUTCDATE()) > (copia.Requisition.receivedStamp + 345600000)/1000 --older than 4 days worth of miliseconds
	AND DATEDIFF(s, '1970-01-01 00:00:00', GETUTCDATE()) < (copia.Requisition.receivedStamp + 2592000000)/1000 --newer than 30 days worth of miliseconds
	AND copia.SubSpecimen.ID not like 'iop%'
ORDER BY [SID]"""
# cursor.execute(query)
# data = cursor.fetchall()

data = pd.read_sql(query,cnxn,columns=list)
#Data = pd.DataFrame(data)
# print(Data)

export_excel = data.to_excel (outpath+'/{}_oldorders.xlsx'.format(filename),index=False)
# print(data)

# conn.close()