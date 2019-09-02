import smb
from smb.SMBConnection import SMBConnection
import socket
import os

sharePath = 'hqshare01'
shareUID = 'orchadmin'
sharePWD ='%TRE5tyu'
shareDomain = 'HDL'
outPath = '//hqshare01/lab share/Incomplete Orders'
IP = socket.gethostbyname(sharePath)
print(IP)
conn = SMBConnection(shareUID,sharePWD,shareUID,sharePath, domain=shareDomain, use_ntlm_v2=True, is_direct_tcp=True)
conn.connect('hqshare01', 445)
results = conn.listPath('hqshare01.hdl.local', '/')

for x in results:
	print(x.filename)

# if conn.connect(IP, 445):
#     print('Conncection to {} Established'.format(sharePath))
# else:
#     print('Connection to {} could not be established'.format(sharePath))

# os.listdir(r'\lab share')