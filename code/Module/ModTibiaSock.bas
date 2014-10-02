Attribute VB_Name = "ModTibiaSock"
 Public DllVersion As Boolean
 Option Explicit


#If DllVersion Then
 'Api TibiaSocket para Visual Basic 6 Modificado por AvalonTM (21/09/14)
 Public Declare Sub Address_OUTGOINGDATASTREAM Lib "tibiasock_old.dll" (ByVal address As Long)
 Public Declare Sub Address_OUTGOINGDATALEN Lib "tibiasock_old.dll" (ByVal address As Long)
 Public Declare Sub Address_SENDOUTGOINGPACKET Lib "tibiasock_old.dll" (ByVal address As Long)
 Public Declare Sub Address_INCOMINGDATASTREAM Lib "tibiasock_old.dll" (ByVal address As Long)
 Public Declare Sub Address_PARSERFUNC Lib "tibiasock_old.dll" (ByVal address As Long)
 
 
 Public Declare Sub SendPacketToServer Lib "tibiasock_old.dll" (ByVal Handle As Long, ByRef Bytes As Byte, ByVal lenght As Integer)
 Public Declare Sub SendPacketToClient Lib "tibiasock_old.dll" (ByVal Handle As Long, ByRef Bytes As Byte, ByVal lenght As Integer)
 
#Else
 'Api TibiaSocket para Visual Basic 6 Modificado por AvalonTM (21/09/14)
 Public Declare Sub Address_OUTGOINGDATASTREAM Lib "tibiasock_new.dll" (ByVal address As Long)
 Public Declare Sub Address_OUTGOINGDATALEN Lib "tibiasock_new.dll" (ByVal address As Long)
 Public Declare Sub Address_SENDOUTGOINGPACKET Lib "tibiasock_new.dll" (ByVal address As Long)
 Public Declare Sub Address_INCOMINGDATASTREAM Lib "tibiasock_new.dll" (ByVal address As Long)
 Public Declare Sub Address_PARSERFUNC Lib "tibiasock_new.dll" (ByVal address As Long)
 

 Public Declare Sub SendPacketToServer Lib "tibiasock_new.dll" (ByVal Handle As Long, ByRef Bytes As Byte, ByVal lenght As Integer)
 Public Declare Sub SendPacketToClient Lib "tibiasock_new.dll" (ByVal Handle As Long, ByRef Bytes As Byte, ByVal lenght As Integer)

#End If



 Public Const Proseso = "tibiaclient" 'Nombre de la Classe del proceso. Tibia = tibiaclient
 Public TibiaHwnd As Long
