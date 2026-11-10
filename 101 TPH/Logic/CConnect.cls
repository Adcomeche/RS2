VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjDA As CDataAccess

Public Property Get ConnDB() As String
    ConnDB = mobjDA.ConnDB
End Property

Public Property Get ConnPWD() As String
    ConnPWD = mobjDA.ConnPWD
End Property

Public Property Get SQLDecSeparator() As String
    SQLDecSeparator = mobjDA.SQLDecSeparator
End Property

Public Property Get SQLJollyChar() As String
    SQLJollyChar = mobjDA.SQLJollyChar
End Property

Public Function MakeSQL(Data, DataType As Byte) As String
    MakeSQL = mobjDA.MakeSQL(Data, DataType)
End Function

Public Property Get ConnServer() As String
    ConnServer = mobjDA.ConnServer
End Property

Public Property Get ConnUserName() As String
    ConnUserName = mobjDA.ConnUserName
End Property

Public Property Get DA() As CDataAccess
    Set DA = mobjDA
End Property


Public Function Connect(ByVal UserName As String, _
                        ByVal Password As String, _
                        ByVal DriverName As String, _
                        ByVal ServerName As String, _
                        ByVal DBName As String) As Boolean
    Connect = mobjDA.Connect(UserName, Password, DriverName, ServerName, DBName)
End Function

Public Property Set ErrH(objNewValue As CErrorHandler)
    Set mobjDA.ErrH = objNewValue
End Property

Private Sub Class_Initialize()
    Set mobjDA = New CDataAccess
End Sub


Private Sub Class_Terminate()
    Set mobjDA = Nothing
End Sub


Public Property Let DataEngine(varNewValue As DBEngine)
    mobjDA.DataEngine = varNewValue
End Property

Public Property Get DataEngine() As DBEngine
    DataEngine = mobjDA.DataEngine
End Property

