VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMotivoEventoSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public DS As CDataStruct

Public SQL_GET_TS As String
Public SQL_INSERT As String
Public SQL_UPDATE As String
Public SQL_FIND As String
Public SQL_GETLIST As String

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub

Private Sub Class_Initialize()
    SQL_GETLIST = "SELECT smeidasc, smecodso, smeannul, timestmp FROM multsme0 "
    SQL_FIND = "SELECT smeidame, smeidasc, smecodso, smeannul, timestmp FROM multsme0 WHERE "
    SQL_UPDATE = "UPDATE multsme0 SET "
    SQL_INSERT = "INSERT INTO multsme0 "
    SQL_GET_TS = "SELECT timestmp FROM multsme0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multsme0.smeidame", 0, True, 3, "Codice"
        .Add "Società", "multsme0.smeidasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multsme0.smecodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multsme0.smeannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Società", "2"
    End With
    DS.CancelField = "multsme0.smeannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub




