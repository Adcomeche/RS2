VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTipoEventoSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    SQL_GETLIST = "SELECT steidasc, stecodso, steannul, timestmp FROM multste0 "
    SQL_FIND = "SELECT steidate, steidasc, stecodso, steannul, timestmp FROM multste0 WHERE "
    SQL_UPDATE = "UPDATE multste0 SET "
    SQL_INSERT = "INSERT INTO multste0 "
    SQL_GET_TS = "SELECT timestmp FROM multste0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multste0.steidate", 0, True, 3, "Codice"
        .Add "Società", "multste0.steidasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multste0.stecodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multste0.steannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Società", "2"
    End With
    DS.CancelField = "multste0.steannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub



