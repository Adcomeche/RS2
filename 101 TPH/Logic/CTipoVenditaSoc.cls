VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTipoVenditaSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    SQL_GETLIST = "SELECT stvidasc, stvcodso, stvannul, timestmp FROM multstv0 "
    SQL_FIND = "SELECT stvidatv, stvidasc, stvcodso, stvannul, timestmp FROM multstv0 WHERE "
    SQL_UPDATE = "UPDATE multstv0 SET "
    SQL_INSERT = "INSERT INTO multstv0 "
    SQL_GET_TS = "SELECT timestmp FROM multstv0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multstv0.stvidatv", 0, True, 3, "Codice"
        .Add "Società", "multstv0.stvidasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multstv0.stvcodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multstv0.stvannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Società", "2"
    End With
    DS.CancelField = "multstv0.stvannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


