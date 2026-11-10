VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTipoContrattoSoc"
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
    SQL_GETLIST = "SELECT stnidasc, stncodso, stnannul, timestmp FROM multstn0 "
    SQL_FIND = "SELECT stnidatn, stnidasc, stncodso, stnannul, timestmp FROM multstn0 WHERE "
    SQL_UPDATE = "UPDATE multstn0 SET "
    SQL_INSERT = "INSERT INTO multstn0 "
    SQL_GET_TS = "SELECT timestmp FROM multstn0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multstn0.stnidatn", 0, True, 3, "Codice"
        .Add "Società", "multstn0.stnidasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multstn0.stncodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multstn0.stnannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Società", "2"
    End With
    DS.CancelField = "multstn0.stnannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


