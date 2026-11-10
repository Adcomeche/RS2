VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDepositoSoc"
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
    SQL_GETLIST = "SELECT sdeidasc, sdecodso, sdeannul, timestmp FROM multsde0 "
    SQL_FIND = "SELECT sdeidade, sdeidasc, sdecodso, sdeannul, timestmp FROM multsde0 WHERE "
    SQL_UPDATE = "UPDATE multsde0 SET "
    SQL_INSERT = "INSERT INTO multsde0 "
    SQL_GET_TS = "SELECT timestmp FROM multsde0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multsde0.sdeidade", 0, True, 3, "Codice"
        .Add "Società", "multsde0.sdeidasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multsde0.sdecodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multsde0.sdeannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Società", "2"
    End With
    DS.CancelField = "multsde0.sdeannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


