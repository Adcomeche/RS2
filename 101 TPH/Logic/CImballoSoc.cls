VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImballoSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public DS As CDataStruct






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
    SQL_GETLIST = "SELECT simidasc, simcodso, simannul, timestmp FROM multsim0 "
    SQL_FIND = "SELECT simidaim, simidasc, simcodso, simannul, timestmp FROM multsim0 WHERE "
    SQL_UPDATE = "UPDATE multsim0 SET "
    SQL_INSERT = "INSERT INTO multsim0 "
    SQL_GET_TS = "SELECT timestmp FROM multsim0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multsim0.simidaim", 0, True, 3, "Codice"
        .Add "Società", "multsim0.simidasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multsim0.simcodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multsim0.simannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Società", "2"
    End With
    DS.CancelField = "multsim0.simannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


