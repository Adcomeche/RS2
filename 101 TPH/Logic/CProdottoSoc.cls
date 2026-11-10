VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProdottoSoc"
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
    SQL_GETLIST = "SELECT spridasc, spridaim, spridapd, sprcodso, sprannul, timestmp FROM multspr0 "
    SQL_FIND = "SELECT spridapr, spridaim, spridapd, spridasc, sprcodso, sprannul, timestmp FROM multspr0 WHERE "
    SQL_UPDATE = "UPDATE multspr0 SET "
    SQL_INSERT = "INSERT INTO multspr0 "
    SQL_GET_TS = "SELECT timestmp FROM multspr0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "multspr0.spridapr", 0, True, 6, "Codice"
        .Add "Imballo", "multspr0.spridaim", 0, True, 3, "Imballo"
        .Add "Posizione Doganale", "multspr0.spridapd", 0, True, 3, "PosizioneDoganale"
        .Add "Società", "multspr0.spridasc", 0, True, 3, "Società"
        .Add "CodPerSoc", "multspr0.sprcodso", 0, True, 15, "CodPerSoc"
        .Add "Annullato", "multspr0.sprannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Imballo", "2"
        .Add "PosizioneDoganale", "3"
        .Add "Società", "4"
    End With
    DS.CancelField = "multspr0.sprannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


