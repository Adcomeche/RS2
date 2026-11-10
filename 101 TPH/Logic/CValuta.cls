VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CValuta"
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
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

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
    Dim i As Integer
    
    SQL_FIND = "SELECT avaidava, avadescr, avaabbre, avaannul, timestmp FROM anagava0 WHERE "
    SQL_GETLIST = "SELECT avaidava, avadescr FROM anagava0 "
    SQL_UPDATE = "UPDATE anagava0 SET "
    SQL_INSERT = "INSERT INTO anagava0 "
    SQL_GET_TS = "SELECT timestmp FROM anagava0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT avaidava, avadescr, avaannul FROM anagava0 "
    
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anagava0.avaidava", 0, True, 3, "Codice"
        .Add "Descrizione", "anagava0.avadescr", 0, True, 30, "Descrizione"
        .Add "Abbreviazione", "anagava0.avaabbre", 0, True, 5, "Abbreviazione"
        .Add "Annullato", "anagava0.avaannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    DS.CancelField = "anagava0.avaannul"
End Sub



Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


Private Sub DS_ClientClear()
    RaiseEvent ClientClear

End Sub

Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh

End Sub


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property


