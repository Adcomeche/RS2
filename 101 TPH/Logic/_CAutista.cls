VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutista"
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
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Private Sub DS_ClientClear()
    RaiseEvent ClientClear
End Sub

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function


Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT aauidaau, aaunomea, aaunomab," & _
               "aauidave, aaunumpa, aaufldis, aaunote1," & _
               "aaunote2, aauannul, timestmp FROM anagaau0 WHERE "
    SQL_GETLIST = "SELECT aauidaau, aaunomea FROM anagaau0 "
    SQL_UPDATE = "UPDATE anagaau0 SET "
    SQL_INSERT = "INSERT INTO anagaau0 "
    SQL_GET_TS = "SELECT timestmp FROM anagaau0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aauidaau, aaunomea, aauidave,aaufldis, aauannul FROM anagaau0 "
    
    mvarFieldsIndexes = Array("0", "1", "3", "5", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anagaau0.aauidaau", 0, True, 3, "0"
        .Add "Nome", "anagaau0.aaunomea", 0, True, 30, "1"
        .Add "Nome Abbreviato", "anagaau0.aaunomab", 0, , 6, "2"
        .Add "Vettore", "anagaau0.aauidave", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Numero Patente", "anagaau0.aaunumpa", 0, , 15, "4"
        .Add "Disponibilità", "anagaau0.aaufldis", 0, True, 1, "5"
        .Add "Note 1", "anagaau0.aaunote1", 0, , 40, "6"
        .Add "Note 2", "anagaau0.aaunote2", 0, , 40, "7"
        .Add "Annullato", "anagaau0.aauannul", 3, , , "8"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagaau0.aauannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub




