VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModalitaFatturazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CModalitaFatturazione
' Data Creazione 06/03/06
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
    
    SQL_GETLIST = "SELECT amfidamf, amfdescr, amfannul, timestmp FROM anagamf0 "
    SQL_FIND = "SELECT amfidamf, amfdescr, amfflrof, amfflor1, amfflor2, amfflor3, amfannul, timestmp FROM anagamf0 WHERE "
    SQL_UPDATE = "UPDATE anagamf0 SET "
    SQL_INSERT = "INSERT INTO anagamf0 "
    SQL_GET_TS = "SELECT timestmp FROM anagamf0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT amfidamf, amfdescr, amfannul FROM anagamf0 "
    SQL_GET_SEARCH_LIST_GROUPBY = ""
    
    Set DS = New CDataStruct
    'Lista dei campi della istruzione SQL_GET_SEARCH_LIST
    mvarFieldsIndexes = Array("0", "1", "6")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    'costruisce la collezione contenente la definizione di tutti campi della tabella
    
    With DS.Fields
        .Add "Codice", "anagamf0.amfidamf", 0, True, 3, "0"
        .Add "Descrizione", "anagamf0.amfdescr", 0, True, 25, "1"
        .Add "Raggr. Ord. Fattura", "anagamf0.amfflrof", 0, , 2, "2"
        .Add "Ord. Fattura 1", "anagamf0.amfflor1", 0, , 2, "3"
        .Add "Ord. Fattura 2", "anagamf0.amfflor2", 0, , 2, "4"
        .Add "Ord. Fattura 2", "anagamf0.amfflor3", 0, , 2, "5"
        .Add "Annullato", "anagamf0.amfannul", 3, , , "6"
    End With
    
    'specifica il campo chiave
    With DS.KeyFields
        .Add "0", "1"
    End With
    'specifica il campo di annullamento logico
    DS.CancelField = "anagamf0.amfannul"
End Sub
Private Sub Class_Terminate()
    Set DS = Nothing
    Set mcolListFields = Nothing
End Sub
Private Sub DS_ClientClear()
    RaiseEvent ClientClear
End Sub

Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub




