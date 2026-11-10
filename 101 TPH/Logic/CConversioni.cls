VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConversioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CConversioni
' Data Creazione 4/11/99
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
    ' SQL_FIND = Restituisce i valori che vengono associati ai controlli
    SQL_FIND = "SELECT aucidau1,aucidau2,aucnrcon,aucdescr,aucannul,timestmp FROM anagauc0 WHERE "
    ' SQL_GETLIST = Restituisce i valori dei campi ceh vengono associati alla GRIGLIA
    SQL_GETLIST = "SELECT aucidau1,aucidau2,aucnrcon,aucdescr FROM anagauc0 "
    SQL_UPDATE = "UPDATE anagauc0 SET "
    SQL_INSERT = "INSERT INTO anagauc0 "
    SQL_GET_TS = "SELECT timestmp FROM anagauc0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aucidau1,aucidau2,aucnrcon,aucdescr,aucannul FROM anagauc0 "
    
    Set DS = New CDataStruct
    ' sono i campi che riempiono la griglia
    mvarFieldsIndexes = Array("CodiceUM1", "CodiceUM2", "Conversione", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    'costruisce la collezione contenente la definizione di tutti campi della tabella
    
    With DS.Fields
        .Add "Codice UM1", "anagauc0.aucidau1", 0, True, 3, "CodiceUM1"
        .Add "Codice UM2", "anagauc0.aucidau2", 0, True, 3, "CodiceUM2"
        .Add "Fattore di Conversione", "anagauc0.aucnrcon", 1, True, , "Conversione"
        .Add "Descrizione", "anagauc0.aucdescr", 0, True, 25, "Descrizione"
        .Add "Annullato", "anagauc0.aucannul", 3, , , "Annullato"
    End With
    'specifica il campo chiave
    With DS.KeyFields
        
        .Add "CodiceUM1", 1
        .Add "CodiceUM2", 2
    End With
    'specifica il campo di annullamento logico
    DS.CancelField = "anagauc0.aucannul"
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
