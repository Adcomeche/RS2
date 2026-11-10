VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSerbatoio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CSerbatoio
' Data Creazione 03/11/99
Option Explicit


Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String


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
    
    SQL_GETLIST = "SELECT asridasr,asrdescr,asridarr,asrannul,timestmp FROM anagasr0 "
    SQL_FIND = "SELECT asridasr,asridade,asrdescr FROM anagasr0 WHERE "
    SQL_UPDATE = "UPDATE anagasr0 SET "
    SQL_INSERT = "INSERT INTO anagasr0 "
    SQL_GET_TS = "SELECT timestmp FROM anagasr0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT asridasr,asridade,asrdescr, asrannul FROM anagasr0 "
    
    Set DS = New CDataStruct
    'Lista dei campi della istruzione SQL_GET_SEARCH_LIST
    mvarFieldsIndexes = Array("Codice", "Deposito", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    'costruisce la collezione contenente la definizione di tutti campi della tabella
    
    With DS.Fields
        .Add "Codice", "anagasr0.asridasr", 0, True, 3, "Codice"
        .Add "Deposito", "anagasr0.asridade", 0, True, 3, "Deposito"
        With .Item("Deposito")
            .ForeignTableEntity = "ADE"  ' Specifica che il campo è chiave primaria e chiave esterna
        End With
        .Add "Descrizione", "anagasr0.asrdescr", 0, , 25, "Descrizione"
        .Add "Raggruppamento", "anagasr0.asridarr", 0, , 3, "Raggruppamento"
        With .Item("Raggruppamento")
            .ForeignTableEntity = "ARR"
            .ForeignTable = "anagarr0"
            .ForeignTableFld = "arridarr"
            .ForeignTableDescField = "arrdescr"
            .ForeignTableDeleteFld = "arrannul"
        End With
        .Add "Annullato", "anagasr0.asrannul", 3, , , "Annullato"
    End With
    'specifica il campo chiave
    With DS.KeyFields
        .Add "Deposito", "1"
        .Add "Codice", "2"
    End With
    'specifica il campo di annullamento logico
    DS.CancelField = "anagasr0.asrannul"
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


