VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CTipoDocumento
' Data Creazione 03/11/99
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
    SQL_FIND = "SELECT atoidato,atodescr,atoflut1,atoflut2,atoflut3,atoflut4,atoflut5,atoannul,timestmp FROM docuato0 WHERE "
    ' SQL_GETLIST = Restituisce i valori dei campi ceh vengono associati alla GRIGLIA
    SQL_GETLIST = "SELECT atoidato,atodescr FROM docuato0 "
    SQL_UPDATE = "UPDATE docuato0 SET "
    SQL_INSERT = "INSERT INTO docuato0 "
    SQL_GET_TS = "SELECT timestmp FROM docuato0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT atoidato,atodescr, " & _
                          "atoannul FROM docuato0 "
    
    Set DS = New CDataStruct
    ' sono i campi che riempiono la griglia
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    'costruisce la collezione contenente la definizione di tutti campi della tabella
    
    With DS.Fields
        .Add "Codice", "docuato0.atoidato", 0, True, 3, "Codice"
        .Add "Descrizione", "docuato0.atodescr", 0, True, 25, "Descrizione"
        .Add "Flag Utente 1", "docuato0.atoflut1", 3, , , "FlagUtente1"
        .Add "Flag Utente 2", "docuato0.atoflut2", 3, , , "FlagUtente2"
        .Add "Flag Utente 3", "docuato0.atoflut3", 3, , , "FlagUtente3"
        .Add "Flag Utente 4", "docuato0.atoflut4", 3, , , "FlagUtente4"
        .Add "Flag Utente 5", "docuato0.atoflut5", 3, , , "FlagUtente5"
        .Add "Annullato", "docuato0.atoannul", 3, , , "Annullato"
    End With
    'specifica il campo chiave
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    'specifica il campo di annullamento logico
    DS.CancelField = "docuato0.atoannul"
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


