VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFasciaPrezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CFasciaPrezzi
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
    
    SQL_GETLIST = "SELECT fpzidfpz, fpzdescr FROM prezfpz0 "
    SQL_FIND = "SELECT fpzidfpz, fpzdescr, fpzannul, timestmp FROM prezfpz0 WHERE "
    SQL_UPDATE = "UPDATE prezfpz0 SET "
    SQL_INSERT = "INSERT INTO prezfpz0 "
    SQL_GET_TS = "SELECT timestmp FROM prezfpz0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT fpzidfpz, fpzdescr, " & _
                          "fpzannul FROM prezfpz0 "
    
    Set DS = New CDataStruct
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    With DS.Fields
        .Add "Codice", "prezfpz0.fpzidfpz", 0, True, 7, "Codice"
        .Add "Descrizione", "prezfpz0.fpzdescr", 0, True, 35, "Descrizione"
        .Add "Annullato", "prezfpz0.fpzannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    DS.CancelField = "prezfpz0.fpzannul"
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


