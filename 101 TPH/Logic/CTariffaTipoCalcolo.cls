VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTariffaTipoCalcolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    
    SQL_FIND = "SELECT ttcidttc, ttcdescr, ttcannul, timestmp FROM tarittc0 WHERE "
    SQL_GETLIST = "SELECT ttcidttc, ttcdescr FROM tarittc0 "
    SQL_UPDATE = "UPDATE tarittc0 SET "
    SQL_INSERT = "INSERT INTO tarittc0 "
    SQL_GET_TS = "SELECT timestmp FROM tarittc0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT ttcidttc, ttcdescr, " & _
                          "ttcannul FROM tarittc0 "
    Set DS = New CDataStruct
    mvarFieldsIndexes = Array("0", "1", "2")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    With DS.Fields
        .Add "Codice", "tarittc0.ttcidttc", 0, True, 3, "0"
        .Add "Descrizione", "tarittc0.ttcdescr", 0, True, 25, "1"
        .Add "Annullato", "tarittc0.ttcannul", 3, , , "2"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "tarittc0.ttcannul"
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

Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property



