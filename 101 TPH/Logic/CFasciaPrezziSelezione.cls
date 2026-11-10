VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFasciaPrezziSelezione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
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
Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Public Function GetSelection(Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, False, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, False, MaxRecords)
    End If
End Function
Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property
Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function
Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT FPZIDFPZ,FPZDESCR,FPZANNUL, " & _
               "       timestmp FROM PREZFPZ WHERE "
    SQL_GET_SELECTION = "SELECT FPZIDFPZ,FPZDESCR " & _
                        "  FROM prezfpz0  "
 
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Dal Numero", "prezfpz0.fpzidfpz", 0, , 7, "0", " >= "
        .Add "Al Numero", "prezfpz0.fpzidfpz", 0, , 7, "1", " <= "
    End With
    DS.CancelField = "ordioct0.octannul"
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
    End With
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
