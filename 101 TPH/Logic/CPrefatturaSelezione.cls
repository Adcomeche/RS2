VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaSelezione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_SET_AVANZAMENTO As String
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
    
    SQL_FIND = "SELECT pptidppt, pptdtinp, pptdtfip, pptdescr, " & _
               "       timestmp FROM prefppt0 WHERE "
    SQL_GETLIST = "SELECT octidoct, octidasc FROM ordioct0 "
    SQL_GET_SELECTION = "SELECT pptidppt, pptdtinp, pptdtfip, pptdescr " & _
                        "  FROM prefppt0  "
'    mvarFieldsIndexes = Array("0", "1", "2", "17", "12", "13", "5", _
'                              "3", "8", "7", "18", "21", "23")
'    Set mcolListFields = New Collection
'    For i = 0 To UBound(mvarFieldsIndexes)
'        mcolListFields.Add mvarFieldsIndexes(i)
'    Next i
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Dal Numero", "prefppt0.pptidppt", 0, , 7, "0", " >= "
        .Add "Al Numero", "prefppt0.pptidppt", 0, , 7, "1", " <= "
        .Add "Dalla Data", "prefppt0.pptdtinp", 2, , 10, "2", " >= "
        .Add "Alla Data", "prefppt0.pptdtfip", 2, , 10, "3", " <= "
        .Add "Società", "prefppt0.pptidasc", 0, True, 3, "4"
        With .Item("4")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Tipo prefattura", "prefppt0.pptfltip", 0, True, 1, "5"
    End With
    DS.CancelField = "ordioct0.octannul"
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
        .Add "4", "5"
        .Add "5", "6"
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




