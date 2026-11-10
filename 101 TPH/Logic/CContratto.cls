VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContratto"
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
Public SQL_GET_SELECTION_LIST As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

'Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
'    On Error GoTo ErrHandler
'
'    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
'
'    Exit Function
'ErrHandler:
'    DecodeFirmCode = vbNullString
'End Function
'
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
    
    SQL_FIND = "SELECT acnidacl, acnidacn, acnnumco, " & _
               "       acnnumse, acndescr, acnnotec, " & _
               "       acnkmcon, acnidatn, acnidasc, " & _
               "       acnidave, acnfltit, acnidamf, " & _
               "       acnannul, timestmp FROM anagacn0 WHERE "
    SQL_GETLIST = "SELECT acnidacl, acnidacn, acndescr FROM anagacn0 "
    SQL_UPDATE = "UPDATE anagacn0 SET "
    SQL_INSERT = "INSERT INTO anagacn0 "
    SQL_GET_TS = "SELECT timestmp FROM anagacn0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT acnidacl, acnidacn, acndescr, " & _
                          "acnannul FROM anagacn0 "
    SQL_GET_SELECTION_LIST = "SELECT acnidacl, acnidacn, acndescr FROM anagacn0 "
    
    Set DS = New CDataStruct
    mvarFieldsIndexes = Array("11", "0", "3", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Cliente", "anagacn0.acnidacl", 0, True, 10, "11"
        With .Item("11")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Codice", "anagacn0.acnidacn", 0, True, 3, "0"
        .Add "Numero", "anagacn0.acnnumco", 0, True, 10, "1"
        .Add "Numero secondario", "anagacn0.acnnumse", 0, False, 10, "2"
        .Add "Descrizione", "anagacn0.acndescr", 0, True, 25, "3"
        .Add "Note", "anagacn0.acnnotec", 0, , 50, "4"
        .Add "KM a contratto", "anagacn0.acnkmcon", 1, True, 8, "5"
        .Add "Tipo contratto", "anagacn0.acnidatn", 0, True, 3, "6"
        With .Item("6")
            .ForeignTableEntity = "ATN"
            .ForeignTable = "anagatn0"
            .ForeignTableFld = "atnidatn"
            .ForeignTableDescField = "atndescr"
            .ForeignTableDeleteFld = "atnannul"
        End With
        .Add "Società", "anagacn0.acnidasc", 0, True, 3, "7"
        With .Item("7")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Vettore", "anagacn0.acnidave", 0, True, 3, "8"
        With .Item("8")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Tipo tariffa", "anagacn0.acnfltit", 0, , 2, "10"
        .Add "Modalità Fatturazione", "anagacn0.acnidamf", 0, True, 3, "12"
        With .Item("12")
            .ForeignTableEntity = "AMF"
            .ForeignTable = "anagamf0"
            .ForeignTableFld = "amfidamf"
            .ForeignTableDescField = "amfdescr"
            .ForeignTableDeleteFld = "amfannul"
        End With

        .Add "Annullato", "anagacn0.acnannul", 3, , , "9"
    End With
    With DS.KeyFields
        .Add "11", "1"
        .Add "0", "2"
    End With
    DS.CancelField = "anagacn0.acnannul"
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


Public Function GetSelectionList(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    
    If Not IsMissing(OrderFields) Then
        GetSelectionList = DS.GetSelection(SQL_GET_SELECTION_LIST, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionList = DS.GetSelection(SQL_GET_SELECTION_LIST, ActiveRecordsOnly, , MaxRecords)
    End If

End Function


