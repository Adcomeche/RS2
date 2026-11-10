VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContrattoDettaglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION_EXTENDED As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_DELETE As String

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

Public Function Delete(varIDACL As Variant, _
                        varIDACN As Variant, _
                        varIDADD As Variant) As Boolean
On Error GoTo ErrHandler
    
    With DS.DA
        If Not .ExecuteCommand(Me.SQL_DELETE, _
                                varIDACL, _
                                varIDACN, _
                                varIDADD) Then
            GoTo ErrHandler
        End If
        
    End With
    
    Delete = True
    Exit Function
ErrHandler:
    Delete = False
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT acdidacl, acdidacn, acdidadd, acdfludm, " & _
               "       acdminta, acdflpat, acdfldif, acdidttc, " & _
               "       timestmp FROM anagacd0 WHERE "
    SQL_GETLIST = "SELECT acdidacl, acdidacn, acdidadd, acdfludm, " & _
                  "       acdminta, acdflpat, acdidttc FROM anagacd0 "
    SQL_GET_SELECTION_EXTENDED = "SELECT acdidacl, acdidacn, acdidadd, adddescr, acdfludm, " & _
                  "       ANAGAVV_FLUDM.AVVDESCR, acdminta, acdflpat, ANAGAVV_FLPAT.AVVDESCR, " & _
                  "       acdfldif, ANAGAVV_FLDIF.AVVDESCR FROM anagacd0 " & _
                  "       INNER JOIN ANAGADD0 ON ADDIDADD = ACDIDADD" & _
                  "       LEFT JOIN ANAGAVV ANAGAVV_FLPAT ON ANAGAVV_FLPAT.AVVIDTBL = 'ANAGACD0'" & _
                  "            AND ANAGAVV_FLPAT.AVVIDFLD = 'ACDFLPAT' AND" & _
                  "            ANAGAVV_FLPAT.AVVVALUE = ANAGACD0.ACDFLPAT" & _
                  "       LEFT JOIN ANAGAVV ANAGAVV_FLDIF ON ANAGAVV_FLDIF.AVVIDTBL = 'ANAGACD0'" & _
                  "            AND ANAGAVV_FLDIF.AVVIDFLD = 'ACDFLDIF' AND" & _
                  "            ANAGAVV_FLDIF.AVVVALUE = ANAGACD0.ACDFLDIF" & _
                  "       LEFT JOIN ANAGAVV ANAGAVV_FLUDM ON ANAGAVV_FLUDM.AVVIDTBL = 'ANAGACD0'" & _
                  "            AND ANAGAVV_FLUDM.AVVIDFLD = 'ACDFLUDM' AND" & _
                  "            ANAGAVV_FLUDM.AVVVALUE = ANAGACD0.ACDFLUDM"
    SQL_UPDATE = "UPDATE anagacd0 SET "
    SQL_INSERT = "INSERT INTO anagacd0 "
    SQL_GET_TS = "SELECT timestmp FROM anagacd0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT acdidacl, acdidacn, acdidadd FROM anagacd0 "
    SQL_DELETE = "DELETE anagacd0 WHERE acdidacl = ? AND acdidacn = ? AND acdidadd = ?"

    Set DS = New CDataStruct
    mvarFieldsIndexes = Array("0", "1", "2")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Cliente", "anagacd0.acdidacl", 0, True, 10, "0"
        .Add "Contratto", "anagacd0.acdidacn", 0, True, 3, "1"
        .Add "Addebito", "anagacd0.acdidadd", 0, True, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "ADD"
            .ForeignTable = "anagadd0"
            .ForeignTableFld = "addidadd"
            .ForeignTableDescField = "adddescr"
            .ForeignTableDeleteFld = "addannul"
        End With
        .Add "Flag UDM", "anagacd0.acdfludm", 0, False, 2, "3"
        .Add "Minimo tassabile", "anagacd0.acdminta", 1, , 13, "4"
        .Add "Parametri tassazione", "anagacd0.acdflpat", 0, True, 2, "5"
        .Add "Separazione Fattura", "anagacd0.acdfldif", 0, True, 2, "6"
        .Add "Tipo Calcolo", "anagacd0.acdidttc", 0, True, 3, "7"
        With .Item("7")
            .ForeignTableEntity = "TTC"
            .ForeignTable = "tarittc0"
            .ForeignTableFld = "ttcidttc"
            .ForeignTableDescField = "ttcdescr"
            .ForeignTableDeleteFld = "ttcannul"
        End With
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
    End With
    DS.CancelField = ""
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

Public Function GetSelectionExtended(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    
    If Not IsMissing(OrderFields) Then
        GetSelectionExtended = DS.GetSelection(SQL_GET_SELECTION_EXTENDED, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionExtended = DS.GetSelection(SQL_GET_SELECTION_EXTENDED, ActiveRecordsOnly, , MaxRecords)
    End If

End Function

