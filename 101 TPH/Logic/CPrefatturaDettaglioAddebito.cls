VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaDettaglioAddebito"
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
Public SQL_DELETE As String

Private SQL_GET_SELECTION As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

'Public Function Delete(Optional IDTestata As String, _
'                       Optional IDRiga As String) As Boolean
'    On Error GoTo ErrHandler
'    Dim strIDTestata As String
'    Dim strIDRiga As String
'
'    If Len(IDTestata) = 0 Then
'        strIDTestata = DS.Fields("0").Value
'    Else
'        strIDTestata = IDTestata
'    End If
'
'    If Len(IDRiga) = 0 Then
'        strIDRiga = DS.Fields("1").Value
'    Else
'        strIDRiga = IDRiga
'    End If
'
'    With DS.DA
'        If Not .ExecuteCommand(Me.SQL_DELETE, strIDTestata, strIDRiga) Then
'            GoTo ErrHandler
'        End If
'
'    End With
'
'    Delete = True
'    Exit Function
'ErrHandler:
'    Delete = False
'End Function

Public Function Delete(varIDPPT As Variant, _
                        varIDPPD As Variant, _
                        varIDACL As Variant, _
                        varIDACN As Variant, _
                        varIDADD) As Boolean
On Error GoTo ErrHandler
    
    With DS.DA
        If Not .ExecuteCommand(Me.SQL_DELETE, varIDPPT, _
                                varIDPPD, _
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

Public Function Find() As Boolean
    Dim blnResult As Boolean
    
    blnResult = DS.Find(Me.SQL_FIND)
    Find = blnResult
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean

'    If DS.Fields("44").NewValue <> DS.Fields("44").Value Then
'        If DS.Fields("45").Value <> "S" Then
'            DS.Fields("45").NewValue = "S"
'            DS.Fields("45").Updated = True
'        End If
'        If CDbl(DS.Fields("44").NewValue) = CDbl(DS.Fields("36").Value) Then
'            DS.Fields("50").NewValue = CLng(DS.Fields("50").Value) And 65503
'            DS.Fields("50").Updated = True
'        Else
'            DS.Fields("50").NewValue = CLng(DS.Fields("50").Value) Or 32
'            DS.Fields("50").Updated = True
'        End If
'    End If
    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    BeginSave = True
    ' ---
End Function
Private Sub Class_Initialize()
    SQL_FIND = "SELECT padidppt, padidppd, padidacl, padidacn, " & _
               "       padidadd, padnrcon, padnrkms, padnrlts, " & _
               "       padimtar, padimimp, timestmp " & _
               " FROM prefpad0 WHERE "
    SQL_GETLIST = "SELECT padidppt, padidpad, padidadd, padnrcon, " & _
               "       padnrkms, padnrlts " & _
                  "  FROM prefpad0  "
    SQL_UPDATE = "UPDATE prefpad0 SET "
    SQL_INSERT = "INSERT INTO prefpad0 "
    SQL_GET_TS = "SELECT timestmp FROM prefpad0 WHERE "
    SQL_DELETE = "DELETE PREFPAD0 WHERE padidppt = ? AND padidppd = ? AND padidacl = ? AND padidacn = ? AND padidadd = ?"

    SQL_GET_SELECTION = "SELECT padidppt, padidppd, padidacl, aclrasoc, " & _
                        " padidacn, acndescr, padidadd, adddescr, " & _
                        " padnrcon, padnrkms, padnrlts, padimtar, " & _
                        " padimimp " & _
                        " FROM PREFPAD0 " & _
                        " LEFT JOIN ANAGACL0 ON aclidacl = padidacl " & _
                        " LEFT JOIN ANAGACN0 ON acnidacl = padidacl AND acnidacn = padidacn " & _
                        " LEFT JOIN ANAGADD0 ON addidadd = padidadd "

    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Prefattura", "prefpad0.padidppt", 0, True, 7, "0", , True
        .Add "Numero Dettaglio", "prefpad0.padidppd", 1, True, 7, "1", , True
        .Add "Cliente", "prefpad0.padidacl", 0, True, 10, "2", , True
        With .Item("2")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Contratto", "prefpad0.padidacn", 0, True, 3, "3", , True
        With .Item("3")
            .ForeignTableEntity = "ACN"
            .ForeignTable = "anagacn0"
            .ForeignTableFld = "acnidacn"
            .ForeignTableDescField = "acndescr"
            .ForeignTableDeleteFld = "acnannul"
        End With
        .Add "Codice Addebito", "prefpad0.padidadd", 0, True, 3, "4", , True
        With .Item("4")
            .ForeignTableEntity = "ADD"
            .ForeignTable = "anagadd0"
            .ForeignTableFld = "addidadd"
            .ForeignTableDescField = "adddescr"
            .ForeignTableDeleteFld = "addannul"
        End With
        .Add "Numero consegne", "prefpad0.padnrcon", 1, False, 13, "5"
        .Add "Kilometri", "prefpad0.padnrkms", 1, False, 13, "6"
        .Add "Litri", "prefpad0.padnrlts", 1, False, 13, "7"
        .Add "Valore Tariffa", "prefpad0.padimtar", 1, False, 11, "8"
        .Add "Importo", "prefpad0.padimimp", 1, False, 11, "9"
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
        .Add "4", "5"
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


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, , MaxRecords)
    End If

End Function

