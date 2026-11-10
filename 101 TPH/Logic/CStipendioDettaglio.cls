VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStipendioDettaglio"
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
    SQL_FIND = "SELECT olvidaau, olvdtolv, olvdtinl, olvhrinl, " & _
               "       olvdtilr, olvhrilr, olvdtfil, olvhrfil, " & _
               "       olvnrkms, olvflmas, olvannul, timestmp " & _
               " FROM stipolv0 WHERE "
    SQL_GETLIST = "SELECT olvidaau, olvdtolv, olvdtinl, olvhrinl, " & _
                  "       olvdtilr, olvhrilr, olvdtfil, olvhrfil, " & _
                  "       olvnrkms " & _
                  "  FROM stipolv0  "
    SQL_UPDATE = "UPDATE stipolv0 SET "
    SQL_INSERT = "INSERT INTO stipolv0 "
    SQL_GET_TS = "SELECT timestmp FROM stipolv0 WHERE "
    SQL_DELETE = "DELETE stipolv0 WHERE olvidaau = ? AND olvdtolv = ? "

    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice Autista", "stipolv0.olvidaau", 0, True, 3, "0", , True
        .Add "Data", "stipolv0.olvdtolv", 2, True, 10, "1", , True
        .Add "Data Inizio Lavoro Prevista", "stipolv0.olvdtinl", 2, True, 10, "2", , True
        .Add "Ora Inizio Lavoro Prevista", "stipolv0.olvhrinl", 4, True, 5, "3", , True
        .Add "Data Inizio Lavoro", "stipolv0.olvdtilr", 2, True, 10, "4"
        .Add "Ora Inizio Lavoro", "stipolv0.olvhrilr", 4, True, 5, "5"
        .Add "Data Fine Lavoro", "stipolv0.olvdtfil", 2, True, 10, "6"
        .Add "Ora Fine Lavoro", "stipolv0.olvhrfil", 4, True, 5, "7"
        .Add "Km percorsi", "stipolv0.olvnrkms", 1, False, 6, "8"
        .Add "Flag motivo assenza", "stipolv0.olvflmas", 0, False, 2, "10"
        .Add "Annullato", "stipolv0.olvannul", 3, , , "9"
    End With
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

