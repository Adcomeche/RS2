VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStipendioRiga"
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

Private SQL_GET_SELECTION_VIAGGI As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

Public Function Delete(varIDSDT As Variant) As Boolean
On Error GoTo ErrHandler
    
    With DS.DA
        If Not .ExecuteCommand(Me.SQL_DELETE, varIDSDT) Then
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
    On Error GoTo ErrHandler
    
    Insert = DS.Insert(Me.SQL_INSERT)
    
    Exit Function
ErrHandler:
    Insert = False
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean

    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    BeginSave = True
    ' ---
End Function
Private Sub Class_Initialize()
    SQL_FIND = "SELECT SDRIDSDT, SDRIDSDR, SDRIDADS, SDRDTRIF, " & _
               "       SDRNRQTA, SDRIDAUM, SDRNRIMU, SDRNRIMP, " & _
               "       timestmp " & _
               " FROM stipsdr0 WHERE "
    SQL_UPDATE = "UPDATE stipsdr0 SET "
    SQL_INSERT = "INSERT INTO stipsdr0 "
    SQL_GET_TS = "SELECT timestmp FROM stipsdr0 WHERE "
    SQL_DELETE = "DELETE stipsdr0 WHERE SDRIDSDT = ? AND SDRIDSDR = ? "

    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Stipendio", "stipsdr0.SDRIDSDT", 0, True, 7, "0", , True
        .Add "Progressivo Riga", "stipsdr0.SDRIDSDR", 0, True, 7, "1", , True
        .Add "Codice Addebito", "stipsdr0.SDRIDADS", 0, True, 3, "2"
        .Add "Data Riferimento", "stipsdr0.SDRDTRIF", 2, True, 10, "3"
        .Add "Quantità", "stipsdr0.SDRNRQTA", 1, True, 12, "4"
        .Add "Unità di Misura", "stipsdr0.SDRIDAUM", 0, True, 3, "5"
        .Add "Importo unitario", "stipsdr0.SDRNRIMU", 1, True, 12, "6"
        .Add "Importo", "stipsdr0.SDRNRIMP", 1, True, 12, "7"
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

