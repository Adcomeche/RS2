VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStipendioDettaglioAddebito"
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

Public Function Delete(varIDAAU As Variant, _
                        varDTOLV As Variant, _
                        varIDADS As Variant) As Boolean
On Error GoTo ErrHandler
    
    With DS.DA
        If Not .ExecuteCommand(Me.SQL_DELETE, _
                                varIDAAU, _
                                varDTOLV, _
                                varIDADS) Then
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

    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    BeginSave = True
    ' ---
End Function
Private Sub Class_Initialize()
    SQL_FIND = "SELECT olaidaau, oladtolv, olaidads, olanrqta, " & _
               "       olaidaum, oladescr, olanrimp, timestmp " & _
               " FROM stipola0 WHERE "
    SQL_GETLIST = "SELECT olaidaau, oladtolv, olaidads, olanrqta, " & _
                  "       olaidaum, oladescr, olanrimp " & _
                  "  FROM stipola0  "
    SQL_UPDATE = "UPDATE stipola0 SET "
    SQL_INSERT = "INSERT INTO stipola0 "
    SQL_GET_TS = "SELECT timestmp FROM stipola0 WHERE "
    SQL_DELETE = "DELETE stipola0 WHERE olaidaau = ? AND oladtolv = ? AND olaidads = ? "

    SQL_GET_SELECTION = "SELECT adsidads, adsdescr, '', adsidaum, '','' " & _
                        "  FROM stipads0 " & _
                        " WHERE adsannul = 0 AND adsflaut = 'M' "

    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice Autista", "stipola0.olaidaau", 0, True, 3, "0", , True
        .Add "Data", "stipola0.oladtolv", 2, True, 10, "1", , True
        .Add "Codice Addebito", "stipola0.olaidads", 0, True, 3, "2", , True
        With .Item("2")
            .ForeignTableEntity = "ADS"
            .ForeignTable = "stipads0"
            .ForeignTableFld = "adsidads"
            .ForeignTableDescField = "adsdescr"
            .ForeignTableDeleteFld = "adsannul"
        End With
        .Add "Quantità", "stipola0.olanrqta", 1, True, 3, "3", , True
        .Add "Unità di Misura", "stipola0.olaidaum", 0, True, 3, "4", , True
        With .Item("4")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Descrizione", "stipola0.oladescr", 0, , 64, "5"
        .Add "Importo", "stipola0.olanrimp", 1, , 12, "6"
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
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

