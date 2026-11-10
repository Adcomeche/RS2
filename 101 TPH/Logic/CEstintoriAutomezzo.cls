VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEstintoriAutomezzo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public DS As CDataStruct

Public SQL_GET_TS As String
Public SQL_INSERT As String
Public SQL_UPDATE As String
Public SQL_FIND As String
Public SQL_GETLIST As String

Private Const SQL_CLEAR_ESTINTORI = "DELETE FROM MULTEAM0 WHERE EAMIDAAM = ?"

Private Const SQL_GET_ESTINTORI = "SELECT EAMIDEST FROM MULTEAM0 WHERE EAMANNUL = 0 AND EAMIDAAM = ?"

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
    SQL_GETLIST = "SELECT eamidest, eamidaam, eamannul, timestmp FROM multeam0 "
    SQL_FIND = "SELECT eamidest, eamidaam, eamannul, timestmp FROM multeam0 WHERE "
    SQL_UPDATE = "UPDATE multeam0 SET "
    SQL_INSERT = "INSERT INTO multeam0 "
    SQL_GET_TS = "SELECT timestmp FROM multeam0 WHERE "
    
    Set DS = New CDataStruct
    With DS.Fields
        .Add "Codice Estintore", "multeam0.eamidest", 0, True, 7, "Codice"
        .Add "Targa", "multeam0.eamidaam", 0, True, 10, "Targa"
        .Add "Annullato", "multeam0.eamannul", 3, , , "Annullato"
    End With
    
    With DS.KeyFields
        .Add "Codice", "1"
        .Add "Targa", "2"
    End With
    DS.CancelField = "multeam0.eamannul"
End Sub

Private Sub Class_Terminate()
    Set DS = Nothing
End Sub

Public Function ClearEstintori(varIDAAM As Variant) As Boolean
    On Error GoTo ErrHandler

    If Not DS.DA.ExecuteCommand(SQL_CLEAR_ESTINTORI, varIDAAM) Then
        GoTo ErrHandler
    End If

    ClearEstintori = True
    Exit Function

ErrHandler:
    ClearEstintori = False
End Function

Public Function blnGetEstintori(ByRef varBuffer As Variant, varIDAAM As Variant) As Boolean
    On Error GoTo ErrHandler

    If Not DS.DA.QueryCommand(SQL_GET_ESTINTORI, varBuffer, 0, varIDAAM) Then
        GoTo ErrHandler
    End If

    blnGetEstintori = True
    Exit Function

ErrHandler:
    blnGetEstintori = False
End Function
