VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumeratore"
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

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

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
    ' --- Impedisco le modifiche della descrizione del numeratore
    DS.Fields("1").Updated = False
    ' ---
    
    ' --- Imposto la data ultima modifica alla data odierna
    DS.Fields("3").NewValue = Format$(Date, "dd/mm/yyyy")
    DS.Fields("3").Updated = True
    '---
    
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    SQL_FIND = "SELECT anuidanu, anudescr, anuprogr, anudtagg, timestmp FROM anaganu0 WHERE "
    SQL_GETLIST = "SELECT anuidanu, anudescr FROM anaganu0 "
    SQL_UPDATE = "UPDATE anaganu0 SET "
    SQL_INSERT = "INSERT INTO anaganu0 "
    SQL_GET_TS = "SELECT timestmp FROM anaganu0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anaganu0.anuidanu", 0, True, 3, "0"
        .Add "Descrizione", "anaganu0.anudescr", 0, True, 30, "1"
        .Add "Progressivo", "anaganu0.anuprogr", 0, True, 7, "2"
        .Add "Data Aggiornamento", "anaganu0.anudtagg", 2, True, 10, "3"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = ""
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




