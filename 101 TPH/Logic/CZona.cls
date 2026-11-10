VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZona"
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

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
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
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT azoidazo, azodescr, azofleim, azoflemo, azocodma, azocodco, azoannul, timestmp FROM anagazo0 WHERE "
    SQL_GETLIST = "SELECT azoidazo, azodescr, azoflemo, azocodma FROM anagazo0 "
    SQL_UPDATE = "UPDATE anagazo0 SET "
    SQL_INSERT = "INSERT INTO anagazo0 "
    SQL_GET_TS = "SELECT timestmp FROM anagazo0 WHERE "
    
    Set DS = New CDataStruct
    SQL_GET_SEARCH_LIST = "SELECT azoidazo, azodescr, azoannul FROM anagazo0 "
    
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    With DS.Fields
        .Add "Codice", "anagazo0.azoidazo", 0, True, 3, "Codice"
        .Add "Descrizione", "anagazo0.azodescr", 0, True, 25, "Descrizione"
        .Add "Escludi Importazione", "anagazo0.azofleim", 3, , , "2"
        .Add "Escludi Mover", "anagazo0.azoflemo", 3, , , "3"
        .Add "Presidio", "anagazo0.azocodma", 0, , 15, "4"
        .Add "Centro operativo", "anagazo0.azocodco", 0, , 5, "5"
        .Add "Annullato", "anagazo0.azoannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    DS.CancelField = "anagazo0.azoannul"
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


