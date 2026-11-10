VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParametriTariffe"
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

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

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
    
    SQL_FIND = "SELECT axtidaxt, axtdtvar, axtimpt1, axtidav1, " & _
               "       axtimpt2, axtidav2, axtannul, timestmp FROM tariaxt0 WHERE "
    SQL_GETLIST = "SELECT axtidaxt, axtdtvar FROM tariaxt0 "
    SQL_UPDATE = "UPDATE tariaxt0 SET "
    SQL_INSERT = "INSERT INTO tariaxt0 "
    SQL_GET_TS = "SELECT timestmp FROM tariaxt0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT axtidaxt, axtdtvar, axtannul FROM tariaxt0 "
    
    mvarFieldsIndexes = Array("0", "1", "6")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "tariaxt0.axtidaxt", 0, True, 3, "0"
        .Add "Data ultima variazione", "tariaxt0.axtdtvar", 2, True, 10, "1"
        .Add "Importo Tariffa minima 1", "tariaxt0.axtimpt1", 1, True, 13, "2"
        .Add "Valuta Tariffa minima 1", "tariaxt0.axtidav1", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Importo Tariffa minima 2", "tariaxt0.axtimpt2", 1, , 13, "4"
        .Add "Valuta Tariffa minima 2", "tariaxt0.axtidav2", 0, , 3, "5"
        With .Item("5")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Annullato", "tariaxt0.axtannul", 3, , , "6"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
    End With
    
    DS.CancelField = "tariaxt0.axtannul"
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


