VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMagazzinoDeposito"
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
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

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
    
    SQL_FIND = "SELECT amdidamd, amdidade, amddescr, amdidasc, " & _
               "amdidatv, amdfldis, amdannul, timestmp FROM anagamd0 WHERE "
    SQL_GETLIST = "SELECT amdidamd, amddescr, amdidasc, amdidatv, amdfldis, amdannul FROM anagamd0 "
    SQL_UPDATE = "UPDATE anagamd0 SET "
    SQL_INSERT = "INSERT INTO anagamd0 "
    SQL_GET_TS = "SELECT timestmp FROM anagamd0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT amdidamd, amdidade, amddescr, amdidasc, amdidatv, amdfldis, " & _
                          "       amdannul FROM anagamd0 "
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "5", "6")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anagamd0.amdidamd", 0, True, 3, "0"
        .Add "Codice Deposito", "anagamd0.amdidade", 0, True, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ADE"
        End With
        .Add "Descrizione", "anagamd0.amddescr", 0, True, 25, "2"
        .Add "Società", "anagamd0.amdidasc", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Tipo Vendita", "anagamd0.amdidatv", 0, True, 3, "4"
        With .Item("4")
            .ForeignTableEntity = "ATV"
            .ForeignTable = "anagatv0"
            .ForeignTableFld = "atvidatv"
            .ForeignTableDescField = "atvdescr"
            .ForeignTableDeleteFld = "atvannul"
        End With
        .Add "Disponibilità", "anagamd0.amdfldis", 0, True, 1, "5"
        .Add "Annullato", "anagamd0.amdannul", 3, , , "6"
    End With
    With DS.KeyFields
        .Add "1", "1"
        .Add "0", "2"
    End With
    DS.CancelField = "anagamd0.amdannul"
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




