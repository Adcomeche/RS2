VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEnteEmittente"
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

Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
    
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function


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
    
    SQL_FIND = "SELECT aeeidaee, aeedescr, aeeindir, aeelocal," & _
               "aeeprovi, aeenrcap, aeeidasc, aeeannul, timestmp FROM anagaee0 WHERE "
    SQL_GETLIST = "SELECT aeeidaee, aeedescr FROM anagaee0 "
    SQL_UPDATE = "UPDATE anagaee0 SET "
    SQL_INSERT = "INSERT INTO anagaee0 "
    SQL_GET_TS = "SELECT timestmp FROM anagaee0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aeeidaee, aeedescr, aeeidasc, aeeindir, aeelocal," & _
               "aeeprovi, aeenrcap, aeeannul FROM anagaee0 "
    
    mvarFieldsIndexes = Array("0", "1", "6", "2", "3", "4", "5", "7")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CEnteEmittenteSoc
    
    With DS.Fields
        .Add "Codice", "anagaee0.aeeidaee", 0, True, 3, "0"
        .Add "Descrizione", "anagaee0.aeedescr", 0, True, 25, "1"
        .Add "Indirizzo", "anagaee0.aeeindir", 0, False, 40, "2"
        .Add "Località", "anagaee0.aeelocal", 0, False, 25, "3"
        .Add "Provincia", "anagaee0.aeeprovi", 0, False, 2, "4"
        .Add "CAP", "anagaee0.aeenrcap", 0, False, 10, "5"
        .Add "Società", "anagaee0.aeeidasc", 0, True, 3, "6"
        With .Item("6")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Annullato", "anagaee0.aeeannul", 3, , , "7"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagaee0.aeeannul"
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




