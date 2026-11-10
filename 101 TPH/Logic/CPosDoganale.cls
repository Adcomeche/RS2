VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPosDoganale"
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
    
    SQL_FIND = "SELECT apdidapd, apddescr, apdannul, timestmp FROM anagapd0 WHERE "
    SQL_GETLIST = "SELECT apdidapd, apddescr FROM anagapd0 "
    SQL_UPDATE = "UPDATE anagapd0 SET "
    SQL_INSERT = "INSERT INTO anagapd0 "
    SQL_GET_TS = "SELECT timestmp FROM anagapd0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT apdidapd, apddescr, apdannul " & _
                          " FROM anagapd0 "
    
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CPosDoganaleSoc
    
    With DS.Fields
        .Add "Codice", "anagapd0.apdidapd", 0, True, 3, "Codice"
        .Add "Descrizione", "anagapd0.apddescr", 0, True, 25, "Descrizione"
        .Add "Annullato", "anagapd0.apdannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    DS.CancelField = "anagapd0.apdannul"
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




