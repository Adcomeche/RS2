VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocieta"
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
    
    SQL_GETLIST = "SELECT ascidasc, ascrasoc FROM anagasc0 "
    SQL_FIND = "SELECT ascidasc, ascrasoc, ascindir, asclocal, " & _
                "       ascprovi, ascnrcap, asccofis, ascpaiva, " & _
                "       ascidamp, ascidaiv, ascannul, timestmp " & _
                "  FROM anagasc0 WHERE "
                             
    SQL_UPDATE = "UPDATE anagasc0 SET "
    SQL_INSERT = "INSERT INTO anagasc0 "
    SQL_GET_TS = "SELECT timestmp FROM anagasc0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT ascidasc, ascrasoc, ascindir, asclocal, " & _
                          "       ascprovi, ascnrcap, ascannul FROM anagasc0 "
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CSocietaSoc
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "5", "10")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i

    With DS.Fields
        .Add "Codice", "anagasc0.ascidasc", 0, True, 3, "0"
        .Add "Ragione Sociale", "anagasc0.ascrasoc", 0, True, 25, "1"
        .Add "Indirizzo", "anagasc0.ascindir", 0, False, 40, "2"
        .Add "Località", "anagasc0.asclocal", 0, False, 25, "3"
        .Add "Provincia", "anagasc0.ascprovi", 0, False, 2, "4"
        .Add "CAP", "anagasc0.ascnrcap", 0, False, 10, "5"
        .Add "Codice Fiscale", "anagasc0.asccofis", 0, False, 16, "6"
        .Add "Partita IVA", "anagasc0.ascpaiva", 0, False, 20, "7"
        .Add "Modalità di Pagamento", "anagasc0.ascidamp", 0, True, 3, "8"
        With .Item("8")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "Aliquota IVA", "anagasc0.ascidaiv", 0, True, 3, "9"
        With .Item("9")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Annullato", "anagasc0.ascannul", 3, , , "10"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagasc0.ascannul"
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




