VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumeratoreDocumentiAmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe Numeratori Documenti
' Data Creazione  4/9/99
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
    
    SQL_FIND = " SELECT NDAIDASC, NDAIDAEE, NDAIDATO, NDADESCR, " & _
               "        NDANRPRG, NDADTAGG, TIMESTMP " & _
               "   FROM docunda0 WHERE "

    SQL_GETLIST = "SELECT NDAIDATO, NDADESCR FROM DOCUNDA0 "
    
    SQL_UPDATE = "UPDATE docunda0 SET "
    
    SQL_INSERT = "INSERT INTO docunda0 "
    
    SQL_GET_TS = "SELECT timestmp FROM docunda0 WHERE "
    
    SQL_GET_SEARCH_LIST = "SELECT NDAIDATO, NDAIDASC, NDAIDAEE, NDADESCR FROM docunda0 "
    
    mvarFieldsIndexes = Array("0", "1", "2", "3")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Società", "docunda0.ndaidasc", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "docunda0.ndaidaee", 0, True, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Tipo Documento", "docunda0.ndaidato", 0, True, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "ATO"
            .ForeignTable = "docuato0"
            .ForeignTableFld = "atoidato"
            .ForeignTableDescField = "atodescr"
            .ForeignTableDeleteFld = "atoannul"
        End With
        .Add "Descrizione", "docunda0.ndadescr", 0, True, 40, "3"
        .Add "Progressivo", "docunda0.ndanrprg", 0, True, 7, "4"
        .Add "Data Aggiornamento", "docunda0.ndadtagg", 2, True, 10, "5"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
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


