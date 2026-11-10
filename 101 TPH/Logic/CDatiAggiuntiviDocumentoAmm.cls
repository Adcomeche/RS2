VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDatiAggiuntiviDocumentoAmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Classe CDatiAggiuntiviDocumentoAmm
'Data Creazione 4/11/99
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
Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function
Private Sub DS_ClientClear()
    RaiseEvent ClientClear
End Sub
Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function
Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function
Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    SQL_FIND = "SELECT DAGIDAEE, DAGIDAND, DAGIDDAG, DAGRASOC, " & _
               " DAGINDIR, DAGLOCAL, DAGPROVI, DAGNRCAP, DAGANNUL, TIMESTMP " & _
               " FROM DOCUDAG0 WHERE "
               
    SQL_GETLIST = ""
    
    SQL_UPDATE = "UPDATE docudag0 SET "
    
    SQL_INSERT = "INSERT INTO docudag0 "
    
    SQL_GET_TS = "SELECT timestmp FROM docudag0 WHERE "
    
    SQL_GET_SEARCH_LIST = "SELECT DAGIDAEE, DAGIDAND, DAGIDDAG, DAGANNUL FROM DOCUDAG0 "
    
    mvarFieldsIndexes = Array("0", "1", "2", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Ente Emittente", "docudag0.dagidaee", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Anno Documento", "docudag0.dagidand", 0, True, 4, "1"
        .Add "Progressivo Documento", "docudag0.dagiddag", 0, True, 7, "2"
        .Add "Rag.Soc. Cliente", "docudag0.dagrasoc", 0, , 40, "3"
        .Add "Indir. Cliente", "docudag0.dagindir", 0, , 40, "4"
        .Add "Località Cliente", "docudag0.daglocal", 0, , 25, "5"
        .Add "Provincia Cliente", "docudag0.dagprovi", 0, , 2, "6"
        .Add "CAP Cliente", "docudag0.dagnrcap", 0, , 5, "7"
        .Add "Annullato", "docudag0.dagannul", 0, , , "8"
    End With
    With DS.KeyFields    ' CAMPI CHIAVE
        .Add "0", "1"    ' Ente Emittente
        .Add "1", "2"    ' Anno documento
        .Add "2", "3"    ' Progressivo documento
    End With
    DS.CancelField = "docudag0.dagannul"
End Sub
Private Sub Class_Terminate()
    Set DS = Nothing
End Sub
Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub


