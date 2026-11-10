VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStoricoSerbatoio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CStoricoSerbatoio
' Data Creazione 5/11/99
' Autore Fabio Mozzi

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
    
    SQL_GETLIST = ""
    SQL_FIND = " SELECT ASSIDADE, ASSIDASR, ASSIDASS, ASSDTDAT, ASSHRORA, " & _
               " ASSNRDEN, ASSNRDEA, ASSNRTEM, ASSANNUL, TIMESTMP FROM " & _
               " anagass0 WHERE "
    
    SQL_UPDATE = "UPDATE anagass0 SET "
    SQL_INSERT = "INSERT INTO anagass0 "
    SQL_GET_TS = "SELECT timestmp FROM anagass0 WHERE "
    
    SQL_GET_SEARCH_LIST = "SELECT ASSIDADE, ASSIDASR, ASSDTDAT,ASSHRORA, " & _
                          " ASSNRDEN, ASSNRDEA, ASSNRTEM,ASSANNUL FROM anagass0 "
    
    Set DS = New CDataStruct
    'Lista dei campi della istruzione SQL_GET_SEARCH_LIST
    mvarFieldsIndexes = Array("Deposito", "CodSerbatoio", "Data", "Ora", _
                               "Densita15", "DensitaAm", "Temperatura", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    'costruisce la collezione contenente la definizione di tutti campi della tabella
    
    With DS.Fields
        .Add "Deposito", "anagass0.assidade", 0, True, 3, "Deposito"
         With .Item("Deposito")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
         End With
        .Add "Codice Serbatoio", "anagass0.assidasr", 0, True, 3, "CodSerbatoio"
         With .Item("CodSerbatoio")
            .ForeignTableEntity = "ASR"
            .ForeignTable = "anagasr0"
            .ForeignTableFld = "asridasr"
            .ForeignTableDescField = "asrdescr"
            .ForeignTableDeleteFld = "asrannul"
         End With
        .Add "Progressivo", "anagass0.assidass", 0, True, 7, "Progressivo"
        .Add "Data", "anagass0.assdtdat", 2, , 10, "Data"
        .Add "Ora", "anagass0.asshrora", 4, , 5, "Ora"
        .Add "Densità 15°", "anagass0.assnrden", 1, , , "Densita15"
        .Add "Densità Ambiente", "anagass0.assnrdea", 1, , , "DensitaAm"
        .Add "Temperatura", "anagass0.assnrtem", 1, , , "Temperatura"
        .Add "Annullato", "anagass0.assannul", 3, , , "Annullato"
    End With
    
    'specifica il campo chiave
    
    With DS.KeyFields
        .Add "Deposito", "1"
        .Add "CodSerbatoio", "2"
        .Add "Progressivo", "3"
    End With
    
    'specifica il campo di annullamento logico
    DS.CancelField = "anagass0.assannul"
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

