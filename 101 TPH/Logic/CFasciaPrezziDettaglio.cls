VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFasciaPrezziDettaglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CFasciaPrezziDettaglio
' Data Creazione 5/11/99
' Autore Fabio Mozzi

Option Explicit

Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String



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
    SQL_FIND = " SELECT  FPDIDFPZ,FPDIDAPR, FPDIDAIM, FPDIDAPD, " & _
               " FPDPZACQ, FPDPZVEN, FPDNRSC1, FPDNRSC2, FPDNRSC3, " & _
               " FPDNRSV1, FPDNRSV2, FPDNRSV3, FPDNRMV1, FPDNRMV2, " & _
               " FPDNRMV3, FPDANNUL, TIMESTMP FROM prezfpd0 WHERE "
               
    SQL_GETLIST = " SELECT FPDIDAPR, FPDIDAIM, FPDIDAPD, " & _
                  " FPDPZACQ, FPDPZVEN, FPDANNUL, TIMESTMP FROM prezfpd0  "
    
    SQL_UPDATE = "UPDATE prezfpd0 SET "
    
    SQL_INSERT = "INSERT INTO prezfpd0 "
    
    SQL_GET_TS = "SELECT timestmp FROM prezfpd0 WHERE "
    
    
    SQL_GET_SEARCH_LIST = "SELECT FPDIDFPZ, FPDIDAPR, FPDIDAIM, FPDIDAPD, FPDPZACQ, FPDPZVEN FPDANNUL FROM prezfpd0"
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "5", "15")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Fascia di Prezzo", "prezfpd0.fpdidfpz", 0, True, 7, "0"
        With .Item("0")
            .ForeignTableEntity = "FPZ"
            .ForeignTable = "prezfpz0"
            .ForeignTableFld = "fpzidfpz"
            .ForeignTableDescField = "fpzdescr"
            .ForeignTableDeleteFld = "fpzannul"
        End With
        .Add "Prodotto", "prezfpd0.fpdidapr", 0, True, 6, "1"
        With .Item("1")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "prezfpd0.fpdidaim", 0, True, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "prezfpd0.fpdidapd", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        
        .Add "Prezzo di Acquisto", "prezfpd0.fpdpzacq", 1, True, 13, "4"
        .Add "Prezzo di Vendita", "prezfpd0.fpdpzven", 1, True, 13, "5"
        .Add "Sconto 1", "prezfpd0.fpdnrsc1", 1, True, 7, "6"
        .Add "Sconto 2", "prezfpd0.fpdnrsc2", 1, True, 7, "7"
        .Add "Sconto 3", "prezfpd0.fpdnrsc3", 1, True, 7, "8"
        .Add "Sconto 1 Valore", "prezfpd0.fpdnrsv1", 1, True, 13, "9"
        .Add "Sconto 2 Valore", "prezfpd0.fpdnrsv2", 1, True, 13, "10"
        .Add "Sconto 3 Valore", "prezfpd0.fpdnrsv3", 1, True, 13, "11"
        .Add "Margine 1 Valore", "prezfpd0.fpdnrmv1", 1, True, 13, "12"
        .Add "Margine 2 Valore", "prezfpd0.fpdnrmv2", 1, True, 13, "13"
        .Add "Margine 3 Valore", "prezfpd0.fpdnrmv3", 1, True, 13, "14"
        
        .Add "Annullato", "prezfpd0.fpdannul", 3, True, , "15"
    End With
    With DS.KeyFields    ' CAMPI CHIAVE
        .Add "0", "1"    ' Fascia Prezzi
        .Add "1", "2"    ' Prodotto
        .Add "2", "3"    ' Imballo
        .Add "3", "4"    ' Posizione Doganale
    End With
    DS.CancelField = "prezfpd0.fpdannul"
End Sub
Private Sub Class_Terminate()
    Set DS = Nothing
End Sub
Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub


