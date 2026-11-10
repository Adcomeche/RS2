VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStoricoPrezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Classe CStoricoPrezzi
' Data Creazione 4/11/99
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
    SQL_FIND = " SELECT  ASZIDADZ, ASZIDAPR, ASZIDAIM, ASZIDAPD, " & _
               " ASZDTVAR, ASZPGVAR, ASZPZACQ, ASZPZVEN, ASZANNUL, " & _
               " TIMESTMP FROM prezasz0 WHERE "
               
    SQL_GETLIST = " SELECT ASZDTVAR, ASZPGVAR, ASZPZACQ, ASZPZVEN, ASZANNUL, " & _
                  " TIMESTMP FROM prezasz0 "
    
    SQL_UPDATE = "UPDATE prezasz0 SET "
    
    SQL_INSERT = "INSERT INTO prezasz0 "
    
    SQL_GET_TS = "SELECT timestmp FROM prezasz0 WHERE "
    
    SQL_GET_SEARCH_LIST = "SELECT ASZIDADZ , ASZIDAPR, ASZIDAIM, ASZIDAPD, ASZDTVAR, ASZPGVAR, ASZPZACQ, ASZPZVEN, ASZANNUL FROM prezasz0"
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "5", "6", "7", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    With DS.Fields
        .Add "Destinazione", "prezasz0.aszidadz", 0, True, 10, "0"
        With .Item("0")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Prodotto", "prezasz0.aszidapr", 0, True, 6, "1"
        With .Item("1")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "prezasz0.aszidaim", 0, True, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "prezasz0.aszidapd", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Data Variazione", "prezasz0.aszpzacq", 2, True, 10, "4"
        .Add "Progressivo Variazione", "prezasz0.aszpzven", 1, True, 1, "5"
        .Add "Prezzo di Acquisto", "prezasz0.aszpzacq", 1, True, 13, "6"
        .Add "Prezzo di Vendita", "prezasz0.aszpzven", 1, True, 13, "7"
        .Add "Annullato", "prezasz0.aszannul", 3, True, , "8"
    End With
    With DS.KeyFields    ' CAMPI CHIAVE
        .Add "0", "1"    ' Destinazione
        .Add "1", "2"    ' Prodotto
        .Add "2", "3"    ' Imballo
        .Add "3", "4"    ' Posizione Doganale
        .Add "4", "5"    ' Data variazione
        .Add "5", "6"    ' Progressivo variazione
    End With
    DS.CancelField = "prezasz0.aszannul"
End Sub
Private Sub Class_Terminate()
    Set DS = Nothing
End Sub
Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub


