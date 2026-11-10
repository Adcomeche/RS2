VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe CPrezzi
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
    SQL_FIND = " SELECT APZIDADZ, APZIDAPR, APZIDAIM, APZIDAPD, " & _
               " APZIDAIV, APZIDAMP, APZIDFPZ, APZPZACQ, APZPZVEN, " & _
               " APZIDAU1, APZIDAU2, APZQTCIS, APZCACIS, APZQTMIN, " & _
               " APZQTDRO, APZNRGRF, APZFLSOS, APZNRSC1, APZNRSC2, " & _
               " APZNRSC3, APZNRSV1, APZNRSV2, APZNRSV3, APZNRMV1, " & _
               " APZNRMV2, APZNRMV3, APZANNUL, TIMESTMP FROM prezapz0 WHERE "
               
    SQL_GETLIST = "SELECT APZIDAPR, APZIDAIM, APZIDAPD,APZPZACQ,APZIDAMP,APZIDAIV,APZNRSC1, APZNRSC2,APZNRSC3,APZANNUL FROM prezapz0 "
    
    SQL_UPDATE = "UPDATE prezapz0 SET "
    
    SQL_INSERT = "INSERT INTO prezapz0 "
    
    SQL_GET_TS = "SELECT timestmp FROM prezapz0 WHERE "
    
    
    SQL_GET_SEARCH_LIST = "SELECT APZIDADZ, APZIDAPR, APZIDAIM, APZIDAPD, APZIDAIV, APZIDAMP, APZIDFPZ, APZANNUL FROM prezapz0"
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "5", "6", "26")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Destinazione", "prezapz0.apzidadz", 0, True, 10, "0"
        With .Item("0")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Prodotto", "prezapz0.apzidapr", 0, True, 6, "1"
        With .Item("1")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "prezapz0.apzidaim", 0, True, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "prezapz0.apzidapd", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Aliquota IVA", "prezapz0.apzidaiv", 0, True, 3, "4"
        With .Item("4")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Modalità di Pagamento", "prezapz0.apzidamp", 0, True, 3, "5"
        With .Item("5")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "Fascia di Prezzo", "prezapz0.apzidfpz", 0, False, 7, "6"
        With .Item("6")
            .ForeignTableEntity = "FPZ"
            .ForeignTable = "prezfpz0"
            .ForeignTableFld = "fpzidfpz"
            .ForeignTableDescField = "fpzdescr"
            .ForeignTableDeleteFld = "fpzannul"
        End With
        .Add "Prezzo di Acquisto", "prezapz0.apzpzacq", 1, False, 13, "7"
        .Add "Prezzo di Vendita", "prezapz0.apzpzven", 1, False, 13, "8"
        
        .Add "UM Acquisto", "prezapz0.apzidau1", 0, False, 3, "9"
        With .Item("9")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "UM Vendita", "prezapz0.apzidau2", 0, False, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Q.tà Cisterne", "prezapz0.apzqtcis", 1, False, 3, "11"
        .Add "Capacità Cisterne", "prezapz0.apzcacis", 1, False, 8, "12"
        .Add "Q.tà Minima", "prezapz0.apzqtmin", 1, False, 8, "13"
        .Add "Q.tà Drop Medio", "prezapz0.apzqtdro", 1, False, 8, "14"
        .Add "Giorni Rifornimento", "prezapz0.apznrgrf", 1, False, 7, "15"
        
        .Add "Flag Sospensione", "prezapz0.apzflsos", 3, False, 7, "16"
        
        .Add "Sconto 1", "prezapz0.apznrsc1", 1, False, 7, "17"
        .Add "Sconto 2", "prezapz0.apznrsc2", 1, False, 7, "18"
        .Add "Sconto 3", "prezapz0.apznrsc3", 1, False, 7, "19"
        .Add "Sconto 1 Valore", "prezapz0.apznrsv1", 1, False, 13, "20"
        .Add "Sconto 2 Valore", "prezapz0.apznrsv2", 1, False, 13, "21"
        .Add "Sconto 3 Valore", "prezapz0.apznrsv3", 1, False, 13, "22"
        .Add "Margine 1 Valore", "prezapz0.apznrmv1", 1, False, 13, "23"
        .Add "Margine 2 Valore", "prezapz0.apznrmv2", 1, False, 13, "24"
        .Add "Margine 3 Valore", "prezapz0.apznrmv3", 1, False, 13, "25"
        
        .Add "Annullato", "prezapz0.apzannul", 3, False, , "26"
    End With
    With DS.KeyFields ' CAMPI CHIAVE
        .Add "0", "0"    ' Destinazione
        .Add "1", "1"    ' Prodotto
        .Add "2", "2"    ' Imballo
        .Add "3", "3"    ' Posizione Doganale
    End With
    DS.CancelField = "prezapz0.apzannul"
End Sub
Private Sub Class_Terminate()
    Set DS = Nothing
End Sub
Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub


