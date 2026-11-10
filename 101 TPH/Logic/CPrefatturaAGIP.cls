VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaAGIP"
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
Public SQL_CHECK_SOC_CODE As String
Public SQL_DELETE As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

Public Function CheckSocCode(OrdCode As String, IDProdotto As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    CheckSocCode = True
    lngHowMany = 1
    strSQL = SQL_CHECK_SOC_CODE
    blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, OrdCode, IDProdotto)
    If blnResult Then
        CheckSocCode = (varBuffer(0, 0) > 0)
    End If
    Exit Function
ErrHandler:
End Function

Public Function Delete(OrdCode As String, IDProdotto As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim blnResult As Boolean
    
    strSQL = SQL_DELETE
    blnResult = DS.DA.ExecuteCommand(strSQL, OrdCode, IDProdotto)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    Delete = True
    Exit Function
ErrHandler:
    Delete = False
End Function


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
    SQL_FIND = "SELECT gp1idazo, gp1idade, gp1rife1, gp1dtcon, " & _
               "       gp1bosca, gp1qtsca, gp1rife2, gp1tarap, " & _
               "       gp1impor, gp1idapr, gp1idadz, gp1rasde, " & _
               "       gp1idaut, gp1qkmra, gp1flimf, gp1rife3, " & _
               "       gp1annoe, gp1idasc, gp1riave, gp1fltis, " & _
               "       gp1riade, gp1riaut, gp1ridut, gp1riam1, " & _
               "       gp1riam2, gp1riaau, gp1niaau, gp1nrvia, " & _
               "       gp1riope, gp1rifor, gp1rioct, gp1rivvc, " & _
               "       gp1ritik, gp1nomeb, gp1qtskg, gp1fltit, " & _
               "       gp1imppe, gp1fltir, gp1annul, timestmp FROM prefgp10 WHERE "
    SQL_GETLIST = "" '"SELECT ampidamp, ampdescr FROM prefgp10 "
    SQL_UPDATE = "UPDATE prefgp10 SET "
    SQL_INSERT = "INSERT INTO prefgp10 "
    SQL_GET_TS = "SELECT timestmp FROM prefgp10 WHERE "
    SQL_CHECK_SOC_CODE = "SELECT Count(*) FROM prefgp10 WHERE gp1rife1 = ? AND gp1idapr = ?"
    SQL_DELETE = "DELETE FROM prefgp10 WHERE gp1rife1 = ? AND gp1idapr = ?"
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Zona", "prefgp10.gp1idazo", 0, , 3, "0"
        With .Item("0")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Deposito", "prefgp10.gp1idade", 0, , 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Riferimento Esterno 1", "prefgp10.gp1rife1", 0, , 15, "2"
        .Add "Data Consegna", "prefgp10.gp1dtcon", 2, , 10, "3"
        .Add "Bolletta di scarico", "prefgp10.gp1bosca", 0, , 15, "4"
        .Add "Quantità scaricata", "prefgp10.gp1qtsca", 1, , 13, "5"
        .Add "Riferimento Esterno 2", "prefgp10.gp1rife2", 0, , 8, "6"
        .Add "Tariffa applicata", "prefgp10.gp1tarap", 1, , 13, "7"
        .Add "Importo", "prefgp10.gp1impor", 1, , 13, "8"
        .Add "Prodotto", "prefgp10.gp1idapr", 0, , 6, "9"
        With .Item("9")
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Destinazione", "prefgp10.gp1idadz", 0, , 8, "10"
        With .Item("10")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Ragione Sociale Dest.", "prefgp10.gp1rasde", 0, , 23, "11"
        .Add "Unità di Trasporto", "prefgp10.gp1idaut", 0, , 3, "12"
        With .Item("12")
            .ForeignTableEntity = "AUT"
            .ForeignTable = "anagaut0"
            .ForeignTableFld = "autidaut"
            .ForeignTableDescField = "autdescr"
            .ForeignTableDeleteFld = "autannul"
        End With
        .Add "KM Raggio", "prefgp10.gp1qkmra", 0, , 3, "13"
        .Add "Importo forzato", "prefgp10.gp1flimf", 3, , , "14"
        .Add "Riferimento esterno 3", "prefgp10.gp1rife3", 0, , 15, "15"
        .Add "Anno prefattura", "prefgp10.gp1annoe", 1, , 4, "16"
        
        
        .Add "Società", "prefgp10.gp1idasc", 0, True, 3, "18"
        With .Item("18")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Riferimento Vettore", "prefgp10.gp1riave", 0, , 6, "19"
        .Add "Tipo servizio", "prefgp10.gp1fltis", 0, , 2, "20"
        .Add "Riferimento Deposito", "prefgp10.gp1riade", 0, , 4, "21"
        .Add "Riferimento UdT Cepsa", "prefgp10.gp1riaut", 0, , 12, "22"
        .Add "Riferimento UdT Deposito", "prefgp10.gp1ridut", 0, , 3, "23"
        .Add "Targa Motrice", "prefgp10.gp1riam1", 0, , 15, "24"
        .Add "Targa Rimorchio", "prefgp10.gp1riam2", 0, , 15, "25"
        .Add "Riferimento Autista Cepsa", "prefgp10.gp1riaau", 0, , 7, "26"
        .Add "NIF Autista", "prefgp10.gp1niaau", 0, , 14, "27"
        .Add "Numero Viaggio", "prefgp10.gp1nrvia", 0, , 5, "28"
        .Add "Codice operatore", "prefgp10.gp1riope", 0, , 6, "29"
        .Add "Codice fornitore", "prefgp10.gp1rifor", 0, , 6, "30"
        .Add "Riferimento ordine Cepsa", "prefgp10.gp1rioct", 0, , 16, "31"
        .Add "Riferimento consegna Cepsa", "prefgp10.gp1rivvc", 0, , 12, "32"
        .Add "Riferimento Ticket", "prefgp10.gp1ritik", 0, , 15, "33"
        .Add "Nome imbarcazione", "prefgp10.gp1nomeb", 0, , 30, "34"
        .Add "Quantità scaricata Kg", "prefgp10.gp1qtskg", 1, , 13, "35"
        .Add "Tipo tariffa", "prefgp10.gp1fltit", 0, , 2, "36"
        .Add "Importo pedaggi", "prefgp10.gp1imppe", 1, , 13, "37"
        .Add "Tipo riga", "prefgp10.gp1fltir", 0, , 2, "38"
        
        .Add "Annullato", "prefgp10.gp1annul", 3, , , "17"
    End With
    With DS.KeyFields
        .Add "2", "1"
        .Add "9", "2"
    End With
    DS.CancelField = "prefgp10.gp1annul"
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


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property


