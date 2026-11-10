VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaSoc"
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
Public SQL_FIND_CONSEGNA As String
Public SQL_FIND_CONSEGNA_2 As String
Public SQL_FIND_CONSEGNA_3 As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

Private Const SQL_DELETE_PREF_CEPSA = "DELETE FROM PREFGP10 " & _
                                      " WHERE GP1IDASC = '038' AND GP1FLTIR = 'P' AND MONTH(GP1DTCON) = ? AND YEAR(GP1DTCON) = ? AND GP1IDAEE = ? "

Private Const SQL_GET_PEDAGGIO = "SELECT SUM(GP1IMPPE) " & _
                                 "  FROM PREFGP10 WHERE GP1IDASC = '038' AND GP1RIFE1 = ? "
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


Public Function DeletePrefCEPSA(Anno As Variant, Mese As Variant, IDEnteEmittente As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim blnResult As Boolean
    
    If Not (IsNumeric(Anno) And IsNumeric(Mese)) Then
        GoTo ErrHandler
    End If
    strSQL = SQL_DELETE_PREF_CEPSA
    blnResult = DS.DA.ExecuteCommand(strSQL, Mese, Anno, IDEnteEmittente)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    DeletePrefCEPSA = True
    Exit Function
ErrHandler:
    DeletePrefCEPSA = False
End Function

Public Function GetPedaggioCEPSA(RifOrdine As Variant) As Variant
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    strSQL = SQL_GET_PEDAGGIO
    blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, RifOrdine)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetPedaggioCEPSA = varBuffer(0, 0)
    Exit Function
ErrHandler:
    GetPedaggioCEPSA = Empty
End Function


Public Function GetConsegnaCEPSA(RifOrdine As Variant, TipoConsegna As Variant, _
                                 ByRef IDVVT As String, ByRef IDVVC As String, ByRef IDVVP As String, _
                                 ByRef IDAUM As String, _
                                 ByRef IDOCT As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    Dim strRifOrdine As String
    Dim strRifRiga As String
    Dim strRifPartizione As String
    
    If Len(RifOrdine) = 0 Then
        GoTo ErrHandler
    End If
    strRifOrdine = RifOrdine
    Select Case TipoConsegna
        Case "910151", "910152", "910153"
        
            strSQL = SQL_FIND_CONSEGNA_3
            blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strRifOrdine)
            If IsEmpty(varBuffer) Then
                lngHowMany = 0
                strRifOrdine = Mid(RifOrdine, 3, 11)
                blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strRifOrdine)
            End If
            
            
        Case Else
            strRifOrdine = Left(RifOrdine, 10)
            strRifRiga = Mid(RifOrdine, 11, 3)
            strRifPartizione = Mid(RifOrdine, 14, 2)
    
    
            If (strRifPartizione = "00") Or (Len(strRifPartizione) = 0) Then
                strSQL = SQL_FIND_CONSEGNA_2
                blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strRifOrdine, strRifRiga)
                If IsEmpty(varBuffer) Then
                    strSQL = SQL_FIND_CONSEGNA
                    blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strRifOrdine, strRifRiga, strRifPartizione)
                End If
            Else
                If Not IsNumeric(strRifPartizione) Then
                    GoTo ErrHandler
                End If
                strRifPartizione = CStr(CInt(strRifPartizione))
                strSQL = SQL_FIND_CONSEGNA
                blnResult = DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strRifOrdine, strRifRiga, strRifPartizione)
            End If
    End Select
    
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    
    IDVVT = varBuffer(0, 0)
    IDVVC = varBuffer(1, 0)
    IDVVP = varBuffer(2, 0)
    If IsNull(varBuffer(3, 0)) Then
        IDAUM = "L"
    ElseIf varBuffer(3, 0) = vbNullString Then
        varBuffer(3, 0) = "L"
    Else
        IDAUM = varBuffer(3, 0)
    End If
    IDOCT = varBuffer(4, 0)
    
    GetConsegnaCEPSA = True
    Exit Function
ErrHandler:
    IDVVT = vbNullString
    IDVVC = vbNullString
    IDVVP = vbNullString
    IDAUM = vbNullString
    IDOCT = vbNullString
    GetConsegnaCEPSA = False
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
               "       gp1imppe, gp1fltir, gp1idam1, gp1idam2, " & _
               "       gp1idau1, gp1rifat, gp1impoe, gp1impbe, " & _
               "       gp1imphe, gp1impse, gp1idaee, gp1imsex, " & _
               "       gp1annul, timestmp FROM prefgp10 WHERE "
    SQL_GETLIST = "" '"SELECT ampidamp, ampdescr FROM prefgp10 "
    SQL_UPDATE = "UPDATE prefgp10 SET "
    SQL_INSERT = "INSERT INTO prefgp10 "
    SQL_GET_TS = "SELECT timestmp FROM prefgp10 WHERE "
    SQL_CHECK_SOC_CODE = "SELECT Count(*) FROM prefgp10 WHERE gp1rife1 = ? AND gp1idapr = ?"
    SQL_DELETE = "DELETE FROM prefgp10 WHERE gp1rife1 = ? AND gp1idapr = ?"
    
    SQL_FIND_CONSEGNA = "SELECT VVPIDVVT, VVPIDVVC, VVPIDVVP, OCRIDAUM, VVCIDOCT " & _
                        "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT " & _
                        "       INNER JOIN VIAGVVP0 ON VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC " & _
                        "       INNER JOIN ORDIOCR0 ON VVPIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR " & _
                        "       INNER JOIN ORDIOCT0 ON OCTIDOCT = OCRIDOCT " & _
                        " WHERE OCTRIFE1 = ? AND OCRRIFE2 = ? AND CONVERT(int, OCTRAUT4) = ? AND VVTANNUL = 0 AND VVCANNUL = 0 AND VVPANNUL = 0 AND OCTANNUL = 0 AND OCRANNUL = 0"
    SQL_FIND_CONSEGNA_2 = "SELECT VVPIDVVT, VVPIDVVC, VVPIDVVP, OCRIDAUM, VVCIDOCT " & _
                        "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT " & _
                        "       INNER JOIN VIAGVVP0 ON VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC " & _
                        "       INNER JOIN ORDIOCR0 ON VVPIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR " & _
                        "       INNER JOIN ORDIOCT0 ON OCTIDOCT = OCRIDOCT " & _
                        " WHERE OCTRIFE1 = ? AND OCRRIFE2 = ? AND ((OCTRAUT4 IS NULL) OR (OCTRAUT4 = '')) AND VVTANNUL = 0 AND VVCANNUL = 0 AND VVPANNUL = 0 AND OCTANNUL = 0 AND OCRANNUL = 0"
    SQL_FIND_CONSEGNA_3 = "SELECT VVPIDVVT, VVPIDVVC, VVPIDVVP, OCRIDAUM, VVCIDOCT " & _
                        "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT " & _
                        "       INNER JOIN VIAGVVP0 ON VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC " & _
                        "       INNER JOIN ORDIOCR0 ON VVPIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR " & _
                        "       INNER JOIN ORDIOCT0 ON OCTIDOCT = OCRIDOCT " & _
                        " WHERE OCTRIFE1 = ? AND VVTANNUL = 0 AND VVCANNUL = 0 AND VVPANNUL = 0 AND OCTANNUL = 0 AND OCRANNUL = 0"
    
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
        .Add "Destinazione", "prefgp10.gp1idadz", 0, , 10, "10"
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
        .Add "Motrice", "prefgp10.gp1idam1", 0, False, 10, "39"
        With .Item("39")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Rimorchio", "prefgp10.gp1idam2", 0, False, 10, "40"
        With .Item("40")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Autista", "prefgp10.gp1idau1", 0, False, 3, "41"
        With .Item("41")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Riferimento Fattura", "prefgp10.gp1rifat", 0, , 25, "42"
        .Add "Importo Servizio Operario", "prefgp10.gp1impoe", 1, False, 11, "43"
        .Add "Importo Servizio Buques", "prefgp10.gp1impbe", 1, False, 11, "44"
        .Add "Ore Servizi Speciali", "prefgp10.gp1imphe", 1, False, 6, "45"
        .Add "Importo Servizi Speciali", "prefgp10.gp1impse", 1, False, 11, "46"
        .Add "Ente Emittente", "prefgp10.gp1idaee", 0, True, 3, "47"
        With .Item("47")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        
        '-- DP 20/08/08
        .Add "Servicios Exentos", "prefgp10.gp1imsex", 1, False, 13, "48"
        
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
