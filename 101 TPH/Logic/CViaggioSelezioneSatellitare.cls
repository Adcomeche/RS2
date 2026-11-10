VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioSelezioneSatellitare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'03/09/2003  CZUCCHINI   modificato GetSelectionExport1
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
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



'Public Function GetSelectionCreatePPD(Optional TipoPrefattura As String = "A") As Boolean
'    ReDim varOrderFields(0 To 1) As Variant
'    With DS
'        ' --- Seleziono tutti gli scarichi dei viaggi completati
'        .Fields("10").FindValue = "C"
'        If TipoPrefattura = "P" Then
'            .Fields("10").FindValue(2) = "E"
'        End If
'        .Fields("18").FindValue = False
'        .Fields("19").FindValue = False
'        varOrderFields(0) = "2"
'        varOrderFields(1) = "20"
'        GetSelectionCreatePPD = .GetSelection(SQL_GET_SELECTION_CREATE_PPT, True, varOrderFields)
'    End With
'End Function

'Public Function GetSelectionExport1(Optional Presidio As String = vbNullString) As Boolean
'    Const SQL_GET_SELECTION_EXPORT_1 = _
'        " SELECT DISTINCT " & _
'                    " VVTIDVVT, VVTIDAUT, VVTSEQVI, VVTHRINV, " & _
'                    " VVTIDADE, VVTIDAEE, VVTDTINV, VVTIDDEF, " & _
'                    " VVTIDVE1, VVTIDAD1, VVTIDEE1, VVTKMPRE, " & _
'                    " VVTKMDEF, VVTIDAM1, ANAGATM0_1.ATMIDATM, VVTIDAM2, " & _
'                    " VVTHRINC, VVTIDAU1, ANAGAAU0.AAUNOMEA, ANAGAAU0.AAUNRCID, " & _
'                    " ANAGAAU0_1.AAUNOMEA, VVTIDAMD, OCTRIFE1 + ISNULL(OCRRIFE2,''), VVCIDADZ, " & _
'                    " OCTRIFE3, VVCTEMVG, left(OCTNOTA2, 6), substring(OCTNOTA2, 7, 6), " & _
'                    " right(OCTNOTA2, 5), OCTDTCRI, OCTFLPOR, VVPIDAPR, " & _
'                    " VVPQTRLT, OCRQCOLT, 'L', OCTRAUT4, VVCNRSEQ " & _
'        " FROM VIAGVVT0 INNER JOIN " & _
'                    " VIAGVVC0 ON VVCIDVVT = VVTIDVVT INNER JOIN " & _
'                    " ANAGAZO0 ON AZOIDAZO = VVCIDAZO INNER JOIN " & _
'                    " VIAGVVP0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC INNER JOIN " & _
'                    " ORDIOCT0 ON VVCIDOCT = OCTIDOCT INNER JOIN " & _
'                    " ORDIOCR0 ON OCTIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR INNER JOIN " & _
'                    " ANAGAAM0 ON VVTIDAM1 = ANAGAAM0.AAMIDAAM INNER JOIN " & _
'                    " ANAGATM0 ON ANAGAAM0.AAMIDATM = ANAGATM0.ATMIDATM LEFT OUTER JOIN " & _
'                    " ANAGAAM0 AS ANAGAAM0_1 ON VVTIDAM2 = ANAGAAM0_1.AAMIDAAM LEFT OUTER JOIN " & _
'                    " ANAGATM0 AS ANAGATM0_1 ON ANAGAAM0_1.AAMIDATM = ANAGATM0_1.ATMIDATM INNER JOIN " & _
'                    " ANAGAAU0 ON VVTIDAU1 = ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
'                    " ANAGAAU0 AS ANAGAAU0_1 ON VVTIDAU2 = ANAGAAU0_1.AAUIDAAU "
'
'    ReDim varOrderFields(1) As Variant
'    With DS
'        .Fields("10").FindValue = "D"
'        .Fields("10").FindValue(2) = "E"
'        .Fields("10").FindValue(3) = "C"
'
'        .Fields("90").FindValue = Presidio
'        '.Fields("91").FindValue = 0
'        '.Fields("92").FindValue = "RP"
'        varOrderFields(0) = "0"
'        varOrderFields(1) = "94"
'        GetSelectionExport1 = .GetSelection(SQL_GET_SELECTION_EXPORT_1, True, varOrderFields)
'    End With
'End Function

' Seleziona i viaggi allo scarico che devono essere modificati e trasmessi
'Public Function GetSelectionViaggiAlloScarico(Optional ActiveRecordsOnly As Boolean = False, _
'                                              Optional OrderFields As Variant, _
'                                              Optional MaxRecords As Long = 0) As Boolean
'    Const SQL_GET_SELECTION_VIAGGI_ALLO_SCARICO = _
'        "SELECT vvtidvvt, vvcidvvc, vvpidvvp, vvtdtinv, " & _
'        "       vvtidade, vvtidvvt, vvtiddef, vvtidaee, " & _
'        "       vvtseqvi, vvthrinv, vvthrinc, vvtidaut, " & _
'        "       vvtidam1, vvtidam2, vvtidau1, vvtkmdef, " & _
'        "       aaunrcid, sutcodso, vvcnrseq, OCTRIFE1 + ISNULL(OCRRIFE2,''), " & _
'        "       vvcidadz, vvcderas, vvctemvg, adzkmrag, " & _
'        "       vvckmpad, vvcflst5, vvcnote1, vvcnotee, " & _
'        "       vvpidapr, ocrqrilt, vvpqtrlt, vvpqcolt, " & _
'        "       vvpflst5, vvtidasc, octraut4, vvtnripa, " & _
'        "       vvcdtcon, vvtflst3, vvtflst2 " & _
'        " FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN anagade0 ON vvtidade = adeidade " & _
'        "               INNER JOIN VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN ANAGADZ0 ON VVCIDADZ = ADZIDADZ " & _
'        "               INNER JOIN ORDIOCR0 ON VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT AND VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR " & _
'        "               INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU INNER JOIN ORDIOCT0 ON ORDIOCR0.OCRIDOCT = ORDIOCT0.OCTIDOCT " & _
'        "               LEFT OUTER JOIN MULTSUT0 ON VVTIDADE = SUTIDASC AND VVTIDAUT = SUTIDAUT "
'    Dim strSQL As String
'
'    strSQL = SQL_GET_SELECTION_VIAGGI_ALLO_SCARICO
'    If Not IsMissing(OrderFields) Then
'        GetSelectionViaggiAlloScarico = DS.GetSelection(strSQL, ActiveRecordsOnly, OrderFields, MaxRecords)
'    Else
'        GetSelectionViaggiAlloScarico = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)
'    End If
'End Function

'Public Function GetSelectionExport(Optional Presidio As String = vbNullString) As Boolean
'    ReDim varOrderFields(0) As Variant
'    With DS
'        .Fields("10").FindValue = "D"
'        .Fields("90").FindValue = Presidio
'        .Fields("91").FindValue = 0
'        .Fields("92").FindValue = "RP"
'        varOrderFields(0) = "0"
'        GetSelectionExport = .GetSelection(SQL_GET_SELECTION_EXPORT, True, varOrderFields)
'    End With
'End Function

'Public Function GetSelectionExportExViagDetail(Optional FlagFirstTimeOnly As Boolean = True) As Boolean
'    ReDim varOrderFields(1) As Variant
'    With DS
''        .Fields("10").FindValue(1) = "I"
''        .Fields("10").FindValue(2) = "C"
'        .Fields("33").FindValue = ""
'        .Fields("92").FindValue = "GC"
'        If FlagFirstTimeOnly Then
'            .Fields("35").FindValue = "N"
'        Else
'            .Fields("35").FindValue = Empty
'        End If
'        varOrderFields(0) = "13"
'        varOrderFields(1) = "5"
'        GetSelectionExportExViagDetail = .GetSelection(SQL_GET_SELECTION_EXPORT_EXVIAGDETAIL, False, varOrderFields)
'    End With
'End Function


'Public Function GetSelectionExportDest() As Boolean
'    ReDim varOrderFields(0) As Variant
'    With DS
'        .Fields("10").FindValue = "D"
''        .Fields("16").FindValue = "T"
'        varOrderFields(0) = "0"
'        GetSelectionExportDest = .GetSelection(SQL_GET_SELECTION_EXPORT_DEST, True, varOrderFields)
'    End With
'End Function

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
                             
    Dim strSQL As String

    strSQL = SQL_GET_SELECTION

    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)
    End If
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

Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT vvtidvvt, vvtiddef, vvtseqvi, vvtidasc, " & _
               "       vvtdtinc, vvthrinc, vvtdtinv, vvthrinv, " & _
               "       vvtdtfiv, vvthrfiv, vvtidazo, vvtbamez, " & _
               "       vvtbaaut, vvtidatv, vvtidaut, vvtidam1, " & _
               "       vvtidam2, vvtidve1, vvtidve2, vvtidau1, " & _
               "       vvtidau2, vvtidaee, vvtidade, vvtidamd, " & _
               "       vvtidvv1, vvtkmpre, vvtflcap, vvtkmdef, " & _
               "       vvtflcad, vvttmpvi, vvttmpca, vvtflrid, " & _
               "       vvtidad1, vvtidee1, vvtidadz, vvtrife1, " & _
               "       vvtrife2, vvtflst1, vvtflst2, vvtflst3, " & _
               "       vvtflst4, vvtflst5, vvtflstv, vvtfltiv, " & _
               "       vvtannul, timestmp FROM viagvvt0 WHERE "
    SQL_GETLIST = "SELECT vvtidvvt, vvtidasc FROM viagvvt0 "

'    SQL_GET_SELECTION = "SELECT VVTDTINV, VVTIDVVT, VVTIDDEF, VVTSEQVI, VVTIDADE," & _
'                " VVTIDAUT, VVTIDAM1, VVTIDAM2, VVTIDAU1, ASCRASOC," & _
'                " VVTKMDEF, VVTNRIPA, ANAGAVV_A.AVVDESCR, VVTIDAEE, VVTHRINC, VVTTMPCA," & _
'                " VVTHRINV, AAUNRCID, SUTCODSO, VVTTMPVI, NULL, VVTFLSTA, ANAGAVV_B.AVVDESCR" & _
'                " FROM VIAGVVT0" & _
'                " INNER JOIN ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM AND AAMFLSAT = 'K'" & _
'                " INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU" & _
'                " INNER JOIN ANAGASC0 ON VIAGVVT0.VVTIDASC = ANAGASC0.ASCIDASC" & _
'                " LEFT OUTER JOIN MULTSUT0 ON VVTIDADE = SUTIDASC AND VVTIDAUT = SUTIDAUT" & _
'                " LEFT OUTER JOIN ANAGAVV ANAGAVV_A ON AVVIDTBL = 'VIAGVVT0' AND AVVIDFLD = 'VVTFLSTA' AND AVVVALUE = VVTFLSTA" & _
'                " LEFT OUTER JOIN KNOSVIA0 ON KNOSVIA0.NUMVIAJE = VVTNRKVI AND FECHA = (SELECT MAX(FECHA) FROM KNOSVIA0 KVIA WHERE KVIA.NUMVIAJE = KNOSVIA0.NUMVIAJE)" & _
'                " LEFT OUTER JOIN ANAGAVV ANAGAVV_B ON ANAGAVV_B.AVVIDTBL = 'KNOSVIA0' AND ANAGAVV_B.AVVIDFLD = 'ESTA' AND ANAGAVV_B.AVVVALUE = KNOSVIA0.ESTA"

    SQL_GET_SELECTION = "SELECT VVTDTINV, VVTIDVVT, VVTIDDEF, VVTSEQVI, VVTNRKVI, VVTIDADE," & _
                " NULL, VVTIDAUT, VVTIDAM1, VVTIDAM2, VVTIDAU1, ASCRASOC," & _
                " VVTKMDEF, VVTNRIPA, ANAGAVV_A.AVVDESCR, VVTIDAEE, VVTHRINC, VVTTMPCA," & _
                " VVTHRINV, AAUNRCID, SUTCODSO, VVTTMPVI, VVTFLSTA, KNOSVIA0.ESTA, ANAGAVV_B.AVVDESCR" & _
                " FROM VIAGVVT0" & _
                " INNER JOIN ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM AND AAMFLSAT = 'K'" & _
                " INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU" & _
                " INNER JOIN ANAGASC0 ON VIAGVVT0.VVTIDASC = ANAGASC0.ASCIDASC" & _
                " LEFT OUTER JOIN MULTSUT0 ON VVTIDADE = SUTIDASC AND VVTIDAUT = SUTIDAUT" & _
                " LEFT OUTER JOIN ANAGAVV ANAGAVV_A ON AVVIDTBL = 'VIAGVVT0' AND AVVIDFLD = 'VVTFLSTA' AND AVVVALUE = VVTFLSTA" & _
                " LEFT OUTER JOIN KNOSVIA0 ON KNOSVIA0.NUMVIAJE = VVTNRKVI AND FECHA = (SELECT MAX(FECHA) FROM KNOSVIA0 KVIA WHERE KVIA.NUMVIAJE = KNOSVIA0.NUMVIAJE)" & _
                " LEFT OUTER JOIN ANAGAVV ANAGAVV_B ON ANAGAVV_B.AVVIDTBL = 'KNOSVIA0' AND ANAGAVV_B.AVVIDFLD = 'ESTA' AND ANAGAVV_B.AVVVALUE = KNOSVIA0.ESTA"
    
    SQL_UPDATE = "UPDATE viagvvt0 SET "
    SQL_INSERT = "INSERT INTO viagvvt0 "
    SQL_GET_TS = "SELECT timestmp FROM viagvvt0 WHERE "
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Dal Numero Viaggio", "viagvvt0.vvtidvvt", 0, , 7, "0", " >= "
        .Add "Al Numero Viaggio", "viagvvt0.vvtidvvt", 0, , 7, "1", " <= "
        .Add "Numero Viaggio", "viagvvt0.vvtidvvt", 0, , 7, "2"
        .Add "Società", "viagvvt0.vvtidasc", 0, , 3, "3"
        With .Item("3")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "viagvvt0.vvtidaee", 0, , 3, "4"
        With .Item("4")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Dalla Data Inizio Viaggio", "viagvvt0.vvtdtinv", 2, , 10, "5", " >= "
        .Add "Alla Data Inizio Viaggio", "viagvvt0.vvtdtinv", 2, , 10, "6", " <= "
        .Add "Dalla Data Inizio Carico", "viagvvt0.vvtdtinc", 2, , 10, "97", " >= "
        .Add "Alla Data Inizio Carico", "viagvvt0.vvtdtinc", 2, , 10, "98", " <= "
        .Add "Deposito", "viagvvt0.vvtidade", 0, , 3, "7"
        With .Item("7")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Zona", "viagvvt0.vvtidazo", 0, , 3, "8"
        With .Item("8")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Tipo Viaggio", "viagvvt0.vvtfltiv", 0, , 1, "9"
        .Add "Stato Viaggio", "viagvvt0.vvtflstv", 0, , 1, "10"
        .Add "Dal Numero Viaggio Definitivo", "viagvvt0.vvtiddef", 0, , 7, "11", " >= "
        .Add "Al Numero Viaggio Definitivo", "viagvvt0.vvtiddef", 0, , 7, "12", " <= "
        .Add "Unità di Trasporto", "viagvvt0.vvtidaut", 0, , 3, "13"
        With .Item("13")
            .ForeignTableEntity = "AUT"
            .ForeignTable = "anagaut0"
            .ForeignTableFld = "autidaut"
            .ForeignTableDescField = "autdescr"
            .ForeignTableDeleteFld = "autannul"
        End With
        .Add "Motrice", "viagvvt0.vvtidam1", 0, , 10, "14"
        With .Item("14")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Rimorchio", "viagvvt0.vvtidam2", 0, , 10, "15"
        With .Item("15")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Flag stato 5", "viagvvt0.vvtflst5", 0, , 2, "16"
        .Add "Zona consegna", "viagvvc0.vvcidazo", 0, , 3, "17"
        .Add "Consegna annullata", "viagvvc0.vvcannul", 3, , , "18"
        .Add "Scarico annullato", "viagvvp0.vvpannul", 3, , , "19"
        .Add "Numero consegna", "viagvvc0.vvcidvvc", 0, , 3, "20"
        .Add "Vettore", "viagvvt0.vvtidve1", 0, False, 3, "32"
        With .Item("32")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Società di competenza", "viagvvt0.vvtidasc", 0, , 3, "33"
        With .Item("33")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Autista", "viagvvt0.vvtidau1", 0, , 3, "34"
        With .Item("34")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Flag stato 4", "viagvvt0.vvtflst4", 0, False, 2, "35"
        .Add "Autista 2", "viagvvt0.vvtidau2", 0, , 3, "36"
        With .Item("36")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Sequenza Viaggio", "viagvvt0.vvtseqvi", 0, True, 3, "38"
        .Add "Data inizio viaggio", "viagvvt0.vvtdtinv", 2, False, 10, "37"
        
        .Add "Presidio", "anagazo0.azocodma", 0, , 15, "90"
        .Add "Consegna Annullata", "viagvvc0.vvcannul", 3, , , "91"
        .Add "Raggruppamento Ordine", "ordioct0.octidaro", 0, , 3, "92"
        With .Item("92")
            .ForeignTableEntity = "ARO"
            .ForeignTable = "anagaro0"
            .ForeignTableFld = "aroidaro"
            .ForeignTableDescField = "arodescr"
            .ForeignTableDeleteFld = "aroannul"
        End With
        .Add "Presidio GC", "anagade0.adecodma", 0, , 15, "93"
        .Add "Numero Sequenza", "viagvvc0.vvcnrseq", 0, True, 3, "94"
        .Add "Prodotto", "viagvvp0.vvpidapr", 0, True, 6, "95", , True
        With .Item("95")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Progressivo Scarico", "viagvvp0.vvpidvvp", 0, False, 3, "96"
    End With
    DS.CancelField = "viagvvt0.vvtannul"
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
        .Add "4", "5"
        .Add "5", "6"
        .Add "6", "7"
        .Add "7", "8"
        .Add "8", "9"
        .Add "9", "10"
        .Add "10", "11"
        .Add "11", "12"
        .Add "12", "13"
        .Add "13", "14"
        .Add "14", "15"
        .Add "15", "16"
        .Add "16", "17"
        .Add "17", "18"
        .Add "18", "19"
        .Add "19", "20"
        .Add "20", "21"
        .Add "32", "22"
        .Add "33", "23"
        .Add "34", "24"
        .Add "90", "25"
        .Add "91", "26"
        .Add "92", "27"
        .Add "93", "28"
        .Add "35", "29"
        .Add "94", "30"
        .Add "97", "31"
        .Add "98", "32"
    End With
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






