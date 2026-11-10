VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioSelezione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'03/09/2003  CZUCCHINI   modificato GetSelectionExport1
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
Public SQL_GET_SELECTION_GC As String
Public SQL_GET_SELECTION_EXPORT As String
Public SQL_GET_SELECTION_EXPORT_DEST As String
Public SQL_GET_SELECTION_CREATE_PPT As String
Public SQL_GET_SELECTION_CREATE_PPT_CONTRATTO As String
Public SQL_SELECT_TRIP_PRODUCTS As String
Public SQL_INSERT_TRIP_PRODUCTS As String
Public SQL_CLEAR_TRIP_PRODUCTS As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_GET_SELECTION_CREATE_SDT As String
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)



Private Const SQL_GET_SELECTION_EXPORT_EXVIAGDETAIL = _
"SELECT DISTINCT VVTIDVVT, VVTDTINV, VVTIDDEF , VVTKMDEF, " & _
      "VVTIDASC , VVTIDAUT, VVTIDAM1 , VVTIDAM2, " & _
      "VVTIDAEE , VVTFLST4 " & _
      "FROM VIAGVVC0 INNER JOIN VIAGVVT0 ON VIAGVVC0.VVCIDVVT = VIAGVVT0.VVTIDVVT INNER JOIN " & _
      "     ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT RIGHT OUTER JOIN " & _
      "     ANAGAVE0 ON VIAGVVT0.VVTIDVE2 = ANAGAVE0.AVEIDAVE "


Public Function DEUSTO_InsertViaggio(varIDViaggio As Variant, Service As VLocalize.VLocalize, ByRef strError As String) As Boolean

    Const SQL_DEUSTO_GET_DATA_VIAGGIO = "SELECT VVTIDDEF, VVTIDAEE, VVTDTINC, VVTHRINC, " & _
                                        "       VVTDTFIV, VVTHRFIV, VVTDTINV, VVTHRINV, " & _
                                        "       VVTIDADE, VVCIDVVC, VVCIDADZ, OCTRIFE1, " & _
                                        "       VVCDTCON, VVCTEMVG, VVPIDAPR, VVPQTRLT, " & _
                                        "       AAMIDDEU, VVTIDAM1 " & _
                                        "  FROM VIAGVVT0 INNER JOIN " & _
                                        "       VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN " & _
                                        "       VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN " & _
                                        "       ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM INNER JOIN ORDIOCT0 ON VVCIDOCT = OCTIDOCT " & _
                                        " WHERE (VIAGVVT0.VVTIDVVT = ?) AND (VIAGVVC0.VVCANNUL = 0) AND (VIAGVVP0.VVPANNUL = 0) " & _
                                        " ORDER BY VVCIDVVC, VVPIDAPR "

    Dim SQLStatement As String
    Dim strInsertDateTime As String
    Dim varBuffer As Variant
    Dim varResults As Variant
    Dim strData As String
    Dim strProdDesc As String
    Dim strDataOraConsegna As String
    Dim strDataInizioConsegna As String
    Dim strOraInizioConsegna As String
    Dim strDataFineConsegna As String
    Dim stroraFineConsegna As String
    Dim lngHowMany As Long
    Dim lngNumConsegne As Long
    Dim lngCountTask As Long
    Dim i As Long
    Dim j As Long
    Dim blnNewTask As Boolean

    SQLStatement = SQL_DEUSTO_GET_DATA_VIAGGIO
    If Not DS.DA.QueryCommand(SQLStatement, varBuffer, lngHowMany, varIDViaggio) Then
        strError = "Si è verificato un errore nella lettura dei dati del viaggio " & varIDViaggio
        GoTo ErrHandler
    End If

    If IsEmpty(varBuffer) Then
        strError = "Si è verificato un errore nella lettura dei dati del viaggio " & varIDViaggio
        GoTo ErrHandler
    End If

    If Not IsNumeric(varBuffer(16, 0)) Then
        DEUSTO_InsertViaggio = True
        strError = "Codice Satellitare assente sull'automezzo " & varBuffer(15, 0)
        GoTo EndProc
    End If

    ' Determiniamo il numero di consegne del viaggio e impostiamo la struttura per i dati
    ' da passare al servizio
    lngNumConsegne = 1
    For i = 0 To UBound(varBuffer, 2)
        If i > 0 Then
            If varBuffer(9, i) <> varBuffer(9, i - 1) Then
                lngNumConsegne = lngNumConsegne + 1
            End If
        End If
    Next i
    ReDim varobjDataInput(0 To lngNumConsegne, 0 To 19) As Variant

    If IsNull(varBuffer(0, 0)) Or Len(varBuffer(0, 0)) = 0 Then
        strError = "Il viaggio " & varIDViaggio & " è provvisorio, impossibile trasferirlo"
        GoTo ErrHandler
    End If

    ' Inseriamo il task di carico
    varobjDataInput(0, 0) = "1"
    varobjDataInput(0, 1) = varBuffer(0, 0)
    Select Case varBuffer(1, 0)
        Case "004"
            varobjDataInput(0, 2) = "3"
        Case "624"
            varobjDataInput(0, 2) = "2"
        Case Else
            GoTo ErrHandler
    End Select
    varobjDataInput(0, 3) = "Viaje " & varBuffer(0, 0)
    varobjDataInput(0, 4) = " "
    varobjDataInput(0, 5) = Format(varBuffer(2, 0), "dd/mm/yyyy") & " " & Format(varBuffer(3, 0), "hh:nn:ss")
    varobjDataInput(0, 6) = Format(varBuffer(4, 0), "dd/mm/yyyy") & " " & Format(varBuffer(5, 0), "hh:nn:ss")
    varobjDataInput(0, 7) = varBuffer(16, 0)
    varobjDataInput(0, 8) = "Carga"
    varobjDataInput(0, 9) = "S"
    varobjDataInput(0, 10) = Format(varBuffer(2, 0), "dd/mm/yyyy") & " " & Format(varBuffer(3, 0), "hh:nn:ss")
    varobjDataInput(0, 11) = Format(varBuffer(6, 0), "dd/mm/yyyy") & " " & Format(varBuffer(7, 0), "hh:nn:ss")
    varobjDataInput(0, 12) = varBuffer(8, 0)
    For i = 13 To 19
        varobjDataInput(0, i) = " "
    Next i

    lngCountTask = 1
    For i = 0 To UBound(varBuffer, 2)
        strProdDesc = strProdDesc & varBuffer(14, i) & " " & varBuffer(15, i) & "; "
        blnNewTask = False
        If i = UBound(varBuffer, 2) Then
            blnNewTask = True
        ElseIf (varBuffer(9, i) <> varBuffer(9, i + 1)) Then
            blnNewTask = True
        End If
        If blnNewTask Then
            ' Inseriamo la consegna
            For j = 0 To 7
                varobjDataInput(lngCountTask, j) = varobjDataInput(0, j)
            Next j
            varobjDataInput(lngCountTask, 8) = "Entrega " & lngCountTask
            varobjDataInput(lngCountTask, 9) = "N"
            varobjDataInput(lngCountTask, 10) = " "
            varobjDataInput(lngCountTask, 11) = " "
            varobjDataInput(lngCountTask, 12) = varBuffer(10, i)
            varobjDataInput(lngCountTask, 13) = IIf(IsNull(varBuffer(11, i)), " ", varBuffer(11, i))
            varobjDataInput(lngCountTask, 14) = strProdDesc
            varobjDataInput(lngCountTask, 15) = varIDViaggio & varBuffer(9, i)
            strDataOraConsegna = Format(varBuffer(12, i), "dd/mm/yyyy") & " " & Format(varBuffer(13, i), "hh:nn:ss")

            strDataInizioConsegna = Replace(DateAdd("h", -5, strDataOraConsegna), ".", ":")
            strOraInizioConsegna = Right(strDataInizioConsegna, 8)
            strDataInizioConsegna = Left(strDataInizioConsegna, 10)

            strDataFineConsegna = Replace(DateAdd("h", 5, strDataOraConsegna), ".", ":")
            stroraFineConsegna = Right(strDataFineConsegna, 8)
            strDataFineConsegna = Left(strDataFineConsegna, 10)

            varobjDataInput(lngCountTask, 16) = strDataInizioConsegna
            varobjDataInput(lngCountTask, 17) = strOraInizioConsegna
            varobjDataInput(lngCountTask, 18) = strDataFineConsegna
            varobjDataInput(lngCountTask, 19) = stroraFineConsegna

            strProdDesc = vbNullString
            lngCountTask = lngCountTask + 1
        End If
    Next i

    If Not Service.isOpen Then
        Service.Open
    End If
    If Not Service.CreaViaggio(varobjDataInput, varResults) Then
        strError = "Si è verificato un errore durante la creazione del servizio"
        GoTo ErrHandler
    End If

    DEUSTO_InsertViaggio = True
    GoTo EndProc

ErrHandler:
    If Len(strError) = 0 Then
        strError = "Si è verificato un errore sconosciuto"
    End If
    DEUSTO_InsertViaggio = False
EndProc:

End Function

Public Function GetElencoNoteConsegna(IDViaggio As String, _
                                      RefVR As String) As Variant
    
    On Error GoTo ErrHandler
    
    Const SQL_GET_SELECTION = _
        " SELECT DISTINCT " & _
        "        OCTRIFE1 + ISNULL(OCRRIFE2, ''), OCTRAUT4, SDZCODSO, SPRCODSO, " & _
        "        VVPQTRLT, VVPQCOLT, VVTDTINV, VVCTEMVG " & _
        "   FROM VIAGVVT0 INNER JOIN " & _
        "        VIAGVVC0 ON VVCIDVVT = VVTIDVVT INNER JOIN " & _
        "        VIAGVVP0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC INNER JOIN " & _
        "        ORDIOCR0 ON VVPIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR INNER JOIN " & _
        "        ORDIOCT0 ON OCRIDOCT = OCTIDOCT INNER JOIN " & _
        "        MULTSDZ0 ON VVCIDADZ = SDZIDADZ INNER JOIN " & _
        "        MULTSPR0 ON OCRIDAPR = SPRIDAPR " & _
        "  WHERE (OCTFLVER = '1') AND (VVCANNUL = 0) AND (VVPANNUL = 0) AND (VVTIDVVT = ?) AND " & _
        "        (VVPNRVER = ?) AND (SDZIDASC = '901') AND (SPRIDASC = '901') "

    Dim strQ As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    strQ = SQL_GET_SELECTION
    If Not DS.DA.QueryCommand(strQ, varBuffer, lngHowMany, IDViaggio, RefVR) Then
        GoTo ErrHandler
    End If
    GetElencoNoteConsegna = varBuffer
    Exit Function
ErrHandler:
    GetElencoNoteConsegna = Empty
End Function

Public Function GetSelectionExportNEN(IDSocietà As String, _
                                      IDEnteEmittente As String, _
                                      DallaData As Variant, _
                                      AllaData As Variant, _
                                      Optional DaNumeroViaggioDef As Variant, _
                                      Optional ANumeroViaggioDef As Variant) As Variant
    
    On Error GoTo ErrHandler
    
    Const SQL_GET_SELECTION_EXPORT_1 = _
        " SELECT DISTINCT " & _
        "        VVTIDVVT, VVTDTINV, VVTIDDEF, OCRNRTCK, " & _
        "        VVPNRVER, VCAQCALT, COUNT(VVPIDVVC), VVTFLSTV " & _
        "   FROM VIAGVVT0 INNER JOIN " & _
        "        VIAGVVC0 ON VVCIDVVT = VVTIDVVT INNER JOIN " & _
        "        VIAGVVP0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC INNER JOIN " & _
        "        ORDIOCR0 ON VVPIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR INNER JOIN " & _
        "        ORDIOCT0 ON OCRIDOCT = OCTIDOCT LEFT OUTER JOIN " & _
        "        VIAGVCA0 ON VVPIDAPR = VCAIDAPR AND VVTIDVVT = VCAIDVVT " & _
        "  WHERE (OCTFLVER = '1') AND (VVTANNUL = 0) AND (VVCANNUL = 0) AND (VVPANNUL = 0) AND " & _
        "        (VVTFLSTV <> 'P') AND (VVTDTINV BETWEEN ? AND ?) AND (VVTIDASC = ?) AND (VVTIDAEE = ?) "

    Const SQL_GET_SELECTION_GROUP_BY = _
        "  GROUP BY VVTIDVVT, VVTDTINV, VVTIDDEF, VVPNRVER, OCRNRTCK, VCAQCALT, VVTFLSTV"
 
    Dim strQ As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    strQ = SQL_GET_SELECTION_EXPORT_1
    If Not IsMissing(DaNumeroViaggioDef) Then
        If Len(DaNumeroViaggioDef) > 0 Then
            strQ = strQ & " AND (VVTIDDEF >= " & DS.DA.MakeSQL(DaNumeroViaggioDef, dtAlfa) & ")"
        End If
    End If
    If Not IsMissing(ANumeroViaggioDef) Then
        If Len(ANumeroViaggioDef) > 0 Then
            strQ = strQ & " AND (VVTIDDEF <= " & DS.DA.MakeSQL(ANumeroViaggioDef, dtAlfa) & ")"
        End If
    End If
 
    strQ = strQ & SQL_GET_SELECTION_GROUP_BY
    
    If Not DS.DA.QueryCommand(strQ, varBuffer, lngHowMany, DallaData, AllaData, IDSocietà, IDEnteEmittente) Then
        GoTo ErrHandler
    End If
    GetSelectionExportNEN = varBuffer
    Exit Function
ErrHandler:
    GetSelectionExportNEN = Empty
End Function


Public Function GetSelectionCreatePPD(Optional TipoPrefattura As String = "A") As Boolean
    ReDim varOrderFields(0 To 1) As Variant
    With DS
        ' --- Seleziono tutti gli scarichi dei viaggi completati
        .Fields("10").FindValue = "C"
        If TipoPrefattura = "P" Then
            .Fields("10").FindValue(2) = "E"
        End If
        .Fields("18").FindValue = False
        .Fields("19").FindValue = False
        varOrderFields(0) = "2"
        varOrderFields(1) = "20"
        If TipoPrefattura = "P" Then
            If Len(DS.Field("Vettore Trasportatore").FindValue) = 0 Then
                DS.Field("Vettore Trasportatore").Operator = " <> "
                If DS.Field("Ente Emittente").FindValue = "624" Then
                    DS.Field("Vettore Trasportatore").FindValue = "A01"
                ElseIf DS.Field("Ente Emittente").FindValue = "004" Then
                    DS.Field("Vettore Trasportatore").FindValue = "A02"
                Else
                    Exit Function
                End If
            End If
        End If
        GetSelectionCreatePPD = .GetSelection(SQL_GET_SELECTION_CREATE_PPT, True, varOrderFields)
    End With
End Function

Public Function GetSelectionCreateSDT() As Boolean
    ReDim varOrderFields(0 To 0) As Variant
    
    varOrderFields(0) = "80"
    With DS
        .Field("Autista Annullato").FindValue = 0
        GetSelectionCreateSDT = .GetSelection(SQL_GET_SELECTION_CREATE_SDT, False, varOrderFields)
    End With
End Function

Public Function GetSelectionCreatePPDContratto(Optional TipoPrefattura As String = "A") As Boolean
    ReDim varOrderFields(0 To 1) As Variant
    With DS
        ' --- Seleziono tutti gli scarichi dei viaggi completati
        .Fields("10").FindValue = "C"
        If TipoPrefattura = "P" Then
            .Fields("10").FindValue(2) = "E"
        End If
        .Fields("18").FindValue = False
        .Fields("19").FindValue = False
        varOrderFields(0) = "2"
        varOrderFields(1) = "20"
        GetSelectionCreatePPDContratto = .GetSelection(SQL_GET_SELECTION_CREATE_PPT_CONTRATTO, True, varOrderFields)
    End With
End Function

Public Function GetSelectionExport1(Optional Presidio As String = vbNullString) As Boolean
    Const SQL_GET_SELECTION_EXPORT_1 = _
        " SELECT DISTINCT " & _
                    " VVTIDVVT, VVTIDAUT, VVTSEQVI, VVTHRINV, " & _
                    " VVTIDADE, VVTIDAEE, VVTDTINV, VVTIDDEF, " & _
                    " VVTIDVE1, VVTIDAD1, VVTIDEE1, VVTKMPRE, " & _
                    " VVTKMDEF, VVTIDAM1, ANAGATM0_1.ATMIDATM, VVTIDAM2, " & _
                    " VVTHRINC, VVTIDAU1, ANAGAAU0.AAUNOMEA, ANAGAAU0.AAUNRCID, " & _
                    " ANAGAAU0_1.AAUNOMEA, VVTIDAMD, OCTRIFE1 + ISNULL(OCRRIFE2,''), VVCIDADZ, " & _
                    " OCTRIFE3, VVCTEMVG, left(OCTNOTA2, 6), substring(OCTNOTA2, 7, 6), " & _
                    " right(OCTNOTA2, 5), OCTDTCRI, OCTFLPOR, VVPIDAPR, " & _
                    " VVPQTRLT, OCRQCOLT, 'L', OCTRAUT4, VVCNRSEQ, VVPNRVER " & _
        " FROM VIAGVVT0 INNER JOIN " & _
                    " VIAGVVC0 ON VVCIDVVT = VVTIDVVT INNER JOIN " & _
                    " ANAGADE0 ON VVTIDADE = ADEIDADE INNER JOIN " & _
                    " ANAGAZO0 ON AZOIDAZO = VVCIDAZO INNER JOIN " & _
                    " VIAGVVP0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC INNER JOIN " & _
                    " ORDIOCT0 ON VVCIDOCT = OCTIDOCT INNER JOIN " & _
                    " ORDIOCR0 ON OCTIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR INNER JOIN " & _
                    " ANAGAAM0 ON VVTIDAM1 = ANAGAAM0.AAMIDAAM INNER JOIN " & _
                    " ANAGATM0 ON ANAGAAM0.AAMIDATM = ANAGATM0.ATMIDATM LEFT OUTER JOIN " & _
                    " ANAGAAM0 AS ANAGAAM0_1 ON VVTIDAM2 = ANAGAAM0_1.AAMIDAAM LEFT OUTER JOIN " & _
                    " ANAGATM0 AS ANAGATM0_1 ON ANAGAAM0_1.AAMIDATM = ANAGATM0_1.ATMIDATM INNER JOIN " & _
                    " ANAGAAU0 ON VVTIDAU1 = ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
                    " ANAGAAU0 AS ANAGAAU0_1 ON VVTIDAU2 = ANAGAAU0_1.AAUIDAAU INNER JOIN ANAGADZ0 " & _
                    " ON VVCIDADZ = ADZIDADZ"
 
    ReDim varOrderFields(1) As Variant
    With DS
        .Fields("10").FindValue = "D"
        .Fields("10").FindValue(2) = "E"
        .Fields("10").FindValue(3) = "C"
        
        .Fields("93").FindValue = Presidio
        .Field("Tipo Destinazione").FindValue(1) = "R"
        .Field("Tipo Destinazione").FindValue(2) = "B"
        .Field("Tipo Destinazione").FindValue(3) = "E"
        
        '.Fields("91").FindValue = 0
        '.Fields("92").FindValue = "RP"
        varOrderFields(0) = "0"
        varOrderFields(1) = "94"
        GetSelectionExport1 = .GetSelection(SQL_GET_SELECTION_EXPORT_1, True, varOrderFields)
    End With
End Function

' Seleziona i viaggi allo scarico che devono essere modificati e trasmessi
Public Function GetSelectionViaggiAlloScarico(Optional ActiveRecordsOnly As Boolean = False, _
                                              Optional OrderFields As Variant, _
                                              Optional MaxRecords As Long = 0) As Boolean
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
'        "       vvcdtcon, vvtflst3, vvtflst2, octcgpsx, " & _
'        "       octcgpsy, vvchrpre, vvtidau2, vvtidve1, " & _
'        "       vvtidve2, vvtfludg, vvpidoct, ocridaum, " & _
'        "       aaunote1, octnota1, octhrfr1, octhrto1, " & _
'        "       octhrfr2, octhrto2 " & _
'        " FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN anagade0 ON vvtidade = adeidade " & _
'        "               INNER JOIN VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN ANAGADZ0 ON VVCIDADZ = ADZIDADZ " & _
'        "               INNER JOIN ORDIOCR0 ON VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT AND VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR " & _
'        "               INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU INNER JOIN ORDIOCT0 ON ORDIOCR0.OCRIDOCT = ORDIOCT0.OCTIDOCT " & _
'        "               LEFT OUTER JOIN MULTSUT0 ON VVTIDADE = SUTIDASC AND VVTIDAUT = SUTIDAUT "
    
'-- DP 15/05/08
'-- Aggiunti i campi VVCHRARR, VVCDTPRE, VVCDTARR
    Const SQL_GET_SELECTION_VIAGGI_ALLO_SCARICO = _
        "SELECT vvtidvvt, vvcidvvc, vvpidvvp, vvtdtinv, " & _
        "       vvtidade, vvtidvvt, vvtiddef, vvtidaee, " & _
        "       vvtseqvi, vvthrinv, vvthrinc, vvtidaut, " & _
        "       vvtidam1, vvtidam2, vvtidau1, vvtkmdef, " & _
        "       aaunrcid, sutcodso, vvcnrseq, OCTRIFE1 + ISNULL(OCRRIFE2,''), " & _
        "       vvcidadz, vvcderas, vvchrpre, adzkmrag, " & _
        "       vvckmpad, vvcflst5, vvcnote1, vvcnotee, " & _
        "       vvpidapr, ocrqrilt, vvpqtrlt, vvpqcolt, " & _
        "       vvpflst5, vvtidasc, octraut4, vvtnripa, " & _
        "       vvcdtcon, vvtflst3, vvtflst2, octcgpsx, " & _
        "       octcgpsy, vvctemvg, vvtidve1, vvtidve2, " & _
        "       vvtfludg, vvpidoct, ocridaum, aaunote1, " & _
        "       octnota1, octhrfr1, octhrto1, octhrfr2, " & _
        "       octhrto2, svgprgme, svcprgme, vvphrfis, " & _
        "       vvtflsta, vvtflst6, vvcflst4, vvcflst6, " & _
        "       vvpflncc, vvpfldes, vvckmpap, vvtkmrie, " & _
        "       vvchrarr, vvcdtpre, vvcdtarr " & _
        " FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN anagade0 ON vvtidade = adeidade " & _
        "               INNER JOIN VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN ANAGADZ0 ON VVCIDADZ = ADZIDADZ " & _
        "               INNER JOIN ORDIOCR0 ON VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT AND VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR " & _
        "               INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU INNER JOIN ORDIOCT0 ON ORDIOCR0.OCRIDOCT = ORDIOCT0.OCTIDOCT " & _
        "               LEFT OUTER JOIN MULTSUT0 ON VVTIDADE = SUTIDASC AND VVTIDAUT = SUTIDAUT " & _
        "               LEFT OUTER JOIN VIAGSVG0 ON SVGIDVVT = VVTIDVVT " & _
        "               LEFT OUTER JOIN VIAGSVC0 ON SVCIDVVT = VVCIDVVT AND SVCIDVVC = VVCIDVVC "
    
    Dim strSQL As String

    strSQL = SQL_GET_SELECTION_VIAGGI_ALLO_SCARICO
    If Not IsMissing(OrderFields) Then
        GetSelectionViaggiAlloScarico = DS.GetSelection(strSQL, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionViaggiAlloScarico = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)
    End If
End Function

Public Function GetSelectionExport(Optional Presidio As String = vbNullString) As Boolean
    ReDim varOrderFields(0) As Variant
    With DS
        .Fields("10").FindValue = "D"
        .Fields("90").FindValue = Presidio
        .Fields("91").FindValue = 0
        .Fields("92").FindValue = "RP"
        varOrderFields(0) = "0"
        GetSelectionExport = .GetSelection(SQL_GET_SELECTION_EXPORT, True, varOrderFields)
    End With
End Function

Public Function GetSelectionExportExViagDetail(Optional FlagFirstTimeOnly As Boolean = True) As Boolean
    ReDim varOrderFields(1) As Variant
    With DS
'        .Fields("10").FindValue(1) = "I"
'        .Fields("10").FindValue(2) = "C"
        .Fields("33").FindValue = ""
'        .Fields("92").FindValue = "GC"
        If FlagFirstTimeOnly Then
            .Fields("35").FindValue = "N"
        Else
            .Fields("35").FindValue = Empty
        End If
        varOrderFields(0) = "13"
        varOrderFields(1) = "5"
        GetSelectionExportExViagDetail = .GetSelection(SQL_GET_SELECTION_EXPORT_EXVIAGDETAIL, False, varOrderFields)
    End With
End Function


Public Function GetSelectionExportDest() As Boolean
    ReDim varOrderFields(0) As Variant
    With DS
        .Fields("10").FindValue = "D"
'        .Fields("16").FindValue = "T"
        varOrderFields(0) = "0"
        GetSelectionExportDest = .GetSelection(SQL_GET_SELECTION_EXPORT_DEST, True, varOrderFields)
    End With
End Function

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0, _
                             Optional FlagGC As Boolean = False) As Boolean
                             
    Dim strSQL As String

    If Not FlagGC Then
        strSQL = SQL_GET_SELECTION
    Else
        strSQL = SQL_GET_SELECTION_GC
    End If
    
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)
    End If
End Function



Public Function PopulateTempTable() As Boolean
    On Error GoTo ErrHandler
    Dim varResults As Variant
    ReDim varidprodotti(0 To 10)
    Dim j As Integer
    Dim lngHowMany As Long
    Dim i As Long
    Dim strIDViaggio As String
    Dim strPrevIDViaggio As String
    Dim strCond As String
    Dim strSQL As String
    
    If Not DS.DA.ExecuteQuery(SQL_CLEAR_TRIP_PRODUCTS) Then
        GoTo ErrHandler
    End If
    
    strCond = DS.MakeKeyCond
    strSQL = SQL_SELECT_TRIP_PRODUCTS & " WHERE " & strCond  'MakeKeyCond
    If (Right$(strSQL, 5) <> " AND ") And (Len(strCond) > 0) Then
        strSQL = strSQL & " AND "
    End If
    strSQL = strSQL & DS.CancelField & " = " & DS.DA.MakeSQL(False, dtBit)
    
    If Not DS.DA.QueryData(strSQL, varResults, lngHowMany) Then
        GoTo ErrHandler
    End If
    
    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If
    
    For j = 0 To 10
        varidprodotti(j) = vbNullString
    Next j
    j = 1
    strPrevIDViaggio = varResults(0, 0)
    For i = 0 To lngHowMany - 1
        strIDViaggio = varResults(0, i)
        If strIDViaggio <> strPrevIDViaggio Then
            varidprodotti(0) = strPrevIDViaggio
            strPrevIDViaggio = strIDViaggio
            If Not DS.DA.ExecuteCommand(SQL_INSERT_TRIP_PRODUCTS, varidprodotti(0), _
                                        varidprodotti(1), varidprodotti(2), _
                                        varidprodotti(3), varidprodotti(4), _
                                        varidprodotti(5), varidprodotti(6), _
                                        varidprodotti(7), varidprodotti(8), _
                                        varidprodotti(9), varidprodotti(10)) Then
                GoTo ErrHandler
            End If
            For j = 0 To 10
                varidprodotti(j) = vbNullString
            Next j
            j = 1
        End If
        varidprodotti(j) = varResults(1, i)
        j = j + 1
    Next i
    If Not DS.DA.ExecuteCommand(SQL_INSERT_TRIP_PRODUCTS, strIDViaggio, _
                                varidprodotti(1), varidprodotti(2), _
                                varidprodotti(3), varidprodotti(4), _
                                varidprodotti(5), varidprodotti(6), _
                                varidprodotti(7), varidprodotti(8), _
                                varidprodotti(9), varidprodotti(10)) Then
        GoTo ErrHandler
    End If
    PopulateTempTable = True
    Exit Function
ErrHandler:
    PopulateTempTable = False
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
'    SQL_GET_SELECTION = "SELECT vvtidvvt, vvtiddef, vvtidasc, vvtdtinv, " & _
'                        "       vvtidaut, vvtidam1, vvtidam2, vvtidau1, " & _
'                        "       vvtidau2, vvtidade, vvtidamd, vvtfltiv, " & _
'                        "       vvtkmdef, vvtflstv, vvtannul FROM viagvvt0 "
    SQL_GET_SELECTION = "SELECT DISTINCT vvtidvvt, vvtiddef, vvtidasc, vvtdtinv, " & _
                        "       vvtidaut, vvtidam1, vvtidam2, vvtidau1, " & _
                        "       vvtidau2, vvtidade, vvtidamd, vvtfltiv, " & _
                        "       vvtkmdef, vvtflstv, vvtannul FROM viagvvt0 INNER JOIN viagvvc0 ON vvtidvvt = vvcidvvt " & _
                        "       INNER JOIN ordioct0 ON vvcidoct = octidoct INNER JOIN anagazo0 ON vvcidazo = azoidazo "
    SQL_GET_SELECTION_GC = "SELECT DISTINCT vvtidvvt, vvtiddef, vvtidasc, vvtdtinv, " & _
                        "       vvtidaut, vvtidam1, vvtidam2, vvtidau1, " & _
                        "       vvtidau2, vvtidade, vvtidamd, vvtfltiv, " & _
                        "       vvtkmdef, vvtflstv, vvtannul FROM viagvvt0 INNER JOIN viagvvc0 ON vvtidvvt = vvcidvvt " & _
                        "       INNER JOIN ordioct0 ON vvcidoct = octidoct INNER JOIN anagade0 ON vvtidade = adeidade "
    SQL_GET_SELECTION_EXPORT = "SELECT DISTINCT vvtidvvt, vvtidaut, vvtseqvi, vvthrinv, " & _
                               "       vvtidade, vvtidaee, vvtdtinv, vvtiddef, " & _
                               "       vvtidve1, vvtidad1, vvtidee1, " & _
                               "       vvtkmpre, vvtkmdef, vvtidam1, anagatm0.atmidatm, " & _
                               "       vvtidam2, anagatm0_1.atmidatm, vvthrinc, " & _
                               "       vvtidau1, anagaau0.aaunomea, vvtidau2, anagaau0_1.aaunomea, " & _
                               "       acnnumco, acnnumse, anagaau0.aaunumpa, anagaau0_1.aaunumpa, " & _
                               "       vvtidamd " & _
                                " FROM ANAGAZO0 INNER JOIN VIAGVVC0 INNER JOIN " & _
                                " VIAGVVT0 ON VIAGVVC0.VVCIDVVT = VIAGVVT0.VVTIDVVT ON ANAGAZO0.AZOIDAZO = VIAGVVC0.VVCIDAZO INNER JOIN " & _
                                " ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT LEFT OUTER JOIN " & _
                                " ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
                                " ANAGATM0 ON ANAGAAM0.AAMIDATM = ANAGATM0.ATMIDATM LEFT OUTER JOIN " & _
                                " ANAGAAM0 ANAGAAM0_1 ON VIAGVVT0.VVTIDAM2 = ANAGAAM0_1.AAMIDAAM LEFT OUTER JOIN " & _
                                " ANAGATM0 ANAGATM0_1 ON ANAGAAM0_1.AAMIDATM = ANAGATM0_1.ATMIDATM LEFT OUTER JOIN " & _
                                " ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
                                " ANAGAAU0 ANAGAAU0_1 ON VIAGVVT0.VVTIDAU2 = ANAGAAU0_1.AAUIDAAU LEFT OUTER JOIN " & _
                                " ANAGACN0 ON VIAGVVT0.VVTIDCN1 = ANAGACN0.ACNIDACN "
                      
'                               " FROM (((((ANAGATM0 RIGHT JOIN " & _
'                               "         (ANAGAAM0 RIGHT JOIN VIAGVVT0 ON ANAGAAM0.AAMIDAAM = VIAGVVT0.VVTIDAM1) " & _
'                               "          ON ANAGATM0.ATMIDATM = ANAGAAM0.AAMIDATM) LEFT JOIN ANAGAAM0 AS ANAGAAM0_1 " & _
'                               "          ON VIAGVVT0.VVTIDAM2 = ANAGAAM0_1.AAMIDAAM) LEFT JOIN ANAGATM0 AS ANAGATM0_1 " & _
'                               "          ON ANAGAAM0_1.AAMIDATM = ANAGATM0_1.ATMIDATM) LEFT JOIN ANAGAAU0 " & _
'                               "          ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU) LEFT JOIN ANAGAAU0 AS ANAGAAU0_1 " & _
'                               "          ON VIAGVVT0.VVTIDAU2 = ANAGAAU0_1.AAUIDAAU) LEFT JOIN ANAGACN0 ON VIAGVVT0.VVTIDCN1 = ANAGACN0.ACNIDACN "
'
    SQL_GET_SELECTION_EXPORT_DEST = "SELECT vvtidvvt, vvtidade, vvcidadz " & _
                               "  FROM viagvvt0 INNER JOIN viagvvc0 ON viagvvt0.vvtidvvt = viagvvc0.vvcidvvt "
    
    SQL_GET_SELECTION_CREATE_PPT = "SELECT VVPIDVVT, VVCDERAS, VVPIDAPR, VVPQCOLT, VVPIDOCT, " & _
                                   "       VVPIDOCR, OCTDTRIC, OCTDTCRI, OCTRIFE1, OCTRIFE2, " & _
                                   "       OCRQCOKG, OCRQCOLT, OCRQCOPZ, OCRQCOCN, VVPIDVVC, " & _
                                   "       VVPIDVVP, VVTDTINC, VVPQTRKG, VVPQTRLT, VVPQTRPZ, " & _
                                   "       VVPQTRCN, VVTIDASC, VVTIDAEE, VVTIDADE, VVCIDACL, " & _
                                   "       VVCIDADZ, SDZCODSO, VVPIDAIM, VVPIDAPD, VVPQCOKG, " & _
                                   "       VVPQCOPZ, VVPQCOCN, VVCIDAZO, VVCKMPAD, VVTIDAUT, " & _
                                   "       OCRQRILT, OCTRAUT3, OCRRIFE2, OCTRAUT4, VVTDTINV, " & _
                                   "       ANAGAAM0.AAMIDATS, ANAGAAM0_1.AAMIDATS, VVTKMDEF, VVTNRIPA, OCTDTGRA, " & _
                                   "       OCTFLTOR, VVTIDVE1, VVTIDVE2, OCTFLRES  " & _
                                   "  FROM VIAGVVT0 INNER JOIN ((VIAGVVC0 LEFT JOIN MULTSDZ0 ON " & _
                                   "      (VIAGVVC0.VVCIDADZ = MULTSDZ0.SDZIDADZ) AND (VIAGVVC0.VVCIDASC = MULTSDZ0.SDZIDASC)) INNER JOIN " & _
                                   "      (ORDIOCT0 RIGHT JOIN (VIAGVVP0 LEFT JOIN ORDIOCR0 ON " & _
                                   "      (VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT) AND (VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR)) ON " & _
                                   "       ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT) ON (VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) AND (VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT)) " & _
                                   "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT LEFT OUTER JOIN " & _
                                   "       ANAGAAM0 ANAGAAM0_1 ON VIAGVVT0.VVTIDAM2 = ANAGAAM0_1.AAMIDAAM LEFT OUTER JOIN " & _
                                   "       ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM "

    SQL_GET_SELECTION_CREATE_PPT_CONTRATTO = "SELECT VVPIDVVT, VVCDERAS, VVPIDAPR, VVPQCOLT, VVPIDOCT, " & _
                                   "       VVPIDOCR, OCTDTRIC, OCTDTCRI, OCTRIFE1, OCTRIFE2, " & _
                                   "       OCRQCOKG, OCRQCOLT, OCRQCOPZ, OCRQCOCN, VVPIDVVC, " & _
                                   "       VVPIDVVP, VVCDTCON, VVPQTRKG, VVPQTRLT, VVPQTRPZ, " & _
                                   "       VVPQTRCN, VVTIDASC, VVTIDAEE, VVTIDADE, VVCIDACL, " & _
                                   "       VVCIDADZ, SDZCODSO, VVPIDAIM, VVPIDAPD, VVPQCOKG, " & _
                                   "       VVPQCOPZ, VVPQCOCN, VVCIDAZO, VVCKMPAD, VVTIDAUT, " & _
                                   "       OCRQRILT, OCTRAUT3, OCRRIFE2, OCTRAUT4, VVTDTINV, " & _
                                   "       ANAGAAM0.AAMIDATS, ANAGAAM0_1.AAMIDATS, VVTKMDEF, ACNIDACN " & _
                                   "  FROM VIAGVVT0 INNER JOIN ((VIAGVVC0 LEFT JOIN MULTSDZ0 ON " & _
                                   "      (VIAGVVC0.VVCIDADZ = MULTSDZ0.SDZIDADZ) AND (VIAGVVC0.VVCIDASC = MULTSDZ0.SDZIDASC) INNER JOIN ANAGADZ0 ON VVCIDADZ = ADZIDADZ LEFT OUTER JOIN ANAGACN0 ON ADZIDACL = ACNIDACL AND ADZRAGU5 = ACNIDACN) INNER JOIN " & _
                                   "      (ORDIOCT0 RIGHT JOIN (VIAGVVP0 LEFT JOIN ORDIOCR0 ON " & _
                                   "      (VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT) AND (VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR)) ON " & _
                                   "       ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT) ON (VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) AND (VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT)) " & _
                                   "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT LEFT OUTER JOIN " & _
                                   "       ANAGAAM0 ANAGAAM0_1 ON VIAGVVT0.VVTIDAM2 = ANAGAAM0_1.AAMIDAAM LEFT OUTER JOIN " & _
                                   "       ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM "
    
    SQL_GET_SELECTION_CREATE_SDT = "SELECT AAUIDAAU, AAUNOMEA " & _
                                   "  FROM ANAGAAU0 "
                                   
    SQL_SELECT_TRIP_PRODUCTS = "SELECT DISTINCT VVPIDVVT, VVPIDAPR " & _
                                  "  FROM VIAGVVT0 INNER JOIN (VIAGVVC0 INNER JOIN VIAGVVP0 " & _
                                  "       ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) " & _
                                  "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT "
    SQL_INSERT_TRIP_PRODUCTS = "INSERT INTO VIAGTMP0 VALUES (?,?,?,?,?,?,?,?,?,?,?)"
    SQL_CLEAR_TRIP_PRODUCTS = "DELETE FROM VIAGTMP0 "
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
        .Add "Dalla Data fine Viaggio", "viagvvt0.vvtdtfiv", 2, , 10, "101", " >= "
        .Add "Alla Data fine Viaggio", "viagvvt0.vvtdtfiv", 2, , 10, "100", " <= "
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
        
        .Add "Codice Autista", "anagaau0.aauidaau", 0, False, 3, "80"
        With .Item("80")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Vettore Autista", "anagaau0.aauidave", 0, False, 3, "81"
        With .Item("81")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Base Autista", "anagaau0.aauidabs", 0, False, 3, "82"
        With .Item("82")
            .ForeignTableEntity = "ABS"
            .ForeignTable = "anagabs0"
            .ForeignTableFld = "absidabs"
            .ForeignTableDescField = "absdescr"
            .ForeignTableDeleteFld = "absannul"
        End With
        .Add "Autista Annullato", "anagaau0.aauannul", 3, , , "83"
        
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
        .Add "Tipo Destinazione", "anagadz0.adzidatd", 0, False, 3, "99"
        With .Item("99")
            .ForeignTableEntity = "ATD"
            .ForeignTable = "anagatd0"
            .ForeignTableFld = "atdidatd"
            .ForeignTableDescField = "atddescr"
            .ForeignTableDeleteFld = "atdannul"
        End With
        
        .Add "Vettore Trasportatore", "viagvvt0.vvtidve2", 0, False, 3, "39"
        With .Item("39")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Prodotto Stato Riga 5", "viagvvp0.vvpflst5", 0, False, 1, "31"
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
        .Add "31", "33"
        .Add "99", "34"
        .Add "80", "35"
        .Add "81", "36"
        .Add "82", "37"
        .Add "83", "38"
        .Add "39", "39"
        .Add "101", "40"
        .Add "100", "41"
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




