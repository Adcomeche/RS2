VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrdineSelezione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
Public SQL_GET_SELECTION_GROUPBY As String
Public SQL_GET_SELECTION_GC As String
Public SQL_GET_SELECTION_GC_GROUPBY As String
Public SQL_GET_SELECTION_EXPORT As String
Public SQL_GET_SELECTION_EXPORT1 As String
Public SQL_GET_SELECTION_EXPORT_EA As String
Public SQL_GET_SELECTION_CREATE_PPT As String
Public SQL_GET_SELECTION_FRAZIONAMENTO As String
Public SQL_GET_SELECTION_FRAZIONAMENTO_GROUPBY As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_SET_AVANZAMENTO As String
Public SQL_GET_AUT_BY_RIFE1 As String
Public SQL_GET_AUT_BY_APR As String

Private Const SQL_GET_SELECTION_EVASIONE = _
"SELECT OCTRIFE2, '' AS Seq, '' AS IDAut, '' AS Espr1, " & _
"       OCTRAUT2, OCTFLPOR, ADZRASOC, ADZLOCAL, " & _
"       ASCRASOC, OCTIDADE, ADEDESCR, OCTDTCRI, " & _
"       OCTKMRAG, OCTNOTA3, OCTIDADZ, OCTIDOCT, " & _
"       0, 0, 0, 0, " & _
"       0, 0, 0, OCTIDASC " & _
"  FROM ANAGADE0 INNER JOIN (ANAGADZ0 INNER JOIN (ANAGASC0 INNER JOIN ORDIOCT0 ON ANAGASC0.ASCIDASC = ORDIOCT0.OCTIDASC) " & _
"       ON ANAGADZ0.ADZIDADZ = ORDIOCT0.OCTIDADZ) ON ANAGADE0.ADEIDADE = ORDIOCT0.OCTIDADE "

'-- DP 19/08/08
Private Const SQL_GET_SELECTION_RIDUZIONI = "SELECT ORDIOCT0.OCTDTCPI, ORDIOCT0.OCTIDADE, ANAGADE0.ADEDESCR, ORDIOCT0.OCTRIFE1, " & _
                                            "       ORDIOCR0.OCRRIFE2, ORDIOCT0.OCTRAUT4, ORDIOCT0.OCTIDADZ, ANAGADZ0.ADZRASOC, " & _
                                            "       ORDIOCR0.OCRIDAPR, ORDIOCR0.OCRQRILT, ORDIOCR0.OCRQCOLT, ORDIOCR0.OCRQRILT - ORDIOCR0.OCRQCOLT " & _
                                            "FROM ORDIOCT0 INNER JOIN ORDIOCR0 ON " & _
                                            "       ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT " & _
                                            "INNER JOIN ANAGADZ0 ON " & _
                                            "       ORDIOCT0.OCTIDADZ = ANAGADZ0.ADZIDADZ " & _
                                            "INNER JOIN ANAGADE0 ON " & _
                                            "       ORDIOCT0.OCTIDADE = ANAGADE0.ADEIDADE "

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function GetSelectionCreatePPD() As Boolean
    ReDim varOrderFields(0 To 0) As Variant
    With DS
        ' --- Seleziono tutti gli scarichi dei viaggi completati
        varOrderFields(0) = "24"
        GetSelectionCreatePPD = .GetSelection(SQL_GET_SELECTION_CREATE_PPT, True, varOrderFields)
    End With
End Function


Public Function GetSelectionEvasione() As Variant
    On Error GoTo ErrHandler
    
    ReDim varOrderFields(0 To 0) As Variant
    Dim blnResult As Boolean
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strIDOCT As String
    Dim varDatiAUT As Variant
    Dim varBuffer As Variant
    Dim varDummy As Variant
    Dim varTot As Variant
    Dim AXX As CParametro
    Dim OCR As COrdineRiga
    Dim AAU As CAutista
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
        .Fields(AXX.DS.KeyFields(1)).FindValue = "001"
        AXX.Find
    End With
        
    Set OCR = New COrdineRiga
    With OCR.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    With DS
        blnResult = .GetSelection(SQL_GET_SELECTION_EVASIONE, True)
        If Not blnResult Then
            GoTo ErrHandler
        End If
        varBuffer = .List
        For i = 0 To UBound(varBuffer, 2)
            If varBuffer(5, i) = "N" Then
                varBuffer(5, i) = vbNullString
            End If
            varDatiAUT = varBuffer(0, i)
            varBuffer(0, i) = Left(varDatiAUT, 3)
            varBuffer(1, i) = Mid(varDatiAUT, 4, 3)
            varBuffer(2, i) = Mid(varDatiAUT, 7, 3)
            If Not IsNull(varBuffer(2, i)) Then
                AAU.DS.Fields("0").FindValue = varBuffer(2, i)
                AAU.Find
                If AAU.DS.HowMany > 0 Then
                    varBuffer(3, i) = AAU.DS.Fields("1").Value
                End If
            End If
            strIDOCT = varBuffer(15, i)
            OCR.DS.Fields("0").FindValue = strIDOCT
            If OCR.GetList(True) Then
                With OCR.DS
                    varTot = 0
                    For j = 1 To 6
                        For k = 0 To .HowMany - 1
                            If .List(1, k) = AXX.DS.Fields(CStr(5 + j)).Value Then
                                varBuffer(15 + j, i) = .List(4, k) / 1000
                                varTot = varTot + varBuffer(15 + j, i)
                            End If
                        Next k
                    Next j
                    varBuffer(15 + j, i) = varTot
                End With
            End If
        Next i
    End With
    GetSelectionEvasione = varBuffer
    Exit Function
ErrHandler:
    GetSelectionEvasione = Empty
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub

Public Function GetAUT(RifOrdineEsterno As String, _
                       IDProdotto As String, _
                       IDDestinazione As String, _
                       Data As String) As String
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varRowBuf As Variant
    Dim lngHowMany As Long
    
    lngHowMany = 0
    blnResult = DS.DA.QueryCommand(Me.SQL_GET_AUT_BY_RIFE1, varRowBuf, lngHowMany, RifOrdineEsterno)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If lngHowMany > 0 Then
        GetAUT = varRowBuf(0, 0)
    Else
'        blnResult = DS.DA.QueryCommand(Me.SQL_GET_AUT_BY_APR, varRowBuf, lngHowMany, _
'                                       IDProdotto, IDDestinazione, Data, DateAdd("ww", 1, Data))
'        If Not blnResult Then
            GoTo ErrHandler
'        End If
'        If lngHowMany > 0 Then
'            GetAUT = varRowBuf(0, 0)
'        Else
'            GoTo ErrHandler
'        End If
    End If
    Exit Function
ErrHandler:
    GetAUT = vbNullString
End Function

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0, _
                             Optional FlagGC As Boolean = False) As Boolean
    Dim strSQL As String
    Dim strSQLGroupBy As String
    
    If Not FlagGC Then
        strSQL = SQL_GET_SELECTION
        strSQLGroupBy = SQL_GET_SELECTION_GROUPBY
    Else
        strSQL = SQL_GET_SELECTION_GC
        strSQLGroupBy = SQL_GET_SELECTION_GC_GROUPBY
    End If
    
'    DS.Field("Riga annullata").FindValue = 0
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, OrderFields, MaxRecords, , strSQLGroupBy)
    Else
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords, , strSQLGroupBy)
    End If
End Function

Public Function GetSelectionFrazionamento(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional MaxRecords As Long = 0, _
                             Optional FlagGC As Boolean = False) As Boolean
    Dim strSQL As String
    ReDim varOrderFields(1) As Variant
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    strSQL = SQL_GET_SELECTION_FRAZIONAMENTO
    
    DS.Field("Avanzamento Ordine").FindValue = "A"
    
    DS.Field("Quantita Iniziale L").Operator = " > "
    DS.Field("Quantita Iniziale L").FindValue = 32000
    
    DS.Field("Quantita Richiesta L").Operator = " > "
    DS.Field("Quantita Richiesta L").FindValue = 0
    
    DS.Field("Raggruppamento Utente 4").Operator = " <> "
    DS.Field("Raggruppamento Utente 4").FindValue(1) = "#NULL#"
    DS.Field("Raggruppamento Utente 4").FindValue(2) = "''"
    
    varOrderFields(0) = "24"    ' Numero Ordine
    varOrderFields(1) = "94"     ' Prodotto
    
    GetSelectionFrazionamento = DS.GetSelection(strSQL, ActiveRecordsOnly, varOrderFields, MaxRecords)
    
End Function
Public Function GetSelectionExport(Optional OnlyAck As Boolean = False, _
                                   Optional GrandesCuentas As Boolean = False, _
                                   Optional Presidio As String = vbNullString) As Boolean
   ReDim varOrderFields(0) As Variant
   With DS
        .Fields("23").FindValue = "A"
        If Not OnlyAck Then
            .Fields("23").FindValue(2) = "M"
        End If
        If GrandesCuentas Then
            .Fields("32").FindValue = "GC"
        Else
            .Fields("32").FindValue = "RP"
        End If
        If Not GrandesCuentas Then
            If Len(Presidio) > 0 Then
                .Fields("90").FindValue = Presidio
            Else
                .Fields("90").FindValue = Empty
            End If
            .Fields("91").FindValue = Empty
        Else
            ' Captura el Presidio RS3
            If Len(Presidio) > 0 Then
                .Fields("89").FindValue = Presidio
            Else
                .Fields("89").FindValue = Empty
            End If

            ' Captura el presidio RS2
'            If Len(Presidio) > 0 Then
'                .Fields("91").FindValue = Presidio
'            Else
'                .Fields("91").FindValue = Empty
'            End If


            .Fields("90").FindValue = Empty
        End If
        .Field("Raggruppamento Utente 5").Operator = " <> "
        .Field("Raggruppamento Utente 5").FindValue(1) = "#NULL#"
        .Field("Raggruppamento Utente 5").FindValue(2) = "S"
        
        varOrderFields(0) = "24"
        GetSelectionExport = .GetSelection(SQL_GET_SELECTION_EXPORT, True, varOrderFields)
    End With
End Function

Public Function GetSelectionExportEA() As Boolean
   ReDim varOrderFields(1) As Variant
   With DS
        .Field("Tipo Ordine").FindValue(1) = "EA"
        .Field("Tipo Ordine").FindValue(2) = "O"
        varOrderFields(0) = "3"
        varOrderFields(1) = "44"
        GetSelectionExportEA = .GetSelection(SQL_GET_SELECTION_EXPORT_EA, True, varOrderFields)
    End With
End Function

Public Function GetSelectionExport1(Optional OnlyAck As Boolean = False) As Boolean
   ReDim varOrderFields(0) As Variant
   With DS
'        .Fields("22").FindValue = "A"
        .Fields("23").FindValue = "A"
        If Not OnlyAck Then
            .Fields("23").FindValue(2) = "M"
        End If
        varOrderFields(0) = "24"
        GetSelectionExport1 = .GetSelection(SQL_GET_SELECTION_EXPORT1, True, varOrderFields)
    End With
End Function

'-- DP 19/08/08
Public Function GetSelection_Riduzioni(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelection_Riduzioni = DS.GetSelection(SQL_GET_SELECTION_RIDUZIONI, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection_Riduzioni = DS.GetSelection(SQL_GET_SELECTION_RIDUZIONI, ActiveRecordsOnly, , MaxRecords)
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

Public Function SetAvanzamento(NewValue As String) As Boolean
    Dim strIDOrdine As String
    Dim blnConcurErr As Boolean
    
    strIDOrdine = DS.Fields("24").FindValue
    DS.Fields.ResetFields
    
    DS.Fields("24").Value = strIDOrdine
    DS.Fields("24").FindValue = strIDOrdine
    DS.Fields("23").NewValue = NewValue
    DS.Fields("23").Updated = True
    BeginSave blnConcurErr
    DS.DA.EndTrans 0
    SetAvanzamento = True
    Exit Function
ErrHandler:
    SetAvanzamento = False
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT octidoct, octidasc, octidaee, octdtric, " & _
               "       octdtric, octdtins, octdtins, octdtcri, " & _
               "       octdtcpi, octdtcef, octrife1, octrife2, " & _
               "       octidacl, octidadz, octnota1, octnota2, " & _
               "       octidatv, octidade, octidazo, octidamp, " & _
               "       octidaro, octfltor, octflsor, octflaor, " & _
               "       octidoc1, octflpor, octraut1, octraut2, " & _
               "       octraut3, octraut4, octraut5, octflorm, " & _
               "       octidaro, octannul, timestmp FROM ordioct0 WHERE "
    SQL_GETLIST = "SELECT octidoct, octidasc FROM ordioct0 "
    SQL_GET_SELECTION = "SELECT octidoct, octidasc, octidacl, octidadz, " & _
                        "       adzrasoc, adzlocal, sum(ocrqcolt), sum(ocrqrilt), " & _
                        "       octidade, octidaee, octdtric, octdtins, " & _
                        "       octdtcri, octflaor, octannul " & _
                        "  FROM ordioct0 INNER JOIN anagadz0 ON ordioct0.octidadz = anagadz0.adzidadz  " & _
                        "       INNER JOIN anagazo0 ON octidazo = azoidazo " & _
                        "       LEFT OUTER JOIN ordiocr0 on octidoct = ocridoct "
                        
    SQL_GET_SELECTION_GROUPBY = "GROUP BY octidoct, octidasc, octidacl, octidadz, " & _
                        "       adzrasoc, adzlocal, octidade, octidaee, octdtric, " & _
                        "       octdtins, octdtcri, octflaor, octannul "
    SQL_GET_SELECTION_GC = "SELECT octidoct, octidade, adedescr, octdtcri, " & _
                           "       octidadz, adzrasoc, adzlocal, sum(ocrqcolt), " & _
                           "       sum(ocrqrilt), octidasc, octidacl, octidaee, " & _
                           "       octdtric, octdtins, octflaor, octannul " & _
                           "  FROM ordioct0 INNER JOIN anagadz0 ON ordioct0.octidadz = anagadz0.adzidadz  " & _
                           "       INNER JOIN anagade0 ON octidade = adeidade LEFT OUTER JOIN ordiocr0 on octidoct = ocridoct "
    SQL_GET_SELECTION_GC_GROUPBY = "GROUP BY octidoct, octidade, adedescr, octdtcri, " & _
                                   "       octidadz, adzrasoc, adzlocal, octidasc, " & _
                                   "       octidacl, octidaee, octdtric, octdtins, " & _
                                   "       octflaor, octannul "
'    SQL_GET_SELECTION_GC = "SELECT octidoct, octidasc, octidacl, octidadz, " & _
'                        "       adzrasoc, adzlocal, octidade, octidaee, octdtric, " & _
'                        "       octdtins, octdtcri, octflaor, octannul " & _
'                        "  FROM ordioct0 INNER JOIN anagadz0 ON ordioct0.octidadz = anagadz0.adzidadz  " & _
'                        "       INNER JOIN anagade0 ON octidade = adeidade "
    SQL_GET_SELECTION_EXPORT = "SELECT octidadz, adzrasoc, adzindir, adzlocal, " & _
                               "       adznrcap, adznrtel, octnota1, octnota2, " & _
                               "       octidoct, octidade, octdtcri, octidazo, " & _
                               "       octraut3, octflpor, octraut2, octnota3, " & _
                               "       octidasc, ascrasoc, octkmrag, octhrfr1, " & _
                               "       octhrto1, octhrfr2, octhrto2, octhrfr3, " & _
                               "       octhrto3, octconte, octfltor, octdtgra, " & _
                               "       octflcim, octflcpr, octnonav, octflazi, " & _
                               "       octflcnf, octhrap1, octhrch1, octhrap2, " & _
                               "       octhrch2, adznrte2, adznrtem " & _
                               "  FROM ANAGASC0 INNER JOIN (ordioct0 INNER JOIN anagadz0 ON ordioct0.octidadz = anagadz0.adzidadz) " & _
                               "       ON ORDIOCT0.octidasc = ANAGASC0.ascidasc INNER JOIN ANAGAZO0 ON ORDIOCT0.OCTIDAZO = ANAGAZO0.AZOIDAZO " & _
                               "       INNER JOIN anagade0 ON octidade = adeidade "
    SQL_GET_SELECTION_EXPORT_EA = "SELECT octrife3, octrife1, octidadz, adzrasoc, " & _
                                  "       octdtcri, octnota1, '',       octidoct, " & _
                                  "       octdtric, octhrric " & _
                                  "  FROM ordioct0 INNER JOIN anagadz0 ON octidadz = adzidadz "
                               
    SQL_GET_SELECTION_EXPORT1 = "SELECT octidoct, octidasc, octrife1, octidadz, adzprovi, " & _
                               "       adznrcap, adznrtel, octnota1, octidoct, " & _
                               "       octidacl, octidade, octdtcri, octraut1,octidaee " & _
                               "  FROM ordioct0 INNER JOIN anagadz0 ON ordioct0.octidadz = anagadz0.adzidadz "
    SQL_GET_SELECTION_FRAZIONAMENTO = "SELECT OCTIDOCT, OCTRAUT5, OCTIDASC, OCTRIFE1, " & _
                                      "       OCTIDADZ, ADZRASOC, OCTNONAV, ADZLOCAL, " & _
                                      "       ADZIDADE, OCTNOTA1, OCTDTCRI, OCRIDOCR, " & _
                                      "       OCRIDAPR, OCRQCOLT, OCRQRILT, OCTANNUL " & _
                                      "  FROM ORDIOCT0 INNER JOIN ORDIOCR0 ON OCTIDOCT = OCRIDOCT " & _
                                      "       INNER JOIN ANAGADZ0 ON OCTIDADZ = ADZIDADZ " & _
                                      "       INNER JOIN ANAGADE0 ON OCTIDADE = ADEIDADE "
'    SQL_GET_SELECTION_FRAZIONAMENTO_GROUPBY = " GROUP BY OCTIDOCT, OCTRAUT5, OCTIDASC, OCTRIFE1, " & _
'                                              "          OCTIDADZ, ADZRASOC, ADZLOCAL, ADZIDADE, " & _
'                                              "          OCTNOTA1, OCTDTCRI, OCRIDOCR, OCRIDAPR, " & _
'                                              "          OCRQCOLT, OCRQRILT, OCTANNUL " & _
'                                              "   HAVING (SUM(ORDIOCR0.OCRQCOLT) > 32000) "
    SQL_UPDATE = "UPDATE ordioct0 SET "
    SQL_INSERT = "INSERT INTO ordioct0 "
    SQL_GET_TS = "SELECT timestmp FROM ordioct0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT octidoct, octidasc, " & _
               "octidaee, octidade, octidacl, octidadz, " & _
               "octdtins, octdtric, octdtcpi, octdtcri, " & _
               "octidazo, octfltor, octflaor, octidapr, " & _
               "octannul FROM ordioct0 "
    SQL_SET_AVANZAMENTO = "UPDATE ordioct0 SET octflaor = "
    SQL_GET_AUT_BY_RIFE1 = "SELECT VIAGVVT0.VVTIDAUT " & _
                  "  FROM VIAGVVT0 INNER JOIN (ORDIOCT0 INNER JOIN VIAGVVC0 ON " & _
                  "  ORDIOCT0.OCTIDOCT = VIAGVVC0.VVCIDOCT) ON VIAGVVT0.VVTIDVVT " & _
                  " = VIAGVVC0.VVCIDVVT " & _
                  " WHERE ORDIOCT0.OCTRIFE1=?"
    SQL_GET_AUT_BY_APR = "SELECT VIAGVVT0.VVTIDAUT " & _
                     "   FROM VIAGVVT0 INNER JOIN (VIAGVVC0 INNER JOIN VIAGVVP0 ON (VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) AND (VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT)) ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT " & _
                     "  WHERE (((VIAGVVP0.VVPIDAPR) = ?) And ((VIAGVVC0.VVCIDADZ) = ?) And ((VIAGVVT0.VVTDTINV) BETWEEN ? AND ?)) " & _
                     "  ORDER BY VIAGVVT0.VVTDTINV"
    SQL_GET_SELECTION_CREATE_PPT = _
            "SELECT OCTIDOCT, ADZRASOC, OCRIDAPR, OCRQCOLT, " & _
            "       OCTDTRIC, OCTDTCRI, OCTRIFE1, OCTRIFE2, " & _
            "       OCRQCOKG, OCRQCOLT, OCRQCOPZ, OCRQCOCN, " & _
            "       SDZCODSO, OCRQRILT, OCRIDOCR, OCTKMRAG, " & _
            "       OCTIDAZO, OCTIDADZ, OCTIDASC, OCTIDAEE, " & _
            "       OCTIDADE, OCTIDACL, OCRIDAIM, OCRIDAPD, OCTRAUT3 " & _
            "  FROM ANAGADZ0 INNER JOIN ((ORDIOCT0 LEFT JOIN MULTSDZ0 " & _
            "       ON (ORDIOCT0.OCTIDASC = MULTSDZ0.SDZIDASC) AND " & _
            "          (ORDIOCT0.OCTIDADZ = MULTSDZ0.SDZIDADZ)) " & _
            "       LEFT JOIN ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT) " & _
            "       ON ANAGADZ0.ADZIDADZ = ORDIOCT0.OCTIDADZ "
    
    mvarFieldsIndexes = Array("0", "1", "2", "17", "12", "13", "5", _
                              "3", "8", "7", "18", "21", "23")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Dal Numero Ordine", "ordioct0.octidoct", 0, , 7, "0", " >= "
        .Add "Al Numero Ordine", "ordioct0.octidoct", 0, , 7, "20", " <= "
        .Add "Numero Ordine", "ordioct0.octidoct", 0, , 7, "24"
        .Add "Società", "ordioct0.octidasc", 0, , 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "ordioct0.octidaee", 0, , 3, "2"
        With .Item("2")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Dalla Data Ricezione", "ordioct0.octdtric", 2, , 10, "3", " >= "
        .Add "Alla Data Ricezione", "ordioct0.octdtric", 2, , 10, "30", " <= "
        .Add "Dalla Data Inserimento", "ordioct0.octdtins", 2, , 10, "5", " >= "
        .Add "Alla Data Inserimento", "ordioct0.octdtins", 2, , 10, "31", " <= "
        .Add "Dalla Data Consegna Richiesta", "ordioct0.octdtcri", 2, , 10, "7", " >= "
        .Add "Alla Data Consegna Richiesta", "ordioct0.octdtcri", 2, , 10, "29", " <= "
'        .Add "Data Consegna Pianificata", "ordioct0.octdtcpi", 2, , 10, "8"
        .Add "Dalla Data Consegna Pianificata", "ordioct0.octdtcpi", 2, , 10, "25", " >= "
        .Add "Alla Data Consegna Pianificata", "ordioct0.octdtcpi", 2, , 10, "26", " <= "
        
        
        .Add "Cliente", "ordioct0.octidacl", 0, , 10, "12"
        With .Item("12")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Destinazione", "ordioct0.octidadz", 0, , 10, "13"
        With .Item("13")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Deposito", "ordioct0.octidade", 0, , 3, "17"
        With .Item("17")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Zona", "ordioct0.octidazo", 0, , 3, "18"
        With .Item("18")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Tipo Ordine", "ordioct0.octfltor", 0, , 2, "21"
        .Add "Stato Ordine", "ordioct0.octflsor", 0, , 1, "22"
        .Add "Avanzamento Ordine", "ordioct0.octflaor", 0, , 1, "23"
        .Add "Riferimento Esterno 1", "ordioct0.octrife1", 0, , 15, "27"
        .Add "Riferimento Esterno 2", "ordioct0.octrife2", 0, , 15, "28"
        .Add "Dal Riferimento Esterno 3", "ordioct0.octrife3", 0, , 15, "44", " >= "
        .Add "Al Riferimento Esterno 3", "ordioct0.octrife3", 0, , 15, "45", " <= "
        .Add "Raggruppamento Utente 4", "ordioct0.octraut4", 0, , 2, "33"
        .Add "Raggruppamento Utente 5", "ordioct0.octraut5", 0, , 2, "34"
        .Add "Raggruppamento Ordine", "ordioct0.octidaro", 0, , 3, "32"
        With .Item("32")
            .ForeignTableEntity = "ARO"
            .ForeignTable = "anagaro0"
            .ForeignTableFld = "aroidaro"
            .ForeignTableDescField = "arodescr"
            .ForeignTableDeleteFld = "aroannul"
        End With
        .Add "Conferma cliente", "ordioct0.octflcnf", 0, False, 2, "59"
        
        .Add "Presidio", "anagazo0.azocodma", 0, , 15, "90"
        .Add "Presidio GC", "anagade0.adecodma", 0, , 15, "91"
        .Add "Quantita Iniziale L", "ordiocr0.ocrqcolt", 1, False, 15, "92"
        .Add "Quantita Richiesta L", "ordiocr0.ocrqrilt", 1, False, 15, "93"
        .Add "Prodotto", "ordiocr0.ocridapr", 0, False, 6, "94"
        .Add "Presidio RS3", "ordioct0.octidpre", 0, , 15, "89"
        '-- DP 19/08/08
        '-- Il seguenti 2 campi sono stati aggiunti per la funzione di ricerca riduzioni
        .Add "Differenza Litri", "ORDIOCR0.OCRQRILT - ORDIOCR0.OCRQCOLT", 1, False, 15, "35", " <> "
        
        .Add "Riga Annullata", "ordiocr0.ocrannul", 3, , , "95"
    End With
    DS.CancelField = "ordioct0.octannul"
    With DS.KeyFields
        .Add "0", "1"
        .Add "20", "2"
        .Add "1", "3"
        .Add "2", "4"
        .Add "3", "5"
        .Add "5", "6"
        .Add "7", "7"
'        .Add "8", "8"
        .Add "12", "9"
        .Add "13", "10"
        .Add "17", "11"
        .Add "18", "12"
        .Add "21", "13"
        .Add "22", "14"
        .Add "23", "15"
        .Add "24", "16"
        .Add "25", "17"
        .Add "26", "18"
        .Add "27", "19"
        .Add "28", "20"
        .Add "29", "21"
        .Add "30", "22"
        .Add "31", "23"
        .Add "32", "24"
        .Add "90", "25"
        .Add "91", "26"
        .Add "92", "27"
        .Add "93", "28"
        .Add "33", "29"
        .Add "34", "30"
        .Add "44", "31"
        .Add "45", "32"
        .Add "95", "33"
        .Add "35", "34"
        .Add "59", "35"
        .Add "89", "36"
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
