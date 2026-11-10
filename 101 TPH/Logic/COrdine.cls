VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrdine"
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
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_SET_AVANZAMENTO As String
Public SQL_SET_AVANZAMENTO_RIGHE As String
Public SQL_GET_TOTALE_QTA As String
Public SQL_GET_TOTALE_QTA_DEST As String
Public SQL_GET_ID_VIAGGIO As String
Public SQL_CHECK_SOC_CODE As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Private mvarTotaliQta As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private Const SQL_INVALIDA_ORDINI = _
"UPDATE ORDIOCT0 " & _
"   SET OCTRAUT4 = '*' " & _
" WHERE OCTIDASC = ? AND OCTDTRIC = ? AND OCTIDATV = 'TD' AND OCTFLAOR = 'A' AND OCTANNUL = 0"

Private Const SQL_ANNULLA_ORDINI_TD = _
"UPDATE ORDIOCT0 " & _
"   SET OCTANNUL = 1 " & _
" WHERE OCTIDASC = ? AND OCTDTRIC = ? AND OCTIDATV = 'TD' AND OCTFLAOR = 'A' AND OCTRAUT4 = '*'"

Private Const SQL_GET_ID_BY_SOC_CODE = _
"SELECT octidoct " & _
"  FROM ordioct0 " & _
" WHERE octidasc = '001' AND octrife1 = "

Private Const SQL_DELETE_ROWS = _
"DELETE FROM ORDIOCR0 " & _
" WHERE ocridoct = ? "

Private Const SQL_COPY_ORDER = _
"INSERT INTO ORDIOCT0 " & _
    "(octidoct, octraut5, octidasc, octidaee, octdtric, " & _
               "       octhrric, octdtins, octhrins, " & _
               "       octdtcef, octrife1, octrife2, " & _
               "       octrife3, octrife4, octidacl, octidadz, " & _
               "       octnota1, octnota2, octidatv, octidade, " & _
               "       octidazo, octidamp, octidaro, octfltor, " & _
               "       octflsor, octflaor, octflpor, octraut1, " & _
               "       octraut2, octraut3, octflorm, " & _
               "       octnota3, octkmrag, octhrfr1, octhrto1, " & _
               "       octhrfr2, octhrto2, octhrfr3, octhrto3, " & _
               "       octnonav, octdtgra, octflcim, octflcpr, " & _
               "       octidpre, octconte, octannul, octidoc1, octraut4, octdtcri, octdtcpi) " & _
"SELECT "

Private Const SQL_COPY_ORDER_2 = _
    ", octraut5, octidasc, octidaee, octdtric, " & _
               "       octhrric, octdtins, octhrins, " & _
               "       octdtcef, octrife1, octrife2, " & _
               "       octrife3, octrife4, octidacl, octidadz, " & _
               "       octnota1, octnota2, octidatv, octidade, " & _
               "       octidazo, octidamp, octidaro, octfltor, " & _
               "       octflsor, octflaor, octflpor, octraut1, " & _
               "       octraut2, octraut3, octflorm, " & _
               "       octnota3, octkmrag, octhrfr1, octhrto1, " & _
               "       octhrfr2, octhrto2, octhrfr3, octhrto3, " & _
               "       octnonav, octdtgra, octflcim, octflcpr, " & _
               "       octidpre, octconte, octannul, "
               
Private Const SQL_COPY_ORDER_2_QTA = _
    ", NULL, octidasc, octidaee, octdtric, " & _
               "       octhrric, octdtins, octhrins, " & _
               "       octdtcef, octrife1, octrife2, " & _
               "       octrife3, octrife4, octidacl, octidadz, " & _
               "       octnota1, octnota2, octidatv, octidade, " & _
               "       octidazo, octidamp, octidaro, octfltor, " & _
               "       octflsor, octflaor, octflpor, octraut1, " & _
               "       octraut2, octraut3, octflorm, " & _
               "       octnota3, octkmrag, octhrfr1, octhrto1, " & _
               "       octhrfr2, octhrto2, octhrfr3, octhrto3, " & _
               "       octnonav, octdtgra, octflcim, octflcpr, " & _
               "       octidpre, octconte, octannul, "

Private Const SQL_COPY_ORDER_3 = _
" FROM ORDIOCT0 " & _
"WHERE OCTIDOCT =  "

Private Const SQL_COPY_ROWS_1 = _
"INSERT INTO ORDIOCR0 " & _
    "(ocridoct, ocridocr, ocridapr, ocridaim, " & _
    "       ocridapd, ocrdesri, ocrqrikg, ocrqrilt, " & _
    "       ocrqcolt, ocrqripz, ocrqricn, ocrqcokg, " & _
    "       ocrqcopz, ocrqcocn, ocridaiv, ocridamp, " & _
    "       ocrprezz, ocrscper, ocrscval, ocrfltir, " & _
    "       ocrflstr, ocrflavr, ocrrife1, ocrrife2, " & _
    "       ocrragg1, ocrragg2, ocrragg3, ocridaum, ocrannul, " & _
    "       ocridot1, ocridor1) " & _
"SELECT "

Private Const SQL_COPY_ROWS = _
"INSERT INTO ORDIOCR0 " & _
    "(ocridoct, ocridocr, ocridapr, ocridaim, " & _
    "       ocridapd, ocrdesri, ocrqrikg, " & _
    "       ocrqripz, ocrqricn, ocrqcokg, " & _
    "       ocrqcopz, ocrqcocn, ocridaiv, ocridamp, " & _
    "       ocrprezz, ocrscper, ocrscval, ocrfltir, " & _
    "       ocrflstr, ocrflavr, ocrrife1, ocrrife2, " & _
    "       ocrragg1, ocrragg2, ocrragg3, ocridaum, ocrannul, " & _
    "       ocridot1, ocridor1, ocrqrilt, ocrqcolt) " & _
"SELECT "

Private Const SQL_COPY_ROWS_2 = _
    ", ocridocr, ocridapr, ocridaim, " & _
    "       ocridapd, ocrdesri, ocrqrikg, ocrqrilt, " & _
    "       ocrqcolt, ocrqripz, ocrqricn, ocrqcokg, " & _
    "       ocrqcopz, ocrqcocn, ocridaiv, ocridamp, " & _
    "       ocrprezz, ocrscper, ocrscval, ocrfltir, " & _
    "       ocrflstr, ocrflavr, ocrrife1, ocrrife2, " & _
    "       ocrragg1, ocrragg2, ocrragg3, ocridaum, ocrannul, "
Private Const SQL_COPY_ROWS_2_QTA = _
    ", ocridocr, ocridapr, ocridaim, " & _
    "       ocridapd, ocrdesri, ocrqrikg, " & _
    "       ocrqripz, ocrqricn, ocrqcokg,  " & _
    "       ocrqcopz, ocrqcocn, ocridaiv, ocridamp, " & _
    "       ocrprezz, ocrscper, ocrscval, ocrfltir, " & _
    "       ocrflstr, ocrflavr, ocrrife1, ocrrife2, " & _
    "       ocrragg1, ocrragg2, ocrragg3, ocridaum, ocrannul, "

Private Const SQL_COPY_ROWS_3 = _
"FROM ORDIOCR0 WHERE OCRIDOCT = "

Private Const SQL_COPY_ROWS_4 = _
"AND OCRIDOCR = "

Private Const SQL_GET_MAX_PARTIZIONE = _
    " SELECT MAX(OCTRAUT4) " & _
    "   FROM ORDIOCT0 " & _
    "  WHERE OCTIDOC1 = ? AND OCTANNUL = 0"
Public Function CopyOrder(IDOCT As String, _
                          Optional AppendRows As Boolean = True, _
                          Optional ByRef IDNewOct As String, _
                          Optional CreaPartizione As Boolean = False, _
                          Optional IDOCR As String, _
                          Optional Quantità As Variant, _
                          Optional Data As Variant, _
                          Optional ByRef ErrDesc As String) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim varCounter As Variant
    Dim varBuf As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim varCountSun As Variant
    Dim blnConcurError As Boolean
    Dim strPartizione As String
    Dim strError As String
    Dim OCT As COrdine
    Dim OCR As COrdineRiga
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    Set OCR = New COrdineRiga
    With OCR.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    DS.DA.BeginTransaction
    
    If Not IsMissing(CreaPartizione) Then
        If CreaPartizione Then
            If Not IsNumeric(Quantità) Then
                GoTo ErrHandler
            End If
            If Not IsDate(Data) Then
                GoTo ErrHandler
            End If
            
            With OCR.DS
                
                .Field("Numero Ordine").FindValue = IDOCT
                .Field("Numero Riga").FindValue = IDOCR
                
                OCR.Find
                
                If CLng(Quantità) > .Field("Quantita Richiesta L").Value Then
                    strError = "La quantità supera la quantità residua dell'ordine"
                    GoTo ErrHandler
                End If
                If CLng(Quantità) > .Field("Quantita Iniziale L").Value Then
                    strError = "La quantità supera la quantità dell'ordine"
                    GoTo ErrHandler
                End If
            End With
                                
            If Not DS.DA.QueryCommand(SQL_GET_MAX_PARTIZIONE, varBuf, lngHowMany, IDOCT) Then
                GoTo ErrHandler
            End If
            If Not IsEmpty(varBuf) Then
                If IsNumeric(varBuf(0, 0)) Then
                    If varBuf(0, 0) >= 99 Then
                        GoTo ErrHandler
                    End If
                    strPartizione = Format(varBuf(0, 0) + 1, "00")
                Else
                    strPartizione = "01"
                End If
            Else
                strPartizione = "01"
            End If
        End If
    End If
    
    
    varCounter = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "001", 0)
    If varCounter = -1 Then
        GoTo ErrHandler
    End If
    
    strQ = SQL_COPY_ORDER & DS.DA.MakeSQL(varCounter, dtAlfa)
    If CreaPartizione Then
        strQ = strQ & SQL_COPY_ORDER_2_QTA & " OCTIDOCT, " & DS.DA.MakeSQL(strPartizione, dtAlfa)
        strQ = strQ & ", " & DS.DA.MakeSQL(Data, dtDate) & ", " & DS.DA.MakeSQL(Data, dtDate)
    Else
        strQ = strQ & SQL_COPY_ORDER_2 & " OCTIDOCT, " & DS.DA.MakeSQL(strPartizione, dtAlfa)
        strQ = strQ & ", " & DS.DA.MakeSQL(Data, dtDate) & ", " & DS.DA.MakeSQL(Data, dtDate)
    End If
    strQ = strQ & SQL_COPY_ORDER_3 & DS.DA.MakeSQL(IDOCT, dtAlfa)
    
    If Not DS.DA.ExecuteQuery(strQ) Then
        GoTo ErrHandler
    End If
    If AppendRows Then
        If Not IsMissing(Quantità) Then
            strQ = SQL_COPY_ROWS & DS.DA.MakeSQL(varCounter, dtAlfa)
        Else
            strQ = SQL_COPY_ROWS_1 & DS.DA.MakeSQL(varCounter, dtAlfa)
        End If

        If Not IsMissing(Quantità) Then
            strQ = strQ & SQL_COPY_ROWS_2_QTA & " OCRIDOCT, OCRIDOCR "
        Else
            strQ = strQ & SQL_COPY_ROWS_2 & " OCRIDOCT, OCRIDOCR "
        End If
        If Not IsMissing(Quantità) Then
            strQ = strQ & ", " & DS.DA.MakeSQL(Quantità, dtNum) & ", " & DS.DA.MakeSQL(Quantità, dtNum) & " "
        End If
               
        strQ = strQ & SQL_COPY_ROWS_3 & DS.DA.MakeSQL(IDOCT, dtAlfa)
        If Not IsMissing(IDOCR) Then
            If Len(IDOCR) > 0 Then
                strQ = strQ & SQL_COPY_ROWS_4 & DS.DA.MakeSQL(IDOCR, dtAlfa)
            End If
        End If
        If Not DS.DA.ExecuteQuery(strQ) Then
            GoTo ErrHandler
        End If
    End If
    
    If CreaPartizione Then
        OCR.DS.Field("Quantita Richiesta L").NewValue = OCR.DS.Field("Quantita Richiesta L").Value - CLng(Quantità)
        OCR.DS.Field("Quantita Richiesta L").Updated = True
        OCR.BeginSave blnConcurError
        OCR.DS.DA.EndTrans etCommit
        OCR.DS.EndSave
        
        OCT.DS.Field("Numero Ordine").FindValue = IDOCT
        
        OCT.Find
        
        OCT.DS.Field("Raggruppamento Utente 5").NewValue = "S"
        OCT.DS.Field("Raggruppamento Utente 5").Updated = True
        OCT.BeginSave blnConcurError
        OCT.DS.DA.EndTrans etCommit
        OCT.DS.EndSave
    End If
    
    DS.DA.EndTrans etCommit
    
    IDNewOct = varCounter
    CopyOrder = True
    
    Exit Function
    
ErrHandler:
    DS.DA.EndTrans etRollback
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    CopyOrder = False
End Function

Public Function CopyOrderRow(varIDOCT As Variant, varIDOCR As Variant, varIDOCTNew As Variant) As Boolean
On Error GoTo ErrHandler
    Dim strQ As String
    
    strQ = SQL_COPY_ROWS & DS.DA.MakeSQL(varIDOCTNew, dtAlfa) & _
           SQL_COPY_ROWS_2 & " OCRIDOCT, OCRIDOCR " & _
           SQL_COPY_ROWS_3 & DS.DA.MakeSQL(varIDOCT, dtAlfa) & _
           " AND OCRIDOCR = " & DS.DA.MakeSQL(varIDOCR, dtNum)
    If Not DS.DA.ExecuteQuery(strQ) Then
        GoTo ErrHandler
    End If
    
    CopyOrderRow = True
    
    Exit Function
    
ErrHandler:
    CopyOrderRow = False
End Function

Public Function CheckSocCode(OrdCode As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    CheckSocCode = True
    strSQL = SQL_CHECK_SOC_CODE & DS.DA.MakeSQL(OrdCode, dtAlfa)
    strSQL = strSQL & " AND " & DS.CancelField & " = " & DS.DA.MakeSQL(False, dtBit)
    blnResult = DS.DA.QueryData(strSQL, varBuffer)
    If blnResult Then
        CheckSocCode = (varBuffer(0, 0) > 0)
    End If
    Exit Function
ErrHandler:
End Function

'-- DP 28/11/08
'-- Aggiunto strIDADE come parametro per la gestione degli ordini multideposito
Public Function CheckSocCode1(ByVal pSoc As String, ByVal OrdCode As String, Optional NumPartizione As String = "00", Optional ByRef ErrDesc As String, Optional strIDADE As String) As Boolean
    Const SQL_CHECK_SOC_CODE_1 = _
    "SELECT Count(*) FROM ordioct0 WHERE octidasc = "
    Const SQL_CHECK_SOC_CODE_1_AND = _
    " AND octrife1 = "
    Const SQL_CHECK_SOC_CODE_1_AND_2 = _
    " AND CONVERT(INT,octraut4) = "
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim strError As String
    
    CheckSocCode1 = True
    strSQL = SQL_CHECK_SOC_CODE_1 & DS.DA.MakeSQL(pSoc, dtAlfa) & SQL_CHECK_SOC_CODE_1_AND & DS.DA.MakeSQL(OrdCode, dtAlfa)
    If NumPartizione <> "00" Then
        strSQL = strSQL & SQL_CHECK_SOC_CODE_1_AND_2 & DS.DA.MakeSQL(NumPartizione, dtNum)
    End If
    strSQL = strSQL & " AND " & DS.CancelField & " = " & DS.DA.MakeSQL(False, dtBit)
    
    '-- DP 28/11/08
    '-- Gestione ordini multideposito
    If strIDADE <> vbNullString Then
         strSQL = strSQL & " AND OCTIDADE = " & DS.DA.MakeSQL(strIDADE, dtAlfa)
    End If
    
    blnResult = DS.DA.QueryData(strSQL, varBuffer)
    If blnResult Then
        If varBuffer(0, 0) = 0 Then
            ' Non abbiamo trovato l'ordine e la partizione: cerchiamo solamente l'ordine
            strSQL = SQL_CHECK_SOC_CODE_1 & DS.DA.MakeSQL(pSoc, dtAlfa) & SQL_CHECK_SOC_CODE_1_AND & DS.DA.MakeSQL(OrdCode, dtAlfa)
            
            '-- DP 28/11/08
            '-- Gestione ordini multideposito
            strSQL = strSQL & " AND " & DS.CancelField & " = " & DS.DA.MakeSQL(False, dtBit)
            If strIDADE <> vbNullString Then
                 strSQL = strSQL & " AND OCTIDADE = " & DS.DA.MakeSQL(strIDADE, dtAlfa)
            End If
            
            varBuffer = Empty
            blnResult = DS.DA.QueryData(strSQL, varBuffer)
            If Not blnResult Then
                strError = "Errore durante la ricerca dell'ordine"
                GoTo ErrHandler
            End If
            If varBuffer(0, 0) > 0 Then
                strError = "Partizione non trovata"
                GoTo ErrHandler
            Else
                strError = "Ordine non trovato"
                GoTo ErrHandler
            End If
        Else
            CheckSocCode1 = True
        End If
    Else
        strError = "Errore durante la ricerca dell'ordine"
        GoTo ErrHandler
    End If
    Exit Function
ErrHandler:
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    CheckSocCode1 = False
End Function

Public Function DeleteRows(IDOrdine As String) As Boolean
    On Error GoTo ErrHandler
    
    If Not DS.DA.ExecuteCommand(SQL_DELETE_ROWS, _
                                IDOrdine) Then
        GoTo ErrHandler
    End If
    DeleteRows = True
    Exit Function
ErrHandler:
    DeleteRows = False
End Function

Public Function GetIDBySocCode(OrdCode As String) As String
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    strSQL = SQL_GET_ID_BY_SOC_CODE & DS.DA.MakeSQL(OrdCode, dtAlfa)
    strSQL = strSQL & " AND " & DS.CancelField & " = " & DS.DA.MakeSQL(False, dtBit)
    blnResult = DS.DA.QueryData(strSQL, varBuffer)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetIDBySocCode = varBuffer(0, 0)
    Exit Function
ErrHandler:
    GetIDBySocCode = vbNullString
End Function

Public Sub GetTotali()
    Dim strQ As String
    Dim varResult As Variant
    Dim blnOk As Boolean
    Dim lngHowMany As Long
    Dim dblTot As Double
    Dim i As Integer
    
    With DS.DA
        strQ = SQL_GET_TOTALE_QTA & .MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
        blnOk = .QueryData(strQ, varResult, lngHowMany)
        If lngHowMany > 0 Then
            For i = LBound(varResult, 1) To UBound(varResult, 1)
                If IsNull(varResult(i, 0)) Then
                    varResult(i, 0) = 0
                End If
            Next i
            mvarTotaliQta = varResult
        
        End If
        
    End With
End Sub

Public Sub GetTotaliDestinazione()
    Dim strQ As String
    Dim varResult As Variant
    Dim blnOk As Boolean
    Dim lngHowMany As Long
    Dim dblTot As Double
    
    With DS.DA
        strQ = SQL_GET_TOTALE_QTA_DEST
        blnOk = .QueryCommand(strQ, varResult, lngHowMany, _
                               DS.Fields("13").Value, _
                               DS.Fields("7").Value)
        If lngHowMany > 0 Then
            mvarTotaliQta = varResult
        End If
    End With
End Sub


Public Function GetTripID() As String
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    strQ = SQL_GET_ID_VIAGGIO & DS.DA.MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
    strQ = strQ & " AND vvcannul = " & DS.DA.MakeSQL(False, 3)
    If Not DS.DA.QueryData(strQ, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If
    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If
    GetTripID = varBuffer(0, 0)
    Exit Function
ErrHandler:
    GetTripID = vbNullString
End Function

Public Function InvalidaOrdini(Società As Variant, DataRicezione As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    
    blnResult = DS.DA.ExecuteCommand(SQL_INVALIDA_ORDINI, _
                                     Società, _
                                     DataRicezione)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    InvalidaOrdini = True
    Exit Function
ErrHandler:
    InvalidaOrdini = False
End Function

Public Function AnnullaOrdiniTD(Società As Variant, DataRicezione As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    
    blnResult = DS.DA.ExecuteCommand(SQL_ANNULLA_ORDINI_TD, _
                                     Società, _
                                     DataRicezione)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    AnnullaOrdiniTD = True
    Exit Function
ErrHandler:
    AnnullaOrdiniTD = False
End Function

Public Function AnnullaOrdine(varIDOCT As Variant) As Boolean
On Error GoTo ErrHandler
    
    Dim strSQLUpdate As String
    
    strSQLUpdate = "UPDATE ORDIOCT0 SET OCTANNUL = 1 WHERE OCTIDOCT = ?"
    
    If IsEmptyVariantString(varIDOCT) Then
        GoTo ErrHandler
    End If
    
    If Not DS.DA.ExecuteCommand(strSQLUpdate, varIDOCT) Then
        GoTo ErrHandler
    End If

    AnnullaOrdine = True
    Exit Function

ErrHandler:
    AnnullaOrdine = False
End Function

Public Property Get TotaliQta(Index As Integer) As Variant
    If (Index >= LBound(mvarTotaliQta, 1)) And _
        (Index <= UBound(mvarTotaliQta, 1)) Then
        TotaliQta = mvarTotaliQta(Index, 0)
    End If
End Property

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
    Insert = DS.Insert(Me.SQL_INSERT, "anaganu0", "anuprogr", "anuidanu", "001", 0)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim strAvanzamento As String
    Dim varQta As Variant
    Dim varRigheOrdine As Variant
    Dim OCR As COrdineRiga
    Dim OCT As COrdine
    Dim strIDOCTRif As String
    Dim strIDAPR As String
    Dim blnConcErr As Boolean
    Dim i As Long
    
    ' --- Impedisco il salvataggio di ordini il cui stato avanzamento
    ' --- sia "V" o "F"
    strAvanzamento = DS.Fields("23").Value
    If (strAvanzamento = "A") Or (strAvanzamento = "M") Or (strAvanzamento = "V") Then
        ' Se annulliamo o ripristiniamo un ordine che deriva da un ordine "aperto"
        ' modifichiamo di conseguenza la quantità residua dell'ordine "aperto"
        If DS.Field("Annullato").Updated Then
            If (Len(DS.Field("Ordine di Riferimento").Value) > 0) And (Len(DS.Field("Raggruppamento Utente 4").Value) > 0) Then
                strIDOCTRif = DS.Field("Ordine di Riferimento").Value
                Set OCR = New COrdineRiga
                With OCR.DS
                    Set .DA = DS.DA
                    Set .ErrH = DS.ErrH
                End With
                OCR.DS.Field("Numero Ordine").FindValue = DS.Field("Numero Ordine").Value
                OCR.GetList True
                varRigheOrdine = OCR.DS.List
                For i = 0 To UBound(varRigheOrdine, 2)
                    strIDAPR = varRigheOrdine(1, i)
                    OCR.DS.Fields.ResetFields
                    OCR.DS.Field("Numero Ordine").FindValue = strIDOCTRif
                    OCR.DS.Field("Prodotto").FindValue = strIDAPR
                    OCR.Find
                    If DS.Field("Annullato").NewValue = True Then
                        OCR.DS.Field("Quantita Richiesta L").NewValue = OCR.DS.Field("Quantita Richiesta L").Value + varRigheOrdine(4, i)
                    Else
                        OCR.DS.Field("Quantita Richiesta L").NewValue = OCR.DS.Field("Quantita Richiesta L").Value - varRigheOrdine(4, i)
                    End If
                    OCR.DS.Field("Quantita Richiesta L").Updated = True
                    OCR.BeginSave blnConcErr
                    OCR.DS.DA.EndTrans etCommit
                Next i
            End If
        End If
        DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
    Else
        BeginSave = False
        Exit Function
    End If
    BeginSave = True
    ' ---
End Function

Public Function SetAvanzamento(Avanzamento As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strTS As String
    
    DS.DA.BeginTransaction
    strQ = SQL_SET_AVANZAMENTO & DS.DA.MakeSQL(Avanzamento, 0)
    strTS = DS.AddTSUpdate
    If strTS <> vbNullString Then
        strQ = strQ & ", " & strTS
    End If
    strQ = strQ & " WHERE octidoct = " & DS.DA.MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
    If Not DS.DA.ExecuteQuery(strQ) Then
        GoTo ErrHandler
    End If
    
    If Avanzamento = "V" Or Avanzamento = "F" Then
        strQ = SQL_SET_AVANZAMENTO_RIGHE & DS.DA.MakeSQL(Avanzamento, 0)
        strTS = DS.AddTSUpdate
        If strTS <> vbNullString Then
            strQ = strQ & ", " & strTS
        End If
        strQ = strQ & " WHERE ocridoct = " & DS.DA.MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
        strQ = strQ & "   AND ocrflstr = " & DS.DA.MakeSQL("A", 0)
        strQ = strQ & "   AND ocrannul = " & DS.DA.MakeSQL(False, 3)
        
        If Not DS.DA.ExecuteQuery(strQ) Then
            GoTo ErrHandler
        End If
    End If
    DS.DA.EndTrans 0
    SetAvanzamento = True
    
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    SetAvanzamento = False
End Function


Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT octidoct, octidasc, octidaee, octdtric, " & _
               "       octhrric, octdtins, octhrins, octdtcri, " & _
               "       octdtcpi, octdtcef, octrife1, octrife2, " & _
               "       octrife3, octrife4, octidacl, octidadz, " & _
               "       octnota1, octnota2, octidatv, octidade, " & _
               "       octidazo, octidamp, octidaro, octfltor, " & _
               "       octflsor, octflaor, octidoc1, octflpor, " & _
               "       octraut1, octraut2, octraut3, octraut4, " & _
               "       octraut5, octflorm, octnota3, octkmrag, " & _
               "       octhrfr1, octhrto1, octhrfr2, octhrto2, " & _
               "       octhrfr3, octhrto3, octconte, octcgpsx, " & _
               "       octcgpsy, octconno, octcontl, octnonav, " & _
               "       octdtgra, octflcim, octflcpr, octflres, " & _
               "       octflver, octtiser, octflazi, octflcnf, " & _
               "       octflinp, octhrap1, octhrch1, octhrap2, " & _
               "       octhrch2, octhrap3, octhrch3, octidpre, " & _
               "       octannul, timestmp FROM ordioct0 WHERE "
               
    SQL_GETLIST = "SELECT octidoct, octidasc FROM ordioct0 "
    SQL_UPDATE = "UPDATE ordioct0 SET "
    SQL_INSERT = "INSERT INTO ordioct0 "
    SQL_GET_TS = "SELECT timestmp FROM ordioct0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT octidoct, " & _
               "adzrasoc, adzlocal, sum(ocrqrilt) AS TotQ, octdtcri, octidazo, " & _
               "octidade, octfltor, octflaor, octflsor, octannul " & _
               " FROM ANAGADZ0 INNER JOIN " & _
               "      (ORDIOCT0 LEFT JOIN ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT " & _
               "       INNER JOIN ANAGAZO0 ON OCTIDAZO = AZOIDAZO INNER JOIN ANAGADE0 ON OCTIDADE = ADEIDADE ) " & _
               "      ON ANAGADZ0.ADZIDADZ = ORDIOCT0.OCTIDADZ "
    SQL_GET_SEARCH_LIST_GROUPBY = "GROUP BY ORDIOCT0.OCTIDOCT, ANAGADZ0.ADZRASOC, ANAGADZ0.ADZLOCAL, ORDIOCT0.OCTDTCRI, ORDIOCT0.OCTIDAZO, ORDIOCT0.OCTIDADE, " & _
                                  " ORDIOCT0.OCTFLTOR, ORDIOCT0.OCTFLAOR, ORDIOCT0.OCTFLSOR, ORDIOCT0.OCTANNUL"
    SQL_SET_AVANZAMENTO = "UPDATE ordioct0 SET octflaor = "
    SQL_SET_AVANZAMENTO_RIGHE = "UPDATE ordiocr0 SET ocrflavr = "
    SQL_GET_TOTALE_QTA = "SELECT SUM(ocrqrikg),SUM(ocrqrilt),SUM(ocrqripz),SUM(ocrqricn), " & _
                         "       SUM(ocrqcokg),SUM(ocrqcolt),SUM(ocrqcopz),SUM(ocrqcocn) " & _
                         "  FROM ordiocr0 WHERE ocridoct = "
    SQL_GET_ID_VIAGGIO = "SELECT vvcidvvt FROM viagvvc0 WHERE vvcidoct = "
    SQL_CHECK_SOC_CODE = "SELECT Count(*) FROM ordioct0 WHERE octidasc = '001' AND octrife1 = "
    SQL_GET_TOTALE_QTA_DEST = "SELECT SUM(ocrqrikg),SUM(ocrqrilt),SUM(ocrqripz),SUM(ocrqricn), " & _
                              "       SUM(ocrqcokg),SUM(ocrqcolt),SUM(ocrqcopz),SUM(ocrqcocn) " & _
                              "  FROM ORDIOCT0 INNER JOIN ORDIOCR0 ON (ordioct0.octIDoct = ordiocr0.ocrIDoct) " & _
                              " WHERE octidadz = ? AND octdtcri = ? AND octannul = 0 AND ocrannul = 0"
            
    mvarFieldsIndexes = Array("0", "33", "34", "35", "7", "18", _
                              "17", "21", "23", "22", "32", "90", "91")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Ordine", "ordioct0.octidoct", 0, True, 7, "0"
        .Add "Società", "ordioct0.octidasc", 0, True, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "ordioct0.octidaee", 0, True, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Data Ricezione", "ordioct0.octdtric", 2, False, 10, "3"
        .Add "Ora Ricezione", "ordioct0.octhrric", 4, False, 5, "4"
        .Add "Data Inserimento", "ordioct0.octdtins", 2, False, 10, "5"
        .Add "Ora Inserimento", "ordioct0.octhrins", 4, False, 5, "6"
        .Add "Data Consegna Richiesta", "ordioct0.octdtcri", 2, False, 10, "7"
        .Add "Data Consegna Pianificata", "ordioct0.octdtcpi", 2, False, 10, "8"
        .Add "Data Consegna Effettiva", "ordioct0.octdtcef", 2, False, 10, "9"
        .Add "Rif. Esterno 1", "ordioct0.octrife1", 0, False, 15, "10"
        .Add "Rif. Esterno 2", "ordioct0.octrife2", 0, False, 15, "11"
        .Add "Rif. Esterno 3", "ordioct0.octrife3", 0, False, 15, "44"
        .Add "Rif. Esterno 4", "ordioct0.octrife4", 0, False, 15, "45"
        .Add "Cliente", "ordioct0.octidacl", 0, True, 10, "12"
        With .Item("12")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "anagacl0.aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Destinazione", "ordioct0.octidadz", 0, True, 10, "13"
        With .Item("13")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Nota 1", "ordioct0.octnota1", 0, False, 40, "14"
        .Add "Nota 2", "ordioct0.octnota2", 0, False, 40, "15", , True
        .Add "Tipo Vendita", "ordioct0.octidatv", 0, True, 3, "16"
        With .Item("16")
            .ForeignTableEntity = "ATV"
            .ForeignTable = "anagatv0"
            .ForeignTableFld = "atvidatv"
            .ForeignTableDescField = "atvdescr"
            .ForeignTableDeleteFld = "atvannul"
        End With
'        .Add "Deposito", "ordioct0.octidade", 0, True, 3, "17"
        .Add "Deposito", "ordioct0.octidade", 0, True, 3, "17"
        With .Item("17")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Zona", "ordioct0.octidazo", 0, True, 3, "18"
        With .Item("18")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Modalità di Pagamento", "ordioct0.octidamp", 0, , 3, "19"
        With .Item("19")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "Raggruppamento Ordine", "ordioct0.octidaro", 0, True, 3, "20"
        With .Item("20")
            .ForeignTableEntity = "ARO"
            .ForeignTable = "anagaro0"
            .ForeignTableFld = "aroidaro"
            .ForeignTableDescField = "arodescr"
            .ForeignTableDeleteFld = "aroannul"
        End With
        .Add "Tipo Ordine", "ordioct0.octfltor", 0, True, 2, "21"
        .Add "Stato Ordine", "ordioct0.octflsor", 0, True, 1, "22"
        .Add "Avanzamento Ordine", "ordioct0.octflaor", 0, True, 1, "23"
        .Add "Ordine di Riferimento", "ordioct0.octidoc1", 0, False, 7, "24"
        .Add "Priorità", "ordioct0.octflpor", 0, False, 1, "25"
        
        ' Pedidos Eventuales
        .Add "Raggruppamento Utente 1", "ordioct0.octraut1", 0, False, 2, "26"
        
        .Add "Raggruppamento Utente 2", "ordioct0.octraut2", 0, False, 2, "27", , True
        .Add "Raggruppamento Utente 3", "ordioct0.octraut3", 0, False, 2, "28", , True
        
        ' Partizione ordine
        .Add "Raggruppamento Utente 4", "ordioct0.octraut4", 0, False, 2, "29"
        
        ' Ordine "aperto"
        .Add "Raggruppamento Utente 5", "ordioct0.octraut5", 0, False, 2, "30"
        
        .Add "Ordine Modificato", "ordioct0.octflorm", 0, True, 1, "31"
        .Add "Nota 3", "ordioct0.octnota3", 0, False, 40, "36"
        .Add "Km Raggio", "ordioct0.octkmrag", 1, False, 6, "37"
        .Add "Dalle Ore 1", "ordioct0.octhrfr1", 4, False, 5, "38"
        .Add "Alle Ore 1", "ordioct0.octhrto1", 4, False, 5, "39"
        .Add "Dalle Ore 2", "ordioct0.octhrfr2", 4, False, 5, "40"
        .Add "Alle Ore 2", "ordioct0.octhrto2", 4, False, 5, "41"
        .Add "Dalle Ore 3", "ordioct0.octhrfr3", 4, False, 5, "42"
        .Add "Alle Ore 3", "ordioct0.octhrto3", 4, False, 5, "43"
        .Add "Condizionante tecnico", "ordioct0.octconte", 0, , 10, "46"
        .Add "Coordinate GPS X", "ordioct0.octcgpsx", 0, False, 15, "47"
        .Add "Coordinate GPS Y", "ordioct0.octcgpsy", 0, False, 15, "48"
        .Add "Consegnatario Ragione Sociale", "ordioct0.octconno", 0, False, 40, "49"
        .Add "Consegnatario Telefono", "ordioct0.octcontl", 0, False, 20, "50"
        .Add "Nome nave", "ordioct0.octnonav", 0, False, 40, "51"
        .Add "Data Grado", "ordioct0.octdtgra", 2, False, 10, "52"
        .Add "Controllo immagine", "ordioct0.octflcim", 0, False, 2, "53", , True
        .Add "Controllo precincto", "ordioct0.octflcpr", 0, False, 2, "54", , True
        .Add "Responsabilità riduzione", "ordioct0.octflres", 0, False, 2, "55"
        .Add "Venta en Ruta", "ordioct0.octflver", 0, False, 2, "56"
        .Add "Tipo servizio", "ordioct0.octtiser", 0, False, 3, "57", , True
        
        .Add "Azione consentita", "ordioct0.octflazi", 0, False, 2, "58", , True
        .Add "Conferma cliente", "ordioct0.octflcnf", 0, False, 2, "59", , True
        .Add "Accordi con cliente", "ordioct0.octflinp", 0, False, 2, "60"
        .Add "Ora Apertura 1", "ordioct0.octhrap1", 4, False, 5, "61", , True
        .Add "Ora Chiusura 1", "ordioct0.octhrch1", 4, False, 5, "62", , True
        .Add "Ora Apertura 2", "ordioct0.octhrap2", 4, False, 5, "63", , True
        .Add "Ora Chiusura 2", "ordioct0.octhrch2", 4, False, 5, "64", , True
        .Add "Ora Apertura 3", "ordioct0.octhrap3", 4, False, 5, "65", , True
        .Add "Ora Chiusura 3", "ordioct0.octhrch3", 4, False, 5, "66", , True
        .Add "Presidio RS3", "ordioct0.octidpre", 0, False, 8, "67", , True
        .Add "Annullato", "ordioct0.octannul", 3, , , "32"
                
        .Add "Ragione Sociale", "anagadz0.adzrasoc", 0, False, 0, "33", , True
        .Add "Località", "anagadz0.adzlocal", 0, False, 0, "34", , True
        .Add "Tot. Q.ta", "TotQ", 1, False, 0, "35", , True
        .Add "Presidio", "anagazo0.azocodma", 0, , 15, "90"
        .Add "Presidio GC", "anagade0.adecodma", 0, , 15, "91"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "ordioct0.octannul"
    
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

Public Function IsOrdineCopiato(varIDOCT As Variant, ByRef blnOrdineCopiato As Boolean) As Boolean
On Error GoTo ErrHandler

    Dim strSQLSelect As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If IsEmptyVariantString(varIDOCT) Then
        GoTo ErrHandler
    End If
    
    strSQLSelect = "SELECT COUNT(*) FROM ORDIOCT0 WHERE OCTANNUL = " & DS.DA.MakeSQL(False, dtBit) & _
        " AND OCTIDOC1 = ?"

    lngHowMany = 0
    If Not DS.DA.QueryCommand(strSQLSelect, varBuffer, lngHowMany, varIDOCT) Then
        GoTo ErrHandler
    End If

    blnOrdineCopiato = IIf(varBuffer(0, 0) > 0, True, False)

    IsOrdineCopiato = True
    Exit Function

ErrHandler:
    IsOrdineCopiato = False
    
End Function
