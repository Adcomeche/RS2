VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioSelezione2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
Public SQL_GET_SELECTION_EXPORT As String
Public SQL_GET_SELECTION_EXPORT_DEST As String
Public SQL_GET_SELECTION_CREATE_PPT As String
Public SQL_SELECT_TRIP_PRODUCTS As String
Public SQL_INSERT_TRIP_PRODUCTS As String
Public SQL_CLEAR_TRIP_PRODUCTS As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_GET_DELIVERY_LIST As String
Public SQL_GET_VVT_HEAD As String
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Public Function GetDeliveryList(IDViaggio As Variant, DataBuffer As Variant) As Boolean

Dim blnresult As Boolean
Dim lngHowMany As Long
Dim strDataViaggio As String

    
    blnresult = DS.DA.QueryCommand(SQL_GET_DELIVERY_LIST, DataBuffer, lngHowMany, IDViaggio)
    If Not blnresult Then
        GoTo ErrHandler
    End If
    If IsEmpty(DataBuffer) Or IsNull(DataBuffer) Then
        GoTo ErrHandler
    End If
    GetDeliveryList = True
    Exit Function

ErrHandler:

    GetDeliveryList = False

End Function
Public Function GetSelectionCreatePPD() As Boolean
    ReDim varOrderFields(0 To 1) As Variant
    With DS
        ' --- Seleziono tutti gli scarichi dei viaggi completati
        .Fields("10").FindValue = "C"
'        .Fields("10").FindValue = "D"
'        .Fields("16").FindValue = "T"
        .Fields("18").FindValue = False
        .Fields("19").FindValue = False
        varOrderFields(0) = "2"
        varOrderFields(1) = "20"
        GetSelectionCreatePPD = .GetSelection(SQL_GET_SELECTION_CREATE_PPT, True, varOrderFields)
    End With
End Function

Public Function GetSelectionExport() As Boolean
    ReDim varOrderFields(0) As Variant
    With DS
        .Fields("10").FindValue = "D"
        varOrderFields(0) = "0"
        GetSelectionExport = .GetSelection(SQL_GET_SELECTION_EXPORT, True, varOrderFields)
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
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, , MaxRecords)
    End If
End Function



Public Function PopulateTempTable() As Boolean
    On Error GoTo ErrHandler
    Dim varResults As Variant
    ReDim varidprodotti(0 To 10)
    Dim j As Integer
    Dim lngHowMany As Long
    Dim I As Long
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
    For I = 0 To lngHowMany - 1
        strIDViaggio = varResults(0, I)
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
        varidprodotti(j) = varResults(1, I)
        j = j + 1
    Next I
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
    Dim I As Integer
    
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
    SQL_GET_SELECTION = "SELECT vvtidvvt, vvtiddef, vvtidasc, vvtdtinv, " & _
                        "       vvtidaut, vvtidam1, vvtidam2, vvtidau1, " & _
                        "       vvtidau2, vvtidade, vvtidamd, vvtfltiv, " & _
                        "       vvtkmdef, vvtflstv, vvtannul FROM viagvvt0 "
    SQL_GET_SELECTION_EXPORT = "SELECT vvtidvvt, vvtidaut, vvtseqvi, vvthrinv, " & _
                               "       vvtidade, vvtidaee, vvtdtinv, vvtiddef, " & _
                               "       vvtidve1, vvtidad1, vvtidee1, " & _
                               "       vvtkmpre, vvtkmdef, vvtidam1, anagatm0.atmidatm, " & _
                               "       vvtidam2, anagatm0_1.atmidatm, vvthrinc, " & _
                               "       vvtidau1, anagaau0.aaunomea, vvtidau2, anagaau0_1.aaunomea, " & _
                               "       acnnumco, acnnumse, anagaau0.aaunumpa, anagaau0_1.aaunumpa, " & _
                               "       vvtidamd " & _
                               " FROM (((((ANAGATM0 RIGHT JOIN " & _
                               "         (ANAGAAM0 RIGHT JOIN VIAGVVT0 ON ANAGAAM0.AAMIDAAM = VIAGVVT0.VVTIDAM1) " & _
                               "          ON ANAGATM0.ATMIDATM = ANAGAAM0.AAMIDATM) LEFT JOIN ANAGAAM0 AS ANAGAAM0_1 " & _
                               "          ON VIAGVVT0.VVTIDAM2 = ANAGAAM0_1.AAMIDAAM) LEFT JOIN ANAGATM0 AS ANAGATM0_1 " & _
                               "          ON ANAGAAM0_1.AAMIDATM = ANAGATM0_1.ATMIDATM) LEFT JOIN ANAGAAU0 " & _
                               "          ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU) LEFT JOIN ANAGAAU0 AS ANAGAAU0_1 " & _
                               "          ON VIAGVVT0.VVTIDAU2 = ANAGAAU0_1.AAUIDAAU) LEFT JOIN ANAGACN0 ON VIAGVVT0.VVTIDCN1 = ANAGACN0.ACNIDACN "
    SQL_GET_SELECTION_EXPORT_DEST = "SELECT vvtidvvt, vvtidade, vvcidadz " & _
                               "  FROM viagvvt0 INNER JOIN viagvvc0 ON viagvvt0.vvtidvvt = viagvvc0.vvcidvvt "
    
    SQL_GET_SELECTION_CREATE_PPT = "SELECT VVPIDVVT, VVCDERAS, VVPIDAPR, VVPQCOLT, VVPIDOCT, " & _
                                   "       VVPIDOCR, OCTDTRIC, OCTDTCRI, OCTRIFE1, OCTRIFE2, " & _
                                   "       OCRQCOKG, OCRQCOLT, OCRQCOPZ, OCRQCOCN, VVPIDVVC, " & _
                                   "       VVPIDVVP, VVCDTCON, VVPQTRKG, VVPQTRLT, VVPQTRPZ, " & _
                                   "       VVPQTRCN, VVTIDASC, VVTIDAEE, VVTIDADE, VVCIDACL, " & _
                                   "       VVCIDADZ, SDZCODSO, VVPIDAIM, VVPIDAPD, VVPQCOKG, " & _
                                   "       VVPQCOPZ, VVPQCOCN, VVCIDAZO, VVCKMPAD, VVTIDAUT, " & _
                                   "       OCRQRILT " & _
                                   "  FROM VIAGVVT0 INNER JOIN ((VIAGVVC0 LEFT JOIN MULTSDZ0 ON " & _
                                   "      (VIAGVVC0.VVCIDADZ = MULTSDZ0.SDZIDADZ) AND (VIAGVVC0.VVCIDASC = MULTSDZ0.SDZIDASC)) INNER JOIN " & _
                                   "      (ORDIOCT0 RIGHT JOIN (VIAGVVP0 LEFT JOIN ORDIOCR0 ON " & _
                                   "      (VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT) AND (VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR)) ON " & _
                                   "       ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT) ON (VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) AND (VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT)) " & _
                                   "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT "
'    SQL_GET_SELECTION_CREATE_PPT = "SELECT VVPIDVVT, VVCDERAS, VVPIDAPR, VVPQCOLT, VVPIDOCT, " & _
'                                   "       VVPIDOCR, OCTDTRIC, OCTDTCRI, OCTRIFE1, OCTRIFE2, " & _
'                                   "       OCRQCOKG, OCRQCOLT, OCRQCOPZ, OCRQCOCN, VVPIDVVC, " & _
'                                   "       VVPIDVVP, VVTDTINV, VVPQTRKG, VVPQTRLT, VVPQTRPZ, " & _
'                                   "       VVPQTRCN, VVTIDASC, VVTIDAEE, VVTIDADE, VVCIDACL, " & _
'                                   "       VVCIDADZ, SDZCODSO, VVPIDAIM, VVPIDAPD, VVPQCOKG, " & _
'                                   "       VVPQCOPZ, VVPQCOCN, VVCIDAZO, VVCKMPAD, VVTIDAUT, " & _
'                                   "       OCRQRILT, VVCDTCON " & _
'                                   "  FROM VIAGVVT0 INNER JOIN ((VIAGVVC0 LEFT JOIN MULTSDZ0 ON " & _
'                                   "      (VIAGVVC0.VVCIDADZ = MULTSDZ0.SDZIDADZ) AND (VIAGVVC0.VVCIDASC = MULTSDZ0.SDZIDASC)) INNER JOIN " & _
'                                   "      (ORDIOCT0 RIGHT JOIN (VIAGVVP0 LEFT JOIN ORDIOCR0 ON " & _
'                                   "      (VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT) AND (VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR)) ON " & _
'                                   "       ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT) ON (VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) AND (VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT)) " & _
'                                   "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT "

    SQL_SELECT_TRIP_PRODUCTS = "SELECT DISTINCT VVPIDVVT, VVPIDAPR " & _
                                  "  FROM VIAGVVT0 INNER JOIN (VIAGVVC0 INNER JOIN VIAGVVP0 " & _
                                  "       ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) " & _
                                  "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT "
    SQL_INSERT_TRIP_PRODUCTS = "INSERT INTO VIAGTMP0 VALUES (?,?,?,?,?,?,?,?,?,?,?)"
    SQL_CLEAR_TRIP_PRODUCTS = "DELETE FROM VIAGTMP0 "
    SQL_UPDATE = "UPDATE viagvvt0 SET "
    SQL_INSERT = "INSERT INTO viagvvt0 "
    SQL_GET_TS = "SELECT timestmp FROM viagvvt0 WHERE "
    SQL_GET_DELIVERY_LIST = " SELECT VIAGVVC0.VVCIDVVT, VIAGVVC0.VVCNRSEQ,VIAGVVC0.VVCIDADZ, VIAGVVC0.VVCDERAS, " & _
                            " VIAGVVC0.VVCDELOC, VIAGVVP0.VVPIDAPR,ANAGAPR0.APRDESCO, VIAGVVP0.VVPQTRLT, VIAGVVP0.VVPQCOLT " & _
                            " FROM VIAGVVP0 INNER JOIN VIAGVVC0 ON VIAGVVP0.VVPIDVVT = VIAGVVC0.VVCIDVVT AND " & _
                            " VIAGVVP0.VVPIDVVC = VIAGVVC0.VVCIDVVC INNER JOIN ANAGAPR0 ON VIAGVVP0.VVPIDAPR = ANAGAPR0.APRIDAPR " & _
                            " WHERE VIAGVVC0.VVCIDVVT = ? ORDER BY VIAGVVC0.VVCNRSEQ, VIAGVVP0.VVPIDAPR "
                            
    SQL_GET_VVT_HEAD = " SELECT VIAGVVT0.VVTIDADE, VIAGVVT0.VVTIDVVT, VIAGVVT0.VVTIDDEF, VIAGVVT0.VVTIDAUT, VIAGVVT0.VVTIDAM1, VIAGVVT0.VVTIDAM2, " & _
                       " VIAGVVT0.VVTDTINV, SUM(VIAGVVP0.VVPQTRLT) AS QTRLT, SUM(VIAGVVP0.VVPQCOLT) AS QCOLT, VIAGVVT0.VVTDTINC, VIAGVVT0.VVTHRINC, VIAGVVT0.VVTFLSTV, " & _
                       " VIAGVVT0.VVTSEQVI, ANAGAAU0.AAUNOMEA,VIAGVVT0.VVTIDAU1 FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND " & _
                       " VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU "
    
    
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
Public Function GetDeliveryHeader(DataBuffer As Variant, DataViaggio As Variant, _
                                  IDDeposito As Variant, NumViaggioDef As Variant, _
                                  IDUnitadiTras As Variant, TMotrice As Variant, _
                                  TRimorchio) As Boolean

Dim blnresult As Boolean
Dim lngHowMany As Long
Dim SQL_ORDER As String
Dim SQL_PARAM As String
Dim SQL_COND As String
Dim SQL_UNION As String
Dim LngCounter As Long
Dim strDataViaggio As String


    SQL_ORDER = " GROUP BY VIAGVVT0.VVTIDADE, VIAGVVT0.VVTIDVVT, VIAGVVT0.VVTIDDEF, VIAGVVT0.VVTIDAUT, " & _
                " VIAGVVT0.VVTIDAM1, VIAGVVT0.VVTIDAM2, VIAGVVT0.VVTDTINV, VIAGVVT0.VVTDTINC, VIAGVVT0.VVTHRINC, VIAGVVT0.VVTFLSTV, VIAGVVT0.VVTSEQVI,ANAGAAU0.AAUNOMEA, VIAGVVT0.VVTIDAU1 " & _
                " ORDER BY VIAGVVT0.VVTDTINV DESC, VVTIDADE, VIAGVVT0.VVTIDAUT , VIAGVVT0.VVTSEQVI "

    SQL_PARAM = ""
    'SQL_COND = " WHERE VIAGVVT0.VVTFLSTV <> 'P' "
    
    If Len(CStr(DataViaggio)) > 0 Then
        strDataViaggio = Format(DataViaggio, "mm/dd/yy")
        SQL_PARAM = SQL_PARAM & " AND (VIAGVVT0.VVTDTINV = '" & strDataViaggio & "')"
    End If
    If Len(CStr(IDDeposito)) > 0 Then
        SQL_PARAM = SQL_PARAM & " AND (VIAGVVT0.VVTIDADE = '" & IDDeposito & "')"
    End If
    If Len(CStr(NumViaggioDef)) > 0 Then
        SQL_PARAM = SQL_PARAM & " AND (VIAGVVT0.VVTIDDEF = '" & NumViaggioDef & "')"
    End If
    If Len(CStr(IDUnitadiTras)) > 0 Then
        SQL_PARAM = SQL_PARAM & " AND (VIAGVVT0.VVTIDAUT = '" & IDUnitadiTras & "')"
    End If
    If Len(CStr(TMotrice)) > 0 Then
        SQL_PARAM = SQL_PARAM & " AND (VIAGVVT0.VVTIDAM1 LIKE '" & TMotrice & "%')"
    End If
    If Len(CStr(TRimorchio)) > 0 Then
        SQL_PARAM = SQL_PARAM & " AND (VIAGVVT0.VVTIDAM2 LIKE '" & TRimorchio & "%')"
    End If
    If Len(SQL_PARAM) > 0 Then
        SQL_PARAM = " WHERE " & Right(SQL_PARAM, Len(SQL_PARAM) - 4)
    End If
    SQL_UNION = SQL_GET_VVT_HEAD & " " & SQL_PARAM & " " & SQL_ORDER
    
    blnresult = DS.DA.QueryData(SQL_UNION, DataBuffer)
    
    
    If Not blnresult Then
        GoTo ErrHandler
    End If
    If IsEmpty(DataBuffer) Or IsNull(DataBuffer) Then
        GoTo ErrHandler
    Else
        For LngCounter = 0 To (UBound(DataBuffer, 2))
            DataBuffer(10, LngCounter) = Format(DataBuffer(10, LngCounter), "HH:MM")
        Next LngCounter
    End If
    GetDeliveryHeader = True
    Exit Function

ErrHandler:

    GetDeliveryHeader = False

End Function




