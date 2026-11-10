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
Option Explicit


Public SQL_GETLIST As String
Public SQL_GET_SELECTION As String
Public SQL_GET_SELECTION_GC As String
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
        GetSelectionCreatePPD = .GetSelection(SQL_GET_SELECTION_CREATE_PPT, True, varOrderFields)
    End With
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
        .Fields("92").FindValue = "GC"
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
                                   "       VVPIDVVP, VVCDTCON, VVPQTRKG, VVPQTRLT, VVPQTRPZ, " & _
                                   "       VVPQTRCN, VVTIDASC, VVTIDAEE, VVTIDADE, VVCIDACL, " & _
                                   "       VVCIDADZ, SDZCODSO, VVPIDAIM, VVPIDAPD, VVPQCOKG, " & _
                                   "       VVPQCOPZ, VVPQCOCN, VVCIDAZO, VVCKMPAD, VVTIDAUT, " & _
                                   "       OCRQRILT, OCTRAUT3 " & _
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




