VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggio"
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
Public SQL_DELETE As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_GET_TOTALE_QTA As String
Public SQL_GET_TOTALE_QTA_PRODOTTO As String
Public SQL_GET_TOTALE_QTA_PRODOTTO_GROUPBY As String
Public SQL_GET_QTA_SCARICHI As String
Public SQL_SET_STATO As String
Public SQL_SET_DATE_CONSEGNE_EFF As String
Public SQL_SET_DATE_CONSEGNE As String
Public SQL_COPY_CONSEGNE As String
Public SQL_COPY_SCARICHI As String
Public SQL_FIELDS_CONSEGNE As String
Public SQL_FIELDS_SCARICHI As String
Public SQL_COUNT_SIF As String
Public SQL_SELECT_TRIP_PRODUCTS As String
Public SQL_INSERT_TRIP_PRODUCTS As String
Public SQL_CLEAR_TRIP_PRODUCTS As String
Public SQL_GETLIST_COMPUTE_CU_REPARTO As String
Private SQL_COUNT_DESTINATARI As String
Private SQL_COUNT_DESTINAZIONI As String
Private SQL_COUNT_CONSEGNE As String
Private SQL_GETLIST_SET_DATI_STIPENDI As String
Private SQL_GET_ULTIMO_VIAGGIO_CARICATO As String
Private SQL_GET_ULTIMO_VIAGGIO_CARICATO_ORDERBY As String

'-- JG 31/03/09 No chequear deposito
Private Const SQL_CHECK_SEQUENZA_INSERT_DEP = "select count(*) from viagvvt0 " & _
                                          " where vvtdtinv = ? and convert(int, vvtseqvi) = ? and vvtidaut = ? AND VVTANNUL = 0 "

Private Const SQL_CHECK_SEQUENZA_INSERT = "select count(*) from viagvvt0 " & _
                                          " where vvtdtinv = ? and convert(int, vvtseqvi) = ? and vvtidaut = ? and vvtidade = ? AND VVTANNUL = 0 "
Private Const SQL_CHECK_SEQUENZA = "select count(*) from viagvvt0 " & _
                                          " where vvtdtinv = ? and convert(int, vvtseqvi) = ? and vvtidaut = ? and vvtidade = ? AND vvtidvvt <> ? AND VVTANNUL = 0 "
Private Const SQL_COUNT_CONSEGNE_DA_COMPLETARE = "SELECT COUNT(*) " & _
                                                 "  FROM VIAGVVC0 " & _
                                                 " WHERE VVCIDVVT = ? AND VVCANNUL = 0 AND ((VVCFLST5 IS NULL) OR (VVCFLST5 = 'N'))"
Private Const SQL_COUNT_CONSEGNE_SENZA_ORARIO = "SELECT COUNT(*) " & _
                                                 "  FROM VIAGVVC0 " & _
                                                 " WHERE VVCIDVVT = ? AND VVCANNUL = 0 AND VVCFLCPP = 'M' AND ((VVCTEMVG IS NULL) OR (VVCTEMVG = '00:00'))"
Private Const SQL_COUNT_ORDINI_VENTA_EN_RUTA = "SELECT DISTINCT VVPIDAPR, VVPIDVVT, VVPIDVVC, VVPIDVVP " & _
                                                 "  FROM VIAGVVC0 INNER JOIN ORDIOCT0 ON VVCIDOCT = OCTIDOCT " & _
                                                 "       INNER JOIN VIAGVVP0 ON VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC " & _
                                                 " WHERE VVCIDVVT = ? AND (OCTFLVER = '1') AND VVCANNUL = 0 AND VVPANNUL = 0 " & _
                                                 " ORDER BY VVPIDAPR "

Private Const SQL_SET_RIF_VENTA_EN_RUTA = "UPDATE VIAGVVP0 " & _
                                          "   SET VVPNRVER = ? " & _
                                          " WHERE VVPIDVVT = ? AND VVPIDVVC = ? AND VVPIDVVP = ? "

Private Const SQL_SEQUENZA_OK = "SELECT 1 FROM VIAGVVT0 WHERE VVTDTINC = ? AND VVTIDUDT = ? AND VVTSEQVI = ? AND VVTANNUL = 0"

'-- DP 14/08/08
Private Const SQL_GET_COUNTER_SEQUENZA_INVIO = "SELECT SEQNRSEQ FROM VIAGSEQ0 WHERE SEQIDUDT = ? AND SEQDTINV = ?"

Private Const SQL_UPDATE_COUNTER_SEQUENZA_INVIO = "UPDATE VIAGSEQ0 SET SEQNRSEQ = ? WHERE SEQIDUDT = ? AND SEQDTINC = ?"

Private Const SQL_SET_COUNTER_SEQUENZA_INVIO = "INSERT INTO VIAGSEQ0 (SEQNRSEQ, SEQIDUDT, SEQDTINC) VALUES (?, ?, ?)"

' [28-04-05 (PM)] Nuovi statement per interfacciamento KNOSOS
Private SQL_KNOSOS_INSERT_KVIAGGIO As String
Private SQL_KNOSOS_SELECT_KVIAGGIO As String

Private SQL_KNOSOS_INSERT_KCARICO As String
Private SQL_KNOSOS_SELECT_KCARICO As String
Private SQL_KNOSOS_UPDATE_KCARICO As String

Private SQL_KNOSOS_INSERT_KSCARICO As String
Private SQL_KNOSOS_SELECT_KSCARICO As String
Private SQL_KNOSOS_UPDATE_KSCARICO As String
Private SQL_KNOSOS_UPDATESTATO_KSCARICO As String

Private SQL_KNOSOS_SELECT_SCARICHI_RESOIL As String
Private SQL_KNOSOS_SELECT_KSCARICHI_IDDESC As String
Private SQL_KNOSOS_RESET_IDDESC_ANNULLATE As String

Private SQL_KNOSOS_UPDATE_NUMERO_SCARICO As String
Private SQL_KNOSOS_GET_NEXTID_SCARICO As String

Private SQL_UPDATE_KNOSOS_DATI As String

Private Const KNOSOS_TIPO_CARICO As Integer = 0
Private Const KNOSOS_TIPO_SCARICO As Integer = 1
Private Const KNOSOS_IDSCARICO_CARICO As Integer = 0

' Stato del viaggio KNOSOS
Private Const KNOSOS_VIAGGIO_STATO_PENDENTE = 0
Private Const KNOSOS_VIAGGIO_STATO_INVIATO = 1
Private Const KNOSOS_VIAGGIO_STATO_RICEVUTO = 2
Private Const KNOSOS_VIAGGIO_STATO_INCORSO = 3
Private Const KNOSOS_VIAGGIO_STATO_TERMINATO = 9
Private Const KNOSOS_VIAGGIO_STATO_ANNUL_PENDENTE = 10
Private Const KNOSOS_VIAGGIO_STATO_ANNUL_INVIATO = 11
Private Const KNOSOS_VIAGGIO_STATO_ANNUL_ACCETTATO = 12

' Stato del carico o scarico KNOSOS
Private Const KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_CREAZIONE = 0
Private Const KNOSOS_OPEVIAGGIO_STATO_CREATA = 1
Private Const KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_MODIFICA = 2
Private Const KNOSOS_OPEVIAGGIO_STATO_MODIFICATA = 3
Private Const KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_ANNULLAMENTO = 4
Private Const KNOSOS_OPEVIAGGIO_STATO_ANNULLATA = 5

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Private mvarTotaliQta As Variant
Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

'Private Const SQL_GET_EXVIAGDETAIL = " SELECT VVTIDVVT, VVTIDDEF, VVTIDASC, AAUNOMAB, " & _
'                                     "        VVTIDAM1, VVTIDAM2, vvtkmdef, SDECODSO, " & _
'                                     "        SEECODSO, SUTCODSO, VVTDTINV, ACNNUMCO, " & _
'                                     "        VVTIDADE, ADELOCAL, ACNNUMSE, ANAGAAM0.AAMDESCR, " & _
'                                     "        ANAGAAM01.AAMDESCR AS DescRimorchio " & _
'                                     "   FROM VIAGVVT0 INNER JOIN ANAGAAM0 ON VVTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
'                                     "        ANAGAAM0 ANAGAAM01 ON VVTIDAM2 = ANAGAAM01.AAMIDAAM LEFT OUTER JOIN ANAGADE0 ON " & _
'                                     " VIAGVVT0.VVTIDADE = ANAGADE0.ADEIDADE LEFT OUTER JOIN MULTSDE0 ON " & _
'                                     " VIAGVVT0.VVTIDADE = MULTSDE0.SDEIDADE LEFT OUTER JOIN ANAGAAU0 ON " & _
'                                     " VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU LEFT OUTER JOIN MULTSUT0 ON " & _
'                                     " VIAGVVT0.VVTIDAUT = MULTSUT0.SUTIDAUT LEFT OUTER JOIN MULTSEE0 ON " & _
'                                     " VIAGVVT0.VVTIDAEE = MULTSEE0.SEEIDAEE LEFT OUTER JOIN " & _
'                                     " ANAGACN0 ON " & _
'                                     " VIAGVVT0.VVTIDCN1 = ANAGACN0.ACNIDACN " & _
'                                     " WHERE (VVTIDVVT = ?) AND (MULTSDE0.SDEIDASC = ?) AND (MULTSEE0.SEEIDASC = ?) AND (MULTSUT0.SUTIDASC = ?)"
'
Private Const SQL_GET_EXVIAGDETAIL = _
"  SELECT VVTIDVVT, VVTIDDEF, VVTIDASC, VVTIDVE1, " & _
"         VVTIDAM1, VVTIDAM2, VVTKMDEF, VVTIDADE, " & _
"         VVTIDAEE, VVTIDAUT, VVTDTINV, ACNNUMCO, " & _
"         VVTIDADE, ADELOCAL, ACNNUMSE, ANAGAAM0.AAMDESCR, " & _
"         ANAGAAM0_1.AAMDESCR AS DescRimorchio, VVCIDADZ, OCTRIFE1, OCTNOTA3, " & _
"         OCTIDAZO, VVPRIFE1, VVPIDAPR, VVPQCOLT, " & _
"         VVCIDACL, ADZRAGU5, VVCDELOC, VVCIDVVC, " & _
"         AAUNOMAB, VVCANNUL, OCTIDARO, VVTFLST2, " & _
"         VVCKMRAG, OCTIDADE " & _
"    FROM ANAGADZ0 INNER JOIN ((ANAGAAM0 INNER JOIN ((((VIAGVVT0 LEFT JOIN ANAGAAM0 AS ANAGAAM0_1 ON VIAGVVT0.VVTIDAM2 = ANAGAAM0_1.AAMIDAAM) INNER JOIN ANAGAAU0 ON VIAGVVT0.VVTIDAU1 = ANAGAAU0.AAUIDAAU) LEFT JOIN ANAGADE0 ON VIAGVVT0.VVTIDADE = ANAGADE0.ADEIDADE) LEFT JOIN ANAGACN0 ON VIAGVVT0.VVTIDCN1 = ANAGACN0.ACNIDACN) ON ANAGAAM0.AAMIDAAM = VIAGVVT0.VVTIDAM1) INNER JOIN ((ORDIOCT0 INNER JOIN VIAGVVC0 ON ORDIOCT0.OCTIDOCT = VIAGVVC0.VVCIDOCT) INNER JOIN VIAGVVP0 ON (VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) AND (VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT)) ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT) ON ANAGADZ0.ADZIDADZ = VIAGVVC0.VVCIDADZ " & _
"   WHERE (VVTIDVVT = ?) " & _
"ORDER BY VVCNRSEQ, VVPIDAPR "

Private Const SQL_UPDATE_SEQOK = "UPDATE VIAGVVT0 SET VVTSEQOK = ? WHERE VVTIDVVT = ?"

Private Const SQL_UPDATE_SEQIN = "UPDATE VIAGVVT0 SET VVTSEQIN = ? WHERE VVTIDVVT = ?"

Public Function CheckScadenze(strMessage As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strError As String
    Dim strIDUdT As String
    Dim strIDASC As String
    Dim strMotrice As String
    Dim strRimorchio As String
    Dim strIDAutista As String
    Dim varData As Date
    Dim varResult As Variant
    Dim AUT As CUnitaDiTrasporto
    Dim AAM As CAutomezzo
    Dim AAU As CAutista
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    With DS
        If .Field("Data Inizio Viaggio").Updated Then
            varData = .Field("Data Inizio Carico").NewValue
        Else
            varData = .Field("Data Inizio Carico").Value
        End If
        If .Field("Societ‡").Updated Then
            strIDASC = .Field("Societ‡").NewValue
        Else
            strIDASC = .Field("Societ‡").Value
        End If
        
        If .Field("Unit‡ di Trasporto").Updated Then
            strIDUdT = .Field("Unit‡ di Trasporto").NewValue
            If Not AUT.CheckScadenze(strIDUdT, varData, strIDASC, varResult) Then
                GoTo ErrHandler
            End If
            If IsEmpty(varResult) Then
                GoTo ErrHandler
            End If
            Select Case varResult(0)
                Case 0
                Case 1
                    strMessage = varResult(1)
                Case 2
                    strMessage = varResult(1)
                    GoTo ErrHandler
                Case 3
                    strMessage = varResult(1)
                Case Else
                    GoTo ErrHandler
            End Select
        ElseIf .Field("Motrice").Updated Then
            strMotrice = .Field("Motrice").NewValue
            If Not AAM.CheckScadenze(strMotrice, varData, strIDASC, varResult) Then
                GoTo ErrHandler
            End If
            If IsEmpty(varResult) Then
                GoTo ErrHandler
            End If
            Select Case varResult(0)
                Case 0
                Case 1
                    strMessage = varResult(1)
                Case 2
                    strMessage = varResult(1)
                    GoTo ErrHandler
                Case 3
                    strMessage = varResult(1)
                Case Else
                    GoTo ErrHandler
            End Select
        ElseIf .Field("Rimorchio").Updated Then
            strRimorchio = .Field("Rimorchio").NewValue
            If Not AAM.CheckScadenze(strRimorchio, varData, strIDASC, varResult) Then
                GoTo ErrHandler
            End If
            If IsEmpty(varResult) Then
                GoTo ErrHandler
            End If
            Select Case varResult(0)
                Case 0
                Case 1
                    strMessage = varResult(1)
                Case 2
                    strMessage = varResult(1)
                    GoTo ErrHandler
                Case 3
                    strMessage = varResult(1)
                Case Else
                    GoTo ErrHandler
            End Select
        ElseIf .Field("Autista 1").Updated Then
            strIDAutista = .Field("Autista 1").NewValue
            If Not AAU.CheckScadenze(strIDAutista, varData, varResult) Then
                GoTo ErrHandler
            End If
            If IsEmpty(varResult) Then
                GoTo ErrHandler
            End If
            Select Case varResult(0)
                Case 0
                Case 1
                    strMessage = varResult(1)
                Case 2
                    strMessage = varResult(1)
                    GoTo ErrHandler
                Case 3
                    strMessage = varResult(1)
                Case Else
                    GoTo ErrHandler
            End Select
        End If
    End With
    
    CheckScadenze = True
    GoTo EndProc
    
ErrHandler:
    If Len(strMessage) = 0 Then
        strError = "Si Ë verificato un errore durante il controllo delle scadenze"
    Else
        strError = strMessage
    End If
    CheckScadenze = False
EndProc:
    strMessage = strError
End Function

    

Public Function CheckSequenza(Data As Variant, NumSequenza As Variant, Unit‡DiTrasporto As Variant, Deposito As Variant, Optional IDViaggio As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim varResult As Variant
    Dim lngHowMany As Long
    
    If Not IsNumeric(NumSequenza) Then
        '-- DP 28/10/08
'        CheckSequenza = True
'        Exit Function
        GoTo ErrHandler
    End If
    
    If NumSequenza <= 0 Then
        '-- DP 28/10/08
'        CheckSequenza = True
'        Exit Function
        GoTo ErrHandler
    End If
    

    If Len(IDViaggio) > 0 Then
        If Not DS.DA.QueryCommand(SQL_CHECK_SEQUENZA, varResult, lngHowMany, Data, NumSequenza, Unit‡DiTrasporto, Deposito, IDViaggio) Then
            GoTo ErrHandler
        End If
    Else
    ' JG 31/3/9 No controlar el depÛsito
        If Len(Deposito) > 0 Then
            If Not DS.DA.QueryCommand(SQL_CHECK_SEQUENZA_INSERT, varResult, lngHowMany, Data, NumSequenza, Unit‡DiTrasporto, Deposito) Then
                GoTo ErrHandler
            End If
        Else
           If Not DS.DA.QueryCommand(SQL_CHECK_SEQUENZA_INSERT_DEP, varResult, lngHowMany, Data, NumSequenza, Unit‡DiTrasporto) Then
                GoTo ErrHandler
            End If
        End If
    End If
    
    CheckSequenza = (varResult(0, 0) = 0)
    Exit Function
    
ErrHandler:
    CheckSequenza = False
End Function

 
Public Function ComputeCU(IDViaggio As String, PortataVeicolo As Variant, _
                          ByRef CUCarga As Variant, ByRef CUReparto As Variant, CUTotale As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim varCUReparto As Variant
    Dim varBuffer As Variant
    Dim varQuantit‡PerKM As Variant
    Dim varQuantit‡PerKMTot As Variant
    Dim strIDConsegna As Variant
    Dim strIDConsegnaPrec As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim varQtaKgConsegna As Variant
    Dim varQtaKGViaggio As Variant
    
    If Not DS.DA.QueryCommand(SQL_GETLIST_COMPUTE_CU_REPARTO, varBuffer, lngHowMany, IDViaggio) Then
        GoTo ErrHandler
    End If
    For i = 0 To UBound(varBuffer, 2)
        If strIDConsegna <> strIDConsegnaPrec Then
            ' Termine elaborazione consegna precedente
            If i > 0 Then
                varQuantit‡PerKM = varQuantit‡PerKM + varQtaKgConsegna * varBuffer(1, i)
            End If
            ' Inizio elaborazione consegna successiva
            strIDConsegnaPrec = strIDConsegna
            varQtaKgConsegna = 0
        End If
        If IsNumeric(varBuffer(3, i)) Then
            varQtaKgConsegna = varQtaKgConsegna + Round(varBuffer(2, i) * varBuffer(3, i))
            varQtaKGViaggio = varQtaKGViaggio + varQtaKgConsegna
        End If
    Next i
    varQuantit‡PerKM = varQuantit‡PerKM + varQtaKgConsegna * varBuffer(1, i - 1)
    varQuantit‡PerKMTot = varQtaKGViaggio * varQtaKGViaggio
    
    CUReparto = 0
    CUCarga = 0
    CUTotale = 0
    If IsNumeric(PortataVeicolo) And (PortataVeicolo > 0) Then
        CUCarga = Round((varQtaKGViaggio / PortataVeicolo) * 100, 2)
    End If
    If IsNumeric(varQuantit‡PerKMTot) And (varQuantit‡PerKMTot > 0) Then
        CUReparto = Round((varQuantit‡PerKM / varQuantit‡PerKMTot) * 10, 2)
    End If
    CUTotale = Round(CUCarga * CUReparto, 2)
    ComputeCU = True
    Exit Function
ErrHandler:
    ComputeCU = False
End Function

Public Function FindBySMSCarico(IDAutista As Variant, Data As Variant, ProgrCarico As Variant) As Variant
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    On Error GoTo ErrHandler
    
    '-- DP 04/06/08
    '-- Ho modificato "VVTDTINV = ?" in "(VVTDTINV >= ? OR VVTDTINV = ?)" perchÈ altrimenti scappano i viaggi caricati "il giorno precedente o i seguenti la data pianificata"
    '-- L'"ORDER BY" finale serve, nel remoto caso venissero trovati pi˘ viaggi con lo stesso autista, lo stesso progressivo mezzo e la data >= a quella passata, per posizionarmi sul pi˘ vecchio che dovrebbe essere quello corretto
    strSQL = "SELECT VVTIDVVT " & _
            " FROM VIAGVVT0 INNER JOIN VIAGSVG0 ON " & _
            "       VVTIDVVT = SVGIDVVT " & _
            " WHERE VVTIDAU1 = ? AND (VVTDTINV >= ? OR VVTDTINV = ?) AND SVGPRGME = ? AND " & _
            "       (VVTFLST3 = 'N' OR VVTFLST3 IS NULL OR VVTFLST3 = '') AND VVTANNUL = 0 " & _
            " ORDER BY VVTDTRIF"
    
    If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, IDAutista, Data, DateAdd("d", -1, CDate(Data)), ProgrCarico) Then
        GoTo ErrHandler
    End If
    
    If IsEmpty(varBuffer) Then
        FindBySMSCarico = "0"
    Else
        FindBySMSCarico = varBuffer(0, 0)
    End If
    
    Exit Function
    
ErrHandler:
    FindBySMSCarico = "-1"
End Function

Public Function GetExportEsysData(IDAsc As String) As Variant
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strIDViaggio As String
    
    strIDViaggio = DS.Fields("0").Value
    blnResult = DS.DA.QueryCommand(SQL_GET_EXVIAGDETAIL, varBuffer, lngHowMany, _
                                   strIDViaggio)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetExportEsysData = varBuffer
    Exit Function
ErrHandler:
    GetExportEsysData = Empty
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

Public Function ContaProdottiSIF(CodSIF As String) As Integer
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_SIF, varBuffer, _
                                   0, CodSIF, DS.Fields("3").Value, _
                                   DS.Fields("0").Value)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    ContaProdottiSIF = varBuffer(0, 0)
    
    Exit Function
ErrHandler:
    ContaProdottiSIF = -1
End Function

Public Function ContaDestinatari() As Integer
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_DESTINATARI, varBuffer, lngHowMany, _
                                   DS.Fields("0").Value)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    ContaDestinatari = varBuffer(0, 0) 'UBound(varBuffer, 2)
    
    Exit Function
ErrHandler:
    ContaDestinatari = -1
End Function
Public Function ContaDestinazioni() As Integer
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_DESTINAZIONI, varBuffer, lngHowMany, _
                                   DS.Fields("0").Value)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    ContaDestinazioni = UBound(varBuffer, 2) + 1
    
    Exit Function
ErrHandler:
    ContaDestinazioni = -1
End Function

Public Function ContaConsegne() As Long
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_CONSEGNE, varBuffer, lngHowMany, _
                                   DS.Fields("0").Value)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    ContaConsegne = varBuffer(0, 0)
    
    Exit Function
ErrHandler:
    ContaConsegne = -1
End Function

Public Function CheckViaggioCompletato(NumeroViaggio As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_CONSEGNE_DA_COMPLETARE, varBuffer, lngHowMany, _
                                   NumeroViaggio)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    CheckViaggioCompletato = (varBuffer(0, 0) = 0)
    
    Exit Function
ErrHandler:
    CheckViaggioCompletato = False
End Function

Public Function CheckOrarioConsegne(NumeroViaggio As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_CONSEGNE_SENZA_ORARIO, varBuffer, lngHowMany, _
                                   NumeroViaggio)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    CheckOrarioConsegne = (varBuffer(0, 0) = 0)
    
    Exit Function
ErrHandler:
    CheckOrarioConsegne = False
End Function
Public Function IsUpdatable() As Boolean
    Dim strAvanzamento As String
    Dim strNumeroViaggioDef As String
    Dim blnTrasmAGIP As Boolean
    
    ' --- Impedisco il salvataggio del viaggio se
    ' ---   Viaggio elaborato
    strNumeroViaggioDef = Trim(DS.Fields("1").Value)
    blnTrasmAGIP = (DS.Fields("41").Value = "T")
    strAvanzamento = (DS.Fields("42").Value)
    
    IsUpdatable = True
    
    If strAvanzamento = "E" Then
        IsUpdatable = False
        Exit Function
    End If
    
'    If strNumeroViaggioDef <> vbNullString Then
'        If (CLng(strNumeroViaggioDef) > 0) And (Not blnTrasmAGIP) Then
'            IsUpdatable = False
'            Exit Function
'        End If
'    End If
    ' ---
End Function
Public Function Copia() As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim strIDTrip As String
    Dim strIDNewTrip As String
    Dim i  As Integer
    
    Copia = True
    With DS.Fields
        strIDTrip = .Item("0").Value
        If .Item("42").Value <> "E" Then
            Exit Function
        End If
        For i = 1 To .Count
            If Not .Item(i).ReadOnly Then
                .Item(i).NewValue = .Item(i).Value
                .Item(i).Updated = True
            Else
                .Item(i).Updated = False
            End If
        Next i
        .Item("1").NewValue = vbNullString
        .Item("1").Updated = True
        .Item("24").NewValue = strIDTrip
        .Item("24").Updated = True
        .Item("41").NewValue = vbNullString
        .Item("41").Updated = True
        .Item("44").NewValue = False
        .Item("44").Updated = True
        DS.DA.BeginTransaction
        If Not Insert Then
            GoTo ErrHandler
        End If
        
        strIDNewTrip = .Item("0").Value
        ' --- Copio le consegne
        strQ = SQL_COPY_CONSEGNE & " SELECT " & DS.DA.MakeSQL(strIDNewTrip, .Item("0").FldType) & ", "
        strQ = strQ & SQL_FIELDS_CONSEGNE
        strQ = strQ & "  FROM viagvvc0 "
        strQ = strQ & " WHERE vvcidvvt = " & DS.DA.MakeSQL(strIDTrip, .Item("0").FldType)
    
        If Not DS.DA.ExecuteQuery(strQ) Then
            GoTo ErrHandler
        End If
        ' ---
        
        ' --- Copio gli scarichi
        strQ = SQL_COPY_SCARICHI & " SELECT " & DS.DA.MakeSQL(strIDNewTrip, .Item("0").FldType) & ", "
        strQ = strQ & SQL_FIELDS_SCARICHI
        strQ = strQ & "  FROM viagvvp0 "
        strQ = strQ & " WHERE vvpidvvt = " & DS.DA.MakeSQL(strIDTrip, .Item("0").FldType)
    
        If Not DS.DA.ExecuteQuery(strQ) Then
            GoTo ErrHandler
        End If
        ' ---
    
    End With
    DS.DA.EndTrans 0
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    Copia = False
End Function

Public Function SetFlagTrasmissione(NewValue As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDViaggio As String
    Dim blnResult As Boolean
    Dim blnConcurErr As Boolean
    
    SetFlagTrasmissione = True
    With DS.Fields
        strIDViaggio = .Item("0").FindValue
        .Item("41").NewValue = NewValue
        .Item("41").Updated = True
    End With
    blnResult = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    If blnResult Then
        DS.DA.EndTrans 0
    Else
        DS.DA.EndTrans 1
    End If
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    SetFlagTrasmissione = False
End Function

Public Function SetFlagTrasmissioneESys(NewValue As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDViaggio As String
    Dim blnResult As Boolean
    Dim blnConcurErr As Boolean
    
    With DS.Fields
        strIDViaggio = .Item("0").FindValue
        .Item("40").NewValue = NewValue
        .Item("40").Updated = True
    End With
    blnResult = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    If blnResult Then
        DS.DA.EndTrans 0
    Else
        GoTo ErrHandler
    End If
    SetFlagTrasmissioneESys = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    SetFlagTrasmissioneESys = False
End Function

Public Function SetStato(NewValue As String, Optional ErrDescr As String) As Boolean
    Dim strIDViaggio As String
    Dim strCounter As String
    Dim strError As String
    Dim strQ As String
    Dim blnConcurErr As Boolean
    Dim blnUpdateCounter As Boolean
    Dim blnResult As Boolean
    Dim blnConcurError As Boolean
    Dim blnCalcolaOrario As Boolean
    Dim varBuffer As Variant
    Dim varQta As Variant
    Dim lngHowMany As Long
    Dim i As Integer
    Dim j As Integer
    
    Dim varBufDestinatari As Variant
    Dim ADZ As CDestinazione
    Dim varKmTemp As Variant
    Dim intKmMax As Integer
    Dim strIDADE As String
   
    On Error GoTo ErrHandler
    
    SetStato = True
    With DS.Fields
        strIDViaggio = .Item("0").FindValue
        blnCalcolaOrario = True
        Select Case NewValue
            Case "P"
                If DS.Fields("42").Value <> "D" Then
                    Exit Function
                End If
                ' Forziamo la trasmissione ad Eurosystem
                SetFlagTrasmissioneESys "N"
'                .ResetFields
                DS.Fields("1").NewValue = vbNullString
                DS.Fields("1").Updated = True
            Case "D"
                If DS.Fields("42").Value <> "P" Then
                    Exit Function
                End If
                
                If Not Me.CheckOrarioConsegne(DS.Field("Numero Viaggio").Value) Then
                    strError = "Esistono consegne senza orario programmato"
                    GoTo ErrHandler
                End If
                If Not Me.CheckScadenze(strError) Then
                    GoTo ErrHandler
                End If
                strCounter = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "003", 0)
                If strCounter = CStr(-1) Then
                    GoTo ErrHandler
                End If
                
                If Not SMS_Insert(DS.Field("Numero Viaggio").Value, DS.Field("Autista 1").Value) Then
                    GoTo ErrHandler
                End If
                
                If Not Me.CheckVentaEnRuta(strIDViaggio, DS.Field("Ente Emittente").Value) Then
                    GoTo ErrHandler
                End If
                
                ' 27/02/07 - La riga seguente Ë stata commentata per evitare che vengano ripuliti
                '            i campi che vengono utilizzati
'                .ResetFields

                DS.Fields("1").NewValue = strCounter
                DS.Fields("1").Updated = True
                
                
                '-- DP 23/07/08
                '-- Se non sono stati impostati i KM (VVTKMDEF), metto quelli della destinazione pi˘ lontana
                If IsNull(DS.Field("Km definitivi").NewValue) Or DS.Field("Km definitivi").NewValue = 0 Then
                    If IsNull(DS.Field("Km definitivi").Value) Or DS.Field("Km definitivi").Value = 0 Then
                        varBufDestinatari = varGetDestinatariViaggio()

                        If Not IsNull(DS.Field("Deposito").NewValue) And DS.Field("Deposito").NewValue <> vbNullString Then
                            strIDADE = CStr(DS.Field("Deposito").NewValue)
                        ElseIf Not IsNull(DS.Field("Deposito").Value) And DS.Field("Deposito").Value <> vbNullString Then
                            strIDADE = CStr(DS.Field("Deposito").Value)
                        End If

                        If Not IsEmpty(varBufDestinatari) Then
                            Set ADZ = New CDestinazione
                            With ADZ.DS
                                Set .DA = DS.DA
                                Set .ErrH = DS.ErrH
                            End With

                            For i = 0 To UBound(varBufDestinatari, 2)
                                varKmTemp = ADZ.GetDistanza(CStr(varBufDestinatari(0, i)), strIDADE)

                                If Not IsEmpty(varKmTemp) Then
                                    If CInt(varKmTemp(0, 0)) > intKmMax Then
                                        intKmMax = varKmTemp(0, 0)
                                    End If
                                End If
                            Next i

                            Set ADZ = Nothing

                            DS.Field("Km definitivi").NewValue = intKmMax
                            DS.Field("Km definitivi").Updated = True
                        End If
                    End If
                End If
            Case "E"
                If DS.Fields("42").Value <> "C" Then
                    Exit Function
                End If
                blnCalcolaOrario = False
            Case "C"
                If DS.Fields("42").Value <> "D" Then
                    Exit Function
                End If
                ' Forziamo la trasmissione ad Eurosystem
                SetFlagTrasmissioneESys "N"
                blnCalcolaOrario = False
            Case Else
                GoTo ErrHandler
        End Select
        .Item("0").Value = strIDViaggio
        .Item("0").FindValue = strIDViaggio
        .Item("42").NewValue = NewValue
        .Item("42").Updated = True
    End With
    
    blnResult = BeginSave(blnConcurErr, blnCalcolaOrario)
    
    If blnResult Then
        
        Dim VCA As CViaggioCarico
        
        Select Case NewValue
            Case "P"
                ' --- Cancello i dati di carico del viaggio
                Set VCA = New CViaggioCarico
                With VCA
                    Set .DS.DA = DS.DA
                    If Not .DeleteLoadData(strIDViaggio) Then
                        GoTo ErrHandler
                    End If
                End With
                Set VCA = Nothing
                ' ---
            Case "D"
                ' --- Inserisco i dati di carico del viaggio
                Set VCA = New CViaggioCarico
                With VCA
                    Set .DS.DA = DS.DA
                    If Not .InsertLoadData(strIDViaggio) Then
                        GoTo ErrHandler
                    End If
                End With
                Set VCA = Nothing
                ' ---
        End Select
        
        DS.DA.EndTrans 0
        SetStato = True
    Else
        DS.DA.EndTrans 1
        SetStato = False
    End If
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    If Not IsMissing(ErrDescr) Then
        ErrDescr = strError
    End If
    SetStato = False
End Function


Public Function CheckVentaEnRuta(NumeroViaggio As Variant, EnteEmittente As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    Dim strIDProdotto As String
    Dim strIDProdottoPrec As String
    Dim strCounterVentaEnRuta As String
    Dim strIDVVC As String
    Dim strIDVVP As String
    Dim strIDNumeratore As String
    Dim i As Long
    
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_ORDINI_VENTA_EN_RUTA, varBuffer, lngHowMany, _
                                   NumeroViaggio)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If Not IsEmpty(varBuffer) Then
        For i = 0 To UBound(varBuffer, 2)
            strIDProdotto = varBuffer(0, i)
            If strIDProdotto <> strIDProdottoPrec Then
                strIDProdottoPrec = strIDProdotto
                Select Case EnteEmittente
                    Case "004"
                        strIDNumeratore = "021"
                    Case "624"
                        strIDNumeratore = "020"
                    Case Else
                        GoTo ErrHandler
                End Select
                strCounterVentaEnRuta = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", strIDNumeratore, 0)
                If strCounterVentaEnRuta = CStr(-1) Then
                    GoTo ErrHandler
                End If
            End If
            strIDVVC = varBuffer(2, i)
            strIDVVP = varBuffer(3, i)
            If Not DS.DA.ExecuteCommand(SQL_SET_RIF_VENTA_EN_RUTA, strCounterVentaEnRuta, NumeroViaggio, strIDVVC, strIDVVP) Then
                GoTo ErrHandler
            End If
            
        Next i
    End If
    
    CheckVentaEnRuta = True
    Exit Function
ErrHandler:
    CheckVentaEnRuta = False
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

Public Sub GetTotaliProdotto()
    Dim strQ As String
    Dim varResult As Variant
    Dim blnOk As Boolean
    Dim lngHowMany As Long
    Dim dblTot As Double
    Dim i As Integer
    
    With DS.DA
        strQ = SQL_GET_TOTALE_QTA_PRODOTTO & .MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType) & SQL_GET_TOTALE_QTA_PRODOTTO_GROUPBY
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

Public Function GetListViaggiAutista(Data As Variant, IDAutista As String, ByRef ListViaggi As Variant, Optional ByRef ErrDesc As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strSQL As String
    Dim strError As String
    
    If Not IsDate(Data) Then
        strError = "Data non valida"
        GoTo ErrHandler
    End If
    
    strSQL = SQL_GETLIST_SET_DATI_STIPENDI & DS.DA.MakeSQL(Data, dtDate)
    If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, IDAutista) Then
        GoTo ErrHandler
    End If
    
    ListViaggi = varBuffer
    GetListViaggiAutista = True
    Exit Function
ErrHandler:
    ListViaggi = Empty
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    GetListViaggiAutista = False
End Function


Public Function GetUltimoViaggioCaricato(Motrice As String, Data As Variant, Ora As Variant, DatiViaggio As Variant, Optional ByRef ErrDesc As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strSQL As String
    Dim strError As String
    
    If Not IsDate(Data) Then
        strError = "Data non valida"
        GoTo ErrHandler
    End If
    
    strSQL = SQL_GET_ULTIMO_VIAGGIO_CARICATO & DS.DA.MakeSQL(Data, dtDate) & SQL_GET_ULTIMO_VIAGGIO_CARICATO_ORDERBY
    If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, Motrice) Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    
    DatiViaggio = varBuffer
    GetUltimoViaggioCaricato = True
    Exit Function
ErrHandler:
    DatiViaggio = Empty
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    GetUltimoViaggioCaricato = False
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

Public Function Insert(Optional ByRef ErrDesc As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim AAM As CAutomezzo
    Dim AAU As CAutista
    Dim strIDAutista As String
    Dim strError As String
    Dim varDataCarico As Variant
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    Set AAU = New CAutista
    With AAU
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH
    End With
    
    
    DS.Fields("42").NewValue = "P"
    DS.Fields("42").Updated = True
    DS.Fields("40").NewValue = "N"
    DS.Fields("40").Updated = True
    
'    If ((DS.Field("Flag Usa Dati GPS").Updated) And (Len(DS.Field("Flag Usa Dati GPS").NewValue) = 0)) Or (Not DS.Field("Flag Usa Dati GPS").Updated) Then
'        DS.Field("Flag Usa Dati GPS").NewValue = "S"
'        DS.Field("Flag Usa Dati GPS").Updated = True
'    End If
    
    AAM.DS.Field("Targa").FindValue = DS.Field("Motrice").NewValue
    AAM.Find
    DS.Field("Flag Usa Dati GPS").NewValue = IIf(AAM.DS.Field("Satellitare installato").Value = "N", "N", "S")
    DS.Field("Flag Usa Dati GPS").NewValue = "S"
    DS.Field("Flag Usa Dati GPS").Updated = True
    
    
    strIDAutista = DS.Field("Autista 1").NewValue
    varDataCarico = DS.Field("Data inizio carico").NewValue
    
    '-- DP 02/07/08
    '-- Aggiunti 2 campi contenenti data ed ora di riferimento
    DS.Field("Data di riferimento").NewValue = DS.Field("Data inizio carico").NewValue
    DS.Field("Data di riferimento").Updated = True
    DS.Field("Ora di riferimento").NewValue = DS.Field("Ora inizio carico").NewValue
    DS.Field("Ora di riferimento").Updated = True
    
    
    '-- DP 26/08/08
    DS.Field("Sequenza Invio VIA/GES").NewValue = Format(DS.Field("Sequenza viaggio").NewValue, "000")
    DS.Field("Sequenza Invio VIA/GES").Updated = True
    
    
    DS.DA.BeginTransaction
    
    
    If Not DS.Insert(Me.SQL_INSERT, "anaganu0", "anuprogr", "anuidanu", "002", 0) Then
        GoTo ErrHandler
    End If
    
    If Not AAU.ComputeOraInizioLavoro(strIDAutista, varDataCarico) Then
        strError = "Calcolo orario inizio lavoro fallito"
        GoTo ErrHandler
    End If
    
    DS.DA.EndTrans etCommit
    Insert = True
    Exit Function
    
ErrHandler:
    DS.DA.EndTrans etRollback
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    Insert = False
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean, Optional CalcolaOrario As Boolean = True) As Boolean
    Dim blnSetDateEff As Boolean
    Dim blnResult As Boolean
    Dim strDataViaggio As String
    Dim strSQLStatement As String
    Dim strIDAutista As String
    Dim strMotrice As String
    Dim strRimorchio As String
    Dim strIDUdT As String
    Dim varDataCarico As Variant
    
    Dim strSEQVI As String
    Dim blnCambiaSequenzaInvio As Boolean
    
    Dim AAU As CAutista
    
    If Not IsUpdatable Then
        Exit Function
    End If

    Set AAU = New CAutista
    With AAU
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH
    End With
    

    ' --- Se modifico i Km calcolati automaticamente
    ' --- setto il Flag calcolo Km
    If DS.Fields("25").Updated And DS.Fields("26").Value = "A" Then
        DS.Fields("26").NewValue = "M"
        DS.Fields("26").Updated = True
    End If
    If DS.Fields("27").Updated And DS.Fields("28").Value = "A" Then
        DS.Fields("28").NewValue = "M"
        DS.Fields("28").Updated = True
    End If
    ' ---
    
    ' Imposta il default dello stato del viaggio satellitare
    If DS.Fields("6").Updated Then
        strDataViaggio = DS.Fields("6").NewValue
        blnSetDateEff = True
    End If
    
    
    ' Impostiamo il flag di trasmissione a Eurosystem
    DS.Fields.Item("40").NewValue = "N"
    DS.Fields.Item("40").Updated = True
    
    
    If DS.Field("Unit‡ di Trasporto").Updated Then
        strIDUdT = DS.Field("Unit‡ di Trasporto").NewValue
    Else
        strIDUdT = DS.Field("Unit‡ di Trasporto").Value
    End If

    If DS.Field("Data inizio carico").Updated Then
        varDataCarico = DS.Field("Data inizio carico").NewValue
    Else
        varDataCarico = DS.Field("Data inizio carico").Value
    End If
    

    '-- DP 26/08/08
    '-- Gestione sequenza di invio
    If DS.Field("Data inizio viaggio").Updated Then
        If DS.Field("Data inizio viaggio").NewValue <> DS.Field("Data inizio viaggio").Value And DS.Field("Sequenza Invio Trasmessa").Value <> "N" Then
            blnCambiaSequenzaInvio = True
        End If
    End If
    
    If DS.Field("Unit‡ di Trasporto").Updated Then
        If DS.Field("Unit‡ di Trasporto").NewValue <> DS.Field("Unit‡ di Trasporto").Value And DS.Field("Sequenza Invio Trasmessa").Value <> "N" Then
            blnCambiaSequenzaInvio = True
        End If
    End If
    
    If DS.Field("Sequenza Viaggio").Updated Then
        If DS.Field("Sequenza Viaggio").NewValue <> DS.Field("Sequenza Viaggio").Value Then
            blnCambiaSequenzaInvio = True
        End If
    
        strSEQVI = DS.Field("Sequenza Viaggio").NewValue
    Else
        strSEQVI = DS.Field("Sequenza Viaggio").Value
    End If
        
    If blnCambiaSequenzaInvio Then
        If Not IsNull(DS.Field("Sequenza Invio VIA/GES").Value) Then
            If DS.Field("Sequenza Invio VIA/GES").Value <> vbNullString Then
                If Right(DS.Field("Sequenza Invio VIA/GES").Value, 1) = Right(strSEQVI, 1) Then
                    DS.Field("Sequenza Invio VIA/GES").NewValue = Format(CInt(DS.Field("Sequenza Invio VIA/GES").Value) + 10, "000")
                    DS.Field("Sequenza Invio VIA/GES").Updated = True
                Else
                    DS.Field("Sequenza Invio VIA/GES").NewValue = Format(strSEQVI, "000")
                    DS.Field("Sequenza Invio VIA/GES").Updated = True
                End If
            Else
                DS.Field("Sequenza Invio VIA/GES").NewValue = Format(CInt(strSEQVI) + 10, "000")
                DS.Field("Sequenza Invio VIA/GES").Updated = True
            End If
        Else
            DS.Field("Sequenza Invio VIA/GES").NewValue = Format(CInt(strSEQVI) + 10, "000")
            DS.Field("Sequenza Invio VIA/GES").Updated = True
        End If
        
        DS.Field("Sequenza Invio Trasmessa").NewValue = "N"
        DS.Field("Sequenza Invio Trasmessa").Updated = True
    End If
    '-- Fine gestione sequenza di invio
    
    
    blnResult = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    If blnResult Then
        If blnSetDateEff Then
            blnResult = DS.DA.ExecuteCommand(SQL_SET_DATE_CONSEGNE, _
                                             strDataViaggio, _
                                             DS.Fields("0").Value)
            strSQLStatement = " UPDATE ORDIOCT0 " & _
                                " SET OCTDTCEF = " & DS.DA.MakeSQL(strDataViaggio, dtDate) & "," & _
                                "     OCTDTCPI = " & DS.DA.MakeSQL(strDataViaggio, dtDate) & _
                                " WHERE OCTIDOCT IN " & _
                                " (SELECT VVPIDOCT FROM VIAGVVP0 WHERE VVPIDVVT = " & DS.DA.MakeSQL(DS.Fields("0").Value, dtAlfa) & " AND VVPANNUL = 0)"

            blnResult = DS.DA.ExecuteCommand(strSQLStatement)
        End If
    End If
    
    If DS.Field("Autista 1").Updated Then
        strIDAutista = DS.Field("Autista 1").NewValue
        SMS_Insert DS.Fields("0").Value, strIDAutista, True
    Else
        strIDAutista = DS.Field("Autista 1").Value
    End If
    If DS.Field("Motrice").Updated Then
        strMotrice = DS.Field("Motrice").NewValue
    Else
        strMotrice = DS.Field("Motrice").Value
    End If
    If DS.Field("Rimorchio").Updated Then
        strRimorchio = DS.Field("Rimorchio").NewValue
    Else
        strRimorchio = DS.Field("Rimorchio").Value
    End If
    
    If Len(strIDUdT) > 0 Then
    End If
    If Len(strIDAutista) > 0 Then
    End If
    If Len(strMotrice) > 0 Then
    End If
    If Len(strRimorchio) > 0 Then
    End If
    
    If CalcolaOrario Then
        blnResult = AAU.ComputeOraInizioLavoro(strIDAutista, varDataCarico)
    End If
    
    BeginSave = blnResult
    
    ' ---
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
               "       vvtidcn1, vvtnripa, vvtnrkvi, vvtflsta, " & _
               "       vvtkmvuo, vvtfludg, vvtkmrie, vvtkmrir, vvtflagg, " & _
               "       vvtdtrif, vvthrrif, vvtseqin, vvtseqok, " & _
               "       vvtannul, timestmp FROM viagvvt0 WHERE "
    
    SQL_GETLIST = "SELECT vvtidvvt, vvtidasc FROM viagvvt0 "
    SQL_UPDATE = "UPDATE viagvvt0 SET "
    SQL_INSERT = "INSERT INTO viagvvt0 "
    SQL_DELETE = "DELETE FROM viagvvt0 WHERE vvtidvvt = "
    SQL_SET_STATO = "UPDATE viagvvt0 SET vvtflstv = "
'    SQL_SET_DATE_CONSEGNE_EFF = "UPDATE ordioct0 SET octdtcef = ? WHERE octidoct IN " & _
'                                " (SELECT vvpidoct FROM viagvvp0 WHERE vvpidvvt = ?) "
    
    SQL_SET_DATE_CONSEGNE_EFF = " UPDATE (VIAGVVT0 INNER JOIN VIAGVVP0 ON " & _
                                " VIAGVVT0.VVTIDVVT = VIAGVVP0.VVPIDVVT) INNER JOIN " & _
                                " ORDIOCT0 ON VIAGVVP0.VVPIDOCT = ORDIOCT0.OCTIDOCT " & _
                                " SET ORDIOCT0.OCTDTCEF = ?, ORDIOCT0.octdtcpi = ? WHERE (((VIAGVVP0.VVPIDVVT)=?));"

    SQL_SET_DATE_CONSEGNE = " UPDATE VIAGVVC0 " & _
                            "    SET VVCDTCON = ? " & _
                            "  WHERE (((VIAGVVC0.VVCIDVVT)=?))"
    
    SQL_GET_TS = "SELECT timestmp FROM viagvvt0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT vvtidvvt, vvtiddef, vvtseqvi, " & _
                                 "vvtidasc, vvtdtinv, vvtdtfiv, vvtidaut, " & _
                                 "vvtidam1, vvtidam2, " & _
                                 "vvtannul FROM viagvvt0 "
    SQL_GET_TOTALE_QTA = "SELECT SUM(vvpqtrkg),SUM(vvpqtrlt),SUM(vvpqtrpz),SUM(vvpqtrcn), " & _
                         "       SUM(vvpqcokg),SUM(vvpqcolt),SUM(vvpqcopz),SUM(vvpqcocn) " & _
                         "  FROM viagvvc0 INNER JOIN viagvvp0 on vvcidvvt  = vvpidvvt AND vvcidvvc = vvpidvvc WHERE vvcannul = 0 and vvpannul = 0 and vvpidvvt = "
    SQL_GET_TOTALE_QTA_PRODOTTO = "SELECT vvpidapr, SUM(vvpqtrkg),SUM(vvpqtrlt),SUM(vvpqtrpz),SUM(vvpqtrcn), " & _
                                  "       SUM(vvpqcokg),SUM(vvpqcolt),SUM(vvpqcopz),SUM(vvpqcocn) " & _
                                  "  FROM viagvvc0 INNER JOIN viagvvp0 on vvcidvvt  = vvpidvvt AND vvcidvvc = vvpidvvc WHERE vvcannul = 0 and vvpannul = 0 and vvpidvvt = "
    SQL_GET_TOTALE_QTA_PRODOTTO_GROUPBY = " GROUP BY vvpidapr"
    SQL_COPY_CONSEGNE = "INSERT INTO VIAGVVC0 (VVCIDVVT, VVCIDVVC, VVCNRSEQ, VVCIDASC, " & _
                       "VVCIDACL, VVCRASOC, VVCINDIR, VVCLOCAL, VVCPROVI, " & _
                       "VVCNRCAP, VVCCOFIS, VVCPAIVA, VVCIDADZ, VVCDERAS, " & _
                       "VVCDEIND, VVCDELOC, VVCDEPRO, VVCDECAP, VVCDECFI, " & _
                       "VVCDEPIV, VVCKMPAP, VVCFLCPP, VVCKMPAD, VVCFLCPD, " & _
                       "VVCTEMVG, VVCTEMSC, VVCIDAZO, VVCIDATV, VVCIDVT1, " & _
                       "VVCIDVC1, VVCRIES1, VVCRIES2, VVCFLST1, VVCFLST2, " & _
                       "VVCFLST3, VVCFLST4, VVCIDOCT, VVCANNUL) "
    SQL_FIELDS_CONSEGNE = " VVCIDVVC, VVCNRSEQ, VVCIDASC, VVCIDACL, VVCRASOC, " & _
                       " VVCINDIR, VVCLOCAL, VVCPROVI, VVCNRCAP, VVCCOFIS, VVCPAIVA, " & _
                       " VVCIDADZ, VVCDERAS, VVCDEIND, VVCDELOC, VVCDEPRO, " & _
                       " VVCDECAP, VVCDECFI, VVCDEPIV, VVCKMPAP, VVCFLCPP, " & _
                       " VVCKMPAD, VVCFLCPD, VVCTEMVG, VVCTEMSC, VVCIDAZO, " & _
                       " VVCIDATV, VVCIDVVT, VVCIDVVC, VVCRIES1, VVCRIES2, VVCFLST1, " & _
                       " VVCFLST2 , VVCFLST3, VVCFLST4, VVCIDOCT, VVCANNUL "
    SQL_COPY_SCARICHI = "INSERT INTO VIAGVVP0 (VVPIDVVT, VVPIDVVC, VVPIDVVP, VVPIDAPR, VVPIDAIM, " & _
                       " VVPIDAPD, VVPQTRKG, VVPQTRLT, VVPQTRPZ, VVPQTRCN, " & _
                       " VVPQCOKG, VVPQCOLT, VVPQCOPZ, VVPQCOCN, VVPPRINC, " & _
                       " VVPPRFIC, VVPIDVT1, VVPIDVC1, VVPIDVP1, VVPRIFE1, VVPRIFE2, " & _
                       " VVPIDAIV, VVPIDAMP, VVPPREZZ, VVPSCPER, VVPSCVAL, " & _
                       " VVPFLTIR, VVPFLST1, VVPFLST2, VVPFLST3, VVPFLST4, " & _
                       " VVPIDOCT, VVPIDOCR, VVPANNUL)"
    SQL_FIELDS_SCARICHI = " VVPIDVVC, VVPIDVVP, VVPIDAPR, VVPIDAIM, VVPIDAPD, " & _
                   " VVPQTRKG, VVPQTRLT, VVPQTRPZ, VVPQTRCN, VVPQCOKG, " & _
                   " VVPQCOLT, VVPQCOPZ, VVPQCOCN, VVPPRINC, VVPPRFIC, " & _
                   " VVPIDVVT, VVPIDVVC, VVPIDVVP, VVPRIFE1, VVPRIFE2, VVPIDAIV, " & _
                   " VVPIDAMP, VVPPREZZ, VVPSCPER, VVPSCVAL, VVPFLTIR, " & _
                   " VVPFLST1, VVPFLST2, VVPFLST3, VVPFLST4, VVPIDOCT, VVPIDOCR, " & _
                   " VVPANNUL "
    
    SQL_GET_QTA_SCARICHI = "SELECT vvpqcokg, vvpqcolt, vvpqcopz, vvpqcocn, " & _
                           "       vvpidoct, vvpidocr FROM viagvvp0 WHERE vvpidvvt = "
    SQL_COUNT_SIF = "SELECT Count(MULTSPD0.SPDCODSO) AS C " & _
                    " FROM VIAGVVP0 LEFT JOIN MULTSPD0 ON VIAGVVP0.VVPIDAPD = MULTSPD0.SPDIDAPD " & _
                    " WHERE (((MULTSPD0.SPDCODSO)=?) AND ((MULTSPD0.SPDIDASC)=?) " & _
                    " AND ((VIAGVVP0.VVPIDVVT)=?) AND ((VIAGVVP0.VVPANNUL)=0))"
    SQL_COUNT_DESTINATARI = "SELECT COUNT(*) " & _
                            "  FROM VIAGVVC0 INNER JOIN VIAGVVP0 ON VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC " & _
                            " WHERE VVCIDVVT = ? AND VVCANNUL = 0 AND VVPANNUL = 0"
    SQL_COUNT_DESTINAZIONI = "SELECT DISTINCT VVCIDADZ " & _
                            "  FROM VIAGVVC0 " & _
                            " WHERE VVCIDVVT = ? AND VVCANNUL = 0 "
    SQL_COUNT_CONSEGNE = "SELECT COUNT(*) " & _
                            "  FROM VIAGVVC0 " & _
                            " WHERE VVCIDVVT = ? AND VVCANNUL = 0 "
    SQL_SELECT_TRIP_PRODUCTS = "SELECT DISTINCT VVPIDVVT, VVPIDAPR " & _
                                  "  FROM VIAGVVT0 INNER JOIN (VIAGVVC0 INNER JOIN VIAGVVP0 " & _
                                  "       ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) " & _
                                  "       ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT "
    SQL_INSERT_TRIP_PRODUCTS = "INSERT INTO VIAGTMP0 VALUES (?,?,?,?,?,?,?,?,?,?,?)"
    SQL_CLEAR_TRIP_PRODUCTS = "DELETE FROM VIAGTMP0 "
    SQL_GETLIST_COMPUTE_CU_REPARTO = "SELECT VVCIDVVC, VVCKMRAG, VVPQCOLT, APRDECON, " & _
                                     "       VVTKMDEF " & _
                                     "  FROM VIAGVVP0 INNER JOIN VIAGVVC0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC INNER JOIN " & _
                                     "       VIAGVVT0 ON VVCIDVVT = VVTIDVVT INNER JOIN ANAGAPR0 ON VVPIDAPR = APRIDAPR " & _
                                     " WHERE (VVCANNUL = 0) And (VVPANNUL = 0) And (VVPIDVVT = ?) " & _
                                     " ORDER BY VVCIDVVC "
    SQL_GETLIST_SET_DATI_STIPENDI = "SELECT VVTIDVVT, VVTIDDEF, VVTSEQVI, VVTIDADE, " & _
                                    "       ADEDESCR, VVTKMDEF, VVTNRIPA, VVTFLST2 " & _
                                    "  FROM VIAGVVT0 LEFT OUTER JOIN ANAGADE0 ON VVTIDADE = ADEIDADE " & _
                                    " WHERE VVTANNUL = 0 AND VVTFLSTV <> 'P' AND VVTIDAU1 = ? AND VVTDTINV = "
    
    ' [28-04-05 (PM)] Nuovi statement per interfacciamento KNOSOS
    
    ' Statement relativi al KVIAGGIO
    SQL_KNOSOS_INSERT_KVIAGGIO = "INSERT INTO KNOSVIA0" & _
                        " (IDVEH, IDVIAJE, NUMVIAJE, FECHA, ESTA)" & _
                        " VALUES (?, ?, ?, {Ts '%fecha%'}, ?)"
    
    SQL_KNOSOS_SELECT_KVIAGGIO = "SELECT IDVEH, IDVIAJE, NUMVIAJE, CONVERT(datetime, FECHA, 21), ESTA" & _
                        " FROM KNOSVIA0" & _
                        " WHERE NUMVIAJE = ?" & _
                        " AND FECHA = (SELECT MAX(FECHA) FROM KNOSVIA0 KVIA WHERE KVIA.NUMVIAJE = KNOSVIA0.NUMVIAJE)"
    
    ' Statement relativi al KCARICO
    SQL_KNOSOS_INSERT_KCARICO = "INSERT INTO KNOSOPE0" & _
                        " (IDVEH, IDVIAJE, NUMVIAJE, IDDESC, FECHA, TOPE," & _
                        " ESTADO, DEPOSITO, FECHA_PREVISTA)" & _
                        " VALUES (?, ?, ?, ?, {Ts '%fecha%'}, ?, ?, ?, ?)"

'    SQL_KNOSOS_SELECT_KCARICO = "SELECT IDVEH, IDVIAJE, NUMVIAJE, CONVERT(datetime, FECHA, 21)," & _
'                        "ESTADO, DEPOSITO, FECHA_PREVISTA" & _
'                        " FROM KNOSOPE0" & _
'                        " WHERE NUMVIAJE = ? AND IDDESC = ?" & _
'                        " AND FECHA = (SELECT MAX(FECHA) FROM KNOSOPE0 KOPE WHERE KOPE.NUMVIAJE = KNOSOPE0.NUMVIAJE AND KOPE.IDDESC = KNOSOPE0.IDDESC)"
    
    SQL_KNOSOS_SELECT_KCARICO = "SELECT IDVEH, IDVIAJE, NUMVIAJE, CONVERT(datetime, FECHA, 21)," & _
                        "ESTADO, DEPOSITO, FECHA_PREVISTA" & _
                        " FROM KNOSOPE0" & _
                        " WHERE NUMVIAJE = ? AND TOPE = ? AND IDDESC = ?"

    SQL_KNOSOS_UPDATE_KCARICO = "UPDATE KNOSOPE0" & _
                        " SET ESTADO = ?, DEPOSITO = ?, FECHA_PREVISTA = ?" & _
                        " WHERE NUMVIAJE = ? AND TOPE = ? AND IDDESC = ?"
    
    ' Statement relativi al KSCARICO
    SQL_KNOSOS_INSERT_KSCARICO = "INSERT INTO KNOSOPE0" & _
                        " (IDVEH, IDVIAJE, NUMVIAJE, IDDESC, FECHA, TOPE," & _
                        " ESTADO, DESTINO, POBLACION, CALLE_NUM, TELEFONO, HSOLICITADO," & _
                        " HAPERTURA, HPREVLLEG, PRODUCTO, PROD_RES, LITROS, COD_ORDEN, NOTAS1," & _
                        " NOTAS2)" & _
                        " VALUES (?, ?, ?, ?, {Ts '%fecha%'}, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
'    SQL_KNOSOS_SELECT_KSCARICO = "SELECT IDVEH, IDVIAJE, NUMVIAJE, IDDESC, CONVERT(datetime, FECHA, 21), " & _
'                        " ESTADO, DESTINO, POBLACION, CALLE_NUM, TELEFONO, HSOLICITADO, " & _
'                        " HAPERTURA, HPREVLLEG, PRODUCTO, PROD_RES, LITROS, COD_ORDEN, NOTAS1, NOTAS2" & _
'                        " FROM KNOSOPE0" & _
'                        " WHERE NUMVIAJE = ? AND IDDESC = ?" & _
'                        " AND FECHA = (SELECT MAX(FECHA) FROM KNOSOPE0 KOPE WHERE KOPE.NUMVIAJE = KNOSOPE0.NUMVIAJE AND KOPE.IDDESC = KNOSOPE0.IDDESC)"
    
    SQL_KNOSOS_SELECT_KSCARICO = "SELECT IDVEH, IDVIAJE, NUMVIAJE, IDDESC, CONVERT(datetime, FECHA, 21), " & _
                        " ESTADO, DESTINO, POBLACION, CALLE_NUM, TELEFONO, HSOLICITADO, " & _
                        " HAPERTURA, HPREVLLEG, PRODUCTO, PROD_RES, LITROS, COD_ORDEN, NOTAS1, NOTAS2" & _
                        " FROM KNOSOPE0" & _
                        " WHERE NUMVIAJE = ? AND TOPE = ? AND IDDESC = ?"

    SQL_KNOSOS_UPDATE_KSCARICO = "UPDATE KNOSOPE0" & _
                        " SET ESTADO = ?, DESTINO = ?, POBLACION = ?, CALLE_NUM = ?," & _
                        " TELEFONO = ?, HSOLICITADO = ?, HAPERTURA = ?, HPREVLLEG = ?," & _
                        " PRODUCTO = ?, PROD_RES = ?, LITROS = ?, COD_ORDEN = ?," & _
                        " NOTAS1 = ?, NOTAS2 = ?" & _
                        " WHERE NUMVIAJE = ? AND TOPE = ? AND IDDESC = ?"

    SQL_KNOSOS_UPDATESTATO_KSCARICO = "UPDATE KNOSOPE0" & _
                        " SET ESTADO = ?" & _
                        " WHERE NUMVIAJE = ? AND TOPE = ? AND IDDESC = ?"
    
    SQL_KNOSOS_SELECT_SCARICHI_RESOIL = "SELECT VIAGVVP0.VVPIDVVT, VIAGVVP0.VVPIDVVC, VIAGVVP0.VVPIDVVP," & _
                            " VIAGVVP0.VVPIDAPR, VIAGVVC0.VVCDERAS, VIAGVVC0.VVCDELOC," & _
                            " VIAGVVC0.VVCDEIND, ANAGADZ0.ADZNRTEL, ORDIOCT0.OCTNOTA1," & _
                            " ORDIOCT0.OCTHRFR1, ORDIOCT0.OCTHRTO1, VIAGVVC0.VVCTEMVG," & _
                            " VIAGVVP0.VVPQTRLT, ORDIOCT0.OCTNOTA3, VIAGVVP0.VVPIDKDE," & _
                            " ANAGAPR0.APRDESAB, VIAGVVP0.VVPFLSTA FROM VIAGVVP0" & _
                            " INNER JOIN VIAGVVC0 ON VIAGVVP0.VVPIDVVT = VIAGVVC0.VVCIDVVT AND VIAGVVP0.VVPIDVVC = VIAGVVC0.VVCIDVVC" & _
                            " INNER JOIN VIAGVVT0 ON VIAGVVC0.VVCIDVVT = VIAGVVT0.VVTIDVVT" & _
                            " INNER JOIN ANAGADZ0 ON VIAGVVC0.VVCIDADZ = ANAGADZ0.ADZIDADZ" & _
                            " INNER JOIN ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT" & _
                            " INNER JOIN ANAGAPR0 ON VIAGVVP0.VVPIDAPR = ANAGAPR0.APRIDAPR" & _
                            " WHERE VIAGVVT0.VVTIDVVT = ?"

'    SQL_KNOSOS_SELECT_KSCARICHI_IDDESC = "SELECT IDDESC FROM KNOSOPE0 " & _
'                                    " WHERE NUMVIAJE = ? AND TOPE = 1" & _
'                                    " AND FECHA = (SELECT MAX(FECHA) FROM KNOSOPE0 KOPE WHERE KOPE.NUMVIAJE = KNOSOPE0.NUMVIAJE AND KOPE.IDDESC = KNOSOPE0.IDDESC)" & _
'                                    " AND ESTADO NOT IN (4,5)"

    SQL_KNOSOS_SELECT_KSCARICHI_IDDESC = "SELECT IDDESC FROM KNOSOPE0" & _
                                    " WHERE NUMVIAJE = ? AND TOPE = 1" & _
                                    " AND ESTADO <> 4 AND ESTADO <> 5"
    
    SQL_KNOSOS_RESET_IDDESC_ANNULLATE = "UPDATE VIAGVVP0 SET VVPIDKDE = NULL WHERE VVPIDVVT = ?" & _
                                    " AND VVPIDKDE IS NOT NULL AND" & _
                                    " (VVPANNUL = 1 OR (SELECT VVCANNUL FROM VIAGVVC0 WHERE VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC) = 1)"

    SQL_KNOSOS_GET_NEXTID_SCARICO = "SELECT MAX(IDDESC) FROM KNOSOPE0 WHERE NUMVIAJE = ?"
    SQL_KNOSOS_UPDATE_NUMERO_SCARICO = "UPDATE VIAGVVP0 SET VVPIDKDE = ? WHERE VVPIDVVT = ? AND VVPIDVVC = ? AND VVPIDVVP = ?"
    SQL_UPDATE_KNOSOS_DATI = "UPDATE VIAGVVT0 SET VVTNRKVI = ?, VVTFLSTA = ? WHERE VVTIDVVT = ?"
    
    SQL_GET_ULTIMO_VIAGGIO_CARICATO = "SELECT TOP 1 VVTIDVVT, VVTIDAEE, VVTIDAUT, SUTCODSO, VVTIDADE, VVTIDAM2 " & _
                                      "  FROM VIAGVVT0 INNER JOIN ANAGADE0 ON VVTIDADE = ADEIDADE " & _
                                      "       INNER JOIN MULTSUT0 ON SUTIDAUT = VVTIDAUT AND SUTIDASC = ADEIDASC " & _
                                      " WHERE VVTANNUL = 0 AND VVTIDAM1 = ? AND VVTFLST3 = 'S' AND VVTDTINV =  "
    SQL_GET_ULTIMO_VIAGGIO_CARICATO_ORDERBY = " ORDER BY VVTDTINC DESC, VVTHRINC DESC "
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "6", "8", "14", _
                              "15", "16", "44")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Viaggio", "viagvvt0.vvtidvvt", 0, True, 7, "0"
        .Add "Numero Viaggio definitivo", "viagvvt0.vvtiddef", 0, , 7, "1", , True
        .Add "Sequenza Viaggio", "viagvvt0.vvtseqvi", 0, True, 2, "2"
        .Add "Societ‡", "viagvvt0.vvtidasc", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Data inizio carico", "viagvvt0.vvtdtinc", 2, False, 10, "4"
        .Add "Ora inizio carico", "viagvvt0.vvthrinc", 4, False, 5, "5"
        .Add "Data inizio viaggio", "viagvvt0.vvtdtinv", 2, False, 10, "6"
        .Add "Ora inizio viaggio", "viagvvt0.vvthrinv", 4, False, 5, "7"
        .Add "Data fine viaggio", "viagvvt0.vvtdtfiv", 2, False, 10, "8"
        .Add "Ora fine viaggio", "viagvvt0.vvthrfiv", 4, False, 5, "9"
        .Add "Zona", "viagvvt0.vvtidazo", 0, False, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Badge Automezzo", "viagvvt0.vvtbamez", 0, , 5, "11"
        .Add "Badge Autista", "viagvvt0.vvtbaaut", 0, , 5, "12"
        .Add "Tipo Vendita", "viagvvt0.vvtidatv", 0, False, 3, "13"
        With .Item("13")
            .ForeignTableEntity = "ATV"
            .ForeignTable = "anagatv0"
            .ForeignTableFld = "atvidatv"
            .ForeignTableDescField = "atvdescr"
            .ForeignTableDeleteFld = "atvannul"
        End With
        .Add "Unit‡ di Trasporto", "viagvvt0.vvtidaut", 0, True, 3, "14"
        With .Item("14")
            .ForeignTableEntity = "AUT"
            .ForeignTable = "anagaut0"
            .ForeignTableFld = "autidaut"
            .ForeignTableDescField = "autdescr"
            .ForeignTableDeleteFld = "autannul"
        End With
        .Add "Motrice", "viagvvt0.vvtidam1", 0, False, 10, "15"
        With .Item("15")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Rimorchio", "viagvvt0.vvtidam2", 0, False, 10, "16"
        With .Item("16")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Vettore", "viagvvt0.vvtidve1", 0, False, 3, "17"
        With .Item("17")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Vettore Trasportatore", "viagvvt0.vvtidve2", 0, False, 3, "18"
        With .Item("18")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Autista 1", "viagvvt0.vvtidau1", 0, False, 3, "19"
        With .Item("19")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Autista 2", "viagvvt0.vvtidau2", 0, False, 3, "20"
        With .Item("20")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Ente Emittente", "viagvvt0.vvtidaee", 0, True, 3, "21"
        With .Item("21")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Deposito", "viagvvt0.vvtidade", 0, True, 3, "22"
        With .Item("22")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Magazzino di deposito", "viagvvt0.vvtidamd", 0, True, 3, "23"
        With .Item("23")
            .ForeignTableEntity = "AMD"
            .ForeignTable = "anagamd0"
            .ForeignTableFld = "amdidamd"
            .ForeignTableDescField = "amddescr"
            .ForeignTableDeleteFld = "amdannul"
        End With
        .Add "Viaggio di Riferimento", "viagvvt0.vvtidvv1", 0, False, 7, "24", , True
        .Add "Km previsti", "viagvvt0.vvtkmpre", 1, False, 6, "25"
        .Add "Flag calcolo km previsti", "viagvvt0.vvtflcap", 0, True, 1, "26"
        .Add "Km definitivi", "viagvvt0.vvtkmdef", 1, False, 6, "27"
        .Add "Flag calcolo km definitivi", "viagvvt0.vvtflcad", 0, True, 1, "28"
        .Add "Tempo viaggio", "viagvvt0.vvttmpvi", 1, False, 6, "29"
        .Add "Tempo carico", "viagvvt0.vvttmpca", 1, False, 6, "30"
        .Add "Rientro/Dirottamento", "viagvvt0.vvtflrid", 0, False, 1, "31"
        .Add "Deposito fine viaggio", "viagvvt0.vvtidad1", 0, False, 3, "32"
        With .Item("32")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Ente Emittente fine viaggio", "viagvvt0.vvtidee1", 0, False, 3, "33"
        With .Item("33")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Destinazione fine viaggio", "viagvvt0.vvtidadz", 0, False, 10, "34"
        With .Item("34")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Rif. Esterno 1", "viagvvt0.vvtrife1", 0, False, 25, "35"
        .Add "Rif. Esterno 2", "viagvvt0.vvtrife2", 0, False, 25, "36"
        .Add "Flag stato 1", "viagvvt0.vvtflst1", 0, False, 2, "37"
        
        ' *** Flag giorno festivo
        .Add "Flag stato 2", "viagvvt0.vvtflst2", 0, False, 2, "38"
        
        ' *** Flag carico in deposito effettuato
        .Add "Flag stato 3", "viagvvt0.vvtflst3", 0, False, 2, "39"
        ' ***
        ' *** Flag trasmissione ad Eurosystem
        .Add "Flag stato 4", "viagvvt0.vvtflst4", 0, False, 2, "40"
        ' ***
        .Add "Flag stato 5", "viagvvt0.vvtflst5", 0, False, 2, "41", , True
        .Add "Stato Viaggio", "viagvvt0.vvtflstv", 0, False, 1, "42"
        .Add "Tipo Viaggio", "viagvvt0.vvtfltiv", 0, True, 1, "43"
        .Add "Contratto", "viagvvt0.vvtidcn1", 0, False, 3, "45"
        With .Item("45")
            .ForeignTableEntity = "ACN"
            .ForeignTable = "anagacn0"
            .ForeignTableFld = "acnidacn"
            .ForeignTableDescField = "acndescr"
            .ForeignTableDeleteFld = "acnannul"
        End With
        .Add "Importo pedaggi", "viagvvt0.vvtnripa", 1, False, 13, "46"
        .Add "KNOSOS Num. Viaggio", "viagvvt0.vvtnrkvi", 0, False, 6, "47"
        .Add "Stato Satellitare", "viagvvt0.vvtflsta", 0, False, 2, "48"
        .Add "KM a vuoto", "viagvvt0.vvtkmvuo", 1, False, 6, "49"
'        .Add "Contratto secondario", "viagvvt0.vvtidcn2", 0, False, 3, "46"
'        With .Item("46")
'            .ForeignTableEntity = "ACN"
'            .ForeignTable = "anagacn0"
'            .ForeignTableFld = "acnidacn"
'            .ForeignTableDescField = "acndescr"
'            .ForeignTableDeleteFld = "acnannul"
'        End With
        .Add "Flag Usa Dati GPS", "viagvvt0.vvtfludg", 0, False, 2, "50"
        .Add "KM rientro", "viagvvt0.vvtkmrie", 1, False, 6, "51"
        
        '-- DP 10/07/09
        '-- Aggiunto questo campo per gestire sia i Km di rientro del Mover che quelli del satellitare
        .Add "KM rientro reali", "viagvvt0.vvtkmrir", 1, False, 6, "57"
        
        '-- DP 05/06/08
        '-- Aggiunto "VVTFLAGG" per l'indicazione del processo che ha aggiornato il record l'ultima volta ('ME' = Messaggio SMS; 'SE' = Finestra di Seguimiento)
        .Add "Flag ultimo aggiornamento", "viagvvt0.vvtflagg", 0, False, 2, "52", , True
        
        '-- DP 02/07/08
        '-- Aggiunti 2 campi per l'indicazione di Ora e Data di Riferimento
        .Add "Data di riferimento", "viagvvt0.vvtdtrif", 2, False, 10, "53"
        .Add "Ora di riferimento", "viagvvt0.vvthrrif", 4, False, 5, "54"
        
        '-- DP 14/08/08
        '-- Aggiunto "VVTSEQIN" per l'indicazione della sequenza di invio su files VIA e GES
        .Add "Sequenza Invio VIA/GES", "viagvvt0.vvtseqin", 0, False, 3, "55", , True
        
        '-- DP 09/09/08
        '-- Aggiunto "VVTSEQOK" per indicare se la sequenza di invio Ë gi‡ stata trasmessa su VIA o GES.
        '-- FinchÈ il valore Ë a 'N' (DEFAULT), la sequenza di invio verr‡ modificata solamente se
        '-- viene modificata la sequenza viaggio.
        '-- Appena avviene una trasmissione VIA o GES riguardante il viaggio, il valore passa ad 'S' e la sequenza
        '-- di invio puÚ essere modificata (per cambio su VVTDTINV o VVTIDUDT)
        .Add "Sequenza Invio Trasmessa", "viagvvt0.vvtseqok", 0, False, 2, "56", , True
        
        .Add "Annullato", "viagvvt0.vvtannul", 3, , , "44"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "viagvvt0.vvtannul"
    
End Sub

Public Function Delete() As Boolean
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim lngHowMany As Long
    Dim i As Long
    Dim strIDViaggio As String
    Dim VVC As CViaggioConsegna
    Dim objSatError As CError
        
    Dim VCA As CViaggioCarico
    
    DS.DA.BeginTransaction
    
    If Not IsUpdatable() Then
        GoTo ErrHandler
    End If
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        strIDViaggio = DS.Fields("0").Value
        .DS.Fields("0").FindValue = strIDViaggio
        If .GetList Then
            lngHowMany = .DS.HowMany - 1
            For i = 0 To lngHowMany
                .DS.Fields("1").FindValue = .DS.List(0, i)
                .Find
                .Delete
            Next i
        End If
    End With
    
    ' --- Cancello i dati di carico del viaggio
    Set VCA = New CViaggioCarico
    With VCA
        Set .DS.DA = DS.DA
        If Not .DeleteLoadData(strIDViaggio) Then
            GoTo ErrHandler
        End If
    End With
    ' ---
    
    strQ = SQL_DELETE & DS.DA.MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
    
    If Not DS.DA.ExecuteQuery(strQ) Then
        GoTo ErrHandler
    End If
    
'    Set objSatError = New CError
'    If Not Delete_Sat(objSatError) Then
'        GoTo ErrHandler
'    End If
    Set objSatError = Nothing

    DS.DA.EndTrans 0
    
    Set VVC = Nothing
    
    Set VCA = Nothing
    
    Delete = True
    
    Exit Function
    
ErrHandler:
    DS.DA.EndTrans 1
    Set VVC = Nothing
    Set VCA = Nothing
    
    If Not objSatError Is Nothing Then
        DS.DA.ErrH = objSatError
        Set objSatError = Nothing
    End If

    Delete = False
End Function



Public Property Get TotaliQta(Index As Integer) As Variant
    If (Index >= LBound(mvarTotaliQta, 1)) And _
        (Index <= UBound(mvarTotaliQta, 1)) Then
        TotaliQta = mvarTotaliQta(Index, 0)
    End If
End Property

Public Property Get TotaliQtaProdotto() As Variant
    TotaliQtaProdotto = mvarTotaliQta
End Property

Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


Private Sub DS_ClientClear()
    RaiseEvent ClientClear

End Sub

Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub
Private Function KNOSOS_GetNextIDScarico(strKNumViaggio As String) As Long

    Dim varBuffer As Variant
    Dim lngHowMany As Long

    If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_NEXTID_SCARICO, varBuffer, lngHowMany, strKNumViaggio) Then
        KNOSOS_GetNextIDScarico = 0
        Exit Function
    End If

    If IsNull(varBuffer(0, 0)) Then
        KNOSOS_GetNextIDScarico = 1
    Else
        KNOSOS_GetNextIDScarico = varBuffer(0, 0) + 1
    End If

End Function
Private Function GetDBServerDateTimeGMT0() As String

    Dim varBuffer As Variant
    Dim lngHowMany As Long

    If Not DS.DA.QueryCommand("SELECT CONVERT( nvarchar, GETUTCDATE(), 21 ) ", varBuffer, lngHowMany) Then
        GetDBServerDateTimeGMT0 = ""
        Exit Function
    End If
        
    GetDBServerDateTimeGMT0 = varBuffer(0, 0)

End Function
Private Function KNOSOS_InsertViaggio(objSatError As CError, _
                                varTargaAutomezzo As Variant, _
                                varIDViaggio As Variant, _
                                varNumViaggio As Variant, _
                                varStato As Variant)

    Dim SQLStatement As String
    Dim strInsertDateTime As String

    strInsertDateTime = GetDBServerDateTimeGMT0()
    SQLStatement = Replace(SQL_KNOSOS_INSERT_KVIAGGIO, "%fecha%", strInsertDateTime)
    If Not DS.DA.ExecuteCommand(SQLStatement, _
                            varTargaAutomezzo, _
                            varIDViaggio, _
                            varNumViaggio, _
                            varStato) Then
        GoTo ErrHandler
    End If

    KNOSOS_InsertViaggio = True
    Exit Function

ErrHandler:
    KNOSOS_InsertViaggio = False
End Function
Private Sub KNOSOS_AdjustCaricoData(varDeposito As Variant, _
                                    varDataOraCarico As Variant)

    varDeposito = Left(varDeposito, 11)

End Sub
Private Function KNOSOS_InsertCarico(objSatError As CError, _
                                varTargaAutomezzo As Variant, _
                                varIDViaggio As Variant, _
                                varNumViaggio As Variant, _
                                varStato As Variant, _
                                varDeposito As Variant, _
                                varDataOraCarico As Variant)

    Dim SQLStatement As String
    Dim strInsertDateTime As String

    strInsertDateTime = GetDBServerDateTimeGMT0()
    KNOSOS_AdjustCaricoData varDeposito, _
                            varDataOraCarico
    
    SQLStatement = Replace(SQL_KNOSOS_INSERT_KCARICO, "%fecha%", strInsertDateTime)
    If Not DS.DA.ExecuteCommand(SQLStatement, _
                                varTargaAutomezzo, _
                                varIDViaggio, _
                                varNumViaggio, _
                                KNOSOS_IDSCARICO_CARICO, _
                                KNOSOS_TIPO_CARICO, _
                                varStato, _
                                varDeposito, _
                                varDataOraCarico) Then
        GoTo ErrHandler
    End If

    KNOSOS_InsertCarico = True
    Exit Function

ErrHandler:
    
    KNOSOS_InsertCarico = False
End Function
Private Function SMS_Insert(varIDViaggio As Variant, _
                            varIDAutista As Variant, _
                            Optional ForceInsert As Boolean = False) As Boolean

    Dim SQLStatement As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim strSQL4 As String
    Dim strSQL5 As String
    Dim strSQL6 As String
    Dim strSQL7 As String
    Dim strSQL8 As String
    Dim strSQL9 As String
    Dim strSQL10 As String
    Dim varBuffer As Variant
    Dim varBufferCheck As Variant
    Dim varProgr As Variant
    Dim varIDVVC As Variant
    Dim lngHowMany As Long
    Dim i As Long
    
    strSQL1 = "DELETE FROM viagsvg0 " & _
              " WHERE svgidvvt = ? "
    
    strSQL2 = "DELETE FROM viagsvc0 " & _
              " WHERE svcidvvt = ? "
              
    strSQL3 = "SELECT vsaprogr " & _
              "  FROM viagvsa0 " & _
              " WHERE vsaidaau = ? "
              
    strSQL4 = "INSERT INTO viagvsa0 " & _
              "       (vsaidaau, vsaprogr) " & _
              " VALUES ( ?, '000')"
    
    strSQL5 = "SELECT vvcidvvt, vvcidvvc " & _
              "  FROM viagvvc0 " & _
              " WHERE vvcidvvt = ? and vvcannul = 0 " & _
              " ORDER BY vvcnrseq "
    
    strSQL6 = "INSERT INTO viagsvg0 " & _
              "       (svgidvvt, svgprgme) " & _
              " VALUES (?, ? )"
    
    strSQL7 = "INSERT INTO viagsvc0 " & _
              "       (svcidvvt, svcidvvc, svcprgme) " & _
              " VALUES (?, ?, ? )"
    
    strSQL8 = "UPDATE viagvsa0 " & _
              "   SET vsaprogr = ? " & _
              " WHERE vsaidaau = ? "
    
    strSQL9 = "SELECT COUNT(*) FROM viagsvg0 " & _
              " WHERE svgidvvt = ? "
    
    strSQL10 = "SELECT COUNT(*) FROM viagsvc0 " & _
              " WHERE svcidvvt = ? AND svcidvvc = ? "
              
              
    DS.DA.BeginTransaction
              
    ' Se richiesto eliminiamo eventuali numeratore precedentemente inseriti per il viaggio
    If ForceInsert Then
        SQLStatement = strSQL1
        If Not DS.DA.ExecuteCommand(SQLStatement, varIDViaggio) Then
            GoTo ErrHandler
        End If
        SQLStatement = strSQL2
        If Not DS.DA.ExecuteCommand(SQLStatement, varIDViaggio) Then
            GoTo ErrHandler
        End If
    End If
              
    ' Cerchiamo il numeratore corrispondente all'autista
    SQLStatement = strSQL3
    If Not DS.DA.QueryCommand(SQLStatement, varBuffer, lngHowMany, _
                                varIDAutista) Then
        GoTo ErrHandler
    End If
    
    ' Se non troviamo il numeratore per l'autista, lo inseriamo
    If IsEmpty(varBuffer) Then
        SQLStatement = strSQL4
        If Not DS.DA.ExecuteCommand(SQLStatement, varIDAutista) Then
            GoTo ErrHandler
        End If
        varProgr = "000"
    Else
        varProgr = varBuffer(0, 0)
    End If
    
    If Not IsNumeric(varProgr) Then
        GoTo ErrHandler
    End If
    If CLng(varProgr) = "99" Then
        varProgr = "000"
    End If
    
    varProgr = Format(CLng(varProgr) + 1, "000")
    
    ' Selezioniamo l'elenco delle consegne del viaggio
    lngHowMany = 0
    SQLStatement = strSQL5
    If Not DS.DA.QueryCommand(SQLStatement, varBuffer, lngHowMany, _
                              varIDViaggio) Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    
    ' Inseriamo il progressivo per il carico (se non Ë gi‡ presente)
    SQLStatement = strSQL9
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQLStatement, varBufferCheck, lngHowMany, varIDViaggio) Then
        GoTo ErrHandler
    End If
    If varBufferCheck(0, 0) = 0 Then
        SQLStatement = strSQL6
        If Not DS.DA.ExecuteCommand(SQLStatement, varIDViaggio, CLng(varProgr)) Then
            GoTo ErrHandler
        End If
    End If
    
    ' Inseriamo il progressivo per ogni consegna (se non Ë gi‡ presente)
    For i = 0 To UBound(varBuffer, 2)
        If CLng(varProgr) = "99" Then
            varProgr = "000"
        End If
        varProgr = Format(CLng(varProgr) + 1, "000")
        varIDVVC = varBuffer(1, i)
        
        SQLStatement = strSQL10
        lngHowMany = 0
        If Not DS.DA.QueryCommand(SQLStatement, varBufferCheck, lngHowMany, varIDViaggio, varIDVVC) Then
            GoTo ErrHandler
        End If
        If varBufferCheck(0, 0) = 0 Then
            SQLStatement = strSQL7
            If Not DS.DA.ExecuteCommand(SQLStatement, varIDViaggio, varIDVVC, CLng(varProgr)) Then
                GoTo ErrHandler
            End If
        End If
    Next i
    
    ' Aggiorniamo il progressivo sulla tabella dei progressivi per autista
    SQLStatement = strSQL8
    If Not DS.DA.ExecuteCommand(SQLStatement, varProgr, varIDAutista) Then
        GoTo ErrHandler
    End If
    
    DS.DA.EndTrans etCommit
    SMS_Insert = True
    
    Exit Function

ErrHandler:
    
    DS.DA.EndTrans etRollback
    SMS_Insert = False
End Function

Private Function KNOSOS_UpdateCarico(objSatError As CError, _
                                varNumViaggio As Variant, _
                                varDeposito As Variant, _
                                varDataOraCarico As Variant)

    Dim varKBuffer As Variant
    Dim lngHowMany As Long
    
    KNOSOS_AdjustCaricoData varDeposito, _
                            varDataOraCarico

    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_KCARICO, varKBuffer, lngHowMany, _
                                varNumViaggio, _
                                KNOSOS_TIPO_CARICO, _
                                KNOSOS_IDSCARICO_CARICO) Then
        GoTo ErrHandler
    End If

    If lngHowMany = 0 Then
        objSatError.Info = "Carico KNOSOS non trovato!"
        GoTo ErrHandler
    End If

    ' Controlla se i campi sono stati variati
    If varDeposito <> varKBuffer(5, 0) Or _
        varDataOraCarico <> varKBuffer(6, 0) Then

        ' Controlla che lo stato sia 1:Creada o 3:Modificada
'        If varKBuffer(4, 0) <> 1 And varKBuffer(4, 0) <> 3 Then
'            objSatError.Info = "Lo stato del carico KNOSOS non permette delle modifiche!"
'            GoTo ErrHandler
'        End If
        
        If Not DS.DA.ExecuteCommand(SQL_KNOSOS_UPDATE_KCARICO, _
                                    KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_MODIFICA, _
                                    varDeposito, _
                                    varDataOraCarico, _
                                    varNumViaggio, _
                                    KNOSOS_TIPO_CARICO, _
                                    KNOSOS_IDSCARICO_CARICO) Then
            GoTo ErrHandler
        End If

'        If Not KNOSOS_InsertCarico(objSatError, _
'                                varKBuffer(0, 0), _
'                                varKBuffer(1, 0), _
'                                varKBuffer(2, 0), _
'                                ) Then
'            GoTo ErrHandler
'        End If
    End If
    
    KNOSOS_UpdateCarico = True
    Exit Function

ErrHandler:
    
    KNOSOS_UpdateCarico = False

End Function
Private Sub KNOSOS_AdjustScaricoData(varNomeDestinatario As Variant, _
                                varLocalitaDestinatario As Variant, _
                                varIndirizzoDestinatario As Variant, _
                                varTelefonoDestinatario As Variant, _
                                varOrarioRichiesto As Variant, _
                                varOrarioApertura As Variant, _
                                varOraArrivo As Variant, _
                                varCodiceProdotto As Variant, _
                                varCodiceProdottoAbbreviato As Variant, _
                                varLitri As Variant, _
                                varCodiceOrdine As Variant, _
                                varNote1 As Variant, _
                                varNote2 As Variant)

    Dim TelefonoDestinatario As Variant
    Dim PositionCounter As Long

    varNomeDestinatario = Left(varNomeDestinatario, 26)
    varLocalitaDestinatario = Left(varLocalitaDestinatario, 35)
    varIndirizzoDestinatario = Left(varIndirizzoDestinatario, 35)
    If Not IsNull(varTelefonoDestinatario) Then
        TelefonoDestinatario = Null
        For PositionCounter = 1 To Len(varTelefonoDestinatario)
            If Mid(varTelefonoDestinatario, PositionCounter, 1) >= "0" And _
                Mid(varTelefonoDestinatario, PositionCounter, 1) <= "9" Then
                TelefonoDestinatario = TelefonoDestinatario & Mid(varTelefonoDestinatario, PositionCounter, 1)
            End If
        Next PositionCounter
        varTelefonoDestinatario = TelefonoDestinatario
    End If
    varTelefonoDestinatario = Left(varTelefonoDestinatario, 15)
    varOrarioRichiesto = Left(varOrarioRichiesto, 30)
    varOrarioApertura = Left(varOrarioApertura, 15)
    varOraArrivo = Left(varOraArrivo, 15)
    varCodiceProdotto = Left(varCodiceProdotto, 10)
    varCodiceProdottoAbbreviato = Left(varCodiceProdottoAbbreviato, 3)
    varLitri = Trim(CStr(varLitri))
    varCodiceOrdine = Left(varCodiceOrdine, 30)
    varNote1 = Left(varNote1, 33)
    varNote2 = Left(varNote2, 33)
End Sub

Private Function KNOSOS_InsertScarico(objSatError As CError, _
                                        varTargaAutomezzo As Variant, _
                                        varIDViaggio As Variant, _
                                        varNumViaggio As Variant, _
                                        varIdScarico As Variant, _
                                        varStato As Variant, _
                                        varNomeDestinatario As Variant, _
                                        varLocalitaDestinatario As Variant, _
                                        varIndirizzoDestinatario As Variant, _
                                        varTelefonoDestinatario As Variant, _
                                        varOrarioRichiesto As Variant, _
                                        varOrarioApertura As Variant, _
                                        varOraArrivo As Variant, _
                                        varCodiceProdotto As Variant, _
                                        varCodiceProdottoAbbreviato As Variant, _
                                        varLitri As Variant, _
                                        varCodiceOrdine As Variant, _
                                        varNote1 As Variant, _
                                        varNote2 As Variant)

    Dim SQLStatement As String
    Dim strInsertDateTime As String

    strInsertDateTime = GetDBServerDateTimeGMT0()

    ' Adjust data
    KNOSOS_AdjustScaricoData varNomeDestinatario, _
                                varLocalitaDestinatario, _
                                varIndirizzoDestinatario, _
                                varTelefonoDestinatario, _
                                varOrarioRichiesto, _
                                varOrarioApertura, _
                                varOraArrivo, _
                                varCodiceProdotto, _
                                varCodiceProdottoAbbreviato, _
                                varLitri, _
                                varCodiceOrdine, _
                                varNote1, _
                                varNote2
    
    SQLStatement = Replace(SQL_KNOSOS_INSERT_KSCARICO, "%fecha%", strInsertDateTime)
    If Not DS.DA.ExecuteCommand(SQLStatement, _
                                varTargaAutomezzo, _
                                varIDViaggio, _
                                varNumViaggio, _
                                varIdScarico, _
                                KNOSOS_TIPO_SCARICO, _
                                varStato, _
                                varNomeDestinatario, _
                                varLocalitaDestinatario, _
                                varIndirizzoDestinatario, _
                                varTelefonoDestinatario, _
                                varOrarioRichiesto, _
                                varOrarioApertura, _
                                varOraArrivo, _
                                varCodiceProdotto, _
                                varCodiceProdottoAbbreviato, _
                                varLitri, _
                                varCodiceOrdine, _
                                varNote1, _
                                varNote2) Then
        GoTo ErrHandler
    End If
    
    KNOSOS_InsertScarico = True
    Exit Function

ErrHandler:
    
    KNOSOS_InsertScarico = False
End Function
Private Function KNOSOS_DeleteScarico(objSatError As CError, _
                                    varNumViaggio As Variant, _
                                    varIdScarico As Variant, _
                                    blnIfAnnullatoExit As Boolean)

    Dim varKBuffer As Variant
    Dim lngHowMany As Long
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_KSCARICO, varKBuffer, lngHowMany, _
                                varNumViaggio, _
                                KNOSOS_TIPO_SCARICO, _
                                varIdScarico) Then
        GoTo ErrHandler
    End If

    If lngHowMany = 0 Then
        objSatError.Info = "Scarico KNOSOS non trovato!"
        GoTo ErrHandler
    End If

    If blnIfAnnullatoExit = True And (varKBuffer(5, 0) = 4 Or varKBuffer(5, 0) = 5) Then
        KNOSOS_DeleteScarico = True
        Exit Function
    End If
    
    ' Controlla che lo stato sia 1:Creada o 3:Modificada
    If varKBuffer(5, 0) <> 1 And varKBuffer(5, 0) <> 3 Then
        objSatError.Info = "Lo stato dello scarico KNOSOS non permette delle modifiche!"
        GoTo ErrHandler
    End If

    If Not DS.DA.ExecuteCommand(SQL_KNOSOS_UPDATESTATO_KSCARICO, _
                                KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_ANNULLAMENTO, _
                                varNumViaggio, _
                                KNOSOS_TIPO_SCARICO, _
                                varIdScarico) Then
        GoTo ErrHandler
    End If

'    If Not KNOSOS_InsertScarico(objSatError, _
'                                varKBuffer(0, 0), _
'                                varKBuffer(1, 0), _
'                                varKBuffer(2, 0), _
'                                varKBuffer(3, 0), _
'                                KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_ANNULLAMENTO, _
'                                varKBuffer(6, 0), _
'                                varKBuffer(7, 0), _
'                                varKBuffer(8, 0), _
'                                varKBuffer(9, 0), _
'                                varKBuffer(10, 0), _
'                                varKBuffer(11, 0), _
'                                varKBuffer(12, 0), _
'                                varKBuffer(13, 0), _
'                                varKBuffer(14, 0), _
'                                varKBuffer(15, 0), _
'                                varKBuffer(16, 0), _
'                                varKBuffer(17, 0), _
'                                varKBuffer(18, 0)) Then
'        GoTo ErrHandler
'    End If

    KNOSOS_DeleteScarico = True
    Exit Function

ErrHandler:
    
    KNOSOS_DeleteScarico = False
End Function
Private Function KNOSOS_UpdateScarico(objSatError As CError, _
                                    varNumViaggio As Variant, _
                                    varIdScarico As Variant, _
                                    varNomeDestinatario As Variant, _
                                    varLocalitaDestinatario As Variant, _
                                    varIndirizzoDestinatario As Variant, _
                                    varTelefonoDestinatario As Variant, _
                                    varOrarioRichiesto As Variant, _
                                    varOrarioApertura As Variant, _
                                    varOraArrivo As Variant, _
                                    varCodiceProdotto As Variant, _
                                    varCodiceProdottoAbbreviato As Variant, _
                                    varLitri As Variant, _
                                    varCodiceOrdine As Variant, _
                                    varNote1 As Variant, _
                                    varNote2 As Variant)

    Dim varKBuffer As Variant
    Dim lngHowMany As Long
    Dim lngCodiceOrdine As Long
    
    ' Adjust data
    KNOSOS_AdjustScaricoData varNomeDestinatario, _
                                varLocalitaDestinatario, _
                                varIndirizzoDestinatario, _
                                varTelefonoDestinatario, _
                                varOrarioRichiesto, _
                                varOrarioApertura, _
                                varOraArrivo, _
                                varCodiceProdotto, _
                                varCodiceProdottoAbbreviato, _
                                varLitri, _
                                varCodiceOrdine, _
                                varNote1, _
                                varNote2

    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_KSCARICO, varKBuffer, lngHowMany, _
                                varNumViaggio, _
                                KNOSOS_TIPO_SCARICO, _
                                varIdScarico) Then
        GoTo ErrHandler
    End If

    If lngHowMany = 0 Then
        objSatError.Info = "Scarico KNOSOS non trovato!"
        GoTo ErrHandler
    End If

    If varKBuffer(16, 0) = "" Then
        lngCodiceOrdine = 0
    Else
        lngCodiceOrdine = CLng(varKBuffer(16, 0))
    End If
    
    ' Controlla se i campi sono stati variati
    If varNomeDestinatario <> varKBuffer(6, 0) Or _
        varLocalitaDestinatario <> varKBuffer(7, 0) Or _
        varIndirizzoDestinatario <> varKBuffer(8, 0) Or _
        varTelefonoDestinatario <> varKBuffer(9, 0) Or _
        varOrarioRichiesto <> varKBuffer(10, 0) Or _
        varOrarioApertura <> varKBuffer(11, 0) Or _
        varOraArrivo <> varKBuffer(12, 0) Or _
        varCodiceProdotto <> varKBuffer(13, 0) Or _
        varCodiceProdottoAbbreviato <> varKBuffer(14, 0) Or _
        varLitri <> varKBuffer(15, 0) Or _
        varCodiceOrdine <> lngCodiceOrdine Or _
        varNote1 <> varKBuffer(17, 0) Or _
        varNote2 <> varKBuffer(18, 0) Then

        ' Controlla che lo stato sia 1:Creada o 3:Modificada
        If varKBuffer(5, 0) <> 1 And varKBuffer(5, 0) <> 3 Then
            objSatError.Info = "Lo stato dello scarico KNOSOS non permette delle modifiche!"
            GoTo ErrHandler
        End If

        If Not DS.DA.ExecuteCommand(SQL_KNOSOS_UPDATE_KSCARICO, _
                                    KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_MODIFICA, _
                                    varNomeDestinatario, _
                                    varLocalitaDestinatario, _
                                    varIndirizzoDestinatario, _
                                    varTelefonoDestinatario, _
                                    varOrarioRichiesto, _
                                    varOrarioApertura, _
                                    varOraArrivo, _
                                    varCodiceProdotto, _
                                    varCodiceProdottoAbbreviato, _
                                    varLitri, _
                                    varCodiceOrdine, _
                                    varNote1, _
                                    varNote2, _
                                    varNumViaggio, _
                                    KNOSOS_TIPO_SCARICO, _
                                    varIdScarico) Then
            GoTo ErrHandler
        End If

'        If Not KNOSOS_InsertScarico(objSatError, _
'                                        varKBuffer(0, 0), _
'                                        varKBuffer(1, 0), _
'                                        varKBuffer(2, 0), _
'                                        varKBuffer(3, 0), _
'                                        KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_MODIFICA, _
'                                        varNomeDestinatario, _
'                                        varLocalitaDestinatario, _
'                                        varIndirizzoDestinatario, _
'                                        varTelefonoDestinatario, _
'                                        varOrarioRichiesto, _
'                                        varOrarioApertura, _
'                                        varOraArrivo, _
'                                        varCodiceProdotto, _
'                                        varCodiceProdottoAbbreviato, _
'                                        varLitri, _
'                                        varCodiceOrdine, _
'                                        varNote1, _
'                                        varNote2) Then
'            GoTo ErrHandler
'        End If
        
    End If
    
    KNOSOS_UpdateScarico = True
    Exit Function

ErrHandler:
    
    KNOSOS_UpdateScarico = False

End Function
Private Function KNOSOS_GetFilterScaricoAnnullato(blnDeleted As Boolean, strOperator As String) As String
    
    Dim AVC As CViaggioConsegna
    Dim AVP As CViaggioProdotto
    
    ' Prelevo il nome del campo annullato del viaggio carico
    Set AVC = New CViaggioConsegna
    Set AVP = New CViaggioProdotto
    
    KNOSOS_GetFilterScaricoAnnullato = ""
    
    If AVC.DS.CancelField <> "" Then
        KNOSOS_GetFilterScaricoAnnullato = AVC.DS.CancelField & " = " & DS.DA.MakeSQL(IIf(blnDeleted, True, False), dtBit)
    End If
    
    If AVP.DS.CancelField <> "" Then
        If KNOSOS_GetFilterScaricoAnnullato <> "" Then
            KNOSOS_GetFilterScaricoAnnullato = KNOSOS_GetFilterScaricoAnnullato & " " & strOperator & " "
        End If
        
        KNOSOS_GetFilterScaricoAnnullato = KNOSOS_GetFilterScaricoAnnullato & AVP.DS.CancelField & " = " & DS.DA.MakeSQL(IIf(blnDeleted, True, False), dtBit)
    End If
    
    Set AVC = Nothing
    Set AVP = Nothing
End Function
Private Function isValidNumber(strNumber As String)
    Dim lngCounter As Long
    For lngCounter = 1 To Len(strNumber)
        If InStr("0123456789", Mid(strNumber, lngCounter, 1)) = 0 Then
            isValidNumber = False
            Exit Function
        End If
    Next lngCounter
    isValidNumber = True
End Function
' ================================================================================
' Insert_Sat
' --------------------------------------------------------------------------------
' Inserisce il viaggio, il carico e gli scarichi nelle tabelle di interfaccia
' di KNOSOS.
' Viene chiamata al momento del primo passaggio da provissorio a definitivo.
'
' Prerequisiti per l'inserimento:
' - Il campo AAMFLSAT nella tabella degli automezzi deve valere 'K', ovvero dotato
'   di terminale satellitare KNOSOS
' - Lo stato del viaggio deve passare da Provvisorio a Definitivo
' - Deve essere il primo passaggio da provvisorio a Definitivo, il campo VVTNRKVI
'   ("47") deve essere NULL
' ================================================================================
Public Function Insert_Sat(objSatError As CError) As Boolean
On Error GoTo ErrHandler
    Dim AAM As CAutomezzo
    Dim ADE As CDeposito

    Dim strKIdViaggio As String
    Dim strKNumViaggio As String
    Dim lngKIdScarico As Long

    Dim strIdAutomezzo As String
    Dim strIDDeposito As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim SQLStatement As String
    Dim strFilter As String

    Dim blnSatellitareInstallato As String
    Dim strDataOraCarico As String
    Dim strDeposito As String
    
    strIdAutomezzo = DS.Fields("15").Value  ' Targa
    
    ' Controlla che lo stato sia Definitivo
    If Not DS.Fields("42").Value = "D" Then
        objSatError.Info = "Impossibile inserire un viaggio KNOSOS non definitivo!"
        GoTo ErrHandler
    End If

    ' ------------------------------------------------------------------------
    ' Carica i dati dell'automezzo
    ' ------------------------------------------------------------------------
    ' Ricerca l'automezzo e carica i dati
    Set AAM = New CAutomezzo
    With AAM
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH

        .DS.Fields("0").FindValue = strIdAutomezzo

        If .Find = True And .DS.HowMany > 0 Then
            
            ' Imposta SatellitareInstallato a True o False
            blnSatellitareInstallato = IIf(.DS.Fields("34").Value = "K", True, False)
        Else
            objSatError.Info = "Automezzo non trovato!"
            GoTo ErrHandler
        End If
    End With
    
    ' ------------------------------------------------------------------------
    ' Controlla se e' possibile continuare
    ' ------------------------------------------------------------------------
    
    ' Se il satellitare non e' installato esce senza errore
    If blnSatellitareInstallato = False Then
        Insert_Sat = True
        Exit Function
    End If
    
    ' Se il campo VVTNRKVI non e' NULL esce, con errore
    If DS.Fields("47").Value <> "" Then
        objSatError.Info = "Impossibile inserire il viaggio, l'id di riferimento KNOSOS non e' vuoto!"
        GoTo ErrHandler
    End If
    
    ' ------------------------------------------------------------------------
    ' Carica i dati del deposito
    ' ------------------------------------------------------------------------
    ' Ricerca il deposito e carica i dati
    Set ADE = New CDeposito
    With ADE
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH

        strIDDeposito = DS.Fields("22").Value
        .DS.Fields("0").FindValue = strIDDeposito

        If .Find = True Then
            strDeposito = .DS.Fields("1").Value
        Else
            objSatError.Info = "Deposito non trovato!"
            GoTo ErrHandler
        End If
    End With
    
    ' Inizia transazione
    DS.DA.BeginTransaction
    
    ' ------------------------------------------------------------------------
    ' Inserimento del viaggio con stato creazione pendente
    ' ------------------------------------------------------------------------
    strKIdViaggio = DS.Fields("2").Value    ' Sequenza
    strKNumViaggio = Format(DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "900", 0), "000000")
    If Not KNOSOS_InsertViaggio(objSatError, _
                                strIdAutomezzo, _
                                strKIdViaggio, _
                                strKNumViaggio, _
                                KNOSOS_VIAGGIO_STATO_PENDENTE) Then
        GoTo ErrHandler
    End If

    ' ------------------------------------------------------------------------
    ' Inserimento del carico
    ' ------------------------------------------------------------------------
    strDataOraCarico = Format(DS.Fields("4").Value, "dd/mm/yy") & " " & Format(DS.Fields("5").Value, "hh:mm")
    If Not KNOSOS_InsertCarico(objSatError, _
                                strIdAutomezzo, _
                                strKIdViaggio, _
                                strKNumViaggio, _
                                KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_CREAZIONE, _
                                strDeposito, _
                                strDataOraCarico) Then
        GoTo ErrHandler
    End If

    ' ------------------------------------------------------------------------
    ' Inserimento di tutti gli scarichi
    ' ------------------------------------------------------------------------
    
    ' Preleva la lista degli scarichi
    SQLStatement = SQL_KNOSOS_SELECT_SCARICHI_RESOIL
    strFilter = KNOSOS_GetFilterScaricoAnnullato(False, "AND")
    If strFilter <> "" Then
        SQLStatement = SQLStatement + " AND (" & strFilter & ")"
    End If
    SQLStatement = SQLStatement + " ORDER BY VVCNRSEQ, VVPIDVVP"
    
    If Not DS.DA.QueryCommand(SQLStatement, varBuffer, lngHowMany, DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then

        For i = 0 To UBound(varBuffer, 2)
            
            ' Acquisisce il nuovo ID descarga
            lngKIdScarico = KNOSOS_GetNextIDScarico(strKNumViaggio)
            
            If lngKIdScarico = 0 Then
                objSatError.Info = "Impossibile prelevare il nuovo numero di scarico!"
                GoTo ErrHandler
            End If

            If Not KNOSOS_InsertScarico(objSatError, _
                                        strIdAutomezzo, _
                                        strKIdViaggio, _
                                        strKNumViaggio, _
                                        lngKIdScarico, _
                                        KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_CREAZIONE, _
                                        varBuffer(4, i), _
                                        varBuffer(5, i), _
                                        varBuffer(6, i), _
                                        varBuffer(7, i), _
                                        varBuffer(8, i), _
                                        Format(varBuffer(9, i), "hh:mm") & " - " & Format(varBuffer(10, i), "hh:mm"), _
                                        Format(varBuffer(11, i), "hh:mm"), _
                                        varBuffer(3, i), _
                                        varBuffer(15, i), _
                                        varBuffer(12, i), _
                                        varBuffer(13, i), _
                                        vbNullString, _
                                        vbNullString) Then
                GoTo ErrHandler
            End If

            ' Aggiorna l'id della consegna prodotto di resoil
            If Not DS.DA.ExecuteCommand(SQL_KNOSOS_UPDATE_NUMERO_SCARICO, _
                                        lngKIdScarico, _
                                        CStr(varBuffer(0, i)), _
                                        CStr(varBuffer(1, i)), _
                                        CStr(varBuffer(2, i))) Then
                GoTo ErrHandler
            End If

        Next i
    End If

    ' Aggiorna il numero viaggio KNOSOS e lo Stato viaggio a Pendente
    If Not DS.DA.ExecuteCommand(SQL_UPDATE_KNOSOS_DATI, _
                                strKNumViaggio, _
                                "0", _
                                DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If

    ' Commit della transazione
    DS.DA.EndTrans etCommit
    
    Set AAM = Nothing
    Set ADE = Nothing
    
    Insert_Sat = True
    Exit Function

ErrHandler:
    Set AAM = Nothing
    Set ADE = Nothing
    
    Insert_Sat = False
    DS.DA.EndTrans etRollback
    
End Function
' ================================================================================
' Update_Sat
' --------------------------------------------------------------------------------
' Aggiorna, se i dati sono variati, il carico e gli scarichi nelle tabelle di
' interfaccia per KNOSOS.
' Viene chiamata ai successivi passaggi da provvisorio a definitivo.
'
' Prerequisiti per la variazione:
' - Il campo AAMFLSAT nella tabella degli automezzi deve valere 'K', ovvero dotato
'   di ricevitore satellitare KNOSOS
' - Lo stato del viaggio deve essere Definitivo
' - Lo stato del viaggio KNOSOS deve essere 0=Pendiente, 1=enviado, 2=recibido
' - Non deve essere il primo passaggio da Provvisorio a Definitivo, il campo
'   VVTNRKVI ("47") non deve essere NULL
' ================================================================================
Public Function Update_Sat(objSatError As CError) As Boolean

    Dim AAM As CAutomezzo
    Dim ADE As CDeposito

    Dim strKIdViaggio As String
    Dim strKNumViaggio As String
    Dim lngKIdScarico As Long

    Dim strIdAutomezzo As String
    Dim strIDDeposito As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim SQLStatement As String
    Dim strFilter As String

    Dim blnSatellitareInstallato As String
    Dim strDataOraCarico As String
    Dim strDeposito As String
    
    Dim varROScarichiBuffer As Variant
    Dim lngROScarichiIdCount As Long
    Dim lngROAllScarichiId() As Long
    Dim lngROAllScarichiCount As Long

    ' Controlla se lo stato e' Definitivo
    If Not DS.Fields("42").Value = "D" Then
        objSatError.Info = "Impossibile aggiornare un viaggio KNOSOS non definitivo!"
        GoTo ErrHandler
    End If

    ' Inizializzazione variabili
    strIdAutomezzo = DS.Fields("15").Value
    strKNumViaggio = DS.Fields("47").Value
    strKIdViaggio = DS.Fields("2").Value  ' Sequenza
    
    ' ------------------------------------------------------------------------
    ' Carica i dati dell'automezzo
    ' ------------------------------------------------------------------------
    Set AAM = New CAutomezzo
    With AAM
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH

        .DS.Fields("0").FindValue = strIdAutomezzo

        If .Find = True And .DS.HowMany > 0 Then
            
            ' Imposta SatellitareInstallato a True o False
            blnSatellitareInstallato = IIf(.DS.Fields("34").Value = "K", True, False)
        Else
            objSatError.Info = "Automezzo non trovato!"
            GoTo ErrHandler
        End If
    End With

    ' ------------------------------------------------------------------------
    ' Controlli se e' possibile procedere
    ' ------------------------------------------------------------------------
    
    ' Se il satellitare non e' installato esce senza errore
    If blnSatellitareInstallato = False Then
        Update_Sat = True
        Exit Function
    End If
    
    ' Se il campo VVTNRKVI e' NULL esce, con errore
    If DS.Fields("47").Value = "" Then
        objSatError.Info = "Impossibile aggiornare il viaggio, manca l'id di riferimento KNOSOS!"
        GoTo ErrHandler
    End If

    ' Controlla se e' possibile modificare il viaggio
    ' Check dello stato del KVIAGGIO
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_KVIAGGIO, varBuffer, lngHowMany, DS.Fields("47").Value) Then
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then
'        If varBuffer(4, 0) <> KNOSOS_VIAGGIO_STATO_PENDENTE And _
'            varBuffer(4, 0) <> KNOSOS_VIAGGIO_STATO_INVIATO And _
'            varBuffer(4, 0) <> KNOSOS_VIAGGIO_STATO_RICEVUTO Then
'
'            objSatError.Info = "Il viaggio KNOSOS risulta in corso, impossibile modificare il viaggio!"
'            GoTo ErrHandler
'        End If
    Else
        objSatError.Info = "Viaggio KNOSOS non trovato (id = " & DS.Fields("47").Value & ")!"
        GoTo ErrHandler
    End If

    ' ------------------------------------------------------------------------
    ' Carica i dati del deposito
    ' ------------------------------------------------------------------------
    ' Ricerca il deposito e carica i dati
    Set ADE = New CDeposito
    With ADE
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH

        strIDDeposito = DS.Fields("22").Value
        .DS.Fields("0").FindValue = strIDDeposito

        If .Find = True Then
            strDeposito = .DS.Fields("1").Value
        Else
            objSatError.Info = "Deposito non trovato!"
            GoTo ErrHandler
        End If
    End With
    
    ' Inizia transazione
    DS.DA.BeginTransaction
    
    ' ------------------------------------------------------------------------
    ' Update del carico, solo se modificato e solo se il viaggio non e' in corso
    ' ------------------------------------------------------------------------
    If varBuffer(4, 0) < KNOSOS_VIAGGIO_STATO_INCORSO Then
        strDataOraCarico = Format(DS.Fields("4").Value, "dd/mm/yy") & " " & Format(DS.Fields("5").Value, "hh:mm")
        If Not KNOSOS_UpdateCarico(objSatError, _
                                    strKNumViaggio, _
                                    strDeposito, _
                                    strDataOraCarico) Then
            GoTo ErrHandler
        End If
    End If
    
    ' ------------------------------------------------------------------------
    ' Controlla se sono stati modificati i dati degli scarichi
    ' ------------------------------------------------------------------------
    SQLStatement = SQL_KNOSOS_SELECT_SCARICHI_RESOIL
    strFilter = KNOSOS_GetFilterScaricoAnnullato(False, "AND")
    If strFilter <> "" Then
        SQLStatement = SQLStatement + " AND (" & strFilter & ")"
    End If
    
    lngROScarichiIdCount = 0
    If Not DS.DA.QueryCommand(SQLStatement, varROScarichiBuffer, lngROScarichiIdCount, DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If

    lngROAllScarichiCount = 0
    If lngROScarichiIdCount > 0 Then

        For i = LBound(varROScarichiBuffer, 2) To UBound(varROScarichiBuffer, 2)
            
            ' Controlla se e' un nuovo scarico da inserire ex novo
            If IsNull(varROScarichiBuffer(14, i)) Or varROScarichiBuffer(14, i) = 0 Then

                ' Acquisisce l'id descarga
                lngKIdScarico = KNOSOS_GetNextIDScarico(strKNumViaggio)

                If lngKIdScarico = 0 Then
                    objSatError.Info = "Impossibile prelevare il nuovo numero di scarico!"
                    GoTo ErrHandler
                End If
                
                ' Inserisce l'id descarga nel vettore di tutti gli scarichi presenti in ResOil
                lngROAllScarichiCount = lngROAllScarichiCount + 1
                ReDim Preserve lngROAllScarichiId(1 To lngROAllScarichiCount)
                lngROAllScarichiId(lngROAllScarichiCount) = lngKIdScarico

                If Not KNOSOS_InsertScarico(objSatError, _
                                            strIdAutomezzo, _
                                            strKIdViaggio, _
                                            strKNumViaggio, _
                                            lngKIdScarico, _
                                            KNOSOS_OPEVIAGGIO_STATO_RICHIESTA_CREAZIONE, _
                                            varROScarichiBuffer(4, i), _
                                            varROScarichiBuffer(5, i), _
                                            varROScarichiBuffer(6, i), _
                                            varROScarichiBuffer(7, i), _
                                            varROScarichiBuffer(8, i), _
                                            Format(varROScarichiBuffer(9, i), "hh:mm") & " - " & Format(varROScarichiBuffer(10, i), "hh:mm"), _
                                            Format(varROScarichiBuffer(11, i), "hh:mm"), _
                                            varROScarichiBuffer(3, i), _
                                            varROScarichiBuffer(15, i), _
                                            varROScarichiBuffer(12, i), _
                                            varROScarichiBuffer(13, i), _
                                            vbNull, _
                                            vbNull) Then
                    GoTo ErrHandler
                End If

                ' Aggiorna l'id dello scarico di resoil
                If Not DS.DA.ExecuteCommand(SQL_KNOSOS_UPDATE_NUMERO_SCARICO, _
                                            lngKIdScarico, _
                                            CStr(varROScarichiBuffer(0, i)), _
                                            CStr(varROScarichiBuffer(1, i)), _
                                            CStr(varROScarichiBuffer(2, i))) Then
                    GoTo ErrHandler
                End If
            
            ' Si tratta di uno scarico gia' presente, controlla se bisogna aggiornarlo
            Else
                ' Inserisce l'id descarga nel vettore di tutti gli scarichi presenti in ResOil
                lngROAllScarichiCount = lngROAllScarichiCount + 1
                ReDim Preserve lngROAllScarichiId(1 To lngROAllScarichiCount) As Long
                lngROAllScarichiId(lngROAllScarichiCount) = varROScarichiBuffer(14, i)
                
                ' Se lo stato dello scarico e' Pendiente, ok, altrimenti skip
                If IsNull(varROScarichiBuffer(16, i)) Or varROScarichiBuffer(16, i) = "0" Then
                    
                    lngKIdScarico = varROScarichiBuffer(14, i)

                    If Not KNOSOS_UpdateScarico(objSatError, _
                                                strKNumViaggio, _
                                                lngKIdScarico, _
                                                varROScarichiBuffer(4, i), _
                                                varROScarichiBuffer(5, i), _
                                                varROScarichiBuffer(6, i), _
                                                varROScarichiBuffer(7, i), _
                                                varROScarichiBuffer(8, i), _
                                                Format(varROScarichiBuffer(9, i), "hh:mm") & " - " & Format(varROScarichiBuffer(10, i), "hh:mm"), _
                                                Format(varROScarichiBuffer(11, i), "hh:mm"), _
                                                varROScarichiBuffer(3, i), _
                                                varROScarichiBuffer(15, i), _
                                                varROScarichiBuffer(12, i), _
                                                varROScarichiBuffer(13, i), _
                                                vbNullString, _
                                                vbNullString) Then
                        GoTo ErrHandler
                    End If
                End If
            End If
        Next i
    End If

    ' Controlla se sono state annullate o cancellate delle consegne
    ' Preleva l'elenco di tutti gli id descarga dalla tabella di interfaccia
    lngHowMany = 0
    varBuffer = Empty
    
    
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_KSCARICHI_IDDESC, varBuffer, lngHowMany, strKNumViaggio) Then
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then
        
        Dim blnScaricoFound As Boolean
        Dim lngROScarichiIdCounter As Long
        
        ' Ricerca che tutti gli scarichi KNOSOS siano presenti in resoil, altrimenti inserisce
        ' un'annullamento
        For i = LBound(varBuffer, 2) To UBound(varBuffer, 2)
            
            blnScaricoFound = False
            
            For lngROScarichiIdCounter = 1 To lngROAllScarichiCount
                If lngROAllScarichiId(lngROScarichiIdCounter) = varBuffer(0, i) Then
                    blnScaricoFound = True
                    Exit For
                End If
            Next lngROScarichiIdCounter

            If Not blnScaricoFound Then
                If Not KNOSOS_DeleteScarico(objSatError, _
                                            strKNumViaggio, _
                                            varBuffer(0, i), _
                                            True) Then
                    GoTo ErrHandler
                End If
            End If
        Next i
    End If

    ' Se eventualmente ci sono degli annullamenti nella consegna o consegna prodotto,
    ' resetto l'id descarga
    lngHowMany = 0
    If Not DS.DA.ExecuteCommand(SQL_KNOSOS_RESET_IDDESC_ANNULLATE, DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If
    
    Set AAM = Nothing
    Set ADE = Nothing

    DS.DA.EndTrans etCommit

    Update_Sat = True
    Exit Function

ErrHandler:
    Set AAM = Nothing
    Set ADE = Nothing
    
    Update_Sat = False
    DS.DA.EndTrans etRollback
    
End Function
' ================================================================================
' Delete_Sat
' --------------------------------------------------------------------------------
' Annulla il viaggio nelle tabelle di interfaccia per KNOSOS.
' E' chiamata nel momento dell'annullamento del viaggio.
'
' Prerequisiti per l'annullamento:
' - Il campo AAMFLSAT nella tabella degli automezzi deve valere 'K', ovvero dotato
'   di ricevitore satellitare KNOSOS
' - Lo stato del viaggio KNOSOS deve essere 0=Pendiente, 1=enviado, 2=recibido
' - Non deve essere il primo inserimento, il campo VVTNRKVI ("47") non deve essere
'   NULL
' ================================================================================
Public Function Delete_Sat(objSatError As CError) As Boolean
    Dim AAM As CAutomezzo
    Dim ADE As CDeposito

    Dim strKIdViaggio As String
    Dim strKNumViaggio As String
    Dim lngKIdScarico As Long

    Dim strIdAutomezzo As String
    Dim strIDDeposito As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim SQLStatement As String
    
    Dim blnSatellitareInstallato As String
    Dim strDataOraCarico As String
    Dim strDeposito As String

    
    strIdAutomezzo = DS.Fields("15").Value
    
    ' ------------------------------------------------------------------------
    ' Carica i dati dell'automezzo
    ' ------------------------------------------------------------------------
    ' Ricerca l'automezzo e carica i dati
    Set AAM = New CAutomezzo
    With AAM
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH

        .DS.Fields("0").FindValue = strIdAutomezzo

        If .Find = True And .DS.HowMany > 0 Then
            ' Imposta SatellitareInstallato a True o False
            blnSatellitareInstallato = IIf(.DS.Fields("34").Value = "K", True, False)
        Else
            objSatError.Info = "Automezzo non trovato!"
            GoTo ErrHandler
        End If
    End With

    ' ------------------------------------------------------------------------
    ' Controlli se e' possibile procedere
    ' ------------------------------------------------------------------------
    
    ' Se il satellitare non e' installato esce senza errore
    If blnSatellitareInstallato = False Then
        Delete_Sat = True
        Exit Function
    End If
    
    ' Se il campo VVTNRKVI e' NULL esce, con errore
    If DS.Fields("47").Value = "" Then
        objSatError.Info = "Impossibile cancellare il viaggio, manca l'id di riferimento KNOSOS!"
        GoTo ErrHandler
    End If

    ' Controlla se e' possibile modificare il viaggio
    ' Check dello stato del KVIAGGIO
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_KVIAGGIO, varBuffer, lngHowMany, DS.Fields("47").Value) Then
        GoTo ErrHandler
    End If
    
    If lngHowMany > 0 Then
'        If varBuffer(4, 0) <> KNOSOS_VIAGGIO_STATO_PENDENTE And _
'            varBuffer(4, 0) <> KNOSOS_VIAGGIO_STATO_INVIATO And _
'            varBuffer(4, 0) <> KNOSOS_VIAGGIO_STATO_RICEVUTO Then
'
'            objSatError.Info = "Il viaggio KNOSOS risulta in corso, impossibile cancellare il viaggio!"
'            GoTo ErrHandler
'        End If
    Else
        objSatError.Info = "Viaggio KNOSOS non trovato!"
        GoTo ErrHandler
    End If

    ' Inizializzazione variabili
    strKIdViaggio = Format(DS.Fields("2").Value, "00")       ' Sequenza
    strKNumViaggio = DS.Fields("47").Value

    ' Inizia transazione
    DS.DA.BeginTransaction

    ' ------------------------------------------------------------------------
    ' Inserimento del viaggio con stato annullamento pendente
    ' ------------------------------------------------------------------------
    If Not KNOSOS_InsertViaggio(objSatError, _
                                strIdAutomezzo, _
                                strKIdViaggio, _
                                strKNumViaggio, _
                                KNOSOS_VIAGGIO_STATO_ANNUL_PENDENTE) Then
        GoTo ErrHandler
    End If

    Set AAM = Nothing
    Set ADE = Nothing

    DS.DA.EndTrans etCommit

    Delete_Sat = True
    Exit Function

ErrHandler:
    Set AAM = Nothing
    Set ADE = Nothing

    Delete_Sat = False
    DS.DA.EndTrans etRollback

End Function

Public Function SetAvanzamentoOrdini(strAvanzamentoOrdine As String) As Boolean
On Error GoTo ErrHandler
    
    Dim varBufferConsegne As Variant
    Dim lngHowMany As Long
    Dim lngCounter As Long
    Dim VVC As CViaggioConsegna
    Dim SQL_SelectConsegne As String
    
    SQL_SelectConsegne = "SELECT VVCIDVVC FROM VIAGVVC0 WHERE VVCIDVVT = ?"
    
    If Len(DS.Field("Numero Viaggio").Value) = 0 Then
        GoTo ErrHandler
    End If
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH
    End With
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SelectConsegne, varBufferConsegne, lngHowMany, DS.Field("Numero Viaggio").Value) Then
        GoTo ErrHandler
    End If
    
    For lngCounter = 0 To lngHowMany - 1
        VVC.DS.Fields.ResetFields
        VVC.DS.Field("Numero Viaggio").FindValue = DS.Field("Numero Viaggio").Value
        VVC.DS.Field("Numero Consegna").FindValue = varBufferConsegne(0, lngCounter)
        
        If Not VVC.Find() Or lngHowMany = 0 Then
            GoTo ErrHandler
        End If
        
        If Not VVC.SetAvanzamentoOrdine(strAvanzamentoOrdine) Then
            GoTo ErrHandler
        End If
    
    Next lngCounter
    
    Set VVC = Nothing
    SetAvanzamentoOrdini = True
    Exit Function

ErrHandler:
    
    Set VVC = Nothing
    SetAvanzamentoOrdini = False
End Function

'-- DP 23/07/08
Private Function varGetDestinatariViaggio() As Variant
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    If Not DS.DA.QueryCommand(SQL_COUNT_DESTINAZIONI, varBuffer, 0, DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If
    
    varGetDestinatariViaggio = varBuffer
    Exit Function
    
ErrHandler:
    varGetDestinatariViaggio = Empty
End Function

''-- DP 26/08/08
'Private Function strGetCounterSequenzaInvio(varIDUDT As Variant, varDTINV As Variant, varSEQVI As Variant) As String
'    Dim varBuffer As Variant
'    Dim strSQL As String
'
'    On Error GoTo ErrHandler
'
'    '-- Recupero il numeratore
'    If DS.DA.QueryCommand(SQL_GET_COUNTER_SEQUENZA_INVIO, varBuffer, 0, varIDUDT, varDTINV) Then
'        If Not IsEmpty(varBuffer) Then
'            strGetCounterSequenzaInvio = Format(CStr(CInt(varBuffer(0, 0)) + 10), "000")
'
'            strSQL = SQL_UPDATE_COUNTER_SEQUENZA_INVIO
'        Else
'            strGetCounterSequenzaInvio = Format(varSEQVI, "000")
'
'            strSQL = SQL_SET_COUNTER_SEQUENZA_INVIO
'        End If
'    Else
'        GoTo ErrHandler
'    End If
'
'    ' -- Aggiorno/Inserisco il numeratore
'    If Not DS.DA.ExecuteCommand(strSQL, strGetCounterSequenzaInvio, varIDUDT, varDTINV) Then
'        GoTo ErrHandler
'    End If
'
'    Exit Function
'
'ErrHandler:
'    strGetCounterSequenzaInvio = vbNullString
'End Function

'-- DP 09/09/08
Public Function blnUpdateSEQOK(strNewValue As String, varIDVVT As Variant) As Boolean
    On Error GoTo ErrHandler
    
    If strNewValue <> "S" And strNewValue <> "N" Then
        GoTo ErrHandler
    End If
    
    If Not DS.DA.ExecuteCommand(SQL_UPDATE_SEQOK, strNewValue, varIDVVT) Then
        GoTo ErrHandler
    End If
    
    blnUpdateSEQOK = True
    Exit Function
    
ErrHandler:
    blnUpdateSEQOK = False
End Function

'-- DP 13/09/08
Public Function blnUpdateSEQVI(strNuovaSequenza As String, varIDVVT As Variant) As Boolean
    On Error GoTo ErrHandler
    
    If Len(strNuovaSequenza) > 3 Then
        GoTo ErrHandler
    End If
    
    If Not DS.DA.ExecuteCommand(SQL_UPDATE_SEQIN, strNuovaSequenza, varIDVVT) Then
        GoTo ErrHandler
    End If
    
    blnUpdateSEQVI = blnUpdateSEQOK("N", varIDVVT)
    Exit Function
    
ErrHandler:
    blnUpdateSEQVI = False
End Function
