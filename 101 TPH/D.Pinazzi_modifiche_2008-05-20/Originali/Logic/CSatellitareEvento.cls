VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSatellitareEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String


Private objLogger As CLogger
Private SQL_KNOSOS_GET_EVENTS As String
Private SQL_KNOSOS_DELETE_EVENT As String
Private SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL As String
Private SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL As String
Private SQL_KNOSOS_GET_LAST_EVENT_DATE As String
Private SQL_KNOSOS_GET_IMPORTO_PEDAGGIO As String
Private SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI As String
Private SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT As String
Private SQL_UPDATE_VIAGGIO_MODIFICATO As String
Private SQL_GET_SELECTION_EVENTS As String
Private SQL_SELECT_KM_EVENTS As String
Private SQL_GET_KM_EVENT As String

Private SQL_SMS_GET_EVENTS As String
Private SQL_SMS_GET_LIST_AGG As String

Private SQL_DEUSTO_GET_VIAGGI As String
Private SQL_SAT_INSERT_POS As String
Private SQL_SAT_GET_MAX_POS_TRASM As String
Private SQL_SAT_GET_POS_TRASM_LIST As String
Private SQL_SAT_UPDATE_POS_STATUS As String
Private SQL_SAT_GET_LAST_POS As String
Private SQL_SAT_GET_LAST_POS_2 As String
Private SQL_SAT_GET_LAST_POS_3 As String
Private SQL_SAT_GET_LAST_POS_4 As String

Private Const ALIQUOTA_IVA_PEDAGGIO As Single = 16
Private Const KNOSOS_VIAGGIO_STATO_TERMINATO = 9

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Private Const VIAGGIO_STATO_PENDENTE = 0
Private Const VIAGGIO_STATO_INIZIOVIAGGIO = 1
Private Const VIAGGIO_STATO_CARICO = 2
Private Const VIAGGIO_STATO_INCORSO = 3
Private Const VIAGGIO_STATO_CONCLUSO = 4
Private Const VIAGGIO_STATO_CANCELLATO = 5

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function SAT_Get_Max_Pos() As Variant
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not DS.DA.QueryData(SQL_SAT_GET_MAX_POS_TRASM, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If
    
    If IsEmpty(varBuffer) Then
        SAT_Get_Max_Pos = 0
    Else
        SAT_Get_Max_Pos = varBuffer(0, 0)
    End If
    
    Exit Function
    
ErrHandler:
    SAT_Get_Max_Pos = Empty
End Function

Public Function SAT_Get_Last_Pos(IDAutomezzo As Variant, Data As Variant, Ora As Variant, ByRef CoordX As Variant, ByRef CoordY As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim strDataOra As String
    
    If Not IsDate(Data) Then
        GoTo ErrHandler
    End If
    If Not IsDate(Ora) Then
        GoTo ErrHandler
    End If
    
    strDataOra = Format(Data, "yyyymmdd") & Format(Ora, "hhnnss")
    
    If Not DS.DA.QueryData(SQL_SAT_GET_LAST_POS & DS.DA.MakeSQL(IDAutomezzo, dtAlfa) & SQL_SAT_GET_LAST_POS_2 & DS.DA.MakeSQL(DateAdd("d", -2, Data), dtDate) & SQL_SAT_GET_LAST_POS_3 & DS.DA.MakeSQL(Data, dtDate) & SQL_SAT_GET_LAST_POS_4, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        CoordX = Empty
        CoordY = Empty
    Else
        If Not IsDate(varBuffer(2, 0)) Then
            CoordX = Empty
            CoordY = Empty
        Else
            If Format(varBuffer(2, 0), "yyyymmdd") < Format(Data, "yyyymmdd") Then
            
                CoordX = varBuffer(0, 0)
                CoordY = varBuffer(1, 0)
            Else
                For i = 0 To UBound(varBuffer, 2)
                    If Format(varBuffer(3, i), "hhnnss") <= Format(Ora, "hhnnss") Then
                        CoordX = varBuffer(0, i)
                        CoordY = varBuffer(1, i)
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
        
    SAT_Get_Last_Pos = True
    
    Exit Function
    
ErrHandler:
    CoordX = Empty
    CoordY = Empty
    SAT_Get_Last_Pos = False
End Function
Public Function SAT_Get_Pos_List(MaxID As Variant) As Variant
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not DS.DA.QueryCommand(SQL_SAT_GET_POS_TRASM_LIST, varBuffer, lngHowMany, MaxID) Then
        GoTo ErrHandler
    End If
    
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
        
    SAT_Get_Pos_List = varBuffer
    
    Exit Function
    
ErrHandler:
    SAT_Get_Pos_List = Empty
End Function
Public Function SAT_Update_Pos_Status(MaxID As Variant) As Variant
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not DS.DA.ExecuteCommand(SQL_SAT_UPDATE_POS_STATUS, MaxID) Then
        GoTo ErrHandler
    End If
    
    SAT_Update_Pos_Status = True
    
    Exit Function
    
ErrHandler:
    SAT_Update_Pos_Status = False
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
Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean
    
    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    BeginSave = True
    
End Function

Private Sub Class_Initialize()
    SQL_FIND = "SELECT SEVIDSEV, SEVIDVVT, SEVIDVVC, SEVIDVVP," & _
                " SEVIDEVE, SEVDTEVE, SEVQTLTC, SEVQTLTS, SEVNRKM," & _
                " SEVFLSAT, TIMESTMP FROM VIAGSEV0 WHERE "
    SQL_GETLIST = "SELECT SEVIDSEV, SEVIDVVT, SEVIDVVC, SEVIDVVP FROM VIAGSEV0 "
    SQL_UPDATE = "UPDATE VIAGSEV0 SET "
    SQL_INSERT = "INSERT INTO VIAGSEV0 "
    SQL_GET_TS = "SELECT TIMESTMP FROM VIAGSEV0 WHERE "
    
    SQL_SMS_GET_EVENTS = "SELECT SMSIDSMS, SMSNRTEL, SMSDTSMD, SMSDTLOC, " & _
                         "       SMSTESTO " & _
                         "  FROM VIAGSMS0 " & _
                         " WHERE SMSIDSMS > ? " & _
                         " ORDER BY SMSIDSMS "
    
    SQL_SMS_GET_LIST_AGG = "SELECT VVTIDAM1, VVCIDVVT, VVCIDVVC, VVTFLST3, " & _
                           "       VVCFLST5, VVTHRINC, CASE VVCHRPRE WHEN '01/01/1900' THEN VVCTEMVG ELSE VVCHRPRE END as ORA_PREV, VVCTEMVG  " & _
                           "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT " & _
                           " WHERE (VVTDTINV = ?) AND (VVTANNUL = 0) AND (VVCANNUL = 0) " & _
                           " ORDER BY VVTIDAM1, VVTDTINC, VVTHRINC, ORA_PREV "
    
    
    SQL_KNOSOS_GET_EVENTS = "SELECT IDVEH, IDVIAJE, NUMVIAJE, CONVERT( nvarchar, FECHA, 21 ), IDDESC, PROD_RES, " & _
                                " TEVE, CANTREALCAR, CANTREALDESC, KM_PARCIALES, " & _
                                " CONVERT( nVarchar, DATEADD(Hour, DATEDIFF(Hour, GETUTCDATE(), GETDATE()), FECHA), 21 )" & _
                                " FROM KNOSEVE0" & _
                                " ORDER BY FECHA"

    SQL_KNOSOS_DELETE_EVENT = "DELETE KNOSEVE0 WHERE IDVEH = ? AND IDVIAJE = ? AND NUMVIAJE = ?" & _
                                " AND FECHA = { Ts '%fecha%' } AND IDDESC = ?"
    
    SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL = "SELECT VVPIDVVT, VVPIDVVC, VVPIDVVP" & _
                                        " FROM VIAGVVT0 INNER JOIN" & _
                                        " VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN" & _
                                        " VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT" & _
                                        " AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC" & _
                                        " WHERE VIAGVVT0.VVTNRKVI = ? AND VIAGVVP0.VVPIDKDE = ?"
    
    SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL = "SELECT VVTIDVVT FROM VIAGVVT0 WHERE VVTNRKVI = ?"
    
    SQL_KNOSOS_GET_LAST_EVENT_DATE = "SELECT MAX(SEVDTEVE)" & _
                                        " FROM VIAGSEV0"
    
    SQL_KNOSOS_GET_IMPORTO_PEDAGGIO = "SELECT PEAJE FROM KNOSVIA0" & _
                                    " WHERE NUMVIAJE = ? and ESTA = ?"
    
    SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI = "SELECT COUNT(*) FROM VIAGVVP0" & _
                                    " INNER JOIN VIAGVVC0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC" & _
                                    " WHERE VVPIDVVT = ?" & _
                                    " AND VVPIDVVC = ?" & _
                                    " AND (VVPANNUL = 0 AND VVCANNUL = 0)" & _
                                    " AND VVPFLSTA <> '2'"
                                    
    SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT = "SELECT COUNT(*) FROM VIAGVVP0" & _
                                    " INNER JOIN VIAGVVC0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC" & _
                                    " WHERE VVPIDVVT = ?" & _
                                    " AND (VVPANNUL = 0 AND VVCANNUL = 0)" & _
                                    " AND VVPFLSTA <> '2'"

    SQL_UPDATE_VIAGGIO_MODIFICATO = "UPDATE VIAGVVP0 SET VVPFLST5 = 'S' WHERE VVPIDVVT = ? AND VVPANNUL = 0"
    SQL_SELECT_KM_EVENTS = "SELECT SEVIDVVC, MIN(SEVNRKM), MAX(SEVNRKM), MIN(SEVDTEVE)" & _
                            " FROM VIAGSEV0" & _
                            " WHERE  SEVIDVVT = ? AND SEVIDVVC <> ''" & _
                            " GROUP  BY SEVIDVVC" & _
                            " ORDER  BY MIN(SEVNRKM), MIN(SEVDTEVE)"
 
    SQL_GET_KM_EVENT = "SELECT SEVNRKM FROM VIAGSEV0" & _
                                " WHERE SEVIDVVT = ? AND SEVIDEVE = ?"
    
    SQL_GET_SELECTION_EVENTS = "SELECT SEVIDSEV, SEVIDVVT, SEVIDVVC, SEVIDVVP," & _
                            " AVV_B.AVVDESCR, SEVDTEVE, IsNull(vvcderas, adedescr), SEVQTLTC, SEVQTLTS, SEVNRKM," & _
                            " AVV_A.AVVDESCR, SEVFLSAT, SEVIDEVE FROM VIAGSEV0" & _
                            " INNER JOIN ANAGAVV AVV_A ON AVV_A.AVVIDTBL = 'ANAGAAM0' AND" & _
                            " AVV_A.AVVIDFLD = 'AAMFLSAT' AND AVV_A.AVVVALUE = VIAGSEV0.SEVFLSAT" & _
                            " INNER JOIN ANAGAVV AVV_B ON AVV_B.AVVIDTBL = 'VIAGSEV0' AND " & _
                            " AVV_B.AVVIDFLD = 'SEVIDEVE' AND AVV_B.AVVVALUE = VIAGSEV0.SEVIDEVE" & _
                            " INNER JOIN VIAGVVT0 ON VVTIDVVT = SEVIDVVT" & _
                            " LEFT JOIN ANAGADE0 ON ADEIDADE = VVTIDAD1 AND (SEVIDEVE = 'CI' OR SEVIDEVE = 'CF')" & _
                            " LEFT JOIN VIAGVVC0 ON VVCIDVVT = SEVIDVVT AND VVCIDVVC = SEVIDVVC" & _
                            " WHERE SEVIDVVT = '%sevidvvt%' ORDER BY SEVDTEVE DESC"

    SQL_DEUSTO_GET_VIAGGI = "SELECT DISTINCT AAMIDDEU, VVTIDAM1 " & _
                            "  FROM VIAGVVT0 INNER JOIN ANAGAAM0 ON VVTIDAM1 = AAMIDAAM " & _
                            " WHERE (VVTDTINV = ? OR VVTDTINC = ?) AND VVTIDASC = '038' AND VVTFLSTV <> 'P' AND VVTANNUL = 0 AND (AAMIDDEU IS NOT NULL) AND (LEN(AAMIDDEU)>0) "
    SQL_SAT_INSERT_POS = "INSERT INTO VIAGPOS0 " & _
                            "       (POSIDAAM, POSDTPOS, POSHRPOS, POSCOORX, POSCOORY, POSFLTRA, POSNRSTA) " & _
                            "VALUES ("
    SQL_SAT_GET_MAX_POS_TRASM = "SELECT MAX(POSIDPOS) " & _
                                   "  FROM VIAGPOS0 " & _
                                   " WHERE POSFLTRA = 'N'"
    SQL_SAT_GET_POS_TRASM_LIST = "SELECT POSIDAAM, POSDTPOS, POSHRPOS, POSCOORX, POSCOORY " & _
                                    "  FROM VIAGPOS0 " & _
                                    " WHERE POSIDPOS <= ? AND POSFLTRA = 'N'"
    SQL_SAT_UPDATE_POS_STATUS = "UPDATE VIAGPOS0 " & _
                                   "   SET POSFLTRA = 'S' " & _
                                   " WHERE POSIDPOS <= ? AND POSFLTRA = 'N'"
'    SQL_SAT_GET_LAST_POS = "SELECT TOP 1 POSCOORX, POSCOORY " & _
'                           "  FROM VIAGPOS0  INNER JOIN ANAGAAM0 ON AAMIDAAM = POSIDAAM " & _
'                           " WHERE (POSIDAAM = ?) AND (AAMIDDEU IS NOT NULL) AND (AAMIDDEU <> '') AND " & _
'                           " (CAST(DATEPART(yyyy, POSDTPOS) AS CHAR(4)) + RIGHT(CAST(DATEPART(mm, POSDTPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(dd, POSDTPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(hh, POSHRPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(n, POSHRPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(ss, POSHRPOS) + 100 AS CHAR(3)), 2) <= ?) " & _
'                           " ORDER BY CAST(DATEPART(yyyy, POSDTPOS) AS CHAR(4)) + RIGHT(CAST(DATEPART(mm, POSDTPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(dd, POSDTPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(hh, POSHRPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(n, POSHRPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(ss, POSHRPOS) + 100 AS CHAR(3)), 2) DESC "
    SQL_SAT_GET_LAST_POS = "SELECT POSCOORX, POSCOORY, POSDTPOS, POSHRPOS " & _
                           "  FROM VIAGPOS0  INNER JOIN ANAGAAM0 ON AAMIDAAM = POSIDAAM " & _
                           " WHERE (POSIDAAM = "
                           
    SQL_SAT_GET_LAST_POS_2 = ") AND (POSDTPOS >= "
    
    SQL_SAT_GET_LAST_POS_3 = ") AND (POSDTPOS <= "
                           
    SQL_SAT_GET_LAST_POS_4 = ") ORDER BY POSDTPOS DESC, POSHRPOS DESC "
    
    Set DS = New CDataStruct

    With DS.Fields
        .Add "Numero Evento", "VIAGSEV0.sevidsev", 0, True, 7, "0", , True
        .Add "Numero Viaggio", "VIAGSEV0.sevidvvt", 0, True, 7, "1", , True
        .Add "Numero Consegna", "VIAGSEV0.sevidvvc", 0, True, 3, "2", , True
        .Add "Numero Scarico", "VIAGSEV0.sevidvvp", 0, True, 3, "3", , True
        .Add "ID evento", "VIAGSEV0.sevideve", 0, True, 3, "4", , True
        .Add "Data evento", "VIAGSEV0.sevdteve", 2, True, 23, "5", , True
        .Add "Litri caricati", "VIAGSEV0.sevqtltc", 1, False, 3, "6", , True
        .Add "Litri scaricati", "VIAGSEV0.sevqtlts", 1, False, 3, "7", , True
        .Add "Numero KM", "VIAGSEV0.sevnrkm", 1, False, 3, "8", , True
        .Add "Tipo satellitare", "VIAGSEV0.sevflsat", 0, True, 1, "9", , True
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
  End With
    DS.CancelField = ""
    
    Set objLogger = New CLogger

End Sub

Private Sub Class_Terminate()
    Set DS = Nothing
    Set objLogger = Nothing
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
Function StringKNOSOSDate_To_Date(strDate As String)
    If Len(strDate) = 23 Then
        strDate = Left(strDate, Len(strDate) - 4)
    End If
    
    StringKNOSOSDate_To_Date = CDate(strDate)
End Function
Function convertLocalDateToUTCDate(strDate As String) As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strSQLStatement As String

    strSQLStatement = Replace("SELECT CONVERT( nVarchar, DATEADD(hour, -DATEDIFF(hour, GETUTCDATE(), GETDATE()), CONVERT(datetime, { Ts '%fecha%' }, 21)), 21)", "%fecha%", strDate)
    If Not DS.DA.QueryCommand(strSQLStatement, varBuffer, lngHowMany) Then
        convertLocalDateToUTCDate = ""
        Exit Function
    End If
        
    convertLocalDateToUTCDate = varBuffer(0, 0)
    
End Function
Function TruncDate(strDate As Date)
    TruncDate = CDate(DatePart("h", strDate) & ":" & DatePart("n", strDate) & "." & DatePart("s", strDate))
End Function
Private Sub LogFieldValue(Field As CField)

    If Field.Updated Then
        objLogger.logDebug "+ """ & Field.Name & """ (" & Field.FieldName & ") = [" & Field.NewValue & "]"
    Else
        objLogger.logDebug "  """ & Field.Name & """ (" & Field.FieldName & ") = [" & Field.Value & "]"
    End If
End Sub
Function ScanEvents_KNOSOS() As Boolean
    
    Dim i As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim blnUsaDatiGPS As Boolean

    objLogger.logDebug "Inizio ScanEvents_KNOSOS."
    
    ' Scan KNOSOS table events
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_EVENTS, varKBuffer, lngHowMany) Then
        objLogger.logError "Query SQL_KNOSOS_GET_EVENTS."
        GoTo ErrHandler
    End If

    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    ' Inizia transazione
    DS.DA.BeginTransaction

    objLogger.logInfo "Trovati " & lngHowMany & " eventi."
    
    If lngHowMany > 0 Then

        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)

            varIDViaggio = ""
            varIDConsegna = ""
            varIDProdotto = ""

            objLogger.logInfo "Evento n. " & i + 1
            objLogger.logInfo "-------------------------------------------------"
            objLogger.logInfo "IDVEH         [" & varKBuffer(0, i) & "]"
            objLogger.logInfo "IDVIAJE       [" & varKBuffer(1, i) & "]"
            objLogger.logInfo "NUMVIAJE      [" & varKBuffer(2, i) & "]"
            objLogger.logInfo "FECHA         [" & varKBuffer(3, i) & "]"
            objLogger.logInfo "IDDESC        [" & varKBuffer(4, i) & "]"
            objLogger.logInfo "TEVE          [" & varKBuffer(5, i) & "]"
            objLogger.logInfo "PROD_RES      [" & varKBuffer(6, i) & "]"
            objLogger.logInfo "CANTREALCAR   [" & varKBuffer(7, i) & "]"
            objLogger.logInfo "CANTREALDESC  [" & varKBuffer(8, i) & "]"
            objLogger.logInfo "KM_PARCIALES  [" & varKBuffer(9, i) & "]"
            objLogger.logInfo "-------------------------------------------------"
            
            ' E' un evento relativo al viaggio o carico o scarico?
            If CSng(varKBuffer(4, i)) > 0 Then
                
                ' E' uno scarico.
                lngHowMany = 0
                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL, varBuffer, lngHowMany, _
                    CStr(varKBuffer(2, i)), CSng(varKBuffer(4, i))) Then
                    objLogger.logError "Query SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL."
                    GoTo ErrHandler
                End If
                If lngHowMany > 0 Then
                    varIDViaggio = varBuffer(0, 0)
                    varIDConsegna = varBuffer(1, 0)
                    varIDProdotto = varBuffer(2, 0)
                Else
                    objLogger.logWarning "Consegna prodotto non trovato in ResOil."
                End If
            Else
                
                ' E' un carico o viaggio.
                lngHowMany = 0
                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL, varBuffer, lngHowMany, _
                    CStr(varKBuffer(2, i))) Then
                    objLogger.logError "Query SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL."
                    GoTo ErrHandler
                End If
                If lngHowMany > 0 Then
                    varIDViaggio = varBuffer(0, 0)
                    varIDConsegna = ""
                    varIDProdotto = ""
                Else
                    objLogger.logWarning "Viaggio non trovato in ResOil."
                End If
            End If
            
            objLogger.logInfo "ResOil Keys VVTIDVVT = """ & varIDViaggio & _
                                """, VVCIDVVC = """ & varIDConsegna & """, VVPIDVVP = """ & varIDProdotto & """"
            
            ' Log record event
            DS.DA.LogExecTime "INFO: KNOSOS Evento", -1
            DS.DA.LogExecTime "INFO: -----------------------------------------------------------------", -1
            DS.DA.LogExecTime "INFO: - IDVEH         [" & varKBuffer(0, i) & "]", -1
            DS.DA.LogExecTime "INFO: - IDVIAJE       [" & varKBuffer(1, i) & "]", -1
            DS.DA.LogExecTime "INFO: - NUMVIAJE      [" & varKBuffer(2, i) & "]", -1
            DS.DA.LogExecTime "INFO: - FECHA         [" & varKBuffer(3, i) & "]", -1
            DS.DA.LogExecTime "INFO: - IDDESC        [" & varKBuffer(4, i) & "]", -1
            DS.DA.LogExecTime "INFO: - TEVE          [" & varKBuffer(5, i) & "]", -1
            DS.DA.LogExecTime "INFO: - PROD_RES      [" & varKBuffer(6, i) & "]", -1
            DS.DA.LogExecTime "INFO: - CANTREALCAR   [" & varKBuffer(7, i) & "]", -1
            DS.DA.LogExecTime "INFO: - CANTREALDESC  [" & varKBuffer(8, i) & "]", -1
            DS.DA.LogExecTime "INFO: - KM_PARCIALES  [" & varKBuffer(9, i) & "]", -1

            ' Se non trovo il viaggio o prodotto, skip
            If varIDViaggio <> "" Then

                dateLocalEventDateTime = StringKNOSOSDate_To_Date(CStr(varKBuffer(10, i)))
                If lngHowMany > 0 Then
        
                    DS.Fields.ResetFields
            
                    ' Creazione dell'evento
                    DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "901", 0)
                    DS.Fields("0").Updated = True
            
                    DS.Fields("1").NewValue = varIDViaggio
                    DS.Fields("1").Updated = True
        
                    DS.Fields("2").NewValue = varIDConsegna
                    DS.Fields("2").Updated = True
        
                    DS.Fields("3").NewValue = varIDProdotto
                    DS.Fields("3").Updated = True
        
                    DS.Fields("5").NewValue = varKBuffer(10, i)
                    DS.Fields("5").Updated = True
        
                    DS.Fields("6").NewValue = varKBuffer(7, i)
                    DS.Fields("6").Updated = True
        
                    DS.Fields("7").NewValue = varKBuffer(8, i)
                    DS.Fields("7").Updated = True
        
                    DS.Fields("8").NewValue = varKBuffer(9, i)
                    DS.Fields("8").Updated = True
                    
                    DS.Fields("9").NewValue = "K"
                    DS.Fields("9").Updated = True
        
                    ' Carica il viaggio
                    With VVT
                        .DS.Fields.ResetFields
                        .DS.Fields("0").FindValue = varIDViaggio
                    
                        If Not .Find Then
                            objLogger.logError "Find Viaggio."
                            GoTo ErrHandler
                        End If
                        
                        If IsEmptyVariantString(.DS.Field("Flag Usa Dati GPS").Value) Or _
                            .DS.Field("Flag Usa Dati GPS").Value = "S" Then
                            blnUsaDatiGPS = True
                        Else
                            blnUsaDatiGPS = False
                        End If
                    
                    End With
                    
                        
                    Select Case varKBuffer(6, i)

                        ' -------------------------------------------------------
                        ' Evento - Inicio del viaggio (Inizio del viaggio)
                        ' -------------------------------------------------------
                        Case "ON "
                            DS.Fields("4").NewValue = "VI"
                            DS.Fields("4").Updated = True
    
                            objLogger.logInfo "Tipo evento: Inizio viaggio"

                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INIZIOVIAGGIO
                            VVT.DS.Fields("48").Updated = True
                            LogFieldValue VVT.DS.Fields("48")
                            
                            If blnUsaDatiGPS Then
                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                    objLogger.logError "Errore nell'aggiornamento del flag Modificato."
                                    GoTo ErrHandler
                                End If
                            End If

                        ' -------------------------------------------------------
                        ' Evento - Llegada en factora (Inizio carico)
                        ' -------------------------------------------------------
                        Case "LLE"
                            DS.Fields("4").NewValue = "CI"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Inizio carico"
                            
                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_CARICO
                            VVT.DS.Fields("48").Updated = True
                            LogFieldValue VVT.DS.Fields("48")
    
                            If blnUsaDatiGPS Then
                                ' Aggiorna l'ora inizio carico sul viaggio (VVTHRINC)
                                VVT.DS.Fields("5").NewValue = TruncDate(dateLocalEventDateTime)
                                VVT.DS.Fields("5").Updated = True
                                LogFieldValue VVT.DS.Fields("5")
                            
                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                    objLogger.logError "Errore nell'aggiornamento del flag Modificato."
                                    GoTo ErrHandler
                                End If
                            End If
    
                        ' -------------------------------------------------------
                        ' Evento - Salida de factora (Fine del carico)
                        ' -------------------------------------------------------
                        Case "SAL"
                            DS.Fields("4").NewValue = "CF"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Fine carico"
                            
                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
                            VVT.DS.Fields("48").Updated = True
                            LogFieldValue VVT.DS.Fields("48")

                            ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
                            If blnUsaDatiGPS Then
                                VVT.DS.Fields("30").NewValue = DateDiff("n", TruncDate(VVT.DS.Fields("5").Value), TruncDate(dateLocalEventDateTime))
                                VVT.DS.Fields("30").Updated = True
                                LogFieldValue VVT.DS.Fields("30")

                                ' Aggiorna l'ora inizio viaggio sul viaggio (VVTHRINV)
                                VVT.DS.Fields("7").NewValue = TruncDate(dateLocalEventDateTime)
                                VVT.DS.Fields("7").Updated = True
                                LogFieldValue VVT.DS.Fields("7")
                                
                                ' Aggiorna flag carico completato
                                VVT.DS.Fields("39").NewValue = "S"
                                VVT.DS.Fields("39").Updated = True
                                LogFieldValue VVT.DS.Fields("39")

                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                    objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                    GoTo ErrHandler
                                End If
                            End If

                        ' -------------------------------------------------------
                        ' Evento - Inicio descarga (Inizio scarico prodotto)
                        ' -------------------------------------------------------
                        Case "INI"
                            DS.Fields("4").NewValue = "SI"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
                            
                            ' Carica la consegna prodotto e aggiorna i dati
                            With VVP
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                                .DS.Fields("2").FindValue = varIDProdotto
                                
                                If Not .Find Then
                                    objLogger.logError "Find Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
        
                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                .DS.Fields("37").NewValue = "1"
                                .DS.Fields("37").Updated = True
                                LogFieldValue .DS.Fields("37")

                                ' Aggiorna il flag Modificato (VVPFLST5)
                                .DS.Field("Stato Riga 5").NewValue = "S"
                                .DS.Field("Stato Riga 5").Updated = True
                                LogFieldValue .DS.Field("Stato Riga 5")

                                ' Aggiorna l'ora inizio scarico (VVPHRINC)
                                If blnUsaDatiGPS Then
                                    .DS.Fields("38").NewValue = TruncDate(dateLocalEventDateTime)
                                    .DS.Fields("38").Updated = True
                                    LogFieldValue .DS.Fields("38")
                                End If

                                ' Salva
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
    
                                .DS.DA.EndTrans etCommit
                            End With
                            
                            ' Carica la consegna e aggiorna i dati
                            With VVC
                                
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                        
                                If Not .Find Then
                                    objLogger.logError "Find Consegna."
                                    GoTo ErrHandler
                                End If
                                
                                ' E' il primo scarico della consegna
                                If .DS.Fields("44").Value = "0" Or IsNull(.DS.Fields("44").Value) Or .DS.Fields("44").Value = "" Then
                                    
                                    If blnUsaDatiGPS Then
                                        ' Aggiorna l'ora di arrivo
                                        .DS.Fields("24").NewValue = TruncDate(dateLocalEventDateTime)
                                        .DS.Fields("24").Updated = True
                                        LogFieldValue .DS.Fields("24")
                                    End If

                                    ' Aggiorna lo stato 'In corso'
                                    .DS.Fields("44").NewValue = "1"
                                    .DS.Fields("44").Updated = True
                                    LogFieldValue .DS.Fields("44")
                                End If

'                                ' Imposta i km parziali
'                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
'                                    .DS.Fields("22").Updated = True
'                                End If
                                
                                ' Salva la consegna
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.DA.EndTrans etCommit
                            End With
    
                        ' -------------------------------------------------------
                        ' Evento - Fin descarga (Fine scarico prodotto)
                        ' -------------------------------------------------------
                        Case "FIN"
                            DS.Fields("4").NewValue = "SF"
                            DS.Fields("4").Updated = True
                            
                            objLogger.logInfo "Tipo evento: Fine consegna prodotto"
                            
                            ' Carica la consegna prodotto e aggiorna i dati
                            With VVP
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                                .DS.Fields("2").FindValue = varIDProdotto
                                
                                If Not .Find Then
                                    objLogger.logError "Find Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
        
                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                .DS.Fields("37").NewValue = "2"
                                .DS.Fields("37").Updated = True
                                LogFieldValue .DS.Fields("37")
                                
                                If blnUsaDatiGPS Then
                                    ' Imposta i litri scaricati
                                    If IsNull(varKBuffer(8, i)) Or IsEmpty(varKBuffer(8, i)) Then
                                        .DS.Fields("11").NewValue = 0
                                    Else
                                        .DS.Fields("11").NewValue = CSng(varKBuffer(8, i))
                                    End If
                                    .DS.Fields("11").Updated = True
                                    LogFieldValue .DS.Fields("11")
        
                                    ' Imposta i litri caricati
                                    .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
                                    .DS.Fields("7").Updated = True
                                    LogFieldValue .DS.Fields("7")
    
                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
                                    .DS.Fields("39").Updated = True
                                    LogFieldValue .DS.Fields("39")
                                
                                    ' Aggiorna il flag Modificato (VVPFLST5)
                                    .DS.Field("Stato Riga 5").NewValue = "S"
                                    .DS.Field("Stato Riga 5").Updated = True
                                    LogFieldValue .DS.Field("Stato Riga 5")
                                End If

                                ' Salva
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
    
                                .DS.DA.EndTrans etCommit
                            End With
        
                            ' Carica la consegna e aggiorna i dati
                            With VVC
                                
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                        
                                If Not .Find Then
                                    objLogger.logError "Find Consegna."
                                    GoTo ErrHandler
                                End If
    
                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
                                lngHowMany = 0
                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
                                    varIDViaggio, varIDConsegna) Then
                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
                                    GoTo ErrHandler
                                End If
                                
                                If CLng(varBuffer(0, 0)) = 0 Then
                                    
                                    If blnUsaDatiGPS Then
                                        ' Aggiorna il tempo di scarico
                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("24").Value), TruncDate(dateLocalEventDateTime))
                                        .DS.Fields("25").Updated = True
                                        LogFieldValue .DS.Fields("25")

                                        ' Aggiorna il flag consegna ultimata
                                        .DS.Fields("36").NewValue = "S"
                                        .DS.Fields("36").Updated = True
                                        LogFieldValue .DS.Fields("36")
                                    End If

                                    ' Aggiorna lo stato 'Completo'
                                    .DS.Fields("44").NewValue = "2"
                                    .DS.Fields("44").Updated = True
                                    LogFieldValue .DS.Fields("44")
                                End If
                            
'                                ' Imposta i km parziali
'                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
'                                    .DS.Fields("22").Updated = True
'                                End If
                                    
                                ' Salva la consegna
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.DA.EndTrans etCommit
                            End With
                        
                        ' -------------------------------------------------------
                        ' Evento - Rechazo de descarga (Dirottamento scarico)
                        ' -------------------------------------------------------
                        Case "REC"
                            DS.Fields("4").NewValue = "SA"
                            DS.Fields("4").Updated = True

                            objLogger.logInfo "Tipo evento: Dirottamento o Rientro"
                            
                            ' Carica la consegna
                            With VVC
                                
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                        
                                If Not .Find Then
                                    objLogger.logError "Find Consegna."
                                    GoTo ErrHandler
                                End If
                            
                            End With
                            
                            ' Carica la consegna prodotto e aggiorna i dati
                            With VVP
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                                .DS.Fields("2").FindValue = varIDProdotto
                                
                                If Not .Find Then
                                    objLogger.logError "Find Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
        
                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                .DS.Fields("37").NewValue = "2"
                                .DS.Fields("37").Updated = True
                                LogFieldValue .DS.Fields("37")

                                ' Imposta i litri caricati e scaricati
                                Select Case VVC.DS.Fields("33").Value
                                    
                                    ' Dirottamento
                                    Case "D"
                                        objLogger.logInfo "Tipo evento: Dirottamento"
                                        If IsNull(varKBuffer(8, i)) Or IsEmpty(varKBuffer(8, i)) Then
                                            .DS.Fields("11").NewValue = 0
                                        Else
                                            .DS.Fields("11").NewValue = CSng(varKBuffer(8, i))
                                        End If
                                        .DS.Fields("11").Updated = True
                                        LogFieldValue .DS.Fields("11")

                                        ' Imposta i litri caricati
                                        If blnUsaDatiGPS Then
                                            .DS.Fields("7").NewValue = 0
                                            .DS.Fields("7").Updated = True
                                            LogFieldValue .DS.Fields("7")
                                        End If

                                    ' Rientro
                                    Case "R"
                                        objLogger.logInfo "Tipo evento: Rientro"
                                        
                                        If blnUsaDatiGPS Then
                                            ' Imposta i litri scaricati
                                            If IsNull(varKBuffer(7, i)) Or IsEmpty(varKBuffer(7, i)) Then
                                                .DS.Fields("11").NewValue = 0
                                            Else
                                                .DS.Fields("11").NewValue = CSng(varKBuffer(7, i))
                                            End If
                                            .DS.Fields("11").Updated = True
                                            LogFieldValue .DS.Fields("11")
    
                                            ' Imposta i litri caricati
                                            .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
                                            .DS.Fields("7").Updated = True
                                            LogFieldValue .DS.Fields("7")
                                        End If
    
                                End Select

                                If blnUsaDatiGPS Then
                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
                                    .DS.Fields("39").Updated = True
                                    LogFieldValue .DS.Fields("39")
                                
                                    ' Aggiorna il flag Modificato (VVPFLST5)
                                    .DS.Field("Stato Riga 5").NewValue = "S"
                                    .DS.Field("Stato Riga 5").Updated = True
                                    LogFieldValue .DS.Field("Stato Riga 5")
                                End If

                                ' Salva
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
    
                                .DS.DA.EndTrans etCommit
                            End With
        
                            ' Carica la consegna e aggiorna i dati
                            With VVC

                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
                                lngHowMany = 0
                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
                                    varIDViaggio, varIDConsegna) Then
                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
                                    GoTo ErrHandler
                                End If
                                
                                If CLng(varBuffer(0, 0)) = 0 Then
                                    
                                    If blnUsaDatiGPS Then
                                        ' Aggiorna il tempo di scarico
                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("24").Value), TruncDate(dateLocalEventDateTime))
                                        .DS.Fields("25").Updated = True
                                        LogFieldValue .DS.Fields("25")
    
                                        ' Aggiorna il flag consegna ultimata
                                        .DS.Fields("36").NewValue = "S"
                                        .DS.Fields("36").Updated = True
                                        LogFieldValue .DS.Fields("36")
                                    End If

                                    ' Aggiorna lo stato 'Completo'
                                    .DS.Fields("44").NewValue = "2"
                                    .DS.Fields("44").Updated = True
                                    LogFieldValue .DS.Fields("44")
                                End If

'                                ' Imposta i km parziali
'                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
'                                    .DS.Fields("22").Updated = True
'                                End If

                                ' Salva la consegna
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.DA.EndTrans etCommit
                            End With
                            
                        ' -------------------------------------------------------
                        ' Evento - Devolucin de la descarga (Rientro scarico)
                        ' -------------------------------------------------------
                        Case "DEV"
                            DS.Fields("4").NewValue = "SD"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Rientro"
                            
                            ' Carica la consegna prodotto e aggiorna i dati
                            With VVP
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                                .DS.Fields("2").FindValue = varIDProdotto
                                
                                If Not .Find Then
                                    objLogger.logError "Find Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
        
                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                .DS.Fields("37").NewValue = "2"
                                .DS.Fields("37").Updated = True
                                LogFieldValue .DS.Fields("37")

                                If blnUsaDatiGPS Then
                                    ' Imposta i litri scaricati
                                    If IsNull(varKBuffer(7, i)) Or IsEmpty(varKBuffer(7, i)) Then
                                        .DS.Fields("11").NewValue = 0
                                    Else
                                        .DS.Fields("11").NewValue = CSng(varKBuffer(7, i))
                                    End If
                                    .DS.Fields("11").Updated = True
                                    LogFieldValue .DS.Fields("11")

                                    ' Imposta i litri caricati
                                    .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
                                    .DS.Fields("7").Updated = True
                                    LogFieldValue .DS.Fields("7")
    
                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
                                    .DS.Fields("39").Updated = True
                                    LogFieldValue .DS.Fields("39")
                                
                                    ' Aggiorna il flag Modificato (VVPFLST5)
                                    .DS.Field("Stato Riga 5").NewValue = "S"
                                    .DS.Field("Stato Riga 5").Updated = True
                                    LogFieldValue .DS.Field("Stato Riga 5")
                                End If

                                ' Salva
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
    
                                .DS.DA.EndTrans etCommit
                            End With
        
                            ' Carica la consegna e aggiorna i dati
                            With VVC
                                
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                        
                                If Not .Find Then
                                    objLogger.logError "Find Consegna."
                                    GoTo ErrHandler
                                End If
    
                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
                                lngHowMany = 0
                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
                                    varIDViaggio, varIDConsegna) Then
                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
                                    GoTo ErrHandler
                                End If
                                
                                If CLng(varBuffer(0, 0)) = 0 Then
                                    
                                    If blnUsaDatiGPS Then
                                        ' Aggiorna il tempo di scarico
                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("24").Value), TruncDate(dateLocalEventDateTime))
                                        .DS.Fields("25").Updated = True
                                        LogFieldValue .DS.Fields("25")
    
                                        ' Aggiorna il flag consegna ultimata
                                        .DS.Fields("36").NewValue = "S"
                                        .DS.Fields("36").Updated = True
                                        LogFieldValue .DS.Fields("36")
                                    End If

                                    ' Aggiorna lo stato 'Completo'
                                    .DS.Fields("44").NewValue = "2"
                                    .DS.Fields("44").Updated = True
                                    LogFieldValue .DS.Fields("44")
                                End If
                            
'                                ' Imposta i km parziali
'                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
'                                    .DS.Fields("22").Updated = True
'                                End If
                                    
                                ' Salva la consegna
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.DA.EndTrans etCommit
                            End With
                        
                        ' -------------------------------------------------------
                        ' Evento - Fin de viaje (Fine del viaggio)
                        ' -------------------------------------------------------
                        Case "OFF"
                            Dim sngImportoPedaggio As Single
                            Dim lngKmInizioViaggio As Long
                            Dim lngKmInizioCarico As Long
                            
                            objLogger.logInfo "Tipo evento: Fine viaggio"
                            
                            DS.Fields("4").NewValue = "VF"
                            DS.Fields("4").Updated = True

                            With VVT
                                
                                ' Aggiorna lo stato del viaggio (VVTFLSTA)
                                .DS.Fields("48").NewValue = VIAGGIO_STATO_CONCLUSO
                                .DS.Fields("48").Updated = True
                                LogFieldValue .DS.Fields("48")
                                
                                If blnUsaDatiGPS Then
                                    ' Aggiorna l'ora fine viaggio (VVTHRFIV)
                                    .DS.Fields("9").NewValue = TruncDate(dateLocalEventDateTime)
                                    .DS.Fields("9").Updated = True
                                    LogFieldValue .DS.Fields("9")
    
                                    ' Aggiorna il tempo del viaggio sul viaggio (VVTTMPVI)
                                    .DS.Fields("29").NewValue = DateDiff("n", TruncDate(.DS.Fields("7").Value), TruncDate(dateLocalEventDateTime))
                                    .DS.Fields("29").Updated = True
                                    LogFieldValue .DS.Fields("29")
    
                                    ' Preleva i km di fine carico
                                    lngHowMany = 0
                                    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CF") Then
                                        objLogger.logError "Query SQL_GET_KM_EVENT."
                                        GoTo ErrHandler
                                    End If
    
                                    If lngHowMany = 0 Then
                                        objLogger.logError "Impossibile calcolare il km definitivi del viaggio. Manca l'evento fine carico."
                                    Else
                                        ' Aggiorna i km definitivi sul viaggio
                                        If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
                                            .DS.Fields("27").NewValue = CLng(varKBuffer(9, i)) - varBuffer(0, 0)
                                            .DS.Fields("27").Updated = True
                                        End If
                                        LogFieldValue .DS.Fields("27")
                                    End If
    
                                    ' Preleva i km di inizio viaggio
                                    lngHowMany = 0
                                    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "VI") Then
                                        objLogger.logError "Query SQL_GET_KM_EVENT."
                                        GoTo ErrHandler
                                    End If
    
                                    If lngHowMany = 0 Then
                                        objLogger.logError "Impossibile aggiornare i km a vuoto. Manca l'evento inizio viaggio."
                                    Else
                                        lngKmInizioViaggio = varBuffer(0, 0)
                                        
                                        ' Preleva i km di inizio carico
                                        lngHowMany = 0
                                        If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CI") Then
                                            objLogger.logError "Query SQL_GET_KM_EVENT."
                                            GoTo ErrHandler
                                        End If
                                        
                                        If lngHowMany = 0 Then
                                            objLogger.logError "Impossibile aggiornare i km a vuoto. Manca l'evento inizio carico."
                                        Else
                                            lngKmInizioCarico = varBuffer(0, 0)
                                            
                                            ' Aggiorna i km a vuoto
                                            .DS.Fields("49").NewValue = lngKmInizioCarico - lngKmInizioViaggio
                                            .DS.Fields("49").Updated = True
                                            LogFieldValue .DS.Fields("49")
                                        End If
                                    End If
                                    
                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                        GoTo ErrHandler
                                    End If
                                End If

'                                ' Aggiorna i km totali del viaggio
'                                .DS.Fields("49").NewValue = CLng(varKBuffer(9, i)) - lngKmInizioViaggio
'                                .DS.Fields("49").Updated = True
'                                LogFieldValue .DS.Fields("49")

                            End With

                            If blnUsaDatiGPS Then
                                
                                ' Preleva l'importo pedaggio
                                lngHowMany = 0
                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_IMPORTO_PEDAGGIO, varKBuffer2, lngHowMany, _
                                                CStr(varKBuffer(2, i)), KNOSOS_VIAGGIO_STATO_TERMINATO) Then
                                    objLogger.logError "Query SQL_KNOSOS_GET_IMPORTO_PEDAGGIO."
                                    GoTo ErrHandler
                                End If
    
                                If lngHowMany > 0 Then
                                    sngImportoPedaggio = CSng(Replace(varKBuffer2(0, 0), ".", ","))
                                Else
                                    sngImportoPedaggio = 0
                                End If
    
                                ' Calcola imponibile pedaggio
                                If sngImportoPedaggio > 0 Then
                                    sngImportoPedaggio = Round(sngImportoPedaggio / (1 + (ALIQUOTA_IVA_PEDAGGIO / 100)), 2)
                                End If
    
                                ' Aggiorna il campo importo pedaggio
                                VVT.DS.Fields("46").NewValue = sngImportoPedaggio
                                VVT.DS.Fields("46").Updated = True
                                LogFieldValue VVT.DS.Fields("46")
    
                                ' Aggiorna i km
                                If Not UpdateKM(varIDViaggio, CSng(varKBuffer(9, i))) Then
                                    GoTo ErrHandler
                                End If
                            End If

                        Case Else
                            objLogger.logError "Tipo evento sconosciuto."
                            GoTo ErrHandler

                    End Select

                    ' Aggiorna il viaggio
                    With VVT
'                        If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                            .DS.Fields("27").NewValue = GetKmViaggio(varIDViaggio) + CSng(varKBuffer(9, i))
'                            .DS.Fields("27").Updated = True
'                        End If

                        If blnUsaDatiGPS Then
                            ' Aggiorna il flag trasmissione viaggio
                            .DS.Fields("41").NewValue = "S"
                            .DS.Fields("41").Updated = True
                            LogFieldValue .DS.Fields("41")
                        End If

                        ' Salva
                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                            objLogger.logError "Save Viaggio."
                            GoTo ErrHandler
                        End If
        
                        .DS.DA.EndTrans etCommit
                    End With
                    
                    ' Elimina l'evento KNOSOS
                    objLogger.logInfo "Cancellazione KNOSOS EVENTO."
                    strSQLStatement = Replace(SQL_KNOSOS_DELETE_EVENT, "%fecha%", varKBuffer(3, i))
                    
                    If Not DS.DA.ExecuteCommand(strSQLStatement, CStr(varKBuffer(0, i)), _
                                                CStr(varKBuffer(1, i)), _
                                                CStr(varKBuffer(2, i)), _
                                                CLng(varKBuffer(4, i))) Then
                        objLogger.logError "Cancellazione Evento KNOSOS."
                        GoTo ErrHandler
                    End If

                    objLogger.logInfo "Creazione dell'evento in ResOil."
                    If Not Insert Then
                        objLogger.logError "Insert ResOil evento."
                        GoTo ErrHandler
                    End If

                End If
            Else
                DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
            End If
        Next i
    End If

    ScanEvents_KNOSOS = True
    
    DS.DA.EndTrans etCommit
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logDebug "Fine ScanEvents_KNOSOS senza errori."
    Exit Function

ErrHandler:
    
    DS.DA.EndTrans etRollback
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    
    objLogger.logError "Fine ScanEvents_KNOSOS con ERRORI."

    ScanEvents_KNOSOS = False

End Function
Function ScanEvents_DEUSTO() As Boolean

    Dim i As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim blnUsaDatiGPS As Boolean

    objLogger.logDebug "Inizio ScanEvents_DEUSTO."

    ' Cerchiamo i viaggi che circolano oggi
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_DEUSTO_GET_VIAGGI, varKBuffer, lngHowMany, Date, Date) Then
        objLogger.logError "Query SQL_DEUSTO_GET_VIAGGI."
        GoTo ErrHandler
    End If

    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With

    Set VVP = New CViaggioProdotto
    With VVP
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With

    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With

    objLogger.logInfo "Trovati " & lngHowMany & " automezzi."

    Dim oservice As Object
    Set oservice = CreateObject("VLocalize.VLocalize")
    Dim serviceLoc As String
    Dim timeOut As Integer
    serviceLoc = "http://www.tphweb.es/mbtph/mbws/mbws.asmx"
    timeOut = 5 'Timeout webService in sec.

    'Parametri WebService
    Dim user As String
    Dim Password As String
    Dim idVehicle As Integer

    user = "wstph"
    Password = "swtph"

    If lngHowMany > 0 Then

        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)

            If IsNumeric(varKBuffer(0, i)) Then
            idVehicle = varKBuffer(0, i)

            'Parametri Restituiti
            Dim X As String
            Dim Y As String
            Dim ErrCode As String
            Dim ErrDescription As String
            Dim PosDate As String
            Dim PosHour As String
            Dim Status As String

                objLogger.logInfo "Veicolo " & idVehicle
                If (oservice.GetVehiclePos(user, Password, idVehicle, serviceLoc, timeOut, X, Y, PosDate, PosHour, Status, ErrCode, ErrDescription)) Then
                    objLogger.logInfo "Coordinate X:" + X + " |Y:" + Y + " |PosDate:" + PosDate + " |PosHour:" + PosHour + " |Stato:" + Status
                    strSQLStatement = SQL_SAT_INSERT_POS & DS.DA.MakeSQL(varKBuffer(1, i), dtAlfa) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(PosDate, dtDate) & ","
                    strSQLStatement = strSQLStatement & "'" & Replace(Format(PosHour, "hh:mm:ss"), ".", ":") & "',"
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(X, dtNum) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(Y, dtNum) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL("N", dtAlfa) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(Status, dtNum) & ")"
                    
                    If DS.DA.ExecuteQuery(strSQLStatement) Then
                        objLogger.logInfo "Posizione veicolo inserita correttamente"
                    Else
                        objLogger.logInfo "Errore durante l'inserimento della posizione del veicolo"
                    End If
                Else
                    objLogger.logError "Errore: " & ErrCode & " - " & Err.Description
                End If
            Else
            End If

        Next i
    End If

    ScanEvents_DEUSTO = True

    objLogger.logDebug "Fine ScanEvents_DEUSTO senza errori."
    Exit Function

ErrHandler:

    DS.DA.EndTrans etRollback

    objLogger.logError "Fine ScanEvents_DEUSTO con ERRORI."

    ScanEvents_DEUSTO = False

End Function

'
Function ScanEvents_SMS(ByRef LastID As Variant) As Boolean
    
    Dim i As Long
    Dim j As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim varBufferConsegna As Variant
    Dim varDataEvento As Variant
    Dim varOraEvento As Variant
    Dim varTelAutista As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim varIDAutista As Variant
    Dim varIDSMS As Variant
    Dim varTestoSMS As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim AAU As CAutista
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim strTipoEvento As String
    Dim blnUsaDatiGPS As Boolean

    objLogger.logDebug "Inizio ScanEvents_SMS."
    
    varIDSMS = LastID
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SMS_GET_EVENTS, varKBuffer, lngHowMany, varIDSMS) Then
        objLogger.logError "Query SQL_SMS_GET_EVENTS."
        GoTo ErrHandler
    End If

    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set AAU = New CAutista
    With AAU
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    objLogger.logInfo "Eventi trovati: " & lngHowMany
    
    If lngHowMany > 0 Then

        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)

            ' Inizia transazione
            DS.DA.BeginTransaction
            
            varIDViaggio = ""
            varIDConsegna = ""
            varIDProdotto = ""
            strTipoEvento = vbNullString
            
            objLogger.logInfo "Evento n. " & i + 1
            objLogger.logInfo "-------------------------------------------------"
            objLogger.logInfo "IDSMS         [" & varKBuffer(0, i) & "]"
            objLogger.logInfo "NRTEL         [" & varKBuffer(1, i) & "]"
            objLogger.logInfo "DATASMS       [" & varKBuffer(2, i) & "]"
            objLogger.logInfo "DATALOC       [" & varKBuffer(3, i) & "]"
            objLogger.logInfo "TESTO         [" & varKBuffer(4, i) & "]"
            objLogger.logInfo "-------------------------------------------------"
            
            varIDSMS = varKBuffer(0, i)
            
            ' Cerchiamo l'autista corrispondente al numero di telefono del messaggio che stiamo elaborando
            varTelAutista = varKBuffer(1, i)
            If Len(varTelAutista) = 6 And Left(varTelAutista, 1) = "+" Then
                varTelAutista = Mid(varTelAutista, 2)
            ElseIf Len(varTelAutista) > 6 And Left(varTelAutista, 3) = "+34" Then
                varTelAutista = Mid(varTelAutista, 4)
            End If
            
            varIDAutista = AAU.FindByNrTelefono(varTelAutista)
            If IsEmpty(varIDAutista) Then
                objLogger.logWarning "Numero di telefono " & varTelAutista & " non associato ad alcun Autista."
                GoTo ErrHandler
            End If
            
            varDataEvento = Format(varKBuffer(2, i), "dd/mm/yyyy")
            varOraEvento = Format(varKBuffer(2, i), "hh:nn")
            varTestoSMS = varKBuffer(4, i)
            varTestoSMS = Replace(varTestoSMS, ".", "")
            varTestoSMS = Replace(varTestoSMS, ",", "")
            varTestoSMS = Trim(varTestoSMS)
            If Not IsNumeric(varTestoSMS) Then
                objLogger.logError "Testo SMS non numerico"
                GoTo ErrHandler
            End If
            
            varIDViaggio = VVT.FindBySMSCarico(varIDAutista, varDataEvento, varTestoSMS)
            
            ' Determiniamo se si tratta di un evento di carico o di scarico
            Select Case varIDViaggio
                Case "0"
                    ' Non è un carico, cerchiamo tra gli scarichi
                    varBufferConsegna = VVC.FindBySMSConsegna(varIDAutista, varDataEvento, varTestoSMS)
                    If IsEmpty(varBufferConsegna) Then
                        objLogger.logWarning "Errore durante la ricerca della Consegna, Autista " & varIDAutista & ", Data " & varDataEvento
                        GoTo ErrHandler
                    Else
                        varIDViaggio = varBufferConsegna(0, 0)
                        varIDConsegna = varBufferConsegna(1, 0)
                        strTipoEvento = "INI"
                    End If
                Case "-1"
                    objLogger.logError "Errore durante la ricerca del Viaggio, Autista " & varIDAutista & ", Data " & varDataEvento
                    GoTo ErrHandler
                Case Else
                    ' Si tratta di un carico
                    varIDConsegna = vbNullString
                    varIDProdotto = vbNullString
                    strTipoEvento = "SAL"
            End Select
            
            objLogger.logInfo "RESOil Keys VVTIDVVT = """ & varIDViaggio & _
                                """, VVCIDVVC = """ & varIDConsegna & """"
            
            ' Se non trovo il viaggio , skip
            If varIDViaggio <> "" Then

                ' Creazione dell'evento
                DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
                DS.Fields("0").Updated = True
        
                DS.Fields("1").NewValue = varIDViaggio
                DS.Fields("1").Updated = True
    
                DS.Fields("2").NewValue = varIDConsegna
                DS.Fields("2").Updated = True
    
                DS.Fields("3").NewValue = varIDProdotto
                DS.Fields("3").Updated = True
    
                DS.Fields("5").NewValue = varDataEvento
                DS.Fields("5").Updated = True
    
                DS.Fields("9").NewValue = "K"
                DS.Fields("9").Updated = True
                
                ' Carica il viaggio
                With VVT
                    .DS.Fields.ResetFields
                    .DS.Fields("0").FindValue = varIDViaggio
                
                    If Not .Find Then
                        objLogger.logError "Find Viaggio."
                        GoTo ErrHandler
                    End If
                    
                    blnUsaDatiGPS = True
                
                End With
                    
                        
                Select Case strTipoEvento



                ' -------------------------------------------------------
                ' Evento - Salida de factora (Fine del carico)
                ' -------------------------------------------------------
                Case "SAL"
                    DS.Fields("4").NewValue = "CF"
                    DS.Fields("4").Updated = True

                    objLogger.logInfo "Tipo evento: Fine carico"
                    If Not Insert Then
                        objLogger.logError "Insert ResOil evento."
                        GoTo ErrHandler
                    End If
                    
                    ' Aggiorna lo stato del viaggio (VVTFLSTA)
                    VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
                    VVT.DS.Fields("48").Updated = True
                    LogFieldValue VVT.DS.Fields("48")

                    ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
                    If blnUsaDatiGPS Then

                        ' Aggiorna l'ora inizio viaggio sul viaggio (VVTHRINV)
                        VVT.DS.Fields("7").NewValue = varOraEvento
                        VVT.DS.Fields("7").Updated = True
                        LogFieldValue VVT.DS.Fields("7")
                        
                        ' Aggiorna flag carico completato
                        VVT.DS.Fields("39").NewValue = "S"
                        VVT.DS.Fields("39").Updated = True
                        LogFieldValue VVT.DS.Fields("39")

                        ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
                        If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                            objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                            GoTo ErrHandler
                        End If
                    End If
                    
                    ' -------------------------------------------------------
                    ' Evento - Inicio descarga (Inizio scarico prodotto)
                    ' -------------------------------------------------------
                    Case "INI"
                        objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
                        
                        VVP.DS.Field("Numero Viaggio").FindValue = varIDViaggio
                        VVP.DS.Field("Numero Consegna").FindValue = varIDConsegna
                        If Not VVP.GetList(True) Then
                            GoTo ErrHandler
                        End If
                                                
                        For j = 0 To VVP.DS.HowMany - 1
                            varIDProdotto = VVP.DS.ListItem(0, j)
                            
                            DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
                            DS.Fields("0").Updated = True
                    
                            DS.Fields("1").NewValue = varIDViaggio
                            DS.Fields("1").Updated = True
                
                            DS.Fields("2").NewValue = varIDConsegna
                            DS.Fields("2").Updated = True
                
                            DS.Fields("3").NewValue = varIDProdotto
                            DS.Fields("3").Updated = True
                
                            DS.Fields("4").NewValue = "SI"
                            DS.Fields("4").Updated = True
    
                            DS.Fields("5").NewValue = varDataEvento
                            DS.Fields("5").Updated = True
                
                            DS.Fields("9").NewValue = "K"
                            DS.Fields("9").Updated = True
                            
                            objLogger.logInfo "Creazione dell'evento in ResOil."
                            If Not Insert Then
                                objLogger.logError "Insert ResOil evento."
                                GoTo ErrHandler
                            End If
                            
                            ' Carica la consegna prodotto e aggiorna i dati
                            With VVP
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                                .DS.Fields("2").FindValue = varIDProdotto
                                
                                If Not .Find Then
                                    objLogger.logError "Find Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
                                If .DS.HowMany = 0 Then
                                    objLogger.logError "Prodotto Consegna non trovato"
                                    GoTo ErrHandler
                                End If
                                
                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                .DS.Fields("37").NewValue = "2"
                                .DS.Fields("37").Updated = True
                                LogFieldValue .DS.Fields("37")
    
                                ' Aggiorna il flag Modificato (VVPFLST5)
                                .DS.Field("Stato Riga 5").NewValue = "S"
                                .DS.Field("Stato Riga 5").Updated = True
                                LogFieldValue .DS.Field("Stato Riga 5")
    
                                ' Aggiorna l'ora inizio scarico (VVPHRINC)
                                If blnUsaDatiGPS Then
                                    .DS.Fields("38").NewValue = varOraEvento
                                    .DS.Fields("38").Updated = True
                                    LogFieldValue .DS.Fields("38")
                                End If
    
                                ' Salva
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna Prodotto."
                                    GoTo ErrHandler
                                End If
    
                                .DS.DA.EndTrans etCommit
                            End With
                        Next j
                        
                        ' Carica la consegna e aggiorna i dati
                        With VVC
                            
                            .DS.Fields.ResetFields
                            .DS.Fields("0").FindValue = varIDViaggio
                            .DS.Fields("1").FindValue = varIDConsegna
                    
                            If Not .Find Then
                                objLogger.logError "Find Consegna."
                                GoTo ErrHandler
                            End If
                            
                            ' Se e' l'ultima consegna del viaggio impostiamo il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                            lngHowMany = 0
                            If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT, varBuffer, lngHowMany, _
                                varIDViaggio) Then
                                objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT."
                                GoTo ErrHandler
                            End If
                            If varBuffer(0, 0) = 0 Then
                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                    objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                    GoTo ErrHandler
                                End If
                            End If
                                
                            If blnUsaDatiGPS Then
                                ' Aggiorna l'ora di arrivo
                                .DS.Fields("24").NewValue = varOraEvento
                                .DS.Fields("24").Updated = True
                                LogFieldValue .DS.Fields("24")
                                
                                ' Aggiorna il flag consegna ultimata
                                .DS.Fields("36").NewValue = "S"
                                .DS.Fields("36").Updated = True
                                LogFieldValue .DS.Fields("36")

                                ' Aggiorna lo stato 'Completo'
                                .DS.Fields("44").NewValue = "2"
                                .DS.Fields("44").Updated = True
                                LogFieldValue .DS.Fields("44")
                            End If

                            ' Salva la consegna
                            If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                objLogger.logError "Save Consegna."
                                GoTo ErrHandler
                            End If
                            
                            .DS.DA.EndTrans etCommit
                        End With



                    Case Else
                        objLogger.logError "Tipo evento sconosciuto."
                        GoTo ErrHandler

                End Select

                
                ' Aggiorna il viaggio
                With VVT
                    If .CheckViaggioCompletato(varIDViaggio) Then
                        If Not .SetStato("C") Then
                            objLogger.logError "Set Viaggio Completato."
                            GoTo ErrHandler
                        End If
                    End If
                    
                    If blnUsaDatiGPS Then
                        ' Aggiorna il flag trasmissione viaggio
                        .DS.Fields("41").NewValue = "S"
                        .DS.Fields("41").Updated = True
                        LogFieldValue .DS.Fields("41")
                    End If

                    ' Salva
                    If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                        objLogger.logError "Save Viaggio."
                        GoTo ErrHandler
                    End If
    
                    .DS.DA.EndTrans etCommit
                End With
                
            Else
                DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
            End If
            
            DS.DA.EndTrans etCommit
        Next i
    End If

    ScanEvents_SMS = True
    
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logDebug "Fine ScanEvents_SMS senza errori."
    GoTo EndProc

ErrHandler:
    
    DS.DA.EndTrans etRollback
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    
    objLogger.logError "Fine ScanEvents_SMS con ERRORI."

    ScanEvents_SMS = False

EndProc:
    LastID = varIDSMS
End Function

Public Function AggiornaOrari_SMS(Data As Variant, OraAttuale As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dblMinutiDelay As Double
    Dim varBuffer As Variant
    Dim varConsegnaEffettuata As Variant
    Dim varCaricoEffettuato As Variant
    Dim varOraPrevistaCarico As Variant
    Dim varOraPrevistaConsegna As Variant
    Dim lngHowMany As Long
    Dim lngBufferCount As Long
    Dim lngPointer As Long
    Dim strTargaMotrice As String
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    
'    dblMinutiDelay = 10
    dblMinutiDelay = 60
    
    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    lngHowMany = 0
    If Not IsDate(Data) Then
        GoTo ErrHandler
    End If
    Data = Format(Data, "dd/mm/yyyy")
    If Not DS.DA.QueryCommand(SQL_SMS_GET_LIST_AGG, varBuffer, lngHowMany, Data) Then
        objLogger.logError "Query SQL_SMS_GET_LIST_AGG."
        GoTo ErrHandler
    End If
    If Not IsEmpty(varBuffer) Then
        OraAttuale = Format(OraAttuale, "hh:nn")
        If Not IsDate(OraAttuale) Then
            GoTo ErrHandler
        End If
        lngBufferCount = UBound(varBuffer, 2)
        
'        Dim stra As String
'        For i = 0 To 50
'            stra = Format(i, "00") & " "
'            For j = 0 To 7
'                stra = stra & varBuffer(j, i) & vbTab
'            Next j
'            Debug.Print stra
'        Next i
        
        For i = 0 To lngBufferCount
            strTargaMotrice = varBuffer(0, i)
            
            ' Mi posiziono sulla prima consegna non effettuata
            lngPointer = i - 1
            For j = i To lngBufferCount
                If strTargaMotrice <> varBuffer(0, j) Then
                    Exit For
                End If
                varCaricoEffettuato = varBuffer(3, j)
                varConsegnaEffettuata = varBuffer(4, j)
                If (varConsegnaEffettuata = "S") Then
                    lngPointer = j
                End If
            Next j
'            lngPointer = j
            
            ' Elaboro tutte le consegne non effettuate per la motrice in corso di elaborazione
            If lngPointer <= lngBufferCount Then
                
                If varConsegnaEffettuata = "S" Then
                Else
                
                    For j = (lngPointer + 1) To lngBufferCount
                        If strTargaMotrice <> varBuffer(0, j) Then
                            Exit For
                        End If
                        varOraPrevistaCarico = Format(varBuffer(5, j), "hh:nn")
                        varOraPrevistaConsegna = Format(varBuffer(6, j), "hh:nn")
                        
                        ' Individuo le consegne in ritardo di almeno 50 minuti rispetto all'orario passato come parametro
'                        If (DateDiff("n", varOraPrevistaConsegna, OraAttuale) > 0) Or (DateDiff("n", varOraPrevistaCarico, OraAttuale) > 0) Then
                        If (DateDiff("n", varOraPrevistaConsegna, OraAttuale) >= 50) Or (DateDiff("n", varOraPrevistaCarico, OraAttuale) >= 50) Then
                            For k = j To lngBufferCount
                                If strTargaMotrice <> varBuffer(0, k) Then
                                    Exit For
                                End If
                                strIDViaggio = varBuffer(1, k)
                                strIDConsegna = varBuffer(2, k)
                                
                                DS.DA.BeginTransaction
'                                If IsNull(varCaricoEffettuato) Or (varCaricoEffettuato <> "S") Then
                                If IsNull(varBuffer(3, k)) Or (varBuffer(3, k) <> "S") Then
                                    varOraPrevistaCarico = Format(varBuffer(5, k), "hh:nn")
                                    varOraPrevistaCarico = Format(DateAdd("n", dblMinutiDelay, varOraPrevistaCarico), "hh:nn")
'                                    If DateDiff("n", varOraPrevistaCarico, OraAttuale) > 0 Then
                                    If DateDiff("n", varOraPrevistaCarico, OraAttuale) >= 50 Then
                                        varOraPrevistaCarico = OraAttuale
                                        Debug.Print "C" & " " & strTargaMotrice & " " & strIDViaggio & " " & strIDConsegna & " " & varOraPrevistaCarico
                                        
                                        With VVT
                                            .DS.Fields.ResetFields
                                            .DS.Fields("0").FindValue = strIDViaggio
                                        
                                            If Not .Find Then
                                                objLogger.logError "Find Viaggio."
                                                GoTo ErrHandler
                                            End If
                                            .DS.Field("Ora inizio carico").NewValue = varOraPrevistaCarico
                                            .DS.Field("Ora inizio carico").Updated = True
                                            If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                                objLogger.logError "Save Viaggio."
                                                GoTo ErrHandler
                                            End If
                            
                                            .DS.DA.EndTrans etCommit
                                        End With
                                    End If
                                End If
                                
                                varOraPrevistaConsegna = Format(varBuffer(6, k), "hh:nn")
'                                varOraPrevistaConsegna = Format(DateAdd("n", dblMinutiDelay, varOraPrevistaConsegna), "hh:nn")
'                                If DateDiff("n", varOraPrevistaConsegna, OraAttuale) > 0 Then
                                If DateDiff("n", varOraPrevistaConsegna, OraAttuale) >= 50 Then
                                    varOraPrevistaConsegna = Format(DateAdd("n", dblMinutiDelay, OraAttuale), "hh:nn")
                                    Debug.Print "S" & " "; strTargaMotrice & " " & strIDViaggio & " " & strIDConsegna & " " & varOraPrevistaConsegna
                                    With VVC
                                        
                                        .DS.Fields.ResetFields
                                        .DS.Fields("0").FindValue = strIDViaggio
                                        .DS.Fields("1").FindValue = strIDConsegna
                                
                                        If Not .Find Then
                                            objLogger.logError "Find Consegna."
                                            GoTo ErrHandler
                                        End If
                                        .DS.Field("Orario previsto").NewValue = varOraPrevistaConsegna
                                        .DS.Field("Orario previsto").Updated = True
                                        .DS.Field("Flag stato 4").NewValue = "S"
                                        .DS.Field("Flag stato 4").Updated = True
                                        
                                        ' Salva la consegna
                                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                            objLogger.logError "Save Consegna."
                                            GoTo ErrHandler
                                        End If
                                        
                                        .DS.DA.EndTrans etCommit
                                    End With
                                    
                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, strIDViaggio) Then
                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                        GoTo ErrHandler
                                    End If
                                End If
                                
                                DS.DA.EndTrans etCommit
                                j = k
                            Next k
                        End If
                        lngPointer = j
                    Next j
                End If
                i = lngPointer
            End If
        Next i
    End If
    
    AggiornaOrari_SMS = True
    
    objLogger.logDebug "Fine AggiornaOrari_SMS senza errori."
    GoTo EndProc

ErrHandler:
    
    DS.DA.EndTrans etRollback
    
'    Set VVP = Nothing
'    Set VVT = Nothing
'    Set VVC = Nothing
    
    objLogger.logError "Fine AggiornaOrari_SMS con ERRORI."

    AggiornaOrari_SMS = False

EndProc:
End Function

Public Function GetSelectionEvents(ViaggioId As String, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional MaxRecords As Long = 0) As Boolean
    Dim strSQL As String

    strSQL = Replace(SQL_GET_SELECTION_EVENTS, "%sevidvvt%", ViaggioId)

    GetSelectionEvents = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)

End Function
Public Function UpdateKM(varIDViaggio As Variant, sngKmFineViaggio As Single)
On Error GoTo ErrHandler
    
    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim lngRowsCounter As Long
    Dim VVC As CViaggioConsegna
    Dim lngSequenzaConsegna As Long
    Dim lngKm As Long

    objLogger.logDebug "Inizio UpdateKM"

    ' Preleva i km di fine carico
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CF") Then
        objLogger.logError "Query SQL_GET_KM_EVENT."
        GoTo ErrHandler
    End If
    
    If lngHowMany = 0 Then
        objLogger.logError "Impossibile calcolare il km delle consegne. Manca l'evento fine carico."
        UpdateKM = True
        Exit Function
    End If

    lngKm = CLng(varBuffer(0, 0))
    
    ' Preleva i totali dei km
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SELECT_KM_EVENTS, varBuffer, lngHowMany, varIDViaggio) Then
        objLogger.logError "Mancano le consegne."
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then

        Set VVC = New CViaggioConsegna
        With VVC
            Set .DS.DA = DS.DA
            Set .DS.ErrH = DS.DA.ErrH
        End With
        
        lngSequenzaConsegna = 10
        For lngRowsCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)

            objLogger.logDebug "ResOil Keys, VVTIDVVT = """ & varIDViaggio & _
                                """, VVCIDVVC = """ & varBuffer(0, lngRowsCounter) & """"

            ' Carica la consegna e aggiorna i km
            With VVC

                .DS.Fields.ResetFields
                .DS.Fields("0").FindValue = varIDViaggio
                .DS.Fields("1").FindValue = varBuffer(0, lngRowsCounter)

                If Not .Find Then
                    objLogger.logError "Consegna non trovata."
                    GoTo ErrHandler
                End If

                If .DS.HowMany > 0 Then
                    
                    ' Imposta i km
                    If lngRowsCounter = UBound(varBuffer, 2) Then
                        .DS.Fields("22").NewValue = sngKmFineViaggio - lngKm
                    Else
                        .DS.Fields("22").NewValue = CLng(varBuffer(2, lngRowsCounter)) - lngKm
                    End If
                    .DS.Fields("22").Updated = True
                    LogFieldValue .DS.Fields("22")

                    ' Imposta la sequenza del viaggio
                    .DS.Fields("2").NewValue = Format(lngSequenzaConsegna, "000")
                    .DS.Fields("2").Updated = True
                    LogFieldValue .DS.Fields("2")
    
                    ' Imposta i km a vuoto della consegna
                    If lngRowsCounter = UBound(varBuffer, 2) Then
                        .DS.Fields("45").NewValue = sngKmFineViaggio - CLng(varBuffer(2, lngRowsCounter))
                    Else
                        .DS.Fields("45").NewValue = 0
                    End If
                    .DS.Fields("45").Updated = True
                    LogFieldValue .DS.Fields("45")

                    ' Salva la consegna
                    If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                        objLogger.logError "Save consegna."
                        GoTo ErrHandler
                    End If
    
                    .DS.DA.EndTrans etCommit
    
                End If
            
            End With

            lngKm = CLng(varBuffer(2, lngRowsCounter))
            lngSequenzaConsegna = lngSequenzaConsegna + 10
        Next lngRowsCounter
    End If
    
    objLogger.logDebug "Fine UpdateKM senza errori"
    UpdateKM = True
    Set VVC = Nothing
    Exit Function

ErrHandler:
    objLogger.logInfo "Fine UpdateKM con ERRORI"
    UpdateKM = False
    Set VVC = Nothing
End Function
Public Property Get Logger() As CLogger
    Set Logger = objLogger
End Property
Public Property Set Logger(objLoggerNew As CLogger)
    Set objLogger = objLoggerNew
End Property


Private Function IsEmptyVariantString(varValue As Variant) As Variant
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(varValue) = 0 Then
        IsEmptyVariantString = True
    Else
        IsEmptyVariantString = False
    End If
End Function

