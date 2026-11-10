VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSatellitareMessaggio"
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

Private SQL_KNOSOS_INSERT_MESSAGE_TOSEND As String
Private SQL_KNOSOS_SELECT_MESSAGE_STATUS_CHANGED As String
Private SQL_KNOSOS_GET_VEICOLI As String
Private SQL_KNOSOS_GET_LAST_DT_MESSAGE_RECEIVED As String
Private SQL_KNOSOS_GET_NEW_MESSAGE_RECEIVED As String
Private SQL_GET_SELECTION As String
Private SQL_KNOSOS_GET_LAST_MESSAGE_COUNT As String
Private SQL_GET_AUTOMEZZI_MESSAGES_COUNT As String

Private Const KNOSOS_MESSAGE_STATUS_PENDIENTE = "0"

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function
Private Function GetDBServerDateTime(blnGMT0 As Boolean, Optional DaysToAdd As Long = 0) As String

    Dim varBuffer As Variant
    Dim lngHowMany As Long

    If DaysToAdd = 0 Then
        If Not DS.DA.QueryCommand("SELECT CONVERT( varchar, " & IIf(blnGMT0, "GETUTCDATE()", "GETDATE()") & ", 21 ) ", varBuffer, lngHowMany) Then
            GetDBServerDateTime = ""
            Exit Function
        End If
    Else
        If Not DS.DA.QueryCommand("SELECT CONVERT( varchar, DATEADD(day, " & DaysToAdd & ", " & IIf(blnGMT0, "GETUTCDATE())", "GETDATE()") & ", 21 ) ", varBuffer, lngHowMany) Then
            GetDBServerDateTime = ""
            Exit Function
        End If
    End If
        
    GetDBServerDateTime = varBuffer(0, 0)

End Function
Public Function Insert() As Boolean

    If Not DS.Fields("8").Updated Or IsNull(DS.Fields("8").NewValue) Then
        DS.Fields("8").NewValue = "N"
        DS.Fields("8").Updated = True
    End If

    Insert = DS.Insert(Me.SQL_INSERT)
    
    Exit Function

End Function
Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean

    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    BeginSave = True
    
End Function

Private Sub Class_Initialize()
    SQL_FIND = "SELECT SMEIDSME, SMEIDAAM, SMEDTINS, SMEFLPRO, SMEIDSTA, " & _
            " SMEMSG, SMEIDEXT, SMEFLREA, SMEOBJ, TIMESTMP FROM VIAGSME0 WHERE "
    SQL_UPDATE = "UPDATE VIAGSME0 SET "
    SQL_INSERT = "INSERT INTO VIAGSME0 "
    SQL_GET_TS = "SELECT TIMESTMP FROM VIAGSME0 WHERE "
    
    SQL_KNOSOS_INSERT_MESSAGE_TOSEND = "INSERT INTO KNOSMET0 (IDVEH, FECHA, ESTA, IDMENSA, MENSA, ASUNTO)" & _
                                    " VALUES( ?, {Ts '%fecha%'}, ?, ?, ?, ?)"

'    SQL_KNOSOS_SELECT_MESSAGE_STATUS_CHANGED = "SELECT SMEIDSME, ESTA" & _
'                                        " FROM VIAGSME0" & _
'                                        " INNER JOIN KNOSMET0 ON VIAGSME0.SMEIDAAM = KNOSMET0.IDVEH" & _
'                                        " AND VIAGSME0.SMEIDEXT = KNOSMET0.IDMENSA" & _
'                                        " WHERE VIAGSME0.SMEFLPRO = ? AND VIAGSME0.SMEFLSAT = ? AND SMEIDSTA <= ? AND SMEIDSTA != ESTA AND FECHA = " & _
'                                        " (SELECT MAX(FECHA) FROM KNOSMET0 A WHERE A.IDVEH = VIAGSME0.SMEIDAAM " & _
'                                        " AND A.IDMENSA = VIAGSME0.SMEIDEXT)" & _
'                                        " GROUP BY VIAGSME0.SMEIDSME, KNOSMET0.ESTA"
    
    SQL_KNOSOS_SELECT_MESSAGE_STATUS_CHANGED = "SELECT SMEIDSME, ESTA" & _
                                        " FROM VIAGSME0" & _
                                        " INNER JOIN KNOSMET0 ON VIAGSME0.SMEIDAAM = KNOSMET0.IDVEH" & _
                                        " AND VIAGSME0.SMEIDEXT = KNOSMET0.IDMENSA" & _
                                        " WHERE VIAGSME0.SMEFLPRO = ? AND VIAGSME0.SMEFLSAT = ? AND SMEIDSTA <= ? AND SMEIDSTA != ESTA" & _
                                        " GROUP BY VIAGSME0.SMEIDSME, KNOSMET0.ESTA"
    
    SQL_KNOSOS_GET_LAST_DT_MESSAGE_RECEIVED = "SELECT ANAGAAM0.AAMIDAAM, MAX(VIAGSME0.SMEDTINS)" & _
                                    " FROM ANAGAAM0 LEFT JOIN VIAGSME0 ON VIAGSME0.SMEIDAAM = ANAGAAM0.AAMIDAAM AND VIAGSME0.SMEFLSAT = 'K' AND VIAGSME0.SMEFLPRO = 'A'" & _
                                    " WHERE ANAGAAM0.AAMFLSAT = 'K'" & _
                                    " GROUP BY ANAGAAM0.AAMIDAAM"

    SQL_KNOSOS_GET_NEW_MESSAGE_RECEIVED = "SELECT IDVEH, CONVERT(varchar, FECHA, 21), IDMENSA, MENSA, ASUNTO" & _
                                    " FROM KNOSMER0 WHERE IDVEH = ? AND FECHA > {Ts '%fecha%'}"

    SQL_GET_SELECTION = "SELECT SMEIDSME, SMEDTINS, AVV_A.AVVDESCR AS SMEFLPRODESC," & _
                " SMEIDAAM, AVV_B.AVVDESCR AS SMEIDSTADESC, SMEOBJ, SMEMSG, AVV_C.AVVDESCR, " & _
                " SMEFLREA, SMEFLPRO, SMEIDSTA, SMEFLSAT" & _
                " FROM VIAGSME0" & _
                " INNER JOIN ANAGAVV AVV_A ON AVV_A.AVVIDTBL = 'VIAGSME0' AND " & _
                " AVV_A.AVVIDFLD = 'SMEFLPRO' AND AVV_A.AVVVALUE = VIAGSME0.SMEFLPRO" & _
                " INNER JOIN ANAGAVV AVV_B ON AVV_B.AVVIDTBL = 'VIAGSME0' AND " & _
                " AVV_B.AVVIDFLD = 'SMEIDSTA' AND AVV_B.AVVVALUE = VIAGSME0.SMEIDSTA" & _
                " INNER JOIN ANAGAVV AVV_C ON AVV_C.AVVIDTBL = 'ANAGAAM0' AND " & _
                " AVV_C.AVVIDFLD = 'AAMFLSAT' AND AVV_C.AVVVALUE = VIAGSME0.SMEFLSAT"
    
    SQL_KNOSOS_GET_VEICOLI = "SELECT AAMIDAAM FROM ANAGAAM0 WHERE AAMFLSAT = 'K'"
    
    SQL_KNOSOS_GET_LAST_MESSAGE_COUNT = "SELECT COUNT(*) FROM VIAGSME0 WHERE SMEIDAAM = ? AND SMEDTINS > ?"
    
    SQL_GET_AUTOMEZZI_MESSAGES_COUNT = "SELECT AAMIDAAM, COUNT(*)" & _
                            " FROM ANAGAAM0 INNER JOIN" & _
                            " VIAGSME0 ON ANAGAAM0.AAMIDAAM = VIAGSME0.SMEIDAAM" & _
                            " WHERE SMEFLPRO = ? AND SMEFLREA = ?" & _
                            " GROUP BY AAMIDAAM"
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Id Messaggio", "VIAGSME0.SMEIDSME", 0, True, 7, "0", , True
        .Add "Targa automezzo", "VIAGSME0.SMEIDAAM", 0, True, 10, "1", , True
        .Add "Data di inserimento", "VIAGSME0.SMEDTINS", 0, True, 23, "2", , True
        .Add "Provenienza", "VIAGSME0.SMEFLPRO", 0, True, 2, "3", , True
        .Add "Stato", "VIAGSME0.SMEIDSTA", 0, True, 2, "4", , True
        .Add "Messaggio", "VIAGSME0.SMEMSG", 0, True, 150, "5", , True
        .Add "Tipo satellitare", "VIAGSME0.SMEFLSAT", 0, True, 2, "6", , True
        .Add "ID Esterno", "VIAGSME0.SMEIDEXT", 1, False, 15, "7", , True
        .Add "Flag lettura", "VIAGSME0.SMEFLREA", 0, False, 2, "8", , True
        .Add "Oggetto", "VIAGSME0.SMEOBJ", 0, True, 30, "9", , True
    End With

    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
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

Private Function KNOSOS_InsertMessageToSend(strTargaAutomezzo As String, _
                                        strDateTime As String, _
                                        strStatus As String, _
                                        sngMessageId As Single, _
                                        strMessageText As String, _
                                        strObjectText As String)

On Error GoTo ErrHandler

    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim strSQLStatement As String
    
    strSQLStatement = Replace(SQL_KNOSOS_INSERT_MESSAGE_TOSEND, "%fecha%", strDateTime)
    
    lngHowMany = 0
    If Not DS.DA.ExecuteCommand(strSQLStatement, _
                            strTargaAutomezzo, _
                            strStatus, _
                            sngMessageId, _
                            strMessageText, _
                            strObjectText) Then
        GoTo ErrHandler
    End If

    KNOSOS_InsertMessageToSend = True
    Exit Function
ErrHandler:
    
    KNOSOS_InsertMessageToSend = False

End Function

Function convertUTCDateToLocalDate(strDate As String) As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strSQLStatement As String

    strSQLStatement = Replace("SELECT CONVERT( nVarchar, DATEADD(hour, DATEDIFF(hour, GETUTCDATE(), GETDATE()), CONVERT(datetime, { Ts '%fecha%' }, 21)), 21)", "%fecha%", strDate)
    If Not DS.DA.QueryCommand(strSQLStatement, varBuffer, lngHowMany) Then
        convertUTCDateToLocalDate = ""
        Exit Function
    End If
        
    convertUTCDateToLocalDate = varBuffer(0, 0)
    
End Function
Function Insert_Message_ToSend(strTargaAutomezzo As String, MessageText As String, ObjectText As String)
On Error GoTo ErrHandler
    
    Dim i As Long
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strTipoSatellitare As String
    Dim strDBDateTime As String
    Dim sngIDExt As Single
    
    Dim AAM As CAutomezzo
    
    ' Inizia transazione
    DS.DA.BeginTransaction

    DS.Fields.ResetFields

    ' ------------------------------------------------------------------------
    ' Carica i dati dell'automezzo
    ' ------------------------------------------------------------------------
    Set AAM = New CAutomezzo
    With AAM
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.ErrH

        .DS.Fields("0").FindValue = strTargaAutomezzo

        If .Find = True And .DS.HowMany > 0 Then
            
            ' Controlla il tipo di satellitare installato
            Select Case .DS.Fields("34").Value
                Case "K"
                    strTipoSatellitare = .DS.Fields("34").Value
                    sngIDExt = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "903", 0)
                    DS.Fields("7").NewValue = sngIDExt
                    DS.Fields("7").Updated = True

                Case Else
                    GoTo ErrHandler
            End Select
        Else
            GoTo ErrHandler
        End If
    End With

    ' Imposta i campi con i valori corretti e inserisce il record
    With DS
        strDBDateTime = GetDBServerDateTime(True)
        
        .Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "902", 0)
        .Fields("0").Updated = True
    
        .Fields("1").NewValue = strTargaAutomezzo
        .Fields("1").Updated = True
    
        .Fields("2").NewValue = convertUTCDateToLocalDate(strDBDateTime)
        .Fields("2").Updated = True
    
        .Fields("3").NewValue = "D"
        .Fields("3").Updated = True
    
        .Fields("4").NewValue = KNOSOS_MESSAGE_STATUS_PENDIENTE
        .Fields("4").Updated = True
    
        .Fields("5").NewValue = Left(MessageText, 150)
        .Fields("5").Updated = True
    
        .Fields("6").NewValue = strTipoSatellitare
        .Fields("6").Updated = True
        
        .Fields("8").NewValue = "N"
        .Fields("8").Updated = True

        .Fields("9").NewValue = ObjectText
        .Fields("9").Updated = True
        
        If Not .Insert(SQL_INSERT) Then
            GoTo ErrHandler
        End If

    End With
    
    Select Case strTipoSatellitare

        ' E' KNOSOS?
        Case "K"
            
            ' Inserisce il messaggio nella tabella KNOSOS
            If Not KNOSOS_InsertMessageToSend(strTargaAutomezzo, strDBDateTime, KNOSOS_MESSAGE_STATUS_PENDIENTE, sngIDExt, MessageText, ObjectText) Then
                GoTo ErrHandler
            End If

        Case Else
            GoTo ErrHandler
    End Select
    
    Set AAM = Nothing
    
    DS.DA.EndTrans etCommit
    
    Insert_Message_ToSend = True
    Exit Function

ErrHandler:
    
    DS.DA.EndTrans etRollback
    
    Set AAM = Nothing
    
    Insert_Message_ToSend = False

End Function
Function Scan_MessageSent_KNOSOS() As Boolean
On Error GoTo ErrHandler

    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim varBufferVeicoliKnosos As Variant
    Dim i As Long
    Dim lngVeicoliCounter As Long
    
    ' Inizia transazione
    DS.DA.BeginTransaction

    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_KNOSOS_SELECT_MESSAGE_STATUS_CHANGED, varBuffer, lngHowMany, "D", "K", "1") Then
        objLogger.logError "Query SQL_KNOSOS_SELECT_MESSAGE_STATUS_CHANGED."
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then
        
        objLogger.logError "Trovati " & lngHowMany & " messaggi da spedire."
        For i = 0 To UBound(varBuffer, 2)
            
            With DS
                .Fields.ResetFields
                
                .Fields("0").FindValue = CStr(varBuffer(0, i))

                If Find Then
                    .Fields("4").NewValue = varBuffer(1, i)
                    .Fields("4").Updated = True

                    If Not BeginSave(True) Then
                        objLogger.logError "Save Messaggio."
                        DS.DA.EndTrans etRollback
                        GoTo ErrHandler
                    Else
                        DS.DA.EndTrans etCommit
                    End If
                
                End If

            End With
        
        Next i
    
    End If

    DS.DA.EndTrans etCommit
    
    Scan_MessageSent_KNOSOS = True
    Exit Function

ErrHandler:
    
    DS.DA.EndTrans etRollback
    
    Scan_MessageSent_KNOSOS = False
    
End Function
Function Scan_MessageReceived_KNOSOS()
On Error GoTo ErrHandler

    Dim lngHowMany As Long
    Dim varBuffer1 As Variant
    Dim varBuffer2 As Variant
    Dim lngAutomezziCounter As Long
    Dim lngMessagesCounter As Long
    Dim strSQLStatement As String
    
    ' Inizia transazione
    DS.DA.BeginTransaction
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_LAST_DT_MESSAGE_RECEIVED, varBuffer1, lngHowMany) Then
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then

        For lngAutomezziCounter = 0 To UBound(varBuffer1, 2)
            
            If IsNull(varBuffer1(1, lngAutomezziCounter)) Then
                strSQLStatement = Replace(SQL_KNOSOS_GET_NEW_MESSAGE_RECEIVED, _
                                    "%fecha%", "1900-01-01 00:00:00.000")
            Else
                strSQLStatement = Replace(SQL_KNOSOS_GET_NEW_MESSAGE_RECEIVED, _
                                    "%fecha%", varBuffer1(1, lngAutomezziCounter))
            End If
            
            lngHowMany = 0
            If Not DS.DA.QueryCommand(strSQLStatement, varBuffer2, lngHowMany, _
                            CStr(varBuffer1(0, lngAutomezziCounter))) Then
                GoTo ErrHandler
            End If

            If lngHowMany > 0 Then

                objLogger.logError "Trovati " & lngHowMany & " messaggi ricevuti."
                
                For lngMessagesCounter = 0 To UBound(varBuffer2, 2)
                    
                    With DS
                        .Fields.ResetFields
                
                        .Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "902", 0)
                        .Fields("0").Updated = True
                    
                        .Fields("1").NewValue = varBuffer2(0, lngMessagesCounter)
                        .Fields("1").Updated = True
                    
                        .Fields("2").NewValue = convertUTCDateToLocalDate(CStr(varBuffer2(1, lngMessagesCounter)))
                        .Fields("2").Updated = True
                    
                        .Fields("3").NewValue = "A"
                        .Fields("3").Updated = True
                    
                        .Fields("4").NewValue = "4"
                        .Fields("4").Updated = True
                    
                        .Fields("5").NewValue = varBuffer2(3, lngMessagesCounter)
                        .Fields("5").Updated = True
                    
                        .Fields("6").NewValue = "K"
                        .Fields("6").Updated = True
                        
                        .Fields("7").NewValue = varBuffer2(2, lngMessagesCounter)
                        .Fields("7").Updated = True
                    
                        .Fields("8").NewValue = "N"
                        .Fields("8").Updated = True
                        
                        .Fields("9").NewValue = varBuffer2(4, lngMessagesCounter)
                        .Fields("9").Updated = True
                        
                        If Not .Insert(SQL_INSERT) Then
                            GoTo ErrHandler
                        End If
                    End With
                            
                Next lngMessagesCounter
            
            End If
        
        Next lngAutomezziCounter
    
    End If

    DS.DA.EndTrans etCommit
    
    Scan_MessageReceived_KNOSOS = True
    Exit Function

ErrHandler:
    
    DS.DA.EndTrans etRollback
    
    Scan_MessageReceived_KNOSOS = False
    
End Function
Public Function GetSelection(strAutomezzoId As String, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    Dim strSQL As String

    strSQL = SQL_GET_SELECTION

    DS.Fields("1").FindValue = strAutomezzoId

    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)
    End If

End Function
Public Function GetSelectionMessaggiAutomezzoOrderDesc(strAutomezzoId As String, lngNDays As Long, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional MaxRecords As Long = 0) As Boolean
On Error GoTo ErrHandler
    Dim strSQL As String
    Dim strCond As String
    Dim blnWhere As String
    Dim strCondDays As String
    Dim SME As CSatellitareMessaggio

    Set SME = New CSatellitareMessaggio
    Set SME.DS.DA = DS.DA
    Set SME.DS.ErrH = DS.ErrH

    DS.Fields("1").FindValue = strAutomezzoId
    
    blnWhere = False
    
    strSQL = SQL_GET_SELECTION
    strCond = DS.MakeKeyCond

    If strCond <> "" Then
        strSQL = strSQL & " WHERE " & strCond
        blnWhere = True
    End If

    If lngNDays > -1 Then
        Dim strDate As String
        
        strDate = "'" & Format(DateAdd("d", -lngNDays, Now()), "yyyy-mm-dd 00:00:00.000") & "'"
        strSQL = strSQL & IIf(blnWhere, " AND ", " WHERE ") & DS.Fields("2").FieldName & " >= " & strDate
    End If
    
    strSQL = strSQL & " ORDER BY SMEDTINS DESC"

    DS.Fields.ResetFields
    
    GetSelectionMessaggiAutomezzoOrderDesc = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)

    Set SME = Nothing
    Exit Function
ErrHandler:
    GetSelectionMessaggiAutomezzoOrderDesc = False
    Set SME = Nothing
End Function

Public Function GetSelectionMessaggiOrderDesc(lngNDays As Long, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional MaxRecords As Long = 0) As Boolean
On Error GoTo ErrHandler
    Dim strSQL As String
    Dim strCond As String
    Dim blnWhere As String
    Dim strCondDays As String
    Dim SME As CSatellitareMessaggio

    Set SME = New CSatellitareMessaggio
    Set SME.DS.DA = DS.DA
    Set SME.DS.ErrH = DS.ErrH
    
    blnWhere = False

    strSQL = SQL_GET_SELECTION
    strCond = DS.MakeKeyCond

    If strCond <> "" Then
        strSQL = strSQL & " WHERE " & strCond
        blnWhere = True
    End If

    If lngNDays > -1 Then
        Dim strDate As String
        
        strDate = "'" & Format(DateAdd("d", -lngNDays, Now()), "yyyy-mm-dd 00:00:00.000") & "'"
        strSQL = strSQL & IIf(blnWhere, " AND ", " WHERE ") & DS.Fields("2").FieldName & " >= " & strDate
    End If
    
    strSQL = strSQL & " ORDER BY SMEDTINS DESC"

    DS.Fields.ResetFields
    
    GetSelectionMessaggiOrderDesc = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)

    Set SME = Nothing
    Exit Function
ErrHandler:
    GetSelectionMessaggiOrderDesc = False
    Set SME = Nothing
End Function

Public Function GetAutomezziMessaggesNotRead(varBuffer As Variant) As Boolean

    Dim lngHowMany As Long
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_GET_AUTOMEZZI_MESSAGES_COUNT, varBuffer, lngHowMany, _
            "A", "N") Then
        GoTo ErrHandler
    End If
    
    GetAutomezziMessaggesNotRead = True
    Exit Function

ErrHandler:
    GetAutomezziMessaggesNotRead = False
End Function

Public Function SetReadMessage(strMessageId As String) As Boolean
On Error GoTo ErrHandler
    
    Dim SME As CSatellitareMessaggio
    
    Set SME = New CSatellitareMessaggio
    
    With SME.DS
        Set .DA = DS.DA
        Set .ErrH = DS.DA.ErrH
    End With
    
    With SME
    
        .DS.Fields("0").FindValue = strMessageId
        If Not .Find Then
            GoTo ErrHandler
        End If
        
        If .DS.Fields("3").Value = "A" And .DS.Fields("7").Value = "N" Then
            .DS.Fields("8").NewValue = "L"
            .DS.Fields("8").Updated = True
        
            If Not .BeginSave(True) Then
                GoTo ErrHandler
            End If
        
            .DS.DA.EndTrans etCommit
        Else
            GoTo ErrHandler
        End If
    
    End With
        
    Set SME = Nothing
    
    SetReadMessage = True
    Exit Function

ErrHandler:
    Set SME = Nothing
    SetReadMessage = False
End Function
Public Property Get Logger() As CLogger
    Set Logger = objLogger
End Property
Public Property Set Logger(objLoggerNew As CLogger)
    Set objLogger = objLoggerNew
End Property

