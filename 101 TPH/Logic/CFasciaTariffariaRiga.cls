VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFasciaTariffariaRiga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' 08/07/02  GR  Calcola
Option Explicit


Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_GET_LIST_FASCIA As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Private Function MyRound(Number As Variant, _
                        Optional DecimalPlaces As Integer = 0) As Variant
    On Error GoTo ErrHandler

    Dim N As Single
    Dim lngDec As Long
    Dim Result As Single
    
    lngDec = 10 ^ DecimalPlaces
    
    N = CSng((Number * lngDec) - Int(Number * lngDec))
    Select Case N
        Case Is <= 0.5
            Result = Int(Number * lngDec) / lngDec
        Case Else
            Result = (Int(Number * lngDec) + 1) / lngDec
    End Select

    MyRound = Result
    
    Exit Function
ErrHandler:
    MyRound = Empty
End Function


Public Function CalcolaImporto(IDZona As String, _
                               IDFascia As String, _
                               QScaricata As Single, _
                               QIniziale As Single, _
                               QTotScaricata As Single, _
                               QTotIniziale As Single, _
                               Km As Long, _
                               PrimaVolta As Boolean, _
                               OrdineProgrammato As Boolean, _
                               FatturazioneAttiva As Boolean, _
                               ByRef Tariffa As Single, _
                               ByRef Importo As Single, _
                               ByRef IDValuta As String, _
                               ByRef TipoPercorso As String, _
                               ByRef Anomalia As AnomaliePrefattura, Optional ByVal QScaricataKG As Single, Optional ByVal Zona As String, Optional ByVal TipoligiaProd As String) As Boolean
                               
    On Error GoTo ErrHandler
    
    Static varBuffer As Variant
    Static strIDZonaPrec As String
    Static strIDFasciaPrec As String
    Dim blnResult As Boolean
    Dim blnFound As Boolean
    Dim lngHowMany As Long
    Dim i As Integer
    Dim dblQDifferenza As Single
    Dim dblQuantita As Single
    Dim dblImporto As Single
    Dim dblImportoSca As Single
    Dim dblImportoIni As Single
    Dim dblTariffa As Single
    Dim dblTariffaSca As Single
    Dim dblTariffaIni As Single
    Dim strIDValuta As String
    Dim strIDValutaSca As String
    Dim strIDValutaIni As String
    Dim strTipoPercorso As String
    Dim strTipoPercorsoSca As String
    Dim strTipoPercorsoIni As String
    Dim intAnomalia As Integer
    Dim dblTolleranza As Single
    Dim dblTolleranzaSca As Single
    Dim dblTolleranzaIni As Single
    Dim sngLimiteFasciaInf As Single
    Dim sngTariffaInf As Single
    Dim sngBonusOrdProgrammati As Single
    
    ' --- Leggo i records che formano la fascia tariffaria cercata
    If PrimaVolta Then
        strIDZonaPrec = vbNullString
        strIDFasciaPrec = vbNullString
    End If
    
    If (IDZona <> strIDZonaPrec) Or (IDFascia <> strIDFasciaPrec) Then
        strIDZonaPrec = IDZona
        strIDFasciaPrec = IDFascia
        blnResult = DS.DA.QueryCommand(SQL_GET_LIST_FASCIA, varBuffer, lngHowMany, _
                                       IDZona, IDFascia)
    
        If Not blnResult Then
            GoTo ErrHandler
        End If
        
        If lngHowMany = 0 Then
            GoTo ErrHandler
        End If
    End If
    ' ---
    
    ' --- Mi posiziono sull'intervallo corrispondente alla distanza
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    Do While (i <= UBound(varBuffer, 2)) And (Not blnFound)
        blnFound = (Km >= varBuffer(1, i)) And (Km <= varBuffer(2, i))
        If Not blnFound Then
            i = i + 1
        End If
    Loop
    If Not blnFound Then
        GoTo ErrHandler
    End If
    ' ---
    
    If FatturazioneAttiva Then
        ' --- Se la quantità scaricata è inferiore alla quantità ordinata
        ' --- calcolo l'importo corrispondente
        dblQDifferenza = QScaricata - QIniziale
        If (dblQDifferenza > 0) And (QTotIniziale <= 22000) Then
            blnResult = FindTariffa(QTotScaricata, varBuffer, i, dblTariffaSca, _
                                    dblTolleranzaSca, strIDValutaSca, _
                                    strTipoPercorsoSca, sngLimiteFasciaInf, _
                                    sngTariffaInf, sngBonusOrdProgrammati)
            If Not blnResult Then
                GoTo ErrHandler
            End If
            ' *** (GR) 08/07/02 - Modifica determinazione degli importi
    '        If (QScaricata - QIniziale <= dblTolleranzaSca) _
    '                And (sngLimiteFasciaInf > 0) _
    '                And (QIniziale <= sngLimiteFasciaInf) _
    '                And (QScaricata > sngLimiteFasciaInf) Then
    '            dblTariffaSca = sngTariffaInf
    '            If UCase(strIDValutaSca) = "E" Then
    '                dblImportoSca = MyRound((sngLimiteFasciaInf / 1000) * dblTariffaSca, 2)
    '            Else
    '                dblImportoSca = sngLimiteFasciaInf * dblTariffaSca 'dblTariffaSca
    '            End If
    '        Else
            blnResult = FindTariffa(QTotIniziale, varBuffer, i, dblTariffaIni, _
                                    dblTolleranzaSca, strIDValutaSca, _
                                    strTipoPercorsoSca, sngLimiteFasciaInf, _
                                    sngTariffaInf, sngBonusOrdProgrammati)
            If Not blnResult Then
                GoTo ErrHandler
            End If
            ' *** (GR) 08/07/02 - Fine
            If OrdineProgrammato Then
                dblTariffaSca = dblTariffaSca + sngBonusOrdProgrammati
                dblTariffaIni = dblTariffaIni + sngBonusOrdProgrammati
'            Else
'                dblTariffaSca = dblTariffaSca
'                dblTariffaIni = dblTariffaIni
            End If
            
            If UCase(strIDValutaSca) = "E" Then
                dblImportoSca = MyRound((QScaricata / 1000) * dblTariffaSca, 2)
                dblImportoIni = MyRound((QIniziale / 1000) * dblTariffaIni, 2)
                ' Fueles Canarias
                If Zona = "CAN" And TipoligiaProd = "020" Then
                dblImportoSca = MyRound((QScaricataKG / 1000) * dblTariffaSca, 2)
                End If
            Else
                dblImportoSca = QScaricata * dblTariffaSca
                dblImportoIni = QIniziale * dblTariffaIni
            End If
            
            ' *** (GR) 08/07/02 - Modifica determinazione degli importi
    '        End If
            ' *** (GR) 08/07/02 - Fine
        Else
            blnResult = FindTariffa(QTotScaricata, varBuffer, i, dblTariffaSca, _
                                    dblTolleranzaSca, strIDValutaSca, _
                                    strTipoPercorsoSca, sngLimiteFasciaInf, _
                                    sngTariffaInf, sngBonusOrdProgrammati)
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If OrdineProgrammato Then
                dblTariffaSca = dblTariffaSca + sngBonusOrdProgrammati
                dblTariffaIni = dblTariffaIni + sngBonusOrdProgrammati
            Else
                dblTariffaSca = dblTariffaSca
                dblTariffaIni = dblTariffaIni
            End If
            If UCase(strIDValutaSca) = "E" Then
                dblImportoSca = MyRound((QScaricata / 1000) * dblTariffaSca, 2)
                ' fueles Canarias
                If Zona = "CAN" And TipoligiaProd = "020" Then
                dblImportoSca = MyRound((QScaricataKG / 1000) * dblTariffaSca, 2)
                End If
            Else
                dblImportoSca = QScaricata * dblTariffaSca
            ' fueles Canarias
                If Zona = "CAN" And TipoligiaProd = "020" Then
                dblImportoSca = MyRound((QScaricataKG / 1000) * dblTariffaSca, 2)
                End If
            End If
        End If
        ' ---
    Else
        ' *** Fatturazione passiva per trasportatori
        blnResult = FindTariffa(QTotScaricata, varBuffer, i, dblTariffaSca, _
                                dblTolleranzaSca, strIDValutaSca, _
                                strTipoPercorsoSca, sngLimiteFasciaInf, _
                                sngTariffaInf, sngBonusOrdProgrammati)
        If Not blnResult Then
            GoTo ErrHandler
        End If
        dblTariffaSca = dblTariffaSca
        If UCase(strIDValutaSca) = "E" Then
            dblImportoSca = MyRound((QScaricata / 1000) * dblTariffaSca, 2)
        Else
            dblImportoSca = QScaricata * dblTariffaSca
        End If
    End If
    
    dblTolleranza = dblTolleranzaSca
    strIDValuta = strIDValutaSca
    ' *** (GR) 08/07/02 - Modifica determinazione degli importi
    If dblImportoSca > dblImportoIni Then
        dblImporto = dblImportoSca
        dblTariffa = dblTariffaSca
    Else
        dblImporto = dblImportoIni
        dblTariffa = dblTariffaIni
    End If
    ' *** (GR) 08/07/02 - Fine
    
    strTipoPercorso = strTipoPercorsoSca
    
    intAnomalia = 0
    
    Tariffa = dblTariffa
    Importo = dblImporto
    IDValuta = strIDValuta
    TipoPercorso = strTipoPercorso
    Anomalia = intAnomalia
    CalcolaImporto = True
    Exit Function
ErrHandler:
    Tariffa = 0
    Importo = 0
    IDValuta = vbNullString
    TipoPercorso = vbNullString
    Anomalia = apNoAnomalie
    CalcolaImporto = False
End Function


Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Private Function FindTariffa(Quantita As Single, _
                             Tariffe As Variant, _
                             Row As Integer, _
                             ByRef Tariffa As Single, _
                             ByRef Tolleranza As Single, _
                             ByRef IDValuta As String, _
                             ByRef TipoPercorso As String, _
                             ByRef LimiteFasciaInf As Single, _
                             ByRef TariffaInf As Single, _
                             ByRef BonusOrdiniProgrammati As Single) As Boolean
    On Error GoTo ErrHandler
    
    Dim j As Integer
    Dim blnFound As Boolean
    blnFound = False
    j = 3
'    Do While (j <= 58) And (Not blnFound)
'    Do While (j <= 73) And (Not blnFound)
    Do While (j <= 88) And (Not blnFound)

        Tolleranza = Tariffe(j + 2, Row)
'        blnFound = ((Quantita - Tolleranza) >= Tariffe(j, Row)) And ((Quantita - Tolleranza) <= Tariffe(j + 1, Row))
        blnFound = ((Quantita) >= Tariffe(j, Row)) And ((Quantita) <= Tariffe(j + 1, Row))
        If Not blnFound Then
'            j = j + 5
            j = j + 6
        End If
    Loop
    If Not blnFound Then
        GoTo ErrHandler
    End If
'    If j < 5 Then
    If j < 6 Then
        LimiteFasciaInf = 0
        TariffaInf = 0
    Else
'        LimiteFasciaInf = Tariffe((j - 5) + 1, Row)
'        TariffaInf = Tariffe((j - 5) + 3, Row)
        LimiteFasciaInf = Tariffe((j - 6) + 1, Row)
        TariffaInf = Tariffe((j - 6) + 3, Row)
    End If
'    Tolleranza = Tariffe(j + 2, Row)
    Tariffa = Tariffe(j + 3, Row)
    IDValuta = Tariffe(j + 4, Row)
    If IsNull(Tariffe(j + 5, Row)) Then
        BonusOrdiniProgrammati = 0
    Else
        BonusOrdiniProgrammati = Tariffe(j + 5, Row)
    End If
'    TipoPercorso = Tariffe(63, Row)
'    TipoPercorso = Tariffe(78, Row)
    TipoPercorso = Tariffe(93, Row)
    FindTariffa = True
    Exit Function
ErrHandler:
    FindTariffa = False
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

Public Sub RefreshTimeStamp()
    DS.RefreshTimeStamp SQL_GET_TS
End Sub


Private Sub Class_Initialize()
    Dim i As Integer
    Dim strCounter As String
    
    SQL_FIND = "SELECT tfridazo, tfridtfa, tfrprogd, tfrinfkm, tfrsupkm, " & _
               "       tfrlin01, tfrlsu01, tfrtol01, tfrimp01, tfridv01, tfrimb01, " & _
               "       tfrlin02, tfrlsu02, tfrtol02, tfrimp02, tfridv02, tfrimb02, " & _
               "       tfrlin03, tfrlsu03, tfrtol03, tfrimp03, tfridv03, tfrimb03, " & _
               "       tfrlin04, tfrlsu04, tfrtol04, tfrimp04, tfridv04, tfrimb04, " & _
               "       tfrlin05, tfrlsu05, tfrtol05, tfrimp05, tfridv05, tfrimb05, " & _
               "       tfrlin06, tfrlsu06, tfrtol06, tfrimp06, tfridv06, tfrimb06, " & _
               "       tfrlin07, tfrlsu07, tfrtol07, tfrimp07, tfridv07, tfrimb07, " & _
               "       tfrlin08, tfrlsu08, tfrtol08, tfrimp08, tfridv08, tfrimb08, " & _
               "       tfrlin09, tfrlsu09, tfrtol09, tfrimp09, tfridv09, tfrimb09, " & _
               "       tfrlin10, tfrlsu10, tfrtol10, tfrimp10, tfridv10, tfrimb10, " & _
               "       tfrlin11, tfrlsu11, tfrtol11, tfrimp11, tfridv11, tfrimb11, " & _
               "       tfrlin12, tfrlsu12, tfrtol12, tfrimp12, tfridv12, tfrimb12, " & _
               "       tfrlin13, tfrlsu13, tfrtol13, tfrimp13, tfridv13, tfrimb13, " & _
               "       tfrlin14, tfrlsu14, tfrtol14, tfrimp14, tfridv14, tfrimb14, " & _
               "       tfrlin15, tfrlsu15, tfrtol15, tfrimp15, tfridv15, tfrimb15, " & _
               "       tfrtipop, tfrannul, timestmp FROM taritfr0 WHERE "
    SQL_GETLIST = "SELECT tfrprogd, tfrinfkm, tfrsupkm, tfrtipop, tfrannul " & _
                  "  FROM taritfr0 "
    SQL_UPDATE = "UPDATE taritfr0 SET "
    SQL_INSERT = "INSERT INTO taritfr0 "
    SQL_GET_TS = "SELECT timestmp FROM taritfr0 WHERE "
    SQL_GET_LIST_FASCIA = "SELECT tfrprogd, tfrinfkm, tfrsupkm, " & _
                          "       tfrlin01, tfrlsu01, tfrtol01, tfrimp01, tfridv01, tfrimb01, " & _
                          "       tfrlin02, tfrlsu02, tfrtol02, tfrimp02, tfridv02, tfrimb02, " & _
                          "       tfrlin03, tfrlsu03, tfrtol03, tfrimp03, tfridv03, tfrimb03, " & _
                          "       tfrlin04, tfrlsu04, tfrtol04, tfrimp04, tfridv04, tfrimb04, " & _
                          "       tfrlin05, tfrlsu05, tfrtol05, tfrimp05, tfridv05, tfrimb05, " & _
                          "       tfrlin06, tfrlsu06, tfrtol06, tfrimp06, tfridv06, tfrimb06, " & _
                          "       tfrlin07, tfrlsu07, tfrtol07, tfrimp07, tfridv07, tfrimb07, " & _
                          "       tfrlin08, tfrlsu08, tfrtol08, tfrimp08, tfridv08, tfrimb08, " & _
                          "       tfrlin09, tfrlsu09, tfrtol09, tfrimp09, tfridv09, tfrimb09, " & _
                          "       tfrlin10, tfrlsu10, tfrtol10, tfrimp10, tfridv10, tfrimb10, " & _
                          "       tfrlin11, tfrlsu11, tfrtol11, tfrimp11, tfridv11, tfrimb11, " & _
                          "       tfrlin12, tfrlsu12, tfrtol12, tfrimp12, tfridv12, tfrimb12, " & _
                          "       tfrlin13, tfrlsu13, tfrtol13, tfrimp13, tfridv13, tfrimb13, " & _
                          "       tfrlin14, tfrlsu14, tfrtol14, tfrimp14, tfridv14, tfrimb14, " & _
                          "       tfrlin15, tfrlsu15, tfrtol15, tfrimp15, tfridv15, tfrimb15, " & _
                          "       tfrtipop FROM taritfr0 WHERE tfridazo = ? AND tfridtfa = ? AND tfrannul = 0"
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Zona", "taritfr0.tfridazo", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Codice", "taritfr0.tfridtfa", 0, True, 3, "1"
        .Add "Progressivo", "taritfr0.tfrprogd", 1, True, 3, "2"
        .Add "Limite inferiore Km", "taritfr0.tfrinfkm", 1, True, 7, "3"
        .Add "Limite superiore Km", "taritfr0.tfrsupkm", 1, True, 7, "4"
        For i = 1 To 15  '12
            strCounter = Format$(i, "00")
            .Add "Limite lat. inf. quantità " & strCounter, "taritfr0.tfrlin" & strCounter, 1, , 7, CStr((6 * i) - 1)
            .Add "Limite lat. sup. quantità " & strCounter, "taritfr0.tfrlsu" & strCounter, 1, , 7, CStr((6 * i + 1) - 1)
            .Add "Limite tolleranza quantità " & strCounter, "taritfr0.tfrtol" & strCounter, 1, , 7, CStr((6 * i + 2) - 1)
            .Add "Importo tariffa quantità " & strCounter, "taritfr0.tfrimp" & strCounter, 1, , 13, CStr((6 * i + 3) - 1)
            .Add "Valuta tariffa quantità " & strCounter, "taritfr0.tfridv" & strCounter, 0, , 3, CStr((6 * i + 4) - 1)
            With .Item(CStr((6 * i + 4)) - 1)
                .ForeignTableEntity = "AVA"
                .ForeignTable = "anagava0"
                .ForeignTableFld = "avaidava"
                .ForeignTableDescField = "avadescr"
                .ForeignTableDeleteFld = "avaannul"
            End With
            .Add "Importo bonus programmati " & strCounter, "taritfr0.tfrimb" & strCounter, 1, , 13, CStr((6 * i + 5) - 1)
        Next i
        .Add "Tipo percorso", "taritfr0.tfrtipop", 0, True, 1, "95" '"65"
        .Add "Annullato", "taritfr0.tfrannul", 3, , , "96" '"66"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
    End With
    
    DS.CancelField = "taritfr0.tfrannul"
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


