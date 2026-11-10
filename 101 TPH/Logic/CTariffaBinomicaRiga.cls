VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTariffaBinomicaRiga"
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


Public Function CalcolaImporto(IDDeposito As String, _
                               ByRef IDFascia As String, _
                               QScaricataL As Single, _
                               QTotQScaricataViaggio As Single, QOrdinataViaggioL As Single, QTotScaricataLViaggioMultiDestino As Single, _
                               QTotScaricataL As Single, _
                               QOrdinataViaggioLMultiDestino As Single, _
                               QScaricataKG As Single, _
                               QTotScaricataLViaggioDestinazione As Single, _
                               QTotOrdinataL As Single, _
                               Km As Long, _
                               PrimaVolta As Boolean, _
                               FlagKilometrica As Boolean, _
                               FlagFestivo As Boolean, _
                               NumConsegne As Long, _
                               ConsegnaPiùDistante As Boolean, _
                               UltimoScarico As Boolean, _
                               TipoCisterna As Variant, _
                               TipoTariffa As String, _
                               TipologiaProdotto As String, _
                               IDDestinazione As String, FlagMultiDestino As Boolean, IDVettore As String, FlagResponsabilità As String, QOCTNota2 As String, _
                               ByRef Tariffa As Single, _
                               ByRef Importo As Single, _
                               ByRef IDValuta As String, _
                               ByRef TipoPercorso As String, _
                               ByRef Anomalia As AnomaliePrefattura) As Boolean
    On Error GoTo ErrHandler
    
    Static varBuffer As Variant
    Static strIDDepositoPrec As String
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
    Dim sngTariffaDistanza As Single
    
    ' --- Leggo i records che formano la fascia tariffaria cercata
    If PrimaVolta Then
        strIDDepositoPrec = vbNullString
        strIDFasciaPrec = vbNullString
    End If
    
    If (IDDeposito <> strIDDepositoPrec) Or (IDFascia <> strIDFasciaPrec) Then
        strIDDepositoPrec = IDDeposito
        strIDFasciaPrec = IDFascia
        
        blnResult = DS.DA.QueryCommand(SQL_GET_LIST_FASCIA, varBuffer, lngHowMany, _
                                       IDDeposito, IDFascia)
    
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
    
    
    If QTotScaricataL > 32000 Then
        If FlagMultiDestino Then
            QTotScaricataL = QTotScaricataLViaggioMultiDestino
        Else
            QTotScaricataL = QTotScaricataLViaggioDestinazione 'QTotScaricataLViaggio
        End If
    End If
    
    If FlagMultiDestino Then
        If (QTotScaricataLViaggioMultiDestino < QOrdinataViaggioLMultiDestino) And (QOrdinataViaggioLMultiDestino >= QTotOrdinataL) Then
            QTotScaricataL = QOrdinataViaggioLMultiDestino
        ElseIf QTotOrdinataL > 32000 Then
            QTotScaricataL = QTotScaricataLViaggioMultiDestino
        End If
    Else
        ' JG 27/1/12 quito que compare la cantidad la cantidad total de los pedido con la del pedido servido.
        'If (QTotScaricataLViaggioDestinazione < QOrdinataViaggioL) And (QOrdinataViaggioL >= QTotOrdinataL) Then
        If (QTotScaricataLViaggioDestinazione < QOrdinataViaggioL) Then
            QTotScaricataL = QOrdinataViaggioL
        ElseIf QTotOrdinataL > 32000 Then
            QTotScaricataL = QTotScaricataLViaggioDestinazione
        ElseIf QTotOrdinataL <= 32000 And FlagResponsabilità = "C" Then
            QTotScaricataL = QTotScaricataLViaggioDestinazione
        End If
    End If
    
    ' Cálculo mínimo de 23 tm
    ' If ((TipologiaProdotto = "020") Or (TipologiaProdotto = "040")) And (QScaricataKG < 23000) And (TipoTariffa = "P") Then
    
    If ((TipologiaProdotto = "020") Or (TipologiaProdotto = "040")) And (QScaricataKG < 23000) Then
        QScaricataKG = 23000
    End If
    
    blnResult = FindTariffa(QTotScaricataL, varBuffer, i, dblTariffa, _
                            dblTolleranza, strIDValuta, _
                            strTipoPercorso, sngLimiteFasciaInf, _
                            sngTariffaInf, sngTariffaDistanza)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    If Not FlagKilometrica Then
        
'        If (IDDestinazione = "CPR0905783") And (TipoTariffa <> "P") Then
'
'        ' POne tarifas especiales
'
''            dblTariffa = 22.86923 '22.53126 ' 22.98276 '22.71922 '23.00681 '22.24815 '21.61064 '20.97306
''            dblImporto = dblTariffa * Format(QScaricataL / 1000, "0.000")
''        ElseIf (IDDeposito = "030") And (TipoTariffa <> "P") And (QTotOrdinataL < 15000) Then
''            dblTariffa = 0.01396 '0.01375 '0.01403 '0.0138645  '0.01404 '0.01358 '0.01319 '0.0128
''            dblImporto = dblTariffa * QScaricataL
'        Else


            '-- JG Control viaje inferior a 9500 litros/cambio tarifa QTotOrdinataL
            '             If QTotScaricataLViaggioDestinazione < 9500 And QTotScaricataLViaggioMultiDestino < 9500 And (TipoTariffa = "A") And IDDeposito <> "260" And (TipologiaProdotto <> "020") And (TipologiaProdotto <> "040") And (Left(QOCTNota2, 12) = "910101910101" Or (Left(IDDestinazione, 2) = "EA" Or Left(IDDestinazione, 2) = "OT")) Then
            
            
            
            ' Controlar zona de Barcelona solo si son menos de 4950 (IDDeposito = "666" Or IDDeposito = "121" Or IDDeposito = "D01" Or IDDeposito = "N0")
            Dim Zona As String
            Zona = ""
            If (IDDeposito = "666" Or IDDeposito = "121" Or IDDeposito = "D01" Or IDDeposito = "N0") Then
            Zona = "CAT"
            End If
            If (IDDeposito = "401" Or IDDeposito = "402" Or IDDeposito = "404" Or IDDeposito = "405" Or IDDeposito = "406" Or IDDeposito = "407" Or IDDeposito = "408" Or IDDeposito = "409" Or IDDeposito = "410") Then
            Zona = "CAN"
            End If
            
            If QTotScaricataLViaggioDestinazione < 4950 And QTotOrdinataL < 4950 And (TipoTariffa = "A") And Zona = "CAT" And (TipologiaProdotto <> "020") And (TipologiaProdotto <> "040") And (Left(QOCTNota2, 12) = "910101910101" Or (Left(IDDestinazione, 2) = "EA" Or Left(IDDestinazione, 2) = "OT")) Then
            IDFascia = "LM"
            strIDValuta = "LM"
            Else
    
                ' if anterior sin control zona barcelona  If QTotScaricataLViaggioDestinazione < 9500 And QTotOrdinataL < 9500 And (TipoTariffa = "A") And IDDeposito <> "260" And (TipologiaProdotto <> "020") And (TipologiaProdotto <> "040") And (Left(QOCTNota2, 12) = "910101910101" Or (Left(IDDestinazione, 2) = "EA" Or Left(IDDestinazione, 2) = "OT")) Then
    
                If QTotScaricataLViaggioDestinazione < 9500 And QTotOrdinataL < 9500 And (TipoTariffa = "A") And IDDeposito <> "260" And Zona <> "CAT" And Zona <> "CAN" And (TipologiaProdotto <> "020") And (TipologiaProdotto <> "040") And (Left(QOCTNota2, 12) = "910101910101" Or (Left(IDDestinazione, 2) = "EA" Or Left(IDDestinazione, 2) = "OT")) Then
                IDFascia = "LM"
                strIDValuta = "LM"
                Else
                
                'Control Tarifa Ceuta
                If IDDeposito = "260" Then
                IDFascia = "CE"
                strIDValuta = "CE"
                Else
                dblImporto = MyRound((QScaricataKG / 1000) * dblTariffa + (QScaricataKG / 1000) * Km * sngTariffaDistanza, 2)
                End If
                End If
            End If
            

'        End If
        
'        If Len(TipoCisterna) > 0 Then
'            If (TipoTariffa = "P") And (QTotScaricataL < 3000) Then
'                dblImporto = dblImporto + 0.006 * QScaricataL
'            End If
'        End If
    Else
'        If ConsegnaPiùDistante And UltimoScarico Then
        If UltimoScarico Then
            
            
            If (IDVettore <> "A10") Or (Km > 300) Then
                dblImporto = MyRound(Km * sngTariffaDistanza, 2)
            Else
                dblImporto = sngTariffaDistanza
            End If
            dblTariffa = sngTariffaDistanza
'            If Km <= 25 Then
'                dblImporto = dblImporto + 48.08
'            ElseIf Km <= 50 Then
'                dblImporto = dblImporto + 42.07
'            ElseIf Km <= 100 Then
'                dblImporto = dblImporto + 30.05
'            End If
'            If NumConsegne > 1 Then
'                dblImporto = dblImporto + ((NumConsegne - 1) * 12)
'            End If
'            If FlagFestivo Then
'                dblImporto = dblImporto + 30
'            End If
        Else
            Tariffa = 0
            Importo = 0
            IDValuta = vbNullString
            TipoPercorso = vbNullString
            Anomalia = apNoAnomalie
            CalcolaImporto = True
            Exit Function
        End If
    End If
    
    dblTolleranza = dblTolleranza
    strIDValuta = strIDValuta
    
    strTipoPercorso = strTipoPercorso
    
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
                             ByRef TariffaDistanza As Single) As Boolean
    On Error GoTo ErrHandler
    
    Dim j As Integer
    Dim blnFound As Boolean
        
    j = 3
    Do While (j <= 88) And (Not blnFound)
        Tolleranza = Tariffe(j + 2, Row)
        blnFound = ((Quantita) >= Tariffe(j, Row)) And ((Quantita) <= Tariffe(j + 1, Row))
        If Not blnFound Then
            j = j + 6
        End If
    Loop
    If Not blnFound Then
        GoTo ErrHandler
    End If
    If j < 6 Then
        LimiteFasciaInf = 0
        TariffaInf = 0
    Else
        LimiteFasciaInf = Tariffe((j - 6) + 1, Row)
        TariffaInf = Tariffe((j - 6) + 3, Row)
    End If
    Tariffa = Tariffe(j + 3, Row)
    IDValuta = Tariffe(j + 4, Row)
    If IsNull(Tariffe(j + 5, Row)) Then
        TariffaDistanza = 0
    Else
        TariffaDistanza = Tariffe(j + 5, Row)
    End If
    If Len(Tariffe(93, Row)) > 0 Then
        TipoPercorso = Tariffe(93, Row)
    Else
        TipoPercorso = vbNullString
    End If
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
    
    SQL_FIND = "SELECT tbridade, tbridtfb, tbrprogd, tbrinfkm, tbrsupkm, " & _
               "       tbrlin01, tbrlsu01, tbrtol01, tbrimp01, tbridv01, tbrimb01, " & _
               "       tbrlin02, tbrlsu02, tbrtol02, tbrimp02, tbridv02, tbrimb02, " & _
               "       tbrlin03, tbrlsu03, tbrtol03, tbrimp03, tbridv03, tbrimb03, " & _
               "       tbrlin04, tbrlsu04, tbrtol04, tbrimp04, tbridv04, tbrimb04, " & _
               "       tbrlin05, tbrlsu05, tbrtol05, tbrimp05, tbridv05, tbrimb05, " & _
               "       tbrlin06, tbrlsu06, tbrtol06, tbrimp06, tbridv06, tbrimb06, " & _
               "       tbrlin07, tbrlsu07, tbrtol07, tbrimp07, tbridv07, tbrimb07, " & _
               "       tbrlin08, tbrlsu08, tbrtol08, tbrimp08, tbridv08, tbrimb08, " & _
               "       tbrlin09, tbrlsu09, tbrtol09, tbrimp09, tbridv09, tbrimb09, " & _
               "       tbrlin10, tbrlsu10, tbrtol10, tbrimp10, tbridv10, tbrimb10, " & _
               "       tbrlin11, tbrlsu11, tbrtol11, tbrimp11, tbridv11, tbrimb11, " & _
               "       tbrlin12, tbrlsu12, tbrtol12, tbrimp12, tbridv12, tbrimb12, " & _
               "       tbrlin13, tbrlsu13, tbrtol13, tbrimp13, tbridv13, tbrimb13, " & _
               "       tbrlin14, tbrlsu14, tbrtol14, tbrimp14, tbridv14, tbrimb14, " & _
               "       tbrlin15, tbrlsu15, tbrtol15, tbrimp15, tbridv15, tbrimb15, " & _
               "       tbrtipop, tbrannul, timestmp FROM taritbr0 WHERE "
    SQL_GETLIST = "SELECT tbrprogd, tbrinfkm, tbrsupkm, tbrtipop, tbrannul " & _
                  "  FROM taritbr0 "
    SQL_UPDATE = "UPDATE taritbr0 SET "
    SQL_INSERT = "INSERT INTO taritbr0 "
    SQL_GET_TS = "SELECT timestmp FROM taritbr0 WHERE "
    SQL_GET_LIST_FASCIA = "SELECT tbrprogd, tbrinfkm, tbrsupkm, " & _
                          "       tbrlin01, tbrlsu01, tbrtol01, tbrimp01, tbridv01, tbrimb01, " & _
                          "       tbrlin02, tbrlsu02, tbrtol02, tbrimp02, tbridv02, tbrimb02, " & _
                          "       tbrlin03, tbrlsu03, tbrtol03, tbrimp03, tbridv03, tbrimb03, " & _
                          "       tbrlin04, tbrlsu04, tbrtol04, tbrimp04, tbridv04, tbrimb04, " & _
                          "       tbrlin05, tbrlsu05, tbrtol05, tbrimp05, tbridv05, tbrimb05, " & _
                          "       tbrlin06, tbrlsu06, tbrtol06, tbrimp06, tbridv06, tbrimb06, " & _
                          "       tbrlin07, tbrlsu07, tbrtol07, tbrimp07, tbridv07, tbrimb07, " & _
                          "       tbrlin08, tbrlsu08, tbrtol08, tbrimp08, tbridv08, tbrimb08, " & _
                          "       tbrlin09, tbrlsu09, tbrtol09, tbrimp09, tbridv09, tbrimb09, " & _
                          "       tbrlin10, tbrlsu10, tbrtol10, tbrimp10, tbridv10, tbrimb10, " & _
                          "       tbrlin11, tbrlsu11, tbrtol11, tbrimp11, tbridv11, tbrimb11, " & _
                          "       tbrlin12, tbrlsu12, tbrtol12, tbrimp12, tbridv12, tbrimb12, " & _
                          "       tbrlin13, tbrlsu13, tbrtol13, tbrimp13, tbridv13, tbrimb13, " & _
                          "       tbrlin14, tbrlsu14, tbrtol14, tbrimp14, tbridv14, tbrimb14, " & _
                          "       tbrlin15, tbrlsu15, tbrtol15, tbrimp15, tbridv15, tbrimb15, " & _
                          "       tbrtipop FROM taritbr0 WHERE tbridade = ? AND tbridtfb = ? AND tbrannul = 0"
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Zona", "taritbr0.tbridade", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Codice", "taritbr0.tbridtfb", 0, True, 3, "1"
        .Add "Progressivo", "taritbr0.tbrprogd", 1, True, 3, "2"
        .Add "Limite inferiore Km", "taritbr0.tbrinfkm", 1, True, 7, "3"
        .Add "Limite superiore Km", "taritbr0.tbrsupkm", 1, True, 7, "4"
        For i = 1 To 15  '12
            strCounter = Format$(i, "00")
            .Add "Limite lat. inf. quantità " & strCounter, "taritbr0.tbrlin" & strCounter, 1, , 7, CStr((6 * i) - 1)
            .Add "Limite lat. sup. quantità " & strCounter, "taritbr0.tbrlsu" & strCounter, 1, , 7, CStr((6 * i + 1) - 1)
            .Add "Limite tolleranza quantità " & strCounter, "taritbr0.tbrtol" & strCounter, 1, , 7, CStr((6 * i + 2) - 1)
            .Add "Importo tariffa quantità " & strCounter, "taritbr0.tbrimp" & strCounter, 1, , 13, CStr((6 * i + 3) - 1)
            .Add "Valuta tariffa quantità " & strCounter, "taritbr0.tbridv" & strCounter, 0, , 3, CStr((6 * i + 4) - 1)
            With .Item(CStr((6 * i + 4)) - 1)
                .ForeignTableEntity = "AVA"
                .ForeignTable = "anagava0"
                .ForeignTableFld = "avaidava"
                .ForeignTableDescField = "avadescr"
                .ForeignTableDeleteFld = "avaannul"
            End With
            .Add "Importo tariffa distanza " & strCounter, "taritbr0.tbrimb" & strCounter, 1, , 13, CStr((6 * i + 5) - 1)
        Next i
        .Add "Tipo percorso", "taritbr0.tbrtipop", 0, True, 1, "95" '"65"
        .Add "Annullato", "taritbr0.tbrannul", 3, , , "96" '"66"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
    End With
    
    DS.CancelField = "taritbr0.tbrannul"
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


