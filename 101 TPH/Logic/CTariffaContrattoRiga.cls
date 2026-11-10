VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTariffaContrattoRiga"
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


Public Function CalcolaImporto(IDCliente As String, _
                               IDContratto As String, _
                               IDAddebito As String, _
                               IDFascia As String, _
                               QScaricataL As Single, _
                               QScaricataKG As Single, _
                               Km As Long, _
                               PrimaVolta As Boolean, _
                               FlagKilometrica As Boolean, _
                               FlagFestivo As Boolean, _
                               NumConsegne As Long, _
                               ConsegnaPiùDistante As Boolean, _
                               UltimoScarico As Boolean, _
                               FlagPompa As Boolean, _
                               TipoTariffa As String, _
                               IDTipoCalcolo As String, _
                               ByRef Tariffa As Single, _
                               ByRef Importo As Single, _
                               ByRef IDValuta As String, _
                               ByRef TipoPercorso As String, _
                               ByRef Anomalia As AnomaliePrefattura) As Boolean
    On Error GoTo ErrHandler
    
    Static varBuffer As Variant
    Static strIDClientePrec As String
    Static strIDContrattoPrec As String
    Static strIDFasciaPrec As String
    Static strIDAddebitoPrec As String
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
        strIDClientePrec = vbNullString
        strIDContrattoPrec = vbNullString
        strIDFasciaPrec = vbNullString
        strIDAddebitoPrec = vbNullString
    End If
    
    If (IDCliente <> strIDClientePrec) Or (IDContratto <> strIDContrattoPrec) Or (IDFascia <> strIDFasciaPrec) Then
        strIDClientePrec = IDCliente
        strIDContrattoPrec = IDContratto
        strIDAddebitoPrec = IDAddebito
        strIDFasciaPrec = IDFascia
        
        blnResult = DS.DA.QueryCommand(SQL_GET_LIST_FASCIA, varBuffer, lngHowMany, _
                                       IDCliente, IDContratto, IDAddebito, IDFascia)
    
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
    
    blnResult = FindTariffa(QScaricataL, varBuffer, i, dblTariffa, _
                            dblTolleranza, strIDValuta, _
                            strTipoPercorso, sngLimiteFasciaInf, _
                            sngTariffaInf, sngTariffaDistanza)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    
    If QScaricataL > 0 Then
        Select Case IDTipoCalcolo
            Case "02"
                If IsNumeric(sngTariffaDistanza) And sngTariffaDistanza > 0 Then
                    dblTariffa = sngTariffaDistanza
                    dblImporto = sngTariffaDistanza
                Else
                    dblImporto = dblTariffa
                End If
            Case "03"
                If Km <> 0 Then
                    If IsNumeric(sngTariffaDistanza) And sngTariffaDistanza > 0 Then
                        dblTariffa = sngTariffaDistanza
                        dblImporto = sngTariffaDistanza
                    Else
                        dblImporto = MyRound(Km * dblTariffa, 2)
                    End If
                End If
            Case "22"
                If IsNumeric(sngTariffaDistanza) And sngTariffaDistanza > 0 Then
                    dblTariffa = sngTariffaDistanza
                End If
                dblImporto = MyRound(NumConsegne * dblTariffa, 2)
            Case "91"
                dblImporto = MyRound((QScaricataL / 1000) * dblTariffa, 2)
            Case "98"
                If Km <> 0 Then
                    dblImporto = MyRound((QScaricataL / 1000) * dblTariffa, 2)
                End If
            Case Else
                GoTo ErrHandler
        End Select
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
    
    SQL_FIND = "SELECT tcridacl, tcridacn, tcridadd, tcridtfc, tcrprogd, tcrinfkm, tcrsupkm, " & _
               "       tcrlin01, tcrlsu01, tcrtol01, tcrimp01, tcridv01, tcrimb01, " & _
               "       tcrlin02, tcrlsu02, tcrtol02, tcrimp02, tcridv02, tcrimb02, " & _
               "       tcrlin03, tcrlsu03, tcrtol03, tcrimp03, tcridv03, tcrimb03, " & _
               "       tcrlin04, tcrlsu04, tcrtol04, tcrimp04, tcridv04, tcrimb04, " & _
               "       tcrlin05, tcrlsu05, tcrtol05, tcrimp05, tcridv05, tcrimb05, " & _
               "       tcrlin06, tcrlsu06, tcrtol06, tcrimp06, tcridv06, tcrimb06, " & _
               "       tcrlin07, tcrlsu07, tcrtol07, tcrimp07, tcridv07, tcrimb07, " & _
               "       tcrlin08, tcrlsu08, tcrtol08, tcrimp08, tcridv08, tcrimb08, " & _
               "       tcrlin09, tcrlsu09, tcrtol09, tcrimp09, tcridv09, tcrimb09, " & _
               "       tcrlin10, tcrlsu10, tcrtol10, tcrimp10, tcridv10, tcrimb10, " & _
               "       tcrlin11, tcrlsu11, tcrtol11, tcrimp11, tcridv11, tcrimb11, " & _
               "       tcrlin12, tcrlsu12, tcrtol12, tcrimp12, tcridv12, tcrimb12, " & _
               "       tcrlin13, tcrlsu13, tcrtol13, tcrimp13, tcridv13, tcrimb13, " & _
               "       tcrlin14, tcrlsu14, tcrtol14, tcrimp14, tcridv14, tcrimb14, " & _
               "       tcrlin15, tcrlsu15, tcrtol15, tcrimp15, tcridv15, tcrimb15, " & _
               "       tcrtipop, tcrannul, timestmp FROM taritcr0 WHERE "
    SQL_GETLIST = "SELECT tcrprogd, tcrinfkm, tcrsupkm, tcrtipop, tcrannul " & _
                  "  FROM taritcr0 "
    SQL_UPDATE = "UPDATE taritcr0 SET "
    SQL_INSERT = "INSERT INTO taritcr0 "
    SQL_GET_TS = "SELECT timestmp FROM taritcr0 WHERE "
    SQL_GET_LIST_FASCIA = "SELECT tcrprogd, tcrinfkm, tcrsupkm, " & _
                          "       tcrlin01, tcrlsu01, tcrtol01, tcrimp01, tcridv01, tcrimb01, " & _
                          "       tcrlin02, tcrlsu02, tcrtol02, tcrimp02, tcridv02, tcrimb02, " & _
                          "       tcrlin03, tcrlsu03, tcrtol03, tcrimp03, tcridv03, tcrimb03, " & _
                          "       tcrlin04, tcrlsu04, tcrtol04, tcrimp04, tcridv04, tcrimb04, " & _
                          "       tcrlin05, tcrlsu05, tcrtol05, tcrimp05, tcridv05, tcrimb05, " & _
                          "       tcrlin06, tcrlsu06, tcrtol06, tcrimp06, tcridv06, tcrimb06, " & _
                          "       tcrlin07, tcrlsu07, tcrtol07, tcrimp07, tcridv07, tcrimb07, " & _
                          "       tcrlin08, tcrlsu08, tcrtol08, tcrimp08, tcridv08, tcrimb08, " & _
                          "       tcrlin09, tcrlsu09, tcrtol09, tcrimp09, tcridv09, tcrimb09, " & _
                          "       tcrlin10, tcrlsu10, tcrtol10, tcrimp10, tcridv10, tcrimb10, " & _
                          "       tcrlin11, tcrlsu11, tcrtol11, tcrimp11, tcridv11, tcrimb11, " & _
                          "       tcrlin12, tcrlsu12, tcrtol12, tcrimp12, tcridv12, tcrimb12, " & _
                          "       tcrlin13, tcrlsu13, tcrtol13, tcrimp13, tcridv13, tcrimb13, " & _
                          "       tcrlin14, tcrlsu14, tcrtol14, tcrimp14, tcridv14, tcrimb14, " & _
                          "       tcrlin15, tcrlsu15, tcrtol15, tcrimp15, tcridv15, tcrimb15, " & _
                          "       tcrtipop FROM taritcr0 WHERE tcridacl = ? AND tcridacn = ? AND tcridadd = ? AND tcridtfc = ? AND tcrannul = 0"
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Cliente", "taritcr0.tcridacl", 0, True, 10, "0"
        With .Item("0")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "acldescr"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Contratto", "taritcr0.tcridacn", 0, True, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ACN"
            .ForeignTable = "anagacn0"
            .ForeignTableFld = "acnidacn"
            .ForeignTableDescField = "acndescr"
            .ForeignTableDeleteFld = "acnannul"
        End With
        .Add "Addebito", "taritcr0.tcridadd", 0, True, 3, "98"
        With .Item("98")
            .ForeignTableEntity = "ADD"
            .ForeignTable = "anagadd0"
            .ForeignTableFld = "addidadd"
            .ForeignTableDescField = "adddescr"
            .ForeignTableDeleteFld = "addannul"
        End With
        .Add "Codice", "taritcr0.tcridtfc", 0, True, 3, "2"
        .Add "Progressivo", "taritcr0.tcrprogd", 1, True, 3, "3"
        .Add "Limite inferiore Km", "taritcr0.tcrinfkm", 1, True, 7, "4"
        .Add "Limite superiore Km", "taritcr0.tcrsupkm", 1, True, 7, "5"
        For i = 1 To 15  '12
            strCounter = Format$(i, "00")
            .Add "Limite lat. inf. quantità " & strCounter, "taritcr0.tcrlin" & strCounter, 1, , 7, CStr((6 * i))
            .Add "Limite lat. sup. quantità " & strCounter, "taritcr0.tcrlsu" & strCounter, 1, , 7, CStr((6 * i + 1))
            .Add "Limite tolleranza quantità " & strCounter, "taritcr0.tcrtol" & strCounter, 1, , 7, CStr((6 * i + 2))
            .Add "Importo tariffa quantità " & strCounter, "taritcr0.tcrimp" & strCounter, 1, , 13, CStr((6 * i + 3))
            .Add "Valuta tariffa quantità " & strCounter, "taritcr0.tcridv" & strCounter, 0, , 3, CStr((6 * i + 4))
            With .Item(CStr((6 * i + 4)) - 1)
                .ForeignTableEntity = "AVA"
                .ForeignTable = "anagava0"
                .ForeignTableFld = "avaidava"
                .ForeignTableDescField = "avadescr"
                .ForeignTableDeleteFld = "avaannul"
            End With
            .Add "Importo tariffa distanza " & strCounter, "taritcr0.tcrimb" & strCounter, 1, , 13, CStr((6 * i + 5))
        Next i
        .Add "Tipo percorso", "taritcr0.tcrtipop", 0, True, 1, "96" '"65"
        .Add "Annullato", "taritcr0.tcrannul", 3, , , "97" '"66"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "98", "3"
        .Add "2", "4"
        .Add "3", "5"
    End With
    
    DS.CancelField = "taritcr0.tcrannul"
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


