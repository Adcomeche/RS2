VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStipendioTestata"
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

Private SQL_DELETE As String
Private SQL_DELETE_DETAIL As String
Private SQL_GET_SELECTION As String

Private SQL_GET_SELECTION_VIAGGI As String
Private SQL_GET_SELECTION_ORARI As String
Private SQL_GET_SELECTION_AUTOMEZZI As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

Public Function Delete(varIDSDT As Variant) As Boolean
On Error GoTo ErrHandler
    
    With DS.DA
        .BeginTransaction
        
        If Not .ExecuteCommand(SQL_DELETE_DETAIL, varIDSDT) Then
            GoTo ErrHandler
        End If
        If Not .ExecuteCommand(SQL_DELETE, varIDSDT) Then
            GoTo ErrHandler
        End If
        
        .EndTrans etCommit
    End With
    
    Delete = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    Delete = False
End Function

Public Function CreateStipendio(IDAutista As Variant, _
                                DataEmissione As Variant, _
                                DallaData As Variant, _
                                AllaData As Variant, _
                                Optional ErrorMessage As String) As Boolean
On Error GoTo ErrHandler

    Dim AAU As CAutista
    
    Dim lngVolumeTrasportato As Long
    Dim lngNumConsegne280 As Long
    Dim lngNumConsegne285 As Long
    Dim lngMaxKmViaggio As Long
    Dim lngCounter As Long
    Dim lngHowMany As Long
    Dim lngHowManyAutomezzi As Long
    Dim NDestinos As Integer
    NDestinos = 0
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim strError As String
    Dim strSQL As String
    Dim strProgressivo As String
    Dim strIDVettore As String
    Dim strIDBase As String
    Dim strIDViaggio As String
    Dim strIDViaggioPrec As String
    Dim strIDDestinazione As String
    Dim strIDDestinazionePrec As String
    Dim strIDAddebito As String
    Dim strIDUdM As String
    Dim strIDTipoMezzo As String
    Dim strRagDest5 As String
    Dim strFlagContatore As String
    Dim strIDAccreditoManuale As String
    Dim ModeloFact As String
    
    Dim sngImportoUnitario As Single
    Dim sngQuantita As Single
    Dim sngMinuti As Single
    
    Dim varBufferViaggi As Variant
    Dim varBufferOrari As Variant
    Dim varBufferAccrediti As Variant
    Dim varDataViaggio As Variant
    Dim varKmContratto As Variant
    Dim varIDStipendio As Variant
    Dim varProgressivoRiga As Variant
    Dim varDataInizioLavoro As Variant
    Dim varOraInizioLavoro As Variant
    Dim varDataFineLavoro As Variant
    Dim varOraFineLavoro As Variant
    Dim varKmGiorno As Variant
    Dim varDataRiferimento As Variant
    Dim varDataRiferimentoAnt As Variant
    Dim varGiornoSettimana As Variant
    Dim varDataOraInizio As Variant
    Dim varDataOraFine As Variant
    Dim varOreLavorate As Variant
    Dim varListAutomezzi As Variant
    
    Dim blnFineViaggio As Boolean
    Dim blnFineGiorno As Boolean
    Dim blnFlagPranzo As Boolean
    Dim blnFlag100 As Boolean
    Dim blnFlag270 As Boolean
    Dim blnFlag260 As Boolean
    Dim blnFlag235 As Boolean
    Dim blnFlag520 As Boolean
    Dim blnFlag580 As Boolean
    Dim blnFlagFuel As Boolean
    Dim blnFestivo As Boolean
    
    Dim OLA As CStipendioDettaglioAddebito
    Set OLA = New CStipendioDettaglioAddebito
    With OLA.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    If Not IsDate(DataEmissione) Then
        strError = "Data emissione non valida"
        GoTo ErrHandler
    End If
    If Not IsDate(DallaData) Then
        strError = "Data inizio periodo non valida"
        GoTo ErrHandler
    End If
    If Not IsDate(AllaData) Then
        strError = "Data fine periodo non valida"
        GoTo ErrHandler
    End If
    
    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    strSQL = SQL_GET_SELECTION_VIAGGI
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(strSQL, varBufferViaggi, lngHowMany, _
                              DallaData, AllaData, IDAutista) Then
        strError = "Errore nella selezione dei trasporti effettuati"
        GoTo ErrHandler
    End If
    
    strSQL = SQL_GET_SELECTION_ORARI
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(strSQL, varBufferOrari, lngHowMany, _
                              IDAutista, DallaData, AllaData) Then
        strError = "Errore nella selezione degli orari di lavoro"
        GoTo ErrHandler
    End If
    
    If (Not IsEmpty(varBufferViaggi)) Or (Not IsEmpty(varBufferOrari)) Then
        DS.DA.BeginTransaction
        
        ' Inserisce la testata
        DS.Field("Codice Autista").NewValue = IDAutista
        DS.Field("Codice Autista").Updated = True
        DS.Field("Data inizio periodo").NewValue = DallaData '"01/03/2007"
        DS.Field("Data inizio periodo").Updated = True
        DS.Field("Data fine periodo").NewValue = AllaData '"31/03/2007"
        DS.Field("Data fine periodo").Updated = True
        DS.Field("Data Emissione").NewValue = DataEmissione
        DS.Field("Data Emissione").Updated = True
        If Not Insert() Then
            strError = "Errore durante l'inserimento della testata"
            GoTo ErrHandler
        End If
        varIDStipendio = DS.Field("Numero Stipendio").NewValue
        varProgressivoRiga = 1
        
'        If Not InsertRiga("630", 1, "NR", 156.3, _
'                          varIDStipendio, varProgressivoRiga, DallaData, strError) Then
'            GoTo ErrHandler
'        End If
    End If
    
    If Not IsEmpty(varBufferOrari) Then
        lngHowMany = UBound(varBufferOrari, 2)
        
        strIDVettore = UCase(varBufferOrari(8, 0))
        strIDBase = UCase(varBufferOrari(7, 0))
        
        ' Modelos de facturacion diferentes
        

        
        If strIDVettore = "A01" Then
            ModeloFact = "TPH"
        End If
        
        If strIDVettore = "A02" Then
            ModeloFact = "EUT"
        End If
        
        If strIDVettore = "A02" And (strIDBase = "002" Or strIDBase = "004" Or strIDBase = "010") Then
            ModeloFact = "TPH"
        End If
        
        For i = 0 To lngHowMany
            strIDAddebito = vbNullString
            
            varDataRiferimento = varBufferOrari(0, i)
            varDataInizioLavoro = varBufferOrari(1, i)
            varOraInizioLavoro = varBufferOrari(2, i)
            varDataFineLavoro = varBufferOrari(3, i)
            varOraFineLavoro = varBufferOrari(4, i)
            If IsNull(varBufferOrari(5, i)) Then
                varKmGiorno = 0
            Else
                varKmGiorno = varBufferOrari(5, i)
            End If
            
            varBufferAccrediti = Empty
            blnFlagPranzo = False
            blnFlag100 = False
            OLA.DS.Fields.ResetFields
            OLA.DS.Field("Codice Autista").FindValue = IDAutista
            OLA.DS.Field("Data").FindValue = varDataRiferimento
            If Not OLA.GetList(True) Then
                GoTo ErrHandler
            End If
            If OLA.DS.HowMany > 0 Then
                varBufferAccrediti = OLA.DS.List
            End If
            If Not IsEmpty(varBufferAccrediti) Then
                For k = 0 To UBound(varBufferAccrediti, 2)
                    strIDAccreditoManuale = varBufferAccrediti(2, k)
                    If strIDAccreditoManuale = "001" Then
                        blnFlagPranzo = True
                    End If
                    strIDAddebito = strIDAccreditoManuale
                    sngQuantita = varBufferAccrediti(3, k)
                    If IsNull(varBufferAccrediti(4, k)) Or (Len(varBufferAccrediti(4, k)) = 0) Then
                        strIDUdM = "NR"
                    Else
                        strIDUdM = varBufferAccrediti(4, k)
                    End If
                    Select Case ModeloFact
                        Case "TPH"  'TPH
                            Select Case strIDAddebito
                                Case "001"
                                    sngQuantita = 0
                                    sngImportoUnitario = 0
                                Case "003"
                                    sngImportoUnitario = 36.72
                                Case "004"
                                    sngImportoUnitario = 1.65
                                Case "010"
                                    sngImportoUnitario = 7.48
                                Case "020"
                                    sngImportoUnitario = 7.4
                                Case "100"
                                    blnFlag100 = True
                                    sngImportoUnitario = 13.81 '13.25 '12.66
                                Case "110"
                                    sngImportoUnitario = 3.97 '3.9 '3.64
                                Case "300"
                                    sngImportoUnitario = 0.1
                                Case "320"
                                    sngImportoUnitario = 6.25 '6.14 '22.72
                                Case "600"
                                    sngImportoUnitario = 59.85 '56.32 '55 ' 51.85
                                Case "610"
                                    sngImportoUnitario = 416.97 '307.2 '300
                                Case "620"
                                    sngImportoUnitario = 312.7 '409.6 '400
                                Case "640"
                                    sngImportoUnitario = 52.12 '51.2 '50
                                Case "650"
                                    sngImportoUnitario = 31.32 '30.77 '30.05
                                Case "900"
                                    If IsNumeric(varBufferAccrediti(6, k)) Then
                                        sngQuantita = 1
                                        sngImportoUnitario = varBufferAccrediti(6, k)
                                    Else
                                        sngImportoUnitario = 0
                                    End If
                                Case Else
                                    GoTo ErrHandler
                            End Select
                        Case "EUT"  'Eutrasur
                            Select Case strIDAddebito
                                Case "001"
                                    sngImportoUnitario = 11.15 '11 '10.56
                                Case "003"
                                    sngImportoUnitario = 49 '38.8 '38.26
                                Case "004"
                                    sngImportoUnitario = 1.65
                                Case "010"
                                    sngImportoUnitario = 9.5 '6.75 '7.03 '7.9  '7.79 '7.48
                                Case "020"
                                    sngImportoUnitario = 9.5 '6.75 '7.03 '7.9 '7.83 '7.72 '7.4
                                Case "100"
                                    blnFlag100 = True
                                    sngImportoUnitario = 13.25 '12.66
                                Case "110"
                                    sngImportoUnitario = 3.64
                                Case "300"
                                    sngImportoUnitario = 0.07 '0.08 '0.09 '0.104 '0.1
                                Case "320"
                                    sngImportoUnitario = 0 '17.28 '21.87 '23.64 '22.72
                                Case "600"
                                    sngImportoUnitario = 55 '51.85
                                Case "610"
                                    sngImportoUnitario = 300
                                Case "620"
                                    sngImportoUnitario = 400
                                Case "640"
                                    sngImportoUnitario = 44.5 '50
                                Case "650"
                                    sngImportoUnitario = 30.47 '30.05
                                Case "900"
                                    If IsNumeric(varBufferAccrediti(6, k)) Then
                                        sngQuantita = 1
                                        sngImportoUnitario = varBufferAccrediti(6, k)
                                    Else
                                        sngImportoUnitario = 0
                                    End If
                                Case Else
                                    GoTo ErrHandler
                            End Select
                    End Select
                    If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                      varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                        GoTo ErrHandler
                    End If
                Next k
            End If
            
            If (varDataInizioLavoro <> "01/01/1900") And (varOraInizioLavoro <> "01/01/1900") Then
                
            
                Select Case ModeloFact
                    Case "TPH"
                        ' *** 520 - Fine lavoro oltre le 15.00 e pausa pranzo
                        blnFlag520 = False
                        If (DateDiff("n", Format(varOraFineLavoro, "hh:nn"), "15:00") < 0) And (blnFlagPranzo) Then
                            strIDAddebito = "520"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            sngImportoUnitario = 14.13 '13.53 '13.3  '12.99 '12.4
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                            blnFlag520 = True
                        End If
                        ' ***
                        
                        ' *** 500 - Presenza oltre le 8 ore
                        varDataOraInizio = varDataInizioLavoro & " " & Format(varOraInizioLavoro, "hh:nn")
                        varDataOraFine = varDataFineLavoro & " " & Format(varOraFineLavoro, "hh:nn")
                        If DateDiff("n", varDataOraInizio, varDataOraFine) < 0 Then
                            strError = "Orari di lavoro incongruenti, data " & varDataInizioLavoro
                            GoTo ErrHandler
                        End If
'                        varOreLavorate = DateDiff("h", varDataOraInizio, varDataOraFine)
                        varOreLavorate = Format(DateDiff("n", varDataOraInizio, varDataOraFine) / 60, "0.00")
                        If varOreLavorate > 8 Then
                            strIDAddebito = "500"
                            sngQuantita = varOreLavorate - 8
                            
                            If (sngQuantita >= 1) And blnFlag520 Then
                                sngQuantita = sngQuantita - 1
                            End If
                            
                            strIDUdM = "NR"
                            sngImportoUnitario = 11.56 '13.56 '13.33 '13.02 '12.44
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                        
                        ' *** 510 - Lavoro notturno (22:00 - 06:00)
                        sngQuantita = 0
                        varDataOraInizio = varDataInizioLavoro & " " & Format(varOraInizioLavoro, "hh:nn")
                        varDataOraFine = varDataFineLavoro & " " & Format(varOraFineLavoro, "hh:nn")
                        sngMinuti = DateDiff("n", varDataOraInizio, varDataOraFine)
                        If sngMinuti < 0 Then
                            strError = "Orari di lavoro incongruenti, data " & varDataInizioLavoro
                            GoTo ErrHandler
                        End If
                        For j = 1 To sngMinuti
                            varDataOraInizio = DateAdd("n", 1, varDataOraInizio)
                            If (DateDiff("n", Format(varDataOraInizio, "hh:nn"), "22:00") < 0) Or (DateDiff("n", Format(varDataOraInizio, "hh:nn"), "06:00") >= 0) Then
                                sngQuantita = sngQuantita + 1
                            End If
                        Next j
                        If sngQuantita > 0 Then
                            
                            strIDAddebito = "510"
                            sngQuantita = Format(sngQuantita / 60, "0.00")
                            strIDUdM = "HR"
                            sngImportoUnitario = 2.43 '2.32 '2.28 '2.23 '2.13
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                        
                        ' *** 530 - Inizio lavoro prima delle 20:00 e fine lavoro oltre le 22.00
                        varDataOraFine = varDataFineLavoro & " " & Format(varOraFineLavoro, "hh:nn")
                        If (DateDiff("n", varOraInizioLavoro, "20:00") > 0) And (DateDiff("n", varDataOraFine, varDataInizioLavoro & " 22:00") < 0) Then
                            strIDAddebito = "530"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            'Dieta cena
                            sngImportoUnitario = 14.13 '13.53 '13.3 '12.99 '12.4
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                        
                        ' *** 540 - Fine lavoro oltre le 15.00 e no pausa pranzo
                        varDataOraFine = varDataFineLavoro & " " & Format(varOraFineLavoro, "hh:nn")
'                        If (DateDiff("n", Format(varOraFineLavoro, "hh:nn"), "15:00") < 0) And (Not blnFlagPranzo) And (Not blnFlag100) Then
                        If (DateDiff("n", varDataOraFine, varDataInizioLavoro & " 15:00") < 0) And (Not blnFlagPranzo) And (Not blnFlag100) Then
                            strIDAddebito = "540"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            sngImportoUnitario = 10.96 '10.5 '10.32 '10.08 '9.63
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                        
                        ' *** 550 - Inizio lavoro tra le 04:00 e le 06:00 o tra le 22:00 e le 00:00
                        If ((DateDiff("n", Format(varOraInizioLavoro, "hh:nn"), "04:00") < 0) And (DateDiff("n", Format(varOraInizioLavoro, "hh:nn"), "06:00") > 0)) Or _
                            ((DateDiff("n", Format(varOraFineLavoro, "hh:nn"), "22:00") < 0) And (DateDiff("n", Format(varOraFineLavoro, "hh:nn"), "23:59") >= 0)) Then
                            strIDAddebito = "550"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            sngImportoUnitario = 2.09 '2 '1.97 '1.92 '1.83
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                    
                        ' *** 590 - Lavoro notturno (24:00 - 04:00)
                        varDataOraInizio = varDataInizioLavoro & " " & Format(varOraInizioLavoro, "hh:nn")
                        varDataOraFine = varDataFineLavoro & " " & Format(varOraFineLavoro, "hh:nn")
                        If DateDiff("n", varDataOraInizio, varDataOraFine) < 0 Then
                            strError = "Orari di lavoro incongruenti, data " & varDataInizioLavoro
                            GoTo ErrHandler
                        End If
                        If (DateDiff("n", varDataOraInizio, varDataFineLavoro & " 04:00") >= 0) And (DateDiff("n", varDataOraFine, varDataFineLavoro - 1 & " 00:00") <= 0) Then
                            If (DateDiff("n", varDataOraInizio, varDataFineLavoro - 1 & " 00:00") >= 0) Then
                                varDataOraInizio = varDataFineLavoro - 1 & " 00:00"
                            End If
                            If (DateDiff("n", varDataOraFine, varDataFineLavoro & " 04:00") <= 0) Then
                                varDataOraFine = varDataFineLavoro & " 04:00"
                            End If
                            
                            strIDAddebito = "590"
                            sngQuantita = DateDiff("n", varDataOraInizio, varDataOraFine)
'                            If sngQuantita >= 15 Then
                            If sngQuantita >= 0 Then
                                sngQuantita = 1
                                strIDUdM = "NR"
                                ' Compensacion nocturna
                                sngImportoUnitario = 10.96 '10.5 '10.32 '9.63
                                If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                                  varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                    GoTo ErrHandler
                                End If
                            End If
                        End If
                        ' ***
                    
                    Case "EUT"
                    
                        ' *** 030 - Giorno festivo (sabato, domenica)
                        varGiornoSettimana = Weekday(varDataRiferimento)
                        Select Case varGiornoSettimana
                            Case 1, 7
                                strIDAddebito = "030"
                                sngQuantita = 1
                                strIDUdM = "NR"
                                sngImportoUnitario = 42.72 '44.5 '50 '58.76 '57.94 '55.612
                                If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                                  varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                    GoTo ErrHandler
                                End If
                        End Select
                        ' ***
                        
                        ' *** 310 - Lavoro notturno (22:00 - 06:00)
                        sngQuantita = 0
                        varDataOraInizio = varDataInizioLavoro & " " & Format(varOraInizioLavoro, "hh:nn")
                        varDataOraFine = varDataFineLavoro & " " & Format(varOraFineLavoro, "hh:nn")
                        sngMinuti = DateDiff("n", varDataOraInizio, varDataOraFine)
                        If sngMinuti < 0 Then
                            strError = "Orari di lavoro incongruenti, data " & varDataInizioLavoro
                            GoTo ErrHandler
                        End If
                        For j = 1 To sngMinuti
                            varDataOraInizio = DateAdd("n", 1, varDataOraInizio)
                            If (DateDiff("n", Format(varDataOraInizio, "hh:nn"), "22:00") < 0) Or (DateDiff("n", Format(varDataOraInizio, "hh:nn"), "06:00") >= 0) Then
                                sngQuantita = sngQuantita + 1
                            End If
                        Next j
                        If sngQuantita > 0 Then
                            
                            strIDAddebito = "310"
                            ' *** 28/08/2007 - Il lavoro notturno si paga per parte intera
'                            sngQuantita = Format(sngQuantita / 60, "0.00")
                            sngQuantita = (sngQuantita \ 60) + 1
                            ' ***
                            ' Nocturnidad
                            strIDUdM = "HR"
                            sngImportoUnitario = 1.008 '1.05 '1.19 '1.29  '1.28 ' 1.23
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                        
                        ' *** 225 Km reale trailer -- 235 Km reale rigido
                        strSQL = SQL_GET_SELECTION_AUTOMEZZI
                        lngHowManyAutomezzi = 0
                        If Not DS.DA.QueryCommand(strSQL, varListAutomezzi, lngHowManyAutomezzi, IDAutista, varDataRiferimento) Then
                            strError = "Errore nella selezione degli automezzi"
                            GoTo ErrHandler
                        End If
                        
                        If Not IsEmpty(varListAutomezzi) Then
                            blnFlag235 = False
                            For j = 0 To UBound(varListAutomezzi, 2)
                                Select Case varListAutomezzi(0, j)
                                    Case "1", "5", "6", "7"
                                        blnFlag235 = True
                                End Select
                                Select Case varListAutomezzi(1, j)
                                    Case "1", "5", "6", "7"
                                        blnFlag235 = True
                                End Select
                            Next j
                            
                            If blnFlag235 Then
                                strIDAddebito = "235"
                                Select Case strIDBase
'                                    Case "008"  ' Granada
'                                        sngImportoUnitario = 0.042 '0.048 '0.062 '0.06
                                    Case Else
                                        sngImportoUnitario = 0
                                End Select
                            Else
                                strIDAddebito = "225"
                                Select Case strIDBase
'                                    Case "008"  ' Granada
'                                        sngImportoUnitario = 0.042 '0.048 '0.052 ' 0.05
                                    Case Else
                                        sngImportoUnitario = 0
                                End Select
                            End If
                            
                            strIDUdM = "KM"
                            sngQuantita = varKmGiorno
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataRiferimento, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        ' ***
                End Select
            End If
        Next i
    End If
            
    blnFineGiorno = False
    If Not IsEmpty(varBufferViaggi) Then
        
        lngHowMany = UBound(varBufferViaggi, 2)
        
        strIDVettore = UCase(varBufferViaggi(15, 0))
        strIDBase = UCase(varBufferViaggi(14, 0))
        
        For i = 0 To lngHowMany
            strIDAddebito = vbNullString
            strIDViaggio = varBufferViaggi(0, i)
            varDataViaggio = varBufferViaggi(1, i)
            strIDDestinazione = varBufferViaggi(5, i)
            lngVolumeTrasportato = lngVolumeTrasportato + varBufferViaggi(9, i)
            If IsNumeric(varBufferViaggi(10, i)) Then
                varKmContratto = varBufferViaggi(10, i)
            Else
                varKmContratto = 0
            End If
            If varKmContratto > lngMaxKmViaggio Then
                lngMaxKmViaggio = varKmContratto
            End If
            If IsNull(varBufferViaggi(11, i)) Then
                strRagDest5 = vbNullString
            Else
                strRagDest5 = UCase(varBufferViaggi(11, i))
            End If
            If IsNull(varBufferViaggi(17, i)) Then
                strIDTipoMezzo = varBufferViaggi(16, i)
            Else
                strIDTipoMezzo = varBufferViaggi(17, i)
            End If
            If IsNull(varBufferViaggi(13, i)) Or (Len(varBufferViaggi(13, i)) = 0) Then
                strFlagContatore = IIf(IsNull(varBufferViaggi(18, i)), vbNullString, varBufferViaggi(18, i))
            Else
                strFlagContatore = IIf(IsNull(varBufferViaggi(19, i)), vbNullString, varBufferViaggi(19, i))
            End If
            
            If strFlagContatore = "S" Then
                blnFlag270 = True
                blnFlag580 = True
            End If
            Select Case strRagDest5
                Case "SN"
                    blnFlag260 = True
            End Select
            If varBufferViaggi(20, i) = "020" Then
                blnFlagFuel = True
            End If
            
            If UCase(varBufferViaggi(21, i)) = "F" Then
                If Weekday(varDataViaggio) > 1 And Weekday(varDataViaggio) < 7 Then
                    blnFestivo = True
                End If
            End If
            
            ' *** Check fine viaggio
            If (i = lngHowMany) Then
                blnFineViaggio = True
                blnFineGiorno = True
            ElseIf varBufferViaggi(0, i + 1) <> strIDViaggio Then
                blnFineViaggio = True
            Else
                blnFineViaggio = False
            End If
            ' ***
                                        
            ' *** Check fine giornata
            If i < lngHowMany Then
                If varBufferViaggi(1, i + 1) <> varDataViaggio Then
                    blnFineGiorno = True
                Else
                    blnFineGiorno = False
                End If
            End If
            ' ***
            
            If (strIDViaggio <> strIDViaggioPrec) Or (strIDDestinazione <> strIDDestinazionePrec) Then
                Select Case strRagDest5
                    'Case "BG"
                    Case "B"
                        lngNumConsegne280 = lngNumConsegne280 + 1
                    Case "G"
                        lngNumConsegne285 = lngNumConsegne285 + 1
                End Select
            End If
            
            Select Case ModeloFact
                Case "TPH"
                    If blnFineGiorno Then
                        ' *** 580 - Veicolo con contatore
                        If blnFlag580 Then
                            strIDAddebito = "580"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            sngImportoUnitario = 5 '6.41 '6.14 '6 '5.19
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                            ' *** Se riconosciamo il viaggio con contatore escludiamo eventuali addebiti per
                            ' *** viaggi di fuel
                            blnFlagFuel = False
                            ' ***
                        End If
                        blnFlag580 = False
                        ' ***
                        ' *** 580 - Consegne Fuel
                        If blnFlagFuel Then
                            strIDAddebito = "580"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            sngImportoUnitario = 6.14 '6 '5.19
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        blnFlagFuel = False
                        ' ***
                    End If
                Case "EUT"
                    If blnFineViaggio Then
                        ' *** 220 - Addebito per km percorsi (trailer) -- 230 - Addebito per km percorsi (rigido)
                        Select Case strIDTipoMezzo
                            Case "1", "5", "6", "7"
                                strIDAddebito = "230"
                                Select Case strIDBase
'                                    Case "008"  ' Granada
'                                        sngImportoUnitario = 0
                                    Case Else
                                        sngImportoUnitario = 0.108  '0.1075 '0.112 '0.126 '0.136 '0.135 '0.129
                                End Select
                            Case "2", "3", "4", "8"
                                strIDAddebito = "220"
                                Select Case strIDBase
'                                    Case "008"  ' Granada
'                                        sngImportoUnitario = 0
                                    Case Else
                                        sngImportoUnitario = 0.108  '0.1075 '0.112 '0.126 '0.136 '0.135 '0.129
                                End Select
                            Case Else
                                strIDAddebito = vbNullString
                        End Select
                        
                        If lngMaxKmViaggio < 100 Then
                            sngQuantita = 100
                        Else
                            sngQuantita = lngMaxKmViaggio
                        End If
                        strIDUdM = "KM"
                        
                        If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                          varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                            GoTo ErrHandler
                        End If
                        
                        lngMaxKmViaggio = 0
                        ' ***
                        
                        ' *** 240 - Un addebito per ogni viaggio effettuato
                        strIDAddebito = "240"
                        sngQuantita = 1
                        strIDUdM = "NR"
                        Select Case strIDBase
'                            Case "008"  ' Granada
'                                sngImportoUnitario = 3.44 '3.87 '4.18 '4.13 '3.96
                            Case Else
                                sngImportoUnitario = 0
                        End Select
                        If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                          varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                            GoTo ErrHandler
                        End If
                        ' ***
                        
                        ' *** 250 - Volume trasportato
                        strIDAddebito = "250"
                        sngQuantita = Format(lngVolumeTrasportato / 1000, "0.000")
                        strIDUdM = "MC"
                        Select Case strIDBase
'                            Case "008"  ' Granada
'                                sngImportoUnitario = 0.333 '0.375 '0.405 '0.4  '0.38
                            Case Else
                                sngImportoUnitario = 0
                        End Select
                        If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                          varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                            GoTo ErrHandler
                        End If
                        lngVolumeTrasportato = 0
                        ' ***
                        
                        ' *** 260 Sierra Nevada
                        If blnFlag260 Then
                            strIDAddebito = "260"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            Select Case strIDBase
'                                Case "008"  ' Granada
'                                    sngImportoUnitario = 6.52 '6.62 '7.16 '7.07 '6.788
                                Case Else
                                    sngImportoUnitario = 0
                            End Select
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        blnFlag260 = False
                        ' ***
                    End If
                    
                    If blnFineGiorno Then
                        ' *** 030 - Addebito per giorno festivo (diverso da sabato o domenica)
                        If blnFestivo Then
                            strIDAddebito = "030"
                            sngQuantita = 1
                            strIDUdM = "NR"
                            sngImportoUnitario = 58.76 '57.94 '55.61
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        blnFestivo = False
                        ' ***
                        
                        ' *** 280 Buques -- 290 Buques 2+
                        For j = 1 To lngNumConsegne280
                            If j = 1 Then
                                strIDAddebito = "280"
                                sngImportoUnitario = 0 '17.28 '19.42 '20.99 '20.7 '19.862
                            Else
                                strIDAddebito = "290"
                                sngImportoUnitario = 0 '8.64 '9.71 '10.49 '10.35 '9.931
                            End If
                            sngQuantita = 1
                            strIDUdM = "NR"
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                        Next j
                        lngNumConsegne280 = 0
                        ' ***
                        
                        ' *** 285 Gibraltar -- 295 Gibraltar 2+
                        For j = 1 To lngNumConsegne285
                            If j = 1 Then
                                strIDAddebito = "285"
                                sngImportoUnitario = 16.59 '17.28 '19.42 '20.99 '20.7 '19.862
                            Else
                                strIDAddebito = "295"
                                sngImportoUnitario = 8.29 '8.64 '9.71 '10.49 '10.35 '9.931
                            End If
                            sngQuantita = 1
                            strIDUdM = "NR"
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                        Next j
                        lngNumConsegne285 = 0
                        ' ***
                        
                        ' *** 270 - Veicolo con contatore
                        ' JG 25/1/12 - No se paga a los conductores los servicios con MOTOBOMBA
'                        If blnFlag270 Then
'                            strIDAddebito = "270"
'                            sngQuantita = 1
'                            strIDUdM = "NR"
'                            sngImportoUnitario = 7.33 '8.24 '8.91 '8.79 '8.56
'                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
'                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
'                                GoTo ErrHandler
'                            End If
'                        End If
'                        blnFlag270 = False
                        ' ***
                    End If
            End Select
            
            If (strIDViaggio <> strIDViaggioPrec) Or (strIDDestinazione <> strIDDestinazionePrec) Then
                ' Inizio elaborazione di una consegna
'                If strIDViaggio <> strIDViaggioPrec Then
'                    lngNumConsegne280 = 0
'                End If
                
                
                Select Case ModeloFact
                    Case "TPH"
                    Case "EUT"
                        ' 210 - Un addebito per ogni destinazione del viaggio
                        strIDAddebito = "210"
                        sngQuantita = 1
                        strIDUdM = "NR"
                        Select Case strIDBase
'                            Case "008"  ' Granada
'                                sngImportoUnitario = 1.03 '1.16 '1.26 '1.24 '1.191
                            Case Else
                                sngImportoUnitario = 1.91 '1.99 '2.24 '2.42 '2.39 '2.29
                        End Select
                        If NDestinos >= 0 And NDestinos <> 1 Then
                            If Not InsertRiga(strIDAddebito, sngQuantita, strIDUdM, sngImportoUnitario, _
                                              varIDStipendio, varProgressivoRiga, varDataViaggio, strError) Then
                                GoTo ErrHandler
                            End If
                        End If
                        NDestinos = NDestinos + 1
                    Case Else
                        strError = "Vettore sconosciuto"
                        GoTo ErrHandler
                End Select
                
                strIDDestinazionePrec = strIDDestinazione
            End If
            
            ' destino por Dia
'            If varDataRiferimento <> varDataRiferimentoAnt Then
'                If varDataRiferimentoAnt <> "" Then NDestinos = 0
'                varDataRiferimentoAnt = varDataRiferimento
'            End If
            
            ' Destino por viaje
            
            If strIDViaggio <> strIDViaggioPrec Then
                If strIDViaggioPrec <> "" Then NDestinos = 0
                strIDViaggioPrec = strIDViaggio
            End If
            
            If i Mod 10 = 0 Then
                DoEvents
            End If
            
        Next i
    End If
    
    DS.DA.EndTrans etCommit
    CreateStipendio = True
    Exit Function

ErrHandler:

    DS.DA.EndTrans etRollback
    If Not IsMissing(ErrorMessage) Then
        If Len(strError) > 0 Then
            ErrorMessage = strError '& ", Autista " & IDAutista
        End If
    End If
    CreateStipendio = False
End Function


Private Function InsertRiga(IDAddebito As Variant, _
                            Quantita As Single, _
                            IDUnitaDiMisura As String, _
                            ImportoUnitario As Single, _
                            IDStipendio As Variant, _
                            ByRef ProgressivoRiga As Variant, _
                            DataRiferimento As Variant, _
                            Optional ErrorMessage As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim SDR As CStipendioRiga
    Dim sngImporto As Single
    Dim strError As String
    
    Set SDR = New CStipendioRiga
    With SDR.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    If Len(IDAddebito) > 0 Then
        sngImporto = Format(Quantita * ImportoUnitario, "0.000")
        If sngImporto <> 0 Then
            With SDR.DS
                .Fields.ResetFields
                .Field("Numero Stipendio").NewValue = IDStipendio
                .Field("Numero Stipendio").Updated = True
                .Field("Progressivo Riga").NewValue = Format(ProgressivoRiga, "0000000")
                .Field("Progressivo Riga").Updated = True
                .Field("Codice Addebito").NewValue = IDAddebito
                .Field("Codice Addebito").Updated = True
                .Field("Data Riferimento").NewValue = DataRiferimento
                .Field("Data Riferimento").Updated = True
                .Field("Quantità").NewValue = Quantita
                .Field("Quantità").Updated = True
                .Field("Unità di Misura").NewValue = IDUnitaDiMisura
                .Field("Unità di Misura").Updated = True
                .Field("Importo unitario").NewValue = ImportoUnitario
                .Field("Importo unitario").Updated = True
                .Field("Importo").NewValue = sngImporto
                .Field("Importo").Updated = True
            End With
            If Not SDR.Insert Then
                strError = "Errore durante l'inserimento di una riga di dettaglio"
                GoTo ErrHandler
            End If
            ProgressivoRiga = ProgressivoRiga + 1
        End If
    End If
    
    InsertRiga = True
    
    Exit Function

ErrHandler:

    DS.DA.EndTrans etRollback
    If Not IsMissing(ErrorMessage) Then
        If Len(strError) > 0 Then
            ErrorMessage = strError
        End If
    End If
    InsertRiga = False

End Function

Public Function Find() As Boolean
    Dim blnResult As Boolean
    
    blnResult = DS.Find(Me.SQL_FIND)
    Find = blnResult
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    DS.Field("Stato").NewValue = "P"
    DS.Field("Stato").Updated = True
    
    Insert = DS.Insert(Me.SQL_INSERT, "anaganu0", "anuprogr", "anuidanu", "700", 0)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean

    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    BeginSave = True
    ' ---
End Function
Private Sub Class_Initialize()
    SQL_FIND = "SELECT SDTIDSDT, SDTIDAAU, SDTDTINP, SDTDTFIP, " & _
               "       SDTDTEMI, SDTFLSTA, timestmp " & _
               " FROM stipsdt0 WHERE "
    SQL_UPDATE = "UPDATE stipsdt0 SET "
    SQL_INSERT = "INSERT INTO stipsdt0 "
    SQL_GET_TS = "SELECT timestmp FROM stipsdt0 WHERE "
    SQL_DELETE = "DELETE stipsdt0 WHERE SDTIDSDT = ?  "
    SQL_DELETE_DETAIL = "DELETE stipsdr0 WHERE sdridsdt = ? "
    SQL_GET_SELECTION_VIAGGI = " SELECT  VVTIDVVT, VVTDTINV, VVTIDDEF, VVTIDAU1, " & _
                               "         AAUNOMEA, VVCIDADZ, VVCDERAS, VVCDELOC, " & _
                               "         VVPIDAPR, VVPQCOLT, DEDKMATT, ADZRAGU5, " & _
                               "         VVTIDAM1, VVTIDAM2, AAUIDABS, AAUIDAVE, " & _
                               "         ANAGAAM0.AAMIDATM,  ANAGAAM0_1.AAMIDATM, ANAGAAM0.AAMFLSCA, ANAGAAM0_1.AAMFLSCA, " & _
                               "         APRIDTTP, VVTFLST2 " & _
                               "   FROM  VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT INNER JOIN " & _
                               "         VIAGVVP0 ON VVCIDVVT = VVPIDVVT AND VVCIDVVC = VVPIDVVC INNER JOIN ANAGAPR0 ON VVPIDAPR = APRIDAPR INNER JOIN " & _
                               "         ANAGAAU0 ON VVTIDAU1 = AAUIDAAU INNER JOIN ANAGADZ0 ON VVCIDADZ = ADZIDADZ " & _
                               "         LEFT OUTER JOIN ANAGAAM0 ON VVTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
                               "         ANAGAAM0 ANAGAAM0_1 ON VVTIDAM2 = ANAGAAM0_1.AAMIDAAM LEFT OUTER JOIN " & _
                               "         ANAGDED0 ON VVCIDADZ = DEDIDADZ AND VVTIDADE = DEDIDADE " & _
                               "  WHERE  VVTDTINV >= ? AND VVTDTINV <= ? AND (VVTANNUL = 0) AND (VVCANNUL = 0) AND (VVPANNUL = 0) AND " & _
                               "         (VVTIDAU1 = ?) " & _
                               " ORDER BY VVTDTINV, VVTIDVVT, VVCIDADZ, VVPIDAPR "
                               
    SQL_GET_SELECTION_ORARI = " SELECT OLVDTOLV, OLVDTILR, OLVHRILR, OLVDTFIL, " & _
                              "        OLVHRFIL, OLVNRKMS, OLVFLMAS, AAUIDABS, " & _
                              "        AAUIDAVE " & _
                              "   FROM STIPOLV0 INNER JOIN ANAGAAU0 ON OLVIDAAU = AAUIDAAU " & _
                              "  WHERE (OLVIDAAU = ?) AND (OLVDTOLV >= ?) AND (OLVDTOLV <= ?) AND (STIPOLV0.OLVANNUL = 0) " & _
                              "  ORDER BY OLVDTOLV "
    
    SQL_GET_SELECTION_AUTOMEZZI = " SELECT DISTINCT ANAGAAM0.AAMIDATM, ANAGAAM0_1.AAMIDATM " & _
                                  "   FROM VIAGVVT0 LEFT OUTER JOIN ANAGAAM0 ON VVTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
                                  "        ANAGAAM0 ANAGAAM0_1 ON VVTIDAM2 = ANAGAAM0_1.AAMIDAAM " & _
                                  "  WHERE VVTIDAU1 = ? AND VVTDTINV = ? AND VVTANNUL = 0 AND VVTFLSTV <> 'P' "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Stipendio", "stipsdt0.SDTIDSDT", 0, True, 7, "0", , True
        .Add "Codice Autista", "stipsdt0.SDTIDAAU", 0, True, 3, "1", , True
        With .Item("1")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Data Inizio Periodo", "stipsdt0.SDTDTINP", 2, True, 10, "2", , True
        .Add "Data Fine Periodo", "stipsdt0.SDTDTFIP", 2, True, 10, "3", , True
        .Add "Data Emissione", "stipsdt0.SDTDTEMI", 2, True, 10, "4", , True
        .Add "Stato", "stipsdt0.SDTFLSTA", 0, True, 2, "5"
    End With
    With DS.KeyFields
        .Add "0", "1"
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


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, , MaxRecords)
    End If

End Function
