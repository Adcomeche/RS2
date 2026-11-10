Attribute VB_Name = "OIL2GLOBAL"
Option Explicit


'Public DA As CDataAccess
Public DataConnect As CConnect
Public gmblnConnectOk As Boolean
Public gmblnCancelConnect As Boolean
Public gmRegistry As CRegSettings
Public ClientControls As CControls
Public CommonFilesDir As String
Public gmstrCodInstallazione As CodiceInstallazione
Public gmblnDemoMode As Boolean
Public gmblnMultipleWindows As Boolean
Public gmstrPicFileName As String
Public gmstrReportsPath As String
Public gmstrLogFileName As String
Public gmstrServerLogFileName As String
Public ID_PARAM As String

Public Const ETCOMMIT = 0
Public Const ETROLLBACK = 1
    
Public Const DEJET = 1
Public Const DESQLSRV = 0

Public Enum SysColors
    scBlue = &HFF0000
    scRed = vbRed
End Enum

Public Enum FunctionCodes
    fcANAGASC = "211"
    fcANAGAEE = "212"
    fcANAGADE = "213"
    fcANAGACL = "214"
    fcANAGADZ = "215"
    fcORDINSE = "221"
    fcORDSELA = "3221"
    fcORDSELV = "3222"
    fcORDSELM = "3223"
    fcORDSELE = "223"
    fcORDSEVA = "224"
    fcORDFRAZ = "225"
    fcORDRIDU = "226"   '-- DP 19/08/08: Riduzioni
    fcVIAINSE = "231"
    fcVIASELP = "232"
    fcVIASELD = "3131"
    fcVIASELT = "3132"
    fcVIASELC = "3133"
    fcVIACOMP = "234"
    fcVIASELE = "235"
    fcVIASSAT = "236"   ' (PM) Viaggi Sat
    fcCONTVIA = "237"   '-- DP 18/08/08: Controllo Viaggi
    fcGP1SELE = "241"
    
    fcPPTCREA = "2421"   ' Agip/Cepsa
    fcPPTCRE3 = "2422"   ' Altre compagnie
    fcPPTCRE2 = "243"
    fcPPTSELE = "244"
    fcSTIPINS = "247"
    fcSTIPCRE = "248"
    fcSTIPSEL = "249"
    fcPREFPPT = "998000"
    fcPREFPPD = "999888"
    fcPREFGP1 = "999777"
    fcPREFPAD = "998100"
    fcPRNFORM = "999000"
    fcORDIOCT = "999222"
    fcORDIOCR = "999333"
    fcVIAGVVT = "999444"
    fcVIAGVVC = "999555"
    fcVIAGVVP = "999666"
    fcVIAGVPI = "999667"
    fcVIAGCOM = "998111"
    fcPREZAPZ = "998111"
    fcPREZASZ = "998222"
    fcPREZFPD = "998333"
    fcOCTFRAZ = "998444"
    fcSTIPOLA = "998555"
    fcANAGAPR = "3171"
    fcANAGAIM = "3172"
    fcANAGAPD = "3173"
    fcANAGAVE = "3181"
    fcANAGAAU = "3182"
    fcANAGAAM = "3183"
    fcANAGAUT = "3184"
    fcANAGACN = "3185"
    fcANAGACD = "3186"      ' (PM) 06/03/06
    fcTARITFA = "3191"
    fcTARITFB = "3192"
    fcTARITTC = "3193"      ' (PM) 06/03/06
    fcSTATFRM = "251"
    fcAGIPOIM = "261"
    fcMOVEREX = "262"
    fcANAGRIE = "2620"      ' (PM) Interfaccia ricezione/trasmissione anagrafiche
    fcMOVERIM = "263"
    fcAGIPVEX = "264"
    fcDECOVEX = "265"
    fcANAGRIM = "266"
    fcAGIPPIM = "268"
    fcMOVRAGO = "269"
    fcPROCSAT = "275"       ' (PM) Procedure KNOSOS lato server
    fcMOVRAGI = "2610"
    fcINSORD2 = "2612"      ' (DB) 25/08/2003
    fcEXPVIA1 = "2613"      ' (DB) 25/08/2003
    fcRICORD3 = "2614"      ' (DB) 25/08/2003
    fcEXPSNPV = "2615"      ' (DB) 25/08/2003
    fcIMPORD3 = "2622"      ' (PM) 21/06/2003
    fcIMPVVT1 = "2623"      ' (PM) 21/06/2003
    fcRICTICK = "2624"      ' (PM) 21/06/2003
    fcEXPSNP2 = "2625"      ' (PM) 21/06/2003
    fcIMPPRCE = "2617"
    fcEXPDICE = "2618"
    fcEXPNCON = "2619"
    fcEXPSNPVA = "2627"      ' (PM) Trasm. dati gestione anomalie x Agip
    fcORESFRO = "3622"
    fcANAGAMP = "31101"
    fcANAGAIV = "31102"
    fcANAGATC = "41951"
    fcANAGATD = "41952"
    fcANAGATM = "41953"
    fcANAGATS = "41954"
    fcANAGATV = "41955"
    fcANAGATN = "41956"
    fcANAGATE = "41957"
    fcANAGAME = "41958"
    fcANAGAZO = "31103"
    fcANAGAVA = "31104"
    fcANAGADD = "31108"      ' (PM) 06/03/06
    fcANAGAMF = "31109"      ' (PM) 06/03/06
    fcANAGABS = "311010"
    fcANAGAXX = "41971"
    fcANAGANU = "41972"
    fcANAGAXT = "41973"
    fcDOCUNDA = "41974"
    fcANAGARC = "41961"
    fcANAGARD = "41962"
    fcANAGARP = "41963"
    fcANAGARR = "41964"
    fcANAGARO = "41965"
    fcANAGAFP = "41966"
    ' --- Documenti Amministrativi
    fcDAMINSE = "291"
    fcDAMSEL1 = "292"
    fcDAMSEL2 = "293"
    fcDAMDAR0 = "404444"
    fcDAMDAM1 = "404443"
    
    fcTARITFC = "2169"      ' (PM) 06/03/06
    fcEXPSNPA = "276"       ' (PM) 23/06/06
    fcEXPSNPA2 = "277"
    fcIMPDEUS = "278"
    fcPOSOES = "279"        'SM 22/02/10
    fcGENGES = "2710"       'SM 24/02/10
    fcOCTTQ = "2630"
    fcOCTLOG = "2631"
    fcGENINT = "2632"
    
End Enum

Public Enum CodiceInstallazione
    ciInst1 = 1 '"001"
    ciInst2 = 2 '"002"
    ciInst3 = 3 '"003"
    ciInst4 = 4 '"004"
End Enum

Public Const MAX_RECORDS = 999
Public Const MAX_DBSIZE_DEMOMODE_JET = 1500000
Public Const MAX_DBSIZE_DEMOMODE_SQL = 3000000
Public Const ID_SOC = "001"
Public Const DATE_MASK = "__/__/__"
Public Const TIME_MASK = "__:__"
Public Const FLD_DELIMITER = "|"
Public Const MSG_CONCUR_ERR = "I dati sono cambiati. Vuoi salvare le tue modifiche ?" & vbNewLine & _
                               vbNewLine & _
                              "Premi OK per sovrascrivere i dati nel database" & vbNewLine & _
                              "Premi Cancel per annullare le tue modifiche e visualizzare i dati aggiornati"


Public Enum Errors
    NoError = -1
    FileNotFound = vbObjectError + 513
    InvalidHeader = vbObjectError + 514
    InvalidCode = vbObjectError + 515
    ParameterNotFound = vbObjectError + 516
    MissingDefaults = vbObjectError + 517
    NoImportData = vbObjectError + 518
    OutputError = vbObjectError + 519
    NoRigaOrdine = vbObjectError + 520
    RowsNotFound = vbObjectError + 521
    NoSetState = vbObjectError + 522
    MissingQuantity = vbObjectError + 523
    NoTipoVendita = vbObjectError + 524
    NoZona = vbObjectError + 525
    NoDeposito = vbObjectError + 526
    NoEnteEmittente = vbObjectError + 527
    NoProdotto = vbObjectError + 528
    NoMagazzino = vbObjectError + 529
    noUT = vbObjectError + 530
    NoDestinazione = vbObjectError + 531
    BadIDOrderSoc = vbObjectError + 532
    DestInsertError = vbObjectError + 533
    DestSocInsertError = vbObjectError + 534
    OrderInsertError = vbObjectError + 535
    RowOrderInsertError = vbObjectError + 536
    TripInsertError = vbObjectError + 537
    RowTripInsertError = vbObjectError + 538
    RowProductInsertError = vbObjectError + 539
    BadIDTripMover = vbObjectError + 540
    NoOrderID = vbObjectError + 541
    NoCliente = vbObjectError + 542
    OrderUpdateError = vbObjectError + 543
    TripDefinitionError = vbObjectError + 544
    NoProductInParam = vbObjectError + 545
    TripCopyError = vbObjectError + 546
    NoMotrice = vbObjectError + 547
    NoRimorchio = vbObjectError + 548
    ZonaEsclusaImport = vbObjectError + 549
    ZonaEsclusaExport = vbObjectError + 550
    NoTripID = vbObjectError + 551
    GP1InsertError = vbObjectError + 552
    DeliveryUpdateError = vbObjectError + 553
    NoPosizioneDoganale = vbObjectError + 554
    OrderAlreadyExisting = vbObjectError + 555
    NoTipoAutomezzo = vbObjectError + 556
'    DestinationAlreadyExisting = vbObjectError + 557
    TripDeleteError = vbObjectError + 558
    InsertError = vbObjectError + 559
    NoFasciaTariffaria = vbObjectError + 560
    FasciaTariffariaError = vbObjectError + 561
    NoPrefatturaEsterna = vbObjectError + 562
    PrintError = vbObjectError + 563
    InvalidFileName = vbObjectError + 564
    InvalidAddress = vbObjectError + 565
    OrderNotActive = vbObjectError + 566
    DestHasChanged = vbObjectError + 567
    DeleteError = vbObjectError + 568
    UpdateError = vbObjectError + 569
    OrderInTrip = vbObjectError + 570
    LoadError = vbObjectError + 571
    DeliveryError = vbObjectError + 572
    InvalidDep = vbObjectError + 573
    NoAutista = vbObjectError + 574
    InvalidDetail = vbObjectError + 575
    InvalidDate = vbObjectError + 576
    InvalidFooter = vbObjectError + 577
    SelectError = vbObjectError + 578
    CounterNotFound = vbObjectError + 579
    NoMotivoEvento = vbObjectError + 580
    NoRagDeposito = vbObjectError + 581
    Invalidmattract = vbObjectError + 582
    UnknownError = vbObjectError + 999
End Enum

Public Const ID_SOC_CEPSA = "038"
Public Const ID_SOC_CLH = "901"
Public Const ID_SOC_AGIP = "001"
Public Const ID_SOC_GUARNER = "A05"
Public Const ID_SOC_TQ = "899"
Public Const ID_SOC_LOG = "056"

Public Const ID_AVE_CEPSA = "041883"
Public Function CheckDBSize(Engine As Byte) As Boolean
    On Error GoTo ErrHandler
    Dim lngDBSize As Long
    
    Select Case Engine
        Case DEJET
'            lngDBSize = FileLen(DA.ConnDB)
            lngDBSize = FileLen(DataConnect.ConnDB)
            If lngDBSize > MAX_DBSIZE_DEMOMODE_JET Then
                GoTo ErrHandler
            End If
        Case DESQLSRV
    End Select
    CheckDBSize = True
    Exit Function
ErrHandler:
    CheckDBSize = False
End Function


Private Function CheckLicense() As Boolean
    On Error GoTo ErrHandler
    Dim strPWD As String
    Dim strKey As String
    Dim intYear As Integer
    
'    Select Case gmstrCodInstallazione
'        Case ciInst2
    #If (VersionID = 2) Then ' Or (VersionID = 5) Then
            intYear = Year(Now)
            If intYear >= 2000 Then
                If intYear > 2009 Then
                    GoTo ErrHandler
                End If
                
                ' --- Cerco la chiave corrispondente all'anno in corso
                strPWD = "RES" & CStr(intYear)
                strKey = gmRegistry.GetSetting("General", strPWD, "", True)
                If Len(strKey) = 0 Then
                    ' --- Se la chiave non esiste richiedo la password
                    strKey = InputBox("Inserisci la password per l'anno " & intYear, "Controllo della licenza")
                    If DecryptPassword(Right(CStr(intYear), 1), strKey) <> strPWD Then
                        GoTo ErrHandler
                    End If
                    ' --- Se la password inserita è corretta la memorizzo e ci
                    ' --- rivediamo l'anno prossimo
                    gmRegistry.SaveSetting "General", strPWD, strKey, True
                Else
                    ' --- Se la password memorizzata non è corretta, significa che è
                    ' --- stata modificata a mano nel registro di configurazione.
                    ' --- Studiare un'adeguata pena corporale.
                    If DecryptPassword(Right(CStr(intYear), 1), strKey) <> strPWD Then
                        GoTo ErrHandler
                    End If
                End If
            End If
        #End If
'    End Select
    CheckLicense = True
    Exit Function
ErrHandler:
    MsgBox "La password per l'anno " & intYear & " non è corretta", vbCritical + vbOKOnly, "Controllo della licenza"
    CheckLicense = False
End Function

Public Sub CleanUp()
    SetOptions
    Set ClientControls = Nothing
'    Set DA = Nothing
    Set DataConnect = Nothing
    Set frmSearch = Nothing
    Set gmRegistry = Nothing
End Sub

Public Sub DisplayError(mvarError As CError)
    Dim i As Integer
    
    Load frmDisplay
    
    With frmDisplay
        .lblErrCode = mvarError.Code - vbObjectError
        .lblErrDesc = mvarError.MainDescription
        .lblModule = mvarError.SourceModule
        .lblRoutine = mvarError.SouceProc
        .lblSeverity = mvarError.Severity
        .txtInfo = mvarError.info
        .lblNative = mvarError.NativeCode
        
        For i = 1 To mvarError.Descriptions.Count
            .txtSuppInfo = .txtSuppInfo & mvarError.Descriptions(i).Description & vbNewLine
        Next i
        .Show vbModal
    End With
End Sub


Public Function CheckExit(State As FormState) As Boolean
    Dim blnExit As Boolean
    
    blnExit = True
    Select Case State
        Case fsInsert
            blnExit = MsgBox("L'inserimento non verrà completato. Uscire?", vbYesNo + vbExclamation, "Inserimento in corso") = vbYes
        Case fsEdit
    End Select
    CheckExit = blnExit
End Function

Public Function CRPrint(ReportsList As Variant, _
                        SelectionFormula As String, _
                        ParentForm As Form, _
                        Optional SortFields As Variant = Empty, _
                        Optional Formulas As Variant = Empty, _
                        Optional ParamDescription As String) As Boolean
    
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim LinkForm As Form
    
    If Not IsLoaded("frmPrintReport") Then
        Set LinkForm = DisplayForm(fcPRNFORM)
        If LinkForm Is Nothing Then
            GoTo ErrHandler
        End If
        With LinkForm
            Set .ParentForm = ParentForm
            .LoadReportsList ReportsList, SelectionFormula, SortFields, Formulas, ParamDescription
            .ZOrder 0
            ParentForm.Enabled = False
        End With
    End If
    CRPrint = True
    GoTo EndProc

ErrHandler:
    CRPrint = False
EndProc:
End Function

Public Function EncryptPassword(Number As Byte, _
                                DecryptedPassword As String)
                                
    Dim Password As String
    Dim Counter As Byte
    Dim Temp As Integer
    
    Counter = 1
    Do Until Counter = Len(DecryptedPassword) + 1
        Temp = ASC(Mid(DecryptedPassword, Counter, 1))
        If Counter Mod 2 = 0 Then
            Temp = Temp - Number
        Else
            Temp = Temp + Number
        End If
        Temp = Temp Xor (10 - Number)
        Password = Password & Chr$(Temp)
        Counter = Counter + 1
    Loop
    EncryptPassword = Password
End Function

Public Function DecryptPassword(Number As Byte, _
                                EncryptedPassword As String)
                                
    Dim Password As String
    Dim Counter As Byte
    Dim Temp As Integer
    
    Counter = 1
    Do Until Counter = Len(EncryptedPassword) + 1
        Temp = ASC(Mid(EncryptedPassword, Counter, 1))
        Temp = Temp Xor (10 - Number)
        If Counter Mod 2 = 0 Then
            Temp = Temp + Number
        Else
            Temp = Temp - Number
        End If
        Password = Password & Chr$(Temp)
        Counter = Counter + 1
    Loop
    DecryptPassword = Password
End Function

Public Sub GetCalendarDate(M As MaskEdBox, strTitle As String)
    On Error GoTo ErrHandler
    
    Dim D As Date
    
    If M <> DATE_MASK Then
        D = CVDate(M)
    Else
        D = Now
    End If
    If frmCalendar.GetDate(D, strTitle) Then
        M = Format$(D, "dd\/mm\/yy")
        M.SetFocus
    End If
    Exit Sub
ErrHandler:

End Sub

Private Sub GetOptions()
    With gmRegistry
        gmblnMultipleWindows = CBool(.GetSetting("General", "MultipleWindows", "-1", True))
        gmstrPicFileName = .GetSetting("General", "BackgroundPicture", , True)
        gmstrReportsPath = .GetSetting("General", "ReportsPath", App.Path & "\reports", True)
    End With
End Sub

Public Function IsLoaded(FormName As String) As Boolean
    Dim x As Form
    
    FormName = UCase$(FormName)
    For Each x In Forms
        If UCase$(x.Name) = FormName Then
            IsLoaded = True
            GoTo EndProc
        End If
    Next x
    IsLoaded = False
EndProc:
    
End Function

Public Function IsUserGrandesCuentas() As Boolean
    Dim strNomeUtente As String
    
    strNomeUtente = UCase(Left(DataConnect.ConnUserName, 2))
    IsUserGrandesCuentas = (strNomeUtente = "GC") Or (strNomeUtente = "CP") Or (strNomeUtente = "TP") Or (strNomeUtente = "EU")
End Function

Public Sub SetBackPic()
    On Error GoTo ErrHandler
    
    If Len(gmstrPicFileName) > 0 Then
        If Len(Dir$(gmstrPicFileName)) > 0 Then
            RESOIL2MAIN.Picture = LoadPicture(gmstrPicFileName)
        End If
    End If
    Exit Sub
ErrHandler:
End Sub

Public Sub SetMenuConfiguration()
    With RESOIL2MAIN
        #If VersionID = 1 Then 'ciInst1 Then
'        Select Case gmstrCodInstallazione
'            Case ciInst1
            .mnul26(6).Enabled = False
'            .mnuL26(5).Visible = False
            .mnul26(5).Caption = "6.&5 EUROSYSTEM - Trasmissione Consegne"
            .mnuL313(3).Visible = True
        #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
'            Case ciInst2
            .mnuL1(4).Enabled = False
            .mnuL1(5).Enabled = False
            .mnuL21(9).Enabled = False
            .mnuL23(4).Enabled = False
            .mnul26(2).Enabled = False
            .mnul26(3).Enabled = False
            .mnul26(4).Caption = "6.&4 Trasmissione Viaggi a SAP"
            .mnul26(8).Enabled = False
            .mnul26(9).Enabled = False
            .mnul26(10).Enabled = False
            .mnuL322(3).Enabled = False
            .mnuL313(3).Visible = False
        #ElseIf VersionID = 3 Then
            .mnuL1(4).Enabled = False
            .mnuL21(9).Enabled = False
            .mnul26(2).Enabled = False
            .mnul26(3).Enabled = False
            .mnul26(5).Visible = False
            .mnul26(6).Enabled = False
            .mnul26(8).Enabled = False
            .mnul26(9).Enabled = False
            .mnul26(10).Enabled = False
            .mnuL322(3).Enabled = False
            .mnuL313(3).Visible = False
        #ElseIf VersionID = 4 Then
            .mnuL1(4).Enabled = False
            .mnuL1(5).Enabled = False
            .mnuL21(9).Enabled = False
            .mnul26(5).Visible = False
            .mnul26(6).Enabled = False
            .mnul26(8).Enabled = False
            .mnuL313(3).Visible = False
        #End If
'        End Select
    End With
End Sub

Public Sub SetMenuState(NewState As Boolean)
    Dim i As Integer
    
    For i = 1 To 6
        RESOIL2MAIN.mnuL1(i).Enabled = NewState
    Next i
End Sub

Public Sub SetOptions()
    With gmRegistry
'        .SaveSetting "Database", "DBName", DA.ConnDB, True
'        .SaveSetting "Database", "Server", DA.ConnServer, True
'        .SaveSetting "Database", "UID", DA.ConnUserName, True
'        .SaveSetting "Database", "Engine", CStr(DA.DataEngine), True
        .SaveSetting "Database", "DBName", DataConnect.ConnDB, True
        .SaveSetting "Database", "Server", DataConnect.ConnServer, True
        .SaveSetting "Database", "UID", DataConnect.ConnUserName, True
        .SaveSetting "Database", "Engine", CStr(DataConnect.DataEngine), True
        .SaveSetting "General", "MultipleWindows", IIf(gmblnMultipleWindows, "-1", "0"), True
        .SaveSetting "General", "BackgroundPicture", gmstrPicFileName, True
        .SaveSetting "General", "ReportsPath", gmstrReportsPath, True
    End With
End Sub

Public Sub DoSearch(Entity As String, Frm As Form, _
                    FieldKey As String, FieldTabIndex As Integer)
    Dim DC As Object
    Dim blnShowData As Boolean
    Dim blnShowPrint As Boolean
    Dim varReportsList As Variant
    
    Select Case Entity
        Case "AAM"
            Set DC = New CAutomezzo
        Case "AAU"
            Set DC = New CAutista
        Case "ABS"
            Set DC = New CBase
            blnShowData = True
        Case "ACL"
            Set DC = New CCliente
        Case "ACN"
            Set DC = New CContratto
            blnShowData = True
        Case "ACD"
            Set DC = New CContrattoDettaglio
        Case "ADD"              ' (PM) - 06/03/06
            Set DC = New CAddebito
            blnShowData = True
        Case "ADZ"
            Set DC = New CDestinazione
            SetRelationCond "ACL", DC, Frm
            ReDim varReportsList(0 To 2, 0 To 0) As Variant
            varReportsList(0, 0) = "Stampa elenco Destinazioni"
            varReportsList(1, 0) = "oil2adz1.rpt"
            varReportsList(2, 0) = 1
            blnShowPrint = True
        Case "ADE"
            Set DC = New CDeposito
'            SetRelationCond "ASC", DC, Frm
            blnShowData = True
        Case "AEE"
            Set DC = New CEnteEmittente
            blnShowData = True
        Case "AFP"
            Set DC = New CRagFisProdotto
            blnShowData = True
        Case "AIM"
            Set DC = New CImballo
            blnShowData = True
        Case "AIV"
            Set DC = New CIVA
            blnShowData = True
        Case "AMD"
            Set DC = New CMagazzinoDeposito
            SetRelationCond "ADE", DC, Frm
            blnShowData = True
        Case "AME"
            Set DC = New CMotivoEvento
            SetRelationCond "ASC", DC, Frm
            blnShowData = True
        Case "AMP"
            Set DC = New CPagamento
            blnShowData = True
        Case "APD"
            Set DC = New CPosDoganale
            blnShowData = True
        Case "ARC"
            Set DC = New CRagCliente
            blnShowData = True
        Case "APR"
            Set DC = New CProdotto
            blnShowData = True
        Case "ARD"
            Set DC = New CRagDestinazione
            blnShowData = True
        Case "ARO"
            Set DC = New CRagOrdine
            blnShowData = True
        Case "ARP"
            Set DC = New CRagDeposito
            blnShowData = True
        Case "ARR"
            Set DC = New CRagProdotto
            blnShowData = True
        Case "ASC"
            Set DC = New CSocieta
            blnShowData = True
        Case "ATC"
            Set DC = New CTipoCliente
            blnShowData = True
        Case "ATD"
            Set DC = New CTipoDestinazione
            blnShowData = True
        Case "ATE"
            Set DC = New CTipoEvento
            SetRelationCond "ASC", DC, Frm
            blnShowData = True
        Case "ATM"
            Set DC = New CTipoAutomezzo
            blnShowData = True
        Case "ATN"
            Set DC = New CTipoContratto
            blnShowData = True
        Case "ATS"
            Set DC = New CTipoCisterna
            blnShowData = True
        Case "ATV"
            Set DC = New CTipoVendita
            blnShowData = True
        Case "AUM"
            Set DC = New CUnitaDiMisura
            blnShowData = True
        Case "AUT"
            Set DC = New CUnitaDiTrasporto
            ReDim varReportsList(0 To 2, 0 To 0) As Variant
            varReportsList(0, 0) = "Stampa elenco Unità di Trasporto"
            varReportsList(1, 0) = "oil2aut1.rpt"
            varReportsList(2, 0) = 1
            blnShowPrint = True
            blnShowData = True
        Case "AVE"
            Set DC = New CVettore
        Case "AZO"
            Set DC = New CZona
            blnShowData = True
        Case "AVA"
            Set DC = New CValuta
            blnShowData = True
        Case "AXT"
            Set DC = New CParametriTariffe
            blnShowData = True
        Case "PPT"
            Set DC = New CPrefattura
            blnShowData = True
        Case "PAD"
            Set DC = New CPrefatturaDettaglioAddebito
        
        Case "TTC"  ' (PM) - 06/03/06
            Set DC = New CTariffaTipoCalcolo
            blnShowData = True
        Case "TTP"
            Set DC = New CTipologiaProdotto
            blnShowData = True
        Case "AMF"  ' (PM) - 06/03/06
            Set DC = New CModalitaFatturazione
            blnShowData = True
        
        ' (PM) - 03/02/06
        Case "ATO"
            Set DC = New CTipoDocumento
        Case "DAM"
            Set DC = New CDocumentoAmministrativo
        Case "DAR"
            Set DC = New CRigaDocumentoAmministrativo

        Case Else
            Set DC = Nothing
    End Select
    If Not DC Is Nothing Then
        With DC.DS
            Set .FormHandle = Frm
'            Set .DA = DataConnect.DA
            Set .DA = DataConnect.DA
        End With
        ShowSearchForm Frm, FieldKey, DC, FieldTabIndex, blnShowData, blnShowPrint, varReportsList
        Set DC = Nothing
    End If
End Sub

Sub Main()
    On Error GoTo ErrHandler

    Const NR_MAX_PARAM As Integer = 15
    Const NR_MAX_ATTEMPT_CONNECT As Integer = 3
    
    Dim varWindowsState As Variant
    Dim varCommandLine As Variant
    Dim intNumArgs As Integer
    Dim i As Integer
    Dim blnLetMultistance As Boolean
    ReDim varRunFunction(1 To 1) As Variant
    Dim lngRunFunctionCount As Long
    Dim blnAutoLogon As Boolean
    Dim blnAutoConnectOk As Boolean
    
    intNumArgs = NR_MAX_PARAM
    varWindowsState = vbMaximized
    blnAutoLogon = False
    gmstrLogFileName = vbNullString
    
    If GetCommandLine(varCommandLine, intNumArgs) Then
        If Not IsEmpty(varCommandLine) Then
            For i = 1 To intNumArgs
                Select Case UCase(varCommandLine(i))
                    Case "-M"
                        blnLetMultistance = True

                    Case "-L"
                        blnAutoLogon = True
                    
                    Case Else
                        If Left(UCase(varCommandLine(i)), 2) = "-K" And Len(varCommandLine(i)) > 2 Then
                            gmstrLogFileName = CStr(Mid(varCommandLine(i), 3))
                        Else
                            If Left(UCase(varCommandLine(i)), 2) = "-S" And Len(varCommandLine(i)) > 2 Then
                                gmstrServerLogFileName = CStr(Mid(varCommandLine(i), 3))
                            Else
                                ' Eseguire una function
                                If Left(UCase(varCommandLine(i)), 2) = "-F" And Len(varCommandLine(i)) > 2 Then
                                    lngRunFunctionCount = lngRunFunctionCount + 1
                                    ReDim Preserve varRunFunction(1 To lngRunFunctionCount) As Variant
                                    varRunFunction(lngRunFunctionCount) = Mid(varCommandLine(i), 3)
                                End If
                            End If
                        End If
                End Select
            Next i
        End If
    End If
    
    If Not blnLetMultistance Then
        If App.PrevInstance Then
            End
        End If
    End If
    
    Dim strLicenseTo As String
    Dim strLicenseName As String
    Dim strGetParam As String
    
    Set gmRegistry = New CRegSettings
    
    If Not GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYCOMMONFILES, gREGVALCOMMONFILES, CommonFilesDir) Then
    End If
    
    GetLicenseInfo strLicenseTo, strLicenseName
    gmblnDemoMode = Len(gmRegistry.GetSetting("General", "RES1999", "", True)) > 0
    ID_PARAM = gmRegistry.GetSetting("General", "IDParam", "001", True)
    
    With frmSplash
        .lblLicenseTo = .lblLicenseTo & " " & strLicenseTo
        .lblLicenseName = .lblLicenseName & " " & strLicenseName
        If gmblnDemoMode Then
            .lblWarning = "Versione dimostrativa"
            .lblWarning.Visible = True
        End If
        .Show
    End With
    
    Set ClientControls = New CControls
'    Set DA = New CDataAccess
    Set DataConnect = New CConnect
    GetOptions
    
    With frmLogon
        .Caption = App.ProductName & " - LOGON"
        .cmbProvider.ListIndex = CInt(gmRegistry.GetSetting("Database", "Engine", "0", True))
        .txtDatabase = gmRegistry.GetSetting("Database", "DBName", , True)
        .txtServer = gmRegistry.GetSetting("Database", "Server", , True)
        If blnAutoLogon Then
            .txtUID = gmRegistry.GetSetting("Database", "AutoLogonUser", , True)
            .txtPWD = gmRegistry.GetSetting("Database", "AutoLogonPwd", , True)
        Else
            .txtUID = gmRegistry.GetSetting("Database", "UID", , True)
            .txtPWD = ""
        End If
    End With

    Do While frmSplash.Visible
        DoEvents
    Loop
    
    DataConnect.DA.LogFileName = gmstrLogFileName
    If blnAutoLogon Then
        If frmLogon.txtUID <> vbNullString Then
            ' Provo a connettermi facendo un numero massimo di tentativi
            For i = 1 To NR_MAX_ATTEMPT_CONNECT
                If frmLogon.DoConnect() Then
                    blnAutoConnectOk = True
                    Exit For
                Else
                    ' Prendo tempo: forse si sta avviando il servizio del Server SQL
                    DoEvents
                    DoEvents
                    DoEvents
                End If
            Next i
        End If
        If Not blnAutoConnectOk Then
            frmLogon.txtUID = gmRegistry.GetSetting("Database", "UID", , True)
            frmLogon.txtPWD = ""
            frmLogon.Show
        End If
    Else
        frmLogon.Show
    End If
    
    Do While (Not gmblnConnectOk) And (Not gmblnCancelConnect)
        DoEvents
    Loop

    If Not gmblnCancelConnect Then
        If gmblnDemoMode Then
'            If Not CheckDBSize(DA.DataEngine) Then
            If Not CheckDBSize(DataConnect.DataEngine) Then
                MsgBox "Il database ha raggiunto la dimensione massima consentita per questa versione dell'applicazione", _
                       vbCritical + vbOKOnly
                GoTo ErrHandler
            End If
        End If
        
        With RESOIL2MAIN
            If Not CheckLicense Then
                GoTo ErrHandler
            End If
            SetBackPic
            SetMenuConfiguration
            .AutoShowChildren = False
'            .StatusBar1.Panels(2).Text = "Database: " & DA.ConnDB
            .StatusBar1.Panels(2).Text = "Database: " & DataConnect.ConnDB & " Server: " & DataConnect.ConnServer
            .Show
            Call AutoRunFunction(DataConnect, varRunFunction)
        End With
        
        DataConnect.DA.ExecuteCommand "set dateformat DMY"
    End If
    Unload frmSplash
    
    
'    Dim SDT As CStipendioTestata
'    Set SDT = New CStipendioTestata
'    Dim ErrH As New CErrorHandler
'    With SDT.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    SDT.CreateStipendio "J55", "20/04/2007", "01/03/2007", "31/03/2007"
    Exit Sub
ErrHandler:
    CleanUp
    Set frmSplash = Nothing
    Set RESOIL2MAIN = Nothing
    End
End Sub

Public Sub TestViaggioSAT()

        Dim Service As VLocalize.VLocalize
        Set Service = New VLocalize.VLocalize

        ' -- DP 10/02/2009
        'Service.serviceLocation = "http://www.tphweb.es/mbtph/mbws/mbws.asmx"
        Service.serviceLocation = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
        Service.serviceTimeOut = 5

        Service.serviceUsername = "wstph"
        Service.servicePassword = "swtph"

        Dim varObjDataInput(0 To 1, 0 To 20) As Variant

        'Campi Riga Come da tracciato
        Dim rowField1 As Variant
        Dim rowField2 As Variant
        Dim rowField3 As Variant
        Dim rowField4 As Variant
        Dim rowField5 As Variant

        rowField1 = Array("1", "0075116", "3", "Viaje 0065860", "", "13/12/2007 14:00:00", "13/12/2007 16:19:00", "43", "Carga", "S", "13/12/2007 14:00:00", "13/12/2007 14:28:00", "021", "", "", "", "", "", "", "", "")
        rowField2 = Array("2", "0075116", "3", "Viaje 0065860", "", "13/12/2007 14:00:00", "13/12/2007 16:19:00", "43", "Entrega 1", "N", "", "", "G0090680", "2429499324", "GASOLA", "17000", "0075116010", "13/12/2007", "14:05:00", "13/12/2007", "")
'        rowField3 = Array("3", "0063754", "004", "Viaje 04", "Viaje 04 Test", "05/12/2007 13:00:00", "05/12/2007 13:31:00", "95", "Entrega 2", "N", "", "", "CPF3203564", "0175384427", "GASOLC", "5000", "0063754", "05/12/2007", "10:19:00", "05/12/2007", "10:19:00")
'        rowField4 = Array("4", "0063754", "004", "Viaje 04", "Viaje 04 Test", "05/12/2007 13:00:00", "05/12/2007 13:31:00", "95", "Entrega 3", "N", "", "", "F0503114", "0175386740", "GASOLC", "5000", "0063754", "05/12/2007", "10:50:00", "05/12/2007", "10:50:00")
'        rowField5 = Array("5", "0063754", "004", "Viaje 04", "Viaje 04 Test", "05/12/2007 13:00:00", "05/12/2007 13:31:00", "95", "Entrega 4", "N", "", "", "CPF3247908", "0175386888", "GASOLC", "3000", "0063754", "05/12/2007", "13:11:00", "05/12/2007", "13:11:00")

        Dim i As Integer
        For i = 0 To 20
            varObjDataInput(0, i) = rowField1(i)
            varObjDataInput(1, i) = rowField2(i)
'            varObjDataInput(2, i) = rowField3(i)
'            varObjDataInput(3, i) = rowField4(i)
'            varObjDataInput(4, i) = rowField5(i)
        Next i
        
        Dim varResults As Variant '(,) '= Nothing

    On Error GoTo ErrHandler
        
            Service.Open
            If Service.CreaViaggio(varObjDataInput, varResults) Then
                '//***
                Dim ifx As Integer
                Dim strx As String
                strx = ""
                For ifx = 0 To UBound(varResults, 1)
                    strx = strx & ">>>|" & varResults(ifx, 0) & "|" & varResults(ifx, 1) & "|" & varResults(ifx, 2)
                Next ifx
                Debug.Print strx
                MsgBox "ESEGUITO"
            Else
                MsgBox "NON ESEGUITO"
            End If
            Service.Close
            Exit Sub
ErrHandler:
            Service.Close
            MsgBox Err.Description, , Err.Source
            Exit Sub
End Sub
Private Sub AutoRunFunction(ByVal DC As CConnect, ByVal varBufFunction As Variant)
    Dim varBuf As Variant
    Dim i As Long
    Dim F As Form
    Dim intFlDoc As Integer
    Dim lngPrintPolling As Long
    On Error GoTo ErrHandler
           
    ' Visualizzo le finestre richieste da riga di comando
    If Not IsEmpty(varBufFunction(1)) Then
        For i = LBound(varBufFunction) To UBound(varBufFunction)
            DisplayForm CStr(varBufFunction(i))
        Next i
    End If
Exit Sub
ErrHandler:
End Sub

Public Function GetCommandLine(ByRef rCommandLine As Variant, Optional ByVal MaxArgs As Integer = 10) As Boolean
    'Dichiara le variabili.
    Dim i, CmdLine, NumArgs As Integer
    Dim InArg As Boolean
    Dim C, CmdLnLen As String
    'Imposta le dimensioni giuste per la matrice.
    ReDim ArgArray(1 To MaxArgs) As Variant
    On Error GoTo ErrHandler
    
    ' Inizializza l'array
    For i = 1 To MaxArgs
        ArgArray(i) = vbNullString
    Next i
    NumArgs = 0
    InArg = False
    'Legge gli argomenti della riga di comando.
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    'Si sposta nella riga di comando un carattere alla volta.
    For i = 1 To CmdLnLen
        C = Mid(CmdLine, i, 1)
        'Controlla se vi sono spazi o tabulazioni.
        If (C <> " " And C <> vbTab) Then
            'Nessuno spazio o tabulazione.
            'Controlla se è già nell'argomento.
            If Not InArg Then
                'Inizia un nuovo argomento.
                'Controlla se vi sono troppi argomenti.
                If NumArgs = MaxArgs Then
                    Exit For
                End If
                NumArgs = NumArgs + 1
                InArg = True
            End If
            'Concatena un carattere all'argomento corrente.
            ArgArray(NumArgs) = ArgArray(NumArgs) & C
        Else
            'Ha trovato uno spazio o una tabulazione.
            'Imposta il flag InArg su False.
            InArg = False
        End If
    Next i
    'Restituisce la matrice
    If NumArgs = 0 Then
        rCommandLine = Empty
    Else
        rCommandLine = ArgArray
    End If
    GetCommandLine = True
Exit Function
ErrHandler:
    rCommandLine = Empty
    GetCommandLine = False
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (ASC(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(ASC(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function



Public Function GetLicenseInfo(ByRef strCompany As String, ByRef strUser As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strVersion As String
    
    strVersion = App.major & "." & App.minor & "." & App.Revision
    strCompany = gmRegistry.GetSetting(strVersion, "Company", "RES Data srl", True)
    strUser = gmRegistry.GetSetting(strVersion, "Name", , True)
    GetLicenseInfo = True
    Exit Function
ErrHandler:
    GetLicenseInfo = False
End Function
Public Sub ClientKeyUp(KeyCode As Integer, F As Form)
    Select Case KeyCode
        Case vbKeyF1
            F.DoCommand fcAddNew
        Case vbKeyF2
            F.DoCommand fcSave
        Case vbKeyF3
            F.DoCommand fcExit
        Case vbKeyF4
            F.DoCommand fcUndo
        Case vbKeyF5
            F.DoCommand fcPrint
        Case vbKeyF6
            F.DoCommand fcOpen
        Case vbKeyF7
            F.DoCommand fcOpenLinkEdit
        Case vbKeyF8
            F.DoCommand fcLinkRefresh
        Case vbKeyF9
            F.DoCommand fcFind
        Case vbKeyF11
            F.DoCommand fcCustom1
        Case vbKeyF12
            F.DoCommand fccustom2
    End Select
End Sub

Public Function AddClientControl(ByVal Obj As Object, ByRef F As Form, Optional ByVal Key As String = "") As CControl
    Set AddClientControl = ClientControls.Add(Obj, F)
End Function

Public Sub ConfigureObjects(ByRef F As Form, _
                            Optional ColTags As Variant, _
                            Optional SetMultiSocTags As Boolean = False, _
                            Optional ColTags1 As Variant, Optional ColTags2 As Variant)
    Dim C As Control
    Dim ClControl As CControl
    
    For Each C In F
        If C.Tag <> vbNullString Then
            Select Case UCase(TypeName(C))
                Case "LABEL"
                    If UCase(C.Tag) <> "LINK" Then
                        C.Font.Bold = False
                        AddClientControl C, F
                    End If
                Case "COMMANDBUTTON"
                    AddClientControl C, F
                Case "MASKEDBOX"
                    AddClientControl C, F
                Case "COMBOBOX"
                    AddClientControl C, F
                Case "TDBGRID"
                    C.CellTips = dbgFloating
                    C.CellTipsDelay = 750
                    Set ClControl = AddClientControl(C, F)
                    With ClControl
                        Select Case C.Tag
                            Case "MULTISOC"
                                If (Not IsMissing(ColTags)) And (SetMultiSocTags) Then
                                    .SetColTags UBound(ColTags, 1)
                                    .ColTags = ColTags
                                End If
                                .SetMultiSocGrid F.Soc, Not SetMultiSocTags
                            Case "DETAIL", "INSDETAIL"
                                .SetColTags UBound(ColTags, 1)
                                .ColTags = ColTags
                            Case "DETAIL1"
                                If Not IsMissing(ColTags1) Then
                                    .SetColTags1 UBound(ColTags1, 1)
                                    .ColTags1 = ColTags1
                                End If
                            Case "DETAIL2"
                                If Not IsMissing(ColTags2) Then
                                    .SetColTags2 UBound(ColTags2, 1)
                                    .ColTags2 = ColTags2
                                End If
                            Case "SELECT"
                                If Not IsMissing(ColTags) Then
                                    .SetColTags UBound(ColTags, 1)
                                    .ColTags = ColTags
                                End If
                        End Select
                    End With
            End Select
        End If
    Next C
End Sub

Public Function DisplayForm(strFunctionCode As FunctionCodes, _
                            Optional ModalForm As Boolean = False) As Form
    On Error GoTo ErrHandler
    
    Dim F As Form
    Dim blnLocked As Boolean
    
    Select Case strFunctionCode
        Case fcANAGASC
            Set F = OIL2ASC0
        Case fcANAGAEE
            Set F = OIL2AEE0
        Case fcANAGADE
            Set F = OIL2ADE0
        Case fcANAGACL
            Set F = OIL2ACL0
        Case fcANAGADZ
            Set F = OIL2ADZ0
        Case fcANAGAME
            Set F = OIL2AME0
        Case fcANAGATE
            Set F = OIL2ATE0
        Case fcANAGABS
            Set F = OIL2ABS0
        Case fcORDINSE
            Set F = frmOCTInsert
        Case fcORDSELA
            Set F = frmOCTSelect
            F.FunctionCode = strFunctionCode
        Case fcORDSELV
            Set F = frmOCTSelect
            F.FunctionCode = strFunctionCode
        Case fcORDSELM
            Set F = frmOCTSelect
            F.FunctionCode = strFunctionCode
        Case fcORDSELE
            Set F = frmOCTSelect
            F.FunctionCode = strFunctionCode
        Case fcORDSEVA
            Set F = frmOCTSelect2
            F.FunctionCode = strFunctionCode
        Case fcORDFRAZ
            Set F = frmOCTSelect3
            F.FunctionCode = strFunctionCode
        Case fcORDRIDU
            Set F = frmOCTSelect4
        Case fcOCTFRAZ
            Set F = frmOCTFraz
        Case fcVIAINSE
            Set F = frmVVTInsert
        Case fcVIASELP
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
        Case fcVIASELT
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
        Case fcVIASELC
            #If VersionID = 1 Then
                Set F = frmVVTSelect
                F.FunctionCode = strFunctionCode
            #End If
        Case fcVIACOMP
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
        Case fcVIASELE
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
        Case fcVIASELD
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
        Case fcORDIOCT
            Set F = OIL2OCT0
        Case fcORDIOCR
            Set F = OIL2OCR0
        Case fcVIAGVVT
            Set F = OIL2VVT0
        Case fcVIAGVPI
            Set F = OIL2VPI0
        Case fcVIAGVVP
            Set F = OIL2VVP0
        Case fcVIAGVVC
            Set F = OIL2VVC0
        #If VersionID = 1 Then
        Case fcVIAGCOM
            Set F = frmVVTComp
        #End If
        Case fcPPTSELE
            Set F = frmPPTSelect
        Case fcGP1SELE
            Set F = frmGP1Select
        Case fcPPTCREA
            Set F = frmPPTCreate
        Case fcPPTCRE2
            Set F = frmPPTCreate2
        Case fcPPTCRE3
            Set F = frmPPTCreate3
        Case fcPREFPPT
            Set F = OIL2PPT0
        Case fcPREFPPD
            Set F = OIL2PPD0
        Case fcPREFPAD
            Set F = OIL2PAD0
        Case fcTARITFA
            Set F = OIL2TFA0
        Case fcTARITFB
            Set F = OIL2TFB0
        Case fcTARITTC                  ' (PM) - 06/03/06
            Set F = OIL2TTC0
        Case fcTARITFC                  ' (PM) - 06/03/06
            Set F = OIL2TFC0
        Case fcANAGAPR
            Set F = OIL2APR0
        Case fcANAGAIM
            Set F = OIL2AIM0
        Case fcANAGAPD
            Set F = OIL2APD0
        Case fcANAGAVE
            Set F = OIL2AVE0
        Case fcANAGAAU
            Set F = OIL2AAU0
        Case fcANAGAAM
            #If VersionID = 4 Then
                Set F = OIL2AAM1
            #Else
                Set F = OIL2AAM0
            #End If
        Case fcANAGAUT
            Set F = OIL2AUT0
        Case fcANAGACN
            Set F = OIL2ACN0
        Case fcANAGACD          ' (PM) 06/03/06
            Set F = OIL2ACD0
        Case fcANAGADD          ' (PM) 06/03/06
            Set F = OIL2ADD0
        Case fcANAGAMF          ' (PM) 06/03/06
            Set F = OIL2AMF0
        Case fcANAGAMP
            Set F = OIL2AMP0
        Case fcANAGAIV
            Set F = OIL2AIV0
        Case fcANAGATC
            Set F = OIL2ATC0
        Case fcANAGATD
            Set F = OIL2ATD0
        Case fcANAGATM
            Set F = OIL2ATM0
        Case fcANAGATS
            Set F = OIL2ATS0
        Case fcANAGATV
            Set F = OIL2ATV0
        Case fcANAGATN
            Set F = OIL2ATN0
        Case fcANAGAZO
            Set F = OIL2AZO0
        Case fcANAGAVA
            Set F = OIL2AVA0
        Case fcANAGANU
            Set F = OIL2ANU0
        Case fcANAGAXX
            Set F = OIL2AXX0
        Case fcANAGAXT
            Set F = OIL2AXT0
        Case fcANAGARC
            Set F = OIL2ARC0
        Case fcANAGARD
            Set F = OIL2ARD0
        Case fcANAGARP
            Set F = OIL2ARP0
        Case fcANAGARR
            Set F = OIL2ARR0
        Case fcANAGARO
            Set F = OIL2ARO0
        Case fcANAGAFP
            Set F = OIL2AFP0
        Case fcSTATFRM
            Set F = frmStats
        Case fcINSORD2                      ' (DB) 25/08/2003
            Set F = frmOCTImport2           ' (DB) 25/08/2003
        Case fcEXPVIA1                      ' (DB) 25/08/2003
            Set F = frmVVTExport1           ' (DB) 25/08/2003
        Case fcRICORD3                      ' (DB) 25/08/2003
            Set F = frmOCTImport3           ' (DB) 25/08/2003
        Case fcEXPSNPV                      ' (DB) 25/08/2003
            Set F = frmVVTSnapshotExport    ' (DB) 25/08/2003
            F.FunctionCode = strFunctionCode
        Case fcEXPSNPVA                     ' (PM) 10/06/06
            Set F = frmVVTSnapshotExport    ' (PM) 10/06/06
            F.FunctionCode = strFunctionCode
        Case fcIMPPRCE
            Set F = frmGP1ImportCEPSA
        Case fcEXPDICE
            Set F = frmGP1ExportCEPSA
        Case fcAGIPOIM
            #If (VersionID = 4) Or (VersionID = 5) Then
                Set F = frmOCTImport1
            #Else
                Set F = frmOCTImport
            #End If
        Case fcAGIPVEX
            Set F = frmVVTExport
            #If VersionID = 2 Then
                F.ExportType = etSAP
            #End If
        Case fcDECOVEX
            Set F = VVTExportESys
            
            #If VersionID = 2 Then
                Set F = frmVVTExport
                F.ExportType = etDECO
            #End If
        Case fcAGIPPIM
            Set F = frmGP1Import
        Case fcMOVEREX
            #If (VersionID = 4) Or (VersionID = 5) Then
                Set F = frmOCTExportMover1
            #Else
                Set F = frmOCTExportMover0
            #End If
        Case fcMOVERIM
            #If (VersionID = 4) Or (VersionID = 5) Then
                Set F = frmVVTImportMover1
            #Else
                Set F = frmVVTImportMover
            #End If
        Case fcMOVRAGO
            Set F = frmVVCExportMap
        Case fcMOVRAGI
            Set F = frmVVCImportMap
        Case fcPRNFORM
            Set F = frmPrintReport
        Case fcPREFGP1
            Set F = OIL2GP10
        Case fcPREZAPZ
            Set F = OIL2APZ0
        Case fcPREZASZ
            Set F = OIL2ASZ0
        Case fcPREZFPD
            Set F = OIL2FPD0
        Case fcVIASSAT                      ' (PM) - Viaggi Sat
            Set F = frmVVTSatellitareViewer ' (PM) - Viaggi Sat
        Case fcCONTVIA
            Set F = frmVVTSelect2
        Case fcPROCSAT
            Set F = frmSMSProcedure         ' Gestione SMS
        Case fcANAGRIE                      ' (PM) - Imp/Exp anagrafiche
            Set F = frmEU2KImpExp
        Case fcDAMINSE
            Set F = OIL2DAM1
            F.FunctionCode = strFunctionCode
        Case fcDAMSEL1
            Set F = DAMSelect1
        Case fcDAMSEL2
            Set F = DAMSelect2
        Case fcDAMDAR0
            Set F = OIL2DAR0
        Case fcDAMDAM1
            Set F = OIL2DAM1
            F.FunctionCode = strFunctionCode
        Case fcIMPORD3
            Set F = frmOCTImport4
            F.FunctionCode = strFunctionCode
        Case fcIMPVVT1                      ' (PM) 19/06/2006
            Set F = frmOCTExportEA
'            F.FunctionCode = strFunctionCode
        Case fcRICTICK                      ' (PM) 21/06/2003
            Set F = frmOCTImport3
            F.FunctionCode = strFunctionCode
        Case fcEXPSNP2                      ' (PM) 21/06/2003
            Set F = frmVVTSnapshotExport
            F.FunctionCode = strFunctionCode
        Case fcEXPSNPA                      ' (PM) 21/06/2003
            Set F = frmVVTSnapshotExport
            F.FunctionCode = strFunctionCode
        Case fcEXPSNPA2
            Set F = frmVVTSnapshotExport
            F.FunctionCode = strFunctionCode
        Case fcEXPNCON
            Set F = frmNENExport
        Case fcIMPDEUS
            Set F = frmSatellitareProcedure
        Case fcDOCUNDA
            Set F = OIL2NDA0
            F.FunctionCode = strFunctionCode
            
        Case fcSTIPINS
            Set F = frmSetDatiStipendi
            F.FunctionCode = strFunctionCode
        Case fcSTIPCRE
            Set F = frmSDTCreate
        Case fcSTIPSEL
            Set F = frmSDTSelect
        Case fcSTIPOLA
            Set F = OIL2OLA0
        Case fcPOSOES
            Set F = frmSatellitarePos
        Case fcGENGES
            Set F = frmGeneraGes
        Case fcGENINT
            Set F = frmVVTExportRes
        Case fcOCTTQ
            Set F = frmOCTImporttq
        Case fcOCTLOG
            Set F = frmOCTImportLog
        Case Else
            Exit Function
    End Select

    RESOIL2MAIN.MousePointer = vbHourglass

    ' --- Se l'opzione "Visualizza più finestre contemporaneamente"
    ' --- è disattivata disabilito i menu
    If Not gmblnMultipleWindows Then
        SetMenuState False
    End If
    ' ---
    
    If Not ModalForm Then
        blnLocked = True
        LockWindowUpdate F.hwnd
        F.ZOrder
        F.Icon = RESOIL2MAIN.Icon
        F.Show
        LockWindowUpdate 0
    Else
        F.Show 1
    End If
    
    Set DisplayForm = F
    
    With RESOIL2MAIN
        .Arrange vbCascade
        .MousePointer = vbDefault
    End With
    
    DoEvents

    Exit Function
ErrHandler:
    MsgBox Err.Description
    If blnLocked Then
        LockWindowUpdate 0
    End If
    RESOIL2MAIN.MousePointer = vbDefault
End Function


Public Sub FieldSetDescription(FieldKey As String, _
                               ByRef Frm As Form, _
                               strText As String, _
                               msk As MaskEdBox)
    Dim varRelatedEntity As Variant
    Dim DC As Object
    Dim i As Integer
    ReDim varText(0 To 0) As Variant
    Dim C As Control
    
    If strText <> vbNullString Then
        If (Frm.State = fsInsert) Or (Frm.State = fsEdit) Then
            With Frm.DataClass.DS
                ' --- Se la validazione è stata positiva, cerco le informazioni deducibili
                varText(0) = strText
                Select Case .Fields(FieldKey).ForeignTableEntity
                    Case "ADZ"
                        #If VersionID = 4 Then
                            varRelatedEntity = Array("ACL", "AEE", "ADE", "AZO", "AMP", "AIV", "ATV")
                        #Else
                            varRelatedEntity = Array("ACL", "ADE", "AZO", "AMP", "AIV")
                        #End If
                        Set DC = New CDestinazione
                    Case "AEE"
                        varRelatedEntity = Array("ASC")
                        Set DC = New CEnteEmittente
                    Case "ACL"
                        #If VersionID = 4 Then
                            varRelatedEntity = Array("ASC")
                        #Else
                            varRelatedEntity = Array("ASC", "AEE")
                        #End If
                        Set DC = New CCliente
                    Case "AAU"
                        varRelatedEntity = Array("AVE")
                        Set DC = New CAutista
                    Case "AMD"
                        ReDim varText(0 To 1) As Variant
                        varText(0) = vbNullString
                        For Each C In Frm.Controls
                            If UCase(TypeName(C)) = "MASKEDBOX" Then
                                If .Fields(C.Tag).ForeignTableEntity = _
                                    "ADE" Then
                                        varText(0) = C
                                    Exit For
                                End If
                            End If
                        Next C
                        varText(1) = strText
                        varRelatedEntity = Array("ATV")
                        Set DC = New CMagazzinoDeposito
                    Case "AUT"
'                        varRelatedEntity = Array("AAM", "AAU", "AVE", "ADE", "AEE", "ASC")
                        varRelatedEntity = Array("AVE", "ADE", "AEE", "ASC")
                        Set DC = New CUnitaDiTrasporto
                End Select
                ' ---
                
                If Not DC Is Nothing Then
                    With DC.DS
                        Set .FormHandle = Frm
'                        Set .DA = DataConnect.DA
                        Set .DA = DataConnect.DA
                    End With
                    Frm.MousePointer = vbHourglass
                    For i = LBound(varRelatedEntity) To UBound(varRelatedEntity)
                        FieldFindRelatedInformation CStr(varRelatedEntity(i)), DC, varText, Frm, msk
                    Next i
                    Frm.MousePointer = vbDefault
                    Set DC = Nothing
                End If
            End With
        End If
    End If
End Sub

Public Sub RowToRecord(ByVal WriteLocation As Long, _
                        ByRef DS As CDataStruct, _
                        ByRef GA As ArrayGridCls, _
                        Frm As Form, _
                        ColTags As Variant)
    Dim i As Integer
    Dim DSTemp As CDataStruct
    Dim intCount As Integer
    
    Set DSTemp = Frm.DataClass.DS
    With DS
        intCount = DSTemp.KeyFields.Count
        For i = 1 To intCount
            .Fields(.KeyFields(i)).NewValue = _
                DSTemp.Fields(DSTemp.KeyFields(i)).value
            .Fields(.KeyFields(i)).Updated = True
        Next i
'        .Fields(.KeyFields(1)).NewValue = _
'            DSTemp.Fields(DSTemp.KeyFields(1)).value
'        .Fields(.KeyFields(1)).Updated = True
        intCount = UBound(ColTags)
        For i = 1 To intCount
            If ColTags(i) <> vbNullString Then
                .Fields(ColTags(i)).NewValue = _
                    Trim(GA.value(WriteLocation, i - 1))
'                    Trim (UCase(GA.value(WriteLocation, i - 1)))
                .Fields(ColTags(i)).Updated = True
            End If
        Next i
    End With
End Sub

Public Sub FieldFindRelatedInformation(ForeignEntity As String, _
                                       DC As Object, _
                                       FindValue As Variant, _
                                       ByRef Frm As Form, _
                                       msk As MaskEdBox)
    On Error GoTo ErrHandler
    
    Dim C As Control
    Dim i As Integer
    Dim j As Integer
    Dim intFieldCount As Integer
    Dim DS As New CDataStruct
    
    Set DS = Frm.DataClass.DS
    
    For Each C In Frm.Controls
        If UCase(TypeName(C)) = "MASKEDBOX" Then
            If DS.Fields(C.Tag).ForeignTableEntity = _
                ForeignEntity Then
                
                With DC.DS
                    For i = 0 To UBound(FindValue)
                        If FindValue(i) = vbNullString Then
                            GoTo ErrHandler
                        End If
                        .Fields(.KeyFields(i + 1)).FindValue = FindValue(i)
                    Next i
                    For i = 1 To .Fields.Count
                        If .Fields(i).ForeignTableEntity = ForeignEntity Then
                            If Trim$(C) = vbNullString Then
                                DC.Find
                                C = .Fields(i).value
                                If DS.ValidateClientField(C.Tag, .Fields(i).value) Then
                                    FieldUpdateDescription C.Tag, Frm, msk
                                End If
                            End If
                        End If
                    Next i
                End With
            End If
        End If
    Next C
ErrHandler:
    Set DS = Nothing
End Sub

Public Sub FieldUpdateDescription(FieldKey As String, Frm As Form, msk As MaskEdBox)
    Dim L As Label
    Dim varFldValue As Variant
    
    With Frm.DataClass.DS
        Set L = .FindControl(FieldKey, "LABEL")
        If Not L Is Nothing Then
            L.MousePointer = vbDefault
            With .Fields(FieldKey)
                L = " " & .ForeignTableDescValue
                varFldValue = .ForeignTableFldValue
                If varFldValue <> vbNullString Then
                    msk = varFldValue
                    varFldValue = vbNullString
                End If
            End With
        End If
    End With
End Sub

Public Sub SetValueItems(Col As Column, Entity As Object, _
                         Optional Translate As Boolean = True)
    Dim vItem As ValueItem
    Dim DS As CDataStruct
    Dim i As Long
    Dim intRows As Integer
    
    With Col.ValueItems
        .Presentation = dbgComboBox
        .Translate = Translate
        .Validate = True
        Set DS = Entity.DS
        Entity.GetList True
        If DS.howmany > 0 Then
        
            ' el combo no puede tener mas de 100
            If DS.howmany > 100 Then
            .MaxComboItems = 100
            Else

            .MaxComboItems = DS.howmany
            End If
            intRows = UBound(DS.List, 2)
            For i = 0 To intRows 'UBound(DS.List, 2)
                 Set vItem = New ValueItem
                 vItem.value = DS.ListItem(0, i)
                 vItem.DisplayValue = DS.ListItem(1, i)
                .Add vItem
            Next i
        End If
    End With

End Sub

Public Sub ShowSearchForm(ParentForm As Form, FieldKey As String, _
                          DataClass As Object, FieldTabIndex As Integer, _
                          Optional ShowData As Boolean = False, _
                          Optional ShowPrint As Boolean = False, _
                          Optional varReportsList As Variant)
    Dim i As Integer
    Dim vItem As ValueItem
    Dim strFieldName As String
    Dim intFieldsCount As Integer
    
    With frmSearch
        Set .ParentForm = ParentForm
'        Set .DataAccess = DataClass.DS.DA
        Set .DataClass = DataClass
        .ActiveControlTabIndex = FieldTabIndex
        .FieldKey = FieldKey
        .ShowPrint = ShowPrint
        If ShowPrint Then
            .ReportsList = varReportsList
        End If
        With .TDBCond(2).Columns(0).ValueItems
            .Presentation = dbgComboBox
            .Translate = False
            .Validate = True
            intFieldsCount = DataClass.DS.Fields.Count
            .MaxComboItems = 10
            For i = 1 To intFieldsCount
                Set vItem = New ValueItem
                strFieldName = DataClass.DS.Fields(i).Name
                With vItem
                    .value = strFieldName
                    .DisplayValue = i
                End With
                .Add vItem
            Next i
        End With
        If ShowData Then
            .DoCommand fcFind
        End If
        .ZOrder 0
        .Show
        RESOIL2MAIN.Arrange vbCascade
    End With
    ParentForm.Enabled = False
End Sub

Public Sub SetRelationCond(ForeignEntity As String, DC As Object, Frm As Form)
    Dim C As Control
    Dim i As Integer
    Dim intFieldCount As Integer
    Dim DS As New CDataStruct
    
    Set DS = Frm.DataClass.DS
    For Each C In Frm.Controls
        If UCase(TypeName(C)) = "MASKEDBOX" Then
            If Len(C.Tag) > 0 Then
                If DS.Fields(C.Tag).ForeignTableEntity = ForeignEntity Then
                    With DC.DS.Fields
                        intFieldCount = .Count
                        For i = 1 To intFieldCount
                            If .Item(i).ForeignTableEntity = ForeignEntity Then
                                .Item(i).FindValue = C
                                Exit For
                            End If
                        Next i
                    End With
                    Exit For
                End If
            End If
        End If
    Next C
    Set DS = Nothing
End Sub

Public Sub ShowValidationError(DS As CDataStruct)
    Dim i As Integer
    Dim strMsg As String
    
    For i = 1 To DS.Fields.ValidationErrors.Count
        strMsg = strMsg & DS.Fields.ValidationErrors(i) & vbNewLine
    Next i
    MsgBox strMsg, vbCritical + vbOKOnly, "ERRORE"

End Sub


Public Sub ShowWindow(ByVal M As Menu)
    Dim strFunctionCode As FunctionCodes

    strFunctionCode = Mid$(M.Name, 5) & M.Index
    DisplayForm strFunctionCode
End Sub

Public Function IsEmptyVariantString(varValue As Variant) As Variant
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(varValue) = 0 Then
        IsEmptyVariantString = True
    Else
        IsEmptyVariantString = False
    End If
End Function

Public Function GetComputerName(rpsComputerName) As Boolean
  Dim sComputerName As String
  Dim lComputerNameLen As Long
  Dim lResult As Long
  Dim fRV As Boolean
  
  lComputerNameLen = 256
  sComputerName = Space(lComputerNameLen)
  
  lResult = w32_GetComputerName(sComputerName, lComputerNameLen)
  If lResult <> 0 Then
    rpsComputerName = Left$(sComputerName, lComputerNameLen)
    fRV = True
  Else
    fRV = False
  End If

  GetComputerName = fRV
End Function

'-- DP 08/08/08
Public Function GetFolder(ByVal hWndModal As Long) As String
     Dim bInf As BROWSEINFO
     Dim retval As Long
     Dim PathID As Long
     Dim RetPath As String
     Dim offset As Integer
     'Set the properties of the folder dialog
     bInf.hOwner = hWndModal
     bInf.lpszTitle = "Seleziona un Percorso"
     ' bInf.lpszTitle = MakeTranslatedString(2, "$G47")
     bInf.ulFlags = &H1
     bInf.pszDisplayName = Space$(260)
     'Show the Browse For Folder dialog
     PathID = SHBrowseForFolder(bInf)
     RetPath = Space$(512)
     retval = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
     If retval Then
          'Trim off the null chars ending the path
          'and display the returned folder
          offset = InStr(RetPath, Chr$(0))
          GetFolder = Left$(RetPath, offset - 1)
          'Free memory allocated for PIDL
          CoTaskMemFree PathID
     Else
          GetFolder = ""
     End If
End Function
