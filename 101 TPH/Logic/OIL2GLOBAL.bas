Attribute VB_Name = "OIL2GLOBAL"
Option Explicit

Public DA As CDataAccess
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
    fcVIAINSE = "231"
    fcVIASELP = "232"
    fcVIASELD = "3131"
    fcVIASELT = "3132"
    fcVIACOMP = "234"
    fcVIASELE = "235"
    fcGP1SELE = "241"
    fcPPTCREA = "242"
    fcPPTSELE = "243"
    fcPREFPPT = "998000"
    fcPREFPPD = "999888"
    fcPREFGP1 = "999777"
    fcPRNFORM = "999000"
    fcORDIOCT = "999222"
    fcORDIOCR = "999333"
    fcVIAGVVT = "999444"
    fcVIAGVVC = "999555"
    fcVIAGVVP = "999666"
    fcANAGAPR = "3171"
    fcANAGAIM = "3172"
    fcANAGAPD = "3173"
    fcANAGAVE = "3181"
    fcANAGAAU = "3182"
    fcANAGAAM = "3183"
    fcANAGAUT = "3184"
    fcANAGACN = "3185"
    fcTARITFA = "3191"
    fcSTATFRM = "251"
    fcAGIPOIM = "261"
    fcMOVEREX = "262"
    fcMOVERIM = "263"
    fcAGIPVEX = "264"
    fcANAGRIM = "265"
    fcAGIPPIM = "267"
    fcMOVRAGO = "268"
    fcMOVRAGI = "269"
    fcORESFRO = "3622"
    fcANAGAMP = "31101"
    fcANAGAIV = "31102"
    fcANAGATC = "41951"
    fcANAGATD = "41952"
    fcANAGATM = "41953"
    fcANAGATS = "41954"
    fcANAGATV = "41955"
    fcANAGATN = "41956"
    fcANAGAZO = "31103"
    fcANAGAVA = "31104"
    fcANAGAXX = "41971"
    fcANAGANU = "41972"
    fcANAGAXT = "41973"
    fcANAGARC = "41961"
    fcANAGARD = "41962"
    fcANAGARP = "41963"
    fcANAGARR = "41964"
    fcANAGARO = "41965"
    fcANAGAFP = "41966"
End Enum

Public Enum CodiceInstallazione
    ciInst1 = "001"
    ciInst2 = "002"
    ciInst3 = "003"
End Enum

Public Const MAX_RECORDS = 999
Public Const MAX_DBSIZE_DEMOMODE_JET = 1500000
Public Const MAX_DBSIZE_DEMOMODE_SQL = 3000000
Public Const ID_SOC = "001"
Public Const ID_PARAM = "001"
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
    NoUT = vbObjectError + 530
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
    UnknownError = vbObjectError + 999
End Enum


Public Function CheckDBSize(Engine As DBEngine) As Boolean
    On Error GoTo ErrHandler
    Dim lngDBSize As Long
    
    Select Case Engine
        Case deJet
            lngDBSize = FileLen(DA.ConnDB)
            If lngDBSize > MAX_DBSIZE_DEMOMODE_JET Then
                GoTo ErrHandler
            End If
        Case deSQLServer
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
    
    Select Case gmstrCodInstallazione
        Case ciInst2
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
    End Select
    CheckLicense = True
    Exit Function
ErrHandler:
    MsgBox "La password per l'anno " & intYear & " non è corretta", vbCritical + vbOKOnly, "Controllo della licenza"
    CheckLicense = False
End Function

Public Sub CleanUp()
    SetOptions
    Set ClientControls = Nothing
    Set DA = Nothing
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
        Temp = Asc(Mid(DecryptedPassword, Counter, 1))
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
        Temp = Asc(Mid(EncryptedPassword, Counter, 1))
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

Private Sub SetMenuConfiguration()
    With RESOIL2MAIN
        Select Case gmstrCodInstallazione
            Case ciInst1
'                .mnuL26(5).Visible = False
            Case ciInst2
'                .mnuL21(9).Visible = False
'                .mnuL3110(8).Visible = False
'                .mnuL41910(3).Visible = False
'                .mnuL1(4).Visible = False
'                .mnuL1(5).Visible = False
                .mnuL23(4).Visible = False
                .mnuL26(2).Visible = False
                .mnuL26(3).Visible = False
                .mnuL26(6).Visible = False
                .mnuL26(7).Visible = False
                .mnuL26(8).Visible = False
                .mnuL26(9).Visible = False
                .mnuL322(3).Visible = False
        End Select
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
        .SaveSetting "Database", "DBName", DA.ConnDB, True
        .SaveSetting "Database", "Server", DA.ConnServer, True
        .SaveSetting "Database", "UID", DA.ConnUserName, True
        .SaveSetting "Database", "Engine", CStr(DA.DataEngine), True
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
        Case "ACL"
            Set DC = New CCliente
        Case "ACN"
            Set DC = New CContratto
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
            SetRelationCond "ASC", DC, Frm
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
        Case Else
            Set DC = Nothing
    End Select
    If Not DC Is Nothing Then
        With DC.DS
            Set .FormHandle = Frm
            Set .DA = DA
        End With
        ShowSearchForm Frm, FieldKey, DC, FieldTabIndex, blnShowData, blnShowPrint, varReportsList
        Set DC = Nothing
    End If
End Sub


Sub Main()
    On Error GoTo ErrHandler
        
    If App.PrevInstance Then
        End
    End If
    
    Dim strLicenseTo As String
    Dim strLicenseName As String
    
    Dim AXX As CParametro
    
    Set gmRegistry = New CRegSettings
    
    If Not GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYCOMMONFILES, gREGVALCOMMONFILES, CommonFilesDir) Then
    End If
    
    GetLicenseInfo strLicenseTo, strLicenseName
    gmblnDemoMode = Len(gmRegistry.GetSetting("General", "RES1999", "", True)) > 0
    
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
    Set DA = New CDataAccess
    
    GetOptions
    
    With frmLogon
        .Caption = App.ProductName & " - LOGON"
        .cmbProvider.ListIndex = CInt(gmRegistry.GetSetting("Database", "Engine", "0", True))
        .txtDatabase = gmRegistry.GetSetting("Database", "DBName", , True)
        .txtServer = gmRegistry.GetSetting("Database", "Server", , True)
        .txtUID = gmRegistry.GetSetting("Database", "UID", , True)
        .txtPWD = ""
    End With
    
    Do While frmSplash.Visible
        DoEvents
    Loop
    
    frmLogon.Show
    Do While (Not gmblnConnectOk) And (Not gmblnCancelConnect)
        DoEvents
    Loop
    
    If Not gmblnCancelConnect Then
        If gmblnDemoMode Then
            If Not CheckDBSize(DA.DataEngine) Then
                MsgBox "Il database ha raggiunto la dimensione massima consentita per questa versione dell'applicazione", _
                       vbCritical + vbOKOnly
                GoTo ErrHandler
            End If
        End If
        
        Set AXX = New CParametro
        
        With AXX.DS
            Set .DA = DA
            .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
            If AXX.Find Then
                gmstrCodInstallazione = .Fields("5").value
            End If
        End With
        Set AXX = Nothing
        
        With RESOIL2MAIN
            If Not CheckLicense Then
                GoTo ErrHandler
            End If
            SetBackPic
            SetMenuConfiguration
            .AutoShowChildren = False
            .StatusBar1.Panels(2).Text = "Database: " & DA.ConnDB
            .Show
        End With
        
    End If
    Unload frmSplash
    Exit Sub
ErrHandler:
    CleanUp
    Set AXX = Nothing
    Set frmSplash = Nothing
    Set RESOIL2MAIN = Nothing
    End
End Sub

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
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
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
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
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
            F.DoCommand fcCustom2
    End Select
End Sub

Public Function AddClientControl(ByVal Obj As Object, ByRef F As Form, Optional ByVal Key As String = "") As CControl
    Set AddClientControl = ClientControls.Add(Obj, F)
End Function

Public Sub ConfigureObjects(ByRef F As Form, _
                            Optional ColTags As Variant, _
                            Optional SetMultiSocTags As Boolean = False, _
                            Optional ColTags1 As Variant)
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
        Case fcVIAINSE
            Set F = frmVVTInsert
        Case fcVIASELP
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
        Case fcVIASELT
            Set F = frmVVTSelect
            F.FunctionCode = strFunctionCode
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
        Case fcVIAGVVP
            Set F = OIL2VVP0
        Case fcVIAGVVC
            Set F = OIL2VVC0
        Case fcGP1SELE
            Set F = frmGP1Select
        Case fcPPTCREA
            Set F = frmPPTCreate
        Case fcPPTSELE
            Set F = frmPPTSelect
        Case fcPREFPPT
            Set F = OIL2PPT0
        Case fcPREFPPD
            Set F = OIL2PPD0
        Case fcTARITFA
            Set F = OIL2TFA0
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
            Set F = OIL2AAM0
        Case fcANAGAUT
            Set F = OIL2AUT0
        Case fcANAGACN
            Set F = OIL2ACN0
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
        Case fcAGIPOIM
            Set F = frmOCTImport
        Case fcAGIPVEX
            Set F = frmVVTExport
        Case fcANAGRIM
            Set F = frmADZImport
        Case fcAGIPPIM
            Set F = frmGP1Import
        Case fcMOVEREX
            Set F = frmOCTExportMover
        Case fcMOVERIM
            Set F = frmVVTImportMover
        Case fcMOVRAGO
            Set F = frmVVCExportMap
        Case fcMOVRAGI
            Set F = frmVVCImportMap
        Case fcPRNFORM
            Set F = frmPrintReport
        Case fcPREFGP1
            Set F = OIL2GP10
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
                        varRelatedEntity = Array("ACL", "ADE", "AZO", "AMP", "AIV")
                        Set DC = New CDestinazione
                    Case "AEE"
                        varRelatedEntity = Array("ASC")
                        Set DC = New CEnteEmittente
                    Case "ACL"
                        varRelatedEntity = Array("ASC", "AEE")
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
                        Set .DA = DA
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
                    DC.Find
                    For i = 1 To .Fields.Count
                        If .Fields(i).ForeignTableEntity = ForeignEntity Then
                            If Trim$(C) = vbNullString Then
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
        If DS.HowMany > 0 Then
            .MaxComboItems = DS.HowMany
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
        Set .DataAccess = DataClass.DS.DA
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
