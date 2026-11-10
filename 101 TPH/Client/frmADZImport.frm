VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADZImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE ANAGRAFICHE"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   5760
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7680
         Picture         =   "frmADZImport.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmADZImport.frx":0102
         TabIndex        =   12
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6540
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11113
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17.28"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09/03/00"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame5 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   5175
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmADZImport.frx":0204
            TabIndex        =   5
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2775
         Left            =   5400
         TabIndex        =   11
         Top             =   600
         Width           =   3855
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ente Emittente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Destinazioni da importare:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Destinazioni importate correttamente:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   2655
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   9135
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmADZImport.frx":2C04
            TabIndex        =   23
            Top             =   240
            Width           =   8895
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3120
         Picture         =   "frmADZImport.frx":B51E
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmADZImport.frx":BA50
         TabIndex        =   2
         Tag             =   "IMPORT"
         Top             =   675
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   1680
         Picture         =   "frmADZImport.frx":BB52
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmADZImport.frx":C084
         TabIndex        =   1
         Tag             =   "OPEN"
         Top             =   675
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmADZImport.frx":C186
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmADZImport.frx":C288
         TabIndex        =   0
         Tag             =   "BROWSE"
         Top             =   675
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8640
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmADZImport.frx":C38A
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmADZImport.frx":C49C
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmADZImport.frx":C5AE
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File &Ordini"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmADZImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3
Private Const COL_COD_ORD_LOG = 1
Private Const COL_COD_PROD_LOG = 4
Private Const COL_RESULT_LOG = 10

Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
Private ADZ As CDestinazione
Private AEE As CEnteEmittente

Private mobjInFile As CTextFile
Private mobjInFileDettaglio As CTextFile
Private mstrIDEnte As String
Private mintFileInput As Integer

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
'                Select Case gmstrCodInstallazione
'                    Case ciInst1
                #If VersionID = 1 Then
                        .DialogTitle = "Seleziona il file degli ordini... "
'                    Case ciInst2
                #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
                        Select Case mintFileInput
                            Case 1
                                .DialogTitle = "Seleziona il file degli ordini... "
                            Case 2
                                .DialogTitle = "Seleziona il file di dettaglio... "
                        End Select
'                End Select
                #End If
                .ShowOpen
                mskFileName = .fileName
            End With
        Case fcImport
            ImportRecords
            cmdImport.Enabled = False
        Case fcOpen
            OpenTextFile
        Case fcPrint
            If State = fsEdit Then
            End If
        Case fcExit
            Unload Me
    End Select
End Sub






Private Sub ImportRecords()
    On Error GoTo ErrHandler
    
    Dim strIDClienteDefault As String
'    Dim strIDImballoDefault As String
'    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
'    Dim strIDTipoVendita As String
'    Dim strIDEnteEmittente As String
'    Dim strIDOrder As String
'    Dim strIDOrderAGIP As String
'    Dim strIDPrevOrderAGIP As String
'    Dim strIDProdotto As String
'    Dim strIDRiga  As String
'    Dim strCodDest As String
'    Dim strCodDestPrec As String
'    Dim strIDDeposito As String
'    Dim strIDZona As String
    Dim strErrorMessage As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
'    Dim blnNewOrder As Boolean
    Dim blnConcurError As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    
    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        strIDClienteDefault = Trim$(.Fields("3").value)
    End With
    
    If (strIDClienteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        
        ' --- Leggo una riga del file delle destinazioni
        For Col = 0 To 17
            mobjInFile(Col + 1).value = XA(Row, Col)
        Next Col
        ' ---
        
        ' --- Cerco la destinazione
        strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC, mobjInFile(1).value, True)
        
        ' --- La destinazione è già stata inserita ?
        If strIDDestinazione = vbNullString Then
            ' --- No : devo inserirla ora
            ADZ.DS.DA.BeginTransaction
'            With ADZ.DS.Fields
                RecordToADZ ADZ, mobjInFile, strIDClienteDefault
                If Not ADZ.Insert Then
                    intRowError = DestInsertError
                    GoTo LogResult
                End If
                If Not ADZ.DS.MultiSoc.Insert Then
                    intRowError = DestSocInsertError
                    GoTo LogResult
                End If
'            End With
        Else
            ' --- Sì : la aggiorno nel caso i dati siano cambiati
            RecordToADZ ADZ, mobjInFile, strIDClienteDefault, False
            ADZ.DS.Fields("0").value = strIDDestinazione
            ADZ.DS.Fields("0").FindValue = strIDDestinazione
            ADZ.BeginSave blnConcurError
            ADZ.DS.DA.EndTrans ETCOMMIT
            ADZ.DS.EndSave
'            intRowError = DestinationAlreadyExisting
        End If
        ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
            ADZ.DS.DA.EndTrans ETCOMMIT
        Else
            lngCountFail = lngCountFail + 1
            ADZ.DS.DA.EndTrans ETROLLBACK
        End If
    
        lngCountToGo = lngCountToGo - 1
    
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case DestInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
'            Case DestinationAlreadyExisting
'                strErrorMessage = "Destinazione già presente in anagrafica"
        End Select
        XALog.value(Row, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = Row
        End With
        TDBGrid1.Bookmark = Row
        ' ---
        
        DoEvents
    Next Row
    ' ---
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    ADZ.DS.DA.EndTrans ETROLLBACK
    Select Case ErrCode
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
'            If (strIDImballoDefault = vbNullString) Then
'                strInfo = strInfo & "Imballo" & vbNewLine
'            End If
'            If (strIDPosDefault = vbNullString) Then
'                strInfo = strInfo & "Posizione Doganale" & vbNewLine
'            End If
            If (strIDClienteDefault = vbNullString) Then
                strInfo = strInfo & "Codice Cliente" & vbNewLine
            End If
        Case NoImportData
            strMessage = "Non ci sono dati da importare."
            strInfo = vbNullString
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTImport", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Sub OpenTextFile()
    On Error GoTo ErrHandler
    
    Dim strFileName As String
    Dim strLine As String
    Dim strIDEnteDefault As String
    Dim strEnte As String
    Dim varBuffer As Variant
    Dim blnValidHeader As Boolean
    Dim blnIDOk As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim lngRows As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        strIDEnteDefault = Trim$(.Fields("16").value)
        strEnte = AEE.DecodeFirmCode(ID_SOC, strIDEnteDefault)
        If strEnte <> vbNullString Then
            lblField(1) = strEnte
        End If
    End With
    
    If (strIDEnteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    cmdImport.Enabled = False
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    Do While Not mobjInFile.EOF
        mobjInFile.ParseFixedLenFields mobjInFile.GetLine(False), 1, 19
        If strIDEnteDefault = mobjInFile("15").value Then
            With XA
                .ReDim 1, lngRows, 0, 17
                For i = 0 To 17
                    .value(.UpperBound(1), i) = mobjInFile(i + 2).value
                Next i
            End With
            With XALog
                .ReDim 1, lngRows, 0, 1
                .value(.UpperBound(1), 0) = mobjInFile(2).value
            End With
            lngRows = lngRows + 1
        End If
        DoEvents
    Loop
        
    cmdImport.Enabled = True
    With TDBGrid1
        .Array = XA
        .ReBind
        .Bookmark = 1
    End With
    With tdbLog
        .Array = XALog
        .ReBind
        .Bookmark = 1
    End With
    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows - 1
    
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmADZImport", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case MissingDefaults
            Set C = ErrH.SetError(MissingDefaults, "frmADZImport", _
                    "OpenTextFile", Err.Number, "Parametri di sistema mancanti", 10, vbNullString)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmADZImport", _
                            "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub



Private Sub RecordToADZ(ADZ As CDestinazione, _
                        InFile As CTextFile, _
                        IDClienteDefault As String, _
                        Optional InsertFlag As Boolean = True)
    With ADZ.DS.Fields
        .ResetFields
        If InsertFlag Then
            .Item("0").NewValue = Trim$(Left$(InFile("1").value, .Item("0").MaxLength))
            .Item("0").Updated = True
        End If
        .Item("1").NewValue = Trim$(Left$(InFile("3").value, .Item("1").MaxLength))
        .Item("1").Updated = True
        .Item("4").NewValue = Trim$(Left$(InFile("4").value, .Item("4").MaxLength))
        .Item("4").Updated = True
        .Item("5").NewValue = Trim$(Left$(InFile("5").value, .Item("5").MaxLength))
        .Item("5").Updated = True
        .Item("6").NewValue = Trim$(Left$(InFile("6").value, .Item("6").MaxLength))
        .Item("6").Updated = True
        .Item("7").NewValue = Trim$(Left$(InFile("7").value, .Item("7").MaxLength))
        .Item("7").Updated = True
        .Item("8").NewValue = IDClienteDefault
        .Item("8").Updated = True
'        .Item("16").NewValue = Trim$(Left$(InFile("2").Value, .Item("16").MaxLength))
'        .Item("16").Updated = True
        .Item("18").NewValue = Trim$(Left$(InFile("16").value, .Item("18").MaxLength))
        .Item("18").Updated = True
    End With
    If InsertFlag Then
        With ADZ.DS.MultiSoc.DS.Fields
            .ResetFields
            .Item("Codice").NewValue = Trim$(Left$(InFile("1").value, .Item("Codice").MaxLength))
            .Item("Codice").Updated = True
            .Item("Società").NewValue = ID_SOC
            .Item("Società").Updated = True
            .Item("CodPerSoc").NewValue = Trim$(Left$(InFile("1").value, .Item("CodPerSoc").MaxLength))
            .Item("CodPerSoc").Updated = True
        End With
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    
    ConfigureObjects Me
    mskFileName = gmRegistry.GetSetting("Options", "DestImportFile", , True)
    
    Set ErrH = New CErrorHandler
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AEE = New CEnteEmittente
    With AEE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
    With mobjInFile
        .Add "Tipo record", 4, "1"
        .Add "Codice", 5, "2"
        .Add "Filler1", 3, "3"
        .Add "Nome Gestore", 30, "4"
        .Add "Indirizzo", 30, "5"
        .Add "Località", 30, "6"
        .Add "Provincia", 2, "7"
        .Add "CAP", 5, "8"
        .Add "Codice Cliente", 7, "9", , fftNumerico
        .Add "Codice Provincia", 3, "10", , fftNumerico
        .Add "Postazione", 3, "11", , fftNumerico
        .Add "CC", 1, "12", , fftNumerico
        .Add "Filler2", 1, "13"
        .Add "Ente Emittente", 3, "14", , fftNumerico
        .Add "Presidio Logistico", 3, "15", , fftNumerico
        .Add "Blocco", 2, "16"
        .Add "Telefono", 11, "17"
        .Add "FAX", 11, "18"
        .Add "Ass. Rete", 11, "19"
    End With
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "DestImportFile", mskFileName, True
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set AEE = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set mobjInFileDettaglio = Nothing
    Set frmADZImport = Nothing
End Sub







Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub


Private Sub mskFileName_GotFocus()
    mintFileInput = 1
End Sub





