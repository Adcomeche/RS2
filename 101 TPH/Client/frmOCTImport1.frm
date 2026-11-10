VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOCTImport1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE ORDINI"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9030
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   4800
      Width           =   9015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3120
         Picture         =   "frmOCTImport1.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmOCTImport1.frx":0532
         TabIndex        =   19
         Tag             =   "IMPORT"
         Top             =   240
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
         Picture         =   "frmOCTImport1.frx":0634
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmOCTImport1.frx":0B66
         TabIndex        =   21
         Tag             =   "OPEN"
         Top             =   240
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
         Picture         =   "frmOCTImport1.frx":0C68
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTImport1.frx":0D6A
         TabIndex        =   23
         Tag             =   "BROWSE"
         Top             =   240
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7320
         Picture         =   "frmOCTImport1.frx":0E6C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7200
         Picture         =   "frmOCTImport1.frx":0F6E
         TabIndex        =   5
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5625
      Width           =   9030
      _ExtentX        =   15928
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
            Object.Width           =   10478
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15.39"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "05/01/00"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   4560
         TabIndex        =   4
         Top             =   600
         Width           =   4335
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Tag             =   "2"
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   3840
            Tag             =   "2"
            Top             =   480
            Width           =   240
         End
         Begin VB.Label lblField4 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   720
            TabIndex        =   28
            Tag             =   "2"
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "&Ente Emit."
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini da importare:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   12
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini importati correttamente:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   4335
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "frmOCTImport1.frx":1070
            TabIndex        =   17
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1935
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   8775
         Begin TrueDBGrid60.TDBGrid TDBGrid2 
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmOCTImport1.frx":3A74
            TabIndex        =   16
            Top             =   240
            Width           =   8535
         End
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7920
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox mskDataConsegna 
         Height          =   255
         Left            =   6840
         TabIndex        =   24
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
         Top             =   3840
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
               Picture         =   "frmOCTImport1.frx":8DE6
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport1.frx":8EF8
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport1.frx":900A
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Cons. &Rich."
         Height          =   255
         Index           =   11
         Left            =   5520
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File &Ordini"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOCTImport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3
Private Const COL_COD_ORD_LOG = 1
Private Const COL_COD_PROD_LOG = 4
Private Const COL_RESULT_LOG = 1

Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineRiga
Private APR As CProdotto
Private AEE As CEnteEmittente

Private mobjInFile As CTextFile
Private mobjInFileDettaglio As CTextFile
Private mstrIDEnte As String

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public DataClass As COrdineSelezione

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1


Private Function CheckDestinazione(strCodDest As String, _
                                   strIDEnteEmittenteDefault As String, _
                                   strIDSocDefault As String) As Long
    Dim intRowError As Errors
    Dim strIDEnteEmittente As String
    Dim strIDDestinazione As String
    
    intRowError = NoError
    ' --- Cerco la destinazione
    strIDDestinazione = ADZ.DecodeFirmCode(strIDSocDefault, strCodDest, True)
    
    ' --- La destinazione è già stata inserita ?
    If strIDDestinazione = vbNullString Then
        ' --- No : segnalo l'errore
        intRowError = NoDestinazione
        GoTo ExitFunction
    Else
        ' --- Controllo che l'ente emittente sia uguale a quello di default
        ADZ.DS.Fields("0").FindValue = strIDDestinazione
        If Not ADZ.Find Then
            intRowError = NoDestinazione
            GoTo ExitFunction
        End If
        strIDEnteEmittente = ADZ.DS.Fields("27").value
        If strIDEnteEmittente <> strIDEnteEmittenteDefault Then
            intRowError = NoEnteEmittente
            GoTo ExitFunction
        End If
    End If
    ' ---
ExitFunction:
    CheckDestinazione = intRowError
End Function



    
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
'                .InitDir = App.Path
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .DialogTitle = "Seleziona il file degli ordini... "
                .ShowOpen
                mskFileName = .fileName
            End With
        Case fcImport
            ImportRecords
            mskDataConsegna.Text = DATE_MASK
            cmdImport.Enabled = False
        Case fcOpen
            OpenTextFile
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub ImportRecords()
    On Error GoTo ErrHandler

    Dim strIDImballo As String
    Dim strIDPosDog As String
    Dim strIDSocDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDOrder As String
    Dim strIDOrderAGIP As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDCliente As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strNewLocalita As String
    Dim strErrorMessage As String
    Dim varProdotti(1 To 10) As String
    Dim varDummy  As Variant
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim lngQtaOrdinata As Long
    Dim Col As Integer
    Dim i As Integer
    Dim blnNewOrder As Boolean
    Dim blnLocHasChanged As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    ReDim ListNewLoc(0 To 0) As String

    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError

    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = strIDParam
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDSocDefault = Trim$(.Fields("4").value)
        For i = 1 To 10
            If Trim$(.Fields(i + 6).value) <> vbNullString Then
                varProdotti(i) = Trim$(.Fields(i + 6).value)
            End If
        Next i
    End With

    If (strIDSocDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---

    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)

    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError

        ' --- Leggo una riga del file degli ordini
        For Col = 1 To 8
            mobjInFile(Col).value = XA(Row, Col)
        Next Col
        ' ---
        strIDDestinazione = ADZ.DecodeFirmCode(strIDSocDefault, Right$(mobjInFile("3").value, 4), True)
        With ADZ.DS
            .Fields.ResetFields
            .Fields("0").FindValue = strIDDestinazione
            If Not ADZ.Find Then
                intRowError = NoDestinazione
                GoTo ErrHandler
            End If
            If .HowMany = 0 Then
                intRowError = NoDestinazione
                GoTo ErrHandler
            End If
            strIDDeposito = .Fields("15").value
            strIDTipoVendita = .Fields("28").value
            strIDEnteEmittente = .Fields("27").value
            strIDCliente = .Fields("8").value
            strIDZona = .Fields("16").value
        End With
        
        ' --- Inserisco la testata dell'ordine
        OCT.DS.Fields.ResetFields
        RecordToOCT OCT, mobjInFile, strIDDeposito, strIDDestinazione, _
                    strIDTipoVendita, strIDEnteEmittente, strIDCliente, strIDZona, _
                    strIDSocDefault
        
        OCT.DS.DA.BeginTransaction
        
        If Not OCT.Insert Then
            intRowError = OrderInsertError
            GoTo LogResult
        End If
        strIDOrder = OCT.DS.Fields("0").NewValue
        strIDRiga = 10

        For i = 1 To 4
            ' --- Inserisco una nuova riga ordine
            APR.DS.Fields.ResetFields
            lngQtaOrdinata = mobjInFile(i + 3).value
            If lngQtaOrdinata > 0 Then
                strIDProdotto = varProdotti(i)
                
                With APR.DS.MultiSoc
                    .DS.Fields("Codice").FindValue = strIDProdotto
                    .DS.Fields("Società").FindValue = strIDSocDefault
                    If Not .GetList Then
                        intRowError = NoProdotto
                        GoTo LogResult
                    End If
                    varDummy = .DS.List
                    strIDImballo = varDummy(1, 0)
                    strIDPosDog = varDummy(2, 0)
                End With
                
                RecordToOCR OCR, strIDOrder, strIDRiga, _
                            strIDProdotto, strIDImballo, strIDPosDog, _
                            lngQtaOrdinata * 1000
                
                strIDRiga = strIDRiga + 10
                
                If Not OCR.Insert Then
                    intRowError = RowOrderInsertError
                    GoTo LogResult
                End If
            End If
            ' ---
        Next i

LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError <> NoError) Then
            lngCountFail = lngCountFail + 1
            OCT.DS.DA.EndTrans ETROLLBACK
        Else
            lngCountOK = lngCountOK + 1
            OCT.DS.DA.EndTrans ETCOMMIT
        End If
        lngCountToGo = lngCountToGo - 1

        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo

        Select Case intRowError
            Case NoError
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case OrderInsertError
                strErrorMessage = "Errore durante l'inserimento dell'ordine "
            Case RowOrderInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di ordine"
        End Select
        XALog.value(Row, COL_RESULT_LOG) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = Row
        End With
        TDBGrid2.Bookmark = Row
        ' ---

        DoEvents
    Next Row
    ' ---

    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long

    'OCT.DS.DA.EndTrans 1
    Select Case ErrCode
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
            If (strIDSocDefault = vbNullString) Then
                strInfo = strInfo & "Imballo" & vbNewLine
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
    Dim strEnte As String
    Dim strIDOrdine As String
    Dim strIDEnte As String
    Dim strIDSocDefault As String
    Dim varBuffer As Variant
    Dim varProdotti(1 To 10) As String
    Dim blnValidHeader As Boolean
    Dim blnIDOk As Boolean
    Dim blnFound As Boolean
    Dim blnFlagProdottiOK As Boolean
    Dim blnFirstRow As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim j As Long
    Dim lngRows As Long
    Dim lngCountOrders As Long
    Dim lngUbound As Long
    Dim lngPosition As Long
    Dim FileLocation As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    cmdImport.Enabled = False
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        strIDEnte = Trim$(.Fields("16").value)
        strIDSocDefault = .Fields("4").value
'        strEnte = AEE.DecodeFirmCode(strIDSocDefault, strIDEnte)
'        If strEnte <> vbNullString Then
'            lblField(1) = strEnte
'        End If
        For i = 1 To 10
            If Trim$(.Fields(i + 6).value) <> vbNullString Then
                blnFlagProdottiOK = True
                varProdotti(i) = Trim$(.Fields(i + 6).value)
            End If
        Next i
    End With
    
    If Not blnFlagProdottiOK Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    strIDEnte = MaskEdBox1
    If (strIDEnte = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    
    For i = 1 To 4
        Me.TDBGrid2.Columns(i + 3).Caption = varProdotti(i)
    Next i
    ' ---

    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    lngCountOrders = 1
    blnFirstRow = True
    Do While Not mobjInFile.EOF(True, lngPosition)
        If blnFirstRow Then
            blnFirstRow = False
            mobjInFile.ParseFixedLenFields Right(mobjInFile.GetLine(False, 31, lngPosition), 29), 1, 8
        Else
        mobjInFile.ParseFixedLenFields mobjInFile.GetLine(False, 29, lngPosition), 1, 8
        End If
        If CheckDestinazione(Right(mobjInFile("3").value, 4), strIDEnte, strIDSocDefault) = NoError Then
            With XA
                .ReDim 1, lngCountOrders, 0, 8
                .value(.UpperBound(1), 0) = lngCountOrders
                For i = 0 To 7
                    .value(.UpperBound(1), i + 1) = mobjInFile(i + 1).value 'varBuffer(i + 1)
                Next i
            End With
            With XALog
                .ReDim 1, lngCountOrders, 0, 1
                .value(.UpperBound(1), 0) = lngCountOrders
            End With
            lngCountOrders = lngCountOrders + 1
        End If
        DoEvents
    Loop
    
    If lngCountOrders = 1 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    
    cmdImport.Enabled = True
    With TDBGrid2
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
    lblLog(5) = lngCountOrders - 1
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case MissingDefaults
            Set C = ErrH.SetError(MissingDefaults, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Parametri di sistema mancanti", 10, "Nome del file: " & strFileName)
        Case NoImportData
            Set C = ErrH.SetError(NoImportData, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Nessun ordine da importare.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmOCTImport", _
                            "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub


Private Sub RecordToADZ(ADZ As CDestinazione, InFile As CTextFile, IDClienteDefault As String, IDDeposito As String)
    Dim strIDDestinazione As String
    
    With ADZ.DS.Fields
        .ResetFields
        strIDDestinazione = Format$(InFile("4").value, "00000000")
        .Item("0").NewValue = strIDDestinazione 'Trim$(Left$(InFile("4").value, .Item("0").MaxLength))
        .Item("0").Updated = True
        .Item("1").NewValue = Trim$(Left$(InFile("5").value, .Item("1").MaxLength))
        .Item("1").Updated = True
        .Item("3").NewValue = Trim$(Left$(InFile("5").value, .Item("3").MaxLength))
        .Item("3").Updated = True
        .Item("4").NewValue = Trim$(Left$(InFile("6").value, .Item("4").MaxLength))
        .Item("4").Updated = True
        .Item("5").NewValue = Trim$(Left$(InFile("7").value, .Item("5").MaxLength))
        .Item("5").Updated = True
        .Item("7").NewValue = Trim$(Left$(InFile("8").value, .Item("7").MaxLength))
        .Item("7").Updated = True
        .Item("8").NewValue = IDClienteDefault
        .Item("8").Updated = True
        .Item("15").NewValue = IDDeposito
        .Item("15").Updated = True
        .Item("16").NewValue = Trim$(Left$(InFile("2").value, .Item("16").MaxLength))
        .Item("16").Updated = True
        .Item("18").NewValue = Trim$(Left$(InFile("9").value, .Item("18").MaxLength))
        .Item("18").Updated = True
    End With
    With ADZ.DS.MultiSoc.DS.Fields
        .ResetFields
        .Item("Codice").NewValue = strIDDestinazione 'Trim$(Left$(InFile("4").value, .Item("Codice").MaxLength))
        .Item("Codice").Updated = True
        .Item("Società").NewValue = ID_SOC
        .Item("Società").Updated = True
        .Item("CodPerSoc").NewValue = Trim$(Left$(InFile("4").value, .Item("CodPerSoc").MaxLength))
        .Item("CodPerSoc").Updated = True
    End With
End Sub

Private Sub RecordToOCR(OCR As COrdineRiga, IDordine As String, _
                        IDRiga As String, IDProdotto As String, _
                        IDImballo As String, IDPosDoganale As String, _
                        Qta As Long)
    With OCR.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDordine
        .Item("0").Updated = True
        .Item("1").NewValue = IDRiga
        .Item("1").Updated = True
        .Item("2").NewValue = IDProdotto
        .Item("2").Updated = True
        .Item("3").NewValue = IDImballo
        .Item("3").Updated = True
        .Item("4").NewValue = IDPosDoganale
        .Item("4").Updated = True
        .Item("7").NewValue = Qta
        .Item("7").Updated = True
        .Item("11").NewValue = Qta
        .Item("11").Updated = True
        .Item("21").NewValue = ""
        .Item("21").Updated = True
    End With
End Sub
Private Sub RecordToOCT(OCT As COrdine, InFile As CTextFile, _
                        IDDeposito As String, IDDestinazione As String, _
                        IDTipoVendita As String, IDEnteEmittente As String, _
                        IDCliente As String, IDZona As String, IDSoc As String)
    
    Dim strDtModifica As String
    Dim strHrModifica As String
    Dim strDtConsegna As String
    
    With OCT.DS.Fields
        .ResetFields
        .Item("1").NewValue = IDSoc
        .Item("1").Updated = True
        .Item("2").NewValue = IDEnteEmittente
        .Item("2").Updated = True
        .Item("3").NewValue = Format$(Date, "dd/mm/yyyy")
        .Item("3").Updated = True
        .Item("4").NewValue = Format$(Time, "hh:mm")
        .Item("4").Updated = True
        ' --- Data inserimento ordine
        If Len(Trim$(InFile("1").value)) > 0 Then
            strDtModifica = InFile("1").value
            strDtModifica = Right$(strDtModifica, 2) & _
                            "/" & Mid$(strDtModifica, 5, 2) & _
                            "/" & Left$(strDtModifica, 4)

            strHrModifica = InFile("2").value
            strHrModifica = Left$(strHrModifica, 2) & ":" & Mid$(strHrModifica, 3, 2)
            .Item("5").NewValue = strDtModifica
            .Item("5").Updated = True
            .Item("6").NewValue = strHrModifica
            .Item("6").Updated = True
        End If
        ' ---

        ' --- Data consegna richiesta
        strDtConsegna = Format$(mskDataConsegna.Text, "dd/mm/yyyy")
        .Item("7").NewValue = strDtConsegna
        .Item("7").Updated = True
        ' ---
        ' --- Data consegna pianificata
        .Item("8").NewValue = strDtConsegna
        .Item("8").Updated = True
        ' ---
'        .Item("10").NewValue = Trim$(Left$(InFile("12").value, .Item("10").MaxLength))
'        .Item("10").Updated = True
        .Item("12").NewValue = IDCliente
        .Item("12").Updated = True
        .Item("13").NewValue = IDDestinazione
        .Item("13").Updated = True
        ' --- Nota 1
        Select Case InFile("8").value
            Case "1"
                .Item("14").NewValue = "Mattina"
            Case "2"
                .Item("14").NewValue = "Pomeriggio"
            Case Else
                .Item("14").NewValue = ""
        End Select
        .Item("14").Updated = True
        ' ---
        .Item("16").NewValue = IDTipoVendita
        .Item("16").Updated = True
        .Item("17").NewValue = IDDeposito
        .Item("17").Updated = True
        .Item("18").NewValue = IDZona
        .Item("18").Updated = True
        .Item("21").NewValue = "R"
        .Item("21").Updated = True
        .Item("26").NewValue = "CC"
        .Item("26").Updated = True
    End With

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    
    Set DataClass = New COrdineSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .ConfigureInterface
        ConfigureObjects Me
    End With
    mskFileName = gmRegistry.GetSetting("Options", "OrderImportFile", , True)
    
    Set ErrH = New CErrorHandler
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = ID_PARAM
        If AXX.Find Then
            MaskEdBox1 = AXX.DS.Fields("16").value
            If DataClass.DS.ValidateClientField("2", MaskEdBox1) Then
                FieldUpdateDescription "2", Me, MaskEdBox1
            End If
        End If
    End With
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineRiga
    With OCR.DS
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
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
    With mobjInFile
        .Add "Data inserimento", 8, "1"
        .Add "Ora inserimento", 6, "2"
        .Add "Codice Destinazione", 6, "3"
        .Add "Qta Prodotto 1", 2, "4"
        .Add "Qta Prodotto 2", 2, "5"
        .Add "Qta Prodotto 3", 2, "6"
        .Add "Qta Prodotto 4", 2, "7"
        .Add "Preferenze Gestore", 1, "8"
    End With
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "OrderImportFile", mskFileName, True
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set OCR = Nothing
    Set OCT = Nothing
    Set APR = Nothing
    Set AEE = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set DataClass = Nothing
    Set frmOCTImport1 = Nothing
End Sub







Private Sub mskDataConsegna_Change()
    cmdReadFile.Enabled = ((Len(mskFileName) > 0) And IsDate(mskDataConsegna.Text))
    cmdImport.Enabled = False
End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = ((Len(mskFileName) > 0) And IsDate(mskDataConsegna.Text))
    cmdImport.Enabled = False
End Sub


