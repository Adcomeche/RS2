VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrintReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stampa"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSendTo 
      Height          =   1695
      Left            =   2280
      TabIndex        =   30
      Top             =   3000
      Width           =   6735
      Begin VB.TextBox txtBody 
         Height          =   555
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox txtSubj 
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox txtSendTo 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "M&essaggio:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "O&ggetto:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Destinatario:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraFileName 
      Caption         =   "Es&porta sul file:"
      Height          =   855
      Left            =   5280
      TabIndex        =   15
      Top             =   2040
      Width           =   3735
      Begin VB.TextBox txtFileName 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton cmdSelectDB 
         Height          =   300
         Left            =   3195
         Picture         =   "frmPrintReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "&Invia E-mail"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "Esporta su fi&le"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "&Stampante"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.OptionButton optDestination 
      Caption         =   "&Anteprima"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame fraFormat 
      Caption         =   "&Formato:"
      Height          =   855
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   2895
      Begin VB.ComboBox cmbFormat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "&Orientamento:"
      Height          =   855
      Left            =   5280
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
      Begin VB.OptionButton optOrientation 
         Caption         =   "Ori&zzontale"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "&Verticale"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Tipo di stampa:"
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbReport 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame fraCopies 
      Caption         =   "&Numero di copie:"
      Height          =   855
      Left            =   6720
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
      Begin VB.ComboBox cmbCopies 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "Sta&mpante:"
      Height          =   855
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
      Begin VB.ComboBox cmbPrinters 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   26
      Top             =   4680
      Width           =   9015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7320
         Picture         =   "frmPrintReport.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   240
         Picture         =   "frmPrintReport.frx":0204
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Esegui"
         Height          =   375
         Left            =   120
         Picture         =   "frmPrintReport.frx":0736
         TabIndex        =   24
         Tag             =   "SAVE"
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   5625
      Width           =   9060
      _ExtentX        =   15981
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
            Object.Width           =   10530
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.23"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15/05/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlOpenFile 
      Left            =   240
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleziona il database..."
   End
   Begin VB.Image imgDestination 
      Height          =   480
      Index           =   2
      Left            =   240
      Picture         =   "frmPrintReport.frx":0C68
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image imgDestination 
      Height          =   480
      Index           =   3
      Left            =   240
      Picture         =   "frmPrintReport.frx":10AA
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgDestination 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmPrintReport.frx":14EC
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image imgDestination 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmPrintReport.frx":17F6
      Top             =   1800
      Width           =   480
   End
End
Attribute VB_Name = "frmPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSelectionFormula As String
Private mstrComment As String
Private mvarSortFields As Variant
Private mvarReportsList As Variant
Private mvarFormulas As Variant
Private LogOnInfo As PELogOnInfo

Private mobjParentForm As Object
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private varFormats As Variant
Private varFileExt As Variant
Private varFormatDLL As Variant
Private varFormatType As Variant
Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Public Sub DoCommand(Cmd As FormCommand)
    On Error GoTo ErrHandler
    
    Dim intNumberOfCopy As Integer
    Dim intDestination As Integer
    Dim strReportFileName As String
    Dim lngWidth As Integer
    Dim lngHeight As Integer
    
    If Cmd <> fcExit Then
        strReportFileName = gmstrReportsPath & "\" & mvarReportsList(1, cmbReport.ListIndex)
    End If
    
    Select Case Cmd
        Case fcSave
            Me.Enabled = False
            intNumberOfCopy = cmbCopies.ListIndex + 1
            intDestination = -1
            Do
                intDestination = intDestination + 1
            Loop Until optDestination(intDestination).value
            If UBound(mvarReportsList, 1) > 2 Then
                lngWidth = mvarReportsList(3, cmbReport.ListIndex)
                lngHeight = mvarReportsList(4, cmbReport.ListIndex)
            End If
            DoPrint strReportFileName, LogOnInfo, intDestination, mstrSelectionFormula, _
                    mvarFormulas, intNumberOfCopy, mvarSortFields, , cmbFormat.ListIndex, _
                    txtFileName, txtSendTo, txtSubj, txtBody, mstrComment, lngWidth, lngHeight
            Me.Enabled = True
        Case fcExit
            Unload Me
    End Select
    
    Exit Sub
ErrHandler:

End Sub

Private Function DoPrint(ReportFileName As String, _
                         LogOnInfo As PELogOnInfo, _
                         Optional Destination As Integer = 0, _
                         Optional ByVal SelectionFormula As String, _
                         Optional Formulas As Variant, _
                         Optional NumberOfCopies As Integer = 1, _
                         Optional SortFields As Variant = Empty, _
                         Optional PaperSize As Integer = DMPAPER_A4, _
                         Optional Format As Integer, _
                         Optional fileName As String, _
                         Optional SendTo As String, _
                         Optional subject As String, _
                         Optional Body As String, _
                         Optional Comment As String, _
                         Optional Width As Integer, _
                         Optional Height As Integer) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim result%, jobnum%, mainjob%, dialogflag%, resultlong&
    Dim textHandle&, textLength%, ErrorText$
    Dim JobInfo As PEJobInfo, TempText$
    Dim PrintData As Printer
    Dim curprinter%, defprinterpos%
    Dim summaryInfo As PEReportSummaryInfo
    Dim crDEVMODE As crDEVMODE
    Dim strTitle$
    Dim i As Integer
    Dim lngTextHandle As Long
    Dim intTextLength%
    Dim strBuffer$
    Dim strFormatDLLName As String
    Dim strReportTitle$
    Dim lngFormatType As Long
    Dim C As CError
    Dim ErrCode As Errors
    
    RESOIL2MAIN.MousePointer = vbHourglass
    ErrCode = PrintError
    
    ' Define size of JobInfo structure
    JobInfo.StructSize = PE_SIZEOF_JOB_INFO

    jobnum% = PEOpenPrintJob(ReportFileName)
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If

    result% = PESetNthTableLogOnInfo(jobnum%, 0, LogOnInfo, 1)
    
    If Len(SelectionFormula) > 0 Then
        result% = PEGetSelectionFormula(jobnum%, lngTextHandle, intTextLength%)
        strBuffer$ = Space$(intTextLength)
        result% = PEGetHandleString(lngTextHandle, strBuffer$, intTextLength%)
        If ASC(strBuffer$) <> 0 Then
            SelectionFormula = Left$(strBuffer$, intTextLength% - 1) & " and " & SelectionFormula
        End If
        result% = PESetSelectionFormula(jobnum%, SelectionFormula)
    End If
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If
    
    If Not IsEmpty(Formulas) Then
        For i = LBound(Formulas, 2) To UBound(Formulas, 2)
            result% = PESetFormula(jobnum%, Formulas(0, i), Formulas(1, i))
            If ErrorTrap(jobnum%) Then
                GoTo ErrHandler
            End If
        Next i
    End If
    
    If Not IsMissing(SortFields) Then
        If Not IsEmpty(SortFields) Then
            For i = LBound(SortFields) To UBound(SortFields)
                result% = PESetNthSortField(jobnum%, i, CStr(SortFields(i)), PE_SF_ASCENDING)
            Next i
        End If
    End If
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If
    
    dialogflag% = False
    result% = PEEnableProgressDialog(jobnum%, dialogflag%) ' Set whether or not to display progress dialog
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If
    
    Select Case Destination
        Case 0  ' Preview Window
            strTitle$ = mvarReportsList(0, cmbReport.ListIndex)
            result% = PEOutputToWindow(jobnum%, strTitle$, 0, 0, 500, 300, 0, RESOIL2MAIN.hwnd)
        Case 1  ' Printer
            result% = PEOutputToPrinter(jobnum%, NumberOfCopies) ' Second value is number of copies
        Case 2  ' Export
            strFormatDLLName = varFormatDLL(Format)
            lngFormatType = varFormatType(Format)
            If Len(fileName) = 0 Then
                ErrCode = InvalidFileName
                GoTo ErrHandler
            End If
            result% = crPEExportToDisk(jobnum%, fileName, strFormatDLLName, _
                                      lngFormatType, 0, 0, "", "")
        Case 3  ' E-mail
            strFormatDLLName = varFormatDLL(Format)
            lngFormatType = varFormatType(Format)
            If Len(SendTo) = 0 Then
                ErrCode = InvalidAddress
                GoTo ErrHandler
            End If
            result% = crPEExportToMapi(jobnum%, SendTo + Chr(0), _
                                       "" + Chr(0), subject, Body + Chr(0), strFormatDLLName, _
                                       lngFormatType, 0, 0, "", "")
    End Select
    
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If
    
    If Len(Comment) > 0 Then
        result% = PEGetReportTitle(jobnum%, textHandle&, textLength%)
        If ErrorTrap(jobnum%) Then
            GoTo ErrHandler
        End If
        strReportTitle$ = String$(textLength%, 0)
        result% = PEGetHandleString(textHandle&, strReportTitle$, textLength%)
        ErrorText$ = Trim$(strReportTitle$)
        ErrorText$ = Left$(strReportTitle$, Len(strReportTitle$) - 1)
        
        With summaryInfo
            .StructSize = PE_SIZEOF_REPORT_SUMMARY_INFO
            .comments = Comment & Chr(0)
            .Title = strReportTitle
        End With
        result% = PESetReportSummaryInfo(jobnum%, summaryInfo)
        If ErrorTrap(jobnum%) Then
            GoTo ErrHandler
        End If
    End If
    
    ' Get which printer was selected from the list
    curprinter% = cmbPrinters.ListIndex
    
    ' Set DevMode to whichever orientation was selected on the form
    If Width <= 0 Then
        crDEVMODE.dmFields = DM_PAPERSIZE Or DM_ORIENTATION
        crDEVMODE.dmPaperSize = PaperSize
    Else
        crDEVMODE.dmFields = DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_PAPERSIZE Or DM_ORIENTATION
        crDEVMODE.dmPaperSize = DMPAPER_USER
        crDEVMODE.dmPaperWidth = Width
        crDEVMODE.dmPaperLength = Height
    End If
    If optOrientation(0).value = True Then
        crDEVMODE.dmOrientation = DMORIENT_PORTRAIT
    Else
        crDEVMODE.dmOrientation = DMORIENT_LANDSCAPE
    End If
    
    ' NOTE - the dmFields element of the devmode structure must contain what elements of the devmode
    ' structure are being used. To set multiple elements, use the OR statement to combine the elements together
    With crDEVMODE
'        .dmDeviceName = Printers(curprinter%).DeviceName ' Name of the printer
'        .dmSpecVersion = &H320 ' Version of the DEVMODE structure - the structure used in this code is for Win32
'        .dmSize = 68 ' Size of the DevMode structure
'        .dmFields = .dmFields
    End With
    
    ' Do the SelectPrinter call
    result% = crPESelectPrinter(jobnum%, Printers(curprinter%).driverName & vbNullChar, Printers(curprinter%).DeviceName & vbNullChar, Printers(curprinter%).Port & vbNullChar, crDEVMODE)
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If
    ' Print the job
    result% = PEStartPrintJob(jobnum%, True) ' Note that flag must ALWAYS be true - setting to false will yield unpredictable results
    If ErrorTrap(jobnum%) Then
        GoTo ErrHandler
    End If
    ' Get job status information and display it
    Do
        result% = PEGetJobStatus(jobnum%, JobInfo)
        If ErrorTrap(jobnum%) Then
            GoTo ErrHandler
        End If
        DoEvents
    Loop While result% = PE_JOBINPROGRESS
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    Dim intErrorCode As Integer
    
    strMessage = "Si è verificato un errore durante la stampa"
    intErrorCode = PEGetErrorCode(jobnum%)
    lngNativeCode = CLng(intErrorCode)
    If ErrCode = PrintError Then
        result% = PEGetErrorText(jobnum%, textHandle&, textLength%)
        If result% = 0 Then
            ErrorText$ = "Errore sconosciuto"
        Else
            ' Get the text from the handle
            ErrorText$ = String$(textLength%, 0)
            result% = PEGetHandleString(textHandle&, ErrorText$, textLength%)
            ErrorText$ = Trim$(ErrorText$)
            ErrorText$ = Left$(ErrorText$, Len(ErrorText$) - 1)
        End If
    Else
        Select Case ErrCode
            Case InvalidFileName
                ErrorText$ = "Il nome del file non è valido"
            Case InvalidAddress
                ErrorText$ = "L'indirizzo non è valido"
        End Select
    End If
    strInfo = ErrorText$
    Set C = ErrH.SetError(ErrCode, "frmPrintReport", _
                "DoPrint", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError

EndProc:
    PEClosePrintJob jobnum%
    RESOIL2MAIN.MousePointer = vbDefault
    
End Function

Private Function ErrorTrap(ByVal jobnum%) As Boolean
    Dim errorcode%

    ' Check for any error codes
    errorcode% = PEGetErrorCode(jobnum%)
    ErrorTrap = errorcode% <> 0
End Function

Public Sub LoadReportsList(ReportsList As Variant, _
                           Optional SelectionFormula As String, _
                           Optional SortFields As Variant = Empty, _
                           Optional Formulas As Variant = Empty, _
                           Optional Comment As String = vbNullString)
    Dim i As Integer
    
    mstrSelectionFormula = SelectionFormula
    mvarSortFields = SortFields
    mvarReportsList = ReportsList
    mvarFormulas = Formulas
    mstrComment = Comment
    For i = LBound(mvarReportsList, 2) To UBound(mvarReportsList, 2)
        cmbReport.AddItem mvarReportsList(0, i)
    Next i
    cmbReport.ListIndex = 0
End Sub


Private Sub cmbReport_Click()
    Dim intOrientation As Integer
    
    intOrientation = mvarReportsList(2, cmbReport.ListIndex)
    optOrientation(intOrientation).value = True
    txtSubj = cmbReport.Text & " - Prodotto il " & Format$(Now, "dd/mm/yyyy")
End Sub


Private Sub cmdSelectDB_Click()
    With cdlOpenFile
        .fileName = txtFileName
        .DefaultExt = ".mdb"
        .DialogTitle = "Seleziona il file per l'esportazione... "
        .Filter = varFileExt(cmbFormat.ListIndex) & "Tutti i files (*.*)|*.*| "
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
        .ShowOpen
        txtFileName = .fileName
    End With
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Dim P As Printer
    Dim intDefaultPrinter As Integer
    Dim i As Integer
    Dim intErrorCode As Integer
    Dim strServer As String
    Dim strDB As String
    Dim strUID As String
    Dim strPWD As String
    Dim result%, jobnum%, mainjob%, dialogflag%, resultlong&
    Dim textHandle&, textLength%, ErrorText$
    
    
    varFormats = Array("Microsoft Word for Windows", _
                       "Microsoft Excel 5.0", _
                       "HTML", _
                       "Text", _
                       "Rich Text Format", _
                       "Data Interchange Format")
    varFileExt = Array("Documento Microsoft Word (*.doc)|*.doc|", _
                       "Documento Microsoft Excel (*.xls)|*.xls|", _
                       "Documento HTML (*.htm)|*.htm|", _
                       "Documento di Testo (*.txt)|*.txt|", _
                       "Rich Text Format (*.rtf)|*.rtf|", _
                       "Data Interchange Format (*.dif)|*.dif|")
    varFormatDLL = Array("U2FWORDW.DLL", _
                         "U2FXLS.DLL", _
                         "U2FHTML.DLL", _
                         "U2FTEXT.DLL", _
                         "U2FRTF.DLL", _
                         "U2FDIF.DLL")
    varFormatType = Array(crUXFWordWinType, _
                         crUXFXls5Type, _
                         crUXFHTML3Type, _
                         crUXFTextType, _
                         crUXFRichTextFormatType, _
                         crUXFDIFType)
                         
    Set ErrH = New CErrorHandler
    Dim C As CError
    
    ConfigureObjects Me
    
    ' --- Apro il motore di stampa dei report
    result% = PEOpenEngine()
    If result% = 0 Then
        intErrorCode = PEGetErrorCode(0)
        result% = PEGetErrorText(0, textHandle&, textLength%)
        If result% = 0 Then
            ErrorText$ = "Errore sconosciuto"
        Else
            ' Get the text from the handle
            ErrorText$ = String$(textLength%, 0)
            result% = PEGetHandleString(textHandle&, ErrorText$, textLength%)
            ErrorText$ = Trim$(ErrorText$)
            ErrorText$ = Left$(ErrorText$, Len(ErrorText$) - 1)
        End If
        Set C = ErrH.SetError(0, "frmPrintReport", _
                    "Form.Load", CLng(intErrorCode), "Si è verificato " & _
                    "un errore durante l'apertura del motore di stampa", 10, ErrorText$)
        DisplayError ErrH.LastError
        GoTo ErrHandler
    End If
    ' ---
    
    ' --- Mi collego alla data source ODBC
    With LogOnInfo
        .StructSize = PE_SIZEOF_LOGON_INFO
'        strDB = DA.ConnDB
'        strUID = DA.ConnUserName
'        strPWD = DA.ConnPWD
        strDB = DataConnect.ConnDB
        strUID = DataConnect.ConnUserName
        strPWD = DataConnect.ConnPWD
        
'        If DA.DataEngine = deJet Then
        If DataConnect.DataEngine = DEJET Then
            strServer = "RESOIL2_JET"
        Else
'            strServer = DA.ConnServer
            strServer = DataConnect.ConnServer
        End If
        
        .ServerName = strServer & Chr$(0)
        .DatabaseName = strDB & Chr$(0)
        .UserID = strUID & Chr$(0)
        .Password = strPWD & Chr$(0)
    End With
    ' ---
    
    ' --- Carico la lista delle stampanti definite nel sistema
    For Each P In Printers
        cmbPrinters.AddItem P.DeviceName
        If P.DeviceName = Printer.DeviceName Then
            intDefaultPrinter = cmbPrinters.NewIndex
        End If
    Next P
    cmbPrinters.ListIndex = intDefaultPrinter
    ' ---
    
    ' --- Carico la combobox per la scelta del numero di
    ' --- da stampare
    For i = 1 To 10
        cmbCopies.AddItem CStr(i)
    Next i
    cmbCopies.ListIndex = 0
    ' ---
    
    ' --- Carico la lista dei possibili formati per l'esportazione
    For i = LBound(varFormats) To UBound(varFormats)
        cmbFormat.AddItem varFormats(i)
    Next i
    cmbFormat.ListIndex = 0
    ' ---
    
    optDestination(0).value = True
    
    Exit Sub
ErrHandler:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PECloseEngine
    Set ErrH = Nothing
    mobjParentForm.Enabled = True
    Set frmPrintReport = Nothing
End Sub
Private Sub imgDestination_Click(Index As Integer)
    optDestination(Index).value = True
End Sub


Private Sub optDestination_Click(Index As Integer)
    Select Case Index
        Case 0  'Anteprima
            fraPrinter.Visible = True
            fraOrientation.Visible = True
            fraCopies.Visible = False
            fraFormat.Visible = False
            fraFileName.Visible = False
            fraSendTo.Visible = False
        Case 1  'Stampante
            fraPrinter.Visible = True
            fraOrientation.Visible = True
            fraCopies.Visible = True
            fraFormat.Visible = False
            fraFileName.Visible = False
            fraSendTo.Visible = False
        Case 2  'File
            fraPrinter.Visible = True
            fraOrientation.Visible = True
            fraCopies.Visible = False
            fraFormat.Visible = True
            fraFileName.Visible = True
            fraSendTo.Visible = False
        Case 3  'Email
            fraPrinter.Visible = True
            fraOrientation.Visible = True
            fraCopies.Visible = False
            fraFormat.Visible = True
            fraFileName.Visible = False
            fraSendTo.Visible = True
    End Select
End Sub


