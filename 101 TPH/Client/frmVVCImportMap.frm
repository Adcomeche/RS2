VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVVCImportMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE DISTANZE DA MAPSHORT"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7680
         Picture         =   "frmVVCImportMap.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmVVCImportMap.frx":0102
         TabIndex        =   13
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4365
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
      Height          =   3375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   2175
         Left            =   6960
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Distanze da importare:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   19
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Distanze importate:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   6735
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmVVCImportMap.frx":0204
            TabIndex        =   5
            Top             =   240
            Width           =   6495
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
         Picture         =   "frmVVCImportMap.frx":3DDE
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmVVCImportMap.frx":4310
         TabIndex        =   2
         Tag             =   "IMPORT"
         Top             =   600
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
         Picture         =   "frmVVCImportMap.frx":4412
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmVVCImportMap.frx":4944
         TabIndex        =   1
         Tag             =   "OPEN"
         Top             =   600
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
         Picture         =   "frmVVCImportMap.frx":4A46
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmVVCImportMap.frx":4B48
         TabIndex        =   0
         Tag             =   "BROWSE"
         Top             =   600
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2760
         Top             =   960
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
               Picture         =   "frmVVCImportMap.frx":4C4A
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVCImportMap.frx":4D5C
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVCImportMap.frx":4E6E
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmVVCImportMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3
Private XA As XArrayDB
Private VVC As CViaggioConsegna
Private mobjInFile As CTextFile
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
                .DialogTitle = "Seleziona il file delle distanze... "
                .InitDir = App.Path
                .Filter = "(*.out)|*.out|Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
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

    Dim strFileName As String
    Dim strBuffer As String
    Dim strIDTrip As String
    Dim strDestinazione As String
    Dim strErrorMessage As String
    Dim lngDistanza As Long
    Dim lngCountTrips As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngNumRows As Long
    Dim Row As Long
'    Dim Col As Integer
    Dim i As Integer
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError

    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError

    lngCountToGo = lblLog(5)
    lngNumRows = lblLog(5)
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If

    For Row = 1 To lngNumRows
        intRowError = NoError
        With mobjInFile
            strBuffer = .GetLine(False)
            .ParseFixedLenFields strBuffer, 1, 5
            strIDTrip = .Item("1").value
            strBuffer = .GetLine(False)
            .ParseFixedLenFields strBuffer, 1, 5
            strDestinazione = Format$(.Item("2").value, "00000000")
            lngDistanza = .Item("4").value
            If Not VVC.UpdateKmDef(strIDTrip, strDestinazione, lngDistanza) Then
                intRowError = DeliveryUpdateError
            End If
        End With
LogResult:
        If intRowError <> NoError Then
            lngCountFail = lngCountFail + 1
        End If
        
        If intRowError = NoError Then
            VVC.DS.DA.EndTrans 0
        Else
            VVC.DS.DA.EndTrans 1
        End If
        
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        End If
        lngCountToGo = lngCountToGo - 1
        lngCountTrips = lngCountTrips + 1
        TDBGrid1.Bookmark = lngCountTrips * 2
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        DoEvents
    Next Row
    ' ---

    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long

    VVC.DS.DA.EndTrans 1
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Il file non è stato trovato."
            strInfo = "Nome del file: " & strFileName
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
    Set C = ErrH.SetError(ErrCode, "frmVVCImportMap", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    DisplayError ErrH.LastError
EndProc:
    VVC.DS.DA.EndTrans 0
    mobjInFile.CloseFile
    lblLog(4) = lngCountFail
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Sub OpenTextFile()
    On Error GoTo ErrHandler
    
    Dim strFileName As String
    Dim strLine As String
    Dim strBuffer As String
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim lngRows As Long
    Dim lngCountRecords As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    cmdImport.Enabled = False
    Set XA = New XArrayDB
    lngRows = 1
    Do While Not mobjInFile.EOF
        strBuffer = mobjInFile.GetLine(False)
        mobjInFile.ParseFixedLenFields strBuffer, 1, 5
        With XA
            .ReDim 1, lngRows, 0, 5
            .value(.UpperBound(1), 0) = mobjInFile.Item("1").value
            .value(.UpperBound(1), 1) = mobjInFile.Item("2").value
            .value(.UpperBound(1), 2) = mobjInFile.Item("3").value
            .value(.UpperBound(1), 3) = mobjInFile.Item("4").value
            .value(.UpperBound(1), 4) = mobjInFile.Item("5").value
        End With
        lngRows = lngRows + 1
    Loop
        
    cmdImport.Enabled = True
    With TDBGrid1
        .Array = XA
        .ReBind
        .Bookmark = 1
    End With
    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = (lngRows - 1) / 2
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmVVCImportMap", _
                        "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmVVCImportMap", _
                                "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub



Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    
    ConfigureObjects Me
    mskFileName = gmRegistry.GetSetting("Options", "MapshortOutput", "mapshort.out", True)
    
    Set ErrH = New CErrorHandler
    
    Set VVC = New CViaggioConsegna
    With VVC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
    With mobjInFile
        .Add "Numero Viaggio", 7, "1"
        .Add "Luogo", 10, "2"
        .Add "xxx", 9, "3"
        .Add "Distanza", 6, "4"
        .Add "yyy", 10, "5"
    End With
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "MapshortOutput", mskFileName, True
    Set XA = Nothing
    Set VVC = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set frmVVCImportMap = Nothing
End Sub







Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub


