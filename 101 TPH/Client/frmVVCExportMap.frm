VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmVVCExportMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRASFERIMENTO DESTINAZIONI A MAPSHORT"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   9270
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   2535
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3360
         Picture         =   "frmVVCExportMap.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3360
         Picture         =   "frmVVCExportMap.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.ListBox lstZone 
         Height          =   1860
         Left            =   5880
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   480
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Tag             =   "5"
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateTo 
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Tag             =   "6"
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Tag             =   "13"
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   327681
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   1920
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "&Zone"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblField3 
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
         Left            =   2640
         TabIndex        =   26
         Tag             =   "13"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unità di Trasporto"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5400
         Tag             =   "13"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3840
         Tag             =   "6"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data Inizio Viaggio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3840
         Tag             =   "5"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data Inizio Viaggio"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   5400
      TabIndex        =   10
      Top             =   2520
      Width           =   3855
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Consegne trasf. correttamente:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Consegne da trasferire:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   4680
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   4560
         Picture         =   "frmVVCExportMap.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmVVCExportMap.frx":0306
         TabIndex        =   20
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2400
         Picture         =   "frmVVCExportMap.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmVVCExportMap.frx":050A
         TabIndex        =   22
         Tag             =   "CUSTOM1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmVVCExportMap.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Consegne"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVCExportMap.frx":070E
         TabIndex        =   23
         Tag             =   "BROWSE"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7560
         Picture         =   "frmVVCExportMap.frx":0810
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmVVCExportMap.frx":0912
         TabIndex        =   6
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
      Top             =   5565
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10901
            Key             =   "Msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10.00"
            Key             =   "Time"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16/12/99"
            Key             =   "Date"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmVVCExportMap.frx":0A14
      TabIndex        =   5
      Tag             =   "SELECT"
      Top             =   2640
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVVCExportMap.frx":2859
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVVCExportMap.frx":296B
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVVCExportMap.frx":2A7D
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVVCExportMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 12 '11
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mobjOutFile As CTextFile
Private mobjOutFile2 As CTextFile

Public GridArraySelect As ArrayGridCls
Public Loading As Boolean

Public DataClass As CViaggioSelezione
Public State As FormState

Private AZO As CZona



Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
                cmdOpenFile.Enabled = True
                If Len(mskFileName) = 0 Then
                    mskFileName = App.Path & "\Mapshort.inp"
                End If
                cmdExport.Enabled = True
            Else
                cmdOpenFile.Enabled = False
                cmdExport.Enabled = False
            End If
        Case fcExport
            ExportRecords
            cmdExport.Enabled = False
        Case fcCustom1
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per Mapshort... "
                .InitDir = App.Path
                .Filter = "(*.inp)|*.inp|Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                .ShowSave
                mskFileName = .fileName
            End With
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub ExportRecords()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim i As Integer
    Dim strIDZona As String
    Dim strFileName As String
    Dim strErrorMessage As String
    Dim intRowError As Errors
    Dim blnFound As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    strFileName = mskFileName
    If Not mobjOutFile.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    ' --- Ciclo di elaborazione degli ordini
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        
        With mobjOutFile
            .ClearFields
            .Item("1").value = GridArraySelect.value(lngRow, 0)
            .Item("2").value = GridArraySelect.value(lngRow, 1)
            .Item("3").value = "0"
        
            If Not mobjOutFile.WriteLine Then
                intRowError = OutputError
                GoTo LogResult
            End If
            
            .ClearFields
            .Item("1").value = GridArraySelect.value(lngRow, 0)
            ' --- Formatto il codice destinazione in modo che
            '     Mapshort lo interpreti correttamente
            If IsNumeric(GridArraySelect.value(lngRow, 2)) Then
                .Item("2").value = CLng(GridArraySelect.value(lngRow, 2))
            End If
            ' ---
            .Item("3").value = "0"
        
            If Not mobjOutFile.WriteLine Then
                intRowError = OutputError
                GoTo LogResult
            End If
        End With
        
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione dell'ordine
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            lngCountFail = lngCountFail + 1
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        TDBGrid1.Bookmark = lngRow
        ' ---
        
        DoEvents
        
    Next lngRow
    ' ---
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & strFileName
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFile.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub


Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim j As Integer
    
    Me.MousePointer = vbHourglass
    With DataClass.DS
        If .Fields.Validate Then
            For i = 0 To lstZone.ListCount - 1
                If lstZone.Selected(i) Then
                    j = j + 1
                    .Fields("17").FindValue(j) = AZO.DS.ListItem(0, i)
                End If
            Next i
            If .ScanClientCond Then
                If Not DataClass.GetSelectionExportDest() Then
                    GoTo ErrHandler
                End If
                DisplaySelection
                GoTo EndProc
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    GoTo EndProc
ErrHandler:
    Me.MousePointer = vbDefault
    SelectRecords = False
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function

Private Sub DisplaySelection()
    Dim varLog As Variant
    Dim i As Long
    Dim lngRows As Long
    
    lngRows = DataClass.DS.HowMany
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
End Sub



Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Dalla data inizio viaggio"
        Case 1
            Set M = mskDateTo
            strTitle = "Alla data inizio viaggio"
    End Select
    GetCalendarDate M, strTitle

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim D As Variant
    Dim i As Long
    
    Loading = True
    Set GridArraySelect = New ArrayGridCls
    
    mskFileName = gmRegistry.GetSetting("Options", "MapshortInput", "mapshort.inp", True)
    
    D = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
    mskDateTo.SelText = D
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    AZO.GetList True
    
    For i = 0 To UBound(AZO.DS.List, 2)
        lstZone.AddItem AZO.DS.ListItem(1, i)
        If AZO.DS.ListItem(2, i) = 0 Then
            lstZone.Selected(lstZone.NewIndex) = True
        End If
    Next i
    
    Set mobjOutFile = New CTextFile
    With mobjOutFile
        .Add "Codice Viaggio", 7, "1"
        .Add "Luogo", 18, "2"
        .Add "Zero", 1, "3"
    End With
    
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "MapshortInput", mskFileName, True
    Set GridArraySelect = Nothing
    Set AZO = Nothing
    Set DataClass = Nothing
    Set mobjOutFile = Nothing
    Set frmVVCExportMap = Nothing
End Sub


Private Sub mskFileName_Change()
    If Not Loading Then
        cmdExport.Enabled = (Len(mskFileName) > 0)
    End If
End Sub


