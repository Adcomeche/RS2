VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOCTExportMover1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRASFERIMENTO ORDINI A MOVER"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   9270
   Begin VB.Frame Frame4 
      Height          =   1575
      Left            =   0
      TabIndex        =   26
      Top             =   3840
      Width           =   5295
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmOCTExportMover1.frx":0000
         TabIndex        =   13
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   1935
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox ChOrdiniAcq 
         Caption         =   "&Solo Ordini Acquisiti"
         Height          =   255
         Left            =   7080
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmOCTExportMover1.frx":2A04
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frmOCTExportMover1.frx":2B06
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "25"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateTo 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Tag             =   "26"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Tag             =   "20"
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   5400
         TabIndex        =   34
         Tag             =   "2"
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3600
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox MaskFileVeicoli 
         Height          =   255
         Left            =   5400
         TabIndex        =   37
         Top             =   1440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File &Veicoli:"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   8880
         Tag             =   "2"
         Top             =   360
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
         Left            =   6000
         TabIndex        =   35
         Tag             =   "2"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emit."
         Height          =   255
         Index           =   16
         Left            =   4560
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   3240
         Tag             =   "0"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &N. Ordine"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   3240
         Tag             =   "20"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N. &Ordine"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File:"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3720
         Tag             =   "26"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data Cons. Pian."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3720
         Tag             =   "25"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data Cons. Pian."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   5400
      TabIndex        =   18
      Top             =   3840
      Width           =   3855
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini trasferiti correttamente:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini da trasferire:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   5520
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
         Picture         =   "frmOCTExportMover1.frx":2C08
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmOCTExportMover1.frx":2D0A
         TabIndex        =   29
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmOCTExportMover1.frx":2E0C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
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
         Picture         =   "frmOCTExportMover1.frx":2F0E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportMover1.frx":3010
         TabIndex        =   31
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Ordini"
         Height          =   375
         Left            =   2280
         Picture         =   "frmOCTExportMover1.frx":3112
         TabIndex        =   32
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
         Picture         =   "frmOCTExportMover1.frx":3214
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmOCTExportMover1.frx":3316
         TabIndex        =   14
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6465
      Width           =   9270
      _ExtentX        =   16351
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
            Object.Width           =   10901
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
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "frmOCTExportMover1.frx":3418
      TabIndex        =   12
      Tag             =   "SELECT"
      Top             =   2040
      Width           =   9015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportMover1.frx":A586
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportMover1.frx":A698
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportMover1.frx":A7AA
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTExportMover1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 12 '11
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mobjOutFile As CTextFile
Private mobjOutFileVeicoli As CTextFile

Public GridArraySelect As ArrayGridCls
Public XA As XArrayDB
Public Loading As Boolean

Public DataClass As COrdineSelezione
Public State As FormState

Private Soc As CSocieta
Private AXX As CParametro
Private ADZ As CDestinazione
Private AZO As CZona
Private OCR As COrdineRiga
Private AUT As CUnitaDiTrasporto
Private APR As CProdotto

Private mstrExportFileName As String
Private blFocusOnFile As Boolean
Private blFocusOnVeicoli As Boolean
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
                Dim i As Long
                Dim Rows As Long
                Rows = GridArraySelect.RowCount - 1
                With TDBGrid1.SelBookmarks
                    For i = 0 To Rows
                        .Add " " & CStr(i)
                    Next i
                End With
                cmdOpenFile.Enabled = True
                If Len(mskFileName) = 0 Then
                    mskFileName = ""
                End If
                If DataClass.DS.HowMany = 0 Then
                    cmdExport.Enabled = False
                Else
                    cmdExport.Enabled = True
                End If
            Else
                cmdExport.Enabled = False
            End If
        Case fcExport
            ExportRecords
            ExportDataVeicoli
            cmdExport.Enabled = False
        Case fcOpen
            If Not blFocusOnVeicoli Then
                With CommonDialog1
                    .fileName = mskFileName
                    .DialogTitle = "Seleziona il file per gli Ordini a Mover... "
                    .InitDir = App.Path
                    .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                    .FilterIndex = 1
                    .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                    .ShowSave
                    mskFileName = .fileName
                End With
            Else
                With CommonDialog1
                    .fileName = MaskFileVeicoli
                    .DialogTitle = "Seleziona il file per i Veicoli a Mover... "
                    .InitDir = App.Path
                    .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                    .FilterIndex = 1
                    .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                    .ShowSave
                    MaskFileVeicoli = .fileName
                End With
            End If
        Case fcExit
            Unload Me
    End Select
End Sub
Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.Bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.Bookmark, 0)
    End If
End Function
Private Sub ExportRecords()
    On Error GoTo ErrHandler

    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim Rows As Long
    Dim lngQta As Long
   
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim intRowsOrder As Integer
    Dim i As Long
    Dim j As Long
    Dim intIndexProdotto As Integer
    Dim strIDOrdine As String
    Dim strIDProdotto As String
    Dim strIDPosDog As String
    Dim strIDSocDefault As String
    Dim lngQtaColt As Long
    Dim strIDParam As String
    Dim strIDZona As String
    Dim StrIDDest As String
    Dim strRowState As String
    Dim strFileName As String
    Dim strErrorMessage As String
    Dim intRowError As Errors
    Dim varProdotti(1 To 10) As String
    Dim lngCap(1 To 10)  As Long
    Dim blnFlagProdottiOK As Boolean
    Dim blnFound As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    Dim VarQtaOrdinate As Variant
    Dim IntResQta As Integer
    Dim intCont As Integer
    Dim lngTmpCap As Long
    Dim varDummy As Variant
        
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDSocDefault = AXX.DS.Fields("4").value   ' Società di Default
        If strIDSocDefault = vbNullString Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        For i = 1 To 10
            If Trim$(AXX.DS.Fields(i + 5).value) <> vbNullString Then
                blnFlagProdottiOK = True
                varProdotti(i) = Trim$(AXX.DS.Fields(i + 5).value) ' Contiene i Prodotti di AXX
            End If
        Next i
        If Not blnFlagProdottiOK Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
    End With
    ' ---

    strFileName = mskFileName
    If Not mobjOutFile.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    lngNumRows = GridArraySelect.RowCount - 1
'    lngCountToGo = lngNumRows + 1
    
    Rows = TDBGrid1.SelBookmarks.Count - 1
    lngCountToGo = Rows + 1
'    If Rows >= 0 Then
    For j = 0 To Rows
'    For lngRow = 0 To lngNumRows  ' --- Ciclo di elaborazione degli ordini
        lngRow = CLng(TDBGrid1.SelBookmarks(j))
        intRowError = NoError
        With mobjOutFile
            .ClearFields
            .Item("0").value = strIDSocDefault                  ' Societa di Def.
            .Item("1").value = GridArraySelect.value(lngRow, 0) ' Dest
            .Item("2").value = GridArraySelect.value(lngRow, 9) ' Clie
            .Item("3").value = GridArraySelect.value(lngRow, 1) ' Rasoc
            .Item("4").value = GridArraySelect.value(lngRow, 2) ' Indir
            .Item("5").value = GridArraySelect.value(lngRow, 3) ' Local
            .Item("6").value = GridArraySelect.value(lngRow, 4) ' Prov
            .Item("7").value = GridArraySelect.value(lngRow, 5) ' CAP
            .Item("8").value = GridArraySelect.value(lngRow, 6) ' Telefono
            .Item("13").value = GridArraySelect.value(lngRow, 8) ' Ordine
            .Item("14").value = GridArraySelect.value(lngRow, 12) ' Tipo Ordine
            .Item("15").value = GridArraySelect.value(lngRow, 10) ' Deposito
            .Item("16").value = GridArraySelect.value(lngRow, 13) ' Ente Emit
            .Item("17").value = Format$(GridArraySelect.value(lngRow, 11), "YYYYMMDD") ' Data Consegna
            .Item("18").value = GridArraySelect.value(lngRow, 7) ' Note
        End With
        StrIDDest = GridArraySelect.value(lngRow, 0)
        For intCont = 2 To UBound(varProdotti)   ' Ciclo di Caricamento delle Capacita cisterne
             If Trim(varProdotti(intCont)) <> vbNullString Then
                lngTmpCap = CLng(ADZ.ContaCapacitaCisterne(StrIDDest, varProdotti(intCont)))
                If lngTmpCap > 0 Then
                    lngCap(intCont) = lngTmpCap / 1000
                Else
                    lngCap(intCont) = 0
                End If
             End If
        Next intCont
        strIDOrdine = GridArraySelect.value(lngRow, 8)
        OCR.DS.Fields(OCR.DS.KeyFields(1)).FindValue = strIDOrdine
        If Not OCR.GetList(True) Then  ' Verifico se esiste una riga d'ordine
            intRowError = RowsNotFound
            GoTo LogResult
        End If

        For i = 2 To 5 ' Per i 4 prodotti di mover
            If varProdotti(i) <> vbNullString Then
                With APR.DS.MultiSoc
                    .DS.Fields("Codice").FindValue = varProdotti(i)
                    .DS.Fields("Società").FindValue = strIDSocDefault
                    If Not .GetList Then
                        intRowError = NoProdotto
                        GoTo LogResult
                    End If
                    varDummy = .DS.List
'                    strIDPosDog = varDummy(2, 0)  ' Posizione doganale del Prodotto
                End With
                ' Ricavo dall'ordine le Qtà Iniziali e Concordate per il prodotto
                IntResQta = OCR.GetQtaFromOrderRow(strIDOrdine, varProdotti(i), VarQtaOrdinate)
                If IntResQta > 0 Then
                    lngQta = CLng(VarQtaOrdinate(0, 0)) / 1000
                    lngQtaColt = CLng(VarQtaOrdinate(1, 0)) / 1000
                    strIDPosDog = VarQtaOrdinate(2, 0)
                Else
                    lngQta = 0
                    lngQtaColt = 0
                    strIDPosDog = varDummy(2, 0)
                End If
            Else
                lngQta = 0
                lngQtaColt = 0
                strIDPosDog = vbNullString
            End If
            mobjOutFile.Item(CStr(7 + i)).value = lngCap(i)
            mobjOutFile.Item(CStr(17 + i)).value = varProdotti(i) ' Codice Prodotto
            If varProdotti(i) <> vbNullString Then
                mobjOutFile.Item(CStr(21 + i)).value = strIDPosDog   ' Posizione Doganale
            End If
            mobjOutFile.Item(CStr(25 + i)).value = lngQta        ' Quantità Iniziale
            mobjOutFile.Item(CStr(29 + i)).value = lngQtaColt    ' Quantità Concordata
        Next i
        
'        intRowsOrder = OCR.DS.HowMany - 1
'        For i = 0 To intRowsOrder                   ' Ciclo Scansione Riga Ordine
'            strRowState = OCR.DS.ListItem(9, i)
'            If strRowState = "A" Then
'                strIDProdotto = OCR.DS.ListItem(1, i) ' Codice Prodotto
'                StrIDPosDog = OCR.DS.ListItem(2, i)   ' Posizione Doganale
'                If IsEmpty(OCR.DS.ListItem(4, i)) Then   ' Quantita Iniziale
'                    intRowError = MissingQuantity
'                    GoTo LogResult
'                End If
'                lngQta = CLng(OCR.DS.ListItem(4, i))
'
'                If lngQta = 0 Then
'                    intRowError = MissingQuantity
'                    GoTo LogResult
'                Else
'                    lngQta = lngQta / 1000
'                End If
'
'                If IsEmpty(OCR.DS.ListItem(5, i)) Then   ' Quantita Concordata
'                    intRowError = MissingQuantity
'                    GoTo LogResult
'                End If
'                lngQtaColt = CLng(OCR.DS.ListItem(5, i))
'                If lngQtaColt = 0 Then
'                    intRowError = MissingQuantity
'                    GoTo LogResult
'                Else
'                    lngQtaColt = lngQtaColt / 1000
'                End If
'
'                intIndexProdotto = 1
'                blnFound = False
'                Do While intIndexProdotto <= 10 And Not blnFound
'                    If strIDProdotto = varProdotti(intIndexProdotto) Then
'                        blnFound = True
'                    Else
'                        intIndexProdotto = intIndexProdotto + 1
'                    End If
'                Loop
'                intIndexProdotto = intIndexProdotto - 1
'                If Not blnFound Then
'                    intRowError = NoProductInParam
'                    GoTo LogResult
'                End If
'                      ' Capacità
'
'
'            End If
'        Next i
        If Not mobjOutFile.WriteLine Then
            intRowError = OutputError
            GoTo LogResult
        End If
        ' --- Modifico lo stato dell'ordine passato a Mover
        With DataClass
            .DS.Fields("24").FindValue = strIDOrdine
            If Not .SetAvanzamento("M") Then
                intRowError = NoSetState
                GoTo LogResult
            End If
        End With
        ' ---

LogResult:
        ' --- Scrivo il risultato dell'elaborazione dell'ordine
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            If intRowError = ZonaEsclusaExport Then
                lngCountExcluded = lngCountExcluded + 1
            Else
                lngCountFail = lngCountFail + 1
            End If
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
         
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case OutputError
                strErrorMessage = "Errore durante la scrittura sul file"
            Case NoProductInParam
                strErrorMessage = "Prodotto non trovato tra i parametri di sistema"
            Case RowsNotFound
                strErrorMessage = "Errore durante la lettura delle righe di ordine"
            Case NoSetState
                strErrorMessage = "Impossibile cambiare lo stato dell'ordine"
            Case MissingQuantity
                strErrorMessage = "Quantità mancante in una riga di ordine"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case ZonaEsclusaExport
                strErrorMessage = "Zona esclusa dal trasferimento a Mover"
            'Case MissingCapacity
            '    strErrorMessage = "Capacità Mancante"
                
        End Select
        XA.value(lngRow, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = lngRow
        End With
        TDBGrid1.Bookmark = lngRow
        ' ---

        DoEvents

'    Next lngRow
    Next j
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
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti: Prodotti"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover1", _
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
    
    Me.MousePointer = vbHourglass
    If DataClass.DS.Fields.Validate Then
        DataClass.DS.Fields.ResetFields
        If DataClass.DS.ScanClientCond Then
            If Not DataClass.GetSelectionExport1(ChOrdiniAcq.value) Then
                GoTo ErrHandler
            End If
            
            DisplaySelection
            GoTo EndProc
        End If
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
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
    If DataClass.DS.HowMany > 0 Then
        cmdExport.Enabled = True
    End If
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
     
    If lngRows > 0 Then
    ReDim varLog(0 To 1, 0 To lngRows - 1)
    With DataClass.DS
        For i = 0 To lngRows - 1
            varLog(0, i) = .ListItem(8, i)
        Next i
    End With
    RefreshXArray XA, tdbLog, varLog
    Else
        ClearXArray XA, tdbLog
    End If
End Sub

Private Sub ChOrdiniAcq_Click()
    If ChOrdiniAcq Then    ' Propongo il nome file in base al flag ordini acquisiti
        mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile1", , True)
    Else
        mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile", , True)
    End If
    blFocusOnVeicoli = False
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Dalla data cons. pian."
        Case 1
            Set M = mskDateTo
            strTitle = "Alla data cons. pian."
    End Select
    GetCalendarDate M, strTitle

End Sub

 

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim D As Variant
    
    Loading = True
    Set GridArraySelect = New ArrayGridCls
    Set XA = New XArrayDB
    mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile", , True)
    If mskFileName = vbNullString Then
       cmdOpenFile.Enabled = True
    End If
    MaskFileVeicoli = gmRegistry.GetSetting("Options", "AUTExportFile", , True)
    If MaskFileVeicoli = vbNullString Then
       cmdOpenFile.Enabled = True
    End If
    cmdExport.Enabled = False
    D = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
    mskDateTo.SelText = D
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New COrdineSelezione
    With DataClass.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set Soc = New CSocieta
'    Set Soc.DS.DA = dataconnect.da
    Set Soc.DS.DA = DataConnect.DA
    SetValueItems TDBGrid1.Columns(1), Soc
    
    Set AXX = New CParametro
    With AXX.DS
'        Set .DA = DataConnect.DA
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
    
    Set OCR = New COrdineRiga
    With OCR.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
        With AUT.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    Set APR = New CProdotto
    With APR.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    
    Set mobjOutFile = New CTextFile ' Struttura del File
    With mobjOutFile
        .Add "Societa", 2, "0"
        .Add "Codice Destinazione", 6, "1"
        .Add "Codice Cliente", 10, "2"
        .Add "Ragione Sociale", 30, "3"
        .Add "Indirizzo", 30, "4"
        .Add "Località", 20, "5"
        .Add "Provincia", 2, "6"
        .Add "CAP", 5, "7"
        .Add "Telefono", 11, "8"
        .Add "Cap. Prodotto 1", 3, "9", 0, fftNumerico
        .Add "Cap. Prodotto 2", 3, "10", 0, fftNumerico
        .Add "Cap. Prodotto 3", 3, "11", 0, fftNumerico
        .Add "Cap. Prodotto 4", 3, "12", 0, fftNumerico
        .Add "Ordine", 7, "13"
        .Add "Tipo Ordine", 2, "14"
        .Add "Deposito", 3, "15"
        .Add "Ente Emittente", 3, "16"
        .Add "Data Consegna", 8, "17"
        .Add "Note", 30, "18"
        
        .Add "Codice Prodotto 1", 6, "19"
        .Add "Codice Prodotto 2", 6, "20"
        .Add "Codice Prodotto 3", 6, "21"
        .Add "Codice Prodotto 4", 6, "22"
        
        .Add "Posizione Doganale Prodotto 1", 3, "23"
        .Add "Posizione Doganale Prodotto 2", 3, "24"
        .Add "Posizione Doganale Prodotto 3", 3, "25"
        .Add "Posizione Doganale Prodotto 4", 3, "26"
        
        .Add "Qta Iniziale Prodotto 1", 2, "27", 0, fftNumerico
        .Add "Qta Iniziale Prodotto 2", 2, "28", 0, fftNumerico
        .Add "Qta Iniziale Prodotto 3", 2, "29", 0, fftNumerico
        .Add "Qta Iniziale Prodotto 4", 2, "30", 0, fftNumerico
        
        .Add "Qta Concordata Prodotto 1", 2, "31", 0, fftNumerico
        .Add "Qta Concordata Prodotto 2", 2, "32", 0, fftNumerico
        .Add "Qta Concordata Prodotto 3", 2, "33", 0, fftNumerico
        .Add "Qta Concordata Prodotto 4", 2, "34", 0, fftNumerico
    End With
    
    Set mobjOutFileVeicoli = New CTextFile
    With mobjOutFileVeicoli
        .Add "Unità di trasporto", 3, "1"
        .Add "Targa Motrice", 8, "2"
        .Add "Targa Rimorchio", 8, "3"
        .Add "Deposito", 3, "4"
        .Add "Ente Emittente", 3, "5"
        .Add "Descrizione", 30, "6"
        .Add "Peso Totale Kg", 5, "7", 0, fftNumerico
        .Add "Volume Totale", 5, "8", 0, fftNumerico
        .Add "Tipo Scarico", 1, "9"
        .Add "Codice Contratto", 3, "10"
        .Add "Codice Autista 1", 3, "11"
        .Add "Nome Autista 1", 30, "12"
        .Add "Codice Autista 2", 3, "13"
        .Add "Nome Autista 2", 30, "14"
        .Add "Peso Motrice in Kg", 5, "15", 0, fftNumerico
        .Add "Volume Motrice", 5, "16", 0, fftNumerico
        .Add "Peso Rimorchio in Kg", 5, "17", 0, fftNumerico
        .Add "Volume Rimorchio", 5, "18", 0, fftNumerico
        .Add "Capacita Scomparto 1 Motrice", 2, "19", 0, fftNumerico
        .Add "Numero Scomparti 1 Motrice", 2, "20", 0, fftNumerico
        .Add "Capacita Scomparto 2 Motrice", 2, "21", 0, fftNumerico
        .Add "Numero Scomparti 2 Motrice", 2, "22", 0, fftNumerico
        .Add "Capacita Scomparto 3 Motrice", 2, "23", 0, fftNumerico
        .Add "Numero Scomparti 3 Motrice", 2, "24", 0, fftNumerico
        .Add "Capacita Scomparto 4 Motrice", 2, "25", 0, fftNumerico
        .Add "Numero Scomparti 4 Motrice", 2, "26", 0, fftNumerico
        .Add "Capacita Scomparto 5 Motrice", 2, "27", 0, fftNumerico
        .Add "Numero Scomparti 5 Motrice", 2, "28", 0, fftNumerico
        .Add "Capacita Scomparto 6 Motrice", 2, "29", 0, fftNumerico
        .Add "Numero Scomparti 6 Motrice", 2, "30", 0, fftNumerico
        .Add "Capacita Scomparto 7 Motrice", 2, "31", 0, fftNumerico
        .Add "Numero Scomparti 7 Motrice", 2, "32", 0, fftNumerico
        .Add "Capacita Scomparto 8 Motrice", 2, "33", 0, fftNumerico
        .Add "Numero Scomparti 8 Motrice", 2, "34", 0, fftNumerico
        .Add "Capacita Scomparto 9 Motrice", 2, "35", 0, fftNumerico
        .Add "Numero Scomparti 9 Motrice", 2, "36", 0, fftNumerico
        .Add "Capacita Scomparto 10 Motrice", 2, "37", 0, fftNumerico
        .Add "Numero Scomparti 10 Motrice", 2, "38", 0, fftNumerico
        .Add "Capacita Scomparto 1 Rimorchio", 2, "39", 0, fftNumerico
        .Add "Numero Scomparti 1 Rimorchio", 2, "40", 0, fftNumerico
        .Add "Capacita Scomparto 2 Rimorchio", 2, "41", 0, fftNumerico
        .Add "Numero Scomparti 2 Rimorchio", 2, "42", 0, fftNumerico
        .Add "Capacita Scomparto 3 Rimorchio", 2, "43", 0, fftNumerico
        .Add "Numero Scomparti 3 Rimorchio", 2, "44", 0, fftNumerico
        .Add "Capacita Scomparto 4 Rimorchio", 2, "45", 0, fftNumerico
        .Add "Numero Scomparti 4 Rimorchio", 2, "46", 0, fftNumerico
        .Add "Capacita Scomparto 5 Rimorchio", 2, "47", 0, fftNumerico
        .Add "Numero Scomparti 5 Rimorchio", 2, "48", 0, fftNumerico
        .Add "Capacita Scomparto 6 Rimorchio", 2, "49", 0, fftNumerico
        .Add "Numero Scomparti 6 Rimorchio", 2, "50", 0, fftNumerico
        .Add "Capacita Scomparto 7 Rimorchio", 2, "51", 0, fftNumerico
        .Add "Numero Scomparti 7 Rimorchio", 2, "52", 0, fftNumerico
        .Add "Capacita Scomparto 8 Rimorchio", 2, "53", 0, fftNumerico
        .Add "Numero Scomparti 8 Rimorchio", 2, "54", 0, fftNumerico
        .Add "Capacita Scomparto 9 Rimorchio", 2, "55", 0, fftNumerico
        .Add "Numero Scomparti 9 Rimorchio", 2, "56", 0, fftNumerico
        .Add "Capacita Scomparto 10 Rimorchio", 2, "57", 0, fftNumerico
        .Add "Numero Scomparti 10 Rimorchio", 2, "58", 0, fftNumerico
    End With
    Loading = False
    blFocusOnVeicoli = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If ChOrdiniAcq Then
        gmRegistry.SaveSetting "Options", "OrderExportFile1", mskFileName, True
    Else
        gmRegistry.SaveSetting "Options", "OrderExportFile", mskFileName, True
    End If
    gmRegistry.SaveSetting "Options", "AUTExportFile", MaskFileVeicoli, True
    Set XA = Nothing
    Set GridArraySelect = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set OCR = Nothing
    Set AUT = Nothing
    Set APR = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set mobjOutFile = Nothing
    Set mobjOutFileVeicoli = Nothing
    Set frmOCTExportMover1 = Nothing
    Set mobjOutFile = Nothing
End Sub

Private Sub mskFileName_Change()
    If Len(Trim(mskFileName)) = 0 Then
        cmdOpenFile.Enabled = True
        cmdExport.Enabled = False
    Else
        cmdExport.Enabled = True
    End If
End Sub
Private Function ExportDataVeicoli() As Boolean
    
    
    On Error GoTo ErrHandler
    
   
    Dim varExportDataVeicoli As Variant
    Dim blnResult As Boolean
    Dim intFileNum As Integer
    Dim intRows As Integer
    Dim intCols As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ErrCode As Errors
    Dim C As CError
    Dim strIDAEE As String
    Dim strErrorMessage As String
    
    DoEvents
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    
    strIDAEE = MaskEdBox1.Text
    If strIDAEE = vbNullString Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
   
    blnResult = AUT.GetExportDataVeicoli(strIDAEE, varExportDataVeicoli)
    If Not blnResult Then
        ErrCode = RowsNotFound
        GoTo ErrHandler
    End If
    
    If Not mobjOutFileVeicoli.OpenFile(MaskFileVeicoli, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    intFileNum = mobjOutFileVeicoli.FileNum
    intRows = UBound(varExportDataVeicoli, 2)
    intCols = UBound(varExportDataVeicoli, 1)
    For i = 0 To intRows
        mobjOutFileVeicoli.ClearFields
        mobjOutFileVeicoli.Item("1").value = varExportDataVeicoli(0, i)   ' UDT
        mobjOutFileVeicoli.Item("2").value = varExportDataVeicoli(1, i)   ' Targa Motrice
        mobjOutFileVeicoli.Item("3").value = varExportDataVeicoli(2, i)   ' Targa Rimorchio
        mobjOutFileVeicoli.Item("4").value = varExportDataVeicoli(3, i)   ' Deposito
        mobjOutFileVeicoli.Item("5").value = varExportDataVeicoli(4, i)   ' Ente Emittente
        mobjOutFileVeicoli.Item("6").value = varExportDataVeicoli(5, i)   ' Descrizione
        mobjOutFileVeicoli.Item("7").value = varExportDataVeicoli(6, i)   ' Peso in Kg
        mobjOutFileVeicoli.Item("8").value = varExportDataVeicoli(7, i)   ' Volume Totale
        If varExportDataVeicoli(2, i) <> vbNullString Then
            mobjOutFileVeicoli.Item("9").value = varExportDataVeicoli(9, i)    ' C'è il rimorchio
        Else
            mobjOutFileVeicoli.Item("9").value = varExportDataVeicoli(8, i)    ' C'è la Motrice
        End If
        For j = 10 To 18
              mobjOutFileVeicoli.Item(CStr(j)).value = varExportDataVeicoli(j, i)
        Next j
        For j = 19 To intCols
            If (j Mod 2) > 0 Then
                If varExportDataVeicoli(j, i) > 0 Then
                    mobjOutFileVeicoli.Item(CStr(j)).value = varExportDataVeicoli(j, i) / 1000
                Else
                    mobjOutFileVeicoli.Item(CStr(j)).value = 0
                End If
            Else
                mobjOutFileVeicoli.Item(CStr(j)).value = varExportDataVeicoli(j, i)
            End If
        Next j
        If Not mobjOutFileVeicoli.WriteLine Then
         
        End If
    Next i
    
    MsgBox "Esportazione veicoli terminata correttamente", vbOKOnly + vbInformation, "Elaborazione terminata"
    ExportDataVeicoli = True
    GoTo EndProc
    
ErrHandler:

    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case RowsNotFound
            strMessage = "Errore durante la lettura delle righe di Unita di Trasporto"
            strInfo = "Nessuna Unità di Trasporto Trasferibile"
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti: Ente Emittente"
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & mstrExportFileName
        Case OutputError
            strMessage = "Si è verificato un errore durante la scrittura del file"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover1", _
                "ExportData", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    ExportDataVeicoli = False
    
EndProc:

    mobjOutFileVeicoli.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
End Function
Private Sub MaskFileVeicoli_GotFocus()
    blFocusOnVeicoli = True
End Sub

Private Sub TDBGrid1_SelChange(Cancel As Integer)
    lblLog(5) = TDBGrid1.SelBookmarks.Count
    lblLog(5).Refresh
End Sub

