VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmOCTExportMover 
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
      Height          =   1935
      Left            =   0
      TabIndex        =   27
      Top             =   3480
      Width           =   5295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmOCTExportMover.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Ordine"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportMover.frx":0102
         TabIndex        =   14
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid50.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmOCTExportMover.frx":0204
         TabIndex        =   13
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   1575
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmOCTExportMover.frx":1C21
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frmOCTExportMover.frx":1D23
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
         _Version        =   327681
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
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1080
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   5880
         TabIndex        =   7
         Tag             =   "0"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Tag             =   "20"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   7200
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &N. Ordine"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7200
         Tag             =   "20"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N. &Ordine"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File:"
         Height          =   255
         Index           =   1
         Left            =   1440
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
      Height          =   1935
      Left            =   5400
      TabIndex        =   19
      Top             =   3480
      Width           =   3855
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   30
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini esclusi:"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   29
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini trasferiti correttamente:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini da trasferire:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   17
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
         Picture         =   "frmOCTExportMover.frx":1E25
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmOCTExportMover.frx":1F27
         TabIndex        =   33
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
         Picture         =   "frmOCTExportMover.frx":2029
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
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
         Picture         =   "frmOCTExportMover.frx":212B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmOCTExportMover.frx":222D
         TabIndex        =   35
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Ordini"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportMover.frx":232F
         TabIndex        =   36
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
         Picture         =   "frmOCTExportMover.frx":2431
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmOCTExportMover.frx":2533
         TabIndex        =   15
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   6465
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
            TextSave        =   "18.45"
            Key             =   "Time"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16/11/99"
            Key             =   "Date"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmOCTExportMover.frx":2635
      TabIndex        =   12
      Tag             =   "SELECT"
      Top             =   1800
      Width           =   9015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   4440
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
            Picture         =   "frmOCTExportMover.frx":6898
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOCTExportMover.frx":69AA
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOCTExportMover.frx":6ABC
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTExportMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 12 '11
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mobjOutFile As CTextFile
'Private mobjOutFile As CTextFile

Public GridArraySelect As ArrayGridCls
Public XA As XArray
Public Loading As Boolean

Public DataClass As COrdineSelezione
Public State As FormState

Private Soc As CSocieta
Private AZO As CZona
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCR As COrdineRiga



Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
'            With DataClass.DS
'                .Fields("22").FindValue = "A"
'            End With
            If SelectRecords Then
                cmdOpenFile.Enabled = True
                If Len(mskFileName) = 0 Then
                    mskFileName = App.Path & "\ORD_ASC.TXT"
                End If
            Else
                cmdOpenFile.Enabled = False
            End If
        Case fcExport
            Select Case gmstrCodInstallazione
                Case ciInst4
                    ExportRecords
                Case Else
                    ExportRecords2
            End Select
            cmdExport.Enabled = False
        Case fcOpen
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per Mover... "
                .InitDir = App.Path
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                .ShowSave
                mskFileName = .fileName
            End With
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcORDIOCT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(tdbLog)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

' --- Esportazione ordini in formato AGIP
Private Sub ExportRecords2()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngQta As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim intRowsOrder As Integer
    Dim i As Long
    Dim intIndexProdotto As Integer
    Dim strIDOrdine As String
    Dim strIDProdotto As String
    Dim strIDParam As String
    Dim strIDZona As String
    Dim strRowState As String
    Dim strFileName As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim intRowError As Errors
    Dim varProdotti(1 To 10) As String
    Dim blnFlagProdottiOK As Boolean
    Dim blnFound As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    
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
        For i = 1 To 10
            If Trim$(AXX.DS.Fields(i + 5).value) <> vbNullString Then
                blnFlagProdottiOK = True
                varProdotti(i) = Trim$(AXX.DS.Fields(i + 5).value)
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
    mobjOutFile.Delimiter = "|"
    
    strHeader = "1|    |        |        |    | |"
    Print #mobjOutFile.FileNum, strHeader
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione degli ordini
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        
        With mobjOutFile
            .ClearFields
            .Item("1").value = "2"
            .Item("2").value = GridArraySelect.value(lngRow, 11)
            .Item("3").value = "00"
            .Item("4").value = GridArraySelect.value(lngRow, 0)
            .Item("5").value = GridArraySelect.value(lngRow, 1)
            .Item("6").value = GridArraySelect.value(lngRow, 2)
            .Item("7").value = GridArraySelect.value(lngRow, 3)
            .Item("8").value = GridArraySelect.value(lngRow, 4)
            .Item("9").value = GridArraySelect.value(lngRow, 5)
            .Item("10").value = GridArraySelect.value(lngRow, 6)
            .Item("11").value = GridArraySelect.value(lngRow, 7)

            ' --- 11/08/99
            ' --- La riga seguente è stata aggiunta per
            ' --- mandare a mover i numeri ordine nello stesso
            ' --- formato del numero ordine Agip (zznnnnnn)
'            .Item("12").value = GridArraySelect.value(lngRow, 8)
            .Item("12").value = GridArraySelect.value(lngRow, 11) & _
                                Right$(GridArraySelect.value(lngRow, 8), 6)
            ' ---
                                
            .Item("13").value = GridArraySelect.value(lngRow, 9)
            .Item("15").value = "0000"
            .Item("16").value = Format$(GridArraySelect.value(lngRow, 10), "dd/mm/yy")
            .Item("19").value = "00"
        End With
        
        strIDOrdine = GridArraySelect.value(lngRow, 8)
        OCR.DS.Fields(OCR.DS.KeyFields(1)).FindValue = strIDOrdine
        
        If Not OCR.GetList(True) Then
            intRowError = RowsNotFound
            GoTo LogResult
        End If
        
        ' --- Controllo che la zona non sia da escludere
        ' --- dall'importazione
        strIDZona = GridArraySelect.value(lngRow, 11)
        AZO.DS.Fields("Codice").FindValue = strIDZona
        If Not AZO.Find Then
            intRowError = NoZona
            GoTo LogResult
        End If
        If AZO.DS.HowMany = 0 Then
            intRowError = NoZona
            GoTo LogResult
        End If
        If AZO.DS.Fields("3").value <> 0 Then
            intRowError = ZonaEsclusaExport
            GoTo LogResult
        End If
        ' ---
        
        intRowsOrder = OCR.DS.HowMany - 1
        For i = 0 To intRowsOrder
            strRowState = OCR.DS.ListItem(9, i)
            If strRowState = "A" Then
                strIDProdotto = OCR.DS.ListItem(1, i)
                If IsEmpty(OCR.DS.ListItem(4, i)) Then
                    intRowError = MissingQuantity
                    GoTo LogResult
                End If
                lngQta = CLng(OCR.DS.ListItem(4, i))
                If lngQta = 0 Then
                    intRowError = MissingQuantity
                    GoTo LogResult
                End If
                intIndexProdotto = 1
                With mobjOutFile
                    .Item("14").value = String$(30, "0")
                    .Item("17").value = strIDProdotto
                    .Item("18").value = Format$(lngQta, "000000")
                    If Not .WriteLine(True) Then
                        intRowError = OutputError
                        GoTo LogResult
                    End If
                 End With
           End If
        Next i
        
        
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
        lblLog(7) = lngCountExcluded
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
        End Select
        XA.value(lngRow, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = lngRow
        End With
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
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFile.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub

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
    Dim lngQta As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim intRowsOrder As Integer
    Dim i As Long
    Dim intIndexProdotto As Integer
    Dim strIDOrdine As String
    Dim strIDProdotto As String
    Dim strIDParam As String
    Dim strIDZona As String
    Dim strRowState As String
    Dim strFileName As String
    Dim strErrorMessage As String
    Dim intRowError As Errors
    Dim varProdotti(1 To 10) As String
    Dim blnFlagProdottiOK As Boolean
    Dim blnFound As Boolean
    Dim ErrCode As Errors
    Dim C As CError

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
        For i = 1 To 10
            If Trim$(AXX.DS.Fields(i + 5).value) <> vbNullString Then
                blnFlagProdottiOK = True
                varProdotti(i) = Trim$(AXX.DS.Fields(i + 5).value)
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
    lngCountToGo = lngNumRows + 1
    ' --- Ciclo di elaborazione degli ordini
    For lngRow = 0 To lngNumRows
        intRowError = NoError

        With mobjOutFile
            .ClearFields
            .Item("1").value = GridArraySelect.value(lngRow, 0)
            .Item("2").value = GridArraySelect.value(lngRow, 1)
            .Item("3").value = GridArraySelect.value(lngRow, 2)
            .Item("4").value = GridArraySelect.value(lngRow, 3)
            .Item("5").value = GridArraySelect.value(lngRow, 4)
            .Item("6").value = GridArraySelect.value(lngRow, 5)
            .Item("7").value = GridArraySelect.value(lngRow, 6)
            .Item("8").value = GridArraySelect.value(lngRow, 7)
            .Item("9").value = GridArraySelect.value(lngRow, 8)
            .Item("10").value = GridArraySelect.value(lngRow, 9)
            .Item("11").value = GridArraySelect.value(lngRow, 10)
        End With

        strIDOrdine = GridArraySelect.value(lngRow, 8)
        OCR.DS.Fields(OCR.DS.KeyFields(1)).FindValue = strIDOrdine

        If Not OCR.GetList(True) Then
            intRowError = RowsNotFound
            GoTo LogResult
        End If

        ' --- Controllo che la zona non sia da escludere
        ' --- dall'importazione
        strIDZona = GridArraySelect.value(lngRow, 11)
        AZO.DS.Fields("Codice").FindValue = strIDZona
        If Not AZO.Find Then
            intRowError = NoZona
            GoTo LogResult
        End If
        If AZO.DS.HowMany = 0 Then
            intRowError = NoZona
            GoTo LogResult
        End If
        If AZO.DS.Fields("3").value <> 0 Then
            intRowError = ZonaEsclusaExport
            GoTo LogResult
        End If
        ' ---

        intRowsOrder = OCR.DS.HowMany - 1
        For i = 0 To intRowsOrder
            strRowState = OCR.DS.ListItem(9, i)
            If strRowState = "A" Then
                strIDProdotto = OCR.DS.ListItem(1, i)
                If IsEmpty(OCR.DS.ListItem(4, i)) Then
                    intRowError = MissingQuantity
                    GoTo LogResult
                End If
                lngQta = CLng(OCR.DS.ListItem(4, i))
                If lngQta = 0 Then
                    intRowError = MissingQuantity
                    GoTo LogResult
                End If
                intIndexProdotto = 1
                blnFound = False
                Do While intIndexProdotto <= 10 And Not blnFound
                    If strIDProdotto = varProdotti(intIndexProdotto) Then
                        blnFound = True
                    Else
                        intIndexProdotto = intIndexProdotto + 1
                    End If
                Loop
                If Not blnFound Then
                    intRowError = NoProductInParam
                    GoTo LogResult
                End If
                mobjOutFile.Item(CStr(11 + intIndexProdotto)).value = lngQta
            End If
        Next i

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
        lblLog(7) = lngCountExcluded
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
        End Select
        XA.value(lngRow, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = lngRow
        End With
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
    
    Me.MousePointer = vbHourglass
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            If Not DataClass.GetSelectionExport() Then
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
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
    lblLog(7) = 0
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
    Set XA = New XArray
    
    mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile", , True)
    
    D = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
'    mskDateFrom.SelText = D
    mskDateTo.SelText = D
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New COrdineSelezione
    With DataClass.DS
        Set .DA = DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set Soc = New CSocieta
    Set Soc.DS.DA = DA
    SetValueItems TDBGrid1.Columns(1), Soc
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineRiga
    With OCR.DS
        Set .DA = DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DA
        Set .ErrH = ErrH
    End With
    
    
    Select Case gmstrCodInstallazione
        Case ciInst4
            Set mobjOutFile = New CTextFile
            With mobjOutFile
                .Add "Codice Destinazione", 8, "1"
                .Add "Codice Cliente", 10, "2"
                .Add "Ragione Sociale", 30, "3"
                .Add "Indirizzo", 30, "4"
                .Add "Località", 20, "5"
                .Add "Provincia", 2, "6"
                .Add "CAP", 5, "7"
                .Add "Telefono", 11, "8"
                .Add "Cap. Prodotto 1", 3, "9"
                .Add "Cap. Prodotto 2", 3, "10"
                .Add "Cap. Prodotto 3", 3, "11"
                .Add "Cap. Prodotto 4", 3, "12"
                .Add "Ordine", 7, "13"
                .Add "Tipo Ordine", 2, "14"
                .Add "Deposito", 3, "15"
                .Add "Ente Emittente", 3, "16"
                .Add "Data Consegna", 8, "17"
                .Add "Note", 30, "18"
                .Add "Qta Iniziale Prodotto 1", 2, "19"
                .Add "Qta Iniziale Prodotto 2", 2, "20"
                .Add "Qta Iniziale Prodotto 3", 2, "21"
                .Add "Qta Iniziale Prodotto 4", 2, "22"
                .Add "Qta Concordata Prodotto 1", 2, "23"
                .Add "Qta Concordata Prodotto 2", 2, "24"
                .Add "Qta Concordata Prodotto 3", 2, "25"
                .Add "Qta Concordata Prodotto 4", 2, "26"
            End With
        Case Else
            Set mobjOutFile = New CTextFile
            With mobjOutFile
                .Add "Tipo Record", 1, "1"
                .Add "Zona", 2, "2"
                .Add "Filler1", 2, "3"
                .Add "Codice Destinazione", 8, "4"
                .Add "Ragione Sociale", 30, "5"
                .Add "Indirizzo", 65, "6"
                .Add "Località", 65, "7"
                .Add "CAP", 5, "8"
                .Add "Telefono", 20, "9"
                .Add "Note 1", 50, "10"
                .Add "Note 2", 50, "11"
                .Add "Numero Ordine", 8, "12"
                .Add "Deposito", 2, "13"
                .Add "Descrizione", 30, "14"
                .Add "Filler 2", 4, "15"
                .Add "DataConsegna", 8, "16"
                .Add "Prodotto", 6, "17"
                .Add "Quantità", 6, "18"
                .Add "Tipo Vendita", 2, "19"
            End With
    End Select
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "OrderExportFile", mskFileName, True
    Set XA = Nothing
    Set GridArraySelect = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set OCR = Nothing
    Set AZO = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set mobjOutFile = Nothing
'    Set mobjOutFile2 = Nothing
    Set frmOCTExportMover = Nothing
End Sub


Private Sub mskFileName_Change()
    cmdExport.Enabled = (Len(mskFileName) > 0)
End Sub


