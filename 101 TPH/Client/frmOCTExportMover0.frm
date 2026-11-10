VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTExportMover0 
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
      TabIndex        =   33
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
         Picture         =   "frmOCTExportMover0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Ordine"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportMover0.frx":0102
         TabIndex        =   20
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmOCTExportMover0.frx":0204
         TabIndex        =   19
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   1815
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox chkGC 
         Caption         =   "Grandes Cuentas"
         Height          =   255
         Left            =   7440
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox ChOrdiniAcq 
         Caption         =   "Solo Ord. Acquisiti"
         Height          =   255
         Left            =   7440
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmOCTExportMover0.frx":2C08
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frmOCTExportMover0.frx":2D0A
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
         Left            =   720
         TabIndex        =   15
         Top             =   1440
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Tag             =   "0"
         Top             =   360
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
         Left            =   5640
         TabIndex        =   9
         Tag             =   "20"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPresidio 
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         Tag             =   "1"
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskFileVeicoli 
         Height          =   255
         Left            =   5640
         TabIndex        =   44
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "File Veicoli:"
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   45
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblField0 
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
         Left            =   6240
         TabIndex        =   43
         Tag             =   "1"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   8760
         Tag             =   "1"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Presidio:"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   6960
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &N. Ordine"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   6960
         Tag             =   "20"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N. &Ordine"
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1440
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
      TabIndex        =   25
      Top             =   3480
      Width           =   3855
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   36
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini esclusi:"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   35
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini trasferiti correttamente:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini da trasferire:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   23
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
         Picture         =   "frmOCTExportMover0.frx":2E0C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmOCTExportMover0.frx":2F0E
         TabIndex        =   39
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
         Picture         =   "frmOCTExportMover0.frx":3010
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
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
         Picture         =   "frmOCTExportMover0.frx":3112
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmOCTExportMover0.frx":3214
         TabIndex        =   41
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Ordini"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportMover0.frx":3316
         TabIndex        =   42
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
         Picture         =   "frmOCTExportMover0.frx":3418
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmOCTExportMover0.frx":351A
         TabIndex        =   21
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
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
            TextSave        =   "16:25"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16/04/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmOCTExportMover0.frx":361C
      TabIndex        =   18
      Tag             =   "SELECT"
      Top             =   1920
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
            Picture         =   "frmOCTExportMover0.frx":10FEE
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportMover0.frx":11100
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportMover0.frx":11212
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTExportMover0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 31 '11
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
Private AZO As CZona
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCR As COrdineriga
Private AUT As CUnitaDiTrasporto



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
    Dim strPresidio As String
    Dim strErrorMessage As String
    
    DoEvents
    
    strPresidio = mskPresidio.Text
    If strPresidio = vbNullString Then
        ExportDataVeicoli = True
        Exit Function
    End If
    
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    
   
    blnResult = AUT.GetExportDataVeicoli(strPresidio, varExportDataVeicoli)
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
        If varExportDataVeicoli(2, i) <> vbNullString Then
            mobjOutFileVeicoli.Item("9").value = varExportDataVeicoli(9, i)    ' C'è il rimorchio
            varExportDataVeicoli(6, i) = varExportDataVeicoli(6, i) - varExportDataVeicoli(59, i) - varExportDataVeicoli(60, i)
            varExportDataVeicoli(15, i) = varExportDataVeicoli(15, i)
            varExportDataVeicoli(17, i) = varExportDataVeicoli(17, i) - varExportDataVeicoli(59, i) - varExportDataVeicoli(60, i)
        Else
            mobjOutFileVeicoli.Item("9").value = varExportDataVeicoli(8, i)    ' C'è SOLO la Motrice
            varExportDataVeicoli(6, i) = varExportDataVeicoli(6, i) - varExportDataVeicoli(59, i)
            varExportDataVeicoli(15, i) = varExportDataVeicoli(15, i) - varExportDataVeicoli(59, i)
        End If
        mobjOutFileVeicoli.Item("7").value = varExportDataVeicoli(6, i)   ' Peso in Kg
        mobjOutFileVeicoli.Item("8").value = varExportDataVeicoli(7, i)   ' Volume Totale
'        mobjOutFileVeicoli.Item("10").value = varExportDataVeicoli(0, i)   ' Codice Contratto (UDT)

        If varExportDataVeicoli(61, i) = "A01" Or varExportDataVeicoli(61, i) = "A02" Then
            If varExportDataVeicoli(62, i) = "A01" Or varExportDataVeicoli(62, i) = "A02" Then
                mobjOutFileVeicoli.Item("10").value = "001"
            Else
                mobjOutFileVeicoli.Item("10").value = "002"
            End If
        Else
            mobjOutFileVeicoli.Item("10").value = "003"
        End If
        
'        For j = 10 To 18
        For j = 11 To 18
              mobjOutFileVeicoli.Item(CStr(j)).value = varExportDataVeicoli(j, i)
        Next j
'        mobjOutFileVeicoli.Item("15").value = varExportDataVeicoli(15, i)   ' Peso MOTRICE in Kg
'        mobjOutFileVeicoli.Item("16").value = varExportDataVeicoli(16, i)   ' Volume MOTRICE
'        mobjOutFileVeicoli.Item("17").value = varExportDataVeicoli(17, i)   ' Peso RIMORCHIO in Kg
'        mobjOutFileVeicoli.Item("18").value = varExportDataVeicoli(18, i)   ' Volume RIMORCHIO
        
        For j = 19 To intCols - 4 '2
            If (j Mod 2) > 0 Then
                If varExportDataVeicoli(j, i) > 0 Then
                    mobjOutFileVeicoli.Item(CStr(j)).value = varExportDataVeicoli(j, i)
                Else
                    mobjOutFileVeicoli.Item(CStr(j)).value = 0
                End If
            Else
                mobjOutFileVeicoli.Item(CStr(j)).value = varExportDataVeicoli(j, i)
            End If
        Next j
        If Not mobjOutFileVeicoli.WriteLine Then
            Debug.Print "a"
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
            strInfo = "Nome del file: " & MaskFileVeicoli
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
    Set C = ErrH.SetError(ErrCode, "OCTExportMover1", _
                "ExportData", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    ExportDataVeicoli = False
    
EndProc:

    mobjOutFileVeicoli.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
End Function


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
            ExportRecords2
            ExportDataVeicoli
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
    Dim j As Long
    Dim intIndexProdotto As Integer
    Dim strIDOrdine As String
    Dim strIDProdotto As String
    Dim strIDParam As String
    Dim strIDZona As String
    Dim strNota3 As String
    Dim strRowState As String
    Dim strFileName As String
    Dim strCondTecnico As String
    Dim strNoteAccordiCliente As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim intRowError As Errors
    Dim varProdotti(1 To 10) As String
    Dim blnFlagProdottiOK As Boolean
    Dim blnFound As Boolean
    Dim blnGC As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
'    With AXX.DS
'        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
'        If Not AXX.Find Then
'            GoTo ErrHandler
'        End If
'        If .HowMany = 0 Then
'            ErrCode = InvalidHeader
'            GoTo ErrHandler
'        End If
'        For i = 1 To 10
'            If Trim$(AXX.DS.Fields(i + 5).value) <> vbNullString Then
'                blnFlagProdottiOK = True
'                varProdotti(i) = Trim$(AXX.DS.Fields(i + 5).value)
'            End If
'        Next i
'        If Not blnFlagProdottiOK Then
'            ErrCode = MissingDefaults
'            GoTo ErrHandler
'        End If
'    End With
    ' ---
    
    blnGC = (chkGC.value = vbChecked)
    strFileName = mskFileName
    If Not mobjOutFile.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    mobjOutFile.Delimiter = "|"
    
'    strHeader = "1|    |        |        |    | |"
'    Print #mobjOutFile.FileNum, strHeader
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione degli ordini
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        
        With mobjOutFile
            .ClearFields
            .Item("1").value = "2"
            .Item("2").value = GridArraySelect.value(lngRow, 11)
            If Not IsNull(GridArraySelect.value(lngRow, 28)) Then
                If GridArraySelect.value(lngRow, 28) = "2" Then
                    .Item("3").value = "00"
                Else
                    .Item("3").value = Right(GridArraySelect.value(lngRow, 16), 2)
                End If
            Else
                .Item("3").value = Right(GridArraySelect.value(lngRow, 16), 2)
            End If
            .Item("31").value = GridArraySelect.value(lngRow, 17)
            .Item("24").value = GridArraySelect.value(lngRow, 18)
            .Item("4").value = UCase(GridArraySelect.value(lngRow, 0))
            .Item("5").value = GridArraySelect.value(lngRow, 1)
            .Item("6").value = GridArraySelect.value(lngRow, 2)
            .Item("7").value = GridArraySelect.value(lngRow, 3)
            .Item("8").value = GridArraySelect.value(lngRow, 4)
            .Item("9").value = GridArraySelect.value(lngRow, 5)
            .Item("10").value = GridArraySelect.value(lngRow, 6)
            If GridArraySelect.value(lngRow, 16) = ID_SOC_CEPSA Then
                If IsNull(GridArraySelect.value(lngRow, 25)) Then
                    strCondTecnico = vbNullString
                Else
                    strCondTecnico = DecodeCondTecnico(GridArraySelect.value(lngRow, 25))
                End If
                Select Case UCase(GridArraySelect.value(lngRow, 26))
                    Case "R"
                        .Item("11").value = "ES"
                    Case "E", "O", "EA"
                        .Item("11").value = "CD"
                    Case "B", "BM"
                        .Item("11").value = "BA"
                        .Item("5").value = GridArraySelect.value(lngRow, 2)
                End Select
                If Len(strCondTecnico) > 0 Then
                    .Item("11").value = .Item("11").value & "/ " & strCondTecnico
                End If
            Else
                .Item("11").value = GridArraySelect.value(lngRow, 7)
            End If

            ' --- 11/08/99
            ' --- La riga seguente è stata aggiunta per
            ' --- mandare a mover i numeri ordine nello stesso
            ' --- formato del numero ordine Agip (zznnnnnn)
            .Item("12").value = Format(GridArraySelect.value(lngRow, 8), String(8, "0"))
'            .Item("12").value = GridArraySelect.value(lngRow, 11) & _
'                                Right$(GridArraySelect.value(lngRow, 8), 6)
            ' ---
                                
            .Item("13").value = GridArraySelect.value(lngRow, 9)
            If mskField9.Text = ID_SOC_AGIP Then
                If IsDate(GridArraySelect.value(lngRow, 27)) And GridArraySelect.value(lngRow, 27) <> CDate("01/01/1900") Then
                    .Item("16").value = Format$(GridArraySelect.value(lngRow, 27), "dd/mm/yy")
                Else
                    .Item("16").value = Format$(GridArraySelect.value(lngRow, 10), "dd/mm/yy")
                End If
            Else
                .Item("16").value = Format$(GridArraySelect.value(lngRow, 10), "dd/mm/yy")
            End If
                
            If Format(GridArraySelect.value(lngRow, 10), "dd/mm/yy") = Format(mskDateTo, "dd/mm/yy") Then
                .Item("15").value = "0000"
            Else
                .Item("15").value = "0900"
            End If
            
            If blnGC Then
                .Item("19").value = "GC"
            Else
                .Item("19").value = "00"
            End If
            
            .Item("20").value = GridArraySelect.value(lngRow, 12)
            .Item("21").value = GridArraySelect.value(lngRow, 13)
            .Item("22").value = GridArraySelect.value(lngRow, 14)
            
            strNota3 = vbNullString
            If Len(GridArraySelect.value(lngRow, 30)) > 0 Then
                strNota3 = strNota3 & "B.: " & GridArraySelect.value(lngRow, 30)
            End If
            Select Case GridArraySelect.value(lngRow, 28)
                Case "1"
                    strNota3 = strNota3 & " Imagen CEPSA"
                Case "2"
                    strNota3 = strNota3 & " Imagen BLANCA"
            End Select
            Select Case GridArraySelect.value(lngRow, 29)
                Case "1"
                    strNota3 = strNota3 & " Obl. precincto"
            End Select
            
            If GridArraySelect.value(lngRow, 16) = ID_SOC_CEPSA Then
                strNoteAccordiCliente = GridArraySelect.value(lngRow, 31) & " " & _
                                        GridArraySelect.value(lngRow, 32) & " " & _
                                        Format(GridArraySelect.value(lngRow, 33), "hhnn") & " " & _
                                        Format(GridArraySelect.value(lngRow, 34), "hhnn") & " " & _
                                        Format(GridArraySelect.value(lngRow, 35), "hhnn") & " " & _
                                        Format(GridArraySelect.value(lngRow, 36), "hhnn") & " "
                strNota3 = strNota3 & strNoteAccordiCliente
            End If
            
            
            .Item("23").value = strNota3 'GridArraySelect.value(lngRow, 15)
            
            .Item("25").value = Format(GridArraySelect.value(lngRow, 19), "hhnn")
            .Item("26").value = Format(GridArraySelect.value(lngRow, 20), "hhnn")
            .Item("27").value = Format(GridArraySelect.value(lngRow, 21), "hhnn")
            .Item("28").value = Format(GridArraySelect.value(lngRow, 22), "hhnn")
            .Item("29").value = Format(GridArraySelect.value(lngRow, 23), "hhnn")
            .Item("30").value = Format(GridArraySelect.value(lngRow, 24), "hhnn")
            If .Item("25").value = "0000" Then .Item("25").value = Space(4)
            If .Item("26").value = "0000" Then .Item("26").value = Space(4)
            If .Item("27").value = "0000" Then .Item("27").value = Space(4)
            If .Item("28").value = "0000" Then .Item("28").value = Space(4)
            If .Item("29").value = "0000" Then .Item("29").value = Space(4)
            If .Item("30").value = "0000" Then .Item("30").value = Space(4)
            
            ' Condizionanti tecnici
            If Len(Right(GridArraySelect.value(lngRow, 7), 5)) = 0 Then
                Select Case UCase(GridArraySelect.value(lngRow, 26))
                    Case "O", "EA"
                        .Item("32").value = "30000"
                    Case Else
                        .Item("32").value = "00000"
                End Select
            Else
                .Item("32").value = Right(GridArraySelect.value(lngRow, 7), 5)
                Select Case UCase(GridArraySelect.value(lngRow, 26))
                    Case "O", "EA"
                        If Left(.Item("32").value, 1) = "0" Then
                            .Item("32").value = "3" & Right(.Item("32").value, 4)
                        End If
                End Select
            End If
            
            ' Tipo servizio
            If Len(Mid(GridArraySelect.value(lngRow, 7), 7, 6)) = 0 Then
                Select Case UCase(GridArraySelect.value(lngRow, 26))
                    Case "R"
                        .Item("33").value = "910103"
                    Case "E"
                        .Item("33").value = "910101"
                    Case "O"
                        .Item("33").value = "910153"
                    Case "EA"
                        .Item("33").value = "910152"
                    Case "B"
                        .Item("33").value = "920101"
                    Case "BM"
                        .Item("33").value = "920105"
                    Case "P"
                        .Item("33").value = "910115"
                End Select
            Else
                .Item("33").value = Mid(GridArraySelect.value(lngRow, 7), 7, 6)
            End If
        End With
        
        strIDOrdine = GridArraySelect.value(lngRow, 8)
        OCR.DS.Fields(OCR.DS.KeyFields(1)).FindValue = strIDOrdine
        
        If Not OCR.GetList(True) Then
            intRowError = RowsNotFound
            GoTo LogResult
        End If
        intRowsOrder = OCR.DS.HowMany - 1
        
        If intRowsOrder > 0 Then
            For i = 1 To intRowsOrder
                For j = 0 To (i - 1)
                    If OCR.DS.ListItem(1, i) = OCR.DS.ListItem(1, j) Then
                        intRowError = BadIDOrderSoc
                        GoTo LogResult
                    End If
                Next j
            Next i
        End If
        
        ' --- Controllo che la zona non sia da escludere
        ' --- dal passaggio a Mover
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
                Dim rowmod As Variant
                Dim hmmod As Long
                
                
                 If DataClass.DS.DA.QueryCommand("SELECT OCRRAGG1 FROM ORDIOCR0 WHERE OCRIDOCT = '" + CStr(strIDOrdine) + "' AND OCRIDAPR='" + CStr(strIDProdotto) + "'", rowmod, hmmod) Then
                 
                    If rowmod(0, 0) = "NO" Then
                    .Item("14").value = "NO" + String$(28, "0")
                    Else
                    .Item("14").value = "SI" + String$(28, "0")
                    End If
                    
                    
                 End If
                 
                
                    
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
            Case BadIDOrderSoc
                strErrorMessage = "L'ordine contiene prodotti duplicati"
        End Select
        XA.value(lngRow, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
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
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
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
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
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
            If Not DataClass.GetSelectionExport(ChOrdiniAcq.value = vbChecked, chkGC.value = vbChecked, mskPresidio) Then
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

Private Function DecodeCondTecnico(CondTecnico As String) As String

    On Error GoTo ErrHandler
    
    Dim strTemp As String
    Dim lngCT As Long
    Dim lngTest As Long
    
    If Not IsNumeric(CondTecnico) Then
        GoTo ErrHandler
    End If
    lngCT = CLng(CondTecnico)
    
    ' Tipo veicolo
    lngTest = (lngCT \ 10000)
    Select Case lngTest
        Case 1
            strTemp = "TV 9TM"
        Case 2
            strTemp = "TV 14TM"
        Case 3
            strTemp = "TV 19TM"
        Case 4
            strTemp = "TV -24TM"
        Case 5
            strTemp = "TV +24TM"
    End Select
    
    lngCT = lngCT - lngTest * 10000
    lngTest = (lngCT \ 1000)
    ' Lunghezza tubo
    Select Case lngTest
        Case 1
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "LM 8"
        Case 2
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "LM 15"
        Case 3
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "LM 30"
        Case 4
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "LM +30"
    End Select
    
    lngCT = lngCT - lngTest * 1000
    lngTest = (lngCT \ 100)
   ' Lunghezza tubo
    Select Case lngTest
        Case 1
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "RG APL"
    End Select
    
    ' Tipo scarico
    lngCT = lngCT - lngTest * 100
    lngTest = (lngCT \ 10)
    Select Case lngTest
        Case 1
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "ED BOMBA"
        Case 2
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "ED BOMBA/CONT"
        Case 3
            If Len(strTemp) > 0 Then
                strTemp = strTemp & " - "
            End If
            strTemp = strTemp & "ED TICKET"
    End Select
    
    DecodeCondTecnico = strTemp
    Exit Function
ErrHandler:
    DecodeCondTecnico = vbNullString
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



Private Sub chkGC_Click()
'    mskPresidio.Enabled = (chkGC.value = vbUnchecked)
End Sub

Private Sub ChOrdiniAcq_Click()
    If ChOrdiniAcq Then    ' Propongo il nome file in base al flag ordini acquisiti
        mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile1", , True)
    Else
        mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile", , True)
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
    Set XA = New XArrayDB
    
    mskFileName = gmRegistry.GetSetting("Options", "OrderExportFile", , True)
    MaskFileVeicoli = gmRegistry.GetSetting("Options", "AUTExportFile", , True)
    If IsUserGrandesCuentas Then
        chkGC.value = vbChecked
    Else
        chkGC = vbChecked
    End If
    mskPresidio = DataConnect.ConnUserName
    D = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
'    mskDateFrom.SelText = D
    mskDateTo.SelText = D
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New COrdineSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set Soc = New CSocieta
    Set Soc.DS.DA = DataConnect.DA
    SetValueItems TDBGrid1.Columns(1), Soc
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineriga
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
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
        With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    
    Set mobjOutFile = New CTextFile
    With mobjOutFile
        .Add "Tipo Record", 1, "1"
        .Add "Zona", 2, "2"
        .Add "Società", 2, "3"
        .Add "Codice Destinazione", 10, "4"
        .Add "Ragione Sociale", 28, "5"
        .Add "Indirizzo", 65, "6"
        .Add "Località", 65, "7"
        .Add "CAP", 5, "8"
        .Add "Telefono", 20, "9"
        .Add "Note 1", 50, "10"
        .Add "Note 2", 50, "11"
        .Add "Numero Ordine", 8, "12"
        .Add "Deposito", 3, "13"
        .Add "Descrizione", 30, "14"
        .Add "Filler 2", 4, "15"
        .Add "DataConsegna", 8, "16"
        .Add "Prodotto", 6, "17"
        .Add "Quantità", 6, "18"
        .Add "Tipo Vendita", 2, "19"
        .Add "Ordine Programmato", 1, "20"
        .Add "Ordine Urgente", 1, "21"
        .Add "Fascia di Consegna", 1, "22"
        .Add "Note 3", 50, "23"
        .Add "KM Raggio", 4, "24", "", fftNumerico
        .Add "Dalle ore 1", 4, "25", "", fftNumerico
        .Add "Alle ore 1", 4, "26", "", fftNumerico
        .Add "Dalle ore 2", 4, "27", "", fftNumerico
        .Add "Alle ore 2", 4, "28", "", fftNumerico
        .Add "Dalle ore 3", 4, "29", "", fftNumerico
        .Add "Alle ore 3", 4, "30", "", fftNumerico
        .Add "Ragione sociale società", 30, "31"
        .Add "Condizionanti tecnici", 5, "32", , fftNumerico
        .Add "Tipo Servizio", 6, "33", , fftNumerico
    End With
    Set mobjOutFileVeicoli = New CTextFile
    With mobjOutFileVeicoli
        .Add "Unità di trasporto", 3, "1"
        .Add "Targa Motrice", 10, "2"
        .Add "Targa Rimorchio", 10, "3"
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
        .Add "Capacita Scomparto 1 Motrice", 5, "19", 0, fftNumerico
        .Add "Numero Scomparti 1 Motrice", 2, "20", 0, fftNumerico
        .Add "Capacita Scomparto 2 Motrice", 5, "21", 0, fftNumerico
        .Add "Numero Scomparti 2 Motrice", 2, "22", 0, fftNumerico
        .Add "Capacita Scomparto 3 Motrice", 5, "23", 0, fftNumerico
        .Add "Numero Scomparti 3 Motrice", 2, "24", 0, fftNumerico
        .Add "Capacita Scomparto 4 Motrice", 5, "25", 0, fftNumerico
        .Add "Numero Scomparti 4 Motrice", 2, "26", 0, fftNumerico
        .Add "Capacita Scomparto 5 Motrice", 5, "27", 0, fftNumerico
        .Add "Numero Scomparti 5 Motrice", 2, "28", 0, fftNumerico
        .Add "Capacita Scomparto 6 Motrice", 5, "29", 0, fftNumerico
        .Add "Numero Scomparti 6 Motrice", 2, "30", 0, fftNumerico
        .Add "Capacita Scomparto 7 Motrice", 5, "31", 0, fftNumerico
        .Add "Numero Scomparti 7 Motrice", 2, "32", 0, fftNumerico
        .Add "Capacita Scomparto 8 Motrice", 5, "33", 0, fftNumerico
        .Add "Numero Scomparti 8 Motrice", 2, "34", 0, fftNumerico
        .Add "Capacita Scomparto 9 Motrice", 5, "35", 0, fftNumerico
        .Add "Numero Scomparti 9 Motrice", 2, "36", 0, fftNumerico
        .Add "Capacita Scomparto 10 Motrice", 5, "37", 0, fftNumerico
        .Add "Numero Scomparti 10 Motrice", 2, "38", 0, fftNumerico
        .Add "Capacita Scomparto 1 Rimorchio", 5, "39", 0, fftNumerico
        .Add "Numero Scomparti 1 Rimorchio", 2, "40", 0, fftNumerico
        .Add "Capacita Scomparto 2 Rimorchio", 5, "41", 0, fftNumerico
        .Add "Numero Scomparti 2 Rimorchio", 2, "42", 0, fftNumerico
        .Add "Capacita Scomparto 3 Rimorchio", 5, "43", 0, fftNumerico
        .Add "Numero Scomparti 3 Rimorchio", 2, "44", 0, fftNumerico
        .Add "Capacita Scomparto 4 Rimorchio", 5, "45", 0, fftNumerico
        .Add "Numero Scomparti 4 Rimorchio", 2, "46", 0, fftNumerico
        .Add "Capacita Scomparto 5 Rimorchio", 5, "47", 0, fftNumerico
        .Add "Numero Scomparti 5 Rimorchio", 2, "48", 0, fftNumerico
        .Add "Capacita Scomparto 6 Rimorchio", 5, "49", 0, fftNumerico
        .Add "Numero Scomparti 6 Rimorchio", 2, "50", 0, fftNumerico
        .Add "Capacita Scomparto 7 Rimorchio", 5, "51", 0, fftNumerico
        .Add "Numero Scomparti 7 Rimorchio", 2, "52", 0, fftNumerico
        .Add "Capacita Scomparto 8 Rimorchio", 5, "53", 0, fftNumerico
        .Add "Numero Scomparti 8 Rimorchio", 2, "54", 0, fftNumerico
        .Add "Capacita Scomparto 9 Rimorchio", 5, "55", 0, fftNumerico
        .Add "Numero Scomparti 9 Rimorchio", 2, "56", 0, fftNumerico
        .Add "Capacita Scomparto 10 Rimorchio", 5, "57", 0, fftNumerico
        .Add "Numero Scomparti 10 Rimorchio", 2, "58", 0, fftNumerico
    End With
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If ChOrdiniAcq Then
        gmRegistry.SaveSetting "Options", "OrderExportFile1", mskFileName, True
    Else
        gmRegistry.SaveSetting "Options", "OrderExportFile", mskFileName, True
    End If
    gmRegistry.SaveSetting "Options", "OrderExportGC", chkGC.value, True
    gmRegistry.SaveSetting "Options", "AUTExportFile", MaskFileVeicoli, True
    Set XA = Nothing
    Set GridArraySelect = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set OCR = Nothing
    Set AZO = Nothing
    Set AUT = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set mobjOutFile = Nothing
    Set mobjOutFileVeicoli = Nothing
    Set frmOCTExportMover0 = Nothing
End Sub


Private Sub mskFileName_Change()
    cmdExport.Enabled = (Len(mskFileName) > 0)
End Sub


