VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmGP1ExportCEPSA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CEPSA - TRASMISSIONE DISCREPANZE"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   8910
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   30
      Top             =   5640
      Width           =   8895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   4560
         Picture         =   "frmGP1ExportCEPSA.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmGP1ExportCEPSA.frx":0102
         TabIndex        =   11
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   2400
         Picture         =   "frmGP1ExportCEPSA.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Prefatture"
         Height          =   375
         Left            =   2280
         Picture         =   "frmGP1ExportCEPSA.frx":0306
         TabIndex        =   9
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
         Index           =   0
         Left            =   240
         Picture         =   "frmGP1ExportCEPSA.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Seleziona Percorso"
         Height          =   375
         Left            =   120
         Picture         =   "frmGP1ExportCEPSA.frx":050A
         TabIndex        =   36
         Tag             =   "OPEN"
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
         Left            =   7200
         Picture         =   "frmGP1ExportCEPSA.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7080
         Picture         =   "frmGP1ExportCEPSA.frx":070E
         TabIndex        =   12
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   4920
      TabIndex        =   23
      Top             =   3600
      Width           =   3975
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Records da trasmettere:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   27
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Records trasmessi correttamente:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   2175
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8895
      Begin VB.CheckBox chkPedaggi 
         Caption         =   "Solo pedaggi discordanti"
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   480
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkImporti 
         Caption         =   "Solo importi discordanti"
         Height          =   255
         Left            =   6720
         TabIndex        =   44
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cmbAnno 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbMese 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   8040
         Picture         =   "frmGP1ExportCEPSA.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   8040
         Picture         =   "frmGP1ExportCEPSA.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   6720
         TabIndex        =   0
         Tag             =   "5"
         Top             =   1080
         Visible         =   0   'False
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
         Left            =   6720
         TabIndex        =   1
         Tag             =   "6"
         Top             =   1440
         Visible         =   0   'False
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
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Tag             =   "14"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Tag             =   "15"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   1320
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
               Picture         =   "frmGP1ExportCEPSA.frx":0A14
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1ExportCEPSA.frx":0B26
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1ExportCEPSA.frx":0C38
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskSoc 
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Tag             =   "20"
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskIDAEE 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Tag             =   "21"
         Top             =   1800
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
         Index           =   2
         Left            =   5640
         Tag             =   "21"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emittente:"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblsocieta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   46
         Tag             =   "21"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Anno"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   42
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mese"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5640
         Tag             =   "20"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblsocieta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   38
         Tag             =   "20"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Società:"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio"
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   6360
         Tag             =   "15"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal Num. &Viaggio"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3360
         Tag             =   "14"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data Inizio Viaggio"
         Height          =   255
         Index           =   12
         Left            =   4800
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   8520
         Tag             =   "5"
         Top             =   1080
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data Inizio Viaggio"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   8520
         Tag             =   "6"
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Percorso File:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   0
      TabIndex        =   20
      Top             =   3600
      Width           =   4815
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmGP1ExportCEPSA.frx":0D4A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Dettaglio"
         Height          =   375
         Left            =   120
         Picture         =   "frmGP1ExportCEPSA.frx":0E4C
         TabIndex        =   10
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmGP1ExportCEPSA.frx":0F4E
         TabIndex        =   8
         Top             =   240
         Width           =   4575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   6495
      Width           =   8910
      _ExtentX        =   15716
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
            Object.Width           =   10266
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "13:54"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/12/2017"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1335
      Left            =   120
      OleObjectBlob   =   "frmGP1ExportCEPSA.frx":3F4A
      TabIndex        =   43
      Top             =   2280
      Width           =   8655
   End
End
Attribute VB_Name = "frmGP1ExportCEPSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const MAX_COLS = 25

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjOutFileTrip As CTextFile

Private mstrProgrInvio As String

Public GridArraySelect As ArrayGridCls
Public XA As XArrayDB
Public Loading As Boolean

Public DataClass As CPrefatturaDettaglioSelezione
Public State As FormState

Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private AUT As CUnitaDiTrasporto
Private AAU As CAutista
Private AAM As CAutomezzo
Private OCT As COrdine
Private ADZ As CDestinazione
Private ATV As CTipoVendita
Private ATM As CTipoAutomezzo
Private AXX As CParametro
Private ADE As CDeposito
Private AEE As CEnteEmittente
Private AVE As CVettore
Private ACL As CCliente
Private APR As CProdotto
Private APD As CPosDoganale
Private AZO As CZona
Private ASC As CSocieta
Private GP1 As CPrefatturaSoc

Private Sub DisplaySelection()
    Dim varLog As Variant
    Dim i As Long
    Dim lngRows As Long
    
    lngRows = DataClass.DS.HowMany
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
'    lblLog(7) = 0
    If lngRows > 0 Then
        ReDim varLog(0 To 2, 0 To lngRows - 1)
        For i = 0 To lngRows - 1
            varLog(0, i) = DataClass.DS.ListItem(0, i)
            varLog(1, i) = DataClass.DS.ListItem(1, i)
        Next i
        RefreshXArray XA, tdbLog, varLog
    End If
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
                cmdExport.Enabled = True
            End If
        Case fcExport
            ExportRecords001
            cmdExport.Enabled = False
        Case fcOpen
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per la trasmissione... "
                .InitDir = App.Path
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.* "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                .ShowSave
                mskFileName = .fileName
            End With
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcVIAGVVT)
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
Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
    End If
End Function

Private Sub ExportRecords001()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim i As Long
    Dim j As Long
    Dim intRowsTrip As Integer
    Dim intProgressivoRecord As Integer
    Dim intProgressivoConsegna As Integer
    Dim intCountSospesi As Integer
    Dim strCodOperatore As String
    Dim strIDZona As String
    Dim strIDZonaPrec As String
    Dim strIDViaggio As String
    Dim strIDViaggioPrec As String
    Dim strIDEnteEmittente As String
    Dim strIDDeposito As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strDataViaggio As String
    Dim strDataInvio As String
    Dim strOraInvio As String
    Dim varBadge As Variant
    Dim strCentroOperativo As String
    Dim strUnitàTrasporto As String
    Dim strOraInizioViaggio As String
    Dim strFileName As String
    Dim strOraConsegna As String
    Dim strRifAGIP As String
    Dim strStatoOrdine As String
    Dim strTipoVendita As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim strTail As String
    Dim varNIF As Variant
    Dim strDummy As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim strKeyNumeratore As String
    Dim strDescNumeratore As String
    Dim ErrCode As Errors
    Dim C As CError
    
    Dim StrTipoRecord As String
    Dim strNota2_1 As String
    Dim strnota2_2 As String
    Dim strcodentexsoc As String
    Dim strNumOrdineCepsa As String
    Dim strNumPartizione As String
    Dim strFiller5 As String
    Dim strRIFEST1 As String
    Dim strCODDESTXSOC As String
    Dim strCODPRODXSOC As String
    Dim varQTAINIZiOoRDINE As Variant
    Dim strUMPRODOTTO As String
    Dim strDATACONSRIChiestaORD As String
    Dim strpriorita As String
    Dim strCODDEPXSOC As String
    Dim strCodAUTSoc As String
    Dim strCodAUTDep As String
    Dim strCodAAUSoc As String
    Dim strIDTipoOrdine As String
    Dim strMotrice As String
    Dim strRimorchio As String
    Dim varNOTA2_3 As Variant
    Dim varRifEst2 As Variant
    Dim varfiller16 As Variant
    Dim strProdotto As String
    Dim varVIAGQTATRASP As Variant
    Dim strDTAVIAGGIO As String
    Dim varNSEQ As Variant
    Dim varOraConsegna As Variant
    Dim varPedaggio As Variant
    Dim varPedaggioEsterno As Variant
    Dim varCounterPedaggio As Variant
    Dim varAnomalia As Variant
    Dim strNIFAutista As String
    Dim strORACONSPREV As String
    Dim strRifVettore As String
    Dim intDifferenza As Integer
    Dim intNumeroDestinatari As Integer
    Dim intCounterDestinatari As Integer
    Dim strProgressivo As String
    Dim blnDiffImporti As Boolean
    Dim blnDiffPedaggi As Boolean
    Dim Deposito As String
    
    Dim ANU As CNumeratore
    
    Dim strIDSocietàDeposito As String
    
    Dim strPcName As String
    Dim strComName As String
    Dim strTempFileName As String
    Dim blnErrori As Boolean
    Dim blnTrasmetti As Boolean
    Dim intFileFromNum As Integer
    Dim intFileToNum As Integer
    Dim strLine As String
    
    
    If MaskSoc = "" Then
        Call MsgBox("E' necessario indicare la società.", vbExclamation + vbOKOnly)
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    
    varCounter = 0
    Select Case mskIDAEE
        Case "004"
            strKeyNumeratore = Right(cmbAnno.Text, 2) & Chr(VBA.ASC("A") + cmbMese.ListIndex)
            strDescNumeratore = "INVIO DISCR. CEPSA EUT"
            strRifVettore = "709704" '"709704" 'strRifVettore = "002881" Oilogist
        Case "624"
            strKeyNumeratore = Right(cmbAnno.Text, 2) & Chr(VBA.ASC("M") + cmbMese.ListIndex)
            strDescNumeratore = "INVIO DISCR. CEPSA TPH"
            strRifVettore = "054320"
        Case Else
            GoTo ErrHandler
    End Select
    
    
    '-- DP 13/06/08
    ' Con mail del 12/06/08 R. Chaparro richiede che in caso di errori il file possa essere generato su richiesta
    ' Creo dunque un file temporaneo che verrà copiato e rinominato nel percorso corretto solamente se non ci sono
    ' errori oppure se l'utente decide comunque di generarlo con errori
    If GetComputerName(strPcName) Then
        strComName = UCase(strPcName)
    End If
    strTempFileName = strComName & "_DISCREPANCIAS_" & mskIDAEE.Text & ".tmp"
    If Not mobjOutFileTrip.OpenFile(strTempFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione dei viaggi
    For lngRow = 0 To lngNumRows
        
        intRowError = NoError
        
        If (chkImporti.value = vbChecked) Then
            blnDiffImporti = False
            varAnomalia = GridArraySelect.value(lngRow, 28)
            blnDiffImporti = ((varAnomalia And apDiffImporto) = apDiffImporto) Or _
                             ((varAnomalia And apTariffa) = apTariffa) Or _
                             ((varAnomalia And apPedaggio) = apPedaggio) Or _
                             ((varAnomalia And apServOperarioBunker) = apServOperarioBunker) Or _
                             ((varAnomalia And apServBuquesPuerto) = apServBuquesPuerto) Or _
                             ((varAnomalia And apServOreServSpeciali) = apServOreServSpeciali) Or _
                             ((varAnomalia And apServImpServSpeciali) = apServImpServSpeciali) Or _
                             ((varAnomalia And apServImpServEsenti) = apServImpServEsenti)
                             
            If Not blnDiffImporti Then
                If (chkPedaggi.value = vbChecked) Then
                    blnDiffPedaggi = False
                    varAnomalia = GridArraySelect.value(lngRow, 28)
                    blnDiffPedaggi = ((varAnomalia And apPedaggio) = apPedaggio)
                    If Not blnDiffPedaggi Then
                        intRowError = OrderAlreadyExisting
                        GoTo LogResult
                    End If
                Else
                    intRowError = OrderAlreadyExisting
                    GoTo LogResult
                End If
            End If
        Else
            If (chkPedaggi.value = vbChecked) Then
                blnDiffPedaggi = False
                varAnomalia = GridArraySelect.value(lngRow, 28)
                blnDiffPedaggi = ((varAnomalia And apPedaggio) = apPedaggio)
                If Not blnDiffPedaggi Then
                    intRowError = OrderAlreadyExisting
                    GoTo LogResult
                End If
            End If
        End If
        
        
        If UCase(GridArraySelect.value(lngRow, 24)) = "D" Then
            intRowError = OrderNotActive
            GoTo LogResult
        End If
        
        With ADE.DS
            strDummy = GridArraySelect.value(lngRow, 14)
            .Fields("0").FindValue = strDummy
            If Not ADE.Find Then
                intRowError = NoDeposito
                GoTo LogResult
            Else
                strCODDEPXSOC = .MultiSoc.DS.GetFirmCode(ID_SOC_CEPSA)
                strIDSocietàDeposito = .Field("Società").value
            End If
        End With
         
        With ADZ.DS
            .Fields(.KeyFields(1)).FindValue = GridArraySelect.value(lngRow, 13)
            If Not ADZ.Find Then
                ErrCode = NoDestinazione
'                GoTo ErrHandler

                ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                strCODDESTXSOC = vbNullString
            Else
                If .HowMany = 0 Then
                    ErrCode = NoDestinazione
'                    GoTo ErrHandler

                    ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                    strCODDESTXSOC = vbNullString
                Else
                    strCODDESTXSOC = UCase(.MultiSoc.DS.GetFirmCode(ID_SOC_CLH))
                    If (UCase(Left(strCODDESTXSOC, 2)) = "CP") Or (UCase(Left(strCODDESTXSOC, 2)) = "EA") Or (UCase(Left(strCODDESTXSOC, 2)) = "OT") Then
                        strCODDESTXSOC = Mid(strCODDESTXSOC, 3)
                    End If
                End If
            End If
        End With
         
        With APR.DS
            .Fields(.KeyFields(1)).FindValue = GridArraySelect.value(lngRow, 9)
            If Not APR.Find Then
                ErrCode = NoProdotto
'                GoTo ErrHandler
                
                ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                strProdotto = vbNullString
            Else
                If .HowMany = 0 Then
                    ErrCode = NoProdotto
'                    GoTo ErrHandler
                
                    ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                    strProdotto = vbNullString
                Else
                    strProdotto = .MultiSoc.DS.GetFirmCode(MaskSoc, 3)
                End If
            End If
        End With
        
        With AUT.DS
            strDummy = GridArraySelect.value(lngRow, 7)
            .Fields("0").FindValue = strDummy
            If Not AUT.Find Then
                intRowError = noUT
'                GoTo LogResult

                ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                strCodAUTSoc = vbNullString
            Else
                strCodAUTSoc = UCase(.MultiSoc.DS.GetFirmCode(ID_SOC_CEPSA))
                If IsEmptyVariantString(strCodAUTSoc) Then
                    intRowError = InvalidAddress
'                    GoTo LogResult

                    ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                    strCodAUTSoc = vbNullString
                Else
                    If Len(strCodAUTSoc) > 0 Then
                        If Left(strCodAUTSoc, 1) <> "D" Then
                            strCodAUTSoc = "D" & String(9 - Len(strCodAUTSoc), "0") & strCodAUTSoc
                        End If
                    End If
                    
                    '-- DP 04/06/08
                    '-- Su richiesta di A. Mangi del 04/06/08 occorre prendere il codice dell'AUT in base alla società del deposito e non con CLH fisso
                    'strCodAUTDep = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                    strCodAUTDep = .MultiSoc.DS.GetFirmCode(strIDSocietàDeposito)
                    If IsEmptyVariantString(strCodAUTDep) Then
                        intRowError = InvalidDep
                        
                        ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
'                        GoTo LogResult
                    End If
                End If
            End If
        End With
        
        With AAU.DS
            strDummy = GridArraySelect.value(lngRow, 18)
            .Fields("0").FindValue = strDummy
            If Not AAU.Find Then
                intRowError = NoAutista
'                GoTo LogResult

                ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
                strCodAAUSoc = vbNullString
                strNIFAutista = vbNullString
            Else
                '-- DP 04/06/08
                '-- Su richiesta di A. Mangi del 04/06/08 occorre prendere il codice dell'AAU solo se la società del deposito è CLH, altrimenti questo campo va lascianto in bianco
                'strCodAAUSoc = .MultiSoc.DS.GetFirmCode(ID_SOC_CEPSA)
                If strIDSocietàDeposito = ID_SOC_CLH Then
                    strCodAAUSoc = .MultiSoc.DS.GetFirmCode(strIDSocietàDeposito)
                Else
                    strCodAAUSoc = vbNullString
                End If
                
                strNIFAutista = UCase(.Fields("9").value)
                If Len(strNIFAutista) > 0 Then
                    If (Left(strNIFAutista, 2) <> "ES") And IsNumeric(Left(strNIFAutista, 2)) Then
                        strNIFAutista = "ES" & strNIFAutista
                    End If
                    strNIFAutista = Replace(Replace(strNIFAutista, " ", ""), "-", "")
                End If
            End If
        End With
        
        AAM.DS.Fields.ResetFields
        AAM.DS.Field("Targa").FindValue = GridArraySelect.value(lngRow, 16)
        If AAM.Find Then
            If AAM.DS.HowMany > 0 Then
                strMotrice = UCase(GridArraySelect.value(lngRow, 16))
                If Left(strMotrice, 2) <> "ES" Then
                    If Len(AAM.DS.Field("Flag Nazionalità").value) > 0 Then
                        strMotrice = UCase(AAM.DS.Field("Flag Nazionalità").value) & strMotrice
                    Else
                        strMotrice = "ES" & strMotrice
                    End If
                End If
                strMotrice = Replace(Replace(strMotrice, " ", ""), "-", "")
            Else
                '-- DP 16/06/08
                '-- Aggiunto controllo su motrice
                intRowError = NoMotrice
            End If
        Else
            '-- DP 16/06/08
            '-- Aggiunto controllo su motrice
            intRowError = NoMotrice
        End If
        
        strRimorchio = vbNullString
        If Len(GridArraySelect.value(lngRow, 17)) > 0 Then
            AAM.DS.Fields.ResetFields
            AAM.DS.Field("Targa").FindValue = GridArraySelect.value(lngRow, 17)
            If AAM.Find Then
                If AAM.DS.HowMany > 0 Then
                    strRimorchio = UCase(GridArraySelect.value(lngRow, 17))
                    If Len(strRimorchio) > 0 Then
                        If Left(strRimorchio, 2) <> "ES" Then
                            If Len(AAM.DS.Field("Flag Nazionalità").value) > 0 Then
                                strRimorchio = UCase(AAM.DS.Field("Flag Nazionalità").value) & strRimorchio
                            Else
                                strRimorchio = "ES" & strRimorchio
                            End If
                        End If
                        strRimorchio = Replace(Replace(strRimorchio, " ", ""), "-", "")
                    End If
                Else
                    '-- DP 16/06/08
                    '-- Aggiunto controllo su rimorchio
                    intRowError = NoRimorchio
                End If
            Else
                '-- DP 16/06/08
                '-- Aggiunto controllo su rimorchio
                intRowError = NoRimorchio
            End If
        End If
        If Len(strRimorchio) = 0 Then
            strRimorchio = strMotrice
        End If
        
        strIDTipoOrdine = GridArraySelect.value(lngRow, 35)
        
        ' ---
        With mobjOutFileTrip
            .ClearFields
            .Item("1").value = strRifVettore
'            .Item("2").value = Format(GridArraySelect.value(lngRow, 6), "yyyymmdd")     ' Data consegna
            ' *** GR 03/09/2006 - utilizziamo la data di inizio carico invece che la data di inizio viaggio
            .Item("2").value = Format(GridArraySelect.value(lngRow, 26), "yyyymmdd")     ' Data consegna
            
'            If ((varAnomalia And apDiffImporto) = apDiffImporto) Then
'                If GridArraySelect.value(lngRow, 33) = "C" Then
'                    .Item("3").value = "PLIMIT"
'                End If
'            End If
                
            If GridArraySelect.value(lngRow, 33) = "C" Then
                 .Item("3").value = "PLIMIT"
            End If
            
            ' Jg 20/4/2016 Ceuta
            
            Deposito = GridArraySelect.value(lngRow, 14)
            If Deposito = "260" Then
                .Item("3").value = "SCEUTA"
            End If
            
            .Item("6").value = strCODDEPXSOC                                            ' Deposito
            .Item("13").value = GridArraySelect.value(lngRow, 19)                       ' Sequenza viaggio
            
            
            ' Buscamos los datos en prefactura Cepsa
            GP1.DS.Fields("2").FindValue = GridArraySelect.value(lngRow, 22)
            If GP1.Find Then
                If GP1.DS.HowMany > 0 Then
                    .Item("2").value = Format(GP1.DS.Fields("3").value, "yyyymmdd")     ' Fecha de viaje recibida de Cepsa
                    .Item("14").value = GP1.DS.Fields("29").value                        'Cod. Operador CEPSA
                    .Item("15").value = GP1.DS.Fields("30").value                        'Cod. Suminitrador CEPSA
                    .Item("17").value = GP1.DS.Fields("32").value                       ' Rif. Consegna SAP
                    .Item("18").value = GP1.DS.Fields("33").value                       ' Ticket
                    .Item("7").value = GP1.DS.Fields("22").value                        ' UDT Cepsa
                    .Item("8").value = GP1.DS.Fields("23").value                        ' UDT por deposito
                    .Item("9").value = GP1.DS.Fields("24").value                        ' MAtricula CEPSA tractora
                    .Item("10").value = GP1.DS.Fields("25").value                       ' MAtricula CEPSA cisterna
                    .Item("11").value = GP1.DS.Fields("26").value                       ' Codigo CEPSA conductor
                    .Item("12").value = GP1.DS.Fields("27").value                       ' NIF Conductor
                    With APR.DS
                        .Fields(.KeyFields(1)).FindValue = GP1.DS.Fields("9").value
                        If Not APR.Find Then
                            ErrCode = NoProdotto
                           
                        Else
                            If .HowMany = 0 Then
                                ErrCode = NoProdotto
                            Else
                                strProdotto = .MultiSoc.DS.GetFirmCode(MaskSoc, 3)
                            End If
                        End If
                    End With
                    .Item("19").value = strProdotto & "0000"
                
                Else
                    .Item("14").value = Left(GridArraySelect.value(lngRow, 21), 6)              ' Cod. Operador RESOIL
                    .Item("15").value = Mid(GridArraySelect.value(lngRow, 21), 7, 6)            ' Cod. Suminstrador RESOIL
            
                    ' Patch : c...ata di CEPSA (per EA viene inviato il codice suministrador sbagliato, dobbiamo inviare quello giusto)
                    If .Item("15").value = "910151" Then
                        .Item("15").value = "910152"
                    End If
                    ' Patch : gestione ordini "OTECLIMA" ricevuti in manuale
                    If (strIDTipoOrdine = "O") And (IsNull(.Item("15").value) Or (Len(.Item("15").value) = 0)) Then
                        .Item("14").value = "910101"
                        .Item("15").value = "910153"
                    End If
                    '-- DP 18/07/08 se non trovo il numero di albaran in PREFGP10 utilizzo OCRNRTCK
                    If GridArraySelect.value(lngRow, 36) <> vbNullString Then
                        .Item("17").value = vbNullString
                        .Item("18").value = Format(GridArraySelect.value(lngRow, 36), "00000000000000") & "0"
                    Else
                        '-- DP 10/06/08
                        '-- Se la Find non trova nulla e OCRNRTCK è vuoto segnalo l'errore come richiesto da R. Chaparro con mail del 12/06/08
                        If GridArraySelect.value(lngRow, 14) <> "260" Then intRowError = RowsNotFound
        
                        ' -- DP 18/07/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
        '                GoTo LogResult
                    End If
                    .Item("19").value = strProdotto & "0000"
                    .Item("7").value = strCodAUTSoc                                             ' UdT per Società
                    .Item("8").value = strCodAUTDep                                             ' UdT per Deposito
                    .Item("9").value = strMotrice                                               ' Motrice
                    .Item("10").value = strRimorchio                                            ' Rimorchio
                    .Item("11").value = Right(strCodAAUSoc, 7)                                  ' Autista per Società
                    .Item("12").value = strNIFAutista                                           ' Documento Autista

                End If
            End If
            .Item("16").value = GridArraySelect.value(lngRow, 22)                       ' Riferimento ordine
            ' --- Eseguo la ricerca del record di prefatturazione
            ' --- esterno per Codice Ordine esterno
            GP1.DS.Fields("2").FindValue = GridArraySelect.value(lngRow, 22)
            
    ' Julio 25/11/11 Desplazo arriba en busqueda de tabla Cepsa.
            
'            If GP1.Find Then
'                If GP1.DS.HowMany > 0 Then
'                    .Item("17").value = GP1.DS.Fields("32").value                       ' Rif. Consegna SAP
'                    .Item("18").value = GP1.DS.Fields("33").value                       ' Ticket
'                Else
'                    '-- DP 18/07/08 se non trovo il numero di albaran in PREFGP10 utilizzo OCRNRTCK
'                    If GridArraySelect.value(lngRow, 36) <> vbNullString Then
'                        .Item("17").value = vbNullString
'                        .Item("18").value = Format(GridArraySelect.value(lngRow, 36), "00000000000000") & "0"
'                    Else
'                        '-- DP 10/06/08
'                        '-- Se la Find non trova nulla e OCRNRTCK è vuoto segnalo l'errore come richiesto da R. Chaparro con mail del 12/06/08
'                        intRowError = RowsNotFound
'
'                        ' -- DP 18/07/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
'        '                GoTo LogResult
'                    End If
'                End If
'            Else
'                '-- DP 18/07/08 se non trovo il numero di albaran in PREFGP10 utilizzo OCRNRTCK
'                If GridArraySelect.value(lngRow, 36) <> vbNullString Then
'                    .Item("17").value = vbNullString
'                    .Item("18").value = Format(GridArraySelect.value(lngRow, 36), "00000000000000") & "0"
'                Else
'                    '-- DP 10/06/08
'                    '-- Se la Find non trova nulla e OCRNRTCK è vuoto segnalo l'errore come richiesto da R. Chaparro con mail del 12/06/08
'                    intRowError = RowsNotFound
'
'                    ' -- DP 13/06/08 - R. Chaparro chiede che la riga possa essere trasmesso anche con errori
'    '                GoTo LogResult
'                End If
'            End If
            
'            .Item("19").value = strProdotto & "0000"                                    ' Codice Prodotto
            
            Select Case .Item("15").value
                Case "910152"
                    .Item("20").value = "00000XG" & "G" & Mid(strCODDESTXSOC, 2)            ' Codice Destinazione
                Case "910153"
                    .Item("20").value = "00000XG" & "O" & Mid(strCODDESTXSOC, 2)            ' Codice Destinazione
                Case Else
                    .Item("20").value = "00000DP" & strCODDESTXSOC                          ' Codice Destinazione
            End Select
            .Item("21").value = GridArraySelect.value(lngRow, 34)                       ' Nome imbarcazione
            .Item("22").value = GridArraySelect.value(lngRow, 10)                       ' Litri scaricati
            .Item("23").value = GridArraySelect.value(lngRow, 20)                       ' Kili  scaricati
            .Item("24").value = GridArraySelect.value(lngRow, 15)                       ' Kilometri
            
            If GridArraySelect.value(lngRow, 38) = "LM" Then                            ' Franja tarifa 2011
                .Item("25").value = "LM"
                Else
                .Item("25").value = "BT"
            End If

            .Item("26").value = GridArraySelect.value(lngRow, 11) * 1000000             ' Tariffa
            .Item("27").value = GridArraySelect.value(lngRow, 12) * 100                 ' Importo
            
            .Item("30").value = GridArraySelect.value(lngRow, 29) * 100
            .Item("31").value = GridArraySelect.value(lngRow, 30) * 100
            .Item("32").value = GridArraySelect.value(lngRow, 31) * 100
            .Item("33").value = GridArraySelect.value(lngRow, 32) * 100
            
            '-- DP 20/08/08
            .Item("34").value = IIf(Not IsNull(GridArraySelect.value(lngRow, 37)), GridArraySelect.value(lngRow, 37) * 1000, vbNullString)                 ' Servizi Esenti
            
'            If (varAnomalia And apPedaggio) = apPedaggio Then
'                .Item("28").value = GridArraySelect.value(lngRow, 25) * 100             ' Pedaggio
'            Else
'                .Item("28").value = 0
'            End If
            
            strIDViaggio = GridArraySelect.value(lngRow, 5) ' Numero viaggio
            If strIDViaggio <> strIDViaggioPrec Then
                strIDViaggioPrec = strIDViaggio
                With VVT
                    .DS.Fields("0").FindValue = strIDViaggio
                    If Not .Find Then
                        intRowError = NoTripID
                        GoTo LogResult
                    End If
                End With
                intNumeroDestinatari = VVT.ContaDestinatari
                intCounterDestinatari = 0
                varCounterPedaggio = 0
            End If
            
            If IsNumeric(GridArraySelect.value(lngRow, 25)) Then
                If GridArraySelect.value(lngRow, 25) > 0 Then
                    varPedaggio = Round((GridArraySelect.value(lngRow, 25)), 2) * 100
                    If GP1.DS.HowMany >= 0 Then
                        varPedaggioEsterno = GP1.GetPedaggioCEPSA(GridArraySelect.value(lngRow, 22))
                        
                         'JG-- 11/1/12 "Parece" que se controla a parte los peajes y si el peaje tiene una diferencia inferior a 0.05 pasa de todo y salta.
                         
'                        If (Not IsEmpty(varPedaggioEsterno)) And IsNumeric(varPedaggioEsterno) Then
'                            If Abs(varPedaggio - varPedaggioEsterno * 100) <= 0.05 Then
'                                intRowError = OrderAlreadyExisting
'                                GoTo LogResult
'                            End If
'                        End If
                    
                    End If
                    
'                    intCounterDestinatari = intCounterDestinatari + 1
                    ' --- Sull'ultima consegna impostiamo la quota di pedaggio rimanente
'                    If intCounterDestinatari = intNumeroDestinatari Then
'                        .Item("28").value = Round(GridArraySelect.value(lngRow, 25), 2) * 100 - varCounterPedaggio
'                    Else
'                        .Item("28").value = varPedaggio
'                        varCounterPedaggio = varCounterPedaggio + varPedaggio
'                    End If
                    .Item("28").value = varPedaggio
                End If
            End If
'            If IsNumeric(GridArraySelect.value(lngRow, 25)) Then
'                If GridArraySelect.value(lngRow, 25) > 0 Then
'                    varPedaggio = Round((GridArraySelect.value(lngRow, 25) / intNumeroDestinatari), 2) * 100
'                    If GP1.DS.HowMany >= 0 Then
'                        varPedaggioEsterno = GP1.GetPedaggioCEPSA(GridArraySelect.value(lngRow, 22))
'
'                        ' -- DP 09/01/09
'                        ' -- La prima versione della funzione usciva se i pedaggi discrepavano di meno di 2 euro.
'                        ' -- Ora, essendoci anche altre discrepanze possibili, è stato modificato il controllo.
'                        blnDiffImporti = ((varAnomalia And apDiffImporto) = apDiffImporto) Or _
'                             ((varAnomalia And apTariffa) = apTariffa) Or _
'                             ((varAnomalia And apServOperarioBunker) = apServOperarioBunker) Or _
'                             ((varAnomalia And apServBuquesPuerto) = apServBuquesPuerto) Or _
'                             ((varAnomalia And apServOreServSpeciali) = apServOreServSpeciali) Or _
'                             ((varAnomalia And apServImpServSpeciali) = apServImpServSpeciali) Or _
'                             ((varAnomalia And apServImpServEsenti) = apServImpServEsenti)
'
'                        If (Not IsEmpty(varPedaggioEsterno)) And IsNumeric(varPedaggioEsterno) Then
'                            If Abs(varPedaggio - varPedaggioEsterno * 100) <= 1 And Not blnDiffImporti Then
'                                intRowError = OrderAlreadyExisting
'                                GoTo LogResult
'                            End If
'                        End If
'                    End If
'
'
'                    If (Not IsEmpty(varPedaggioEsterno)) And IsNumeric(varPedaggioEsterno) Then
'                        If Abs(varPedaggio - varPedaggioEsterno * 100) <= 1 Then
'                            intCounterDestinatari = intCounterDestinatari + 1
'                            ' --- Sull'ultima consegna impostiamo la quota di pedaggio rimanente
'                            If intCounterDestinatari = intNumeroDestinatari Then
'                                .Item("28").value = Round(GridArraySelect.value(lngRow, 25), 2) * 100 - varCounterPedaggio
'                            Else
'                                .Item("28").value = varPedaggio
'                                varCounterPedaggio = varCounterPedaggio + varPedaggio
'                            End If
'                        End If
'                    End If
'                End If
'            End If
          
            If Not .WriteLine Then
                 intRowError = OutputError
                 GoTo LogResult
            End If
        End With
        
        ' --- Modifico lo stato del viaggio
        If Not VVT.SetFlagTrasmissione("D") Then
            intRowError = NoSetState
            GoTo LogResult
        End If
        
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            If (intRowError = OrderAlreadyExisting) Or _
               (intRowError = OrderNotActive) Then
                lngCountExcluded = lngCountExcluded + 1
            Else
                blnErrori = True
                lngCountFail = lngCountFail + 1
            End If
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
'        lblLog(7) = lngCountExcluded
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case OutputError
                strErrorMessage = "Errore durante la scrittura sul file"
            Case noUT
                strErrorMessage = "Unità di trasporto non trovata"
            Case RowsNotFound
                strErrorMessage = "Falta el numero de albaran"
            Case NoSetState
                strErrorMessage = "Impossibile cambiare lo stato del viaggio"
            Case NoOrderID
                strErrorMessage = "Ordine non trovato"
            Case NoTipoVendita
                strErrorMessage = "Codice Tipo Vendita non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case InvalidCode
                strErrorMessage = "N.I.F. assente"
            Case InvalidDep
                strErrorMessage = "Badge automezzo assente"
            Case InvalidAddress
                strErrorMessage = "Codice CEPSA Unità di Trasporto assente"
            Case OrderNotActive
                strErrorMessage = "Trattamento in corso"
            Case OrderAlreadyExisting
                strErrorMessage = "Discrepanza da non trasmettere"
            Case NoTripID
                strErrorMessage = "Viaggio non trovato"
            Case NoAutista
                strErrorMessage = "Autista non trovato"
        End Select
        XA.value(lngRow, 2) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
        ' ---
        
        DoEvents
        
    Next lngRow
    ' ---
    
    mobjOutFileTrip.CloseFile
    
    
    '-- DP 13/06/08
    ' Con mail del 12/06/08 R. Chaparro richiede che in caso di errori il file possa essere generato su richiesta
    ' Creo dunque un file temporaneo che verrà copiato e rinominato nel percorso corretto solamente se non ci sono
    ' errori oppure se l'utente decide comunque di generarlo con errori
    If blnErrori Then
        If MsgBox("ATENCIÓN: Se ha producido errores durante la creación de la transmisión. ¿Se desea crear igualmente el fichero?", vbQuestion + vbYesNo, "ResOil.2") = vbYes Then
            blnTrasmetti = True
        Else
            blnTrasmetti = False
        End If
    Else
        blnTrasmetti = True
    End If
    
    If blnTrasmetti Then
    '    If Len(mstrProgrInvio) = 0 Then
            With DataClass.DS.DA
                varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", strKeyNumeratore, 0)
                Select Case varCounter
                    Case -1
                        Set ANU = New CNumeratore
                        With ANU.DS
                            Set .DA = DataClass.DS.DA
                            Set .ErrH = DataClass.DS.ErrH
                        End With
                        ANU.DS.Fields("0").NewValue = strKeyNumeratore
                        ANU.DS.Fields("0").Updated = True
                        ANU.DS.Fields("1").NewValue = strDescNumeratore
                        ANU.DS.Fields("1").Updated = True
                        ANU.DS.Fields("2").NewValue = "0000000"
                        ANU.DS.Fields("2").Updated = True
                        ANU.DS.Fields("3").NewValue = Format(Date, "dd/mm/yyyy")
                        ANU.DS.Fields("3").Updated = True
                        If Not ANU.Insert Then
                            GoTo ErrHandler
                        End If
                        varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", strKeyNumeratore, 0)
                    Case Is >= 99
                        varCounter = 1
                End Select
            End With
            mstrProgrInvio = Format$(varCounter, String(2, "0"))
    '    End If
    
        If Len(mskFileName) > 0 Then
            strFileName = mskFileName
        End If
    
        strFileName = strFileName & "\" & "TPREF" & strRifVettore & cmbAnno.Text & Format(cmbMese.ListIndex + 1, "00") & Right(mstrProgrInvio, 2) & ".txt"
    
        intFileFromNum = FreeFile
        Open strTempFileName For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Print #intFileToNum, strLine
        Loop
        Close #intFileFromNum
        Close #intFileToNum
    End If
        
    
    
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
    DataClass.DS.DA.EndTrans ETROLLBACK
    Set C = ErrH.SetError(ErrCode, "frmGP1ExportCEPSA", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFileTrip.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            DataClass.DS.Field("Deposito").Operator = " <> "
            DataClass.DS.Field("Deposito").FindValue = "011"
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

Private Sub SetDate()
    Dim strFromDate As String
    Dim strToDate As String
    
    If Len(cmbAnno) = 0 Then Exit Sub
    strFromDate = "01/" & cmbMese.ListIndex + 1 & "/" & cmbAnno
    If cmbMese.ListIndex = 11 Then
        strToDate = DateAdd("d", -1, "01/01/" & CStr(CInt(cmbAnno) + 1))
    Else
        strToDate = DateAdd("d", -1, "01/" & cmbMese.ListIndex + 2 & "/" & cmbAnno)
    End If
    mskDateFrom = Format(strFromDate, "dd/mm/yy")
    mskDateTo = Format(strToDate, "dd/mm/yy")
End Sub

Private Sub cmbAnno_Click()
    SetDate
End Sub


Private Sub cmbMese_Click()
    SetDate
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

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim D As Variant
    Dim i As Integer
    
    Loading = True
    Set GridArraySelect = New ArrayGridCls
    Set XA = New XArrayDB
    
    cmdBrowse.Enabled = True
    MaskSoc = gmRegistry.GetSetting("Options", "SocExport1", , True)
    mskIDAEE = gmRegistry.GetSetting("Options", "IDAEEExport1", , True)
    mskFileName = gmRegistry.GetSetting("Options", "GP1Export", , True)
        
    Set mobjOutFileTrip = New CTextFile
    With mobjOutFileTrip
        .Add "Vettore", 6, "1", ID_AVE_CEPSA, fftNumerico
        .Add "Data", 8, "2", , fftNumerico
        .Add "Tipo di servizio", 6, "3"
        .Add "Tipo consegna", 2, "4", " 0"
        .Add "Stato record", 1, "5" ', "D"
        .Add "Deposito", 4, "6", , fftNumerico
        .Add "Riferimento UDT", 12, "7"
        .Add "Riferimento UDT Deposito", 3, "8"
        .Add "Targa Motrice", 15, "9"
        .Add "Targa Rimorchio", 15, "10"
        .Add "Riferimento autista", 7, "11"
        .Add "NIF Autista", 14, "12"
        .Add "Numero viaggio", 5, "13", , fftNumerico
        .Add "Codice operatore", 6, "14"
        .Add "Codice fornitore", 6, "15"
        .Add "Riferimento ordine", 16, "16"
        .Add "Riferimento consegna SAP", 12, "17"
        .Add "Riferimento ticket", 15, "18"
        .Add "Codice Prodotto", 9, "19"
        .Add "Codice Destinazione", 15, "20"
        .Add "Nome imbarcazione", 30, "21"
        .Add "Quantità Litri", 7, "22", , fftNumerico
        .Add "Quantità Kili", 7, "23", , fftNumerico
        .Add "Distanza Km", 4, "24", , fftNumerico
        .Add "Tipo tariffa", 2, "25", "BT"
        .Add "Tariffa", 9, "26", 0, fftNumerico
        .Add "Importo", 10, "27", 0, fftNumerico
        .Add "Importo pedaggio", 10, "28", 0, fftNumerico
        .Add "Servicios Exentos", 9, "34", 0, fftNumerico
        .Add "Filler", 54, "29"
        .Add "Servizio operario bunker", 10, "30", 0, fftNumerico
        .Add "Servizio buques a puerto", 10, "31", 0, fftNumerico
        .Add "Ore servizi speciali", 5, "32", 0, fftNumerico
        .Add "Importo servizi speciali", 10, "33", 0, fftNumerico
    End With
    
    D = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
    mskDateFrom.SelText = D
    mskDateTo.SelText = D
        
    Set ErrH = New CErrorHandler
    Set DataClass = New CPrefatturaDettaglioSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .Fields("21").Required = True
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
'    With AXX.DS
'        .Fields(.KeyFields(1)).FindValue = ID_PARAM
'        If AXX.Find Then
'            MaskEdBox1 = AXX.DS.Fields("16").value
'            If DataClass.DS.ValidateClientField("4", MaskEdBox1) Then
'                FieldUpdateDescription "4", Me, MaskEdBox1
'                mskFileName = mskFileName & MaskEdBox1
'            End If
'        End If
'    End With
    
    Set VVC = New CViaggioConsegna
    With VVC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ATM = New CTipoAutomezzo
    With ATM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ATV = New CTipoVendita
    With ATV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
       
    Set AVE = New CVettore
    With AVE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.DA = .DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With

    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.DA = .DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AEE = New CEnteEmittente
    With AEE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.DA = .DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APD = New CPosDoganale
    With APD.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set GP1 = New CPrefatturaSoc
    With GP1.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ASC = New CSocieta
    With ASC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    lblsocieta(8).Caption = DescrizioneSocieta(MaskSoc)
    
    For i = 1 To 12
        cmbMese.AddItem Format("01/" & CStr(i) & "/2000", "mmmm")
    Next i
    cmbMese.ListIndex = Month(Now) - 1
    cmbAnno.AddItem CStr(Year(Date) - 1)
    cmbAnno.AddItem CStr(Year(Date))
    cmbAnno.AddItem CStr(Year(Date) + 1)
    cmbAnno.ListIndex = 1
    Loading = False
End Sub

Private Function DescrizioneSocieta(ByVal pIdSoc As String) As String
    Dim strDescrSocieta As String
    On Error GoTo ErrHandler
    
    strDescrSocieta = vbNullString
    If pIdSoc <> vbNullString Then
        With ASC.DS
            .Fields.ResetFields
            .Fields("0").FindValue = pIdSoc
            If ASC.Find And .HowMany > 0 Then
                strDescrSocieta = .Fields("1").value
            End If
        End With
    End If
    DescrizioneSocieta = strDescrSocieta
Exit Function
ErrHandler:
    DescrizioneSocieta = vbNullString
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strFileName As String
    gmRegistry.SaveSetting "Options", "IDAEEExport1", mskIDAEE, True
    gmRegistry.SaveSetting "Options", "GP1Export", mskFileName, True
    gmRegistry.SaveSetting "Options", "SocExport1", MaskSoc, True
    Set AEE = Nothing
    Set AVE = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    Set VVP = Nothing
    Set AUT = Nothing
    Set OCT = Nothing
    Set ADZ = Nothing
    Set ATV = Nothing
    Set ATM = Nothing
    Set AXX = Nothing
    Set ADE = Nothing
    Set ACL = Nothing
    Set APR = Nothing
    Set APD = Nothing
    Set AZO = Nothing
    Set ASC = Nothing
    Set AAM = Nothing
    Set XA = Nothing
    Set frmGP1ExportCEPSA = Nothing
End Sub

Private Sub MaskSoc_LostFocus()
    lblsocieta(8).Caption = DescrizioneSocieta(MaskSoc)
'    Select Case MaskSoc
'        Case "001"
'            lblsocieta(8).Caption = "AGIP ESPAÑA"
'        Case "002"
'            lblsocieta(8).Caption = "CISTERNAS DE ALSAUA "
'        Case "005"
'            lblsocieta(8).Caption = "PETROGAL ESPAÑOLA, S.A."
'        Case "007"
'            lblsocieta(8).Caption = "TAMOIL-OLINVEST ESPAÑA"
'        Case "008"
'            lblsocieta(8).Caption = "SARAS ENERGIA ,S.L."
'        Case "009"
'            lblsocieta(8).Caption = "KUWAIT PETROLEUM ESPAÑA"
'        Case "013"
'            lblsocieta(8).Caption = "MEROIL , S.A."
'        Case "018"
'            lblsocieta(8).Caption = "SOLPETROL COMPANY ,S.L."
'        Case "025"
'            lblsocieta(8).Caption = "ERG PETROLEOS ,S.A"
'        Case "027"
'            lblsocieta(8).Caption = "JAIME ROYO PASCUAL"
'        Case "033"
'            lblsocieta(8).Caption = "TOTALFINA ESPAÑA ,S.A."
'        Case "036"
'            lblsocieta(8).Caption = "CARREFOUR"
'        Case "037"
'            lblsocieta(8).Caption = "CENTROS COMMERCIALES"
'        Case "038"
'            lblsocieta(8).Caption = "CEPSA"
'        Case "900"
'            lblsocieta(8).Caption = "EUROSYSTEM"
'        Case "999"
'            lblsocieta(8).Caption = "GENERICA"
'        Case Else
'            lblsocieta(8).Caption = "SOCIETA' NON TROVATA"
'    End Select
End Sub

Private Sub TDBGrid1_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    GridArraySelect.ReadData RowBuf, StartLocation, offset, ApproximatePosition
End Sub

