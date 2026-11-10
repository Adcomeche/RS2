VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVVTExport2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CEPSA - TRASMISSIONE DISCREPANZE"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   8910
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   29
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
         Picture         =   "frmVVTExport2.frx":0000
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
         Picture         =   "frmVVTExport2.frx":0102
         TabIndex        =   10
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
         Picture         =   "frmVVTExport2.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Viaggi"
         Height          =   375
         Left            =   2280
         Picture         =   "frmVVTExport2.frx":0306
         TabIndex        =   8
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
         Picture         =   "frmVVTExport2.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Seleziona File..."
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTExport2.frx":050A
         TabIndex        =   35
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
         Picture         =   "frmVVTExport2.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7080
         Picture         =   "frmVVTExport2.frx":070E
         TabIndex        =   11
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   4920
      TabIndex        =   22
      Top             =   3600
      Width           =   3975
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Records da trasmettere:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Records trasmessi correttamente:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   1815
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8895
      Begin VB.ComboBox cmbAnno 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbMese 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   8040
         Picture         =   "frmVVTExport2.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   8040
         Picture         =   "frmVVTExport2.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Tag             =   "11"
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
         Left            =   6000
         TabIndex        =   3
         Tag             =   "12"
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
         Left            =   8160
         Top             =   120
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
               Picture         =   "frmVVTExport2.frx":0A14
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTExport2.frx":0B26
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTExport2.frx":0C38
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskSoc 
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Tag             =   "3"
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Anno"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mese"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5640
         Tag             =   "3"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblsocieta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   37
         Tag             =   "3"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Società:"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   36
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio Def."
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   7320
         Tag             =   "12"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal Num. &Viaggio Def."
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3360
         Tag             =   "11"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data Inizio Viaggio"
         Height          =   255
         Index           =   12
         Left            =   4800
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   0
      TabIndex        =   19
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
         Picture         =   "frmVVTExport2.frx":0D4A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Viaggio"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTExport2.frx":0E4C
         TabIndex        =   9
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmVVTExport2.frx":0F4E
         TabIndex        =   7
         Top             =   240
         Width           =   4575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   6525
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
            TextSave        =   "16.14"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "22/09/2004"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmVVTExport2.frx":3956
      TabIndex        =   6
      Tag             =   "SELECT"
      Top             =   1920
      Width           =   8655
   End
End
Attribute VB_Name = "frmVVTExport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const MAX_COLS = 35 ' 4

Private Const COL_ID_NUMEROVIAGGIO = 1
Private Const COL_ID_UNITADITRASPORTO = 2
Private Const COL_SEQVIAGGIO = 3
Private Const COL_ORAINIZIOVIAGGIO = 4
Private Const COL_ID_DEPOSITO = 5
Private Const COL_IDENTEEMITTENTE = 6
Private Const COL_DT_VIAGGIO = 7
Private Const COL_ID_NUMVIAGDEF = 8
Private Const COL_ID_VETTORE = 9
Private Const COL_ID_DEPFINEVIAG = 10
Private Const COL_ID_ENTEEMITFINVIAG = 11
Private Const COL_KM_PREVISTI = 12
Private Const COL_KM_DEFIN = 13
Private Const COL_ID_MOTRICE = 14
Private Const COL_ID_TIPO = 15
Private Const COL_ID_RIMORCHIO = 16
Private Const COL_HR_INIZIOCARICO = 17
Private Const COL_ID_AUTISTA1 = 18
Private Const COL_NOME_AUTISTA1 = 19
Private Const COL_CARTAIDENT1 = 20
Private Const COL_NOME_AUTISTA2 = 21
Private Const COL_ID_MAGAZZINO = 22
Private Const COL_RIFEST1_CONSEGNA = 23
Private Const COL_ID_CLIENTE = 24
Private Const COL_RIFEST2_CONSEGNA = 25
Private Const COL_ORACONSPREV = 26
Private Const COL_NOTA2_1_ORDINE = 27
Private Const COL_NOTA2_2_ORDINE = 28
Private Const COL_NOTA2_3_ORDINE = 29
Private Const COL_DATACONSRICHIORD_VIAGGIO = 30
Private Const COL_PRIORITA_ORDINE = 31
Private Const COL_CODPRODXSOC = 32
Private Const COL_QTA_TRASPORTATA = 33
Private Const COL_QTA_RICHIESTA = 34
Private Const COL_UM = 35

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjOutFileTrip As CTextFile

Private mstrProgrInvio As String

Public GridArraySelect As ArrayGridCls
Public XA As XArrayDB
Public Loading As Boolean

Public DataClass As CViaggioSelezione
Public State As FormState

Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private AUT As CUnitaDiTrasporto
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
        ReDim varLog(0 To 1, 0 To lngRows - 1)
        For i = 0 To lngRows - 1
            varLog(0, i) = DataClass.DS.ListItem(0, i)
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
    If IsNull(DBGrid.Bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.Bookmark, 0)
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
    Dim ErrCode As Errors
    Dim C As CError
    
'     variabili aggiuntive del mobjOutFileTrip CZ    02/09/2003
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
    Dim varNOTA2_3 As Variant
    Dim varRifEst2 As Variant
    Dim varfiller16 As Variant
    Dim strProdotto As String
    Dim varVIAGQTATRASP As Variant
    Dim strDTAVIAGGIO As String
    Dim varNSEQ As Variant
    Dim varOraConsegna As Variant
    Dim strCARTIDAUTISTA As String
    Dim strORACONSPREV As String
    Dim intDifferenza As Integer
    Dim strProgressivo As String

    If MaskSoc = "" Then
        Call MsgBox("E' necessario indicare la società.", vbExclamation + vbOKOnly)
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            ErrCode = ParameterNotFound
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = ParameterNotFound
            GoTo ErrHandler
        End If
        strIDEnteEmittente = .Fields("16").value
    End With
    If Len(strIDEnteEmittente) = 0 Then
        ErrCode = NoEnteEmittente
        GoTo ErrHandler
    End If
    With AEE.DS
        .Fields(.KeyFields(1)).FindValue = strIDEnteEmittente
        If Not AEE.Find Then
            ErrCode = NoEnteEmittente
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = NoEnteEmittente
            GoTo ErrHandler
        End If
        strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(MaskSoc)
    End With
    ' ---
    
'    strCodOperatore = Format$(mskCodOperatore, "00")
'    mobjOutFileTrip.Delimiter = "|"
    
    If Len(mstrProgrInvio) = 0 Then
        With DataClass.DS.DA
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "006", 0)
            If varCounter > 9999 Then
                varCounter = 1
            End If
        End With
        mstrProgrInvio = Format$(varCounter, String(4, "0"))
    End If
    
    If Len(mskFileName) > 0 Then
        strFileName = mskFileName
    End If
    
    strFileName = strFileName & "\" & "VIA_" & Right(mstrProgrInvio, 4) & ".DAT"
    If Not mobjOutFileTrip.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    strDataInvio = Format$(Date, "yyyymmdd")
    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00")
    
    ' Inserisco il record di inizio trasmissione
    strHeader = "00" & mstrProgrInvio & strDataInvio & "VIAJES" & Space(3) & "000001" & "000001" & strIDEnteEmittente & "00" & Space(96)
    Print #mobjOutFileTrip.FileNum, strHeader
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    ' --- Ciclo di elaborazione dei viaggi
    intProgressivoRecord = 1
    For lngRow = 0 To lngNumRows
        strIDZonaPrec = vbNullString
        intRowError = NoError
        
        With ADE.DS
            strDummy = GridArraySelect.value(lngRow, 4) 'ID Deposito
            .Fields("0").FindValue = strDummy
            If Not ADE.Find Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
            strCODDEPXSOC = .MultiSoc.DS.GetFirmCode(MaskSoc)
            strCentroOperativo = .Fields("8").value
        End With
         
        With ADZ.DS
            .Fields(.KeyFields(1)).FindValue = GridArraySelect.value(lngRow, 23)
            If Not ADZ.Find Then
                ErrCode = NoDestinazione
                GoTo ErrHandler
            End If
            If .HowMany = 0 Then
                ErrCode = NoDestinazione
                GoTo ErrHandler
            End If
            strCODDESTXSOC = .MultiSoc.DS.GetFirmCode(MaskSoc)
        End With
         
        With APR.DS
            .Fields(.KeyFields(1)).FindValue = GridArraySelect.value(lngRow, 31)
            If Not APR.Find Then
                ErrCode = NoProdotto
                GoTo ErrHandler
            End If
            If .HowMany = 0 Then
                ErrCode = NoProdotto
                GoTo ErrHandler
            End If
            strProdotto = .MultiSoc.DS.GetFirmCode(MaskSoc, 3)
        End With
        
        With AUT.DS
            strDummy = GridArraySelect.value(lngRow, 1) 'ID Unità di trasporto
            .Fields("0").FindValue = strDummy
            If Not AUT.Find Then
                intRowError = noUT
                GoTo LogResult
            End If
            varBadge = .MultiSoc.DS.GetFirmCode(GridArraySelect.value(lngRow, 4))
'            varBadge = .Fields("14").value
        End With
       
        strOraInizioViaggio = Format$(GridArraySelect.value(lngRow, 3), "hh:mm")                  'Ora inizio viaggio
        With mobjOutFileTrip
            .ClearFields
            .Item("1").value = "21"                                                                'tipo record
            .Item("2").value = GridArraySelect.value(lngRow, 26)                                   'nota2_1
            .Item("3").value = GridArraySelect.value(lngRow, 27)                                   'nota2_2
            .Item("4").value = strIDEnteEmittente                                                  'cod Ente Emittente x Soc.
            .Item("5").value = "00"                                                                'filler5
            .Item("6").value = GridArraySelect.value(lngRow, 22)                                   'consegna rif esterno1
            .Item("7").value = Mid(strCODDESTXSOC, 3)                                              'consegna codice dest per Soc
            .Item("8").value = strProdotto                                                         'cod prodotto per società
            .Item("9").value = Format(GridArraySelect.value(lngRow, 33), String(9, "0"))           'qta iniziale ordine
            .Item("10").value = IIf(UCase(GridArraySelect.value(lngRow, 34)) = "L", "L", "K")      'Unità di misura
            .Item("11").value = Format$(GridArraySelect.value(lngRow, 29), "yyyymmdd")             'Data Cons. Rich.
            .Item("12").value = IIf(UCase(GridArraySelect.value(lngRow, 30) = "S"), "S", "N")      'priorità
            .Item("13").value = strCODDEPXSOC                                                      'deposito
            .Item("14").value = IIf(GridArraySelect.value(lngRow, 28) = vbNullString, "00000", Format$(GridArraySelect.value(lngRow, 28), "00000")) 'nota2_3
            .Item("15").value = GridArraySelect.value(lngRow, 24)                                  'Rif esterno2
            If IsNull(GridArraySelect.value(lngRow, 35)) Or (Len(GridArraySelect.value(lngRow, 35)) = 0) Then
                strNumPartizione = "00"
            Else
                strNumPartizione = Format(GridArraySelect.value(lngRow, 35), "00")
            End If
            .Item("16").value = strNumPartizione                                                   'Partizione ordine
            .Item("17").value = Format(GridArraySelect.value(lngRow, 32), String(9, "0"))          'qta trasportata
            .Item("18").value = Format$(GridArraySelect.value(lngRow, 6), "yyyymmdd")              'data viaggio
            If IsNull(varBadge) Or Len(varBadge) = 0 Then
                intRowError = InvalidDep
                GoTo LogResult
            End If
            .Item("19").value = varBadge                                                           'badge
            .Item("20").value = Format(GridArraySelect.value(lngRow, 2), "00")                     'n° sequenza
            varNIF = GridArraySelect.value(lngRow, 19)
            If IsNull(varNIF) Or Len(varNIF) = 0 Then
                intRowError = InvalidCode
                GoTo LogResult
            End If
            .Item("21").value = "ES" & Replace(Replace(varNIF, " ", ""), "-", "")
            varOraConsegna = Format(GridArraySelect.value(lngRow, 25), "hh:nn")
            .Item("22").value = Left(varOraConsegna, 2) & ":" & Right(varOraConsegna, 2)             'ora consegna prevista
          
            If Not .WriteLine Then
                 intRowError = OutputError
                 GoTo LogResult
            End If
            intProgressivoRecord = intProgressivoRecord + 1
        End With
        strIDViaggio = GridArraySelect.value(lngRow, 0) 'Codice viaggio
        
        
        ' --- Modifico lo stato del viaggio
        With VVT
            .DS.Fields("0").FindValue = strIDViaggio
            If Not .Find Then
                intRowError = NoTripID
                GoTo LogResult
            End If
            If Not .SetFlagTrasmissione("T") Then
                intRowError = NoSetState
                GoTo LogResult
            End If
        End With
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            If (intRowError = ZonaEsclusaExport) Or _
               (intRowError = OrderNotActive) Then
                lngCountExcluded = lngCountExcluded + 1
            Else
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
                strErrorMessage = "Errore durante la lettura delle consegne"
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
    intProgressivoRecord = intProgressivoRecord + 1
    strProgressivo = Format(intProgressivoRecord, "0000")
    'intDifferenza = 4 - Len(strProgressivo)
    ' Inserisco il record di fine trasmissione
    strTail = "90" & mstrProgrInvio & strDataInvio & "VIAJES" & Space(3) & "000001" & "000001" & strIDEnteEmittente & "00" & strProgressivo & Space(92)
    Print #mobjOutFileTrip.FileNum, strTail
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & strFileName
        Case ParameterNotFound
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti: Prodotti"
        Case NoEnteEmittente
            strMessage = "Ente Emittente non trovato."
            strInfo = "Ente emittente: " & strIDEnteEmittente
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
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover", _
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
            If Not DataClass.GetSelectionExport1() Then
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
    MaskSoc = gmRegistry.GetSetting("Options", "Socexport1", , True)
    mskFileName = gmRegistry.GetSetting("Options", "TripExport1", , True)
        
    Set mobjOutFileTrip = New CTextFile
    With mobjOutFileTrip
          
        .Add "Tipo Record", 2, "1"
        .Add "Note2_1", 6, "2"
        .Add "Note2_2", 6, "3"
        .Add "Codice Ente e Società", 4, "4"
        .Add "Filler5", 2, "5"
        .Add "Consegna. Rif.Esterno1", 13, "6"
        .Add "Consegna.Codice Destinatario x Società", 8, "7"
        .Add "Dettaglio.Consegna.Cod.Prodotto x Società", 7, "8"
        .Add "Qta iniziale ordine", 9, "9"
        .Add "Unità di misura", 1, "10"
        .Add "Data consegna richiesta ordine", 8, "11"
        .Add "Ordine.priorità", 1, "12"
        .Add "Codice deposito x società", 4, "13"
        .Add "Ordine.Nota2_3", 5, "14"
        .Add "Rif Esterno2.Consegna", 4, "15"
        .Add "Partizione", 2, "16"
        .Add "Viaggio.prodotto.quantità trasportata", 9, "17"
        .Add "Data Viaggio", 8, "18"
        .Add "Badge", 11, "19"
        .Add "N Sequenza", 2, "20"
        .Add "Cart Identità Autista", 20, "21"
        .Add "Ora consegna prevista", 5, "22"

    End With
    D = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
    mskDateFrom.SelText = D
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
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
       
    Set AVE = New CVettore
    With AVE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
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
    'gmRegistry.SaveSetting "Options", "OperatorCode", mskCodOperatore, True
    gmRegistry.SaveSetting "Options", "TripExport1", mskFileName, True
    gmRegistry.SaveSetting "Options", "Socexport1", MaskSoc, True
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
    Set XA = Nothing
    Set frmVVTExport2 = Nothing
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

