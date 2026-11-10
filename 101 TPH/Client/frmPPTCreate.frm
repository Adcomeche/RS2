VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPPTCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VALORIZZAZIONE CONSEGNE"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   9270
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmPPTCreate.frx":0000
      TabIndex        =   19
      Tag             =   "SELECT"
      Top             =   4440
      Width           =   5175
   End
   Begin TrueDBGrid60.TDBGrid tdbLog 
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "frmPPTCreate.frx":35EA
      TabIndex        =   41
      Top             =   3240
      Width           =   9015
   End
   Begin MSMask.MaskEdBox mskDescrizione 
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   35
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   2775
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   9255
      Begin VB.OptionButton optTipoTariffa 
         Caption         =   "Valorizzazione passiva"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   43
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton optTipoTariffa 
         Caption         =   "Valorizzazione attiva"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   42
         Top             =   1800
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3360
         Picture         =   "frmPPTCreate.frx":6BDA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3360
         Picture         =   "frmPPTCreate.frx":6CDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.ListBox lstZone 
         Height          =   1410
         Left            =   5880
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Tag             =   "97"
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
      Begin MSMask.MaskEdBox mskDateTo 
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Tag             =   "98"
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Tag             =   "13"
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField11 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Tag             =   "92"
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSociet‡ 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Tag             =   "3"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Tag             =   "4"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Tag             =   "39"
         Top             =   2400
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
         Index           =   11
         Left            =   5400
         Tag             =   "39"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vettore"
         Height          =   255
         Index           =   11
         Left            =   1200
         TabIndex        =   48
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblField10 
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
         TabIndex        =   47
         Tag             =   "39"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         TabIndex        =   46
         Tag             =   "4"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emittente"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5400
         Tag             =   "4"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5400
         Tag             =   "3"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Societ‡"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         TabIndex        =   45
         Tag             =   "3"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   5400
         Tag             =   "92"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ra&ggruppamento"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   12
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblField8 
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
         TabIndex        =   44
         Tag             =   "92"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "&Zone"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   36
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
         TabIndex        =   17
         Tag             =   "13"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit‡ di Trasporto"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5400
         Tag             =   "13"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3840
         Tag             =   "98"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data Inizio Carico"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3840
         Tag             =   "97"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data Inizio Carico"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   5400
      TabIndex        =   26
      Top             =   4320
      Width           =   3735
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   39
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Viaggi valorizzati:"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   38
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Scarichi elab. correttamente:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   29
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Scarichi da elaborare:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   24
      Top             =   6000
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   2400
         Picture         =   "frmPPTCreate.frx":6DDE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Esegui"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmPPTCreate.frx":6EE0
         TabIndex        =   21
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
         Picture         =   "frmPPTCreate.frx":6FE2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Consegne"
         Height          =   375
         Left            =   120
         Picture         =   "frmPPTCreate.frx":70E4
         TabIndex        =   20
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
         Picture         =   "frmPPTCreate.frx":71E6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmPPTCreate.frx":72E8
         TabIndex        =   22
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   6855
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
            TextSave        =   "18:26"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "21/01/2020"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   1320
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
            Picture         =   "frmPPTCreate.frx":73EA
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPPTCreate.frx":74FC
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPPTCreate.frx":760E
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblIDTestata 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7800
      TabIndex        =   40
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Num. Valorizzazione"
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   37
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Descrizione"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "frmPPTCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 12 '11
Private Const COL_RESULT_LOG = 3
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Public GridArraySelect As ArrayGridCls
Public Loading As Boolean

Public DataClass As CViaggioSelezione
Public XA As XArrayDB
Public State As FormState

Private AZO As CZona
Private ACN As CContratto
Private ADZ As CDestinazione
Private APR As CProdotto
Private PPT As CPrefattura
Private PPD As CPrefatturaDettaglio
Private PPA As CPrefatturaDettaglioAddebito
Private TFA As CFasciaTariffaria
Private TFR As CFasciaTariffariaRiga
Private TFB As CTariffaBinomica
Private TBR As CTariffaBinomicaRiga
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private PDS As CPrefatturaDettaglioSelezione
Private OCT As COrdine
Private AUT As CUnitaDiTrasporto


Private Const CONTRATTOID_AGIP = "1"
Private Const CONTRATTOID_CEPSA = "3"


Private Sub ElaboraMinimo(IDTestata As String)
    On Error GoTo ErrHandler
        
    Dim lngNumRows As Long
    Dim lngUltimaRiga As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim j As Long
    Dim i As Integer
    Dim strIDUT As String
    Dim strIDZona As String
    Dim strIDFasciaTariffaria As String
    Dim strIDValutaTariffaMinima As String
    Dim strTipoPercorso As String
    Dim strDataViaggio As String
    Dim sngTotImporti As Single
    Dim sngTariffaMinima As Single
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    StatusBar1.Panels(2).Text = "Attribuzione del minimo garantito in corso... "
    
    If Not SelectPrefatture(IDTestata) Then
        GoTo ErrHandler
    End If
    If IsEmpty(PDS.DS.List) Then
        GoTo LogResult
    End If
    
    lngNumRows = UBound(PDS.DS.List, 2)
    For i = 0 To lngNumRows
        intRowError = NoError
        With PDS.DS
            strIDUT = .ListItem(0, j)
            strIDZona = .ListItem(1, j)
            strDataViaggio = .ListItem(2, j)
            sngTotImporti = .ListItem(3, j)
            strTipoPercorso = .ListItem(4, j)
            lngUltimaRiga = .ListItem(5, j)
        End With
        strIDFasciaTariffaria = TFA.FindByDate(strIDZona, strDataViaggio, "A")
        
        With TFA.DS
            .Fields("0").FindValue = strIDZona
            .Fields("1").FindValue = strIDFasciaTariffaria
            If Not TFA.Find Then
                GoTo LogResult
            End If
            If strTipoPercorso = "L" Then
                sngTariffaMinima = .Fields("6").value
                strIDValutaTariffaMinima = .Fields("7").value
            Else
                sngTariffaMinima = .Fields("4").value
                strIDValutaTariffaMinima = .Fields("5").value
            End If
        End With
        If sngTotImporti < sngTariffaMinima Then
            With PPD
                .DS.DA.BeginTransaction
                If Not .AzzeraConsegne(IDTestata, strIDUT, strIDZona, _
                                       strDataViaggio, lngUltimaRiga) Then
                    GoTo LogResult
                End If
                If Not .ApplicaMinimo(sngTariffaMinima, IDTestata, strIDUT, _
                                      strIDZona, strDataViaggio, lngUltimaRiga) Then
                    GoTo LogResult
                End If
            End With
        End If
LogResult:
        ' --- Scrivo il risultato dell'elaborazione del minimo
        If intRowError = NoError Then
            PPD.DS.DA.EndTrans ETCOMMIT
            lngCountOK = lngCountOK + 1
        Else
            PPD.DS.DA.EndTrans ETROLLBACK
            lngCountFail = lngCountFail + 1
        End If
        lngCountToGo = lngCountToGo - 1
        ' ---
        DoEvents
    Next i
    
    GoTo EndProc
    
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmPPTCreate", _
                "ElaboraMinimo", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    StatusBar1.Panels(2).Text = vbNullString
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub

Private Function RowToPPD(RowIndex As Long, _
                          IDDettaglio As Long, _
                          PPD As CPrefatturaDettaglio, _
                          IDTestata As String, _
                          IDZonaI As String, _
                          IDFasciaTariffariaI As String, _
                          TariffaI As Single, _
                          ImportoI As Single, _
                          IDValutaI As String, _
                          TipoPercorsoI As String, _
                          KmRaggioInterni As Long, _
                          RiferimentoEsterno As String, _
                          QtaScaricata As Single, _
                          AnomaliaI As AnomaliePrefattura, _
                          QtaScaricataKg As Single, _
                          Optional Pedaggio As Single, _
                          Optional ServOperario As Single, _
                          Optional ServBuques As Single, _
                          Optional ServOreSeServBuquesrvSpeciali As Single, _
                          Optional ServServSpeciali As Single) As Errors
    On Error GoTo ErrHandler
    
    Dim intError As Errors
    Dim intCols As Integer
    Dim i As Integer
    
    intError = NoError
    
    intCols = UBound(DataClass.DS.List, 1)
    ReDim varrow(intCols, 0) As Variant
    With DataClass.DS
        For i = 0 To intCols
            varrow(i, 0) = .ListItem(i, RowIndex)
        Next i
    End With
    
    With PPD.DS
        .Fields("0").NewValue = IDTestata
        .Fields("1").NewValue = IDDettaglio
        .Fields("2").NewValue = varrow(4, 0)
        .Fields("3").NewValue = varrow(5, 0)
        .Fields("4").NewValue = varrow(6, 0)
        .Fields("5").NewValue = varrow(7, 0)
        .Fields("6").NewValue = RiferimentoEsterno 'varrow(8, 0)
'        .Fields("7").NewValue = varrow(9, 0)
        .Fields("8").NewValue = varrow(10, 0)
        If varrow(11, 0) = 0 Then
            .Fields("9").NewValue = varrow(35, 0)
        Else
            .Fields("9").NewValue = varrow(11, 0)
        End If
        .Fields("10").NewValue = varrow(12, 0)
        .Fields("11").NewValue = varrow(13, 0)
        .Fields("12").NewValue = varrow(0, 0)
        .Fields("13").NewValue = varrow(14, 0)
        .Fields("14").NewValue = varrow(15, 0)
        ' Data consegna
        If Format(varrow(16, 0), "dd/mm/yyyy") = "01/01/1900" Then
            .Fields("15").NewValue = varrow(39, 0)
       Else
            .Fields("15").NewValue = varrow(16, 0)
        End If
        .Fields("16").NewValue = varrow(17, 0)
        .Fields("17").NewValue = varrow(18, 0)
        .Fields("18").NewValue = varrow(19, 0)
        .Fields("19").NewValue = varrow(20, 0)
        .Fields("20").NewValue = varrow(21, 0)
        .Fields("21").NewValue = varrow(22, 0)
        .Fields("22").NewValue = varrow(23, 0)
        .Fields("23").NewValue = varrow(24, 0)
        .Fields("24").NewValue = varrow(25, 0)
        .Fields("25").NewValue = varrow(26, 0)
        .Fields("26").NewValue = varrow(2, 0)
        .Fields("27").NewValue = varrow(27, 0)
        .Fields("28").NewValue = varrow(28, 0)
        .Fields("29").NewValue = QtaScaricataKg 'varrow(29, 0)
        .Fields("30").NewValue = QtaScaricata 'varrow(3, 0)
        .Fields("31").NewValue = varrow(30, 0)
        .Fields("32").NewValue = varrow(31, 0)
        
        .Fields("55").NewValue = varrow(34, 0)
        .Fields("55").Updated = True
        
        ' --- Valorizzazione interna
        .Fields("41").NewValue = IDZonaI
        .Fields("42").NewValue = IDFasciaTariffariaI
        .Fields("43").NewValue = TariffaI
        .Fields("44").NewValue = ImportoI
        .Fields("45").NewValue = ""
        .Fields("46").NewValue = IDValutaI
        .Fields("47").NewValue = KmRaggioInterni 'varRow(33, 0)
        If Not IsMissing(Pedaggio) Then
            .Field("Importo Pedaggio Interno").NewValue = Pedaggio
            .Field("Importo Pedaggio Interno").Updated = True
        End If
        If Not IsMissing(ServOperario) Then
            .Field("Importo Servizio Operario Interno").NewValue = ServOperario
            .Field("Importo Servizio Operario Interno").Updated = True
        End If
        If Not IsMissing(ServBuques) Then
            .Field("Importo Servizio Buques Interno").NewValue = ServBuques
            .Field("Importo Servizio Buques Interno").Updated = True
        End If
        If Not IsMissing(ServOreSeServBuquesrvSpeciali) Then
            .Field("Ore Servizi Speciali Interno").NewValue = ServOreSeServBuquesrvSpeciali
            .Field("Ore Servizi Speciali Interno").Updated = True
        End If
        If Not IsMissing(ServServSpeciali) Then
            .Field("Importo Servizi Speciali Interno").NewValue = ServServSpeciali
            .Field("Importo Servizi Speciali Interno").Updated = True
        End If
        
        .Fields("48").NewValue = TipoPercorsoI
        .Fields("50").NewValue = .Fields("50").NewValue Or AnomaliaI
        ' ---
        
        If IsDate(varrow(44, 0)) Then
            .Field("Data Grado Servizio").NewValue = varrow(44, 0)
            .Field("Data Grado Servizio").Updated = True
        End If
        
        For i = 0 To 32
            .Fields(CStr(i)).Updated = True
        Next i
        For i = 41 To 50
            .Fields(CStr(i)).Updated = True
        Next i
    End With
    RowToPPD = intError
    Exit Function
ErrHandler:
    If intError = NoError Then
        intError = UnknownError
    End If
    RowToPPD = intError
End Function

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
                cmdExport.Enabled = True
            Else
                cmdExport.Enabled = False
            End If
        Case fcExport
            If Len(Trim$(mskDescrizione)) > 0 Then
                CreateRecords
                cmdExport.Enabled = False
            Else
                MsgBox "Occorre inserire una descrizione", vbCritical + vbOKOnly, "ATTENZIONE"
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub CreateRecords()
    On Error GoTo ErrHandler
    
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngIDDettaglio As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountViaggiOK As Long
    Dim lngKm As Long
    Dim lngMaxKM As Long
    Dim lngKmEsterni As Long
    Dim lngKmDefViaggio As Long
    Dim lngNumConsegne As Long
    Dim lngNumScarichi As Long
    Dim j As Long
    Dim i As Long
    Dim k As Long
    Dim strIDZona As String
    Dim strIDTestata As String
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim strIDScarico As String
    Dim strIDOrdine As String
    Dim strIDViaggioPrec As String
    Dim strIDConsegnaPrec As String
    Dim strIDRaggrDest As String
    Dim strRifOrdineEsterno As String
    Dim strRifOrdineRiga As String
    Dim strPartizione As String
'    Dim strDataViaggio As String
    Dim strDataConsegna As String
    Dim strDataConsegnaRichiesta As String
    Dim strErrorMessage As String
    Dim strIDValuta As String
    Dim strIDValutaDefault As String
    Dim strIDFasciaTariffaria As String
    Dim strIDFasciaTariffariaBusqueda As String
    Dim strIDDestinazione As String
    Dim strIDDeposito As String
    Dim StrIDSociet‡ As String
    Dim strTipoPercorso As String
    Dim strIDProdotto As String
    Dim strBollettaScarico As String
    Dim strIDUdT As String
    Dim strIDEnteEmittente As String
    Dim strIDTipologiaProdotto As String
    Dim strIDVettore As String
    Dim strFlagSoloTrattore As String
    Dim strFlagResponsabilit‡ As String
    Dim strContratto As String
    Dim strTipoTariffa As String
    Dim strModalit‡Tariffa As String
    Dim varDensit‡Prodotto As Variant
    Dim varTipoCisterna As Variant
    Dim varBufferDistanze As Variant
    Dim varBufferServiziSpeciali As Variant
    Dim blnFound As Boolean
    Dim blnPropostaMancante As Boolean
    Dim blnResult As Boolean
    Dim blnViaggioOK As Boolean
    Dim blnConcurErr As Boolean
    Dim blnFlagProgrammato As Boolean
    Dim blnFatturazioneAttiva As Boolean
    Dim blnUltimaConsegna As Boolean
    Dim blnConsegnaPi˘Distante As Boolean
    Dim blnFlagKilometrica As Boolean
    Dim blnFestivo As Boolean
    Dim blnMultiDestino As Boolean
    Dim sngQScaricata As Single
    Dim sngQScaricataKg As Single
    Dim sngQCaricataKg As Single
    Dim sngQCaricataL As Single
    Dim sngQIniziale As Single
    Dim sngTotQScaricata As Single
    Dim sngTotQScaricataDestinazione As Single
    Dim sngTotQScaricataConsegna As Single
    Dim sngTotQScaricataConsegnaKG As Single
    Dim sngTotQOrdinata As Single
    Dim sngTotQScaricataViaggio As Single
    Dim sngTotQScaricataViaggioMultiDestino As Single
    Dim sngTotQOrdinataViaggio As Single
    Dim sngTotQOrdinataViaggioMultiDestino As Single
    Dim sngTotQScaricataKG As Single
    Dim sngTotQIniziale As Single
    Dim sngQta As Single
    Dim sngQtaKg As Single
    Dim sngQtaServiziSpeciali As Single
    Dim sngTariffaServiziSpeciali As Single
    Dim sngTariffa As Single
    Dim sngImporto As Single
    Dim sngQuantit‡Vuoto As Single
    Dim sngPedaggio As Single
    Dim sngPeajeManual As Single
    Dim sngPeajeMasAlto As Single
    Dim sngServOperario As Single
    Dim sngServBuquesPuerto As Single
    Dim sngServOreServiziSpeciali As Single
    Dim sngServServiziSpeciali As Single
    Dim sngOCTNota2 As String
    Dim intRowError As Errors
    Dim intRowErrorLog As Errors
    Dim intAnomalia As AnomaliePrefattura
    Dim ErrCode As Errors
    Dim TFB As CTariffaBinomica
    Dim VCE As CViaggioEvento
    Dim C As CError
    Dim EntregasporDestino As Integer
    ' -- DP 07/01/09
    Dim blnFoundOB As Boolean
    Dim blnFoundBP As Boolean
    Dim blnFoundSE As Boolean
    Dim blnFoundSV As Boolean
    Dim blnFoundTO As Boolean
    
    Dim VarkbufferB3 As Variant
    Dim lngHowManyB3 As Long
    Dim CobrarB3 As Boolean
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    
    StatusBar1.Panels(2).Text = "Valorizzazione delle consegne in corso... "
    
    
    Set TFB = New CTariffaBinomica
    With TFB.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    Set VCE = New CViaggioEvento
    With VCE.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    blnFatturazioneAttiva = optTipoTariffa(0).value
    If blnFatturazioneAttiva Then
        strTipoTariffa = "A"
    Else
        strTipoTariffa = "P"
    End If
    
    ' --- Inserisco la testata della valorizzazione
    With PPT.DS
        .Fields("1").NewValue = mskDateFrom
        .Fields("1").Updated = True
        .Fields("2").NewValue = mskDateTo
        .Fields("2").Updated = True
        .Fields("3").NewValue = mskDescrizione
        .Fields("3").Updated = True
        .Fields("4").NewValue = strTipoTariffa
        .Fields("4").Updated = True
        .Fields("5").NewValue = mskSociet‡
        .Fields("5").Updated = True
    End With
    
    If Not PPT.Insert Then
        ErrCode = InsertError
        GoTo ErrHandler
    End If
    strIDTestata = PPT.DS.Fields("0").value
    lblIDTestata = strIDTestata
    ' ---
    
    blnViaggioOK = True
    lngIDDettaglio = 1
    
    ' --- Ciclo di elaborazione delle consegne
    For lngRow = 0 To lngNumRows
    

        
        With DataClass.DS
            blnUltimaConsegna = False
            strIDViaggio = .ListItem(0, lngRow)
            If lngRow = lngNumRows Then
                blnUltimaConsegna = True
            ElseIf strIDViaggio <> .ListItem(0, lngRow + 1) Then
                blnUltimaConsegna = True
            End If
            strIDConsegna = .ListItem(14, lngRow)
            strIDScarico = .ListItem(15, lngRow)
            strIDOrdine = .ListItem(4, lngRow)
            strDataConsegna = .ListItem(16, lngRow)
            If Format(strDataConsegna, "dd/mm/yyyy") = "01/01/1900" Then
                strDataConsegna = .ListItem(39, lngRow)
            End If
            strDataConsegnaRichiesta = .ListItem(7, lngRow)
            strIDVettore = .ListItem(47, lngRow)
            If strIDVettore <> .ListItem(46, lngRow) Then
                strFlagSoloTrattore = "S"
            Else
                strFlagSoloTrattore = "N"
            End If
        End With
        
        ' --- Se inizio l'elaborazione di un nuovo viaggio
        ' --- chiudo l'elaborazione del viaggio precedente
        If strIDViaggio <> strIDViaggioPrec Then
            If Len(strIDViaggioPrec) > 0 Then
                If blnViaggioOK Then
                    If blnFatturazioneAttiva Then
                        VVT.DS.Fields("0").FindValue = strIDViaggioPrec
                        If VVT.Find Then
                            If VVT.SetStato("E") Then
                                PPD.DS.DA.EndTrans ETCOMMIT
                                lngCountViaggiOK = lngCountViaggiOK + 1
                                lblLog(7) = lngCountViaggiOK
                            Else
                                PPD.DS.DA.EndTrans ETROLLBACK
                            End If
                        Else
                            PPD.DS.DA.EndTrans ETROLLBACK
                        End If
                    Else
                        PPD.DS.DA.EndTrans ETCOMMIT
                        lngCountViaggiOK = lngCountViaggiOK + 1
                        lblLog(7) = lngCountViaggiOK
                    End If
                Else
                    PPD.DS.DA.EndTrans ETROLLBACK
                End If
            End If
            blnViaggioOK = True
            strIDViaggioPrec = strIDViaggio
            strIDConsegnaPrec = vbNullString
            sngTotQScaricataKG = 0
            
            ' Contiamo le righe che fanno parte del viaggio che iniziamo ad elaborare
            ' e calcoliamo il pedaggio per scarico
            sngPeajeManual = 0
            sngPedaggio = 0
            sngPeajeMasAlto = 0
            For i = lngRow To lngNumRows
            
                If DataClass.DS.ListItem(0, i) = strIDViaggio Then
                           
                        If IsNumeric(DataClass.DS.ListItem(43, i)) And (DataClass.DS.ListItem(43, i)) > 0 Then
                        sngPedaggio = DataClass.DS.ListItem(43, i)
                        Else
                            varBufferDistanze = ADZ.GetDistanza(DataClass.DS.ListItem(25, i), DataClass.DS.ListItem(23, i))
                            If Not (IsEmpty(varBufferDistanze)) Then
                                If Not IsNumeric(varBufferDistanze(2, 0)) Then
                                    sngPedaggio = sngPeajeMasAlto
                                Else
                                    If varBufferDistanze(2, 0) > sngPeajeMasAlto Then
                                    sngPeajeMasAlto = varBufferDistanze(2, 0)
                                    sngPedaggio = sngPeajeMasAlto
                                    End If
                                End If
                           End If
                        End If
                End If
                
                If DataClass.DS.ListItem(0, i) <> strIDViaggio Then
                    Exit For
                End If
                
            Next i
            lngNumScarichi = i - lngRow
            
            sngPedaggio = Format(sngPedaggio / lngNumScarichi, "0.00")
            
            PPD.DS.DA.BeginTransaction
        End If
        ' ---
        
        intRowError = NoError
        intRowErrorLog = NoError
        intAnomalia = apNoAnomalie
        
        With DataClass.DS
            strIDZona = .ListItem(32, lngRow)
            sngQScaricata = .ListItem(3, lngRow)
            sngQScaricataKg = .ListItem(29, lngRow)
            sngQCaricataKg = .ListItem(17, lngRow)
            sngQCaricataL = .ListItem(18, lngRow)
            If (IsNull(.ListItem(11, lngRow))) Or (.ListItem(11, lngRow) = 0) Then
                sngQIniziale = .ListItem(35, lngRow)
            Else
                sngQIniziale = .ListItem(11, lngRow)
            End If
            lngKm = .ListItem(33, lngRow)
            strIDDestinazione = .ListItem(25, lngRow)
            strIDUdT = .ListItem(34, lngRow)
            strIDProdotto = .ListItem(2, lngRow)
            strRifOrdineEsterno = .ListItem(8, lngRow)
            strIDDeposito = .ListItem(23, lngRow)
            StrIDSociet‡ = .ListItem(21, lngRow)
            strIDEnteEmittente = .ListItem(22, lngRow)
            strFlagResponsabilit‡ = vbNullString
            If Not IsNull(.ListItem(48, lngRow)) Then
                strFlagResponsabilit‡ = .ListItem(48, lngRow)
            End If
            If Len(.ListItem(41, lngRow)) > 0 Then
                varTipoCisterna = .ListItem(41, lngRow)
            Else
                varTipoCisterna = .ListItem(40, lngRow)
            End If
            If IsNumeric(.ListItem(42, lngRow)) Then
                lngKmDefViaggio = .ListItem(42, lngRow)
            Else
                lngKmDefViaggio = 0
            End If
            
            Select Case StrIDSociet‡
                Case ID_SOC_CEPSA
                
                    Select Case .ListItem(45, lngRow)
                    Case "EA", "O"
                    Case Else
                        If IsNumeric(.ListItem(37, lngRow)) Then
                            strRifOrdineRiga = Format(.ListItem(37, lngRow), "000")
                        Else
                            strRifOrdineRiga = "001"
                        End If
                        If IsNumeric(.ListItem(38, lngRow)) Then
                            strPartizione = Format(.ListItem(38, lngRow), "00")
                        Else
                            strPartizione = "00"
                        End If
                        strRifOrdineEsterno = strRifOrdineEsterno & strRifOrdineRiga & strPartizione
                    End Select
                    
                    With VCE.DS
                        .Fields.ResetFields
                        .Fields("0").FindValue = strIDViaggio
                        .Fields("1").FindValue = strIDConsegna
                        .Fields("2").FindValue = strIDScarico
                        VCE.Find
                        If VCE.DS.HowMany > 0 Then
                            ' Se rientro merce
                            If VCE.DS.Fields("4").value = "011" Then
                                sngQScaricata = sngQCaricataL
                                sngQScaricataKg = sngQCaricataKg
                            End If
                        End If
                    End With
                    
            End Select
        End With
        PPD.DS.Fields.ResetFields
        
        AUT.DS.Fields.ResetFields
        AUT.DS.Fields("0").FindValue = strIDUdT
        AUT.Find
        strContratto = AUT.DS.Fields("15").value
        strModalit‡Tariffa = vbNullString
        
        If (StrIDSociet‡ = ID_SOC_CEPSA) Or (StrIDSociet‡ = ID_SOC_GUARNER) Then strModalit‡Tariffa = "B"
        
        Select Case strModalit‡Tariffa
            Case "B"
                APR.DS.Fields.ResetFields
                APR.DS.Fields("0").FindValue = strIDProdotto
                APR.Find
                varDensit‡Prodotto = APR.GetDensity(strDataConsegna)
                strIDTipologiaProdotto = APR.DS.Field("Tipologia Prodotto").value
                
                If Len(varTipoCisterna) > 0 Then
                    strIDFasciaTariffaria = TFB.FindByDate(strIDDeposito, strDataConsegna, strTipoTariffa, strContratto, StrIDSociet‡, strIDEnteEmittente, strIDTipologiaProdotto, strIDDestinazione, CStr(varTipoCisterna), strIDVettore, strFlagSoloTrattore)
                    If Len(strIDFasciaTariffaria) = 0 Then
                        strIDFasciaTariffaria = TFB.FindByDate(strIDDeposito, strDataConsegna, strTipoTariffa, strContratto, StrIDSociet‡, strIDEnteEmittente, strIDTipologiaProdotto, strIDDestinazione, , strIDVettore, strFlagSoloTrattore)
                    End If
                Else
                    strIDFasciaTariffaria = TFB.FindByDate(strIDDeposito, strDataConsegna, strTipoTariffa, strContratto, StrIDSociet‡, strIDEnteEmittente, strIDTipologiaProdotto, strIDDestinazione, , strIDVettore, strFlagSoloTrattore)
                End If
                If Len(strIDFasciaTariffaria) = 0 Then
                    intRowErrorLog = NoFasciaTariffaria
                    GoTo LogResult
                End If
                sngQta = sngQScaricata
                
                ADZ.DS.Fields.ResetFields
                ADZ.DS.Fields("0").FindValue = strIDDestinazione
                ADZ.Find
                
'                If blnFatturazioneAttiva Then 'Or ((strIDDeposito <> "300") And (strIDDeposito <> "030")) Then
                    varBufferDistanze = ADZ.GetDistanza(strIDDestinazione, strIDDeposito)
                    If IsEmpty(varBufferDistanze) Then
                        If Not IsNumeric(ADZ.DS.Fields("29").value) Then
                            lngKm = 0
                        Else
                            lngKm = ADZ.DS.Fields("29").value
                        End If
                    Else
                        If blnFatturazioneAttiva Then
                            If IsNumeric(varBufferDistanze(0, 0)) Then
                                lngKm = varBufferDistanze(0, 0)
                            Else
                                lngKm = 0
                            End If
                        Else
                            If IsNumeric(varBufferDistanze(1, 0)) Then
                                lngKm = varBufferDistanze(1, 0)
                            Else
                                lngKm = 0
                            End If
                        End If
                    End If
'                Else
'                    lngKm = lngKmDefViaggio
'                End If
                
                sngTotQScaricataConsegna = 0
                VVC.DS.Fields.ResetFields
                VVC.DS.Field("Numero Viaggio").FindValue = strIDViaggio
                VVC.DS.Field("Numero Consegna").FindValue = strIDConsegna
                VVC.Find
                VVC.GetTotali
                If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                    sngTotQScaricataConsegna = VVC.TotaliQta(5)
                    sngTotQScaricataConsegnaKG = VVC.TotaliQta(4)
                End If
                
                sngTotQScaricataViaggioMultiDestino = 0
                sngTotQOrdinataViaggioMultiDestino = 0
                
                blnMultiDestino = False
                If Len(ADZ.DS.Fields("11").value) > 0 Then
                    blnMultiDestino = True
                    strIDRaggrDest = ADZ.DS.Fields("11").value
                    VVC.GetTotaliMultiDestinoData strDataConsegna, strIDRaggrDest, strIDDeposito
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQScaricata = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                    VVC.GetOrdinatoMultiDestinoData strDataConsegna, strIDRaggrDest, strIDDeposito
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQOrdinata = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                    VVC.GetTotaliMultiDestinoViaggio strIDViaggio, strIDRaggrDest
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQScaricataViaggioMultiDestino = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                    VVC.GetTotaliOrdineMultiDestinoViaggio strIDViaggio, strIDRaggrDest
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQOrdinataViaggioMultiDestino = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                Else
                    VVC.GetTotaliDestinazioneData strDataConsegna, strIDDeposito
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQScaricata = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                    VVC.GetTotaliDestinazioneViaggio strIDViaggio, strIDDestinazione
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQScaricataDestinazione = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                    VVC.GetTotaliPerOrdine strDataConsegna, strIDDeposito
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQOrdinata = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                    VVC.GetTotaliOrdineViaggio strIDViaggio, strIDDestinazione
                    If (Not IsNull(VVC.TotaliQta(5))) And (Not IsEmpty(VVC.TotaliQta(5))) Then
                        sngTotQOrdinataViaggio = VVC.TotaliQta(5)
                    Else
                        intRowErrorLog = NoOrderID
                        GoTo LogResult
                    End If
                End If
                
                If sngQScaricataKg = 0 Then
                    If Not IsEmpty(varDensit‡Prodotto) Then
                        ' Il round non funziona una beata fava in VB
                        ' aggiunto uno .001 per farlo funzionare correttamente
                        
                        ' Bella soluzione del cavolo, peggio la cura del male
                        'sngQtaKg = Round(sngQScaricata * varDensit‡Prodotto + 0.001)
                        
                        sngQtaKg = Format(sngQScaricata * varDensit‡Prodotto, "0")
                        sngQScaricataKg = sngQtaKg
                    End If
                Else
                    sngQtaKg = sngQScaricataKg
                End If
                sngTotQScaricataKG = sngTotQScaricataKG + sngQtaKg
                
                TFB.DS.Fields("0").FindValue = strIDDeposito
                TFB.DS.Fields("1").FindValue = strIDFasciaTariffaria
                TFB.Find
                blnFlagKilometrica = (TFB.DS.Fields("13").value = "K")
                
                VVT.DS.Fields("0").FindValue = strIDViaggio
                VVT.Find
                blnFestivo = (VVT.DS.Fields("38").value = "F")
                lngNumConsegne = VVT.ContaDestinazioni
                If lngNumConsegne = -1 Then
                    GoTo ErrHandler
                End If
                blnConsegnaPi˘Distante = VVC.IsDestinazionePi˘Lontana2(strIDViaggio, strIDConsegna, strIDDeposito)
                If blnConsegnaPi˘Distante Then
                    lngMaxKM = lngKm
                End If
                VVT.GetTotali
                sngTotQScaricataViaggio = VVT.TotaliQta(5)
                If (Not blnFatturazioneAttiva) And blnFlagKilometrica And blnUltimaConsegna Then
                    lngKm = lngMaxKM
                End If
                
                
                sngServOperario = 0
                sngServBuquesPuerto = 0
                sngServOreServiziSpeciali = 0
                sngServServiziSpeciali = 0
                
                varBufferServiziSpeciali = ADZ.GetServiziSpeciali(strIDDestinazione, strDataConsegna)
                
                If Not IsEmpty(varBufferServiziSpeciali) Then
                    ' -- DP 07/01/09
                    ' -- Queste 3 variabili servono per non sovrascrivere l'importo calcolato:
                    ' -- nei casi in cui sono presenti pi˘ tariffe storicizzate per la stessa ADZ devo prendere solamente la pi˘ recente
                    blnFoundOB = False
                    blnFoundBP = False
                    blnFoundSE = False
                    blnFoundSV = False
                    blnFoundTO = False
                    
                    For k = 0 To UBound(varBufferServiziSpeciali, 2)
                        sngTariffaServiziSpeciali = varBufferServiziSpeciali(1, k)
                        Select Case varBufferServiziSpeciali(2, k)
                            Case "L"
                                sngQtaServiziSpeciali = Format(sngQScaricata / 1000, "0.000")
                                
                            Case "K"
                                sngQtaServiziSpeciali = Format(sngQScaricataKg / 1000, "0.000")
                                
                            Case Else
                                intRowErrorLog = InvalidDetail
                                GoTo LogResult
                                
                        End Select
                        
                        Select Case varBufferServiziSpeciali(0, k)
                            Case "OB" ' Canon puerto
                                ' -- DP 07/01/09
                                If Not blnFoundOB Then
                                    sngServOperario = Format(sngQtaServiziSpeciali * sngTariffaServiziSpeciali, "0.00") + sngServOperario
                                    blnFoundOB = True
                                End If
                                
                            Case "TO" ' Nuevo Canon Puerto extra
                                ' -- JG 04/1/12 Nuevo invento del TBO de poner una tarifa ficticia dentro de la existente de operario.
                                If Not blnFoundTO Then
                                    sngServOperario = Format(sngQtaServiziSpeciali * sngTariffaServiziSpeciali, "0.00") + sngServOperario
                                    blnFoundTO = True
                                End If
                                
                            Case "BP" ' OPerario/Bombeo
                                ' -- DP 07/01/09
                                If Not blnFoundBP Then
                                    
                                    ' JG 10/1/12 No se cobran los bombeos de los servicios que tienen menos de 3 entregas y no son pesados, si se cobran los que est·n en S3
                                    
                                    EntregasporDestino = VVC.Entregasdestinofecha(strIDDestinazione, strDataConsegna, strIDDeposito, Left(strRifOrdineEsterno, 10))
                                    
                                    lngHowManyB3 = 0
                                    If Not DA.QueryCommand("SELECT SXDNRALQ FROM  TARISXD0 WHERE (SXDIDADZ = ?) AND (SXDFLTSE = 'B3')", VarkbufferB3, lngHowManyB3, strIDDestinazione) Then
                                    End If
                                    
                                    If lngHowManyB3 > 0 Then
                                    If IsNumeric(VarkbufferB3(0, 0)) Then EntregasporDestino = VarkbufferB3(0, 0)
                                    End If
                                    
                                    If EntregasporDestino < 3 And strIDTipologiaProdotto <> "020" And strIDTipologiaProdotto <> "040" Then
                                    Else
                                    sngServBuquesPuerto = Format(sngQtaServiziSpeciali * sngTariffaServiziSpeciali, "0.00")
                                    End If

                                    blnFoundBP = True
                                End If
                                
                            Case "SE" ' Servicio especial
                                ' -- DP 07/01/09
                                If Not blnFoundSE Then
                                    sngServServiziSpeciali = Format(sngQtaServiziSpeciali * sngTariffaServiziSpeciali, "0.00")
                                    blnFoundSE = True
                                End If
                            
                            Case "SV"
                                ' -- JG Paralizaciones
                                If Not blnFoundSV Then
                                    sngServServiziSpeciali = Format(sngTariffaServiziSpeciali / lngNumScarichi, "0.00")
                                    blnFoundSV = True
                                End If
                                
                            Case "B3"
                                
                                ' Codigo de los puertos en los cuales se ha de cobrar el bombeo desde el primer suministro
                                
                                
                            Case Else
                                intRowErrorLog = InvalidCode
                                GoTo LogResult
                                
                        End Select
                    Next k
                End If
                
                With OCT
                   .DS.Fields("0").FindValue = strIDOrdine
                   .Find
                   sngOCTNota2 = .DS.Fields("15").value
                End With
              

            
                 
            ' JGA 20/1/20 Control zona Canarias si el viaje se hace en Canarias y la nota empieza por xxx cambiamos la tarifa a REC del mismo deposito
            Dim Zona As String
            Zona = ""
            If (strIDDeposito = "401" Or strIDDeposito = "402" Or strIDDeposito = "404" Or strIDDeposito = "405" Or strIDDeposito = "406" Or strIDDeposito = "407" Or strIDDeposito = "408" Or strIDDeposito = "409" Or strIDDeposito = "410") Then
            Zona = "CAN"
            End If
              
           If (strTipoTariffa = "A") And Zona = "CAN" And (strIDTipologiaProdotto <> "020") And (strIDTipologiaProdotto <> "040") And (Left(sngOCTNota2, 12) = "910101910101") Then
            strIDFasciaTariffaria = "REC"
            End If
               
                blnResult = TBR.CalcolaImporto(strIDDeposito, strIDFasciaTariffaria, _
                                   sngQta, sngTotQScaricataViaggio, sngTotQOrdinataViaggio, sngTotQScaricataViaggioMultiDestino, sngTotQScaricata, sngTotQOrdinataViaggioMultiDestino, sngQtaKg, sngTotQScaricataDestinazione, sngTotQOrdinata, lngKm, (lngRow = 0), _
                                   blnFlagKilometrica, blnFestivo, lngNumConsegne, blnConsegnaPi˘Distante, _
                                   blnUltimaConsegna, varTipoCisterna, strTipoTariffa, strIDTipologiaProdotto, strIDDestinazione, blnMultiDestino, strIDVettore, strFlagResponsabilit‡, sngOCTNota2, sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                If Not blnResult Then
                    intRowErrorLog = FasciaTariffariaError
                    GoTo LogResult
                End If
                
                ' control de Tarifa Lm reparto menor a 9500 Litros
                If strIDFasciaTariffaria = "LM" Then
                ' JG 30/01/12 buscamos la franja de menudeo dependiendo la sociedad
                Select Case strIDEnteEmittente
                    Case "624"
                        strIDFasciaTariffariaBusqueda = "LMT"
                    Case "004"
                        strIDFasciaTariffariaBusqueda = "LME"
                    Case Else
                        strIDFasciaTariffariaBusqueda = "LM"
                End Select
                
                              
                ' Nueva franja 2015 menudeo filiales
                    
                    If strIDDeposito = "004" Or strIDDeposito = "032" Or strIDDeposito = "031" Or strIDDeposito = "030" Or strIDDeposito = "033" Or strIDDeposito = "034" Then
                    strIDFasciaTariffariaBusqueda = "L15"
                    End If
                    If strIDDeposito = "035" Then
                    strIDFasciaTariffariaBusqueda = "L16"
                    End If
                    If strIDDeposito = "666" Or strIDDeposito = "121" Or strIDDeposito = "D01" Or strIDDeposito = "N0" Then
                    strIDFasciaTariffariaBusqueda = "CAT"
                    End If
                
                
                
                ' Jg Y Marc 26/10/2015 se le pasa el dato de cantidad inicial total que nadie le ha pasado nunca
                
                    With OCT
                        .DS.Fields("0").FindValue = strIDOrdine
                        .Find
                        .GetTotaliDestinazione
                    End With
                    sngTotQIniziale = OCT.TotaliQta(1)
                    
                    strIDZona = "00"
                    

                    blnResult = TFR.CalcolaImporto(strIDZona, strIDFasciaTariffariaBusqueda, _
                                       sngQta, sngQIniziale, sngTotQScaricata, _
                                       sngTotQIniziale, (lngKm / 2), (lngRow = 0), blnFlagProgrammato, blnFatturazioneAttiva, sngTariffa, _
                                       sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                    
                    If Not blnResult Then
                        intRowErrorLog = FasciaTariffariaError
                        GoTo LogResult
                    End If
                End If
                
                
                
                ' Calculo fueles canarias por Isla y por TN destinatarios islas
                    
                If Zona = "CAN" And (strIDTipologiaProdotto = "020") Then
                        
                    strIDZona = strIDDeposito
                    
                    strIDFasciaTariffariaBusqueda = TFA.FindByDate(strIDZona, strDataConsegna, strTipoTariffa, strIDDestinazione)
                    
                    
                  
                    blnResult = TFR.CalcolaImporto(strIDZona, strIDFasciaTariffariaBusqueda, _
                                       sngQta, sngQIniziale, sngTotQScaricata, _
                                       sngTotQIniziale, (lngKm / 2), (lngRow = 0), blnFlagProgrammato, blnFatturazioneAttiva, sngTariffa, _
                                       sngImporto, strIDValuta, strTipoPercorso, intAnomalia, sngQtaKg, Zona, strIDTipologiaProdotto)
                    
                    If Not blnResult Then
                        intRowErrorLog = FasciaTariffariaError
                        GoTo LogResult
                    End If
                 
                End If
                
                
                
                ' Control tarifa Ceuta
                
                If strIDFasciaTariffaria = "CE" Then
                strIDFasciaTariffariaBusqueda = "CE"
                
                With OCT
                        .DS.Fields("0").FindValue = strIDOrdine
                        .Find
                        .GetTotaliDestinazione
                    End With
                    sngTotQIniziale = OCT.TotaliQta(1)
                    
                    strIDZona = "00"
                  
                    blnResult = TFR.CalcolaImporto(strIDZona, strIDFasciaTariffariaBusqueda, _
                                       sngQta, sngQIniziale, sngTotQScaricata, _
                                       sngTotQIniziale, (lngKm / 2), (lngRow = 0), blnFlagProgrammato, blnFatturazioneAttiva, sngTariffa, _
                                       sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                    
                    If Not blnResult Then
                        intRowErrorLog = FasciaTariffariaError
                        GoTo LogResult
                    End If

                
                
                End If
                
            Case Else
                strIDFasciaTariffaria = TFA.FindByDate(strIDZona, strDataConsegna, strTipoTariffa, strIDDestinazione)
                If Len(strIDFasciaTariffaria) = 0 Then
                    intRowErrorLog = NoFasciaTariffaria
                    GoTo LogResult
                End If
        
                ' --- Leggo i totali delle quantit‡ scaricate
                ' --- per la destinazione
                If (strIDViaggio <> strIDViaggioPrec) Or (strIDConsegna <> strIDConsegnaPrec) Then
                    If (strIDConsegna <> strIDConsegnaPrec) Then
                        strIDConsegnaPrec = strIDConsegna
                    End If
                    If Len(DataClass.DS.ListItem(36, lngRow)) > 0 Then
                        blnFlagProgrammato = (DataClass.DS.ListItem(36, lngRow) = "S")
                    Else
                        blnFlagProgrammato = False
                    End If
                    With VVC
                        .DS.Fields("0").FindValue = strIDViaggio
                        .DS.Fields("1").FindValue = strIDConsegna
                        .Find
                    End With
                    
                    With OCT
                        .DS.Fields("0").FindValue = strIDOrdine
                        .Find
                        .GetTotaliDestinazione
                    End With
                    
                    VVC.GetTotaliDestinazione
                    If (Not blnFlagProgrammato) Then
                        sngTotQScaricata = VVC.TotaliQta(5)
                    Else
                        If blnFatturazioneAttiva Then
                            With OCT
                                If .TotaliQta(5) > 0 Then
                                    sngTotQScaricata = .TotaliQta(5)
                                Else
                                    sngTotQScaricata = .TotaliQta(1)
                                End If
                            End With
                        Else
                            sngTotQScaricata = VVC.TotaliQta(5)
                        End If
                    End If
                    sngTotQIniziale = OCT.TotaliQta(1)
                End If
                ' ---
                
        '        ' --- Uso il raggio della proposta di fattura anche per la valorizzazione
        '        ' --- interna
        '        lngKm = lngKmEsterni
        '        ' ---
                
                If Len(VVC.DS.Fields("39").value) = 0 Then
                    lngKm = 0
                Else
                    lngKm = VVC.DS.Fields("39").value
                End If
                    
                If Not blnFlagProgrammato Then
                    sngQta = sngQScaricata
                Else
                    If sngQIniziale > sngQScaricata Then
                        sngQta = sngQIniziale
                    Else
                        sngQta = sngQScaricata
                    End If
                End If
                
                blnResult = TFR.CalcolaImporto(strIDZona, strIDFasciaTariffaria, _
                                   sngQta, sngQIniziale, sngTotQScaricata, _
                                   sngTotQIniziale, lngKm, (lngRow = 0), blnFlagProgrammato, blnFatturazioneAttiva, sngTariffa, _
                                   sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                If Not blnResult Then
                    intRowErrorLog = FasciaTariffariaError
                    GoTo LogResult
                End If
        
        End Select
        
        If strIDValuta = vbNullString Then
            strIDValuta = strIDValutaDefault
        End If

        intRowError = RowToPPD(lngRow, lngIDDettaglio, PPD, strIDTestata, _
                               strIDZona, strIDFasciaTariffaria, _
                               sngTariffa, sngImporto, strIDValuta, _
                               strTipoPercorso, lngKm, strRifOrdineEsterno, _
                               sngQScaricata, intAnomalia, sngQScaricataKg, sngPedaggio, _
                               sngServOperario, sngServBuquesPuerto, sngServOreServiziSpeciali, sngServServiziSpeciali)
        If intRowError <> NoError Then
            intRowErrorLog = intRowError
            GoTo LogResult
        End If
        ' --- Fine valorizzazione interna
        
        
        ' --- Inserimento della riga di prefattura
        If Not PPD.Insert Then
            intRowError = InsertError
            intRowErrorLog = InsertError
            GoTo LogResult
        End If
        
        ' Inserisco la riga di addebito
        If sngImporto <> 0 Then
            PPA.DS.Fields.ResetFields
            With PPA.DS.Fields
                .Item("0").NewValue = PPD.DS.Fields("0").value
                .Item("1").NewValue = PPD.DS.Fields("1").value
                .Item("2").NewValue = DataClass.DS.ListItem(24, lngRow)
                
                Select Case StrIDSociet‡
                    Case "001", ID_SOC_GUARNER  ' Agip
                        .Item("3").NewValue = CONTRATTOID_AGIP
                    
                    Case "038"  ' Cepsa
                        .Item("3").NewValue = CONTRATTOID_CEPSA
                        
                        
                End Select
                    
                .Item("4").NewValue = "001"     ' Addebito
                .Item("5").NewValue = lngNumConsegne
                .Item("6").NewValue = lngKm
                .Item("7").NewValue = sngQta
                .Item("8").NewValue = sngTariffa
                .Item("9").NewValue = sngImporto
                
                .Item("0").Updated = True
                .Item("1").Updated = True
                .Item("2").Updated = True
                .Item("3").Updated = True
                .Item("4").Updated = True
                .Item("5").Updated = True
                .Item("6").Updated = True
                .Item("7").Updated = True
                .Item("8").Updated = True
                .Item("9").Updated = True
            End With
            
            If Not PPA.Insert Then
                intRowError = InsertError
                intRowErrorLog = InsertError
                GoTo LogResult
            End If
        End If
        
        If strTipoTariffa = "P" Then
            Select Case strIDVettore
                Case "A10"  'Pantoja
                    If (blnFestivo Or (Format(strDataConsegna, "w", vbSunday) = 1)) And blnUltimaConsegna Then
                        PPA.DS.Fields.ResetFields
                        With PPA.DS.Fields
                            .Item("0").NewValue = PPD.DS.Fields("0").value
                            .Item("1").NewValue = PPD.DS.Fields("1").value
                            .Item("2").NewValue = DataClass.DS.ListItem(24, lngRow)
                            
                            .Item("3").NewValue = CONTRATTOID_CEPSA
                                
                            .Item("4").NewValue = "013"     ' Addebito
                            .Item("5").NewValue = 0
                            .Item("6").NewValue = 0
                            .Item("7").NewValue = 1
                            .Item("8").NewValue = 36.05
                            .Item("9").NewValue = 36.05
                            
                            For i = 0 To 9
                                .Item(CStr(i)).Updated = True
                            Next i
                        End With
                        
                        If Not PPA.Insert Then
                            intRowError = InsertError
                            intRowErrorLog = InsertError
                            GoTo LogResult
                        End If
                    End If
                Case "A31"  'Direccion logistica
                    If strIDEnteEmittente = "004" Then
                        If (blnFestivo Or (Format(strDataConsegna, "w", vbSunday) = 1)) And blnUltimaConsegna Then
                            PPA.DS.Fields.ResetFields
                            With PPA.DS.Fields
                                .Item("0").NewValue = PPD.DS.Fields("0").value
                                .Item("1").NewValue = PPD.DS.Fields("1").value
                                .Item("2").NewValue = DataClass.DS.ListItem(24, lngRow)
                                
                                .Item("3").NewValue = CONTRATTOID_CEPSA
                                    
                                .Item("4").NewValue = "013"     ' Addebito
                                .Item("5").NewValue = 0
                                .Item("6").NewValue = 0
                                .Item("7").NewValue = 1
                                .Item("8").NewValue = 36
                                .Item("9").NewValue = 36
                                
                                For i = 0 To 9
                                    .Item(CStr(i)).Updated = True
                                Next i
                            End With
                            
                            If Not PPA.Insert Then
                                intRowError = InsertError
                                intRowErrorLog = InsertError
                                GoTo LogResult
                            End If
                        End If
                    End If
            End Select
        End If
    
'        If Not blnFatturazioneAttiva Then
'            Select Case strIDVettore
'
'            End Select
'        End If
        

'        If (strModalit‡Tariffa = "B") And (Not blnFlagKilometrica) And (Len(varTipoCisterna) > 0) And (Not blnFatturazioneAttiva) Then
'            If blnUltimaConsegna Then
'                sngQuantit‡Vuoto = 0
'                Select Case varTipoCisterna
'                    Case "001", "002", "003"
'                        If sngTotQScaricataViaggio <= 20010 Then
'                           sngQuantit‡Vuoto = 30000 - sngTotQScaricataViaggio
'                        End If
'                    Case "004"
'                        If sngTotQScaricataViaggio <= 15010 Then
'                           sngQuantit‡Vuoto = 22000 - sngTotQScaricataViaggio
'                        End If
'                End Select
'                If sngQuantit‡Vuoto > 0 Then
'                    sngQtaKg = Format(sngQuantit‡Vuoto * 0.7385, "0")
'                    blnResult = TBR.CalcolaImporto(strIDDeposito, strIDFasciaTariffaria, _
'                                       sngQuantit‡Vuoto, sngTotQScaricata, sngTotQScaricataConsegna, sngTotQScaricataConsegnaKG, sngTotQScaricataViaggio, sngQtaKg, sngTotQScaricataKG, sngTotQOrdinata, lngMaxKM, (lngRow = 0), _
'                                       blnFlagKilometrica, blnFestivo, lngNumConsegne, blnConsegnaPi˘Distante, _
'                                       blnUltimaConsegna, varTipoCisterna, strTipoTariffa, strIDTipologiaProdotto, strIDDestinazione, blnMultiDestino, sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
'                    If Not blnResult Then
'                        intRowErrorLog = FasciaTariffariaError
'                    End If
'                    lngIDDettaglio = lngIDDettaglio + 1
'                    intRowError = RowToPPD(lngRow, lngIDDettaglio, PPD, strIDTestata, _
'                                           strIDZona, strIDFasciaTariffaria, _
'                                           sngTariffa, sngImporto, strIDValuta, _
'                                           strTipoPercorso, lngMaxKM, strRifOrdineEsterno, _
'                                           sngQuantit‡Vuoto, intAnomalia, sngQtaKg)
'                    If intRowError <> NoError Then
'                        intRowErrorLog = intRowError
'                        GoTo LogResult
'                    End If
'                    If Not PPD.Insert Then
'                        intRowError = InsertError
'                        intRowErrorLog = InsertError
'                        GoTo LogResult
'                    End If
'                End If
'            End If
'        End If
        
    ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione delle consegne
        If intRowError = NoError Then
            If intRowErrorLog <> NoError Then
                lngCountFail = lngCountFail + 1
            Else
                lngCountOK = lngCountOK + 1
            End If
        Else
            blnViaggioOK = False
            lngCountFail = lngCountFail + 1
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        
        Select Case intRowError
            Case UnknownError
                strErrorMessage = "Errore sconosciuto"
            Case InsertError
                strErrorMessage = "Errore durante l'inserimento della riga di dettaglio"
            Case UpdateError
                strErrorMessage = "Errore durante l'aggiornamento della q.ta scaricata"
        End Select
        
        Select Case intRowErrorLog
            Case NoError
                strErrorMessage = "OK - Num. dettaglio " & lngIDDettaglio
            Case NoFasciaTariffaria
                strErrorMessage = "Fascia Tariffaria " & strIDFasciaTariffaria & " Zona " & strIDZona & " non trovata"
'            Case BadIDOrderSoc
'                strErrorMessage = "Riferimento esterno all'ordine mancante"
            Case FasciaTariffariaError
                strErrorMessage = "Errore nel calcolo della tariffa, Fascia " & strIDFasciaTariffaria & " Zona " & strIDZona
            Case NoPrefatturaEsterna
                strErrorMessage = "Proposta di fattura non trovata"
            Case InvalidDetail
                strErrorMessage = "Unit‡ di misura Servizi Speciali non valida"
            Case InvalidCode
                strErrorMessage = "Codice Servizi Speciali non valido"
            Case NoOrderID
                strErrorMessage = "Errore nella determinazione delle quantit‡"
        End Select
        
        lngIDDettaglio = lngIDDettaglio + 1
        XA.value(lngRow, COL_RESULT_LOG) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
        DoEvents
        ' ---
        
    Next lngRow
    ' ---
    
    If blnViaggioOK Then
        If blnFatturazioneAttiva Then
            VVT.DS.Fields("0").FindValue = strIDViaggio
            If VVT.Find Then
                VVT.SetStato "E"
                PPD.DS.DA.EndTrans ETCOMMIT
                lngCountViaggiOK = lngCountViaggiOK + 1
                lblLog(7) = lngCountViaggiOK
            Else
                PPD.DS.DA.EndTrans ETROLLBACK
            End If
        Else
            PPD.DS.DA.EndTrans ETCOMMIT
            lngCountViaggiOK = lngCountViaggiOK + 1
            lblLog(7) = lngCountViaggiOK
        End If
    Else
        PPD.DS.DA.EndTrans ETROLLBACK
    End If
    
'    ElaboraMinimo strIDTestata
    GoTo EndProc
    
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    PPD.DS.DA.EndTrans ETROLLBACK
    Select Case ErrCode
        Case InsertError
            strMessage = "Impossibile inserire una nuova prefattura."
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Impossibile trovare la valuta di default"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmPPTCreate", _
                "CreateRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    StatusBar1.Panels(2).Text = vbNullString
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
'            For i = 0 To lstZone.ListCount - 1
'                If lstZone.Selected(i) Then
'                    j = j + 1
'                    .Fields("17").FindValue(j) = AZO.DS.ListItem(0, i)
'                End If
'            Next i
            If Not .ScanClientCond Then
                GoTo ErrHandler
            End If
            If optTipoTariffa(0).value = True Then
                If Not DataClass.GetSelectionCreatePPD("A") Then
                    GoTo ErrHandler
                End If
            Else
                If Not DataClass.GetSelectionCreatePPD("P") Then
                    GoTo ErrHandler
                End If
            End If
            DisplaySelection
            GoTo EndProc
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

Private Function SelectPrefatture(IDTestata As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim j As Integer
    
    Me.MousePointer = vbHourglass
    With PDS.DS
        For i = 0 To lstZone.ListCount - 1
            If lstZone.Selected(i) Then
                j = j + 1
                .Fields("17").FindValue(j) = AZO.DS.ListItem(0, i)
            End If
        Next i
        .Fields("0").FindValue = IDTestata
        .Fields("5").FindValue = mskDateFrom
        .Fields("6").FindValue = mskDateTo
        .Fields("13").FindValue = MaskEdBox7
        If Not PDS.GetSelectionMinimo Then
            GoTo ErrHandler
        End If
        GoTo EndProc
    End With
    GoTo EndProc
ErrHandler:
    Me.MousePointer = vbDefault
    SelectPrefatture = False
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectPrefatture = True
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
    lblIDTestata = vbNullString
    If lngRows > 0 Then
        ReDim varLog(0 To 3, 0 To lngRows - 1)
        For i = 0 To lngRows - 1
            varLog(0, i) = DataClass.DS.ListItem(0, i)
            varLog(1, i) = DataClass.DS.ListItem(14, i)
            varLog(2, i) = DataClass.DS.ListItem(2, i)
        Next i
        RefreshXArray XA, tdbLog, varLog
        tdbLog.bookmark = 0
    End If
End Sub



Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Dalla data inizio carico"
        Case 1
            Set M = mskDateTo
            strTitle = "Alla data inizio carico"
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
    Set XA = New XArrayDB
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .Fields("3").Required = True
        .ConfigureInterface
        ConfigureObjects Me
        .Fields("4").Required = True
        .Fields("5").Required = True
        .Fields("6").Required = True
    End With
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    AZO.GetList True
    
    Set PPT = New CPrefattura
    With PPT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
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
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set TFA = New CFasciaTariffaria
    With TFA.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set TFR = New CFasciaTariffariaRiga
    With TFR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set TFB = New CTariffaBinomica
    With TFB.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set TBR = New CTariffaBinomicaRiga
    With TBR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set PPD = New CPrefatturaDettaglio
    With PPD.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set PPA = New CPrefatturaDettaglioAddebito
    With PPA.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    Set PDS = New CPrefatturaDettaglioSelezione
    With PDS.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ACN = New CContratto
    With ACN.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

'
'    mskField11 = "RP"
    
    With AZO.DS
        For i = 0 To UBound(.List, 2)
            lstZone.AddItem .ListItem(0, i) & " " & .ListItem(1, i)
            If UCase(.ListItem(3, i)) = UCase(AZO.DS.DA.ConnUserName) Then
                lstZone.Selected(lstZone.NewIndex) = True
            End If
        Next i
    End With
    
    If lstZone.ListCount > 0 Then
        lstZone.ListIndex = 0
    End If
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set GridArraySelect = Nothing
    Set AZO = Nothing
    Set ADZ = Nothing
    Set PPT = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    Set VVP = Nothing
    Set PPD = Nothing
    Set PDS = Nothing
    Set TFA = Nothing
    Set TFR = Nothing
    Set TFB = Nothing
    Set TBR = Nothing
    Set OCT = Nothing
    Set AUT = Nothing
    Set ACN = Nothing
    Set APR = Nothing
    Set DataClass = Nothing
    Set XA = Nothing
    Set frmPPTCreate = Nothing
End Sub




