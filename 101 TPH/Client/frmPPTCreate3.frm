VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPPTCreate3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAZIONE PREFATTURE"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   9270
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmPPTCreate3.frx":0000
      TabIndex        =   16
      Tag             =   "SELECT"
      Top             =   3720
      Width           =   5175
   End
   Begin TrueDBGrid60.TDBGrid tdbLog 
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "frmPPTCreate3.frx":35EA
      TabIndex        =   38
      Top             =   2520
      Width           =   9015
   End
   Begin MSMask.MaskEdBox mskDescrizione 
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   2160
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   35
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   2055
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9255
      Begin VB.OptionButton optTipoTariffa 
         Caption         =   "Valorizzazione passiva"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   40
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optTipoTariffa 
         Caption         =   "Valorizzazione attiva"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   39
         Top             =   1320
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3360
         Picture         =   "frmPPTCreate3.frx":6BDA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3360
         Picture         =   "frmPPTCreate3.frx":6CDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.ListBox lstZone 
         Height          =   960
         Left            =   5880
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   9
         Tag             =   "13"
         Top             =   1320
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
         TabIndex        =   11
         Tag             =   "92"
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskSocietà 
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
         Caption         =   "&Società"
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
         TabIndex        =   42
         Tag             =   "3"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   5400
         Tag             =   "92"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ra&ggruppamento"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   10
         Top             =   1680
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
         TabIndex        =   41
         Tag             =   "92"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "&Zone"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   14
         Tag             =   "13"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unità di Trasporto"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5400
         Tag             =   "13"
         Top             =   1320
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
      TabIndex        =   23
      Top             =   3600
      Width           =   3735
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   36
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Viaggi valorizzati:"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Scarichi elab. correttamente:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Scarichi da elaborare:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   5280
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
         Picture         =   "frmPPTCreate3.frx":6DDE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Esegui"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmPPTCreate3.frx":6EE0
         TabIndex        =   18
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
         Picture         =   "frmPPTCreate3.frx":6FE2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Consegne"
         Height          =   375
         Left            =   120
         Picture         =   "frmPPTCreate3.frx":70E4
         TabIndex        =   17
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
         Picture         =   "frmPPTCreate3.frx":71E6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmPPTCreate3.frx":72E8
         TabIndex        =   19
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6150
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
            TextSave        =   "16.02"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12/06/2006"
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
            Picture         =   "frmPPTCreate3.frx":73EA
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPPTCreate3.frx":74FC
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPPTCreate3.frx":760E
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
      TabIndex        =   37
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Num. Valorizzazione"
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   34
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Descrizione"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "frmPPTCreate3"
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
Private ACD As CContrattoDettaglio
Private ADZ As CDestinazione
Private APR As CProdotto
Private PPT As CPrefattura
Private PPD As CPrefatturaDettaglio
Private PPA As CPrefatturaDettaglioAddebito
Private TFC As CTariffaContratto
Private TCR As CTariffaContrattoRiga
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private PDS As CPrefatturaDettaglioSelezione
Private OCT As COrdine
Private AUT As CUnitaDiTrasporto



Private Sub ElaboraMinimo(IDTestata As String)
    On Error GoTo ErrHandler
        
    Dim lngNumRows As Long
    Dim lngUltimaRiga As Long
    Dim lngCountOK As Long
    Dim lngCounTFCil As Long
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
'        strIDFasciaTariffaria = TFC.FindByDate(strIDZona, strDataViaggio, "A")
        
        With TFC.DS
            .Fields("0").FindValue = strIDZona
            .Fields("1").FindValue = strIDFasciaTariffaria
            If Not TFC.Find Then
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
            lngCounTFCil = lngCounTFCil + 1
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
                          IDContratto As String, _
                          IDFasciaTariffaria As String, _
                          TariffaI As Single, _
                          ImportoI As Single, _
                          IDValutaI As String, _
                          TipoPercorsoI As String, _
                          KmRaggioInterni As Long, _
                          RiferimentoEsterno As String, _
                          QtaScaricata As Single, _
                          AnomaliaI As AnomaliePrefattura, _
                          QtaScaricataKg As Single) As Errors
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
        .Fields("43").NewValue = TariffaI
        .Fields("44").NewValue = ImportoI
        .Fields("45").NewValue = ""
        .Fields("46").NewValue = IDValutaI
        .Fields("47").NewValue = KmRaggioInterni 'varRow(33, 0)
        .Fields("48").NewValue = TipoPercorsoI
        .Fields("50").NewValue = .Fields("50").NewValue Or AnomaliaI
        .Fields("63").NewValue = IDContratto
        .Fields("63").Updated = True
        .Fields("64").NewValue = IDFasciaTariffaria
        .Fields("64").Updated = True
        ' ---
        
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
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngIDDettaglio As Long
    Dim lngCountOK As Long
    Dim lngCounTFCil As Long
    Dim lngCountToGo As Long
    Dim lngCountViaggiOK As Long
    Dim lngKM As Long
'    Dim lngMaxKM As Long
    Dim lngKmEsterni As Long
    Dim lngKmDefViaggio As Long
    Dim lngNumConsegne As Long
    Dim lngNumConsegneViaggio As Long
    Dim lngNumConsAddizionali As Long
    Dim lngCountDettContratto As Long
    Dim j As Long
    Dim i As Integer
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
    Dim strErrDesc As String
    Dim strIDValuta As String
    Dim strIDValutaDefault As String
    Dim strIDFasciaTariffaria As String
    Dim strIDDestinazione As String
    Dim strIDDestinazionePrec As String
    Dim strIDCliente As String
    Dim strIDDeposito As String
    Dim StrIDSocietà As String
    Dim strTipoPercorso As String
    Dim strIDProdotto As String
    Dim strBollettaScarico As String
    Dim strIDUdT As String
    Dim strTipoTariffa As String
    Dim strModalitàTariffa As String
    Dim strIDAddebito As String
    Dim strParametriTassazione As String
    Dim strIDTipoCalcolo As String
    Dim varIDContratto As String
    Dim varDensitàProdotto As Variant
    Dim varTipoCisterna As Variant
    Dim varBufferDistanze As Variant
    Dim varDettaglioContratto As Variant
    Dim varMinimoTassabile As Variant
    Dim blnFound As Boolean
    Dim blnPropostaMancante As Boolean
    Dim blnResult As Boolean
    Dim blnViaggioOK As Boolean
    Dim blnConcurErr As Boolean
    Dim blnFlagProgrammato As Boolean
    Dim blnFatturazioneAttiva As Boolean
    Dim blnUltimaConsegna As Boolean
    Dim blnConsegnaPiùDistante As Boolean
    Dim blnFlagKilometrica As Boolean
    Dim blnFestivo As Boolean
    Dim blnPompa As Boolean
    Dim blnCalcolaConsAddizionali As Boolean
    Dim sngQScaricata As Single
    Dim sngQScaricataKg As Single
    Dim sngQCaricataKg As Single
    Dim sngQCaricataL As Single
    Dim sngQIniziale As Single
'    Dim sngTotQScaricata As Single
    Dim sngTotQScaricataViaggio As Single
    Dim sngTotQScaricataKG As Single
    Dim sngTotQIniziale As Single
    Dim sngQta As Single
    Dim sngQtaKg As Single
    Dim sngTariffa As Single
    Dim sngImporto As Single
    Dim sngQuantitàVuoto As Single
    Dim intRowError As Errors
    Dim intRowErrorLog As Errors
    Dim intAnomalia As AnomaliePrefattura
    Dim ErrCode As Errors
    Dim VCE As CViaggioEvento
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    
    StatusBar1.Panels(2).Text = "Valorizzazione delle consegne in corso... "
    
    
'    Set VCE = New CViaggioEvento
'    With VCE.DS
'        Set .DA = DataClass.DS.DA
'        Set .ErrH = DataClass.DS.ErrH
'    End With
    
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
        .Fields("5").NewValue = mskSocietà
        .Fields("5").Updated = True
        .Fields("6").NewValue = "G" ' (PM) 08/03/06 - Imposta lo stato prefattura a Generata
        .Fields("6").Updated = True
    End With
    
    If Not PPT.Insert Then
        ErrCode = InsertError
        GoTo ErrHandler
    End If
    strIDTestata = PPT.DS.Fields("0").value
    lblIDTestata = strIDTestata
    ' ---
    
    blnViaggioOK = True
    blnCalcolaConsAddizionali = True
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
'            strIDViaggioPrec = strIDViaggio
            strIDConsegnaPrec = vbNullString
            sngTotQScaricataKG = 0
            PPD.DS.DA.BeginTransaction
        End If
        ' ---
        
        intRowError = NoError
        intRowErrorLog = NoError
        intAnomalia = apNoAnomalie
        strErrDesc = vbNullString
        
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
'            lngKM = .ListItem(33, lngRow)
            
            strIDCliente = .ListItem(24, lngRow)
            strIDDestinazione = .ListItem(25, lngRow)
            ADZ.DS.Field("Codice").FindValue = strIDDestinazione
            ADZ.Find
            If ADZ.DS.HowMany = 0 Then
                GoTo ErrHandler
            End If
            
            strIDUdT = .ListItem(34, lngRow)
            strIDProdotto = .ListItem(2, lngRow)
            strRifOrdineEsterno = IIf(IsNull(.ListItem(8, lngRow)), "", .ListItem(8, lngRow))
            strIDDeposito = .ListItem(23, lngRow)
            varBufferDistanze = ADZ.GetDistanza(strIDDestinazione, strIDDeposito)
            If Not IsEmpty(varBufferDistanze) Then
                lngKM = varBufferDistanze(0, 0)
            Else
                lngKM = 0
            End If
            
            StrIDSocietà = .ListItem(21, lngRow)
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
            If IsNull(.ListItem(43, lngRow)) Then
                varIDContratto = ""
            Else
                varIDContratto = .ListItem(43, lngRow)
            End If
            If Len(varIDContratto) = 0 Then
                intRowError = UnknownError
                intRowErrorLog = UnknownError
                strErrDesc = "Contratto mancante su destinazione " & strIDDestinazione
                GoTo LogResult
            End If
            ACD.DS.Fields("0").FindValue = strIDCliente
            ACD.DS.Fields("1").FindValue = varIDContratto
            If Not ACD.GetList(True) Then
                intRowError = UnknownError
                intRowErrorLog = UnknownError
                strErrDesc = "Errore nella lettura del dettaglio del contratto " & strIDCliente & " " & varIDContratto
                GoTo LogResult
            End If
            If ACD.DS.HowMany = 0 Then
                intRowError = UnknownError
                intRowErrorLog = UnknownError
                strErrDesc = "Nessun dettaglio definito per il contratto " & strIDCliente & " " & varIDContratto
                GoTo LogResult
            End If
            varDettaglioContratto = ACD.DS.List
        End With
        
        
        ' Elaboro tutti gli addebiti delle righe di contratto
        
        PPD.DS.Fields.ResetFields
        
        strModalitàTariffa = vbNullString
        
        sngQta = sngQScaricata
        
        
        ' Controllo se sto iniziando ad elaborare un viaggio o una consegna
        If (strIDViaggio <> strIDViaggioPrec) Or (strIDConsegna <> strIDConsegnaPrec) Then
            
            ' Controllo se sto iniziando l'elaborazione di un viaggio
            If (strIDViaggio <> strIDViaggioPrec) Then
                strIDViaggioPrec = strIDViaggio
                strIDConsegnaPrec = vbNullString
                strIDDestinazionePrec = vbNullString
                
                VVT.DS.Fields("0").FindValue = strIDViaggio
                If Not VVT.Find Then
                    GoTo ErrHandler
                End If
                
                ' Determino se il viaggio è stato fatto in un giorno festivo
                blnFestivo = (VVT.DS.Fields("38").value = "F") Or (VVT.DS.Fields("38").value = "FB")
                
                ' Determino se il per viaggio è stata utilizzata la pompa/contalitri
                blnPompa = (VVT.DS.Fields("38").value = "B") Or (VVT.DS.Fields("38").value = "FB")
                
                blnCalcolaConsAddizionali = True
                
                ' Calcolo il totale del prodotto scaricato per il viaggio
                VVT.GetTotali
                sngTotQScaricataViaggio = VVT.TotaliQta(5)
                
                ' Determino il numero di consegne del viaggio
                lngNumConsegneViaggio = VVT.ContaDestinazioni
                If lngNumConsegneViaggio = -1 Then
                    GoTo ErrHandler
                End If
                lngNumConsegne = lngNumConsegneViaggio
            End If
            
        
            ' Controllo se sto iniziando l'elaborazione di una consegna
            If (strIDConsegna <> strIDConsegnaPrec) Then
                strIDConsegnaPrec = strIDConsegna
            
                ' Controllo se ho già elaborato una consegna per la stessa destinazione
                ' Attenzione: le consegne devono essere in ordine di destinazione
                If strIDDestinazione <> strIDDestinazionePrec Then
                    strIDDestinazionePrec = strIDDestinazione
                    
'                    ' Determino la quantità scaricata per le consegne per la destinazione
'                    With VVC
'                        .DS.Fields("0").FindValue = strIDViaggio
'                        .DS.Fields("1").FindValue = strIDConsegna
'                        .Find
'                    End With
'                    VVC.GetTotaliDestinazione
'                    sngTotQScaricata = VVC.TotaliQta(5)
                    
                    ' Controllo se la consegna è la più distante nel viaggio
                    If lngNumConsegneViaggio > 1 Then
                        blnConsegnaPiùDistante = VVC.IsDestinazionePiùLontana2(strIDViaggio, strIDConsegna, strIDDeposito)
                    Else
                        blnConsegnaPiùDistante = True
                    End If
                    If blnConsegnaPiùDistante Then
                        ' Se è la consegna più distante fatturo tutto il quantitativo del viaggio con i km della consegna
                        sngQta = sngTotQScaricataViaggio
                    Else
                        ' Se non è la consegna più distante, la ignoro
                        sngQta = 0
                    End If
                Else
                    ' Se ho già elaborato una consegna per questa destinazione, ignoro la consegna
                    sngQta = 0
                End If
            End If
            
        Else
            ' Ignoro le righe di prodotto successive alla prima
            sngQta = 0
        End If
        
        ' Elaboro gli addebiti previsti dal contratto per questa consegna
'        If sngQta <> 0 Then
            For lngCountDettContratto = 0 To UBound(varDettaglioContratto, 2)
                strIDAddebito = varDettaglioContratto(2, lngCountDettContratto)
                If Not IsNull(varDettaglioContratto(4, lngCountDettContratto)) Then
                    varMinimoTassabile = varDettaglioContratto(4, lngCountDettContratto)
                Else
                    varMinimoTassabile = 0
                End If
                If Not IsNull(varDettaglioContratto(5, lngCountDettContratto)) Then
                    strParametriTassazione = varDettaglioContratto(5, lngCountDettContratto)
                Else
                    strParametriTassazione = vbNullString
                End If
                strIDTipoCalcolo = varDettaglioContratto(6, lngCountDettContratto)
                
                ' Determino la fascia tariffaria in vigore al momento della consegna
                strIDFasciaTariffaria = TFC.FindByDate(strIDCliente, varIDContratto, strIDAddebito, strDataConsegna, strParametriTassazione, strIDDeposito, strIDDestinazione)
                If Len(strIDFasciaTariffaria) = 0 Then
                    If (strIDAddebito = "001") Then
                        intRowError = NoFasciaTariffaria
                        intRowErrorLog = NoFasciaTariffaria
                        GoTo LogResult
                    Else
                        intRowErrorLog = NoError
                    End If
                    GoTo NextDettContratto
                End If
                
                sngImporto = 0
                ' Elaboro i vari addebiti definiti nel contratto
                Select Case strIDAddebito
                    ' Trasporto
                    Case "001"
                        
                        If IsNumeric(varMinimoTassabile) And (sngQta <> 0) Then
                            If sngQta < (varMinimoTassabile * 1000) Then
                                sngQta = varMinimoTassabile * 1000
                            End If
                        End If
                        
                        Select Case strIDTipoCalcolo
                            Case "03", "98"
                                If lngKM = 0 Then
                                    intRowError = UnknownError
                                    intRowErrorLog = UnknownError
                                    strErrDesc = "Km mancanti, Dest.: " & strIDDestinazione & ", Dep.: " & strIDDeposito
                                    GoTo LogResult
                                End If
                        End Select
                        
                        blnResult = TCR.CalcolaImporto(strIDCliente, varIDContratto, strIDAddebito, strIDFasciaTariffaria, _
                                                        sngQta, sngQtaKg, lngKM, (lngRow = 0), _
                                                        blnFlagKilometrica, blnFestivo, lngNumConsegne, blnConsegnaPiùDistante, _
                                                        blnUltimaConsegna, blnPompa, strTipoTariffa, strIDTipoCalcolo, _
                                                        sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                        If Not blnResult Then
                            intRowError = FasciaTariffariaError
                            intRowErrorLog = FasciaTariffariaError
                            GoTo LogResult
                        End If
                
                
                
                        If strIDValuta = vbNullString Then
                            strIDValuta = strIDValutaDefault
                        End If
                
                        intRowError = RowToPPD(lngRow, lngIDDettaglio, PPD, strIDTestata, _
                                               strIDZona, varIDContratto, strIDFasciaTariffaria, _
                                               sngTariffa, sngImporto, strIDValuta, _
                                               strTipoPercorso, lngKM, strRifOrdineEsterno, _
                                               sngQScaricata, intAnomalia, sngQScaricataKg)
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
                ' Autopista
                Case "002"
                    blnResult = TCR.CalcolaImporto(strIDCliente, varIDContratto, strIDAddebito, strIDFasciaTariffaria, _
                                                    sngQta, sngQtaKg, lngKM, (lngRow = 0), _
                                                    blnFlagKilometrica, blnFestivo, lngNumConsegne, blnConsegnaPiùDistante, _
                                                    blnUltimaConsegna, blnPompa, strTipoTariffa, strIDTipoCalcolo, _
                                                    sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                    If Not blnResult Then
                        intRowError = FasciaTariffariaError
                        intRowErrorLog = FasciaTariffariaError
                        GoTo LogResult
                    End If
                    sngQta = 0
                    lngKM = 0
                    lngNumConsegne = 0
                ' Uso pompa
                Case "005"
                    If blnPompa Then
                        blnResult = TCR.CalcolaImporto(strIDCliente, varIDContratto, strIDAddebito, strIDFasciaTariffaria, _
                                                        sngTotQScaricataViaggio, sngQtaKg, lngKM, (lngRow = 0), _
                                                        blnFlagKilometrica, blnFestivo, lngNumConsegne, blnConsegnaPiùDistante, _
                                                        blnUltimaConsegna, blnPompa, strTipoTariffa, strIDTipoCalcolo, _
                                                        sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                        If Not blnResult Then
                            intRowError = FasciaTariffariaError
                            intRowErrorLog = FasciaTariffariaError
                            GoTo LogResult
                        End If
                        sngQta = 0
                        lngKM = 0
                        lngNumConsegne = 0
                        blnPompa = False
                    End If
                ' Consegne addizionali
                Case "012"
                    If blnCalcolaConsAddizionali Then
                        lngNumConsAddizionali = VVT.ContaDestinazioni - 1
                        If lngNumConsAddizionali > 0 Then
                            blnResult = TCR.CalcolaImporto(strIDCliente, varIDContratto, strIDAddebito, strIDFasciaTariffaria, _
                                                            sngTotQScaricataViaggio, sngQtaKg, lngKM, (lngRow = 0), _
                                                            blnFlagKilometrica, blnFestivo, lngNumConsAddizionali, blnConsegnaPiùDistante, _
                                                            blnUltimaConsegna, blnPompa, strTipoTariffa, strIDTipoCalcolo, _
                                                            sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                            If Not blnResult Then
                                intRowError = FasciaTariffariaError
                                intRowErrorLog = FasciaTariffariaError
                                GoTo LogResult
                            End If
                            sngQta = 0
                            lngKM = 0
    '                        sngImporto = sngImporto * lngNumConsAddizionali
                            lngNumConsegne = 0
                            blnCalcolaConsAddizionali = False
                        End If
                    End If
                ' Giorno festivo
                Case "013"
                    If blnFestivo Then
                        blnResult = TCR.CalcolaImporto(strIDCliente, varIDContratto, strIDAddebito, strIDFasciaTariffaria, _
                                                        sngTotQScaricataViaggio, sngQtaKg, lngKM, (lngRow = 0), _
                                                        blnFlagKilometrica, blnFestivo, lngNumConsegne, blnConsegnaPiùDistante, _
                                                        blnUltimaConsegna, blnPompa, strTipoTariffa, strIDTipoCalcolo, _
                                                        sngTariffa, sngImporto, strIDValuta, strTipoPercorso, intAnomalia)
                        If Not blnResult Then
                            intRowError = FasciaTariffariaError
                            intRowErrorLog = FasciaTariffariaError
                            GoTo LogResult
                        End If
                        sngQta = 0
                        lngKM = 0
                        lngNumConsegne = 0
                        blnFestivo = False
                    End If
            End Select

            ' Inserisco la riga di addebito
            If sngImporto <> 0 Then
                PPA.DS.Fields.ResetFields
                With PPA.DS.Fields
                    .Item("0").NewValue = PPD.DS.Fields("0").value
                    .Item("1").NewValue = PPD.DS.Fields("1").value
                    .Item("2").NewValue = strIDCliente
                    .Item("3").NewValue = varIDContratto
                    .Item("4").NewValue = strIDAddebito
                    .Item("5").NewValue = lngNumConsegne
                    .Item("6").NewValue = lngKM
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

NextDettContratto:

        Next lngCountDettContratto
'    End If
    ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione delle consegne
        If intRowError = NoError Then
            If intRowErrorLog <> NoError Then
                lngCounTFCil = lngCounTFCil + 1
            Else
                lngCountOK = lngCountOK + 1
            End If
        Else
            blnViaggioOK = False
            lngCounTFCil = lngCounTFCil + 1
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCounTFCil
        lblLog(5) = lngCountToGo
        
        Select Case intRowError
            Case UnknownError
                If strErrDesc <> vbNullString Then
                    strErrorMessage = strErrDesc
                Else
                    strErrorMessage = "Errore sconosciuto"
                End If
            Case InsertError
                strErrorMessage = "Errore durante l'inserimento della riga di dettaglio"
            Case UpdateError
                strErrorMessage = "Errore durante l'aggiornamento della q.ta scaricata"
            Case NoFasciaTariffaria
                strErrorMessage = "Fascia Tariffaria " & strIDFasciaTariffaria & " non trovata"
            Case FasciaTariffariaError
                strErrorMessage = "Errore nel calcolo della tariffa, Fascia " & strIDFasciaTariffaria
        End Select
        
        Select Case intRowErrorLog
            Case NoError
                strErrorMessage = "OK - Num. dettaglio " & lngIDDettaglio
            Case NoFasciaTariffaria
                strErrorMessage = "Fascia Tariffaria " & strIDFasciaTariffaria & " non trovata"
            Case FasciaTariffariaError
                strErrorMessage = "Errore nel calcolo della tariffa, Fascia " & strIDFasciaTariffaria
            Case NoPrefatturaEsterna
                strErrorMessage = "Proposta di fattura non trovata"
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
            For i = 0 To lstZone.ListCount - 1
                If lstZone.Selected(i) Then
                    j = j + 1
                    .Fields("17").FindValue(j) = AZO.DS.ListItem(0, i)
                End If
            Next i
            If Not .ScanClientCond Then
                GoTo ErrHandler
            End If
            If optTipoTariffa(0).value = True Then
                If Not DataClass.GetSelectionCreatePPDContratto Then
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
    
    Set TFC = New CTariffaContratto
    With TFC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set TCR = New CTariffaContrattoRiga
    With TCR.DS
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
    
    Set ACD = New CContrattoDettaglio
    With ACD.DS
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
    Set TFC = Nothing
    Set TCR = Nothing
    Set OCT = Nothing
    Set AUT = Nothing
    Set ACN = Nothing
    Set APR = Nothing
    Set DataClass = Nothing
    Set XA = Nothing
    Set frmPPTCreate3 = Nothing
End Sub




