VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRASMISSIONE VIAGGI"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   8940
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   24
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
         Picture         =   "frmVVTExport.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmVVTExport.frx":0102
         TabIndex        =   30
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
         Picture         =   "frmVVTExport.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Viaggi"
         Height          =   375
         Left            =   2280
         Picture         =   "frmVVTExport.frx":0306
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
         Index           =   0
         Left            =   240
         Picture         =   "frmVVTExport.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Seleziona File..."
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTExport.frx":050A
         TabIndex        =   34
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
         Picture         =   "frmVVTExport.frx":060C
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
         Left            =   7080
         Picture         =   "frmVVTExport.frx":070E
         TabIndex        =   26
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   4920
      TabIndex        =   17
      Top             =   3600
      Width           =   3975
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Viaggi da trasmettere:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Viaggi trasmessi correttamente:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   1815
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3360
         Picture         =   "frmVVTExport.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7320
         Picture         =   "frmVVTExport.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Tag             =   "5"
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
         Left            =   6000
         TabIndex        =   4
         Tag             =   "6"
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
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
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
         TabIndex        =   9
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Tag             =   "4"
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCodOperatore 
         Height          =   255
         Left            =   7440
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
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
               Picture         =   "frmVVTExport.frx":0A14
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTExport.frx":0B26
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTExport.frx":0C38
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskPresidio 
         Height          =   255
         Left            =   6840
         TabIndex        =   40
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "&Presidio:"
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   41
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblNumInvio 
         Caption         =   "Codice &Operatore:"
         Height          =   255
         Left            =   6000
         TabIndex        =   39
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emit."
         Height          =   255
         Index           =   16
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
         Width           =   735
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
         Left            =   2640
         TabIndex        =   36
         Tag             =   "4"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   5160
         Tag             =   "4"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio Def."
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   8
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
         TabIndex        =   6
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3840
         Tag             =   "5"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data Inizio Viaggio"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   7800
         Tag             =   "6"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File:"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   0
      TabIndex        =   12
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
         Picture         =   "frmVVTExport.frx":0D4A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Viaggio"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTExport.frx":0E4C
         TabIndex        =   14
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmVVTExport.frx":0F4E
         TabIndex        =   15
         Top             =   240
         Width           =   4575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6525
      Width           =   8940
      _ExtentX        =   15769
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
            Object.Width           =   10319
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15.05"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "28/05/2008"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmVVTExport.frx":3956
      TabIndex        =   28
      Tag             =   "SELECT"
      Top             =   1920
      Width           =   8655
   End
End
Attribute VB_Name = "frmVVTExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 19 ' 4

Private Const COL_ID_UNITA_TRASPORTO = 1
Private Const COL_NUM_SEQUENZA = 2
Private Const COL_HR_VIAGGIO = 3
Private Const COL_ID_DEPOSITO = 4
Private Const COL_ID_ENTE_EMITTENTE = 5
Private Const COL_DT_VIAGGIO = 6
Private Const COL_ID_VIAGGIO_DEF = 7
Private Const COL_ID_VETTORE = 8
Private Const COL_ID_DEPOSITO_FV = 9
Private Const COL_ID_ENTE_FV = 10
Private Const COL_KM_PREVISTI = 11
Private Const COL_KM_DEFINITIVI = 12
Private Const COL_ID_MOTRICE = 13
Private Const COL_TIPO_MOTRICE = 14
Private Const COL_ID_RIMORCHIO = 15
Private Const COL_TIPO_RIMORCHIO = 16
Private Const COL_ORA_INIZIO_CARICO = 17
Private Const COL_ID_AUTISTA1 = 18
Private Const COL_NOME_AUTISTA1 = 19
Private Const COL_ID_AUTISTA2 = 20
Private Const COL_NOME_AUTISTA2 = 21
Private Const COL_CONTRATTO = 22
Private Const COL_CONTRATTO_SEC = 23
Private Const COL_PAT_AUTISTA1 = 24
Private Const COL_PAT_AUTISTA2 = 25
Private Const COL_ID_MAGAZZINO = 26

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjOutFile As CTextFile
Private mobjOutFile200 As CTextFile
Private mobjOutFile201 As CTextFile
Private mobjOutFile202 As CTextFile
Private mobjOutFile210 As CTextFile
Private mobjOutFile220 As CTextFile
Private mobjOutFile240 As CTextFile

Private mobjOutFile400 As CTextFile ' Trasmissione Viaggi Testata Tamoil
Private mobjOutFile401 As CTextFile ' Trasmissione Viaggi Dettagli Tamoil

Private mstrProgrInvio As String

Public GridArraySelect As ArrayGridCls
Public XA As XArrayDB
Public Loading As Boolean

Public DataClass As CViaggioSelezione
Public State As FormState

#If VersionID = 2 Then
    Public Enum ExportTypes
        etSAP = 0
        etDECO = 1
    End Enum
    Public ExportType As ExportTypes
#End If

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
            #If VersionID = 1 Then
                ExportRecords001
            #ElseIf VersionID = 2 Then
                Select Case ExportType
                    Case etSAP
                        ExportRecords002
                    Case etDECO
                        ExportRecords005
                End Select
            #ElseIf VersionID = 4 Then
                ExportRecords004
'            #ElseIf VersionID = 5 Then
'                ExportRecords005
            #End If
            cmdExport.Enabled = False
        Case fcOpen
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per la trasmissione... "
                .InitDir = App.Path
                .Filter = "Tutti i files (*.*)|*.*| "
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
    Dim intProgressivoViaggio As Integer
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
    Dim strBadge As String
    Dim strCentroOperativo As String
    Dim strUnit‡Trasporto As String
    Dim strOraInizioViaggio As String
    Dim strFileName As String
    Dim strOraConsegna As String
    Dim strRifAGIP As String
    Dim strStatoOrdine As String
    Dim strTipoVendita As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim strDummy As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim ErrCode As Errors
    Dim C As CError
    
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
        strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(ID_SOC)
    End With
    ' ---
    
    strCodOperatore = Format$(mskCodOperatore, "00")
    mobjOutFile.Delimiter = "|"
    
    If Len(mstrProgrInvio) = 0 Then
        With DataClass.DS.DA
'            .BeginTransaction
'            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "004", 0)
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "1" & strCodOperatore, 0)
            If varCounter > 999 Then
                varCounter = 1
            End If
'            .UpdateCounter "anuprogr", varCounter
'            .EndTrans etCommit
        End With
        mstrProgrInvio = Format$(varCounter, "000")
    End If
    
    If Len(mskFileName) > 0 Then
        strFileName = mskFileName
    Else
        strFileName = strIDEnteEmittente & "ser"
        mskFileName = strFileName
        mskFileName.Refresh
    End If
    
    If Not mobjOutFile.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    strDataInvio = Format$(Date, "dd/mm/yy")
    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00")
    strHeader = "1|" & strIDEnteEmittente & "|" & strCodOperatore & "|" & mstrProgrInvio & "|" & strDataInvio & "|" & strOraInvio & "|" & Space(42) & "|"
    
    Print #mobjOutFile.FileNum, strHeader
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione dei viaggi
    For lngRow = 0 To lngNumRows
'        intProgressivoViaggio = intProgressivoViaggio + 1
        strIDZonaPrec = vbNullString
        intRowError = NoError
        
        With ADE.DS
            strDummy = GridArraySelect.value(lngRow, 4) 'ID Deposito
            .Fields("0").FindValue = strDummy
            If Not ADE.Find Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
            strIDDeposito = .MultiSoc.DS.GetFirmCode(ID_SOC)
            strCentroOperativo = .Fields("8").value
        End With
        
        With AUT.DS
            strDummy = GridArraySelect.value(lngRow, 1) 'ID Unit‡ di trasporto
            .Fields("0").FindValue = strDummy
            If Not AUT.Find Then
                intRowError = noUT
                GoTo LogResult
            End If
            strUnit‡Trasporto = .MultiSoc.DS.GetFirmCode(ID_SOC)
            strBadge = .Fields("14").value
        End With
        
        strDataViaggio = Format$(GridArraySelect.value(lngRow, 6), "dd/mm/yy") 'Data viaggio
        strOraInizioViaggio = GridArraySelect.value(lngRow, 3) 'Ora inizio viaggio
        With mobjOutFile
            .ClearFields
            .Item("1").value = "2"
            .Item("2").value = strDataViaggio ' strDataInvio
            .Item("3").value = strUnit‡Trasporto
            .Item("4").value = strBadge
'            .Item("5").value = Format$(intProgressivoViaggio, "00")
            .Item("6").value = Format$(GridArraySelect.value(lngRow, 2), "00") 'Sequenza viaggio
            .Item("8").value = Format$(Hour(strOraInizioViaggio), "00") & ":" & Format$(Minute(strOraInizioViaggio), "00")
            .Item("11").value = strIDDeposito 'Deposito
'            .Item("12").value = strCentroOperativo
        End With
        
        strIDViaggio = GridArraySelect.value(lngRow, 0) 'Codice viaggio
        
        With VVC.DS
            .Fields(.KeyFields(1)).FindValue = strIDViaggio
            
            If Not VVC.GetSequentialList Then
            'If Not VVC.GetList(True) Then
                intRowError = RowsNotFound
                GoTo LogResult
            End If
            
            intRowsTrip = .HowMany - 1
        End With
        
        intCountSospesi = 0
        For i = 0 To intRowsTrip
            With OCT
                strStatoOrdine = vbNullString
                .DS.Fields("0").FindValue = VVC.DS.ListItem(2, i)
                If Not .Find Then
                    intRowError = NoOrderID
                    GoTo LogResult
                End If
                strRifAGIP = .DS.Fields("10").value
                strStatoOrdine = .DS.Fields("22").value
                strTipoVendita = UCase(.DS.Fields("16").value)
                strIDZona = .DS.Fields("18").value
                With AZO.DS
                    .Fields("Codice").FindValue = strIDZona
                    If Not AZO.Find Then
                        intRowError = NoDestinazione
                        GoTo LogResult
                    End If
                    strCentroOperativo = .Fields("5").value
                End With
            End With
            
            ' --- Se gli ordini sono sospesi ("bloccati") o il tipo vendita Ë "FD"
            ' --- non esportiamo la consegna
            If (strStatoOrdine <> "S") And (strTipoVendita <> "FD") Then
            
                strOraConsegna = VVC.DS.ListItem(5, i)
                With ADZ.DS
                    strIDDestinazione = CStr(VVC.DS.ListItem(6, i))
                    .Fields("0").FindValue = strIDDestinazione
                    If Not ADZ.Find Then
                        intRowError = NoDestinazione
                        GoTo LogResult
                    End If
                    strIDDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC)
'                    strIDZona = .Fields("16").value
                End With
                
                With ATV.DS
                    strIDTipoVendita = CStr(VVC.DS.ListItem(7, i))
                    .Fields(.KeyFields(1)).FindValue = strIDTipoVendita
                    If Not ATV.Find Then
                        intRowError = NoTipoVendita
                        GoTo LogResult
                    End If
                    strIDTipoVendita = .MultiSoc.DS.GetFirmCode(ID_SOC)
                End With
                
                With mobjOutFile
    '                .Item("7").value = VVC.DS.ListItem(1, i)
'                    .Item("7").value = Format$(i + 1, "00")
                    intProgressivoConsegna = 1
                    For j = 0 To intRowsTrip
                        If VVC.DS.ListItem(1, j) < VVC.DS.ListItem(1, i) Then
                            intProgressivoConsegna = intProgressivoConsegna + 1
                        End If
                    Next j
                    intProgressivoConsegna = intProgressivoConsegna - intCountSospesi
                    If strIDZonaPrec <> strIDZona Then
                        strIDZonaPrec = strIDZona
                        intProgressivoViaggio = intProgressivoViaggio + 1
                    End If
                    .Item("5").value = Format$(intProgressivoViaggio, "00")
                    .Item("7").value = Format$(intProgressivoConsegna, "00")
                    .Item("9").value = Format$(Hour(strOraConsegna), "00") & ":" & Format$(Minute(strOraConsegna), "00")
                    .Item("10").value = strIDZona
                    .Item("12").value = strCentroOperativo
                    .Item("13").value = strRifAGIP
                    .Item("14").value = strIDDestinazione
                    .Item("15").value = strIDTipoVendita
                    If Not .WriteLine(True) Then
                        intRowError = OutputError
                        GoTo LogResult
                    End If
                End With
'                intProgressivoConsegna = intProgressivoConsegna + 1
            Else
                intCountSospesi = intCountSospesi + 1
            End If
        Next i
        
        
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
        ' ---
        
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
                strErrorMessage = "Unit‡ di trasporto non trovata"
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
    mobjOutFile.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub

End Sub


Private Sub ExportRecords002()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngQtaLt As Long
    Dim intRowsTrip As Integer
    Dim intRowsConsegna As Integer
    Dim intKmPrevisti As Integer
    Dim intKmDefinitivi As Integer
    Dim intKmDefinitiviConsegna As Integer
    Dim intNumSequenza As Integer
    Dim i As Long
    Dim j As Long
    Dim intFileNum As Integer
    Dim intProgressivoConsegna As Integer
    Dim intCountSIF As Integer
    Dim strIDParam As String
    Dim strIDViaggio As String
    Dim strIDViaggioDef As String
    Dim strIDConsegna As String
    Dim strIDEnteEmittente As String
    Dim strIDDeposito As String
    Dim strIDCliente As String
    Dim strIDDestinazione As String
    Dim strIDVettore As String
    Dim strIDFineViaggio As String
    Dim strIDUT As String
    Dim strIDMotrice As String
    Dim strIDRimorchio As String
    Dim strIDTipoMotrice As String
    Dim strIDTipoRimorchio As String
    Dim strIDTipoAutomezzo As String
    Dim strIDAutista1 As String
    Dim strIDAutista2 As String
    Dim strNomeAutista1 As String
    Dim strNomeAutista2 As String
    Dim strContratto As String
    Dim strContrattoSecondario As String
    Dim strIDProdotto As String
    Dim strIDPosizioneDoganale As String
    Dim strDataInizioViaggio As String
    Dim strOraInvio As String
    Dim strUnit‡Trasporto As String
    Dim strOraInizioViaggio As String
    Dim strOraInizioCarico As String
    Dim strDataRicezioneOrdine As String
    Dim strOraRicezioneOrdine As String
    Dim strDataInserimentoOrdine As String
    Dim strOraInserimentoOrdine As String
    Dim strTipoOrdine As String
    Dim strFileName As String
    Dim strOraConsegna As String
    Dim strRifSoc As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim strDummy As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    strFileName = mskFileName
    
    If Not mobjOutFile200.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    intFileNum = mobjOutFile200.FileNum
    mobjOutFile201.FileNum = intFileNum
    mobjOutFile202.FileNum = intFileNum
    mobjOutFile210.FileNum = intFileNum
    mobjOutFile220.FileNum = intFileNum
    mobjOutFile240.FileNum = intFileNum
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione degli ordini
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        
        With AEE.DS
            strDummy = GridArraySelect.value(lngRow, COL_ID_ENTE_EMITTENTE)
            .Fields(.KeyFields(1)).FindValue = strDummy
            If Not AEE.Find Then
                ErrCode = NoEnteEmittente
                GoTo LogResult
            End If
            strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(ID_SOC)
        End With
        
        With ADE.DS
            strDummy = GridArraySelect.value(lngRow, COL_ID_DEPOSITO) 'ID Deposito
            .Fields("0").FindValue = strDummy
            If Not ADE.Find Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
            strIDDeposito = .MultiSoc.DS.GetFirmCode(ID_SOC)
        End With
        
        With AVE.DS
            strDummy = GridArraySelect.value(lngRow, COL_ID_VETTORE)
            .Fields("0").FindValue = strDummy
            If Not AVE.Find Then
                intRowError = noUT
                GoTo LogResult
            End If
            strIDVettore = .MultiSoc.DS.GetFirmCode(ID_SOC)
        End With
        
        With GridArraySelect
            intNumSequenza = .value(lngRow, COL_NUM_SEQUENZA)
            strUnit‡Trasporto = .value(lngRow, COL_ID_UNITA_TRASPORTO)
            strOraInizioViaggio = .value(lngRow, COL_HR_VIAGGIO)
            strOraInizioViaggio = Format$(Hour(strOraInizioViaggio), "00") & _
                                  Format$(Minute(strOraInizioViaggio), "00")
            strDataInizioViaggio = .value(lngRow, COL_DT_VIAGGIO)
            strIDViaggioDef = .value(lngRow, COL_ID_VIAGGIO_DEF)
            strIDMotrice = .value(lngRow, COL_ID_MOTRICE)
            strIDRimorchio = .value(lngRow, COL_ID_RIMORCHIO)
            strIDTipoMotrice = .value(lngRow, COL_TIPO_MOTRICE)
            If Len(.value(lngRow, COL_TIPO_RIMORCHIO)) > 0 Then
                strIDTipoRimorchio = .value(lngRow, COL_TIPO_RIMORCHIO)
            End If
            strOraInizioCarico = .value(lngRow, COL_ORA_INIZIO_CARICO)
            strOraInizioCarico = Format$(Hour(strOraInizioCarico), "00") & _
                                 Format$(Minute(strOraInizioCarico), "00")
            If Len(strIDTipoRimorchio) > 0 Then
                strDummy = strIDTipoRimorchio
            Else
                strDummy = strIDTipoMotrice
            End If
        End With
        With ATM.DS
            .Fields("Codice").FindValue = strDummy
            If Not ATM.Find Then
                intRowError = NoTipoAutomezzo
                GoTo LogResult
            End If
            strIDTipoAutomezzo = .MultiSoc.DS.GetFirmCode(ID_SOC)
        End With
        
'        If Len(GridArraySelect.value(lngRow, COL_ID_DEPOSITO_FV)) > 0 Then
'            With ADE.DS
'                strDummy = GridArraySelect.value(lngRow, COL_ID_DEPOSITO_FV) 'ID Deposito
'                .Fields("0").FindValue = strDummy
'                If Not ADE.Find Then
'                    intRowError = NoDeposito
'                    GoTo LogResult
'                End If
'                strIDFineViaggio = .MultiSoc.DS.GetFirmCode(ID_SOC)
'                strIDFineViaggio = "DU" & strIDFineViaggio
'            End With
'        Else
'            If Len(GridArraySelect.value(lngRow, COL_ID_ENTE_FV)) > 0 Then
'                With AEE.DS
'                    strDummy = GridArraySelect.value(lngRow, COL_ID_ENTE_FV)
'                    .Fields(.KeyFields(1)).FindValue = strIDEnteEmittente
'                    If Not AEE.Find Then
'                        ErrCode = NoEnteEmittente
'                        GoTo ErrHandler
'                    End If
'                    strIDFineViaggio = .MultiSoc.DS.GetFirmCode(ID_SOC)
'                    strIDFineViaggio = "PL" & strIDFineViaggio
'                End With
'            End If
'        End If
        strIDFineViaggio = "PL" & strIDEnteEmittente
        intKmPrevisti = GridArraySelect.value(lngRow, COL_KM_PREVISTI)
        intKmDefinitivi = GridArraySelect.value(lngRow, COL_KM_DEFINITIVI)
        strIDAutista1 = GridArraySelect.value(lngRow, COL_ID_AUTISTA1)
        strNomeAutista1 = GridArraySelect.value(lngRow, COL_NOME_AUTISTA1)
        strIDAutista2 = GridArraySelect.value(lngRow, COL_ID_AUTISTA2)
        
        If Len(GridArraySelect.value(lngRow, COL_NOME_AUTISTA2)) > 0 Then
            strNomeAutista2 = GridArraySelect.value(lngRow, COL_NOME_AUTISTA2)
        Else
            strNomeAutista2 = ""
        End If
        
        If Len(GridArraySelect.value(lngRow, COL_CONTRATTO)) > 0 Then
            strContratto = GridArraySelect.value(lngRow, COL_CONTRATTO)
        Else
            strContratto = ""
        End If
        
        If Len(GridArraySelect.value(lngRow, COL_CONTRATTO_SEC)) > 0 Then
            strContrattoSecondario = GridArraySelect.value(lngRow, COL_CONTRATTO_SEC)
        Else
            strContrattoSecondario = ""
        End If
        
        strIDViaggio = GridArraySelect.value(lngRow, 0) 'Codice viaggio
        With VVT
            .DS.Fields("0").FindValue = strIDViaggio
            If Not .Find Then
                intRowError = NoTripID
                GoTo LogResult
            End If
            intCountSIF = .ContaProdottiSIF("02")
        End With
        
        With VVC.DS
            .Fields(.KeyFields(1)).FindValue = strIDViaggio
            
            If Not VVC.GetList(True) Then
                intRowError = RowsNotFound
                GoTo LogResult
            End If
            
            intRowsTrip = .HowMany - 1
        End With
        
        With mobjOutFile200
            .ClearFields
            .Item("2").value = strIDEnteEmittente
            strDataInizioViaggio = Format$(strDataInizioViaggio, "yyyy") & _
                                   Format$(Month(strDataInizioViaggio), "00") & _
                                   Format$(Day(strDataInizioViaggio), "00")
            .Item("3").value = strDataInizioViaggio
            .Item("4").value = strIDViaggioDef
            .Item("5").value = strContratto
            .Item("6").value = strIDVettore
            .Item("7").value = strIDDeposito
            .Item("9").value = IIf(intCountSIF > 0, "S", "N")
            .Item("10").value = strContrattoSecondario
            .Item("14").value = intRowsTrip + 1
            .Item("15").value = .Item("14").value
            .Item("16").value = strIDFineViaggio
            .Item("18").value = intKmPrevisti
            .Item("20").value = intKmDefinitivi
            .Item("23").value = strIDEnteEmittente
            If Not .WriteLine(True) Then
                intRowError = OutputError
                GoTo LogResult
            End If
        End With
        
        With mobjOutFile201
            .Item("2").value = strIDEnteEmittente
            .Item("3").value = strDataInizioViaggio
            .Item("4").value = strIDViaggioDef
            .Item("5").value = strIDMotrice
            .Item("6").value = strIDRimorchio
            .Item("7").value = strIDTipoAutomezzo
            .Item("8").value = strNomeAutista1
            .Item("9").value = strOraInizioCarico
            .Item("10").value = strOraInizioViaggio
            .Item("11").value = intNumSequenza
            .Item("12").value = strIDEnteEmittente

            If Not .WriteLine(True) Then
                intRowError = OutputError
                GoTo LogResult
            End If
        End With
        
        With mobjOutFile202
            .Item("2").value = strIDEnteEmittente
            .Item("3").value = strDataInizioViaggio
            .Item("4").value = strIDViaggioDef
            .Item("5").value = strNomeAutista2
            .Item("7").value = "UE" & strIDEnteEmittente
'            .Item("8").value = strIDAutista1
'            .Item("9").value = strIDAutista2
            ' --- Provvisorio 13/09/99: Il codice autista puÚ essere pi˘
            ' ---                       lungo di 3
            If Len(strIDAutista1) > 0 Then
                .Item("8").value = "1" & strIDAutista1
            End If
            If Len(strIDAutista2) > 0 Then
                .Item("9").value = "1" & strIDAutista2
            End If
            ' ---
            
            .Item("11").value = strUnit‡Trasporto
            .Item("12").value = strIDEnteEmittente

            If Not .WriteLine(True) Then
                intRowError = OutputError
                GoTo LogResult
            End If
        End With
        
        
        ' --- Elaborazione delle consegne
        
        For i = 0 To intRowsTrip
            With OCT
                .DS.Fields("0").FindValue = VVC.DS.ListItem(2, i)
                If Not .Find Then
                    intRowError = NoOrderID
                    GoTo LogResult
                End If
                
                If Len(.DS.Fields("10").value) > 0 Then
                    ' --- Ordini da Call Center
                    strRifSoc = .DS.Fields("10").value
                    strTipoOrdine = .DS.Fields("11").value
                Else
                    ' --- Ordini manuali
                    With .DS.DA
'                        .BeginTransaction
                        varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "004", 0)
'                        .UpdateCounter "anuprogr", varCounter
'                        .EndTrans etCommit
                    End With
                    strRifSoc = varCounter
                    strTipoOrdine = "CC"
                End If
                strDataRicezioneOrdine = .DS.Fields("3").value
                strDataRicezioneOrdine = Format$(strDataRicezioneOrdine, "yyyy") & _
                                         Format$(Month(strDataRicezioneOrdine), "00") & _
                                         Format$(Day(strDataRicezioneOrdine), "00")
                strOraRicezioneOrdine = .DS.Fields("4").value
                strOraRicezioneOrdine = Format$(Hour(strOraRicezioneOrdine), "00") & _
                                        Format$(Minute(strOraRicezioneOrdine), "00")
                strDataInserimentoOrdine = .DS.Fields("5").value
                strDataInserimentoOrdine = Format$(strDataInserimentoOrdine, "yyyy") & _
                                         Format$(Month(strDataInserimentoOrdine), "00") & _
                                         Format$(Day(strDataInserimentoOrdine), "00")
                If strDataInserimentoOrdine = "19000101" Then
                    strDataInserimentoOrdine = vbNullString
                End If
                strOraInserimentoOrdine = .DS.Fields("6").value
                strOraInserimentoOrdine = Format$(Hour(strOraInserimentoOrdine), "00") & _
                                        Format$(Minute(strOraInserimentoOrdine), "00")
                If strDataRicezioneOrdine = "19000101" Then
                    strDataRicezioneOrdine = strDataInserimentoOrdine
                    strOraRicezioneOrdine = strOraInserimentoOrdine
                End If
            End With
            strOraConsegna = VVC.DS.ListItem(5, i)
            strOraConsegna = Format$(Hour(strOraConsegna), "00") & _
                             Format$(Minute(strOraConsegna), "00")
            intProgressivoConsegna = VVC.DS.ListItem(1, i)
'            strIDCliente = VVC.DS.ListItem(8, i)
            intKmDefinitiviConsegna = VVC.DS.ListItem(9, i)
'            With ACL.DS
'                .Fields("0").FindValue = strIDCliente
'                If Not ACL.Find Then
'                    intRowError = NoCliente
'                    GoTo LogResult
'                End If
'                strIDCliente = .MultiSoc.DS.GetFirmCode(ID_SOC)
'            End With
            With ADZ.DS
                strIDDestinazione = CStr(VVC.DS.ListItem(6, i))
                .Fields("0").FindValue = strIDDestinazione
                If Not ADZ.Find Then
                    intRowError = NoDestinazione
                    GoTo LogResult
                End If
                strIDDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC)
            End With
            
            With mobjOutFile210
                .Item("2").value = strIDEnteEmittente
                .Item("3").value = strDataInizioViaggio
                .Item("4").value = strIDViaggioDef
                .Item("5").value = intProgressivoConsegna
                .Item("6").value = strIDDestinazione
'                .Item("8").value = strIDCliente
                .Item("11").value = strIDEnteEmittente
    
                If Not .WriteLine(True) Then
                    intRowError = OutputError
                    GoTo LogResult
                End If
            End With
            With mobjOutFile240
                .Item("2").value = strIDEnteEmittente
                .Item("3").value = strDataInizioViaggio
                .Item("4").value = strIDViaggioDef
                .Item("5").value = intProgressivoConsegna
                .Item("6").value = strTipoOrdine
                .Item("7").value = strRifSoc
                .Item("8").value = intKmDefinitiviConsegna
                .Item("9").value = strDataRicezioneOrdine
                .Item("10").value = strOraRicezioneOrdine
                .Item("11").value = strDataInserimentoOrdine
                .Item("12").value = strOraInserimentoOrdine
                .Item("13").value = strOraConsegna
                .Item("15").value = strIDEnteEmittente
    
            End With
            
            ' --- Elaborazione dei prodotti consegnati
            strIDConsegna = VVC.DS.ListItem(0, i)
            With VVP.DS
                .Fields(.KeyFields(1)).FindValue = strIDViaggio
                .Fields(.KeyFields(2)).FindValue = strIDConsegna
                
                If Not VVP.GetList(True) Then
                    intRowError = RowsNotFound
                    GoTo LogResult
                End If
                
                intRowsConsegna = .HowMany - 1
            End With
            
            For j = 0 To intRowsConsegna
                strIDProdotto = VVP.DS.ListItem(1, j)
                strIDPosizioneDoganale = VVP.DS.ListItem(7, j) '(8, j)
                lngQtaLt = VVP.DS.ListItem(3, j)
                
                With APR.DS
                    .Fields("0").FindValue = strIDProdotto
                    If Not APR.Find Then
                        intRowError = NoProdotto
                        GoTo LogResult
                    End If
                    strIDProdotto = .MultiSoc.DS.GetFirmCode(ID_SOC, 3)
                End With
                With APD.DS
                    .Fields("Codice").FindValue = strIDPosizioneDoganale
                    If Not APD.Find Then
                        intRowError = NoPosizioneDoganale
                        GoTo LogResult
                    End If
                    strIDPosizioneDoganale = .MultiSoc.DS.GetFirmCode(ID_SOC)
                End With
                With mobjOutFile220
                    .Item("2").value = strIDEnteEmittente
                    .Item("3").value = strDataInizioViaggio
                    .Item("4").value = strIDViaggioDef
                    .Item("5").value = intProgressivoConsegna
                    .Item("6").value = strIDProdotto
                    .Item("7").value = strIDPosizioneDoganale
                    .Item("9").value = lngQtaLt
                    .Item("12").value = strIDEnteEmittente
        
                    If Not .WriteLine(True) Then
                        intRowError = OutputError
                        GoTo LogResult
                    End If
                End With
            Next j
            ' ---
            If Not mobjOutFile240.WriteLine(True) Then
                intRowError = OutputError
                GoTo LogResult
            End If

        Next i
        ' ---
        ' --- Modifico lo stato del viaggio
        If Not VVT.SetFlagTrasmissione("T") Then
            intRowError = NoSetState
            GoTo LogResult
        End If
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
'        lblLog(7) = lngCountExcluded
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case OutputError
                strErrorMessage = "Errore durante la scrittura sul file"
            Case noUT
                strErrorMessage = "Unit‡ di trasporto non trovata"
            Case RowsNotFound
                strErrorMessage = "Errore durante la lettura delle consegne"
            Case NoSetState
                strErrorMessage = "Impossibile cambiare lo stato del viaggio"
            Case NoOrderID
                strErrorMessage = "Ordine non trovato"
            Case NoTripID
                strErrorMessage = "Viaggio non trovato"
            Case NoTipoAutomezzo
                strErrorMessage = "Codice Tipo Automezzo non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Ente Emittente non trovato."
            Case NoCliente
                strErrorMessage = "Codice Cliente non trovato."
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato."
            Case NoPosizioneDoganale
                strErrorMessage = "Codice Posizione Doganale non trovato."
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
    Set C = ErrH.SetError(ErrCode, "frmVVTExportMover", _
                "ExportRecords002", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFile200.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub

End Sub
Private Sub ExportRecords005()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngQtaLt As Long
    Dim lngTotQtaLt As Long
    Dim intRowsTrip As Integer
    Dim intRowsConsegna As Integer
    Dim intTotRowsConsegna As Integer
    Dim intTotRowsTrip As Integer
    Dim intNumSequenza As Integer
    Dim i As Long
    Dim j As Long
    Dim intFileNum As Integer
    Dim intProgressivoConsegna As Integer
    Dim intProgressivoProdotto As Integer
    Dim strIDParam As String
    Dim strIDViaggio As String
    Dim strIDViaggioDef As String
    Dim strIDConsegna As String
    Dim strIDEnteEmittente As String
    Dim strIDDeposito As String
    Dim strIDMagazzino As String
    Dim strDescMagazzino As String
    Dim strIDCliente As String
    Dim strDescCliente As String
    Dim strIndiCliente As String
    Dim strLocaCliente As String
    Dim strCCAPCliente As String
    Dim strProvCliente As String
    Dim strPIVACliente As String
    Dim strIDDestinazione As String
    Dim strDescDest As String
    Dim strIndiDest As String
    Dim strLocaDest As String
    Dim strCCAPDest As String
    Dim strProvDest As String
    Dim strPIVADest As String
    Dim strIDVettore As String
    Dim strIDFineViaggio As String
    Dim strIDUT As String
    Dim strIDMotrice As String
    Dim strIDRimorchio As String
    Dim strIDTipoMotrice As String
    Dim strIDTipoRimorchio As String
    Dim strIDTipoAutomezzo As String
    Dim strIDAutista1 As String
    Dim strIDAutista2 As String
    Dim strNomeAutista1 As String
    Dim strNomeAutista2 As String
    Dim strPatAutista1 As String
    Dim strPatAutista2 As String
    Dim strIDProdotto As String
    Dim strIDPosizioneDoganale As String
    Dim strDataInizioViaggio As String
    Dim strOraInvio As String
    Dim strUnit‡Trasporto As String
    Dim strOraInizioCarico As String
    Dim strFileName As String
    Dim strErrorMessage As String
    Dim strDummy As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    strFileName = mskFileName
    
    If Not mobjOutFile200.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    intFileNum = mobjOutFile200.FileNum
    mobjOutFile201.FileNum = intFileNum
    mobjOutFile202.FileNum = intFileNum
    mobjOutFile210.FileNum = intFileNum
    mobjOutFile220.FileNum = intFileNum
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione dei viaggi
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        intTotRowsTrip = 0
        intTotRowsConsegna = 0
        lngTotQtaLt = 0
        
        With GridArraySelect
            strDataInizioViaggio = .value(lngRow, COL_DT_VIAGGIO)
            strIDViaggioDef = .value(lngRow, COL_ID_VIAGGIO_DEF)
            strIDMotrice = .value(lngRow, COL_ID_MOTRICE)
            strIDRimorchio = .value(lngRow, COL_ID_RIMORCHIO)
            strIDDeposito = GridArraySelect.value(lngRow, COL_ID_DEPOSITO)
            strIDMagazzino = .value(lngRow, COL_ID_MAGAZZINO)
            strOraInizioCarico = .value(lngRow, COL_ORA_INIZIO_CARICO)
            strOraInizioCarico = Format$(Hour(strOraInizioCarico), "00") & _
                                 Format$(Minute(strOraInizioCarico), "00") & _
                                 "00"
            strNomeAutista1 = .value(lngRow, COL_NOME_AUTISTA1)
            strPatAutista1 = .value(lngRow, COL_PAT_AUTISTA1)
            If Len(GridArraySelect.value(lngRow, COL_NOME_AUTISTA2)) > 0 Then
                strNomeAutista2 = GridArraySelect.value(lngRow, COL_NOME_AUTISTA2)
                strPatAutista2 = .value(lngRow, COL_PAT_AUTISTA2)
            Else
                strNomeAutista2 = ""
                strPatAutista2 = ""
            End If
            strIDViaggio = GridArraySelect.value(lngRow, 0) 'Codice viaggio
        End With
        
        With AVE.DS
            strDummy = GridArraySelect.value(lngRow, COL_ID_VETTORE)
            .Fields("0").FindValue = strDummy
            If Not AVE.Find Then
                intRowError = noUT
                GoTo LogResult
            End If
            strIDVettore = .MultiSoc.DS.GetFirmCode(ID_SOC)
        End With
        
'        With AEE.DS
'            strDummy = GridArraySelect.value(lngRow, COL_ID_ENTE_EMITTENTE)
'            .Fields(.KeyFields(1)).FindValue = strDummy
'            If Not AEE.Find Then
'                ErrCode = NoEnteEmittente
'                GoTo LogResult
'            End If
'            strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(ID_SOC)
'        End With
'
'        With ADE.DS
'            strDummy = GridArraySelect.value(lngRow, COL_ID_DEPOSITO) 'ID Deposito
'            .Fields("0").FindValue = strDummy
'            If Not ADE.Find Then
'                intRowError = NoDeposito
'                GoTo LogResult
'            End If
'            strIDDeposito = .MultiSoc.DS.GetFirmCode(ID_SOC)
'        End With
        
        With VVT
            .DS.Fields("0").FindValue = strIDViaggio
            If Not .Find Then
                intRowError = NoTripID
                GoTo LogResult
            End If
        End With
        
        With VVC.DS
            .Fields(.KeyFields(1)).FindValue = strIDViaggio
            
            If Not VVC.GetList(True) Then
                intRowError = RowsNotFound
                GoTo LogResult
            End If
            
            intRowsTrip = .HowMany - 1
        End With
        
        With mobjOutFile200
            .ClearFields
            strDataInizioViaggio = Format$(strDataInizioViaggio, "yyyy") & _
                                   Format$(Month(strDataInizioViaggio), "00") & _
                                   Format$(Day(strDataInizioViaggio), "00")
            .Item("2").value = strDataInizioViaggio
            .Item("4").value = strIDViaggioDef
            .Item("6").value = lngRow + 1
'            .Item("8").value = strIDEnteEmittente
'            .Item("9").value = strIDDeposito
            .Item("10").value = strIDVettore
            .Item("11").value = strPatAutista1
            .Item("12").value = strPatAutista2
            .Item("13").value = strIDMotrice
            .Item("14").value = strIDRimorchio
            .Item("15").value = strOraInizioCarico
            .Item("16").value = strNomeAutista1
            .Item("17").value = strNomeAutista2
            .Item("19").value = "I"
            If Not .WriteLine(True) Then
                intRowError = OutputError
                GoTo LogResult
            End If
            intTotRowsTrip = intTotRowsTrip + 1
        End With
        
        ' --- Elaborazione delle consegne
        For i = 0 To intRowsTrip
            intProgressivoConsegna = i + 1 'VVC.DS.ListItem(1, i)
            strIDCliente = VVC.DS.ListItem(8, i)
            With ACL.DS
                .Fields("0").FindValue = strIDCliente
                If Not ACL.Find Then
                    intRowError = NoCliente
                    GoTo LogResult
                End If
                strIDCliente = .MultiSoc.DS.GetFirmCode(ID_SOC)
                strDescCliente = .Fields("1").value
                strIndiCliente = .Fields("4").value
                strLocaCliente = .Fields("5").value
                strCCAPCliente = .Fields("7").value
                strProvCliente = .Fields("6").value
                strPIVACliente = .Fields("11").value
            End With
            With ADZ.DS
                strIDDestinazione = CStr(VVC.DS.ListItem(6, i))
                .Fields("0").FindValue = strIDDestinazione
                If Not ADZ.Find Then
                    intRowError = NoDestinazione
                    GoTo LogResult
                End If
                strIDDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC)
                strDescDest = .Fields("1").value
                strIndiDest = .Fields("4").value
                strLocaDest = .Fields("5").value
                strCCAPDest = .Fields("7").value
                strProvDest = .Fields("6").value
                strPIVADest = .Fields("10").value
            End With
            
            With mobjOutFile201
                .Item("2").value = strDataInizioViaggio
                .Item("4").value = strIDViaggioDef
                .Item("5").value = intProgressivoConsegna
                .Item("6").value = strIDCliente
                .Item("7").value = strDescCliente
                .Item("8").value = strIndiCliente
                .Item("9").value = strLocaCliente
                .Item("10").value = strCCAPCliente
                .Item("11").value = strProvCliente
                .Item("12").value = strPIVACliente
    
                If Not .WriteLine(True) Then
                    intRowError = OutputError
                    GoTo LogResult
                End If
                intTotRowsTrip = intTotRowsTrip + 1
            End With
            
            With mobjOutFile202
                .Item("2").value = strDataInizioViaggio
                .Item("4").value = strIDViaggioDef
                .Item("5").value = intProgressivoConsegna
                .Item("6").value = strIDDestinazione
                .Item("7").value = strDescDest
                .Item("8").value = strIndiDest
                .Item("9").value = strLocaDest
                .Item("10").value = strCCAPDest
                .Item("11").value = strProvDest
                .Item("12").value = strPIVADest
                
                If Not .WriteLine(True) Then
                    intRowError = OutputError
                    GoTo LogResult
                End If
                intTotRowsTrip = intTotRowsTrip + 1
                intTotRowsConsegna = intTotRowsConsegna + 1
            End With
            
            ' --- Elaborazione dei prodotti consegnati
            strIDConsegna = VVC.DS.ListItem(0, i)
            With VVP.DS
                .Fields(.KeyFields(1)).FindValue = strIDViaggio
                .Fields(.KeyFields(2)).FindValue = strIDConsegna
                
                If Not VVP.GetList(True) Then
                    intRowError = RowsNotFound
                    GoTo LogResult
                End If
                
                intRowsConsegna = .HowMany - 1
            End With
            
            For j = 0 To intRowsConsegna
                intProgressivoProdotto = j + 1
                strIDProdotto = VVP.DS.ListItem(1, j)
                lngQtaLt = VVP.DS.ListItem(3, j)
                With APR.DS
                    .Fields("0").FindValue = strIDProdotto
                    If Not APR.Find Then
                        intRowError = NoProdotto
                        GoTo LogResult
                    End If
                    strIDProdotto = .MultiSoc.DS.GetFirmCode(ID_SOC, 3)
                End With
                With mobjOutFile220
                    .Item("2").value = strDataInizioViaggio
                    .Item("4").value = strIDViaggioDef
                    .Item("5").value = intProgressivoConsegna
                    .Item("6").value = intProgressivoProdotto
                    .Item("7").value = strIDProdotto & "000"
                    .Item("9").value = lngQtaLt '* 1000
        
                    If Not .WriteLine(True) Then
                        intRowError = OutputError
                        GoTo LogResult
                    End If
                    intTotRowsTrip = intTotRowsTrip + 1
                    lngTotQtaLt = lngTotQtaLt + lngQtaLt
                End With
            Next j
            ' --- Fine elaborazione prodotti
        Next i
        ' --- Fine elaborazione consegne
        
        ' --- Scrivo il riepilogo del viaggio
        With mobjOutFile210
            .Item("2").value = strDataInizioViaggio
            .Item("4").value = strIDViaggioDef
            .Item("5").value = lngTotQtaLt '* 1000
            .Item("6").value = intTotRowsConsegna
            .Item("7").value = intTotRowsTrip + 1
            If Not .WriteLine(True) Then
                intRowError = OutputError
                GoTo LogResult
            End If
        End With
        ' ---
        
        ' --- Modifico lo stato del viaggio
        If Not VVT.SetFlagTrasmissione("T") Then
            intRowError = NoSetState
            GoTo LogResult
        End If
        ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione del viaggio
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            lngCountFail = lngCountFail + 1
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
            Case noUT
                strErrorMessage = "Vettore non trovato"
            Case RowsNotFound
                strErrorMessage = "Errore durante la lettura delle consegne"
            Case NoSetState
                strErrorMessage = "Impossibile cambiare lo stato del viaggio"
            Case NoOrderID
                strErrorMessage = "Ordine non trovato"
            Case NoTripID
                strErrorMessage = "Viaggio non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoMagazzino
                strErrorMessage = "Codice Magazzino di Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Ente Emittente non trovato."
            Case NoCliente
                strErrorMessage = "Codice Cliente non trovato."
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato."
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
    Set C = ErrH.SetError(ErrCode, "frmVVTExportMover", _
                "ExportRecords002", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFile200.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub

End Sub

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            If Not DataClass.GetSelectionExport(mskPresidio) Then
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
    
    Loading = True
    Set GridArraySelect = New ArrayGridCls
    Set XA = New XArrayDB
    
'    Select Case gmstrCodInstallazione
'        Case ciInst1
    #If VersionID = 1 Then
            cmdBrowse.Enabled = False
            
            imgState(16).Visible = False ' Ente Emittente
            lblField4.Visible = False    ' Ente Emittente
            Label1(16).Visible = False   ' Ente Emittente
            MaskEdBox1.Visible = False   ' Ente Emittente
            
            mskFileName = gmRegistry.GetSetting("Options", "TripExportFileAGIP", , True)
            mskPresidio = DataConnect.ConnUserName
'            mskCodOperatore = gmRegistry.GetSetting("Options", "OperatorCode", , True)
            Select Case UCase(DataConnect.ConnUserName)
                Case "LOG01"
                    mskCodOperatore = "01"
                Case "LOG02"
                    mskCodOperatore = "03"
                Case "LOG03"
                    mskCodOperatore = "02"
                Case "LOG04"
                    mskCodOperatore = "04"
            End Select
            
            Set mobjOutFile = New CTextFile
            With mobjOutFile
                .Add "Tipo Record", 1, "1"
                .Add "Data Invio", 8, "2"
                .Add "Unit‡ Trasporto", 3, "3"
                .Add "Badge", 3, "4"
                .Add "Progressivo Giorno", 2, "5"
                .Add "Sequenza Viaggio", 2, "6"
                .Add "Sequenza Consegna", 2, "7"
                .Add "Ora inizio Viaggio", 5, "8"
                .Add "Ora prevista Consegna", 5, "9"
                .Add "Zona", 2, "10"
                .Add "Deposito", 2, "11"
                .Add "Centro Operativo", 2, "12"
                .Add "Numero Ordine", 8, "13"
                .Add "Codice Destinazione", 8, "14"
                .Add "Tipo Vendita", 2, "15"
                .Add "Momento", 1, "16", "M"
            End With
'        Case ciInst2
        #ElseIf VersionID = 2 Then
            Select Case ExportType
                Case etSAP
                    Me.Caption = "TRASMISSIONE VIAGGI A SAP"
                    cmdBrowse.Enabled = True
                    lblNumInvio.Visible = False
                    mskCodOperatore.Visible = False
                    
                    imgState(16).Visible = False ' Ente Emittente
                    lblField4.Visible = False    ' Ente Emittente
                    Label1(16).Visible = False   ' Ente Emittente
                    MaskEdBox1.Visible = False   ' Ente Emittente
                    
                    mskFileName = gmRegistry.GetSetting("Options", "TripExportFile", , True)
                    Set mobjOutFile200 = New CTextFile
                    With mobjOutFile200
                        .Add "Tipo Record", 3, "1", "200"
                        .Add "Ente Emittente", 3, "2", , fftNumerico
                        .Add "Data Consegna", 8, "3"
                        .Add "Numero Viaggio", 7, "4", , fftNumerico
                        .Add "Numero Contratto", 8, "5" ', , fftNumerico
                        .Add "Codice Vettore", 3, "6"
                        .Add "Codice Deposito", 3, "7", , fftNumerico
                        .Add "Flag consegna", 1, "8", "D"
                        .Add "Flag SA", 1, "9"
                        .Add "Cod. Contratto Secondario", 8, "10"
                        .Add "Durata viaggio giorni", 2, "11", "01", fftNumerico
                        .Add "Giorni sabato", 2, "12"
                        .Add "Giorni festivi", 2, "13"
                        .Add "Numero consegne", 2, "14", , fftNumerico
                        .Add "Numero consegne Soc", 2, "15", , fftNumerico
                        .Add "Punto fine viaggio", 5, "16"
                        .Add "KM da ultimo scarico", 4, "17", , fftNumerico
                        .Add "KM complessivi previsti", 4, "18", , fftNumerico
                        .Add "Flag KM previsti", 1, "19", "A"
                        .Add "KM effettivi percorsi", 4, "20", , fftNumerico
                        .Add "Flag KM previsti", 1, "21", "A"
                        .Add "Filler", 2, "22", "  "
                        .Add "Presidio Logistico", 3, "23", , fftNumerico
                        .Add "ex lettera 9000", 1, "24", " "
                    End With
                    Set mobjOutFile201 = New CTextFile
                    With mobjOutFile201
                        .Add "Tipo Record", 3, "1", "201"
                        .Add "Ente Emittente", 3, "2", , fftNumerico
                        .Add "Data Consegna", 8, "3"
                        .Add "Numero Viaggio", 7, "4", , fftNumerico
                        .Add "Targa Motrice", 8, "5"
                        .Add "Targa Rimorchio", 8, "6"
                        .Add "Tipo Cisterna", 3, "7"
                        .Add "Autista 1", 24, "8"
                        .Add "Ora inizio carico", 4, "9"
                        .Add "Ora inizio viaggio", 4, "10"
                        .Add "Sequenza viaggio", 4, "11", , fftNumerico
                        .Add "Presidio logistico", 3, "12", , fftNumerico
                        .Add "ex lettera 9000", 1, "13", " "
                    End With
                    Set mobjOutFile202 = New CTextFile
                    With mobjOutFile202
                        .Add "Tipo Record", 3, "1", "202"
                        .Add "Ente Emittente", 3, "2", , fftNumerico
                        .Add "Data Consegna", 8, "3"
                        .Add "Numero Viaggio", 7, "4", , fftNumerico
                        .Add "Autista 2", 25, "5"
                        .Add "Foglio viaggio rif.", 7, "6", "0000000", fftNumerico
                        .Add "Punto inizio viaggio", 5, "7"
                        .Add "Codice Autista 1", 4, "8", , fftNumerico
                        .Add "Codice Autista 2", 4, "9", , fftNumerico
                        .Add "Filler", 4, "10", "    "
                        .Add "Unit‡ di trasporto", 6, "11", , fftNumerico
                        .Add "Presidio logistico", 3, "12", , fftNumerico
                        .Add "ex lettera 9000", 1, "13", " "
                    End With
                    Set mobjOutFile210 = New CTextFile
                    With mobjOutFile210
                        .Add "Tipo Record", 3, "1", "210"
                        .Add "Ente Emittente", 3, "2", , fftNumerico
                        .Add "Data Consegna", 8, "3"
                        .Add "Numero Viaggio", 7, "4", , fftNumerico
                        .Add "Progressivo Consegna", 2, "5", , fftNumerico
                        .Add "Codice destinazione", 5, "6"
                        .Add "Flag autostradale", 1, "7"
                        .Add "Codice Cliente ERG", 10, "8", , fftNumerico
                        .Add "Quantit‡ min. ordine", 3, "9", 0, fftNumerico
                        .Add "Filler", 34, "10", Space(34)
                        .Add "Presidio logistico", 3, "11", , fftNumerico
                        .Add "ex lettera 9000", 1, "12", " "
                    End With
                    Set mobjOutFile220 = New CTextFile
                    With mobjOutFile220
                        .Add "Tipo Record", 3, "1", "220"
                        .Add "Ente Emittente", 3, "2", , fftNumerico
                        .Add "Data Consegna", 8, "3"
                        .Add "Numero Viaggio", 7, "4", , fftNumerico
                        .Add "Progressivo Consegna", 2, "5", , fftNumerico
                        .Add "Codice prodotto", 5, "6", , fftNumerico
                        .Add "Codice posizione doganale", 2, "7", , fftNumerico
                        .Add "Codice Imballo", 3, "8", "000", fftNumerico
                        .Add "Quantit‡ Lt", 5, "9", , fftNumerico
                        .Add "Quantit‡ Kg", 5, "10", "0", fftNumerico
                        .Add "Filler", 33, "11", Space(33)
                        .Add "Presidio logistico", 3, "12", , fftNumerico
                        .Add "ex lettera 9000", 1, "13", " "
                    End With
                    Set mobjOutFile240 = New CTextFile
                    With mobjOutFile240
                        .Add "Tipo Record", 3, "1", "240"
                        .Add "Ente Emittente", 3, "2", , fftNumerico
                        .Add "Data Consegna", 8, "3"
                        .Add "Numero Viaggio", 7, "4", , fftNumerico
                        .Add "Progressivo Consegna", 2, "5", , fftNumerico
                        .Add "Tipo Ordine", 2, "6"
                        .Add "Numero Ordine", 9, "7", , fftNumerico
                        .Add "Km parziali", 5, "8", , fftNumerico
                        .Add "Data acquisizione ordine", 8, "9"
                        .Add "Ora acquisizione ordine", 4, "10"
                        .Add "Data ultima modifica ordine", 8, "11"
                        .Add "Ora ultima modifica ordine", 4, "12"
                        .Add "Ora consegna prevista", 4, "13"
                        .Add "Filler", 9, "14", Space(9)
                        .Add "Presidio logistico", 3, "15", , fftNumerico
                        .Add "ex lettera 9000", 1, "16", " "
                    End With
                Case etDECO
                    Me.Caption = "TRASMISSIONE VIAGGI A DEPOSITO COMUNE"
                    cmdBrowse.Enabled = True
                    lblNumInvio.Visible = False
                    mskCodOperatore.Visible = False
                    
                    imgState(16).Visible = False ' Ente Emittente
                    lblField4.Visible = False    ' Ente Emittente
                    Label1(16).Visible = False   ' Ente Emittente
                    MaskEdBox1.Visible = False   ' Ente Emittente
                    
                    mskFileName = gmRegistry.GetSetting("Options", "TripExportFile", , True)
                    Set mobjOutFile200 = New CTextFile
                    With mobjOutFile200
                        .Add "Tipo Record", 2, "1", "01"
                        .Add "Data Viaggio", 8, "2"
                        .Add "Impianto", 4, "3", Space$(4)
                        .Add "Numero Viaggio", 10, "4", , fftNumerico
                        .Add "Filler", 1, "5", "0"
                        .Add "Progressivo Viaggio", 5, "6", , fftNumerico
                        .Add "Tipo rapporto", 1, "7", "1"
                        .Add "Emittente", 4, "8", "ERG"
                        .Add "Depositante", 4, "9", "ERG"
                        .Add "Codice Vettore", 10, "10"
                        .Add "Patente Primo Autista", 10, "11"
                        .Add "Patente Secondo Autista", 10, "12"
                        .Add "Targa Motrice", 10, "13"
                        .Add "Targa Rimorchio", 10, "14"
                        .Add "Ora carico", 6, "15"
                        .Add "Nome Primo Autista", 30, "16"
                        .Add "Nome Secondo Autista", 30, "17"
                        .Add "Filler", 97, "18", Space(97)
                        .Add "Flag InsCanc", 1, "19"
                        .Add "Fine record", 2, "20", "**"
                    End With
                    Set mobjOutFile201 = New CTextFile
                    With mobjOutFile201
                        .Add "Tipo Record", 2, "1", "02"
                        .Add "Data Viaggio", 8, "2"
                        .Add "Impianto", 4, "3", Space$(4)
                        .Add "Numero Viaggio", 10, "4", , fftNumerico
                        .Add "Progressivo Consegna", 4, "5", , fftNumerico
                        .Add "Codice Cliente", 10, "6"
                        .Add "Nome Cliente", 35, "7"
                        .Add "Indirizzo Cliente", 35, "8"
                        .Add "Localit‡ Cliente", 35, "9"
                        .Add "CAP Cliente", 5, "10"
                        .Add "Provincia Cliente", 3, "11"
                        .Add "Partita IVA Cliente", 11, "12"
                        .Add "Filler", 91, "13", Space(91)
                        .Add "Fine record", 2, "14", "**"
                    End With
                    Set mobjOutFile202 = New CTextFile
                    With mobjOutFile202
                        .Add "Tipo Record", 2, "1", "03"
                        .Add "Data Viaggio", 8, "2"
                        .Add "Impianto", 4, "3", Space$(4)
                        .Add "Numero Viaggio", 10, "4", , fftNumerico
                        .Add "Progressivo Consegna", 4, "5", , fftNumerico
                        .Add "Codice Destinazione", 10, "6"
                        .Add "Nome Destinazione", 35, "7"
                        .Add "Indirizzo Destinazione", 35, "8"
                        .Add "Localit‡ Destinazione", 35, "9"
                        .Add "CAP Destinazione", 5, "10"
                        .Add "Provincia Destinazione", 3, "11"
                        .Add "Partita IVA Destinazione", 11, "12"
                        .Add "Filler", 91, "13", Space(91)
                        .Add "Fine record", 2, "14", "**"
                    End With
                    Set mobjOutFile210 = New CTextFile
                    With mobjOutFile210
                        .Add "Tipo Record", 2, "1", "05"
                        .Add "Data Viaggio", 8, "2"
                        .Add "Impianto", 4, "3", Space$(4)
                        .Add "Numero Viaggio", 10, "4", , fftNumerico
                        .Add "Somma quantit‡ Viaggio", 10, "5", , fftNumerico
                        .Add "Somma consegne Viaggio", 2, "6", , fftNumerico
                        .Add "Somma righe Viaggio", 3, "7", , fftNumerico
                        .Add "Filler", 214, "8", Space(214)
                        .Add "Fine record", 2, "9", "**"
                    End With
                    Set mobjOutFile220 = New CTextFile
                    With mobjOutFile220
                        .Add "Tipo Record", 2, "1", "04"
                        .Add "Data Viaggio", 8, "2"
                        .Add "Impianto", 4, "3", Space$(4)
                        .Add "Numero Viaggio", 10, "4", , fftNumerico
                        .Add "Progressivo Consegna", 4, "5", , fftNumerico
                        .Add "Progressivo Prodotto", 6, "6", , fftNumerico
                        .Add "Codice Prodotto", 18, "7", , fftNumerico
                        .Add "Filler", 2, "8"
                        .Add "Quantit‡", 10, "9", , fftNumerico
                        .Add "Unit‡ di misura", 3, "10", "LT"
                        .Add "Filler", 186, "11", Space(186)
                        .Add "Fine record", 2, "12", "**"
                    End With
            End Select
         #ElseIf VersionID = 4 Then ' Trasmissione Viaggi a Tamoil
            cmdBrowse.Enabled = True
            lblNumInvio.Visible = False
            mskCodOperatore.Visible = False
            mskFileName = gmRegistry.GetSetting("Options", "TripExportFile", , True)
            Set mobjOutFile400 = New CTextFile
            With mobjOutFile400
                .Add "Tipo Record", 2, "1", "T0"
                .Add "Progressivo Viaggio", 3, "2", , fftNumerico
                .Add "Unit‡ di trasporto", 3, "3"
                .Add "Nome Autista 2", 20, "12", Space(20)
                .Add "Sequenza viaggio", 2, "6", , fftNumerico
                .Add "Data Viaggio", 8, "7"
                .Add "Ora inizio carico", 4, "8"
                .Add "Ora inizio viaggio", 4, "9"
                .Add "KM previsti", 5, "10", , fftNumerico
                .Add "Codice Deposito", 3, "11"
                .Add "Targa Motrice", 10, "4"
                .Add "Targa Rimorchio", 10, "5"
                .Add "Codice Magazzino", 3, "13"
                .Add "Filler 1", 2, "14", Space(2)
                .Add "Filler 2", 1, "15", "*"
            End With
            Set mobjOutFile401 = New CTextFile
            With mobjOutFile401
                .Add "Tipo Record", 2, "1", "T1"
                .Add "Progressivo Viaggio", 3, "2", , fftNumerico
                .Add "Progressivo Consegna", 3, "3", , fftNumerico
                .Add "Codice Destinazione", 6, "4"
                .Add "Codice Societ‡", 2, "5"
                .Add "Qt‡ Prodotto 1", 5, "6", 0, fftNumerico
                .Add "Qt‡ Prodotto 2", 5, "7", 0, fftNumerico
                .Add "Qt‡ Prodotto 3", 5, "8", 0, fftNumerico
                .Add "Qt‡ Prodotto 4", 5, "9", 0, fftNumerico
                .Add "Qt‡ Prodotto 5", 5, "10", 0, fftNumerico
                .Add "Qt‡ Prodotto 6", 5, "11", 0, fftNumerico
                .Add "Qt‡ Prodotto 7", 5, "12", 0, fftNumerico
                .Add "Qt‡ Prodotto 8", 5, "13", 0, fftNumerico
                .Add "Qt‡ Prodotto 9", 5, "14", 0, fftNumerico
                .Add "Qt‡ Prodotto 10", 5, "15", 0, fftNumerico
                .Add "Codice Ordine", 7, "16"
                .Add "KM Parziali", 5, "17", 0, fftNumerico
                .Add "Filler 1", 1, "18", Space(2)
                .Add "Filler 2", 1, "19", "*"
            End With
'        #ElseIf VersionID = 5 Then
        #End If
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
    #If VersionID = 4 Then
        With AXX.DS
            .Fields(.KeyFields(1)).FindValue = ID_PARAM
            If AXX.Find Then
                MaskEdBox1 = AXX.DS.Fields("16").value
                If DataClass.DS.ValidateClientField("4", MaskEdBox1) Then
                    FieldUpdateDescription "4", Me, MaskEdBox1
                    mskFileName = mskFileName & MaskEdBox1
                End If
            End If
        End With
    #End If
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
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With

    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
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
        'Set .MultiSoc.DS.DA = DA
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
    
    Loading = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim strFileName As String
    
    #If VersionID = 1 Then
        gmRegistry.SaveSetting "Options", "OperatorCode", mskCodOperatore, True
        gmRegistry.SaveSetting "Options", "TripExportFileAGIP", mskFileName, True
    #ElseIf (VersionID = 2) Then
        gmRegistry.SaveSetting "Options", "TripExportFile", mskFileName, True
    #ElseIf (VersionID = 4) Then 'Or (VersionID = 5) Then
        strFileName = Left$(mskFileName, InStr(1, mskFileName, "."))
        gmRegistry.SaveSetting "Options", "TripExportFile", strFileName, True
    #End If
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
    Set XA = Nothing
    Set frmVVTExport = Nothing
End Sub

Private Sub ExportRecords004()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngQtaLt As Long
    Dim intRowsTrip As Integer
    Dim intRowsConsegna As Integer
    Dim intKmPrevisti As Integer
    Dim intKmDefinitivi As Integer
    Dim intKmDefinitiviConsegna As Integer
    Dim intNumSequenza As Integer
    Dim i As Long
    Dim j As Long
    Dim intFileNum As Integer
    Dim intProgressivoConsegna As Integer
    Dim intCountSIF As Integer
    Dim strIDParam As String
    Dim strIDViaggio As String
    Dim strIDViaggioDef As String
    Dim strIDConsegna As String
    Dim strIDEnteEmittente As String
    Dim strIDDeposito As String
    Dim strIDCliente As String
    Dim strIDDestinazione As String
    Dim strIDVettore As String
    Dim strIDFineViaggio As String
    Dim strIDUT As String
    Dim strIDMotrice As String
    Dim strIDRimorchio As String
    Dim strIDTipoMotrice As String
    Dim strIDTipoRimorchio As String
    Dim strIDTipoAutomezzo As String
    Dim strIDAutista1 As String
    Dim strIDAutista2 As String
    Dim strNomeAutista1 As String
    Dim strNomeAutista2 As String
    Dim strIDMagazzino As String
    Dim StrIDSociet‡ As String
    Dim StrIDSociet‡Default As String
    Dim strContratto As String
    Dim strContrattoSecondario As String
    Dim IntProVVC As Integer
    Dim strIDProdotto As String
    Dim strIDPosizioneDoganale As String
    Dim strDataInizioViaggio As String
    Dim strOraInvio As String
    Dim strUnit‡Trasporto As String
    Dim strOraInizioViaggio As String
    Dim strOraInizioCarico As String
    Dim strDataRicezioneOrdine As String
    Dim strOraRicezioneOrdine As String
    Dim strDataInserimentoOrdine As String
    Dim strOraInserimentoOrdine As String
    Dim strTipoOrdine As String
    Dim strFileName As String
    Dim strOraConsegna As String
    Dim strRifSoc As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim strDummy As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim varProdotti(1 To 10) As String
    Dim ErrCode As Errors
    Dim blnFlagProdottiOK As Boolean
    Dim intIndexProdotto As Integer
    Dim intPrgVVT As Integer
    Dim blnFound As Boolean
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    With AXX.DS  'Ciclo di Ricerca dell'ordine dei prodotti in AXX
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        StrIDSociet‡Default = AXX.DS.Fields("4").value
        If StrIDSociet‡Default = vbNullString Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        For i = 1 To 10
            If Trim$(AXX.DS.Fields(i + 6).value) <> vbNullString Then
                blnFlagProdottiOK = True
                varProdotti(i) = Trim$(AXX.DS.Fields(i + 6).value)
            End If
        Next i
        If Not blnFlagProdottiOK Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
    End With
    
    strFileName = mskFileName
    If Not mobjOutFile400.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    intFileNum = mobjOutFile400.FileNum
    mobjOutFile401.FileNum = intFileNum
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    ' --- Ciclo di elaborazione dei Viaggi VVT
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        intPrgVVT = lngRow + 1
        With ADE.DS
            strDummy = GridArraySelect.value(lngRow, COL_ID_DEPOSITO) 'ID Deposito
            .Fields("0").FindValue = strDummy
            If Not ADE.Find Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
            
        End With
        With GridArraySelect
            intNumSequenza = .value(lngRow, COL_NUM_SEQUENZA)
            strUnit‡Trasporto = .value(lngRow, COL_ID_UNITA_TRASPORTO)
            strOraInizioViaggio = .value(lngRow, COL_HR_VIAGGIO)
            strOraInizioViaggio = Format$(Hour(strOraInizioViaggio), "00") & _
                                  Format$(Minute(strOraInizioViaggio), "00")
            strDataInizioViaggio = .value(lngRow, COL_DT_VIAGGIO)
            strIDViaggioDef = .value(lngRow, COL_ID_VIAGGIO_DEF)
            strIDMotrice = .value(lngRow, COL_ID_MOTRICE)
            strIDRimorchio = .value(lngRow, COL_ID_RIMORCHIO)
'            strIDTipoMotrice = .value(lngRow, COL_TIPO_MOTRICE)
'            If Len(.value(lngRow, COL_TIPO_RIMORCHIO)) > 0 Then
'                strIDTipoRimorchio = .value(lngRow, COL_TIPO_RIMORCHIO)
'            End If
            strOraInizioCarico = .value(lngRow, COL_ORA_INIZIO_CARICO)
            strOraInizioCarico = Format$(Hour(strOraInizioCarico), "00") & _
                                 Format$(Minute(strOraInizioCarico), "00")
'            If Len(strIDTipoRimorchio) > 0 Then
'                strDummy = strIDTipoRimorchio
'            Else
'                strDummy = strIDTipoMotrice
'            End If
        End With
        intKmPrevisti = GridArraySelect.value(lngRow, COL_KM_PREVISTI)
        'strIDAutista2 = GridArraySelect.value(lngRow, COL_ID_AUTISTA2)
        If Len(GridArraySelect.value(lngRow, COL_NOME_AUTISTA2)) > 0 Then
            strNomeAutista2 = GridArraySelect.value(lngRow, COL_NOME_AUTISTA2)
        Else
            strNomeAutista2 = ""
        End If
        strIDViaggio = GridArraySelect.value(lngRow, 0) 'Codice viaggio
        With VVT.DS
            strDummy = strIDViaggio 'IDViaggio
            .Fields("0").FindValue = strDummy
            If Not VVT.Find Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
            strIDMagazzino = .Fields("23").value '    .MultiSoc.DS.GetFirmCode(ID_SOC)
            StrIDSociet‡ = .Fields("3").value
        End With
        strIDDeposito = ADE.DS.MultiSoc.DS.GetFirmCode(StrIDSociet‡)
        
        With mobjOutFile400
            .ClearFields
            .Item("2").value = intPrgVVT                ' Progressivo Viaggio
            .Item("3").value = strUnit‡Trasporto     ' Unita di Trasporto
            .Item("4").value = strIDMotrice          ' Targa Motrice
            .Item("5").value = strIDRimorchio        ' Targa Rimorchio
            .Item("6").value = intNumSequenza        ' Sequenza Viaggio
            
            strDataInizioViaggio = Format$(strDataInizioViaggio, "yyyy") & _
                                   Format$(Month(strDataInizioViaggio), "00") & _
                                   Format$(Day(strDataInizioViaggio), "00")
            .Item("7").value = strDataInizioViaggio  ' Data Viaggio
            .Item("8").value = strOraInizioCarico    ' Ora Inizio Carico
            .Item("9").value = strOraInizioViaggio   ' Ora Inizio Viaggio
            .Item("10").value = intKmPrevisti        ' Km Previsti
            .Item("11").value = strIDDeposito
            .Item("12").value = strNomeAutista2
            .Item("13").value = strIDMagazzino
        End With
        If Not mobjOutFile400.WriteLine(True) Then   'Scrivo la Testata
             intRowError = OutputError
             GoTo LogResult
        End If
        With VVC.DS
            .Fields(.KeyFields(1)).FindValue = strIDViaggio
            If Not VVC.GetList(True) Then
                intRowError = RowsNotFound
                GoTo LogResult
            End If
            intRowsTrip = .HowMany - 1
        End With
        For i = 0 To intRowsTrip ' --- Elaborazione dei prodotti consegnati
            IntProVVC = i + 1
            strIDConsegna = VVC.DS.ListItem(0, i)
            With mobjOutFile401
                .ClearFields
                .Item("2").value = intPrgVVT       'Progressivo Viaggio
                .Item("3").value = IntProVVC       'Progressivo Consegna
                .Item("4").value = CStr(VVC.DS.ListItem(6, i)) ' Destinazione
                .Item("5").value = StrIDSociet‡    ' Societa
                .Item("16").value = VVC.DS.ListItem(10, i)
                .Item("17").value = VVC.DS.ListItem(9, i)
            End With
            With VVP.DS
              .Fields(.KeyFields(1)).FindValue = strIDViaggio
              .Fields(.KeyFields(2)).FindValue = strIDConsegna
              If Not VVP.GetList(True) Then
                intRowError = RowsNotFound
                GoTo LogResult
              End If
              intRowsConsegna = .HowMany - 1
            End With
            For j = 0 To intRowsConsegna
                 strIDProdotto = VVP.DS.ListItem(1, j) ' Codice Prodotto
                 strIDPosizioneDoganale = VVP.DS.ListItem(7, j)
                 lngQtaLt = VVP.DS.ListItem(3, j)      ' Quantit‡ Corrispondente
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
                 
                 ' Se la posizione doganale Ë SIF scriviamo la quantit‡
                 ' prodotto nelle colonne da 5 a 8
                 If (strIDPosizioneDoganale = "2") And (intIndexProdotto <= 4) Then
                    intIndexProdotto = intIndexProdotto + 4
                 End If
                 
                 mobjOutFile401.Item(CStr(5 + intIndexProdotto)).value = lngQtaLt ' Codice Prodotto
             Next j
             If Not mobjOutFile401.WriteLine(True) Then     'Scrivo la Riga
                intRowError = OutputError
                GoTo LogResult
            End If
        Next i

        ' ---
        ' --- Modifico lo stato del viaggio
        If Not VVT.SetFlagTrasmissione("T") Then
            intRowError = NoSetState
            GoTo LogResult
        End If
     
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
'        lblLog(7) = lngCountExcluded
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case OutputError
                strErrorMessage = "Errore durante la scrittura sul file"
            Case noUT
                strErrorMessage = "Unit‡ di trasporto non trovata"
            Case RowsNotFound
                strErrorMessage = "Errore durante la lettura delle consegne"
            Case NoSetState
                strErrorMessage = "Impossibile cambiare lo stato del viaggio"
            Case NoOrderID
                strErrorMessage = "Ordine non trovato"
            Case NoTripID
                strErrorMessage = "Viaggio non trovato"
            Case NoTipoAutomezzo
                strErrorMessage = "Codice Tipo Automezzo non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Ente Emittente non trovato."
            Case NoCliente
                strErrorMessage = "Codice Cliente non trovato."
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato."
            Case NoPosizioneDoganale
                strErrorMessage = "Codice Posizione Doganale non trovato."
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
    Set C = ErrH.SetError(ErrCode, "frmVVTExportMover", _
                "ExportRecords004", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFile400.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub

End Sub

 
Private Sub MaskEdBox1_LostFocus()
    Dim strFileName As String
    
    strFileName = mskFileName
    If Right$(strFileName, 1) = "." Then
        strFileName = strFileName & MaskEdBox1
    Else
        strFileName = Left$(strFileName, InStr(1, strFileName, ".")) & MaskEdBox1
    End If
    mskFileName = strFileName
End Sub


