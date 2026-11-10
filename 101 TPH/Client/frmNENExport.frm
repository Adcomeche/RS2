VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmNENExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CEPSA - TRASMISSIONE NOTE DI CONSEGNA"
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
         Picture         =   "frmNENExport.frx":0000
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
         Picture         =   "frmNENExport.frx":0102
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
         Picture         =   "frmNENExport.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Viaggi"
         Height          =   375
         Left            =   2280
         Picture         =   "frmNENExport.frx":0306
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
         Picture         =   "frmNENExport.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Seleziona File..."
         Height          =   375
         Left            =   120
         Picture         =   "frmNENExport.frx":050A
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
         Picture         =   "frmNENExport.frx":060C
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
         Picture         =   "frmNENExport.frx":070E
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
      Height          =   1815
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3360
         Picture         =   "frmNENExport.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7320
         Picture         =   "frmNENExport.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   2040
         TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   4
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
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
         Left            =   8160
         Top             =   360
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
               Picture         =   "frmNENExport.frx":0A14
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNENExport.frx":0B26
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNENExport.frx":0C38
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskSoc 
         Height          =   255
         Left            =   960
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
      Begin MSMask.MaskEdBox mskEnteEmittente 
         Height          =   255
         Left            =   5400
         TabIndex        =   6
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ente Emitt.:"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblsocieta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   39
         Tag             =   "4"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   8520
         Tag             =   "4"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   4080
         Tag             =   "3"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblsocieta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   38
         Tag             =   "3"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Società:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio Def."
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   13
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
         TabIndex        =   15
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
         Picture         =   "frmNENExport.frx":0D4A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Viaggio"
         Height          =   375
         Left            =   120
         Picture         =   "frmNENExport.frx":0E4C
         TabIndex        =   10
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmNENExport.frx":0F4E
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
            TextSave        =   "11.54"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17/01/2008"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmNENExport.frx":3956
      TabIndex        =   7
      Tag             =   "SELECT"
      Top             =   1920
      Width           =   8655
   End
End
Attribute VB_Name = "frmNENExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'02/09/2003  CZucchini
'Interfaccia esportazione viaggi
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
Private mobjOutFile As CTextFile
Private mobjOutFileDetail As CTextFile
Private mobjOutFileHeader As CTextFile

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

Private Sub DisplaySelection(buffer As Variant)
    Dim varLog As Variant
    Dim i As Long
    Dim lngRows As Long
    
    lngRows = UBound(buffer, 2) + 1
    RefreshUnboundGrid TDBGrid1, buffer, GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
'    lblLog(7) = 0
    If lngRows > 0 Then
        ReDim varLog(0 To 1, 0 To lngRows - 1)
        For i = 0 To lngRows - 1
            varLog(0, i) = buffer(0, i)
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
    
    Const TEMP_FILE = "NEN.tmp"
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim i As Long
    Dim j As Long
    Dim intProgressivoConsegna As Integer
    Dim intProgressivoRecord As Integer
    Dim intFileFromNum As Integer
    Dim intFileToNum As Integer
    Dim strIDViaggio As String
    Dim strIDEnteEmittente As String
    Dim strDataViaggio As String
    Dim strAnnoDoc As String
    Dim strDataInvio As String
    Dim strOraInvio As String
    Dim strOraConsegna As String
    Dim strFileName As String
    Dim strErrorMessage As String
    Dim strRifVR As String
    Dim strNumOrdine As String
    Dim strNumeratore As String
    Dim strProgrInvio As String
    Dim strLine As String
    Dim intRowError As Errors
    Dim varBufferNoteConsegna As Variant
    Dim varCodTrasportatore As Variant
    Dim varQtaCaricata As Variant
    Dim varTotQtaScaricata As Variant
    Dim varQtaCalo As Variant
    Dim varCounter As Variant
    Dim ErrCode As Errors
    Dim C As CError
    

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
    varCodTrasportatore = "054320"
    ' ---
    
    strFileName = TEMP_FILE
    If Len(Dir(TEMP_FILE)) > 0 Then
        Kill TEMP_FILE
    End If
    If Not mobjOutFile.OpenFile(strFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    mobjOutFileDetail.FileNum = mobjOutFile.FileNum
    mobjOutFileHeader.FileNum = mobjOutFile.FileNum
    
    strDataInvio = Format$(Date, "yyyymmdd")
    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
    
    With mobjOutFileHeader
        .Item("2").value = strIDEnteEmittente
        .Item("3").value = mstrProgrInvio
        .Item("4").value = strDataInvio
        .Item("5").value = strOraInvio
        .Item("6").value = "0000"
        If Not .WriteLine Then
            GoTo ErrHandler
        End If
        intProgressivoRecord = intProgressivoRecord + 1
    End With
    
    
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    ' --- Ciclo di elaborazione dei viaggi
    
    For lngRow = 0 To lngNumRows
        If GridArraySelect.value(lngRow, 7) = "D" Then
            ErrCode = ParameterNotFound
            GoTo ErrHandler
        End If
    Next lngRow
    
    intProgressivoRecord = 1
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        
        strIDViaggio = GridArraySelect.value(lngRow, 0)
        strDataViaggio = GridArraySelect.value(lngRow, 1)
        strRifVR = varCodTrasportatore & Format(GridArraySelect.value(lngRow, 4), String(9, "0"))
        strAnnoDoc = Format(strDataViaggio, "yyyy")
        varQtaCaricata = GridArraySelect.value(lngRow, 5)
        If Not IsNumeric(varQtaCaricata) Then
            ' Se non è presente una quantità caricata usciamo con errore
            GoTo ErrHandler
        End If
        If IsNull(GridArraySelect.value(lngRow, 3)) Or (Len(GridArraySelect.value(lngRow, 3)) = 0) Then
            GoTo ErrHandler
        End If
        If Len(strRifVR) = 0 Then
            GoTo ErrHandler
        End If
        With mobjOutFile
            varBufferNoteConsegna = DataClass.GetElencoNoteConsegna(strIDViaggio, GridArraySelect.value(lngRow, 4))
            varTotQtaScaricata = 0
            For i = 0 To UBound(varBufferNoteConsegna, 2)
'                varTotQtaScaricata = varTotQtaScaricata + varBufferNoteConsegna(4, i)
                varTotQtaScaricata = varTotQtaScaricata + varBufferNoteConsegna(5, i)
            Next i
            
            .ClearFields
            .Item("2").value = strAnnoDoc                                                          'anno documento consegna
            .Item("3").value = GridArraySelect.value(lngRow, 3)                                    'ticket
            .Item("4").value = strRifVR                                                            'referencia VR
            .Item("5").value = varQtaCaricata                                                      'quantità caricata
            If varQtaCaricata <> varTotQtaScaricata Then
                .Item("6").value = GridArraySelect.value(lngRow, 6) + 1                            'numero note consegna
            Else
                .Item("6").value = GridArraySelect.value(lngRow, 6)                                'numero note consegna
            End If
            
            If Not .WriteLine Then
                 intRowError = OutputError
                 GoTo LogResult
            End If
            intProgressivoRecord = intProgressivoRecord + 1
            
            If Not IsEmpty(varBufferNoteConsegna) Then
                intProgressivoConsegna = 1
                For i = 0 To UBound(varBufferNoteConsegna, 2)
                    With mobjOutFileDetail
                        .ClearFields
                        .Item("2").value = strAnnoDoc
                        .Item("3").value = GridArraySelect.value(lngRow, 3)
                        .Item("4").value = strRifVR
                        .Item("5").value = intProgressivoConsegna
                        strNumOrdine = varBufferNoteConsegna(0, i)
                        If Not IsNull(varBufferNoteConsegna(1, i)) And Len(varBufferNoteConsegna(1, i)) > 0 Then
                            strNumOrdine = strNumOrdine & varBufferNoteConsegna(1, i)
                        Else
                            strNumOrdine = strNumOrdine & "00"
                        End If
                        .Item("7").value = strNumOrdine
                        .Item("8").value = Right(varBufferNoteConsegna(2, i), 8)
                        .Item("9").value = varBufferNoteConsegna(3, i)
                        '.Item("10").value = varBufferNoteConsegna(4, i)
                        .Item("10").value = varBufferNoteConsegna(5, i)
                        .Item("12").value = Format(strDataViaggio, "yyyymmdd")
                        strOraConsegna = Format(varBufferNoteConsegna(7, i), "HH:MM")
                        strOraConsegna = Left(strOraConsegna, 2) & ":" & Right(strOraConsegna, 2)
                        .Item("13").value = strOraConsegna
                        If Not .WriteLine Then
                             intRowError = OutputError
                             GoTo LogResult
                        End If
                        intProgressivoConsegna = intProgressivoConsegna + 1
                        intProgressivoRecord = intProgressivoRecord + 1
                    End With
                Next i
                
                If varQtaCaricata <> varTotQtaScaricata Then
                    With mobjOutFileDetail
                        .ClearFields
                        .Item("2").value = strAnnoDoc
                        .Item("3").value = GridArraySelect.value(lngRow, 3)
                        .Item("4").value = strRifVR
                        .Item("6").value = "M"
                        .Item("9").value = varBufferNoteConsegna(3, i - 1)
                        varQtaCalo = varQtaCaricata - varTotQtaScaricata
                        .Item("10").value = varQtaCalo
                        If varQtaCalo > 0 Then
                            .Item("11").value = "+"
                        Else
                            .Item("11").value = "-"
                        End If
                        .Item("12").value = Format(strDataViaggio, "yyyymmdd")
                        strOraConsegna = Format(varBufferNoteConsegna(7, i - 1), "HH:MM")
                        strOraConsegna = Left(strOraConsegna, 2) & ":" & Right(strOraConsegna, 2)
                        .Item("13").value = strOraConsegna
                        If Not .WriteLine Then
                             intRowError = OutputError
                             GoTo LogResult
                        End If
                        intProgressivoRecord = intProgressivoRecord + 1
                    End With
                End If
                
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
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case OutputError
                strErrorMessage = "Errore durante la scrittura sul file"
            Case NoOrderID
                strErrorMessage = "Ordine non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
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
    
    lngCountOK = 0
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & strFileName
        Case ParameterNotFound
            strMessage = "Non tutti i viaggi sono completati. Impossibile trasmettere."
            strInfo = vbNullString
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
    
    If lngCountOK > 0 Then
        With DataClass.DS.DA
            strNumeratore = "040"
            Select Case mskEnteEmittente
                Case "624"
                    strNumeratore = "040"
                Case "004"
                    strNumeratore = "041"
                Case Else
                    lngCountOK = 0
            End Select
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", strNumeratore, 0)
            If varCounter = -1 Then
                ErrCode = InvalidCode
                GoTo ErrHandler
            End If
            If varCounter > 9999999999# Then
                varCounter = 1
            End If
        End With
        
        strProgrInvio = Format(varCounter, String(10, "0"))
        
        If Len(mskFileName) > 0 Then
            strFileName = mskFileName
        End If
    
        strFileName = strFileName & "\NEN_" & strProgrInvio & ".DAT"
'        FileCopy TEMP_FILE, strFileName
        
        intFileFromNum = FreeFile
        Open TEMP_FILE For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Select Case Left(strLine, 2)
                Case "01"
                    strLine = Left(strLine, Len(strLine) - 4) & Format(intProgressivoRecord, "0000")
                    Mid(strLine, 7, 10) = strProgrInvio
                    Print #intFileToNum, strLine
                Case Else
                    Print #intFileToNum, strLine
            End Select
        Loop
        Close #intFileFromNum
        
'        intFileFromNum = FreeFile
'        Open TEMP_FILE For Input Access Read As #intFileFromNum
'        Do While Not EOF(intFileFromNum)
'            Line Input #intFileFromNum, strLine
'            If Left(strLine, 2) = "03" Then
'                Print #intFileToNum, strLine
'            End If
'        Loop
'        Close #intFileFromNum
'
        
        Close #intFileToNum
    
    End If
    
    
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    
    Me.MousePointer = vbHourglass
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            varBuffer = DataClass.GetSelectionExportNEN(MaskSoc, mskEnteEmittente, mskDateFrom, mskDateTo, MaskEdBox5, MaskEdBox6)
            If Not IsEmpty(varBuffer) Then
                DisplaySelection varBuffer
            End If
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
    
    cmdBrowse.Enabled = True
    MaskSoc = gmRegistry.GetSetting("Options", "NENExportSoc", , True)
    mskEnteEmittente = gmRegistry.GetSetting("Options", "NENExportAEE", , True)
    mskFileName = gmRegistry.GetSetting("Options", "NENExportFileName", , True)
        
    Set mobjOutFileHeader = New CTextFile
    With mobjOutFileHeader
        .Add "Tipo Record", 2, "1", "01", fftNumerico
        .Add "Codice trasportatore (CLH)", 4, "2"
        .Add "Progressivo invio", 10, "3", , fftNumerico
        .Add "Data generazione", 8, "4"
        .Add "Ora generazione", 8, "5"
        .Add "Numero records", 4, "6", "0000", fftNumerico
    End With
    Set mobjOutFile = New CTextFile
    With mobjOutFile
        .Add "Tipo Record", 2, "1", "02"
        .Add "Anno nota di consegna", 4, "2", , fftNumerico
        .Add "Ticket", 15, "3", , fftNumerico
        .Add "Riferimento venta en ruta", 15, "4", , fftNumerico
        .Add "Quantità litri per Bolla di carico", 9, "5"
        .Add "Numero note di consegna", 2, "6", , fftNumerico
    End With
    Set mobjOutFileDetail = New CTextFile
    With mobjOutFileDetail
        .Add "Tipo Record", 2, "1", "03", fftNumerico
        .Add "Anno nota di consegna", 4, "2", , fftNumerico
        .Add "Ticket", 15, "3", , fftNumerico
        .Add "Riferimento venta en ruta", 15, "4", , fftNumerico
        .Add "Progressivo nota di consegna", 2, "5", , fftNumerico
        .Add "Flag tipo consegna", 1, "6", "U"
        .Add "Numero ordine", 16, "7"
        .Add "Codice destinatario (CLH)", 15, "8"
        .Add "Codice prodotto (CLH)", 7, "9"
        .Add "Quantità litri", 9, "10"
        .Add "Segno", 1, "11"
        .Add "Data consegna", 8, "12"
        .Add "Ora consegna", 5, "13"
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
    gmRegistry.SaveSetting "Options", "NENExportFileName", mskFileName, True
    gmRegistry.SaveSetting "Options", "NENExportAEE", mskEnteEmittente, True
    gmRegistry.SaveSetting "Options", "NENExportSoc", MaskSoc, True
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
    Set frmNENExport = Nothing
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

