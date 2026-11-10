VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTImport2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CEPSA - RICEZIONE ORDINI"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   10950
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   23
      Top             =   7320
      Width           =   10935
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   9240
         Picture         =   "frmOCTImport2.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9120
         Picture         =   "frmOCTImport2.frx":0102
         TabIndex        =   7
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   8145
      Width           =   10950
      _ExtentX        =   19315
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
            Object.Width           =   13864
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17:30"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "31/03/2015"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10935
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   8175
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   6720
            TabIndex        =   40
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Dest. sin KM:"
            Height          =   255
            Index           =   7
            Left            =   6720
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   32
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   38
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nuevas Dest.:"
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   39
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblprogres 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini esclusi:"
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   33
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini da importare:"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   29
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   28
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini imp. correttamente:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   8520
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ora Invio:"
            Height          =   255
            Index           =   4
            Left            =   8520
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   7560
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Invio:"
            Height          =   255
            Index           =   3
            Left            =   7560
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Progressivo Invio:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   7560
            TabIndex        =   16
            Top             =   1080
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Ente Emittente:"
            Height          =   255
            Index           =   0
            Left            =   7560
            TabIndex        =   15
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   10695
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   2055
            Left            =   120
            OleObjectBlob   =   "frmOCTImport2.frx":0204
            TabIndex        =   5
            Tag             =   "SELECT"
            Top             =   240
            Width           =   10455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   4920
         Width           =   10695
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmOCTImport2.frx":8594
            TabIndex        =   6
            Tag             =   "SELECT"
            Top             =   240
            Width           =   10455
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   8520
         Picture         =   "frmOCTImport2.frx":1B3FE
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2115
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         Picture         =   "frmOCTImport2.frx":1B930
         TabIndex        =   4
         Tag             =   "IMPORT"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   8520
         Picture         =   "frmOCTImport2.frx":1BA32
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1635
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         Picture         =   "frmOCTImport2.frx":1BF64
         TabIndex        =   3
         Tag             =   "OPEN"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   8520
         Picture         =   "frmOCTImport2.frx":1C066
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1140
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   8400
         Picture         =   "frmOCTImport2.frx":1C168
         TabIndex        =   2
         Tag             =   "BROWSE"
         Top             =   1080
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10200
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10200
         Top             =   1680
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
               Picture         =   "frmOCTImport2.frx":1C26A
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport2.frx":1C37C
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport2.frx":1C48E
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskCodSoc 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin VB.Label lblsocieta 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   37
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label Label1 
         Caption         =   "So&cietà"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "File &Ordini"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOCTImport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Interfaccia ricezione ordini

Option Explicit

Private Const MAX_COLS = 3
Private Const COL_COD_ORD_LOG = 1
Private Const COL_COD_PROD_LOG = 4
Private Const COL_RESULT_LOG = 10
Private Const COL_COD_RIF1_LOG = 2
Private Const COL_COD_IDADE_SOC = 5
Private Const COL_RIFE1 = 6

Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineriga
Private APR As CProdotto
Private APD As CPosDoganale
Private ADE As CDeposito
Private AEE As CEnteEmittente
Private ATV As CTipoVendita
Private AZO As CZona
Private ASC As CSocieta

Private mobjInFile As CTextFile
Private mobjInFileDettaglio As CTextFile
Private mstrIDEnte As String
Private mstrUltimoProgressivo As String

'-- DP 22/07/08
'-- Gestione segnalazioni in caso di nuova destinazione
Private strMsgDestinazioni As String

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mvarFunction As FunctionCodes

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Function CheckDestinazione001(ByRef strIDDestinazione As String, _
                                      ByVal strRagSocDestinazione As String, _
                                      ByVal strIDClienteDefault As String, _
                                      ByVal strIDDeposito As String, _
                                      ByVal strIndirizzo As String, _
                                      ByVal varCAP As Variant, _
                                      ByRef strNewLocalita As String, _
                                      ByVal IDTipoDestinazione As Variant, _
                                      GPSX As String, GPSY As String, _
                                      NIF As Variant, strTel1 As String, _
                                      strTel2 As String, strTelMov As String, _
                                      OrarioApertura As Variant) As Long
    Dim intRowError As Errors
    Dim strLocalita As String
    Dim strNewZona As String
    Dim strNewIndirizzo As String
    Dim blnConcurErr As Boolean
    Dim lngKmRaggio As Long
    Dim strCodDest As String
    
    intRowError = NoError
    strMsgDestinazioni = vbNullString
    
    ' --- Cerco la destinazione
    strCodDest = strIDDestinazione
    strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC_CLH, strIDDestinazione, True)
    
    If IDTipoDestinazione = "BM" Then
        IDTipoDestinazione = "B"
    End If
    
    ' --- La destinazione è già stata inserita ?
    If strIDDestinazione = vbNullString Then
        ' --- No : devo inserirla ora
        strIDDestinazione = strCodDest
        ADZ.DS.Fields.ResetFields
        RecordToADZ ADZ, strIDDestinazione, strIDClienteDefault, strIDDeposito, strRagSocDestinazione, strIndirizzo, strNewLocalita, varCAP, IDTipoDestinazione, GPSX, GPSY, NIF, strTel1, strTel2, strTelMov, OrarioApertura
        With ADZ.DS.Fields
            If Not ADZ.Insert Then
                intRowError = DestInsertError
                GoTo ExitFunction
            End If
            If Not ADZ.DS.MultiSoc.Insert Then
                intRowError = DestSocInsertError
                GoTo ExitFunction
            End If
            
            '-- DP 22/07/08
            '-- Gestione segnalazioni in caso di nuova destinazione
            lblField(4).Caption = CInt(lblField(4).Caption) + 1
            strMsgDestinazioni = "Insert. nueva dest.: " & strIDDestinazione & " - "
        End With
    Else
        ADZ.DS.Fields.ResetFields
        ADZ.DS.Field("Codice").FindValue = strIDDestinazione
        ADZ.Find
        RecordToADZ ADZ, strIDDestinazione, strIDClienteDefault, strIDDeposito, strRagSocDestinazione, strIndirizzo, strNewLocalita, varCAP, IDTipoDestinazione, GPSX, GPSY, NIF, strTel1, strTel2, strTelMov, OrarioApertura
        With ADZ.DS
            ADZ.BeginSave blnConcurErr
            .DA.EndTrans ETCOMMIT
            .EndSave
        End With
    End If
    ' ---
ExitFunction:
    CheckDestinazione001 = intRowError
End Function


Private Function CheckOrderRow001(ByRef strIDDeposito As String, _
                                  ByRef strIDEnteEmittente As String, _
                                  ByRef strIDProdotto As String, _
                                  ByRef strIDZona As String, _
                                  ByRef strIDTipoVendita As String) As Long
'
    Dim intRowError As Errors
    Dim Row As Long
    intRowError = NoError
    
    ' --- Controllo che esistano Deposito, Ente Emittente, Prodotto
   strIDDeposito = ADE.DecodeFirmCode2(ID_SOC_CLH, strIDDeposito, True)
   strIDProdotto = APR.DecodeFirmCode(ID_SOC_CLH, strIDProdotto, True)
   strIDZona = "00"
   strIDTipoVendita = "F2"
    
    If strIDDeposito = vbNullString Then
        intRowError = NoDeposito
        GoTo ExitFunction
    End If
    
    ' *** Patch per gestione deposito condiviso tra tph ed eutrasur
    If strIDDeposito = "022" And strIDEnteEmittente = "624" Then
        strIDDeposito = "023"
    End If
    If strIDDeposito = "023" And strIDEnteEmittente = "004" Then
        strIDDeposito = "022"
    End If
    If strIDDeposito = "009" And strIDEnteEmittente = "624" Then
        strIDDeposito = "008"
    End If
    If strIDDeposito = "008" And strIDEnteEmittente = "004" Then
        strIDDeposito = "009"
    End If
    If strIDDeposito = "001" And strIDEnteEmittente = "624" Then
        strIDDeposito = "003"
    End If
    If strIDDeposito = "003" And strIDEnteEmittente = "004" Then
        strIDDeposito = "001"
    End If
    ' ***
    
    If strIDEnteEmittente = vbNullString Then
        intRowError = NoEnteEmittente
        GoTo ExitFunction
    End If
    If strIDProdotto = vbNullString Then
        intRowError = NoProdotto
        GoTo ExitFunction
    End If
    ' ---
    
ExitFunction:
    CheckOrderRow001 = intRowError
End Function


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim Respok, Respuesta As String
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
                .Filter = "(*.dat)|*.dat|Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .DialogTitle = "Seleziona il file degli ordini... "
                .ShowOpen
                mskFileName = .fileName
            End With
        Case fcImport
            ' Control RS3
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Esta opción solo esta disponible desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
                ImportRecords001
            End If
    
        'ImportRecords001

        Case fcOpen
            OpenTextFile001
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub ImportRecords001()
    On Error GoTo ErrHandler
    Dim strIDEnteEmittente As String
    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDTipoOrdine As String
    Dim strIDEnteEmittenteDefault As String
    Dim strNote2 As String
    Dim strIDOrder As String
    Dim strIDPrevOrder As String
    Dim strIDProdotto As String
    Dim strIndirizzo As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strNewLocalita As String
    Dim strLocalità As String
    Dim varCAP As Variant
    Dim varNIF As Variant
    Dim strErrorMessage As String
    Dim strRagSocDestinazione As String
    Dim strCodDepositoXSocietà As String
    Dim strRifEsternoOrdine As String
    Dim strRifEsternoRiga As String
    Dim strProgressivoInvio As String
    Dim strIDUDM As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim lngTotQuantità As Long
    Dim Row As Long
    Dim Col As Integer
    Dim i As Integer
    Dim k As Integer
    Dim IDImballo As String
    Dim IDPosDoganale As String
    Dim varQuantità As Variant
    Dim strDescrizione As String
    Dim blnNewOrder As Boolean
    Dim blnLocHasChanged As Boolean
    Dim blnDeleteOrder As Boolean
    Dim blnOrdineDaAggiornare As Boolean
    Dim blnConcErr As Boolean
    Dim blnNotUpdatable As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    Dim strRotturaOrd As String
    Dim strCondTecnico As String
    Dim strGPSX As String
    Dim strGPSY As String
    Dim strConsegnatario As String
    Dim strTelefonoConsegnatario As String
    Dim strNomeNave As String
    Dim strDescrizionePorto As String
    Dim strNoteBunker As String
    Dim strFlagVentaEnRuta As String
    Dim blnSkipInsertRows As Boolean
    Dim blnFlagBunker As Boolean
    Dim Modcantidad As String
    '-- DP 13/08/08
    Dim strTel1 As String
    Dim strTel2 As String
    Dim strTelMov As String
    
    '-- DP 22/07/08
    '-- Controllo KM destinazione, deposito
    Dim varKM_ADZ_ADE As Variant
    
    '-- DP 28/11/08
    '-- Gestione ordini multideposito
    Dim strIDADE_Rottura As String
    Dim strIDADEPrev As String
    Dim StrIDpresidio As String
    '-- JG 03/03/09
    '-- Gestion pedidos eventuales
    Dim strPedEventuales As String
    
    '-- JG 24/3/9
    '-- Gestión de particion CEPSA
    Dim strparticionprev As String
    Dim strparticionleid As String
    Dim strparticionini As String
    
    ReDim ListNewLoc(0 To 0) As String
    
    If MaskCodSoc = "" Then
        Call MsgBox("E' necessario indicare la società.", vbExclamation + vbOKOnly)
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = strIDParam
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDImballoDefault = Trim$(.Fields("1").value)
        strIDPosDefault = Trim$(.Fields("2").value)
        
       Select Case mvarFunction
           Case fcIMPORD3
                strIDClienteDefault = ID_SOC_GUARNER
    
           Case Else
                strIDClienteDefault = ID_SOC_CEPSA
       End Select
        
        strIDEnteEmittenteDefault = Trim$(.Fields("16").value)
        strIDZona = "00"
    End With
    
    If (strIDImballoDefault = vbNullString) Or _
       (strIDPosDefault = vbNullString) Or _
       (strIDClienteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
  
    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    lngCountToGo = lngNumRows
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    
'    If IsUserGrandesCuentas Then
'        If Not OCT.InvalidaOrdini(Maskcodsoc, Format(Date, "dd/mm/yyyy")) Then
'            GoTo ErrHandler
'        End If
'    End If
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        strRifEsternoOrdine = Left(XA(Row, 4), 10)
        strRotturaOrd = strRifEsternoOrdine 'XA(Row, 0)
    
        '-- DP 28/11/08
        '-- Gestione ordini multideposito
        strIDADE_Rottura = Left(XA(Row, 26), 10)
        
        '-- JG 24/3/9
        '-- Control particion antigua
        strparticionleid = XA.value(Row, 42)
        strparticionini = strparticionleid
        
        If strRotturaOrd = vbNullString Or strIDADE_Rottura = vbNullString Then
            intRowError = BadIDOrderSoc
            GoTo LogResult
        End If
        'La variabile strRotturaOrd serve per capire se si deve inserire testata e riga dettaglio o solo riga dettaglio
        
        If strRotturaOrd <> strIDPrevOrder Or strIDADE_Rottura <> strIDADEPrev Or strparticionini <> strparticionprev Then
 
            ' --- Nuovo ordine : Inizio con l'inserimento di una nuova testata e riga ordine
            blnNewOrder = True
            lngTotQuantità = 0
            blnSkipInsertRows = False
            
            '--JG 24/3/9
            '-- Recojo partición anterior
            strparticionprev = strparticionini
            
            
            strIDPrevOrder = strRotturaOrd
            strIDADEPrev = strIDADE_Rottura
            
            strIDRiga = 10
            strNote2 = XA(Row, 1) & XA(Row, 2) & Format(XA(Row, 10), "00000")
            strIDEnteEmittente = Right(XA(Row, 18), 3) 'strIDEnteEmittenteDefault
            Select Case XA(Row, 2)
                Case "910101"
                    strIDTipoOrdine = "E"
                    blnFlagBunker = False
                Case "910103"
                    strIDTipoOrdine = "R"
                    blnFlagBunker = False
                Case "920101", "920102"
                    strIDTipoOrdine = "B"
                    blnFlagBunker = True
                Case "920105"
                    strIDTipoOrdine = "BM"
                    blnFlagBunker = True
                Case "910151", "910152"
                    strIDTipoOrdine = "EA"
                    blnFlagBunker = False
                Case "910153"
                    strIDTipoOrdine = "O"
                    blnFlagBunker = False
                Case "910115"
                    strIDTipoOrdine = "P"
                    blnFlagBunker = False
                Case Else
                    blnFlagBunker = False
            End Select
            strIDDeposito = XA.value(Row, 26)
            strIDProdotto = XA.value(Row, 21)
            intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, strIDProdotto, strIDZona, strIDTipoVendita)
            If intRowError <> NoError Then
                GoTo LogResult
            End If
            strIDDestinazione = "CP" + Format$(XA(Row, 5), "00000000")
            strIndirizzo = XA.value(Row, 7)
            strNewLocalita = XA.value(Row, 8)
            varCAP = XA.value(Row, 9)
            strGPSX = XA(Row, 11)
            strGPSY = XA(Row, 12)
            strConsegnatario = XA(Row, 13)
            strTelefonoConsegnatario = XA(Row, 14)
            If blnFlagBunker Then
                strNomeNave = XA.value(Row, 6)
                ' Per i bunker utilizziamo come ragione sociale della destinazione la descrizione del porto
                If Len(XA.value(Row, 15)) > 0 Then
                    strRagSocDestinazione = XA.value(Row, 15)
                Else
                    strRagSocDestinazione = vbNullString
                End If
            Else
                strNomeNave = vbNullString
                strRagSocDestinazione = XA.value(Row, 6)
            End If
            strFlagVentaEnRuta = XA.value(Row, 33)
            If IsNull(strFlagVentaEnRuta) Or Len(strFlagVentaEnRuta) = 0 Then
                strFlagVentaEnRuta = "0"
            End If
            varNIF = XA.value(Row, 34)
            
            '--DP 13/08/08
            strTel1 = XA.value(Row, 35)
            strTel2 = XA.value(Row, 36)
            strTelMov = XA.value(Row, 37)
            
            intRowError = CheckDestinazione001(strIDDestinazione, strRagSocDestinazione, _
                                               strIDClienteDefault, strIDDeposito, strIndirizzo, _
                                               varCAP, strNewLocalita, strIDTipoOrdine, strGPSX, _
                                               strGPSY, varNIF, strTel1, strTel2, strTelMov, XA(Row, 41))
            If (intRowError <> NoError) And _
               (intRowError <> DestHasChanged) Then
                GoTo LogResult
            Else
                '-- DP 22/07/08
                '-- Controllo i KM per deposito
                If strMsgDestinazioni = vbNullString Then
                    varKM_ADZ_ADE = ADZ.GetDistanza(strIDDestinazione, strIDDeposito)
                    If IsEmpty(varKM_ADZ_ADE) Then
                        lblField(5).Caption = CInt(lblField(5).Caption) + 1
                        strMsgDestinazioni = "Faltan KM por " & strIDDestinazione & " en la base " & strIDDeposito & " - "
                    Else
                        If varKM_ADZ_ADE(0, 0) = 0 Or IsNull(varKM_ADZ_ADE(0, 0)) Then
                            lblField(5).Caption = CInt(lblField(5).Caption) + 1
                            strMsgDestinazioni = "Faltan KM por " & strIDDestinazione & " en la base " & strIDDeposito & " - "
                        End If
                    End If
                End If
            End If
            '--JG 25/3/9
            'Controlo la partición
            
            If OCT.CheckSocCode1(MaskCodSoc.Text, strRifEsternoOrdine, strparticionini, , strIDDeposito) Then
                blnSkipInsertRows = True
                intRowError = OrderAlreadyExisting
                GoTo LogResult
            Else
                strDescrizione = XA.value(Row, 28)
                strProgressivoInvio = lblprogres(6)
                strCondTecnico = XA(Row, 27)
'                For k = 1 To 5
'                    If Not IsNumeric(Mid(strCondTecnico, k, 1)) Then
'                        'Mid(strCondTecnico, k, 1) = "0"
'                        'XA(Row, 22) =  & "0"
'                    End If
'                Next k
                If Not IsNumeric(strCondTecnico) Then
                    intRowError = InvalidCode
                    GoTo LogResult
                End If
                
                '-- JG 03/03/2009
                '-- Gestion del campo "Pedidos Eventuales"
                strPedEventuales = ADZ.DS.Fields("21").value
                
                'viene inserita la testata
                getPresidio "038", strIDEnteEmittente, strIDDeposito, StrIDpresidio
                RecordToOCT OCT, strIDDeposito, strIDDestinazione, strIDEnteEmittente, strIDClienteDefault, _
                            strRifEsternoOrdine, XA(Row, 24), XA(Row, 25), strNote2, strIDZona, strIDTipoVendita, _
                            strDescrizione, strProgressivoInvio, XA(Row, 29), strIDTipoOrdine, strCondTecnico, _
                            strGPSX, strGPSY, strConsegnatario, strTelefonoConsegnatario, strNomeNave, XA(Row, 31), _
                            XA(Row, 32), strFlagVentaEnRuta, XA(Row, 39), XA(Row, 40), XA(Row, 41), strPedEventuales, strparticionleid, StrIDpresidio
                If Not OCT.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
                strIDOrder = OCT.DS.Fields("0").NewValue
                If strIDOrder = vbNullString Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
                varQuantità = XA.value(Row, 22)
                If IsNumeric(varQuantità) Then
                    lngTotQuantità = lngTotQuantità + CLng(varQuantità)
                End If
                strRifEsternoRiga = Right(XA(Row, 4), 3)
                strNoteBunker = XA.value(Row, 30)
                strIDUDM = XA.value(Row, 23)
                Modcantidad = XA.value(Row, 24)
                RecordToOCR OCR, strIDOrder, strIDRiga, strIDProdotto, strIDImballoDefault, _
                            strIDPosDefault, strDescrizione, varQuantità, strRifEsternoRiga, strNoteBunker, strIDUDM, Modcantidad
                If Not OCR.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
            End If
        Else
            If blnSkipInsertRows Then
                intRowError = OrderAlreadyExisting
                GoTo LogResult
            Else
                strIDRiga = strIDRiga + 10
                varQuantità = XA.value(Row, 22)
                If IsNumeric(varQuantità) Then
                    lngTotQuantità = lngTotQuantità + CLng(varQuantità)
                End If
                strDescrizione = XA.value(Row, 28)
                strIDProdotto = XA.value(Row, 21)
                strIDDeposito = XA.value(Row, 26)
'                strIDEnteEmittente = strIDEnteEmittenteDefault
                intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, _
                                               strIDProdotto, strIDZona, strIDTipoVendita)
                strRifEsternoRiga = Right(XA(Row, 4), 3)
                strNoteBunker = XA.value(Row, 30)
                strIDUDM = XA.value(Row, 23)
                
                'SM Flag para no poder modificar la cantidad
                Modcantidad = XA.value(Row, 43)
                
                RecordToOCR OCR, strIDOrder, strIDRiga, strIDProdotto, strIDImballoDefault, _
                            strIDPosDefault, strDescrizione, varQuantità, strRifEsternoRiga, strNoteBunker, strIDUDM, Modcantidad
                If Not OCR.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
            End If
        End If
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError <> NoError) And _
           (intRowError <> ZonaEsclusaImport) And _
           (intRowError <> DestHasChanged) Then
            lngCountFail = lngCountFail + 1
        End If
        If blnNewOrder Then
            If (intRowError = NoError) Or _
               (intRowError = DestHasChanged) Then
                lngCountOK = lngCountOK + 1
            Else
                If intRowError = ZonaEsclusaImport Then
                    lngCountExcluded = lngCountExcluded + 1
                End If
            End If
            lngCountToGo = lngCountToGo - 1
        End If
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(6) = lngCountExcluded
        '-- DP 22/07/08
        '-- Aggiunta la varabile "strMsgDestinazioni" per segnalare le nuove destinazioni inserite
        Select Case intRowError
            Case NoError
                If lngTotQuantità > 32000 Then
                    strErrorMessage = strMsgDestinazioni & "Quantità > 32000 - OK - Num. Ordine : " & strIDOrder
                Else
                    strErrorMessage = strMsgDestinazioni & "OK - Num. Ordine : " & strIDOrder
                End If
            Case DestHasChanged
                blnLocHasChanged = True
                ListNewLoc(UBound(ListNewLoc)) = "Dest. modificata: " & strIDDestinazione & " " & strNewLocalita
                strErrorMessage = strMsgDestinazioni & "OK - Num. Ordine : " & strIDOrder & _
                                  " - " & ListNewLoc(UBound(ListNewLoc))
                ReDim Preserve ListNewLoc(UBound(ListNewLoc) + 1) As String
            Case OrderAlreadyExisting
                strErrorMessage = strMsgDestinazioni & "Ordine già inserito"
            Case NoTipoVendita
                strErrorMessage = strMsgDestinazioni & "Codice Tipo Vendita non trovato"
            Case NoZona
                strErrorMessage = strMsgDestinazioni & "Codice Zona non trovato"
            Case NoDeposito
                strErrorMessage = strMsgDestinazioni & "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = strMsgDestinazioni & "Codice Ente Emittente non trovato"
            Case NoProdotto
                strErrorMessage = strMsgDestinazioni & "Codice Prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = strMsgDestinazioni & "Codice Destinazione mancante"
            Case BadIDOrderSoc
                strErrorMessage = strMsgDestinazioni & "Numero Ordine mancante"
            Case DestInsertError
                strErrorMessage = strMsgDestinazioni & "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = strMsgDestinazioni & "Errore durante l'importazione della destinazione"
            Case OrderInsertError
                strErrorMessage = strMsgDestinazioni & "Errore durante l'inserimento dell'ordine "
            Case RowOrderInsertError
                strErrorMessage = strMsgDestinazioni & "Errore durante l'inserimento della riga di ordine"
            Case ZonaEsclusaImport
                strErrorMessage = strMsgDestinazioni & "Zona esclusa dall'importazione ordini"
            Case InvalidCode
                strErrorMessage = strMsgDestinazioni & "Condizionante tecnico non numerico"
        End Select
        
        XALog.value(Row, COL_RESULT_LOG) = strErrorMessage
        
        With tdbLog
            .Refresh
            .bookmark = Row
        End With
        TDBGrid1.bookmark = Row
        ' ---
        DoEvents
    Next Row
    ' ---
    
'    If IsUserGrandesCuentas Then
'        If Not OCT.AnnullaOrdiniTD(Maskcodsoc, Format(Date, "dd/mm/yyyy")) Then
'            GoTo ErrHandler
'        End If
'    End If
    
    mstrUltimoProgressivo = lblField(0)
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    'OCT.DS.DA.EndTrans 1
    Select Case ErrCode
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
            If (strIDImballoDefault = vbNullString) Then
                strInfo = strInfo & "Imballo" & vbNewLine
            End If
            If (strIDPosDefault = vbNullString) Then
                strInfo = strInfo & "Posizione Doganale" & vbNewLine
            End If
            If (strIDClienteDefault = vbNullString) Then
                strInfo = strInfo & "Codice Cliente" & vbNewLine
            End If
        Case NoImportData
            strMessage = "Non ci sono dati da importare."
            strInfo = vbNullString
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTImport2", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    If blnLocHasChanged Then
        If MsgBox("Alcune destinazioni sono state modificate. Vuoi stamparne l'elenco?", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            For i = 0 To UBound(ListNewLoc) - 1
                Printer.Print
                Printer.Print
                Printer.Print ListNewLoc(i)
            Next i
            Printer.EndDoc
        End If
    End If
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub

      Private Sub getPresidio(ByVal Soc As String, ByVal ee As String, ByVal dep As String, ByRef Presi As String)
    
        Dim DA As CDataAccess
        Set DA = DataConnect.DA
        
        Dim rowalb As Variant
        Dim lngalb As Long
        Dim sql As String

        sql = "SELECT PREPRESI FROM ANAGAPRE WHERE PRESDASC =  ? AND PRESDEE0 = ? AND PRESDADE = ? "

        If Not (DA.QueryCommand(sql, rowalb, lngalb, Soc, ee, dep)) Then
            If ee = "624" Then
            Presi = "TPH01"
            Else
            Presi = "EUT01"
            End If
        End If

        If lngalb = 0 Then
            If ee = "624" Then
            Presi = "TPH01"
            Else
            Presi = "EUT01"
            End If

        Else
            Presi = rowalb(0, 0)
        End If

     End Sub

Private Sub OpenTextFile001()
    On Error GoTo ErrHandler
    
    Dim strFileName As String
    Dim strLine As String
    Dim strEnte As String
    Dim strPreviousOrder As String
    Dim varBuffer As Variant
    Dim blnValidHeader As Boolean
    Dim blnIDOk As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim lngRows As Long
    Dim lngCountOrders As Long
    Dim lngOccorrenza As Long
    Dim strProgressivo As String
    Dim strProgressivoOld As String
    
    'Testata
    Dim strNote2_1 As String
    Dim strNote2_2 As String
    Dim strCodiceCLHtrasportatore As String
    Dim strRiferimentoEsterno1 As String
    Dim strCodDestinazione As String
    Dim strRagSocDestinazione As String
    Dim strIndirizzo As String
    Dim strLocalità As String
    Dim varCAP As Variant
    Dim varNote2_3 As Variant
    Dim varGPSX As Variant
    Dim varGPSY As Variant
    Dim varNIF As Variant
    Dim strConsegnatario As String
    Dim strConsegnatarioTelefono As String
    Dim strDescrizionePorto As String
    Dim strFlagControlloImmagine As String
    Dim strFlagControlloPrecincto As String
    Dim strTelefono1 As String
    Dim strTelefono2 As String
    Dim strTelefonoMovil As String
    Dim varAzioneCons As Variant
    Dim varConfermare As Variant
    Dim varOrarioApertura As Variant
    
    'Dettaglio
    Dim strCodiceOperatore As String
    Dim strFiller3 As String
    Dim strCodiceTrasportatore As String
    Dim strFiller6 As String
    Dim strCodiceDestinazione As String
    Dim strCodiceProdottoXSocietà As String
    Dim varQuantità As Variant
    Dim strUM As String
    Dim varDataConsRichiesta As Variant
    Dim strDummyData As String
    Dim strPriorità As String
    Dim strCodDepositoXSocietà As String
    Dim varFiller14 As Variant
    Dim strDescrizione As String
    Dim strTipoServizio As String
    Dim strFiller16 As String
    Dim strNoteBunker As String
    Dim strFlagVentaEnRuta As String
    '--JG 23/3/09
    Dim strparticion As String
    Dim Modcantidad As String
    Dim Contatore As Integer
    Dim blnEOF As Boolean
    Dim lngPosition As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    ' ** Estrazione del progressivo file
    lngOccorrenza = InStr(1, strFileName, ".", vbTextCompare)
    If lngOccorrenza > 0 Then
        For i = 1 To lngOccorrenza - 1
            strProgressivoOld = strProgressivo
            strProgressivo = Mid(strFileName, lngOccorrenza - i, i)
            If Not IsNumeric(strProgressivo) Then
                strProgressivo = strProgressivoOld
                Exit For
            End If
        Next i
    End If
    If IsNumeric(strProgressivo) Then
        lblprogres(6) = strProgressivo
    Else
        lblprogres(6) = vbNullString
    End If
    ' **

    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    For i = 0 To 5
        lblField(i) = vbNullString
    Next i
    cmdImport.Enabled = True
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    Contatore = 0
    strLine = vbNullString
    lngPosition = 1
    'Do While Not mobjInFile.EOF
    '    varBuffer = mobjInFile.GetLine(False)
    Do While Not mobjInFile.EOF
        varBuffer = mobjInFile.GetLine(False)
'        varBuffer = mobjInFile.GetLineUNIX(False, , lngPosition, strLine, blnEOF)
        If (Left(varBuffer, 2) = "10") Or (Left(varBuffer, 2) = "11") Then
            If Left(varBuffer, 2) = "10" Then
                Contatore = Contatore + 1
                If Not mobjInFile.ParseFixedLenFields(CStr(varBuffer), 1, 27) Then
                    GoTo ErrHandler
                End If
                strNote2_1 = mobjInFile.Item("2").value
                strNote2_2 = mobjInFile.Item("3").value
                strCodiceCLHtrasportatore = mobjInFile.Item("4").value
                strRiferimentoEsterno1 = mobjInFile.Item("6").value
                strCodDestinazione = mobjInFile.Item("7").value
                strRagSocDestinazione = mobjInFile.Item("8").value
                strIndirizzo = mobjInFile.Item("9").value
                strLocalità = mobjInFile.Item("10").value
                varCAP = mobjInFile.Item("11").value
                varNote2_3 = mobjInFile.Item("12").value
                If Len(mobjInFile.Item("13").value) > 0 Then
                    strFlagControlloImmagine = mobjInFile.Item("13").value
                Else
                    strFlagControlloImmagine = vbNullString
                End If
                If Len(mobjInFile.Item("20").value) > 0 Then
                    strFlagControlloPrecincto = mobjInFile.Item("20").value
                Else
                    strFlagControlloPrecincto = vbNullString
                End If
                varGPSX = mobjInFile.Item("14").value
                varGPSY = mobjInFile.Item("15").value
                strConsegnatario = mobjInFile.Item("16").value
                strConsegnatarioTelefono = mobjInFile.Item("17").value
                strDescrizionePorto = mobjInFile.Item("18").value
                varNIF = mobjInFile.Item("21").value
                
                '-- DP 13/08/08
                strTelefono1 = mobjInFile.Item("22").value
                strTelefono2 = mobjInFile.Item("23").value
                strTelefonoMovil = mobjInFile.Item("24").value
                varAzioneCons = mobjInFile.Item("25").value
                varConfermare = mobjInFile.Item("26").value
                varOrarioApertura = mobjInFile.Item("27").value
            Else
                If Not mobjInFileDettaglio.ParseFixedLenFields(CStr(varBuffer), 1, 21) Then
                    GoTo ErrHandler
                End If
                strCodiceOperatore = mobjInFileDettaglio.Item("2").value
                strFiller3 = mobjInFileDettaglio.Item("3").value
                strCodiceTrasportatore = mobjInFileDettaglio.Item("4").value
                strFiller6 = mobjInFileDettaglio.Item("6").value
                strCodiceDestinazione = mobjInFileDettaglio.Item("7").value
                strCodiceProdottoXSocietà = mobjInFileDettaglio.Item("8").value
                If Len(Trim(Left(mobjInFileDettaglio("17").value, 3))) > 0 Then
                    strCodiceProdottoXSocietà = strCodiceProdottoXSocietà & Left(mobjInFileDettaglio("17").value, 3)
                End If
                varQuantità = mobjInFileDettaglio.Item("9").value
                strUM = mobjInFileDettaglio.Item("10").value
                strDummyData = mobjInFileDettaglio.Item("11").value
                varDataConsRichiesta = Right(strDummyData, 2) + "/" + Mid(strDummyData, 5, 2) + "/" + Left(strDummyData, 4)
                strPriorità = mobjInFileDettaglio.Item("12").value
                strCodDepositoXSocietà = mobjInFileDettaglio.Item("13").value
                varFiller14 = mobjInFileDettaglio.Item("14").value
                strDescrizione = mobjInFileDettaglio.Item("15").value
                strTipoServizio = Space(3)
                LSet strTipoServizio = mobjInFileDettaglio.Item("20").value
                strDescrizione = strTipoServizio & strDescrizione
                strFiller16 = mobjInFileDettaglio.Item("16").value
                strNoteBunker = mobjInFileDettaglio.Item("18").value
                strFlagVentaEnRuta = mobjInFileDettaglio.Item("19").value
                '-- JG 23/3/09
                strparticion = mobjInFileDettaglio.Item("21").value
                Modcantidad = mobjInFileDettaglio.Item("22").value
                With XA
                    .ReDim 1, lngRows, 0, 43
                    .value(lngRows, 0) = Contatore
                    .value(lngRows, 1) = strNote2_1
                    .value(lngRows, 2) = strNote2_2
                    .value(lngRows, 3) = strCodiceCLHtrasportatore
                    .value(lngRows, 4) = strRiferimentoEsterno1
                    .value(lngRows, 5) = strCodDestinazione
                    .value(lngRows, 6) = strRagSocDestinazione
                    .value(lngRows, 7) = strIndirizzo
                    .value(lngRows, 8) = strLocalità
                    .value(lngRows, 9) = varCAP
                    .value(lngRows, 10) = varNote2_3
                    .value(lngRows, 11) = varGPSX
                    .value(lngRows, 12) = varGPSY
                    .value(lngRows, 13) = strConsegnatario
                    .value(lngRows, 14) = strConsegnatarioTelefono
                    .value(lngRows, 15) = strDescrizionePorto
                    .value(lngRows, 16) = strCodiceOperatore
                    .value(lngRows, 17) = strFiller3
                    .value(lngRows, 18) = strCodiceTrasportatore
                    .value(lngRows, 19) = strFiller6
                    .value(lngRows, 20) = strCodiceDestinazione
                    .value(lngRows, 21) = strCodiceProdottoXSocietà
                    .value(lngRows, 22) = CDbl(varQuantità)
                    .value(lngRows, 23) = strUM
                    .value(lngRows, 24) = varDataConsRichiesta
                    .value(lngRows, 25) = strPriorità
                    .value(lngRows, 26) = strCodDepositoXSocietà
                    .value(lngRows, 27) = varFiller14
                    .value(lngRows, 28) = strDescrizione
                    .value(lngRows, 29) = strFiller16
                    .value(lngRows, 30) = strNoteBunker
                    .value(lngRows, 31) = strFlagControlloImmagine
                    .value(lngRows, 32) = strFlagControlloPrecincto
                    .value(lngRows, 33) = strFlagVentaEnRuta
                    .value(lngRows, 34) = varNIF
                    
                    '-- DP 13/08/08
                    .value(lngRows, 35) = strTelefono1
                    .value(lngRows, 36) = strTelefono2
                    .value(lngRows, 37) = strTelefonoMovil
                    
                    '-- DP 01/12/08
                    .value(lngRows, 38) = Left(strRiferimentoEsterno1, 10)
                    .value(lngRows, 39) = varAzioneCons
                    .value(lngRows, 40) = varConfermare
                    .value(lngRows, 41) = varOrarioApertura
                    
                    ' -- JG 23/3/9
                    .value(lngRows, 42) = strparticion
                    .value(lngRows, 43) = Modcantidad
                End With
                
                With XALog
                    .ReDim 1, lngRows, 0, 10  '2
                    .value(.UpperBound(1), COL_COD_ORD_LOG) = Contatore '11)
                    .value(.UpperBound(1), COL_COD_RIF1_LOG) = strRiferimentoEsterno1 '11)
                    .value(.UpperBound(1), COL_COD_PROD_LOG) = strCodiceProdottoXSocietà '16)
                    
                    ' -- DP 01/12/2008
                    .value(.UpperBound(1), COL_COD_IDADE_SOC) = strCodDepositoXSocietà
                    .value(.UpperBound(1), COL_RIFE1) = Left(strRiferimentoEsterno1, 10)
                End With
                
                lngRows = lngRows + 1
            End If
        End If
    Loop
    
    ' -- DP 01/12/2008
    ' -- Ordino gli ordini per RIFE1 e IDADE
    XA.QuickSort 1, XA.UpperBound(1), _
                 38, XORDER_ASCEND, XTYPE_DEFAULT, _
                 26, XORDER_ASCEND, XTYPE_DEFAULT
                 
    XALog.QuickSort 1, XALog.UpperBound(1), _
                 COL_RIFE1, XORDER_ASCEND, XTYPE_DEFAULT, _
                 COL_COD_IDADE_SOC, XORDER_ASCEND, XTYPE_DEFAULT
    
    With TDBGrid1
        .Array = XA
        .ReBind
        .bookmark = 1
    End With
    
    With tdbLog
        .Array = XALog
        .ReBind
        .bookmark = 1
    End With
    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows - 1
    lblLog(6) = 0
    lblField(4) = 0
    lblField(5) = 0
        
        
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmOCTImport2", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmOCTImport2", _
                            "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub


Private Sub RecordToADZ(ADZ As CDestinazione, strIDDestinazione, IDClienteDefault As String, _
                        IDDeposito As String, strRagSocDestinazione As String, _
                        strIndirizzo As String, strLocalità As String, varCAP As Variant, _
                        IDTipoDestinazione As Variant, GPSX As String, GPSY As String, _
                        NIF As Variant, strTel1 As String, strTel2 As String, strTelMov As String, _
                        OrarioApertura As Variant)

    Dim varApertura1 As Variant
    Dim varApertura2 As Variant
    Dim varApertura3 As Variant
    Dim varApertura4 As Variant
    
    With ADZ.DS.Fields
'        .ResetFields
        .Item("0").NewValue = strIDDestinazione 'codice destinazione
        .Item("0").Updated = True
        If Len(strRagSocDestinazione) > 0 Then
            .Item("1").NewValue = strRagSocDestinazione 'ragione sociale destinazione
            .Item("1").Updated = True
        End If
        .Item("4").NewValue = strIndirizzo 'indirizzo
        .Item("4").Updated = True
        .Item("5").NewValue = strLocalità 'località
        .Item("5").Updated = True
        .Item("7").NewValue = varCAP 'cap
        .Item("7").Updated = True
        .Item("8").NewValue = IDClienteDefault 'cod. cliente
        .Item("8").Updated = True
        .Item("12").NewValue = IDTipoDestinazione
        .Item("12").Updated = True
        .Item("15").NewValue = IDDeposito 'cod. deposito x Società
        .Item("15").Updated = True
        .Item("16").NewValue = "00"
        .Item("16").Updated = True
        'If Len(GPSX) > 0 Then
        '    .Item("31").NewValue = GPSX
        '    .Item("31").Updated = True
        'End If
        'If Len(GPSY) > 0 Then
        '    .Item("32").NewValue = GPSY
        '    .Item("32").Updated = True
        'End If
        .Item("10").NewValue = NIF
        .Item("10").Updated = True
        
        '-- DP 14/08/08
        .Item("39").NewValue = strTel1
        .Item("39").Updated = True
        .Item("40").NewValue = strTel2
        .Item("40").Updated = True
        .Item("41").NewValue = strTelMov
        .Item("41").Updated = True
        
        varApertura1 = Mid(OrarioApertura, 1, 4)
        varApertura2 = Mid(OrarioApertura, 5, 4)
        varApertura3 = Mid(OrarioApertura, 9, 4)
        varApertura4 = Mid(OrarioApertura, 13, 4)
        
        If Len(varApertura1) > 0 Then
            If varApertura1 = "2400" Then varApertura1 = "2359"
            .Item("33").NewValue = Left(varApertura1, 2) & ":" & Right(varApertura1, 2)
            .Item("33").Updated = True
        End If
        If Len(varApertura2) > 0 Then
            If varApertura2 = "2400" Then varApertura2 = "2359"
            .Item("34").NewValue = Left(varApertura2, 2) & ":" & Right(varApertura2, 2)
            .Item("34").Updated = True
        End If
        If Len(varApertura3) > 0 Then
            If varApertura3 = "2400" Then varApertura3 = "2359"
            .Item("35").NewValue = Left(varApertura3, 2) & ":" & Right(varApertura3, 2)
            .Item("35").Updated = True
        End If
        If Len(varApertura4) > 0 Then
            If varApertura4 = "2400" Then varApertura4 = "2359"
            .Item("36").NewValue = Left(varApertura4, 2) & ":" & Right(varApertura4, 2)
            .Item("36").Updated = True
        End If
    End With
    With ADZ.DS.MultiSoc.DS.Fields
        .ResetFields
        .Item("Codice").NewValue = strIDDestinazione
        .Item("Codice").Updated = True
        .Item("Società").NewValue = ID_SOC_CLH
        .Item("Società").Updated = True
        .Item("CodPerSoc").NewValue = strIDDestinazione
        .Item("CodPerSoc").Updated = True
    End With
End Sub

Private Sub RecordToOCR(OCR As COrdineriga, ByVal strIDOrder As String, _
                        ByVal strIDRiga As String, ByVal strIDProdotto As String, _
                        ByVal strIDImballoDefault As String, ByVal strIDPosDefault As String, _
                        ByVal strDescrizione As String, ByVal varQuantità As Variant, RifEsterno As String, _
                        NoteBunker As String, IDUDM As String, ModCant As String)
    
    Dim dblQta As Double

    With OCR.DS.Fields
        .ResetFields
        .Item("0").NewValue = strIDOrder
        .Item("0").Updated = True
        .Item("1").NewValue = strIDRiga
        .Item("1").Updated = True
        .Item("2").NewValue = strIDProdotto
        .Item("2").Updated = True
        .Item("3").NewValue = strIDImballoDefault
        .Item("3").Updated = True
        .Item("4").NewValue = strIDPosDefault
        .Item("4").Updated = True
        .Item("5").NewValue = strDescrizione 'descrizione
        .Item("5").Updated = True
        .Item("7").NewValue = varQuantità 'qta in Litri
        .Item("7").Updated = True
        .Item("25").NewValue = RifEsterno
        .Item("25").Updated = True
        .Item("30").NewValue = NoteBunker
        .Item("30").Updated = True
        .Item("32").NewValue = IDUDM
        .Item("32").Updated = True
        If Trim(ModCant) <> "X" Then
        .Item("26").NewValue = "SI"
        .Item("26").Updated = True
        Else
        .Item("26").NewValue = "NO"
        .Item("26").Updated = True
        End If
        
        
        
    End With
End Sub

Private Sub RecordToOCT(OCT As COrdine, IDDeposito As String, IDDestinazione As String, _
                        IDEnteEmittente As String, IDCliente As String, _
                        ByVal pRifEsterno1 As String, ByVal pDtConsRich As Variant, _
                        ByVal pPriorita As String, ByVal pNote2 As String, _
                        ByVal strIDZona As String, ByVal strIDTipoVendita As String, _
                        ByVal pNote As String, ByVal pRifEsterno3 As String, _
                        ByVal varOrari As String, ByVal TipoOrdine As Variant, _
                        ByVal CondTecnico As String, GPSX As String, GPSY As String, _
                        Consegnatario As String, TelefonoConsegnatario As String, _
                        NomeNave As String, FlagImmagine As String, _
                        FlagPrecincto As String, FlagVentaEnRuta As String, _
                        AzioneConsentita As Variant, _
                        ContattareCliente As Variant, _
                        OrarioApertura As Variant, _
                        strPedEventuales As String, _
                        strparticion As String, strPresidio As String)
    
    Dim varOrario1 As Variant
    Dim varOrario2 As Variant
    Dim varOrario3 As Variant
    Dim varOrario4 As Variant
    Dim varApertura1 As Variant
    Dim varApertura2 As Variant
    Dim varApertura3 As Variant
    Dim varApertura4 As Variant
    
    With OCT.DS.Fields
        .ResetFields
        
        .Item("1").NewValue = MaskCodSoc
        .Item("1").Updated = True
        .Item("2").NewValue = IDEnteEmittente
        .Item("2").Updated = True
        .Item("3").NewValue = Format(Date, "DD/MM/YYYY")    ' Data ricezione
        .Item("3").Updated = True
        .Item("4").NewValue = Format(Time, "HH:MM")         ' Ora ricezione
        .Item("4").Updated = True
        .Item("5").NewValue = Format(Date, "DD/MM/YYYY")    ' Data inserimento
        .Item("5").Updated = True
        .Item("7").NewValue = pDtConsRich
        .Item("7").Updated = True
        .Item("8").NewValue = pDtConsRich
        .Item("8").Updated = True
        .Item("10").NewValue = pRifEsterno1
        .Item("10").Updated = True
        .Item("12").NewValue = IDCliente
        .Item("12").Updated = True
        .Item("13").NewValue = IDDestinazione
        .Item("13").Updated = True
        .Item("14").NewValue = Mid(pNote, 4)
        .Item("14").Updated = True
        .Item("15").NewValue = pNote2
        .Item("15").Updated = True
        .Item("16").NewValue = strIDTipoVendita
        .Item("16").Updated = True
        .Item("17").NewValue = IDDeposito
        .Item("17").Updated = True
        .Item("18").NewValue = strIDZona
        .Item("18").Updated = True
        .Item("20").NewValue = "GC"
        .Item("20").Updated = True
        .Item("21").NewValue = TipoOrdine
        .Item("21").Updated = True
        .Item("25").NewValue = pPriorita
        .Item("25").Updated = True
        
        '-- JG 03/03/09
        .Item("26").NewValue = strPedEventuales
        .Item("26").Updated = True
         
         '-- JG 25/03/09
         ' Inserta particion
        .Item("29").NewValue = strparticion
        .Item("29").Updated = True

        
        varOrario1 = Trim(Mid(varOrari, 1, 4))
        varOrario2 = Trim(Mid(varOrari, 5, 4))
        varOrario3 = Trim(Mid(varOrari, 9, 4))
        varOrario4 = Trim(Mid(varOrari, 13, 4))
        If Len(varOrario1) > 0 Then
            If varOrario1 = "2400" Then varOrario1 = "2359"
            .Item("38").NewValue = Left(varOrario1, 2) & ":" & Right(varOrario1, 2)
            .Item("38").Updated = True
        End If
        If Len(varOrario2) > 0 Then
            If varOrario2 = "2400" Then varOrario2 = "2359"
            .Item("39").NewValue = Left(varOrario2, 2) & ":" & Right(varOrario2, 2)
            .Item("39").Updated = True
        End If
        If Len(varOrario3) > 0 Then
            If varOrario3 = "2400" Then varOrario3 = "2359"
            .Item("40").NewValue = Left(varOrario3, 2) & ":" & Right(varOrario3, 2)
            .Item("40").Updated = True
        End If
        If Len(varOrario4) > 0 Then
            If varOrario4 = "2400" Then varOrario4 = "2359"
            .Item("41").NewValue = Left(varOrario4, 2) & ":" & Right(varOrario4, 2)
            .Item("41").Updated = True
        End If
        .Item("44").NewValue = pRifEsterno3
        .Item("44").Updated = True
        
        .Item("46").NewValue = CondTecnico
        .Item("46").Updated = True
        
        .Item("47").NewValue = GPSX
        .Item("47").Updated = True
        .Item("48").NewValue = GPSX
        .Item("48").Updated = True
        .Item("49").NewValue = Consegnatario
        .Item("49").Updated = True
        .Item("50").NewValue = TelefonoConsegnatario
        .Item("50").Updated = True
        .Item("51").NewValue = NomeNave
        .Item("51").Updated = True
        If Len(FlagImmagine) > 0 Then
            .Item("53").NewValue = FlagImmagine
            .Item("53").Updated = True
        End If
        If Len(FlagPrecincto) > 0 Then
            .Item("54").NewValue = FlagPrecincto
            .Item("54").Updated = True
        End If
        .Item("56").NewValue = FlagVentaEnRuta
        .Item("56").Updated = True
        .Item("57").NewValue = Left(pNote, 3)
        .Item("57").Updated = True
        
        .Item("58").NewValue = AzioneConsentita
        .Item("58").Updated = True
        .Item("59").NewValue = ContattareCliente
        .Item("59").Updated = True
        If AzioneConsentita = "0" Then
            .Item("60").NewValue = "0"
            .Item("60").Updated = True
        End If
        varApertura1 = Mid(OrarioApertura, 1, 4)
        varApertura2 = Mid(OrarioApertura, 5, 4)
        varApertura3 = Mid(OrarioApertura, 9, 4)
        varApertura4 = Mid(OrarioApertura, 13, 4)
        
        If Len(varApertura1) > 0 Then
            If varApertura1 = "2400" Then varApertura1 = "2359"
            .Item("61").NewValue = Left(varApertura1, 2) & ":" & Right(varApertura1, 2)
            .Item("61").Updated = True
        End If
        If Len(varApertura2) > 0 Then
            If varApertura2 = "2400" Then varApertura2 = "2359"
            .Item("62").NewValue = Left(varApertura2, 2) & ":" & Right(varApertura2, 2)
            .Item("62").Updated = True
        End If
        If Len(varApertura3) > 0 Then
            If varApertura3 = "2400" Then varApertura3 = "2359"
            .Item("63").NewValue = Left(varApertura3, 2) & ":" & Right(varApertura3, 2)
            .Item("63").Updated = True
        End If
        If Len(varApertura4) > 0 Then
            If varApertura4 = "2400" Then varApertura4 = "2359"
            .Item("64").NewValue = Left(varApertura4, 2) & ":" & Right(varApertura4, 2)
            .Item("64").Updated = True
        End If
        .Item("67").NewValue = strPresidio
        .Item("67").Updated = True
    End With
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    
    Loading = True
    
    ConfigureObjects Me
    
    Select Case mvarFunction
        Case fcIMPORD3
            mskFileName = gmRegistry.GetSetting("Options", "OrderImportFile03", , True)
            mstrUltimoProgressivo = gmRegistry.GetSetting("Options", "PrevOrderImportFileID3", "0", True)
            MaskCodSoc = ID_SOC_GUARNER
            MaskCodSoc.Enabled = False
            lblsocieta(8).Caption = DescrizioneSocieta(MaskCodSoc)

        Case Else
            MaskCodSoc = gmRegistry.GetSetting("Options", "Socimportfile2", , True)
            mskFileName = gmRegistry.GetSetting("Options", "OrderImportFile02", , True)
            mstrUltimoProgressivo = gmRegistry.GetSetting("Options", "PrevOrderImportFileID", "0", True)
            lblsocieta(8).Caption = DescrizioneSocieta(MaskCodSoc)
    End Select
        
    Set ErrH = New CErrorHandler
    
    Set AXX = New CParametro
    With AXX.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCT = New COrdine
    With OCT.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineriga
    With OCR.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AEE = New CEnteEmittente
    With AEE.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ATV = New CTipoVendita
    With ATV.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APD = New CPosDoganale
    With APD.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ADE = New CDeposito
    With ADE.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AZO = New CZona
    With AZO.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ASC = New CSocieta
    With ASC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    lblsocieta(8).Caption = DescrizioneSocieta(MaskCodSoc)
    
    Set mobjInFile = New CTextFile

    With mobjInFile
        .Add "Filler", 2, "1"
        .Add "Codice CLH operador", 6, "2"
        .Add "Codice CLH suministrador", 6, "3"
        .Add "Codice CLH del trasportatore", 4, "4"
        .Add "Filler", 2, "5"
        .Add "Riferimento Esterno1", 13, "6"
        .Add "Codice Destinazione", 8, "7"
        .Add "Ragione Sociale Destinazione", 20, "8"
        .Add "Indirizzo", 35, "9"
        .Add "Località", 20, "10"
        .Add "CAP", 5, "11", , fftNumerico
        .Add "Cond. tecnicos", 5, "12", , fftNumerico
        .Add "Immagine", 1, "13"
        .Add "Precincto", 1, "20"
        .Add "GPS X", 9, "14"
        .Add "GPS Y", 9, "15"
        .Add "Nome Consegnatario", 40, "16"
        .Add "Telefono Consegnatario", 10, "17"
        .Add "Descrizione porto (bunker)", 35, "18"
        .Add "NIF", 11, "21"
        
        '-- DP 13/08/08
        .Add "Telefono 1", 9, "22"
        .Add "Telefono 2", 9, "23"
        .Add "Telefono Movil", 9, "24"
        
        .Add "Azione consentita", 1, "25"
        .Add "Conferma cliente", 1, "26"
        .Add "Orario apertura destinatario", 16, "27"
        
        .Add "Filler", 20, "19"
    End With


    Set mobjInFileDettaglio = New CTextFile
    With mobjInFileDettaglio
        .Add "Tipo record", 2, "1"
        .Add "Codice Operatore", 6, "2"
        .Add "Filler3", 6, "3"
        .Add "Codice Trasportatore", 4, "4"
        .Add "Filler5", 2, "5"
        .Add "Filler6", 13, "6"
        .Add "CLH Destinazione", 8, "7"
        .Add "Prodotto", 7, "8"
        .Add "Quantità", 9, "9"
        .Add "U.M.", 1, "10"
        .Add "Data Consegna Richiesta", 8, "11"
        .Add "Priorità", 1, "12"
        .Add "Deposito", 4, "13"
        .Add "Filler14", 5, "14"
        .Add "Tipo Servizio", 3, "20"
'        .Add "Descrizione", 20, "15"
        .Add "Descrizione", 17, "15"
        .Add "Orari", 16, "16"
'        .Add "Filler17", 16, "17"
        .Add "Filler17", 33, "17"
        .Add "Note Bunker", 144, "18"
        .Add "Venta en ruta", 1, "19"
        '--JG 23/03/09
        .Add "Particion", 2, "21"
        .Add "NoModificarCantidad", 1, "22"
        .Add "Relleno", 72, "23"
    End With

    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Select Case mvarFunction
        Case fcIMPORD3
            gmRegistry.SaveSetting "Options", "Socimportfile3", MaskCodSoc, True
            gmRegistry.SaveSetting "Options", "OrderImportFile03", mskFileName, True
            gmRegistry.SaveSetting "Options", "PrevOrderImportFileID3", mstrUltimoProgressivo, True
 
        Case Else
            gmRegistry.SaveSetting "Options", "Socimportfile2", MaskCodSoc, True
            gmRegistry.SaveSetting "Options", "OrderImportFile02", mskFileName, True
            gmRegistry.SaveSetting "Options", "PrevOrderImportFileID", mstrUltimoProgressivo, True
    End Select
    
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set ADE = Nothing
    Set AZO = Nothing
    Set OCR = Nothing
    Set OCT = Nothing
    Set APR = Nothing
    Set APD = Nothing
    Set ATV = Nothing
    Set AEE = Nothing
    Set ASC = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set mobjInFileDettaglio = Nothing
    Set frmOCTImport2 = Nothing
End Sub

Private Sub MaskCodSoc_LostFocus()
    lblsocieta(8).Caption = DescrizioneSocieta(MaskCodSoc)
'    Select Case MaskCodSoc
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

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub

