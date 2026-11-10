VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTImporttq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion Ordenes Transquality"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10350
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   21
      Top             =   5760
      Width           =   10335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   8640
         Picture         =   "frmOCTImporttq.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "frmOCTImporttq.frx":0102
         TabIndex        =   3
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6540
      Width           =   10350
      _ExtentX        =   18256
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
            Object.Width           =   12806
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18:01"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "02/02/2011"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   6360
         TabIndex        =   11
         Top             =   120
         Width           =   3855
         Begin VB.Label lblprogres 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblLog 
            Caption         =   "Albaranes excluidos"
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   31
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   30
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   29
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Albaranes a importar"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   28
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   27
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errores encontrados"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   26
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   25
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Albaranes importados"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   24
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ora Invio"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Invio"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Progressivo Invio"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Ente Emittente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   6135
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmOCTImporttq.frx":0204
            TabIndex        =   33
            Tag             =   "SELECT"
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   10095
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmOCTImporttq.frx":86A8
            TabIndex        =   32
            Tag             =   "SELECT"
            Top             =   240
            Width           =   9855
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
         Left            =   3120
         Picture         =   "frmOCTImporttq.frx":12D5A
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1110
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmOCTImporttq.frx":1328C
         TabIndex        =   2
         Tag             =   "IMPORT"
         Top             =   1035
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
         Left            =   1680
         Picture         =   "frmOCTImporttq.frx":1338E
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1110
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmOCTImporttq.frx":138C0
         TabIndex        =   1
         Tag             =   "OPEN"
         Top             =   1035
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
         Left            =   240
         Picture         =   "frmOCTImporttq.frx":139C2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1110
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTImporttq.frx":13AC4
         TabIndex        =   0
         Tag             =   "BROWSE"
         Top             =   1035
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
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
               Picture         =   "frmOCTImporttq.frx":13BC6
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImporttq.frx":13CD8
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImporttq.frx":13DEA
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File &Tickets"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOCTImporttq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'05/09/2003 CZucchini
'Interfaccia Ricezione Tickets
Option Explicit

Private Const MAX_COLS = 3
Private Const COL_COD_ORD_LOG = 1
Private Const COL_COD_PROD_LOG = 4
Private Const COL_RESULT_LOG = 10

Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineRiga
Private APR As CProdotto
Private APD As CPosDoganale
Private ADE As CDeposito
Private AEE As CEnteEmittente
Private ATV As CTipoVendita
Private AZO As CZona

Private mobjInFile As CTextFile
Private mstrIDEnte As String
Private mstrUltimoProgressivo As String
Private mstrIDSoc As String

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
                                      ByVal strCodDest As String, _
                                      ByVal strIDClienteDefault As String, _
                                      ByVal strIDDeposito As String, _
                                      ByRef strNewLocalita As String, ByVal strNewDirec As String, strCP As String, strnombrecli As String) As Long
    Dim intRowError As Errors
    Dim strLocalita As String
    Dim strNewZona As String
    Dim strNewIndirizzo As String
    Dim blnConcurErr As Boolean
    Dim lngKmRaggio As Long
    Dim rowdes As Variant
    Dim lngdes As Long

Dim DA As CDataAccess
Set DA = DataConnect.DA

'Variables a definir
Dim IDClienteDefault As String
Dim IDDeposito As String
Dim strRagSocDestinazione As String
Dim strIndirizzo As String
Dim varCAP As String
Dim IDTipoDestinazione As String
Dim GPSX, GPSY As String


    intRowError = NoError
    ' --- Cerco la destinazione
    strCodDest = strIDDestinazione
    strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC_TQ, strCodDest, True)
'    If DA.QueryCommand("SELECT * FROM ANAGADZ0 WERE ADZIDADZ = ?", rowdes, lngdes, strIDDestinazione) Then
'    End If
'
'    ' --- La destinazione è già stata inserita ?
'    If lngdes = 0 Then
If strIDDestinazione = vbNullString Then
            If RecordToADZ(ADZ, strCodDest, strnombrecli, ID_SOC_TQ, strIDDeposito, strNewDirec, strNewLocalita, strCP, strIDDeposito) Then
            End If
           
        ' --- No : devo inserirla ora
        With ADZ.DS.Fields
        
            If Not ADZ.Insert Then
                intRowError = DestInsertError
                GoTo ExitFunction
            End If
            If Not ADZ.DS.MultiSoc.Insert Then
                intRowError = DestSocInsertError
                GoTo ExitFunction
            End If
        End With
        strIDDestinazione = strCodDest
    End If
  
ExitFunction:
    CheckDestinazione001 = intRowError
End Function

'Private Function CheckOrderRow001(ByRef strIDDeposito As String, _
'                                  ByRef strIDEnteEmittente As String, _
'                                  ByRef strIDProdotto As String, _
'                                  ByRef strIDTipoVendita As String, _
'                                  ByRef strIDZona As String, _
'                                  ByRef strCodDest As String) As Long
'    Dim intRowError As Errors
'
'    intRowError = NoError
'
'    ' --- Controllo che esistano Deposito, Zona, Tipo Vendita, Ente Emittente, Prodotto
'    strIDDeposito = ADE.DecodeFirmCode2(ID_SOC, mobjInFile("13").value, True)
'    strIDEnteEmittente = AEE.DecodeFirmCode(ID_SOC, mstrIDEnte, True)
'    strIDProdotto = APR.DecodeFirmCode(ID_SOC, mobjInFile("17").value, True)
'
'    If mobjInFile("19").value = vbNullString Then
'        intRowError = NoTipoVendita
'        GoTo ExitFunction
'    End If
'    strIDTipoVendita = ATV.DecodeFirmCode(ID_SOC, mobjInFile("19").value, True)
'    If strIDTipoVendita = vbNullString Then
'        intRowError = NoTipoVendita
'        GoTo ExitFunction
'    End If
'    strIDZona = mobjInFile("2").value
'    AZO.DS.Fields("Codice").FindValue = strIDZona
'    If Not AZO.Find Then
'        intRowError = NoZona
'        GoTo ExitFunction
'    End If
'    If AZO.DS.HowMany = 0 Then
'        intRowError = NoZona
'        GoTo ExitFunction
'    End If
'    ' --- Controllo che la zona non sia da escludere
'    ' --- dall'importazione
'    If AZO.DS.Fields("2").value <> 0 Then
'        intRowError = ZonaEsclusaImport
'        GoTo ExitFunction
'    End If
'    ' ---
'
'    If strIDDeposito = vbNullString Then
'        intRowError = NoDeposito
'        GoTo ExitFunction
'    End If
'    If strIDEnteEmittente = vbNullString Then
'        intRowError = NoEnteEmittente
'        GoTo ExitFunction
'    End If
'    If strIDProdotto = vbNullString Then
'        intRowError = NoProdotto
'        GoTo ExitFunction
'    End If
'    ' ---
'
'    strCodDest = Format$(mobjInFile("4").value, "00000000")
'    If strCodDest = vbNullString Then
'        intRowError = NoDestinazione
'        GoTo ExitFunction
'    End If
'
'
'ExitFunction:
'    CheckOrderRow001 = intRowError
'End Function

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
                .Filter = "(*.txt)|*.txt|Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .DialogTitle = "Seleziona il file dei tickets... "
                .ShowOpen
                mskFileName = .fileName
            End With
        Case fcImport
            ImportRecords001
            cmdImport.Enabled = False
        Case fcOpen
            OpenTextFile001
        Case fcExit
            Unload Me
    End Select
End Sub

Private Function UpdOrdine(ByVal strRifEsterno1 As String, _
                           ByVal strTicket As String, _
                           ByVal strRifEsternoRiga As String, _
                           ByRef strIDOCT As String, _
                           Optional NumPartizione As String = "00") As Boolean
    Const SEL_ORDINE = _
        "SELECT OCTIDOCT " & _
        "  FROM ORDIOCT0 INNER JOIN ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT " & _
        " WHERE OCTRIFE1 = ? AND OCTIDASC = ? AND OCTANNUL = 0 " & _
        "       AND OCRRIFE2 = ? AND OCRANNUL = 0 "
    Const SEL_ORDINE_2 = " AND CONVERT(INT,OCTRAUT4) = "
    
    Const UPD_ORDINE = _
        "UPDATE ORDIOCT0 " & _
        "   SET OCTNOTA3 = ? " & _
        " WHERE OCTIDOCT = ? "

    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
'    Dim strIDOCT As String
    
    strSQL = SEL_ORDINE
    If NumPartizione <> "00" Then
        strSQL = strSQL & SEL_ORDINE_2 & OCT.DS.DA.MakeSQL(NumPartizione, 1)
    End If
    If Not OCT.DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strRifEsterno1, mstrIDSoc, strRifEsternoRiga) Then
        GoTo ErrHandler
    End If
    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If
    strIDOCT = varBuffer(0, 0)
    
    strSQL = UPD_ORDINE
    If Not OCT.DS.DA.ExecuteCommand(strSQL, strTicket, strIDOCT) Then
        GoTo ErrHandler
    End If
    UpdOrdine = True

Exit Function
ErrHandler:
    UpdOrdine = False
End Function

Private Function CheckOrderRow001(ByRef strIDDeposito As String, _
                                  ByRef strIDEnteEmittente As String, _
                                  ByRef strIDProdotto As String, _
                                  ByRef strIDZona As String, _
                                  ByRef strIDTipoVendita As String) As Long



    
    Dim intRowError As Errors
    Dim Row As Long
    intRowError = NoError
    
    ' --- Controllo che esistano Deposito, Ente Emittente, Prodotto
    strIDDeposito = ADE.DecodeFirmCode2(ID_SOC_TQ, strIDDeposito, True)
    strIDProdotto = APR.DecodeFirmCode(ID_SOC_TQ, strIDProdotto, True)
    strIDZona = "00"
    strIDTipoVendita = "F2"
    
    If strIDDeposito = vbNullString Then
        intRowError = NoDeposito
        GoTo ExitFunction
    End If
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
'
Private Sub ImportRecords001()
    On Error GoTo ErrHandler
    Dim strOutFileName As String
    Dim strHeader As String
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
   ' Dim strIDProdotto As String
    Dim strIndirizzo As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strNewLocalita As String
    Dim strCP As String
    Dim strnombrecli As String
    Dim stralbaran As String
    Dim strproducto As String
    Dim strclifactu, strdireccion, strenteemitente, strcodpropcont, strpedido, strobserva, strpoblacion, strpropcont, strfecha, strcantidad, strcliente, strfactoria As String
    Dim strLocalita As String
    Dim varCAP As Variant
    Dim strErrorMessage As String
    Dim strRagSocDestinazione As String
    Dim strCodDepositoXSocietà As String
    Dim strRifEsternoOrdine As String
    Dim strRifEsternoRiga As String
    Dim strProgressivoInvio As String
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
    Dim blnSkipInsertRows As Boolean
    Dim blnFlagBunker As Boolean
    ReDim ListNewLoc(0 To 0) As String
    

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
        
        strIDClienteDefault = ID_SOC_TQ
        
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

    
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        strRifEsternoOrdine = XA(Row, 1)
        strRotturaOrd = strRifEsternoOrdine 'XA(Row, 0)
        'La variabile strRotturaOrd serve per capire se si deve inserire testata e riga dettaglio o solo riga dettaglio
        If strRotturaOrd <> strIDPrevOrder Then
            ' --- Nuovo ordine : Inizio con l'inserimento di una nuova testata e riga ordine
            blnNewOrder = True
            lngTotQuantità = 0
            blnSkipInsertRows = False
            strIDPrevOrder = strRotturaOrd
            strIDRiga = 10
            strNote2 = XA(Row, 1) & XA(Row, 2) & Format(XA(Row, 10), "00000")
            strIDEnteEmittente = "004"
'
           ' strIDProdotto = XA.value(Row, 2)
            
            
            strfecha = XA(Row, 0)
            stralbaran = XA(Row, 1)
            strproducto = XA(Row, 2)
            strcantidad = XA(Row, 3)
            strIDDeposito = XA.value(Row, 4)
            strIDDestinazione = "TQ" + Left(XA(Row, 5), 8)
            strnombrecli = XA(Row, 6)
            strdireccion = XA(Row, 7)
            strNewLocalita = XA.value(Row, 8)
            strobserva = XA(Row, 9)
            strclifactu = XA(Row, 10)
            strpedido = XA(Row, 11)
            strpropcont = XA(Row, 12)
            strcodpropcont = XA(Row, 13)
            strenteemitente = XA(Row, 14)
            strCP = XA(Row, 15)

            'strRagSocDestinazione = XA.value(Row, 6)
            'strIndirizzo = XA.value(Row, 7)
            
            
            intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, _
                                   strproducto, strIDZona, strIDTipoVendita)
            intRowError = CheckDestinazione001(strIDDestinazione, strCodDest, strIDClienteDefault, strIDDeposito, strNewLocalita, strdireccion, strCP, strnombrecli)

            'intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, strIDProdotto, strIDZona, strIDTipoVendita)
            intRowError = CheckAlbaran(stralbaran, strproducto)
            If intRowError <> NoError Then
                GoTo LogResult
            End If
        
            If (intRowError <> NoError) And _
               (intRowError <> DestHasChanged) Then
                GoTo LogResult
            End If
'            If OCT.CheckSocCode1("TQ", strRifEsternoOrdine) Then
'                blnSkipInsertRows = True
'                intRowError = OrderAlreadyExisting
'                GoTo LogResult
'            Else
                strDescrizione = XA.value(Row, 9)
'                strProgressivoInvio = lblprogres(6)
'                strCondTecnico = XA(Row, 23)
'                For k = 1 To 5
'                    If Not IsNumeric(Mid(strCondTecnico, k, 1)) Then
'                        Mid(strCondTecnico, k, 1) = "0"
'                        'XA(Row, 22) =  & "0"
'                    End If
'                Next k
'                If Not IsNumeric(strCondTecnico) Then
'                    intRowError = InvalidCode
'                    GoTo LogResult
'                End If
                
                'viene inserita la testata
                
                
                
                
                
'                RecordToOCT OCT, strIDDeposito, strIDDestinazione, strIDEnteEmittente, strIDClienteDefault, _
'                            strRifEsternoOrdine, XA(Row, 20), XA(Row, 21), strNote2, strIDZona, strIDTipoVendita, _
'                            strDescrizione, strProgressivoInvio, XA(Row, 25), strIDTipoOrdine, strCondTecnico, _
'                            strGPSX, strGPSY, strConsegnatario, strTelefonoConsegnatario, strNomeNave

                
                RecordToOCT OCT, strIDDeposito, strIDDestinazione, XA(Row, 14), ID_SOC_TQ, _
                            XA(Row, 1), XA(Row, 0), "", XA(Row, 10), "00", "TD", _
                            XA(Row, 9), XA(Row, 11), "", "E", "", _
                            "", "", "", "", "", XA(Row, 1)
                            

                If Not OCT.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
                strIDOrder = OCT.DS.Fields("0").NewValue
                If strIDOrder = vbNullString Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
           
                If IsNumeric(strcantidad) Then
                    lngTotQuantità = lngTotQuantità + CLng(varQuantità)
                End If

                RecordToOCR OCR, strIDOrder, strIDRiga, strproducto, strIDImballoDefault, _
                            strIDPosDefault, strDescrizione, strcantidad, strRifEsternoRiga, strNoteBunker
                If Not OCR.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
'                End If
            End If
        Else
                strIDDeposito = XA.value(Row, 4)
                strIDRiga = strIDRiga + 10
                strcantidad = XA.value(Row, 3)
                If IsNumeric(strcantidad) Then
                    lngTotQuantità = lngTotQuantità + CLng(strcantidad)
                End If
                strproducto = XA(Row, 2)
                
                intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, _
                                   strproducto, strIDZona, strIDTipoVendita)
                
                intRowError = CheckAlbaran(stralbaran, strproducto)
                If intRowError <> NoError Then
                    GoTo LogResult
                End If

                RecordToOCR OCR, strIDOrder, strIDRiga, strproducto, strIDImballoDefault, _
                            strIDPosDefault, strDescrizione, strcantidad, strRifEsternoRiga, strNoteBunker
                 If Not OCR.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
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
        
        Select Case intRowError
            Case NoError
                If lngTotQuantità > 32000 Then
                    strErrorMessage = "Quantità > 32000 - OK - Num. Ordine : " & strIDOrder
                Else
                    strErrorMessage = "OK - Num. Ordine : " & strIDOrder
                End If
            Case DestHasChanged
                blnLocHasChanged = True
                ListNewLoc(UBound(ListNewLoc)) = "Dest. modificata: " & strIDDestinazione & " " & strNewLocalita
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder & _
                                  " - " & ListNewLoc(UBound(ListNewLoc))
                ReDim Preserve ListNewLoc(UBound(ListNewLoc) + 1) As String
            Case OrderAlreadyExisting
                strErrorMessage = "Ordine già inserito"
            Case NoTipoVendita
                strErrorMessage = "Codice Tipo Vendita non trovato"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Codice Ente Emittente non trovato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione mancante"
            Case BadIDOrderSoc
                strErrorMessage = "Numero Ordine mancante"
            Case DestInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case OrderInsertError
                strErrorMessage = "Errore durante l'inserimento dell'ordine "
            Case RowOrderInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di ordine"
            Case ZonaEsclusaImport
                strErrorMessage = "Zona esclusa dall'importazione ordini"
            Case InvalidCode
                strErrorMessage = "Condizionante tecnico non numerico"
        End Select
        
'        With mobjOutFile
'            .ClearFields
'            .Item("1").value = strRifEsternoOrdine
'            .Item("2").value = XA(Row, 5)
'            .Item("3").value = strRagSocDestinazione
'            .Item("4").value = XA.value(Row, 17)
'            .Item("5").value = Format$(XA.value(Row, 18), "000000")
'            .Item("6").value = Format(XA(Row, 20), "yyyymmdd")
'            .Item("7").value = XA.value(Row, 24) 'GridArraySelect.value(lngRow, 5)
'            .Item("8").value = vbNullString
'            If Not .WriteLine(True) Then
'            End If
'        End With
        
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
        Case InvalidCode
            Set C = ErrH.SetError(FileNotFound, "frmOCTImport4", _
                    "ImportRecords", Err.Number, "Errore nell'apertura del file.", 10, "Nome del file: " & strOutFileName)
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
'    mobjOutFile.CloseFile
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub
 Private Function CheckAlbaran(Albaran As String, Producto As String) As String
 Dim rowalb As Variant
 Dim lngalb As Long
 Dim sql As String
 sql = "SELECT * FROM ORDIOCT0 INNER JOIN ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT WHERE ORDIOCT0.OCTRIFE1=? AND ORDIOCR0.OCRIDAPR= ?"
    If Not (ADZ.DS.DA.QueryCommand(sql, rowalb, lngalb, Albaran, Producto)) Then
        CheckAlbaran = LoadError
    End If
    
    If lngalb > 0 Then
        CheckAlbaran = OrderAlreadyExisting
    
    Else
        CheckAlbaran = NoError
    End If
    
    
 End Function


Private Sub RecordToOCT(OCT As COrdine, IDDeposito As String, IDDestinazione As String, _
                        IDEnteEmittente As String, IDCliente As String, _
                        ByVal pRifEsterno1 As String, ByVal pDtConsRich As Variant, _
                        ByVal pPriorita As String, ByVal pNote2 As String, _
                        ByVal strIDZona As String, ByVal strIDTipoVendita As String, _
                        ByVal pNote As String, ByVal pRifEsterno3 As String, _
                        ByVal varOrari As String, ByVal TipoOrdine As Variant, _
                        ByVal CondTecnico As String, GPSX As String, GPSY As String, Consegnatario As String, _
                        TelefonoConsegnatario As String, NomeNave As String, Albaran As String)
    
    Dim varOrario1 As Variant
    Dim varOrario2 As Variant
    Dim varOrario3 As Variant
    Dim varOrario4 As Variant
    
    With OCT.DS.Fields
        .ResetFields
        
        .Item("1").NewValue = ID_SOC_TQ
        .Item("1").Updated = True
        .Item("2").NewValue = IDEnteEmittente
        .Item("2").Updated = True
        .Item("3").NewValue = Format(Date, "DD/MM/YYYY")    ' Data ricezione
        .Item("3").Updated = True
        .Item("4").NewValue = Format(Now, "hh:mm")          ' Ora ricezione
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
        .Item("14").NewValue = pNote
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
        .Item("36").NewValue = Albaran
        .Item("36").Updated = True

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
        
    End With
End Sub
 
Private Sub RecordToOCR(OCR As COrdineRiga, ByVal strIDOrder As String, _
                        ByVal strIDRiga As String, ByVal strIDProdotto As String, _
                        ByVal strIDImballoDefault As String, ByVal strIDPosDefault As String, _
                        ByVal strDescrizione As String, ByVal varQuantità As Variant, RifEsterno As String, _
                        NoteBunker As String)
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
        .Item("5").NewValue = Left(strDescrizione, 30) 'descrizione
        .Item("5").Updated = True
        .Item("7").NewValue = varQuantità 'qta in Litri
        .Item("7").Updated = True
        .Item("25").NewValue = RifEsterno
        .Item("25").Updated = True
        .Item("30").NewValue = NoteBunker
        .Item("30").Updated = True
    End With
End Sub

 
Private Function RecordToADZ(ADZ As CDestinazione, strIDDestinazione As String, strRAZsoc As String, IDClienteDefault As String, _
                         Deposito As String, Direcion As String, Localidad As String, CP As String, IDTipoDestinazione As String)

    With ADZ.DS.Fields
        .ResetFields
        .Item("0").NewValue = strIDDestinazione 'codice destinazione
        .Item("0").Updated = True
        .Item("1").NewValue = strRAZsoc 'rAZONSOCIAL
        .Item("1").Updated = True
        .Item("4").NewValue = Direcion
        .Item("4").Updated = True
        .Item("5").NewValue = Localidad
        .Item("5").Updated = True
        .Item("7").NewValue = CP 'Codifgo Postal
        .Item("7").Updated = True
        .Item("8").NewValue = IDClienteDefault 'cliente (TQ = 899)
        .Item("8").Updated = True
        .Item("12").NewValue = "E" 'tIPODESTINACIÓN
        .Item("12").Updated = True
        .Item("15").NewValue = IDTipoDestinazione 'DEPOSITO
        .Item("15").Updated = True
        .Item("16").NewValue = "00"  'ZONA
        .Item("16").Updated = True
    End With
    ' no tengo muy claro como va es esto....
    With ADZ.DS.MultiSoc.DS.Fields
        .ResetFields
        .Item("Codice").NewValue = strIDDestinazione
        .Item("Codice").Updated = True
        .Item("Società").NewValue = ID_SOC_TQ
        .Item("Società").Updated = True
        .Item("CodPerSoc").NewValue = strIDDestinazione
        .Item("CodPerSoc").Updated = True
    End With
End Function


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
    '***05/09/2003 CZ
    Dim strfecha As String
    Dim stralbaran As String
    Dim strproducto As String
    Dim strcantidad As String
    Dim strfactoria As String
    Dim strcliente As String
    Dim strnombrecli As String
    Dim strdireccion As String
    Dim strpoblacion As String
    Dim strobserva As String
    Dim strclifactu As String
    Dim strpedido As String
    Dim strpropcont As String
    Dim strcodpropcont As String
    Dim strDummyData As String
    Dim strenteemitente As String
    Dim strCP As String
    
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
        lblprogres(4) = strProgressivo
    Else
        lblprogres(4) = vbNullString
    End If
    ' **
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    For i = 0 To 3
        lblField(i) = vbNullString
    Next i
    cmdImport.Enabled = True
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    Do While Not mobjInFile.EOF
        varBuffer = mobjInFile.GetLine(False)
        If Not mobjInFile.ParseFixedLenFields(CStr(varBuffer), 1, 16) Then
                GoTo ErrHandler
                End If
            strfecha = mobjInFile.Item("1").value
            stralbaran = mobjInFile.Item("2").value
            strproducto = mobjInFile.Item("3").value
            strcantidad = mobjInFile.Item("4").value
            strfactoria = mobjInFile.Item("5").value
            strcliente = mobjInFile.Item("6").value
            strnombrecli = mobjInFile.Item("7").value
            strdireccion = mobjInFile.Item("8").value
            strpoblacion = mobjInFile.Item("9").value
            strobserva = mobjInFile.Item("10").value
            strclifactu = mobjInFile.Item("11").value
            strpedido = mobjInFile.Item("12").value
            strpropcont = mobjInFile.Item("13").value
            strcodpropcont = mobjInFile.Item("14").value
            strenteemitente = mobjInFile.Item("15").value
            strCP = mobjInFile.Item("16").value
            
            With XA
                .ReDim 1, lngRows, 0, 15
                .value(lngRows, 0) = strfecha
                .value(lngRows, 1) = stralbaran
                .value(lngRows, 2) = strproducto
                .value(lngRows, 3) = strcantidad
                .value(lngRows, 4) = strfactoria
                .value(lngRows, 5) = strcliente
                .value(lngRows, 6) = strnombrecli
                .value(lngRows, 7) = strdireccion
                .value(lngRows, 8) = strpoblacion
                .value(lngRows, 9) = strobserva
                .value(lngRows, 10) = strclifactu
                .value(lngRows, 11) = strpedido
                .value(lngRows, 12) = strpropcont
                .value(lngRows, 13) = strcodpropcont
                .value(lngRows, 14) = strenteemitente
                .value(lngRows, 15) = strCP
            End With
            
            With XALog
                .ReDim 1, lngRows, 0, 10 '2
                .value(.UpperBound(1), COL_COD_ORD_LOG) = stralbaran
                .value(.UpperBound(1), COL_COD_PROD_LOG) = strproducto
            End With
            lngRows = lngRows + 1
       
        
    Loop
    
    XA.QuickSort 1, XA.UpperBound(1), _
                 1, XORDER_ASCEND, XTYPE_DEFAULT, _
                 6, XORDER_ASCEND, XTYPE_DEFAULT, _
                 14, XORDER_ASCEND, XTYPE_DEFAULT, _
                 5, XORDER_ASCEND, XTYPE_DEFAULT
    XALog.QuickSort 1, XALog.UpperBound(1), _
                 0, XORDER_ASCEND, XTYPE_DEFAULT, _
                 1, XORDER_ASCEND, XTYPE_DEFAULT, _
                 2, XORDER_ASCEND, XTYPE_DEFAULT, _
                 3, XORDER_ASCEND, XTYPE_DEFAULT
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
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmOCTImporttq", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmOCTImporttq", _
                            "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    
    ConfigureObjects Me
    
    Select Case mvarFunction
        Case fcRICTICK
            mskFileName = gmRegistry.GetSetting("Options", "Importarfictq", , True)
            mstrIDSoc = ID_SOC_GUARNER
            Caption = "GUARNER - RICEZIONE TICKETS"
        
        Case Else
            mskFileName = gmRegistry.GetSetting("Options", "Importarfictq", , True)
            mstrIDSoc = ID_SOC_CEPSA
    End Select

    Set ErrH = New CErrorHandler
    
    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineRiga
    With OCR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
'    With mobjInFile
'        .Add "Tipo Record", 2, "1"
'        .Add "Filler2", 4, "2"
'        .Add "Operatore", 6, "3"
'        .Add "Filler4", 6, "4"
'        .Add "Rif. Esterno1", 13, "5"
'        .Add "Partizione", 2, "6"
'        .Add "Codice Destinazione", 8, "7"
'        .Add "Codice del prodotto", 7, "8"
'        .Add "Q.ta", 9, "9"
'        .Add "Unità di misura", 1, "10"
'        .Add "Codice veicolo", 11, "11"
'        .Add "Numero viaggio", 2, "12"
'        .Add "Filler13", 8, "13"
'        .Add "Ticket", 15, "14"
'        .Add "Codice Additivo", 3, "15"
'        .Add "Nome commerciale se additivato", 30, "16"
'    End With
    With mobjInFile
        .Add "FECHA", 19, "1"
        .Add "ALBARAN", 22, "2"
        .Add "CODPRODUCTO", 6, "3"
        .Add "CANTIDAD", 22, "4"
        .Add "FACTORIA", 4, "5"
        .Add "CLIENTE", 15, "6"
        .Add "nombe2", 50, "7"
        .Add "direccion", 50, "8"
        .Add "poblacion", 30, "9"
        .Add "observaciones", 50, "10"
        .Add "CLIENTEFACT", 50, "11"
        .Add "Pedido", 20, "12"
        .Add "Propcontrato", 50, "13"
        .Add "Codpropcontrato", 10, "14"
        .Add "Codenteemitente", 3, "15"
        .Add "Codigo Postal", 8, "16"
    End With
    
    
    Loading = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Select Case mvarFunction
        Case fcRICTICK
            gmRegistry.SaveSetting "Options", "Importarfictq", mskFileName, True
        
        Case Else
            gmRegistry.SaveSetting "Options", "Importarfictq", mskFileName, True
    End Select
    
    'gmRegistry.SaveSetting "Options", "OrderImportFileDetail", mskFileNameDettaglio, True
    'gmRegistry.SaveSetting "Options", "PrevOrderImportFileID", mstrUltimoProgressivo, True
    
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set OCT = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set frmOCTImporttq = Nothing
End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub




