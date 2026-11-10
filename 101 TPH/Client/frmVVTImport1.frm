VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTImport1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE VIAGGI"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   10350
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   6180
      Width           =   10335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmVVTImport1.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
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
         Picture         =   "frmVVTImport1.frx":0102
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
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
         Picture         =   "frmVVTImport1.frx":0634
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   8640
         Picture         =   "frmVVTImport1.frx":0B66
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "frmVVTImport1.frx":0C68
         TabIndex        =   5
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmVVTImport1.frx":0D6A
         TabIndex        =   19
         Tag             =   "IMPORT"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmVVTImport1.frx":0E6C
         TabIndex        =   21
         Tag             =   "OPEN"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTImport1.frx":0F6E
         TabIndex        =   23
         Tag             =   "BROWSE"
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6960
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
            TextSave        =   "16.19"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "06/08/2008"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   6195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   120
         TabIndex        =   24
         Top             =   180
         Width           =   10095
         Begin MSMask.MaskEdBox mskFileName 
            Height          =   255
            Left            =   1380
            TabIndex        =   31
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskCodSoc 
            Height          =   255
            Left            =   6240
            TabIndex        =   33
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "So&cietà"
            Height          =   255
            Index           =   5
            Left            =   5280
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblSocieta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   6840
            TabIndex        =   34
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "File &Viaggi"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Progressivo Invio"
            Height          =   255
            Index           =   2
            Left            =   5280
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   6660
            TabIndex        =   29
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Invio"
            Height          =   255
            Index           =   3
            Left            =   7920
            TabIndex        =   28
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   8880
            TabIndex        =   27
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ente Emittente"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   1380
            TabIndex        =   25
            Top             =   600
            Width           =   3615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   10095
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            Caption         =   "Esclusi"
            Height          =   255
            Index           =   7
            Left            =   6120
            TabIndex        =   16
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   7080
            TabIndex        =   15
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   1740
            TabIndex        =   14
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblLog 
            Caption         =   "Viaggi da importare"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   9060
            TabIndex        =   12
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            Caption         =   "In errore"
            Height          =   255
            Index           =   1
            Left            =   8160
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   10
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            Caption         =   "Importati correttamente"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   9
            Top             =   240
            Width           =   1755
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   1740
         Width           =   10095
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1875
            Left            =   120
            OleObjectBlob   =   "frmVVTImport1.frx":1070
            TabIndex        =   17
            Top             =   180
            Width           =   9855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   3900
         Width           =   10095
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmVVTImport1.frx":6658
            TabIndex        =   0
            Top             =   240
            Width           =   9855
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10020
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
               Picture         =   "frmVVTImport1.frx":11466
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTImport1.frx":11578
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTImport1.frx":1168A
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmVVTImport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private ASC As CSocieta
Private ADE As CDeposito
Private AEE As CEnteEmittente
'Private ATV As CTipoVendita
'Private AZO As CZona

Private ACL As CCliente
Private AUT As CUnitaDiTrasporto
Private AMD As CMagazzinoDeposito
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto

Private mobjInFile As CTextFile
Private mobjInFileHeader As CTextFile
Private mobjInFileFooter As CTextFile
Private mobjInFileDetail As CTextFile

Private mstrIDEnte As String
Private mintFileInput As Integer
Private mstrUltimoProgressivo As String

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mvarFunction As FunctionCodes

Private Enum fldLog
    fldLog_ViaggioCounter = 0
    fldLog_DataViaggio = 1
    fldLog_Badge = 2
    fldLog_SequenzaViaggio = 3
    fldLog_OraConsegna = 4
    fldLog_Destinazione = 5
    fldLog_Prodotto = 6
    fldLog_Esito = 7
End Enum

Private Enum fldVVTImport
    fldVVTImport_ViaggioCounter = 0
    fldVVTImport_TipoRecord = 1
    fldVVTImport_Note2_1 = 2
    fldVVTImport_Note2_2 = 3
    fldVVTImport_EnteEmittente = 4
    fldVVTImport_Filler1 = 5
    fldVVTImport_ConsegnaRifEsterno1 = 6
    fldVVTImport_Destinazione = 7
    fldVVTImport_Prodotto = 8
    fldVVTImport_QtaInizialeOrdine = 9
    fldVVTImport_UnitaMisura = 10
    fldVVTImport_DataConsegnaRichiestaOrdine = 11
    fldVVTImport_OrdinePriorita = 12
    fldVVTImport_Deposito = 13
    fldVVTImport_Nota2_3 = 14
    fldVVTImport_ConsegnaRifEsterno2 = 15
    fldVVTImport_Partizione = 16
    fldVVTImport_QtaTrasportata = 17
    fldVVTImport_DataViaggio = 18
    fldVVTImport_Badge = 19
    fldVVTImport_SequenzaViaggio = 20
    fldVVTImport_CIAutista = 21
    fldVVTImport_OraConsegnaPrevista = 22
End Enum

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
                .Filter = "(*.dat)|*.dat|Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .DialogTitle = "Seleziona il file dei viaggi... "
                
                .ShowOpen
                mskFileName = .fileName
            End With
        
        Case fcImport
            ImportRecords
            cmdImport.Enabled = False
        
        Case fcOpen
            OpenTextFile
        
        Case fcPrint
            If State = fsEdit Then
            End If
        
        Case fcExit
            Unload Me
    End Select
End Sub

Private Function ToDate(strDate As String) As Variant
On Error GoTo ErrHandler
    If Len(strDate) = 0 Then
        ToDate = Empty
        Exit Function
    End If
    
    If Len(strDate) <> 8 Then
        ToDate = Empty
        Exit Function
    End If

    ToDate = DateSerial(CInt(Left(strDate, 4)), CInt(Mid(strDate, 5, 2)), CInt(Right(strDate, 2)))
    Exit Function
ErrHandler:
    ToDate = Empty
End Function

Private Sub ImportRecords()
    On Error GoTo ErrHandler

    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDOrder As String
    Dim strIDOrderAGIP As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strNewLocalita As String
    Dim strErrorMessage As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim i As Integer
    Dim blnNewOrder As Boolean
    Dim blnLocHasChanged As Boolean
    Dim blnDeleteOrder As Boolean
    Dim blnOrdineDaAggiornare As Boolean
    Dim blnConcErr As Boolean
    Dim blnNotUpdatable As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    ReDim ListNewLoc(0 To 0) As String
    
    Dim strKeyViaggio As String
    Dim strPrecKeyViaggio As String
    Dim strKeyConsegna As String
    Dim strPrecKeyConsegna As String
    Dim varOrdineBuffer As Variant
    Dim lngHowMany As Long
    Dim SQL_GetOrdine As String
    Dim SQL_GetOrdinePartizione As String
    Dim SQL_UpdateAvanzamentoOrdineRiga As String

    Dim strIDADE As String
    Dim strIDAMD As String
    Dim strIDAUT As String
    Dim varDataViaggio As Variant
    Dim strIDSocieta As String
    Dim strIDCliente As String
    Dim StrIDADZ As String
    Dim lngConsegnaCounter As Long
    Dim lngConsegnaProdottoCounter As Long
    Dim StrIDAPR As String
    Dim strExtraErrorMessage As String
    Dim strIDSocietaTrasc As String
    Dim strIDADZOrdine As String
    Dim blnVVTInError As Boolean
    Dim strIDViaggio As String
    Dim blnLastConsegna As Boolean
    Dim lngIDOrder As Long
    Dim strIDOCT As String
    Dim strNewIDOCT As String
    Dim strIDOCTOri As String
    Dim blnViaggioFound As Boolean

    Dim strOrdineRiferimentoEsterno As String
    Dim strRigaOrdineRiferimentoEsterno As String
    
    SQL_GetOrdine = "SELECT OCTIDOCT, OCTFLAOR FROM ORDIOCT0 WHERE OCTIDASC = ? AND OCTRIFE1 = ? AND OCTANNUL = 0 AND IsNull(OCTIDOC1, '') = ''"
    SQL_GetOrdinePartizione = "SELECT OCTIDOCT, OCTFLAOR FROM ORDIOCT0 WHERE OCTIDASC = ? AND OCTRIFE1 = ? AND OCTANNUL = 0 AND OCTRAUT4 = ? AND OCTIDOC1 = ?"
    SQL_UpdateAvanzamentoOrdineRiga = "UPDATE ORDIOCR0 SET OCRFLAVR = 'A' WHERE OCRIDOCT = ?"
    
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
        strIDClienteDefault = Trim$(.Fields("3").value)
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
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)

    strIDSocieta = MaskCodSoc
    strIDCliente = MaskCodSoc
    strIDSocietaTrasc = ID_SOC_CLH

    With ACL.DS
        .Fields.ResetFields
        .Field("Codice").FindValue = strIDCliente
        
        If Not ACL.Find() Or .HowMany = 0 Then
            ErrCode = NoCliente
            strExtraErrorMessage = "Cliente """ & strIDCliente & """ non trovato."
            GoTo ErrHandler
        End If
    End With
    
    ' --- Ciclo di acquisizione dati
    strPrecKeyViaggio = ""
    strPrecKeyConsegna = ""
    blnVVTInError = False
    For Row = lngStartRow To lngNumRows
        
        intRowError = NoError
        strKeyViaggio = GetKeyViaggio(Row)
        strExtraErrorMessage = ""
        
        strOrdineRiferimentoEsterno = Left(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 10)
        strRigaOrdineRiferimentoEsterno = Mid(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 11)

        If Row = lngNumRows Then
            blnLastConsegna = True
        Else
            If strKeyViaggio <> GetKeyViaggio(Row + 1) Then
                blnLastConsegna = True
            Else
                blnLastConsegna = False
            End If
        End If

        ' Nuovo viaggio
        If strPrecKeyViaggio <> strKeyViaggio Then
            
            DataConnect.DA.BeginTransaction

            strPrecKeyViaggio = strKeyViaggio
            blnVVTInError = False
            blnViaggioFound = False
            
            ' Ricerca UDT tramite Badge
            strIDAUT = AUT.DecodeFirmCode(strIDSocietaTrasc, XA.value(Row, fldVVTImport_Badge), True)
            If Len(strIDAUT) = 0 Then
                intRowError = noUT
                strExtraErrorMessage = "Badge """ & XA.value(Row, fldVVTImport_Badge) & """ non trovato."
                GoTo LogResult
            End If

            With AUT.DS
                .Fields.ResetFields
                .Field("Codice").FindValue = strIDAUT
                If Not AUT.Find() Or .HowMany = 0 Then
                    intRowError = noUT
                    strExtraErrorMessage = "Unità di trasporto """ & strIDAUT & """ non trovata."
                    GoTo LogResult
                End If
            End With
        
            ' Ricerca Deposito
            strIDADE = ADE.DecodeFirmCode(strIDSocietaTrasc, XA.value(Row, fldVVTImport_Deposito), True)
            If Len(strIDADE) = 0 Then
                intRowError = NoDeposito
                strExtraErrorMessage = "Deposito """ & XA.value(Row, fldVVTImport_Deposito) & """ non trovato."
                GoTo LogResult
            End If
            
            With AMD.DS
                .Fields.ResetFields
                .Field("Codice Deposito").FindValue = strIDADE
                AMD.GetList True
                strIDAMD = .List(0, 0)
            End With
            
            With VVT.DS
            
                .Fields.ResetFields
            
                .Field("Sequenza Viaggio").NewValue = XA.value(Row, fldVVTImport_SequenzaViaggio)
                .Field("Sequenza Viaggio").Updated = True

                .Field("Società").NewValue = strIDSocieta
                .Field("Società").Updated = True

                varDataViaggio = ToDate(XA.value(Row, fldVVTImport_DataViaggio))
                If Not IsDate(varDataViaggio) Then
                    intRowError = InvalidDate
                    strExtraErrorMessage = "Data viaggio """ & XA.value(Row, fldVVTImport_DataViaggio) & """ non valida."
                    GoTo LogResult
                End If

                .Field("Data inizio carico").NewValue = varDataViaggio
                .Field("Data inizio carico").Updated = True

                .Field("Data inizio viaggio").NewValue = varDataViaggio
                .Field("Data inizio viaggio").Updated = True

                .Field("Data fine viaggio").NewValue = varDataViaggio
                .Field("Data fine viaggio").Updated = True

                .Field("Unità di Trasporto").NewValue = strIDAUT
                .Field("Unità di Trasporto").Updated = True

                .Field("Motrice").NewValue = AUT.DS.Field("Motrice").value
                .Field("Motrice").Updated = True

                .Field("Rimorchio").NewValue = AUT.DS.Field("Rimorchio").value
                .Field("Rimorchio").Updated = True

                If Not IsEmptyVariantString(AUT.DS.Field("Vettore").value) Then
                    .Field("Vettore").NewValue = AUT.DS.Field("Vettore").value
                    .Field("Vettore").Updated = True
                End If

                If Not IsEmptyVariantString(AUT.DS.Field("Vettore Trasportatore").value) Then
                    .Field("Vettore Trasportatore").NewValue = AUT.DS.Field("Vettore Trasportatore").value
                    .Field("Vettore Trasportatore").Updated = True
                End If

                If Not IsEmptyVariantString(AUT.DS.Field("Autista 1").value) Then
                    .Field("Autista 1").NewValue = AUT.DS.Field("Autista 1").value
                    .Field("Autista 1").Updated = True
                End If

                If Not IsEmptyVariantString(AUT.DS.Field("Autista 1").value) Then
                    .Field("Autista 2").NewValue = AUT.DS.Field("Autista 2").value
                    .Field("Autista 2").Updated = True
                End If
                
                .Field("Flag calcolo km definitivi").NewValue = "M"
                .Field("Flag calcolo km definitivi").Updated = True

                .Field("Flag calcolo km previsti").NewValue = "M"
                .Field("Flag calcolo km previsti").Updated = True

                .Field("Ente Emittente").NewValue = "056"
                .Field("Ente Emittente").Updated = True

                .Field("Deposito").NewValue = strIDADE
                .Field("Deposito").Updated = True

                .Field("Magazzino di deposito").NewValue = strIDAMD
                .Field("Magazzino di deposito").Updated = True

                .Field("Tipo Viaggio").NewValue = "R"
                .Field("Tipo Viaggio").Updated = True

                .Field("Annullato").NewValue = 0
                .Field("Annullato").Updated = True

                '-- DP 05/08/08
                If Not VVT.CheckSequenza(.Fields("6").NewValue, _
                                        .Fields("2").NewValue, _
                                        .Fields("14").NewValue, _
                                        .Fields("22").NewValue, "") Then
                    ErrCode = InsertError
                    strExtraErrorMessage = "Il numero di sequenza è già stato utilizzato"
                    GoTo ErrHandler
                End If
            
                If Not VVT.Insert() Then
                    ErrCode = InsertError
                    strExtraErrorMessage = "Errore durante l'inserimento del viaggio."
                    GoTo ErrHandler
                End If
                strIDViaggio = VVT.DS.Field("Numero Viaggio").value
            End With
            
            strPrecKeyConsegna = ""
            lngConsegnaCounter = 1
        End If

        If Not blnVVTInError Then
            strKeyConsegna = GetKeyConsegna(Row)
            If strPrecKeyConsegna <> strKeyConsegna Then
    
                strPrecKeyConsegna = strKeyConsegna
                
                ' Ricerca Ordine
                lngHowMany = 0
                If Not OCT.DS.DA.QueryCommand(SQL_GetOrdine, varOrdineBuffer, lngHowMany, strIDSocieta, strOrdineRiferimentoEsterno) Then
                    strExtraErrorMessage = "Errore nella ricerca dell'ordine cliente """ & Left(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 10) & """."
                    GoTo ErrHandler
                End If
                
                ' Check Ordine
                If lngHowMany = 0 Then
                    strExtraErrorMessage = "Ordine """ & strOrdineRiferimentoEsterno & """ non trovato."
                    intRowError = NoOrderID
                    GoTo LogResult
                End If
                
                ' Check Avanzamento ordine
                If Not IsEmptyVariantString(varOrdineBuffer(1, 0)) And _
                    varOrdineBuffer(1, 0) <> "A" Then
                    strExtraErrorMessage = "Viaggio già inserito."
                    blnViaggioFound = True
                    intRowError = NoOrderID
                    GoTo LogResult
                End If
                

                strIDOCT = varOrdineBuffer(0, 0)
                strIDOCTOri = strIDOCT

                ' Viaggio con partizione
                If XA.value(Row, fldVVTImport_Partizione) <> "00" Then
                    
                    ' Ricerca Ordine con partizione
                    lngHowMany = 0
                    If Not OCT.DS.DA.QueryCommand(SQL_GetOrdinePartizione, varOrdineBuffer, lngHowMany, strIDSocieta, strOrdineRiferimentoEsterno, XA.value(Row, fldVVTImport_Partizione), strIDOCTOri) Then
                        strExtraErrorMessage = "Errore nella ricerca dell'ordine cliente """ & Left(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 10) & """."
                        ErrCode = SelectError
                        GoTo ErrHandler
                    End If
                
                    ' Check Ordine
                    If lngHowMany = 0 Then
                        
                        With OCT
                            ' Crea nuovo Ordine sulla base di quello trovato
                            If Not .CopyOrder(strIDOCT, , strNewIDOCT, True) Then
                                strExtraErrorMessage = "Errore nella creazione dell'ordine """ & Left(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 10) & """ per viaggio con partizione """ & XA.value(Row, fldVVTImport_Partizione) & """."
                                ErrCode = InsertError
                                GoTo ErrHandler
                            End If
    
                            strIDOCT = strNewIDOCT
                            .DS.Fields.ResetFields
                            .DS.Field("Numero Ordine").FindValue = strIDOCT
                            If Not .Find() Or .DS.HowMany = 0 Then
                                strExtraErrorMessage = "Errore nella creazione dell'ordine """ & Left(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 10) & """ per viaggio con partizione """ & XA.value(Row, fldVVTImport_Partizione) & """."
                                ErrCode = InsertError
                                GoTo ErrHandler
                            End If
                            
                            .DS.Field("Raggruppamento Utente 4").NewValue = XA.value(Row, fldVVTImport_Partizione)
                            .DS.Field("Raggruppamento Utente 4").Updated = True
                            
                            .DS.Field("Ordine di Riferimento").NewValue = strIDOCTOri
                            .DS.Field("Ordine di Riferimento").Updated = True

                            .DS.Field("Avanzamento Ordine").NewValue = "A"
                            .DS.Field("Avanzamento Ordine").Updated = True

                            If Not .BeginSave(True) Then
                                strExtraErrorMessage = "Errore nell'aggiornamento dell'ordine """ & strIDOCT & """."
                                ErrCode = UpdateError
                                GoTo ErrHandler
                            End If
                            .DS.DA.EndTrans ETCOMMIT
                            
                            ' Aggiorna lo stato della riga a Acquisito
                            If Not .DS.DA.ExecuteCommand(SQL_UpdateAvanzamentoOrdineRiga, strIDOCT) Then
                                strExtraErrorMessage = "Errore nell'aggiornamento dell'ordine """ & strIDOCT & """."
                                ErrCode = UpdateError
                                GoTo ErrHandler
                            End If
                        End With
                    Else
                        If Not IsEmptyVariantString(varOrdineBuffer(1, 0)) And _
                            varOrdineBuffer(1, 0) <> "A" Then
                            
                            strExtraErrorMessage = "Ordine """ & strOrdineRiferimentoEsterno & """ con partizione """ & XA.value(Row, fldVVTImport_Partizione) & """ non trovato."
                            intRowError = OrderInTrip
                            GoTo LogResult
                        End If
                    End If
                    
                    With OCT.DS
                        .Fields.ResetFields
                        .Field("Numero Ordine").FindValue = strIDOCT
                        OCT.Find
                    End With
                End If
                    
                OCT.DS.Field("Numero Ordine").FindValue = strIDOCT
                OCT.Find
                
                ' Check cliente ordine
                If OCT.DS.Field("Cliente").value <> strIDCliente Then
                    strExtraErrorMessage = "L'ordine cliente """ & Left(XA.value(Row, fldVVTImport_ConsegnaRifEsterno1), 10) & """ presenta un cliente diverso da """ & strIDCliente & """."
                    intRowError = BadIDOrderSoc
                    GoTo LogResult
                End If

                ' Ricerca Destinazione
                With ADZ.DS
                    .Fields.ResetFields
                    .Field("Codice").FindValue = OCT.DS.Field("Destinazione").value
                    ADZ.Find
'                    strIDADZOrdine = .MultiSoc.DS.GetFirmCode(strIDSocietaTrasc)
                End With
                
'                If Len(strIDADZOrdine) = 0 Then
'                    intRowError = NoDestinazione
'                    strExtraErrorMessage = "Destinazione """ & XA.value(Row, fldVVTImport_Destinazione) & """ non trovato."
'                    GoTo LogResult
'                End If
'
'                If Mid(strIDADZOrdine, 3) <> XA.value(Row, fldVVTImport_Destinazione) Then
'                    intRowError = NoDestinazione
'                    strExtraErrorMessage = "Destinazione """ & XA.value(Row, fldVVTImport_Destinazione) & """ non trovato."
'                    GoTo LogResult
'                End If
                
                ' Inserimento della consegna
                With VVC.DS
                    
                    .Fields.ResetFields
                    
                    .Field("Numero Viaggio").NewValue = VVT.DS.Field("Numero Viaggio").value
                    .Field("Numero Viaggio").Updated = True
                    
                    .Field("Numero Consegna").NewValue = Format(lngConsegnaCounter * 10, "000")
                    .Field("Numero Consegna").Updated = True
                    
                    .Field("Numero Sequenza").NewValue = .Field("Numero Consegna").NewValue
                    .Field("Numero Sequenza").Updated = True
                    
                    .Field("Società").NewValue = strIDSocieta
                    .Field("Società").Updated = True
                    
                    .Field("Cliente").NewValue = strIDCliente
                    .Field("Cliente").Updated = True
                    
                    .Field("Ragione Sociale").NewValue = ACL.DS.Field("Ragione Sociale").value
                    .Field("Ragione Sociale").Updated = True
                    .Field("Indirizzo").NewValue = ACL.DS.Field("Indirizzo").value
                    .Field("Indirizzo").Updated = True
                    .Field("Località").NewValue = ACL.DS.Field("Località").value
                    .Field("Località").Updated = True
                    .Field("Provincia").NewValue = ACL.DS.Field("Provincia").value
                    .Field("Provincia").Updated = True
                    .Field("CAP").NewValue = ACL.DS.Field("CAP").value
                    .Field("CAP").Updated = True
                    .Field("Codice Fiscale").NewValue = ACL.DS.Field("Codice Fiscale").value
                    .Field("Codice Fiscale").Updated = True
                    .Field("Partita IVA").NewValue = ACL.DS.Field("Partita IVA").value
                    .Field("Partita IVA").Updated = True
                    
                    .Field("Destinazione").NewValue = OCT.DS.Field("Destinazione").value
                    .Field("Destinazione").Updated = True
                    .Field("Ragione Sociale Destinazione").NewValue = ADZ.DS.Field("Ragione Sociale").value
                    .Field("Ragione Sociale Destinazione").Updated = True
                    .Field("Indirizzo Destinazione").NewValue = ADZ.DS.Field("Indirizzo").value
                    .Field("Indirizzo Destinazione").Updated = True
                    .Field("Località Destinazione").NewValue = ADZ.DS.Field("Località").value
                    .Field("Località Destinazione").Updated = True
                    .Field("Provincia Destinazione").NewValue = ADZ.DS.Field("Provincia").value
                    .Field("Provincia Destinazione").Updated = True
                    .Field("CAP Destinazione").NewValue = ADZ.DS.Field("CAP").value
                    .Field("CAP Destinazione").Updated = True
                    .Field("Codice Fiscale Destinazione").NewValue = ADZ.DS.Field("Codice Fiscale").value
                    .Field("Codice Fiscale Destinazione").Updated = True
                    .Field("Partita IVA Destinazione").NewValue = ADZ.DS.Field("Partita IVA").value
                    .Field("Partita IVA Destinazione").Updated = True
    
                    .Field("Zona").NewValue = OCT.DS.Field("Zona").value
                    .Field("Zona").Updated = True
                    
                    .Field("Ordine Cliente").NewValue = OCT.DS.Field("Numero Ordine").value
                    .Field("Ordine Cliente").Updated = True
    
                    .Field("Data consegna").NewValue = varDataViaggio
                    .Field("Data consegna").Updated = True
                    
                    .Field("Flag calcolo km previsti").NewValue = "M"
                    .Field("Flag calcolo km previsti").Updated = True
                    
                    .Field("Flag calcolo km definitivi").NewValue = "M"
                    .Field("Flag calcolo km definitivi").Updated = True
                    
                    .Field("Tipo Vendita").NewValue = OCT.DS.Field("Tipo Vendita").value
                    .Field("Tipo Vendita").Updated = True
                    
                    .Field("Riferimento Esterno 1").NewValue = OCT.DS.Field("Rif. Esterno 1").value
                    .Field("Riferimento Esterno 1").Updated = True
                    
                    .Field("Riferimento Esterno 2").NewValue = OCT.DS.Field("Rif. Esterno 2").value
                    .Field("Riferimento Esterno 2").Updated = True
    
                    .Field("Km raggio").NewValue = OCT.DS.Field("Km Raggio").value
                    .Field("Km raggio").Updated = True
    
                    .Field("Flag stato 3").NewValue = OCT.DS.Field("Raggruppamento Utente 3").value
                    .Field("Flag stato 3").Updated = True
    
                    .Field("Orario previsto").NewValue = XA.value(Row, fldVVTImport_OraConsegnaPrevista)
                    .Field("Orario previsto").Updated = True
                    
                    .Field("Orario arrivo").NewValue = .Field("Orario previsto").NewValue
                    .Field("Orario arrivo").Updated = True
                    
                    .Field("Annullato").NewValue = 0
                    .Field("Annullato").Updated = True
                    
                    If Not VVC.Insert(False) Then
                        ErrCode = InsertError
                        strExtraErrorMessage = "Errore durante l'inserimento della consegna."
                        GoTo ErrHandler
                    End If
                    lngConsegnaCounter = lngConsegnaCounter + 1
                End With
    
                lngConsegnaProdottoCounter = 1
            End If
            
            ' Inserimento della consegna prodotto
'            StrIDAPR = APR.DecodeFirmCode(strIDSocietaTrasc, XA.value(Row, fldVVTImport_Prodotto), True)
'            If Len(StrIDAPR) = 0 Then
'                intRowError = NoProdotto
'                strExtraErrorMessage = "Prodotto """ & XA.value(Row, fldVVTImport_Prodotto) & """ non trovato."
'                GoTo LogResult
'            End If

            ' Ricerca la riga ordine
            lngIDOrder = OCR.GetRowOrder(strIDOCT, strRigaOrdineRiferimentoEsterno, strIDSocietaTrasc, XA.value(Row, fldVVTImport_Prodotto), strExtraErrorMessage)
            If lngIDOrder = -1 Then
                intRowError = NoRigaOrdine
                GoTo LogResult
            End If
        
            With OCR.DS
                .Fields.ResetFields
                .Field("Numero Ordine").FindValue = OCT.DS.Field("Numero Ordine").value
                .Field("Numero Riga").FindValue = lngIDOrder
                
                If Not OCR.Find() Or .HowMany = 0 Then
                    intRowError = NoRigaOrdine
                    strExtraErrorMessage = "Prodotto """ & StrIDAPR & """ non trovato nell'ordine """ & OCT.DS.Field("Codice").value & """."
                    GoTo LogResult
                End If
            End With
            
            With VVP.DS
                .Fields.ResetFields
                
                .Field("Numero Viaggio").NewValue = VVT.DS.Field("Numero Viaggio").value
                .Field("Numero Viaggio").Updated = True
                
                .Field("Numero Consegna").NewValue = VVC.DS.Field("Numero Consegna").value
                .Field("Numero Consegna").Updated = True
                
                .Field("Numero Riga").NewValue = Format(lngConsegnaProdottoCounter * 10, "00")
                .Field("Numero Riga").Updated = True
                
                .Field("Prodotto").NewValue = OCR.DS.Field("Prodotto").value
                .Field("Prodotto").Updated = True
                
                .Field("Imballo").NewValue = OCR.DS.Field("Imballo").value
                .Field("Imballo").Updated = True
                
                .Field("Posizione Doganale").NewValue = OCR.DS.Field("Posizione Doganale").value
                .Field("Posizione Doganale").Updated = True
                
                .Field("Quantita Trasportata L").NewValue = CLng(XA.value(Row, fldVVTImport_QtaTrasportata))
                .Field("Quantita Trasportata L").Updated = True
                
                .Field("Quantita Consegnata L").NewValue = CLng(XA.value(Row, fldVVTImport_QtaTrasportata))
                .Field("Quantita Consegnata L").Updated = True
    
                .Field("Ordine Cliente").NewValue = OCT.DS.Field("Numero Ordine").value
                .Field("Ordine Cliente").Updated = True
                
                .Field("Riga Ordine Cliente").NewValue = lngIDOrder
                .Field("Riga Ordine Cliente").Updated = True
                
                .Field("Riferimento Esterno 2").NewValue = OCR.DS.Field("Riferimento Esterno 2").value
                .Field("Riferimento Esterno 2").Updated = True
    
                .Field("IVA").NewValue = OCR.DS.Field("IVA").value
                .Field("IVA").Updated = True
    
                .Field("Prezzo").NewValue = OCR.DS.Field("Prezzo").value
                .Field("Prezzo").Updated = True
    
                .Field("Sconto Percentuale").NewValue = OCR.DS.Field("Sconto Percentuale").value
                .Field("Sconto Percentuale").Updated = True
    
                .Field("Sconto Valore").NewValue = OCR.DS.Field("Sconto Valore").value
                .Field("Sconto Valore").Updated = True
    
                .Field("Tipo Riga").NewValue = OCR.DS.Field("Tipo Riga").value
                .Field("Tipo Riga").Updated = True
    
                .Field("Modalità di Pagamento").NewValue = OCR.DS.Field("Modalità di Pagamento").value
                .Field("Modalità di Pagamento").Updated = True
                
                .Field("Annullato").NewValue = 0
                .Field("Annullato").Updated = True
                    
                If Not VVP.Insert() Then
                    ErrCode = InsertError
                    strExtraErrorMessage = "Errore durante l'inserimento del prodotto consegna."
                    GoTo ErrHandler
                End If
            End With
    
            ' Aggiorna avanzamento riga ordine
            If Not OCR.SetAvanzamento("V") Then
                ErrCode = UpdateError
                strExtraErrorMessage = "Errore durante l'aggiornamento avanzamento riga ordine."
                GoTo ErrHandler
            End If
            
            lngConsegnaProdottoCounter = lngConsegnaProdottoCounter + 1
        End If
        
LogResult:
        
        If intRowError <> NoError Then
            blnVVTInError = True
        End If
        
        If blnLastConsegna Then
            If Not blnVVTInError Then
                ' Imposta l'avanzamento ordine
                If Not VVT.SetAvanzamentoOrdini("V") Then
                    ErrCode = UpdateError
                    strExtraErrorMessage = "Errore nell'impostare l'avanzamento ordine """ & OCT.DS.Field("Codice").value & """."
                    GoTo ErrHandler
                End If

                ' Imposta lo stato del viaggio a definitivo
                VVT.DS.Field("Numero Viaggio").FindValue = strIDViaggio
                If Not VVT.SetStato("D") Then
                    ErrCode = UpdateError
                    strExtraErrorMessage = "Errore durante l'aggiornamento dello stato viaggio a 'Definitivo'"
                    GoTo ErrHandler
                End If
                VVT.DS.DA.EndTrans ETCOMMIT
            Else
                VVT.DS.DA.EndTrans ETROLLBACK
            End If
        End If

        ' --- Scrivo il risultato dell'importazione
        If (intRowError <> NoError) Then
            If blnLastConsegna Then
                If blnViaggioFound Then
                    lngCountExcluded = lngCountExcluded + 1
                Else
                    lngCountFail = lngCountFail + 1
                End If
            End If
        Else
            If blnLastConsegna Then
                If Not blnVVTInError Then
                    lngCountOK = lngCountOK + 1
                Else
                    If blnViaggioFound Then
                        lngCountExcluded = lngCountExcluded + 1
                    Else
                        lngCountFail = lngCountFail + 1
                    End If
                End If
            End If
        End If

        ' lngCountExcluded = lngCountExcluded + 1
        If blnLastConsegna Then
            lngCountToGo = lngCountToGo - 1
        End If

        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(6) = lngCountExcluded

        Select Case intRowError
            Case NoError
                If Not blnVVTInError Then
                    If Row = lngNumRows Then
                        strErrorMessage = "OK - Num. Viaggio: " & VVT.DS.Field("Numero Viaggio").value
                    Else
                        If strKeyViaggio <> GetKeyViaggio(Row + 1) Then
                            strErrorMessage = "OK - Num. Viaggio: " & VVT.DS.Field("Numero Viaggio").value
                        Else
                            strErrorMessage = ""
'                            strErrorMessage = "Ok."
                        End If
                    End If
                Else
                    strErrorMessage = ""
                    'strErrorMessage = "Skip viaggio."
                End If
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
            Case NoOrderID
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
            Case BadIDOrderSoc
                strErrorMessage = "Ordine società non valido."
            Case OrderInTrip
                strErrorMessage = "Viaggio già inserito."
        End Select
        
        If strExtraErrorMessage <> vbNullString Then
            XALog.value(Row, fldLog.fldLog_Esito) = strExtraErrorMessage
        Else
            XALog.value(Row, fldLog.fldLog_Esito) = strErrorMessage
        End If

        With tdbLog
            .Refresh
            .bookmark = Row
        End With
        TDBGrid1.bookmark = Row
        ' ---

        DoEvents
    Next Row
    ' ---

    mstrUltimoProgressivo = lblField(0)
    
    GoTo EndProc

ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long

    DataConnect.DA.EndTrans ETROLLBACK
    
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
        
        Case InsertError
            strMessage = "Errore inserimento."
            strInfo = strExtraErrorMessage
        
        Case UpdateError
            strMessage = "Errore aggiornamento."
            strInfo = strExtraErrorMessage
        
        Case NoCliente
            strMessage = "Codice Cliente non trovato."
            strInfo = strExtraErrorMessage
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTImport", _
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

Private Sub OpenTextFile()
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
    Dim strPrecIdViaggio As String
    Dim strIDViaggio As String
    Dim lngCountViaggi As Long
    Dim lngCounter As Long
    Dim blnFoundHeader As Boolean
    Dim blnFoundFooter As Boolean
    Dim strExtraErrorMessage As String
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    For i = 0 To 2
        lblField(i) = vbNullString
    Next i
    
    cmdImport.Enabled = False
    
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1

    ' Legge Header
    blnFoundHeader = False
    blnFoundFooter = False
    Do While Not mobjInFile.EOF
        strExtraErrorMessage = ""
        Line Input #mobjInFile.FileNum, strLine

        Select Case Left(strLine, 2)

            ' Header
            Case "00"
'                If lngRows <> 1 Then
'                    ErrCode = InvalidHeader
'                    strExtraErrorMessage = "Il primo record del file non è l'header."
'                    GoTo ErrHandler
'                End If

                If Not mobjInFileHeader.ParseFixedLenFields(strLine, 1, mobjInFileHeader.Count) Then
                    ErrCode = InvalidHeader
                    strExtraErrorMessage = "Errore durante il parser dell'header."
                    GoTo ErrHandler
                End If

                mstrIDEnte = mobjInFileHeader.Item(7).value
                strEnte = AEE.DecodeFirmCode(ID_SOC, mstrIDEnte)
                If strEnte <> vbNullString Then
                    lblField(1) = strEnte
                    blnIDOk = True
                Else
                    ErrCode = NoEnteEmittente
                    GoTo ErrHandler
                End If

                lblField(0) = mobjInFileHeader.Item(2).value
                lblField(2) = mobjInFileHeader.Item(7).value

                ' TODO: Verificare ultimo progressivo
                If IsNumeric(mstrUltimoProgressivo) Then
                    If lblField(0) <> CStr(CLng(mstrUltimoProgressivo) + 1) Then
                        MsgBox "Il numero del file atteso è:  " & (CLng(mstrUltimoProgressivo) + 1), vbOKOnly + vbExclamation, "Attenzione"
                    End If
                End If
                blnFoundHeader = True

            ' Detail
            Case "21"

                If Not mobjInFileDetail.ParseFixedLenFields(strLine, 1, mobjInFileDetail.Count) Then
                    ErrCode = InvalidDetail
                    strExtraErrorMessage = "Errore nel parser del dettaglio."
                    GoTo ErrHandler
                End If

                With XA
                    .ReDim 1, lngRows, 0, mobjInFileDetail.Count
                    .value(.UpperBound(1), 0) = 0
                    For i = 0 To mobjInFileDetail.Count - 1
                        .value(.UpperBound(1), i + 1) = mobjInFileDetail.Item(i + 1).value
                    Next i
                End With
                
                With XALog
                    .ReDim 1, lngRows, 0, 7
                    .value(.UpperBound(1), fldLog.fldLog_ViaggioCounter) = 0
                    .value(.UpperBound(1), fldLog.fldLog_DataViaggio) = mobjInFileDetail.Item(fldVVTImport_DataViaggio).value
                    .value(.UpperBound(1), fldLog.fldLog_Badge) = mobjInFileDetail.Item(fldVVTImport_Badge).value
                    .value(.UpperBound(1), fldLog.fldLog_SequenzaViaggio) = mobjInFileDetail.Item(fldVVTImport_SequenzaViaggio).value
                    .value(.UpperBound(1), fldLog.fldLog_OraConsegna) = mobjInFileDetail.Item(fldVVTImport_OraConsegnaPrevista).value
                    .value(.UpperBound(1), fldLog.fldLog_Destinazione) = mobjInFileDetail.Item(fldVVTImport_Destinazione).value
                    .value(.UpperBound(1), fldLog.fldLog_Prodotto) = mobjInFileDetail.Item(fldVVTImport_Prodotto).value
                    .value(.UpperBound(1), fldLog.fldLog_Esito) = vbNullString
                End With
                lngRows = lngRows + 1
            
            ' Footer
'            Case "90"
'                If Not blnFoundHeader Then
'                    ErrCode = InvalidHeader
'                    strExtraErrorMessage = "Nel file è presente il footer ma non l'header."
'                    GoTo ErrHandler
'                End If
'
'                If Not mobjInFileFooter.ParseFixedLenFields(strLine, 1, mobjInFileFooter.Count) Then
'                    ErrCode = InvalidFooter
'                    strExtraErrorMessage = "Errore durante il parser del footer."
'                    GoTo ErrHandler
'                End If
'
'                blnFoundFooter = True
'                If mobjInFileHeader.Item(2).value <> mobjInFileFooter.Item(2).value Or _
'                    mobjInFileHeader.Item(4).value <> mobjInFileFooter.Item(4).value Or _
'                    mobjInFileHeader.Item(7).value <> mobjInFileFooter.Item(7).value Then
'                    ErrCode = InvalidFooter
'                    strExtraErrorMessage = "Nel footer sono presenti dei dati non coerenti rispetto all'header."
'                    GoTo ErrHandler
'                End If
'
''                If CLng(mobjInFileFooter.Item(9).value) <> lngRows Then
''                    ErrCode = InvalidFooter
''                    strExtraErrorMessage = "Il numero di record segnalato nel footer non corrisponde al numero di record presenti nel file."
''                    GoTo ErrHandler
'                End If
                
        End Select
    Loop

'    If Not blnFoundFooter Then
'        ErrCode = InvalidFooter
'        GoTo ErrHandler
'    End If

    cmdImport.Enabled = True

    XA.QuickSort 1, XA.UpperBound(1), _
                    fldVVTImport_DataViaggio, XORDER_ASCEND, XTYPE_STRING, _
                    fldVVTImport_Badge, XORDER_ASCEND, XTYPE_STRING, _
                    fldVVTImport_SequenzaViaggio, XORDER_ASCEND, XTYPE_STRING, _
                    fldVVTImport_OraConsegnaPrevista, XORDER_ASCEND, XTYPE_STRING, _
                    fldVVTImport_Prodotto, XORDER_ASCEND, XTYPE_STRING

    XALog.QuickSort 1, XALog.UpperBound(1), _
                    fldLog.fldLog_DataViaggio, XORDER_ASCEND, XTYPE_STRING, _
                    fldLog.fldLog_Badge, XORDER_ASCEND, XTYPE_STRING, _
                    fldLog.fldLog_SequenzaViaggio, XORDER_ASCEND, XTYPE_STRING, _
                    fldLog.fldLog_OraConsegna, XORDER_ASCEND, XTYPE_STRING, _
                    fldLog.fldLog_Prodotto, XORDER_ASCEND, XTYPE_STRING

    strPrecIdViaggio = ""
    lngCountViaggi = 0
    For lngCounter = XA.LowerBound(1) To XA.UpperBound(1)
    
        strIDViaggio = GetKeyViaggio(lngCounter)
        If strPrecIdViaggio <> strIDViaggio Then
            strPrecIdViaggio = strIDViaggio
            lngCountViaggi = lngCountViaggi + 1
        End If
        XA.value(lngCounter, 0) = lngCountViaggi
        XALog.value(lngCounter, 0) = lngCountViaggi
    Next lngCounter

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
    lblLog(5) = lngCountViaggi
    lblLog(6) = 0
    
    GoTo EndProc

ErrHandler:
    Dim strErrorMessage As String
    Dim strInfoMessage As String
    
    Select Case ErrCode
        Case NoEnteEmittente
            strErrorMessage = "Codice Ente Emittente non trovato."
            strInfoMessage = "Codice Ente: " & mobjInFileHeader.Item(7).value
        
        Case InvalidHeader
            strErrorMessage = "Il file ha un header con formato non valido."
            strInfoMessage = "Nome del file: " & strFileName
        
        Case InvalidFooter
            strErrorMessage = "Il file ha un footer con formato non valido."
            strInfoMessage = "Nome del file: " & strFileName
        
        Case InvalidDetail
            strErrorMessage = "Il file ha un dettaglio con formato non valido."
            strInfoMessage = "Nome del file: " & strFileName & " - Linea: " & lngRows

        Case FileNotFound
            strErrorMessage = "Il file non è stato trovato."
            strInfoMessage = "Nome del file: " & strFileName
        
        Case NoError
            Select Case Err.Number
                Case Else
                    strErrorMessage = "Errore sconosciuto"
                    strInfoMessage = ""
            End Select
    End Select
    If strExtraErrorMessage <> vbNullString Then
        strInfoMessage = strInfoMessage & " - " & strExtraErrorMessage
    End If
    Set C = ErrH.SetError(ErrCode, Me.Name, _
            "OpenTextFile", 0, strErrorMessage, 10, strInfoMessage)
    
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError

EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault

End Sub

Private Function GetKeyViaggio(lngIndex As Long)
    GetKeyViaggio = XA.value(lngIndex, fldVVTImport_DataViaggio) & "." & _
        XA.value(lngIndex, fldVVTImport_Badge) & "." & _
        XA.value(lngIndex, fldVVTImport_SequenzaViaggio)
End Function

Private Function GetKeyConsegna(lngIndex As Long)
    GetKeyConsegna = GetKeyViaggio(lngIndex) & "." & _
        Left(XA.value(lngIndex, fldVVTImport_ConsegnaRifEsterno1), 10) & _
        XA.value(lngIndex, fldVVTImport_OraConsegnaPrevista)
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    
    ConfigureObjects Me
    
    Set ErrH = New CErrorHandler
    
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
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AEE = New CEnteEmittente
    With AEE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ASC = New CSocieta
    With ASC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
'    Set APD = New CPosDoganale
'    With APD.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'        Set .MultiSoc.DS.ErrH = ErrH
'    End With
    
    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
'    Set AZO = New CZona
'    With AZO.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AMD = New CMagazzinoDeposito
    With AMD.DS
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
    
    mskFileName = gmRegistry.GetSetting("Options", "TripImportFile1", , True)
    mstrUltimoProgressivo = gmRegistry.GetSetting("Options", "PrevTripImportFileID", "0", True)

    Set mobjInFile = New CTextFile
    Set mobjInFileHeader = New CTextFile
    Set mobjInFileFooter = New CTextFile
    Set mobjInFileDetail = New CTextFile

    With mobjInFile
    End With

    With mobjInFileHeader
        .Add "Tipo Record", 2, "1"
        .Add "Progressivo", 4, "2"
        .Add "Data Invio", 8, "3"
        .Add "Tipo File", 9, "4"
        .Add "Unknow 1", 6, "5"
        .Add "Unknow 2", 6, "6"
        .Add "Ente Emittente", 4, "7"
        .Add "Unknow 3", 2, "8"
    End With

    With mobjInFileFooter
        .Add "Tipo Record", 2, "1"
        .Add "Progressivo", 4, "2"
        .Add "Data Invio", 8, "3"
        .Add "Tipo File", 9, "4"
        .Add "Unknow 1", 6, "5"
        .Add "Unknow 2", 6, "6"
        .Add "Ente Emittente", 4, "7"
        .Add "Unknow 3", 2, "8"
        .Add "Record Count", 4, "9"
    End With
    
    With mobjInFileDetail
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
        .Add "Quantità trasportata", 9, "17"
        .Add "Data Viaggio", 8, "18"
        .Add "Badge", 11, "19"
        .Add "N Sequenza", 2, "20"
        .Add "Cart Identità Autista", 20, "21"
        .Add "Ora consegna prevista", 5, "22"
    End With
    
    Select Case mvarFunction
        Case fcIMPVVT1
            MaskCodSoc.Text = ID_SOC_GUARNER
            MaskCodSoc.Enabled = False
            
    End Select

    lblsocieta.Caption = DescrizioneSocieta(MaskCodSoc)
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "TripImportFile1", mskFileName, True
    gmRegistry.SaveSetting "Options", "PrevTripImportFileID", mstrUltimoProgressivo, True
    
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set ADE = Nothing
    Set ASC = Nothing
    Set OCR = Nothing
    Set OCT = Nothing
    Set APR = Nothing
    Set AEE = Nothing
    Set ACL = Nothing
    Set AUT = Nothing
    Set AMD = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    Set VVP = Nothing
    
    
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set mobjInFileHeader = Nothing
    Set mobjInFileFooter = Nothing
    Set mobjInFileDetail = Nothing
    Set frmVVTImport1 = Nothing
End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If XA.value(bookmark, 0) Mod 2 = 0 Then
        RowStyle.BackColor = &HC0FFC0
    Else
        RowStyle.BackColor = &HC0FFFF
    End If
End Sub

Private Sub tdbLog_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If XALog.value(bookmark, 0) Mod 2 = 0 Then
        RowStyle.BackColor = &HC0FFC0
    Else
        RowStyle.BackColor = &HC0FFFF
    End If
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

