VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTImport3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CEPSA - RICEZIONE TICKETS"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   12270
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12255
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   6360
         TabIndex        =   11
         Top             =   120
         Width           =   5775
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
            Caption         =   "Tickets Esclusi"
            Height          =   255
            Index           =   7
            Left            =   4200
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   4200
            TabIndex        =   30
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Tickets da importare"
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   26
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   25
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Tickets importati "
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   19
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ora Invio"
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   18
            Top             =   840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Invio"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   16
            Top             =   240
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
            Width           =   495
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
         Top             =   1560
         Width           =   12015
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmOCTImport3.frx":0000
            TabIndex        =   33
            Tag             =   "SELECT"
            Top             =   240
            Width           =   11775
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   12015
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmOCTImport3.frx":68A4
            TabIndex        =   32
            Tag             =   "SELECT"
            Top             =   240
            Width           =   11775
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
         Picture         =   "frmOCTImport3.frx":10F56
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
         Picture         =   "frmOCTImport3.frx":11488
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
         Picture         =   "frmOCTImport3.frx":1158A
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
         Picture         =   "frmOCTImport3.frx":11ABC
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
         Picture         =   "frmOCTImport3.frx":11BBE
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
         Picture         =   "frmOCTImport3.frx":11CC0
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
         Left            =   5640
         Top             =   840
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
               Picture         =   "frmOCTImport3.frx":11DC2
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport3.frx":11ED4
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport3.frx":11FE6
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
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   21
      Top             =   5760
      Width           =   12255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   10560
         Picture         =   "frmOCTImport3.frx":120F8
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
         Left            =   10440
         Picture         =   "frmOCTImport3.frx":121FA
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
      Width           =   12270
      _ExtentX        =   21643
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
            Object.Width           =   16193
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:47"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "01/06/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTImport3"
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
Private Const COL_RESULT_LOG = 4

Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
'Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineriga
Private APR As CProdotto
'Private APD As CPosDoganale
'Private ADE As CDeposito
'Private AEE As CEnteEmittente
'Private ATV As CTipoVendita
'Private AZO As CZona

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

'Private Function CheckDestinazione001(ByRef strIDDestinazione As String, _
'                                      ByVal strCodDest As String, _
'                                      ByVal strIDClienteDefault As String, _
'                                      ByVal strIDDeposito As String, _
'                                      ByRef strNewLocalita As String) As Long
'    Dim intRowError As Errors
'    Dim strLocalita As String
'    Dim strNewZona As String
'    Dim strNewIndirizzo As String
'    Dim blnConcurErr As Boolean
'    Dim lngKmRaggio As Long
'
'    intRowError = NoError
'    ' --- Cerco la destinazione
'    strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC, strCodDest, True)
'
'    ' --- La destinazione è già stata inserita ?
'    If strIDDestinazione = vbNullString Then
'        ' --- No : devo inserirla ora
'        With ADZ.DS.Fields
'            RecordToADZ ADZ, mobjInFile, strIDClienteDefault, strIDDeposito
'            If Not ADZ.Insert Then
'                intRowError = DestInsertError
'                GoTo ExitFunction
'            End If
'            If Not ADZ.DS.MultiSoc.Insert Then
'                intRowError = DestSocInsertError
'                GoTo ExitFunction
'            End If
'        End With
'        strIDDestinazione = strCodDest
'    Else
'        ' --- Controllo che la località non sia cambiata
'        ADZ.DS.Fields("0").FindValue = strIDDestinazione
'        If Not ADZ.Find Then
'            intRowError = NoDestinazione
'            GoTo ExitFunction
'        End If
'        strLocalita = UCase(Trim$(ADZ.DS.Fields("5").value))
'        strNewLocalita = ADZ.DS.DA.MakeSQL(mobjInFile("7").value, 0) 'dtAlfa)
'        strNewLocalita = Mid$(strNewLocalita, 2, Len(strNewLocalita) - 2)
'        If strLocalita <> strNewLocalita Then
'            ' --- Se la località è cambiata aggiorno la destinazione
'            strNewZona = mobjInFile("2").value
'            strNewIndirizzo = mobjInFile("6").value
'            If Len(mobjInFile("15").value) > 0 Then
'                lngKmRaggio = mobjInFile("15").value
'            End If
'            With ADZ.DS
'                .Fields("4").NewValue = strNewIndirizzo
'                .Fields("4").Updated = True
'                .Fields("5").NewValue = strNewLocalita
'                .Fields("5").Updated = True
'                .Fields("15").NewValue = strIDDeposito
'                .Fields("15").Updated = True
'                .Fields("16").NewValue = strNewZona
'                .Fields("16").Updated = True
'                .Fields("29").NewValue = lngKmRaggio
'                .Fields("29").Updated = True
'                ADZ.BeginSave blnConcurErr
'                .DA.EndTrans ETCOMMIT
'                .EndSave
'            End With
'            ' ---
'            intRowError = DestHasChanged
'            GoTo ExitFunction
'        Else
'            If Len(ADZ.DS.Fields("29").value) > 0 Then
'                lngKmRaggio = ADZ.DS.Fields("29").value
'            Else
'                lngKmRaggio = 0
'            End If
'            If lngKmRaggio <> mobjInFile("15").value Then
'                With ADZ.DS
'                    ' Km Raggio
'                    .Fields("29").NewValue = mobjInFile("15").value
'                    .Fields("29").Updated = True
'                    ' Deposito
'                    .Fields("15").NewValue = strIDDeposito
'                    .Fields("15").Updated = True
'                    ADZ.BeginSave blnConcurErr
'                    .DA.EndTrans ETCOMMIT
'                    .EndSave
'                End With
'            End If
'        End If
'        ' ---
'    End If
'    ' ---
'ExitFunction:
'    CheckDestinazione001 = intRowError
'End Function

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
                .Filter = "(*.dat)|*.dat|Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
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
        "       AND OCRRIFE2 = ? AND OCRANNUL = 0 AND ((OCTRAUT5 IS NULL) OR (OCTRAUT5 = '') OR (OCTRAUT5 <> 'S'))  "
    Const SEL_ORDINE_2 = " AND CONVERT(INT,OCTRAUT4) = "
    
    Const UPD_ORDINE = _
        "UPDATE ORDIOCT0 " & _
        "   SET OCTNOTA3 = ? " & _
        " WHERE OCTIDOCT = ? "
'    Const UPD_ORDINE = _
'        "UPDATE ORDIOCT0 " & _
'        "   SET OCTNOTA3 = ? " & _
'        " WHERE OCTRIFE1 = ? AND OCTIDASC = '038'"
'    Const UPD_ORDINE_2 = _
'        " AND OCTRAUT4 = "

    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
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
    If Len(strIDOCT) = 0 Then
        GoTo ErrHandler
    End If
    strSQL = UPD_ORDINE
    If Not OCT.DS.DA.ExecuteCommand(strSQL, strTicket, strIDOCT) Then
        GoTo ErrHandler
    End If
    UpdOrdine = True

Exit Function
ErrHandler:
    UpdOrdine = False
End Function

Private Sub ImportRecords001()
    On Error GoTo ErrHandler
    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDOrder As String
    Dim strNewLocalita As String
    Dim strErrorMessage As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim blnLocHasChanged As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    Dim strRifEsterno1 As String
    Dim strRifEsterno1Prec As String
    Dim strProgrDettaglio As String
    Dim strTicket As String
    Dim strTicketPrec As String
    Dim strPartizione As String
    Dim strPartizionePrec As String
    Dim strAdditivo As String
    Dim strIDProdotto As String
    Dim StrIDAPR As String
    Dim varQuantita As Variant
    ReDim ListNewLoc(0 To 0) As String

    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    
    
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
                
        intRowError = NoError
        strRifEsterno1 = XA.value(Row, 4)
        strProgrDettaglio = XA.value(Row, 5)
        strPartizione = XA.value(Row, 6)
        If strRifEsterno1 <> strRifEsterno1Prec Then
            strRifEsterno1Prec = strRifEsterno1
            strTicket = vbNullString
        End If
        If strPartizione <> strPartizionePrec Then
            strPartizionePrec = strPartizione
            strTicket = vbNullString
        End If
        strAdditivo = XA.value(Row, 15)
        strIDProdotto = XA.value(Row, 8)
        StrIDAPR = APR.DecodeFirmCode("901", strIDProdotto + strAdditivo, True)
        If (strRifEsterno1 = strRifEsterno1Prec) And (strPartizione = strPartizionePrec) Then
            If (Left(XA.value(Row, 14), 7) <> Right(strTicket, 7)) Then
                If Len(strTicket) = 0 Then
                    strTicket = Left(XA.value(Row, 14), 7)
                    If strAdditivo = "521" Then
                        strTicket = "O " & strTicket
                    ElseIf strIDProdotto = "0606420" Then
                        strTicket = "+ " & strTicket
                    End If
                Else
                    If strAdditivo = "521" Then
                        strTicket = strTicket & " / O " & Left(XA.value(Row, 14), 7)
                    ElseIf strIDProdotto = "0606420" Then
                        strTicket = "+ " & strTicket
                    Else
                        strTicket = strTicket & " / " & Left(XA.value(Row, 14), 7)
                    End If
                End If
            End If
        Else
            strTicket = Left(XA.value(Row, 14), 7)
            If strAdditivo = "521" Then
                strTicket = "O " & strTicket
            ElseIf strIDProdotto = "0606420" Then
                strTicket = "+ " & strTicket
            End If
        End If
        
        strErrorMessage = vbNullString
        If OCT.CheckSocCode1(mstrIDSoc, strRifEsterno1, strPartizione, strErrorMessage) Then
            If Not UpdOrdine(strRifEsterno1, strTicket, strProgrDettaglio, strIDOrder, strPartizione) Then
                intRowError = OrderUpdateError
                GoTo LogResult
            End If
            If Not OCR.UpdateNrTicket(strIDOrder, StrIDAPR, XA.value(Row, 9), XA.value(Row, 14), strPartizione, strProgrDettaglio) Then
                intRowError = OrderUpdateError
                GoTo LogResult
            End If
        Else
            intRowError = BadIDOrderSoc
            GoTo LogResult
        End If
        
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError = NoError) Then
            lngCountOK = lngCountOK + 1
        Else
            lngCountFail = lngCountFail + 1
        End If
        lngCountToGo = lngCountToGo - 1
        
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(6) = lngCountExcluded
        
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK "
            Case DestHasChanged
                blnLocHasChanged = True
                ListNewLoc(UBound(ListNewLoc)) = "Dest. modificata: " & strIDDestinazione & " " & _
                                                 strNewLocalita
                strErrorMessage = "OK - Num. Ordine : " & strRifEsterno1 & _
                                  " - " & ListNewLoc(UBound(ListNewLoc))
                ReDim Preserve ListNewLoc(UBound(ListNewLoc) + 1) As String
            Case OrderAlreadyExisting
                strErrorMessage = "Ordine già inserito"
            Case NoTipoVendita
                strErrorMessage = "Codice tipo vendita non trovato"
            Case NoZona
                strErrorMessage = "Codice zona non trovato"
            Case NoDeposito
                strErrorMessage = "Codice deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Codice ente emittente non trovato"
            Case NoProdotto
                strErrorMessage = "Codice prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice destinazione mancante"
            Case BadIDOrderSoc
                If Len(strErrorMessage) = 0 Then
                    strErrorMessage = "Ordine non trovato"
                End If
            Case DestInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case OrderInsertError
                strErrorMessage = "Errore durante l'inserimento del ordine "
            Case RowOrderInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di ordine"
            Case ZonaEsclusaImport
                strErrorMessage = "Zona esclusa dall'importazione ordine"
            Case OrderUpdateError
                strErrorMessage = "Errore durante l'aggiornamento dell' ordine"
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
    Set C = ErrH.SetError(ErrCode, "frmOCTImport3", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    Me.Enabled = True
    Me.MousePointer = vbDefault
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
    '***05/09/2003 CZ
    Dim StrTipoRecord As String
    Dim strFiller2 As String
    Dim strOperatore As String
    Dim strFiller4 As String
    Dim strRifEsterno1 As String
    Dim strProgrDettaglio As String
    Dim strPartizioneOrdine As String
    Dim strCodDest As String
    Dim CodProdotto As String
    Dim VarQta As Variant
    Dim strUM As String
    Dim strCodveicolo As String
    Dim varNumViaggio As Variant
    Dim varFiller13 As Variant
    Dim Nota3 As String
    Dim strDummyData As String
    
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
        If Left(CStr(varBuffer), 2) = "31" Then
            If Not mobjInFile.ParseFixedLenFields(CStr(varBuffer), 1, 20) Then
                GoTo ErrHandler
            End If
            StrTipoRecord = mobjInFile.Item("1").value
            strFiller2 = mobjInFile.Item("2").value
            strOperatore = mobjInFile.Item("3").value
            strFiller4 = mobjInFile.Item("4").value
            strRifEsterno1 = Left(mobjInFile.Item("5").value, 10)
            strProgrDettaglio = Right(mobjInFile.Item("5").value, 3)
            strPartizioneOrdine = mobjInFile.Item("6").value
            strCodDest = mobjInFile.Item("7").value
            CodProdotto = mobjInFile.Item("8").value
            VarQta = mobjInFile.Item("9").value
            strUM = mobjInFile.Item("10").value
            strCodveicolo = mobjInFile.Item("11").value
            varNumViaggio = mobjInFile.Item("12").value
            strDummyData = mobjInFile.Item("13").value
            varFiller13 = Right(strDummyData, 2) + "/" + Mid(strDummyData, 5, 2) + "/" + Left(strDummyData, 4)
            Nota3 = mobjInFile.Item("14").value
            
            With XA
                .ReDim 1, lngRows, 0, 16
                .value(lngRows, 0) = StrTipoRecord
                .value(lngRows, 1) = strFiller2
                .value(lngRows, 2) = strOperatore
                .value(lngRows, 3) = strFiller4
                .value(lngRows, 4) = strRifEsterno1
                .value(lngRows, 5) = strProgrDettaglio
                .value(lngRows, 6) = strPartizioneOrdine
                .value(lngRows, 7) = strCodDest
                .value(lngRows, 8) = CodProdotto
                .value(lngRows, 9) = VarQta
                .value(lngRows, 10) = strUM
                .value(lngRows, 11) = strCodveicolo
                .value(lngRows, 12) = varNumViaggio
                .value(lngRows, 13) = varFiller13
                .value(lngRows, 14) = Nota3
                .value(lngRows, 15) = mobjInFile.Item("15").value
                .value(lngRows, 16) = mobjInFile.Item("16").value
            End With
            
            With XALog
                .ReDim 1, lngRows, 0, 4 '2
                .value(.UpperBound(1), 0) = strRifEsterno1
                .value(.UpperBound(1), 1) = strPartizioneOrdine
                .value(.UpperBound(1), 2) = Nota3
                .value(.UpperBound(1), 3) = strProgrDettaglio
            End With
            lngRows = lngRows + 1
        End If
    Loop
    
    XA.QuickSort 1, XA.UpperBound(1), _
                 4, XORDER_ASCEND, XTYPE_DEFAULT, _
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
            Set C = ErrH.SetError(FileNotFound, "frmOCTImport3", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmOCTImport3", _
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
            mskFileName = gmRegistry.GetSetting("Options", "ImportTicketsFile2", , True)
            mstrIDSoc = ID_SOC_GUARNER
            Caption = "GUARNER - RICEZIONE TICKETS"
        
        Case Else
            mskFileName = gmRegistry.GetSetting("Options", "ImportTicketsFile", , True)
            'mstrUltimoProgressivo = gmRegistry.GetSetting("Options", "PrevOrderImportFileID", "0", True)
            mstrIDSoc = ID_SOC_CEPSA
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
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineriga
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
    With mobjInFile
        .Add "Tipo Record", 2, "1"
        .Add "Filler2", 4, "2"
        .Add "Operatore", 6, "3"
        .Add "Filler4", 6, "4"
        .Add "Rif. Esterno1", 13, "5"
        .Add "Partizione", 2, "6"
        .Add "Codice Destinazione", 8, "7"
        .Add "Codice del prodotto", 7, "8"
        .Add "Q.ta", 9, "9"
        .Add "Unità di misura", 1, "10"
        .Add "Codice veicolo", 11, "11"
        .Add "Numero viaggio", 2, "12"
        .Add "Filler13", 8, "13"
        .Add "Ticket", 15, "14"
        .Add "Codice Additivo", 3, "15"
        .Add "Nome commerciale se additivato", 30, "16"
        .Add "Matricula Tractora", 18, "17"
        .Add "Matricula Cisterna", 18, "18"
        .Add "Matricula 3", 18, "19"
        .Add "Relleno", 50, "20"
    End With
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case mvarFunction
        Case fcRICTICK
            gmRegistry.SaveSetting "Options", "ImportTicketsFile2", mskFileName, True
        
        Case Else
            gmRegistry.SaveSetting "Options", "ImportTicketsFile", mskFileName, True
    End Select
    
    'gmRegistry.SaveSetting "Options", "OrderImportFileDetail", mskFileNameDettaglio, True
    'gmRegistry.SaveSetting "Options", "PrevOrderImportFileID", mstrUltimoProgressivo, True
    
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set OCT = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set frmOCTImport3 = Nothing
End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub


