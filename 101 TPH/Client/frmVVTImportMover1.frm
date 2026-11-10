VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVVTImportMover1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE VIAGGI DA MOVER"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   5160
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7680
         Picture         =   "frmVVTImportMover1.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmVVTImportMover1.frx":0102
         TabIndex        =   14
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6060
      Width           =   9390
      _ExtentX        =   16563
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
            Object.Width           =   11113
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15.39"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "05/01/00"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   2175
         Left            =   7560
         TabIndex        =   12
         Top             =   600
         Width           =   1695
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Viaggi da importare:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Viaggi importati:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   7335
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1335
            Left            =   120
            OleObjectBlob   =   "frmVVTImportMover1.frx":0204
            TabIndex        =   5
            Top             =   240
            Width           =   7095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   9135
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmVVTImportMover1.frx":3204
            TabIndex        =   6
            Top             =   240
            Width           =   8895
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
         Picture         =   "frmVVTImportMover1.frx":9182
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmVVTImportMover1.frx":96B4
         TabIndex        =   2
         Tag             =   "IMPORT"
         Top             =   600
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
         Picture         =   "frmVVTImportMover1.frx":97B6
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmVVTImportMover1.frx":9CE8
         TabIndex        =   1
         Tag             =   "OPEN"
         Top             =   600
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
         Picture         =   "frmVVTImportMover1.frx":9DEA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTImportMover1.frx":9EEC
         TabIndex        =   0
         Tag             =   "BROWSE"
         Top             =   600
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2760
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Tag             =   "4"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
               Picture         =   "frmVVTImportMover1.frx":9FEE
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTImportMover1.frx":A100
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTImportMover1.frx":A212
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emit."
         Height          =   255
         Index           =   16
         Left            =   4680
         TabIndex        =   26
         Top             =   240
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
         Left            =   6120
         TabIndex        =   25
         Tag             =   "4"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   9000
         Tag             =   "4"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmVVTImportMover1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Importazione Viaggi Tamoil
'
'
Option Explicit

Private Const MAX_COLS = 3
Private XA As XArrayDB
Private XALog As XArrayDB

Public DataClass As CViaggioSelezione

Private AXX As CParametro
Private ADE As CDeposito
Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineRiga
Private ACL As CCliente
Private AUT As CUnitaDiTrasporto
Private AAM As CAutomezzo
Private AMD As CMagazzinoDeposito
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private APR As CProdotto
Private mobjInFile As CTextFile
Private mstrIDEnte As String
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private strDataViaggio As String
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file dei viaggi... "
                .InitDir = App.Path
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
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
Private Sub ImportRecords()
    On Error GoTo ErrHandler

    Dim strFileName As String
    Dim strBuffer As String
    Dim strIDImballoDefault As String
    Dim strIDSocieta As String
    Dim strIDPosDefault As String
    Dim strSocDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDTrip As String
    Dim strIDTripMover As String
    Dim strIDProdotti(1 To 10) As String
    Dim strIDProdotto As String
    Dim strIDRigaConsegna  As String
    Dim strIDConsegna As String
    Dim strIDScarico As String
    Dim strIDSequenza As String
    Dim strIDMotrice As String
    Dim strIDRimorchio As String
    Dim strIDVettore As String
    Dim strIDVettoreT As String
    Dim strIDAutista1 As String
    Dim strIDAutista2 As String
    Dim strBadgeAutomezzo As String
    Dim strIDDeposito As String
    Dim strIDMagazzino As String
    Dim strIDZona As String
    Dim strIDCliente As String
    Dim strIDUT As String
    Dim strIDOrdine As String
    Dim strIDOrdineRiga As String
    Dim strInizioViaggio As String
    Dim strIDImballo As String
    Dim strErrorMessage As String
    Dim lngCountTrips As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim i As Integer
    Dim intIDOrdineRiga As Integer
    Dim blnNewTrip As Boolean
    Dim blnTripOK As Boolean
    Dim varDatiCliDes(0 To 13) As Variant
    Dim varQuantita(1 To 4) As Long
    Dim strPosdoganale(1 To 4) As String
    Dim strIDSocietaDef As String   ' Società di Default
    Dim intRowError As Errors
    Dim strNewRow As String
    Dim BlDummy As Boolean
    Dim varDummy As Variant
    Dim ErrCode As Errors
    Dim C As CError

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
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        strIDImballoDefault = Trim$(.Fields("1").value)        ' Codice Imballo di Default
        strIDPosDefault = Trim$(.Fields("2").value)            ' Posizione Imballo di Default
        strSocDefault = Trim$(.Fields("4").value)
'        strIDEnteEmittente = Trim$(.Fields("16").value)        ' Ente Emittente Default
        For i = 1 To 10
            strIDProdotti(i) = Trim$(.Fields(CStr(i + 5)).value)
        Next i
    End With
'       (strIDEnteEmittente = vbNullString) Or _

    If (strIDImballoDefault = vbNullString) Or (strIDPosDefault = vbNullString) Or _
       (strSocDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---

    lngNumRows = XALog.UpperBound(1)
    lngStartRow = XALog.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    strFileName = mskFileName
    strIDEnteEmittente = MaskEdBox1
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If

    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        blnNewTrip = False
        strBuffer = mobjInFile.GetLine(False)
        If Left$(strBuffer, 2) = "T0" Then
            '--- Testata viaggio
            mobjInFile.ParseFixedLenFields strBuffer, 1, 15
            blnNewTrip = True
            VVT.DS.DA.BeginTransaction
            blnTripOK = False
            strIDUT = mobjInFile("3").value                ' Unità di Trasporto
            strIDMotrice = mobjInFile("4").value           ' Codice Motrice
            strIDRimorchio = mobjInFile("5").value         ' Codice Rimorchio
            strIDDeposito = ADE.DecodeFirmCode(strSocDefault, _
                            mobjInFile("11").value, True)   ' Codice Deposito da Multiso
            strIDMagazzino = mobjInFile("13").value        ' Codice Magazzino
            
            If strIDMagazzino = vbNullString Then          ' Controllo che il Codice Magazzino
                With AMD.DS                                ' che arriva da Mover
                    .Fields("0").FindValue = strIDDeposito
                    If Not AMD.Find Then
                        intRowError = NoMagazzino
                        GoTo LogResult
                    End If
                    If .HowMany = 0 Then
                        intRowError = NoMagazzino
                        GoTo LogResult
                    End If
                    strIDMagazzino = .Fields("1").value
                End With
            End If
            
            strInizioViaggio = Format$(Left$(mobjInFile("9").value, 2), "00") & ":" & Format$(Right$(mobjInFile("9").value, 2), "00")
            
            If strIDDeposito = vbNullString Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
            If strIDUT = vbNullString Then
                intRowError = NoUT
                GoTo LogResult
            End If
            
            With AUT.DS
                .Fields("0").FindValue = strIDUT
                If Not AUT.Find Then
                    intRowError = NoUT
                    GoTo LogResult
                End If
                If .HowMany = 0 Then
                    intRowError = NoUT
                    GoTo LogResult
                End If
                If strIDMotrice = vbNullString Then
                    strIDMotrice = .Fields("2").value
                End If
                If strIDRimorchio = vbNullString Then
                    strIDRimorchio = .Fields("3").value
                End If
                strIDAutista1 = .Fields("4").value
                strIDAutista2 = .Fields("12").value
                strIDVettore = .Fields("6").value
                strIDVettoreT = .Fields("7").value
                
            End With
            
             With AAM.DS
                .Fields("0").FindValue = strIDMotrice
                If Not AAM.Find Then
                    intRowError = NoMotrice
                    GoTo LogResult
                End If
                If .HowMany = 0 Then
                    intRowError = NoMotrice
                    GoTo LogResult
                End If
                If Len(strIDRimorchio) > 0 Then
                    .Fields("0").FindValue = strIDRimorchio
                    If Not AAM.Find Then
                        intRowError = NoRimorchio
                        GoTo LogResult
                    End If
                    If .HowMany = 0 Then
                        intRowError = NoRimorchio
                        GoTo LogResult
                    End If
                End If
            End With
            
             
            strIDSocieta = mobjInFile("36").value
            If strIDSocieta = vbNullString Then
                strIDSocieta = strSocDefault
            End If
            ' ---
            ' --- Inserisco la testata del viaggio
            RecordToVVT VVT, mobjInFile, strIDMotrice, strIDRimorchio, strIDDeposito, strIDMagazzino, strIDSocieta, strIDEnteEmittente, strIDVettore, strIDAutista1, strIDAutista2, strIDVettoreT
            If Not VVT.Insert Then
                intRowError = TripInsertError
                GoTo LogResult
            End If
            strIDTrip = VVT.DS.Fields("0").NewValue
            strIDRigaConsegna = "010"
            strIDSequenza = "010"
            blnTripOK = True
            ' ---
        Else
            If blnTripOK Then
                mobjInFile.ParseFixedLenFields strBuffer, 16, 35
                If strIDTrip <> vbNullString Then
                    strIDOrdine = mobjInFile.Item("24").value     'Codice Ordine
'                    If strIDOrdine = vbNullString Then
'                        intRowError = NoOrderID
'                        GoTo LogResult
'                    End If
                    strIDDestinazione = mobjInFile.Item("19").value
                     If IsNumeric(strIDDestinazione) Then
                        strIDDestinazione = Format$(strIDDestinazione, "0000")
                    End If
                    With ADZ.DS  ' Ricavo i Dati sulla destinazione
                       .Fields("0").FindValue = strIDDestinazione
                       If Not ADZ.Find Then
                           intRowError = NoDestinazione
                           GoTo LogResult
                       End If
                       If .HowMany = 0 Then
                           intRowError = NoDestinazione
                           GoTo LogResult
                       End If
                       varDatiCliDes(7) = .Fields("1").value
                       varDatiCliDes(8) = .Fields("4").value
                       varDatiCliDes(9) = .Fields("5").value
                       varDatiCliDes(10) = .Fields("6").value
                       varDatiCliDes(11) = .Fields("7").value
                       varDatiCliDes(12) = .Fields("9").value
                       varDatiCliDes(13) = .Fields("10").value
                       strIDZona = .Fields("16").value
                       strIDCliente = .Fields("8").value
                       strIDTipoVendita = .Fields("28").value
                    End With
                    With ACL.DS   ' --- Leggo i dati anagrafici del cliente
                        .Fields("0").FindValue = strIDCliente
                        If Not ACL.Find Then
                            intRowError = NoCliente
                            GoTo LogResult
                        End If
                        If .HowMany = 0 Then
                            intRowError = NoCliente
                            GoTo LogResult
                        End If
                        varDatiCliDes(0) = .Fields("1").value
                        varDatiCliDes(1) = .Fields("4").value
                        varDatiCliDes(2) = .Fields("5").value
                        varDatiCliDes(3) = .Fields("6").value
                        varDatiCliDes(4) = .Fields("7").value
                        varDatiCliDes(5) = .Fields("10").value
                        varDatiCliDes(6) = .Fields("11").value
                    End With
                    
                    OCT.DS.Fields("0").FindValue = strIDOrdine    ' Check Esistenza Ordine
                    If Not OCT.Find Or OCT.DS.HowMany = 0 Then
                        RecordToOCT OCT, strIDDeposito, strIDDestinazione, strIDTipoVendita, strIDEnteEmittente, strIDCliente, strIDZona, strIDSocieta
                        If Not OCT.Insert Then
                            intRowError = RowOrderInsertError
                            GoTo LogResult
                        End If
                        strIDOrdine = OCT.DS.Fields("0").NewValue
                    Else
                        strIDTipoVendita = OCT.DS.Fields("16").value
                    End If
                    
                    ' ---  Scrittura della riga di Consegna
                    RecordToVVC VVC, mobjInFile, strIDTrip, strIDRigaConsegna, _
                                strIDSequenza, strIDCliente, strIDOrdine, strInizioViaggio, _
                                strIDZona, strIDTipoVendita, strIDSocieta, varDatiCliDes
                                
                    If Not VVC.Insert Then
                        intRowError = RowTripInsertError
                        GoTo LogResult
                    End If
                    strIDConsegna = strIDRigaConsegna
                    strIDRigaConsegna = Format$(CInt(strIDRigaConsegna) + 10, "000")
                    strIDSequenza = Format$(CInt(strIDSequenza) + 10, "000")
                    strIDScarico = 10
                    
                    For i = LBound(varQuantita) To UBound(varQuantita) ' --- Inserisco gli scarichi
                            strIDProdotto = mobjInFile.Item(CStr(26 + i)).value
                            If strIDProdotto <> vbNullString Then
                                varQuantita(i) = mobjInFile.Item(CStr(19 + i)).value
                                 
                                    strPosdoganale(i) = mobjInFile.Item(CStr(30 + i)).value
                                    strIDProdotto = mobjInFile.Item(CStr(26 + i)).value
                                    With APR.DS.MultiSoc
                                        .DS.Fields("Codice").FindValue = strIDProdotto
                                        .DS.Fields("Società").FindValue = strSocDefault
                                        If Not .GetList Then
                                            intRowError = NoProdotto
                                            GoTo LogResult
                                        Else
                                            varDummy = .DS.List
                                            strIDImballo = varDummy(1, 0)
                                        End If
                                    End With
                                    strIDOrdineRiga = OCR.RowFromProduct(strIDOrdine, strIDProdotto)
                                    If strIDOrdineRiga = vbNullString Then
                                        If OCR.GetNewRowOrdine(strIDOrdine, strIDOrdineRiga) Then
                                             If varQuantita(i) > 0 Then
                                                strIDOrdineRiga = strIDOrdineRiga + 10
                                                RecordToOCR OCR, strIDOrdine, strIDOrdineRiga, strIDProdotto, strIDImballo, strPosdoganale(i), varQuantita(i)
                                                If Not OCR.Insert(True) Then
                                                    intRowError = RowOrderInsertError
                                                    GoTo LogResult
                                                End If
                                            End If
                                        End If
                                    Else
                                        Dim blOkUpdateQta As Boolean
                                        blOkUpdateQta = OCR.UpdateQtaConcordata(strIDOrdine, strIDOrdineRiga, varQuantita(i))
                                        If Not blOkUpdateQta Then
                                            intRowError = NoRigaOrdine
                                            GoTo LogResult
                                        End If
                                    End If
                                    If varQuantita(i) > 0 Then
                                        RecordToVVP VVP, mobjInFile, strIDTrip, strIDConsegna, _
                                                strIDScarico, strIDProdotto, strIDImballo, _
                                                strPosdoganale(i), varQuantita(i), strIDOrdine, strIDOrdineRiga
                                        If Not VVP.Insert Then
                                            intRowError = RowProductInsertError
                                            GoTo LogResult
                                        End If
                                        strIDScarico = strIDScarico + 10
                                    End If
                            End If
                    Next i
                    ' ---
                End If
                ' ---
            End If
        End If
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If intRowError <> NoError Then
            lngCountFail = lngCountFail + 1
        End If
        
        If blnNewTrip Then
            If intRowError = NoError Then
                VVT.DS.DA.EndTrans 0
            Else
                VVT.DS.DA.EndTrans 1
            End If
            
            If intRowError = NoError Then
                lngCountOK = lngCountOK + 1
            End If
            lngCountToGo = lngCountToGo - 1
            lngCountTrips = lngCountTrips + 1
            TDBGrid1.Bookmark = lngCountTrips
            lblLog(2) = lngCountOK
            lblLog(4) = lngCountFail
            lblLog(5) = lngCountToGo
        End If

        Select Case intRowError
            Case NoError
                If blnNewTrip Then
                    strErrorMessage = "OK - Num. Viaggio : " & strIDTrip
                Else
                    If blnTripOK Then
                        strErrorMessage = "OK - Num. Consegna : " & strIDConsegna
                    End If
                End If
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoRigaOrdine
                strErrorMessage = "Riga Ordine non trovata"
            Case NoMagazzino
                strErrorMessage = "Codice Magazzino di Deposito non trovato"
            Case NoUT
                strErrorMessage = "Codice Unità di Trasporto non trovato"
            Case NoMotrice
                strErrorMessage = "Codice Motrice non trovato"
            Case NoRimorchio
                strErrorMessage = "Codice Rimorchio non trovato"
            Case NoOrderID
                strErrorMessage = "Ordine non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case NoCliente
                strErrorMessage = "Codice Cliente non trovato"
            Case BadIDTripMover
                strErrorMessage = "Numero Viaggio mancante"
            Case TripInsertError
                strErrorMessage = "Errore durante l'inserimento del viaggio"
            Case RowTripInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di consegna"
            Case RowProductInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di scarico"
            Case NoProductInParam
                strErrorMessage = "Prodotto non trovato tra i parametri di sistema"
            Case OrderUpdateError
                strErrorMessage = "Errore durante l'aggiornamento dello stato dell'Ordine"
        End Select
        XALog.value(Row, 2) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = Row
        End With
        ' ---

        DoEvents
    Next Row
    ' ---

    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long

    VVT.DS.DA.EndTrans 1
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Il file non è stato trovato."
            strInfo = "Nome del file: " & strFileName
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
            If (strIDImballoDefault = vbNullString) Then
                strInfo = strInfo & "Imballo" & vbNewLine
            End If
            If (strIDPosDefault = vbNullString) Then
                strInfo = strInfo & "Posizione Doganale" & vbNewLine
            End If
            If (strIDEnteEmittente = vbNullString) Then
                strInfo = strInfo & "Ente Emittente" & vbNewLine
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
    Set C = ErrH.SetError(ErrCode, "frmVVTImportMover", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    DisplayError ErrH.LastError
EndProc:
    VVT.DS.DA.EndTrans 0
    mobjInFile.CloseFile
    lblLog(4) = lngCountFail
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub
Private Sub RecordToOCR(OCR As COrdineRiga, IDordine As String, _
                        IDRiga As String, IDProdotto As String, _
                        IDImballo As String, IDPosDoganale As String, _
                        Qta As Long)
    With OCR.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDordine
        .Item("0").Updated = True
        .Item("1").NewValue = IDRiga
        .Item("1").Updated = True
        .Item("2").NewValue = IDProdotto
        .Item("2").Updated = True
        .Item("3").NewValue = IDImballo
        .Item("3").Updated = True
        .Item("4").NewValue = IDPosDoganale
        .Item("4").Updated = True
        .Item("7").NewValue = Qta
        .Item("7").Updated = True
        .Item("11").NewValue = Qta
        .Item("11").Updated = True
        .Item("21").NewValue = ""
        .Item("21").Updated = True
    End With
End Sub
Private Sub RecordToOCT(OCT As COrdine, _
                        IDDeposito As String, IDDestinazione As String, _
                        IDTipoVendita As String, IDEnteEmittente As String, _
                        IDCliente As String, IDZona As String, IDSoc As String)
    
    Dim strDtModifica As String
    Dim strHrModifica As String
    Dim strDtConsegna As String
    
    With OCT.DS.Fields
        .ResetFields
         .Item("1").NewValue = IDSoc
        .Item("1").Updated = True
        .Item("2").NewValue = IDEnteEmittente
        .Item("2").Updated = True
        .Item("3").NewValue = Format$(Date, "dd/mm/yyyy")
        .Item("3").Updated = True
        .Item("4").NewValue = Format$(Time, "hh:mm")
        .Item("4").Updated = True
        ' --- Data inserimento ordine
'        If Len(Trim$(InFile("1").value)) > 0 Then
'            strDtModifica = InFile("1").value
'            strDtModifica = Right$(strDtModifica, 2) & _
'                            "/" & Mid$(strDtModifica, 5, 2) & _
'                            "/" & Left$(strDtModifica, 4)
'
'            strHrModifica = InFile("2").value
'            strHrModifica = Left$(strHrModifica, 2) & ":" & Mid$(strHrModifica, 3, 2)
'            .Item("5").NewValue = strDtModifica
'            .Item("5").Updated = True
'            .Item("6").NewValue = strHrModifica
'            .Item("6").Updated = True
'        End If
        ' --- Data consegna richiesta
'        strDtConsegna = Format$(mskDataConsegna.Text, "dd/mm/yyyy")
        .Item("7").NewValue = strDataViaggio
        .Item("7").Updated = True
        ' ---
        ' --- Data consegna pianificata
        .Item("8").NewValue = strDtConsegna
        .Item("8").Updated = True
    
        '.Item("10").NewValue = Trim$(Left$(InFile("12").value, .Item("10").MaxLength))
        '.Item("10").Updated = True
        .Item("12").NewValue = IDCliente
        .Item("12").Updated = True
        .Item("13").NewValue = IDDestinazione
        .Item("13").Updated = True
        .Item("16").NewValue = IDTipoVendita
        .Item("16").Updated = True
        .Item("17").NewValue = IDDeposito
        .Item("17").Updated = True
        .Item("18").NewValue = IDZona
        .Item("18").Updated = True
        
        .Item("21").NewValue = "R"
        .Item("21").Updated = True
        
        .Item("22").NewValue = "A"
        .Item("22").Updated = True

        .Item("23").NewValue = "M"
        .Item("23").Updated = True
        
        .Item("26").NewValue = "MM"
        .Item("26").Updated = True
    End With

End Sub
Private Sub OpenTextFile()
    On Error GoTo ErrHandler
    
    Dim strFileName As String
    Dim strLine As String
    Dim strBuffer As String
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim lngRows As Long
    Dim lngCountTrips As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    cmdImport.Enabled = False
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    lngCountTrips = 1
    Do While Not mobjInFile.EOF
        strBuffer = mobjInFile.GetLine(False)
        If Left$(strBuffer, 2) = "T0" Then ' --- Testata viaggio
            mobjInFile.ParseFixedLenFields strBuffer, 1, 15
            With XA
                .ReDim 1, lngCountTrips, 0, 13
                For i = 0 To 13
                    .value(.UpperBound(1), i) = mobjInFile.Item(i + 2).value
                Next i
            End With
            lngCountTrips = lngCountTrips + 1
            With XALog
                .ReDim 1, lngRows, 0, 2
                .value(.UpperBound(1), 0) = mobjInFile.Item("2").value
            End With
            lngRows = lngRows + 1
            ' ---
        Else
            mobjInFile.ParseFixedLenFields strBuffer, 16, 36 ' --- Riga viaggio
            With XALog
                .ReDim 1, lngRows, 0, 2
                .value(.UpperBound(1), 0) = mobjInFile.Item("18").value
                .value(.UpperBound(1), 1) = mobjInFile.Item("19").value
            End With
            lngRows = lngRows + 1
            ' ---
        End If
    Loop
        
    cmdImport.Enabled = True
    With TDBGrid1
        .Array = XA
        .ReBind
        .Bookmark = 1
    End With
    With tdbLog
        .Array = XALog
        .ReBind
        .Bookmark = 1
    End With
    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngCountTrips - 1
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmVVTImportMover1", _
                        "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmVVTImportMover1", _
                                "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub
Private Sub RecordToVVP(VVP As CViaggioProdotto, InFile As CTextFile, _
                        IDViaggio As String, IDConsegna As String, _
                        IDScarico As String, IDProdotto As String, _
                        IDImballo As String, IDPosDoganale As Variant, _
                        Quantita As Variant, IDordine As String, _
                        IDRigaOrdine As String)
    With VVP.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDViaggio
        .Item("0").Updated = True
        .Item("1").NewValue = IDConsegna
        .Item("1").Updated = True
        .Item("2").NewValue = IDScarico
        .Item("2").Updated = True
        .Item("3").NewValue = IDProdotto
        .Item("3").Updated = True
        .Item("4").NewValue = IDImballo
        .Item("4").Updated = True
        .Item("5").NewValue = IDPosDoganale
        .Item("5").Updated = True
        .Item("7").NewValue = Quantita
        .Item("7").Updated = True
        .Item("11").NewValue = Quantita
        .Item("11").Updated = True
        .Item("32").NewValue = IDordine
        .Item("32").Updated = True
        .Item("33").NewValue = IDRigaOrdine
        .Item("33").Updated = True
    End With
End Sub
Private Sub RecordToVVC(VVC As CViaggioConsegna, InFile As CTextFile, IDViaggio As String, _
                        IDConsegna As String, IDSequenza As String, IDCliente As String, _
                        IDordine As String, InizioViaggio As String, IDZona As String, _
                        IDTipoVendita As String, IDSocDefault As String, ParamArray DatiClienteDest())
    Dim i As Integer
    Dim strTempoViaggio As String
    Dim strOraArrivo As String
    Dim strIDDestinazione As String
    
    With VVC.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDViaggio
        .Item("0").Updated = True
        .Item("1").NewValue = IDConsegna
        .Item("1").Updated = True
        .Item("2").NewValue = IDSequenza
        .Item("2").Updated = True
        .Item("3").NewValue = IDSocDefault
        .Item("3").Updated = True
        .Item("4").NewValue = IDCliente
        .Item("4").Updated = True
        For i = 5 To 11
            .Item(CStr(i)).NewValue = DatiClienteDest(0)(i - 5)
            .Item(CStr(i)).Updated = True
        Next i
        strIDDestinazione = mobjInFile.Item("19").value
        If IsNumeric(strIDDestinazione) Then
            strIDDestinazione = Format$(strIDDestinazione, "0000")
        End If
        .Item("12").NewValue = strIDDestinazione 'Trim$(Left$(InFile("19").value, .Item("12").MaxLength))
        .Item("12").Updated = True
        For i = 13 To 19
            .Item(CStr(i)).NewValue = DatiClienteDest(0)(i - 6)
            .Item(CStr(i)).Updated = True
        Next i
        .Item("20").NewValue = Trim$(InFile("25").value)
        .Item("20").Updated = True
        .Item("21").NewValue = "A"
        .Item("21").Updated = True
        .Item("23").NewValue = "A"
        .Item("23").Updated = True
        If InFile("26").value = vbNullString Then
            InFile("26").value = "0"
        End If
        strOraArrivo = Format$(Left$(InFile("26").value, 2), "00") & ":" & Format$(Right$(InFile("26").value, 2), "00")
        .Item("24").NewValue = strOraArrivo ' strTempoViaggio
        .Item("24").Updated = True
        .Item("26").NewValue = IDZona
        .Item("26").Updated = True
        .Item("27").NewValue = IDTipoVendita
        .Item("27").Updated = True
        .Item("37").NewValue = IDordine
        .Item("37").Updated = True
    End With
End Sub
Private Sub RecordToVVT(VVT As CViaggio, InFile As CTextFile, _
                        IDMotrice As String, IDRimorchio As String, IDDeposito As String, IDMagazzino As String, IDSocieta As String, IDEnteEmittente As String, _
                        IDvettore As String, IDautista1 As String, IDautista2 As String, IDVettoreT As String)

    Dim strOraInizioV As String
    Dim strOraInizioC As String
    Dim strTempoCarico As String
    
    Dim strSequenza As String
    
    With VVT.DS.Fields
        .ResetFields
        'If CInt(InFile("6").value) = 0 Then
         '   strSequenza = "1"
        'Else
        strSequenza = InFile("6").value
        'End If
        .Item("2").NewValue = Trim$(Left$(strSequenza, .Item("2").MaxLength))
        .Item("2").Updated = True
        .Item("3").NewValue = IDSocieta
        .Item("3").Updated = True
        .Item("14").NewValue = Trim$(Left$(InFile("3").value, .Item("14").MaxLength))
        .Item("14").Updated = True
        .Item("15").NewValue = IDMotrice
        .Item("15").Updated = True
        .Item("16").NewValue = IDRimorchio
        .Item("16").Updated = True
        strDataViaggio = Mid$(InFile("7").value, 7, 2) & "/" & Mid$(InFile("7").value, 5, 2) & "/" & Left$(InFile("7").value, 4)
        .Item("4").NewValue = strDataViaggio
        .Item("4").Updated = True
        .Item("6").NewValue = strDataViaggio
        .Item("6").Updated = True
        .Item("8").NewValue = strDataViaggio
        .Item("8").Updated = True
        .Item("21").NewValue = IDEnteEmittente
        .Item("21").Updated = True
        strOraInizioC = Format$(Left$(InFile("8").value, 2), "00") & ":" & Format$(Right$(InFile("8").value, 2), "00")
        strOraInizioV = Format$(Left$(InFile("9").value, 2), "00") & ":" & Format$(Right$(InFile("9").value, 2), "00")
        strTempoCarico = DateDiff("n", strOraInizioC, strOraInizioV)
'        strTempoCarico = CInt(strTempoCarico) \ 60 & DA.SqlDecSeparator & CInt(strTempoCarico) Mod 60
        strTempoCarico = CInt(strTempoCarico) \ 60 & DataConnect.SQLDecSeparator & CInt(strTempoCarico) Mod 60
        .Item("5").NewValue = strOraInizioC
        .Item("5").Updated = True
        .Item("7").NewValue = strOraInizioV
        .Item("7").Updated = True
        .Item("30").NewValue = strTempoCarico
        .Item("30").Updated = True
        .Item("31").NewValue = ""
        .Item("31").Updated = True
        .Item("25").NewValue = InFile("10").value    ' Km Previsti
        .Item("25").Updated = True
        .Item("26").NewValue = "A"
        .Item("26").Updated = True
        .Item("17").NewValue = IDvettore
        .Item("17").Updated = True
        .Item("18").NewValue = IDVettoreT
        .Item("18").Updated = True
        .Item("19").NewValue = IDautista1
        .Item("19").Updated = True
        If IDautista2 <> vbNullString Then
            .Item("20").NewValue = IDautista2
            .Item("20").Updated = True
        End If
        
        .Item("21").NewValue = IDEnteEmittente
        .Item("21").Updated = True
        .Item("27").NewValue = InFile("10").value    ' Km Definitivi
        .Item("27").Updated = True
        .Item("28").NewValue = "A"
        .Item("28").Updated = True
        .Item("22").NewValue = IDDeposito
        .Item("22").Updated = True
        .Item("23").NewValue = IDMagazzino
        .Item("23").Updated = True
        .Item("32").NewValue = IDDeposito
        .Item("32").Updated = True
        .Item("43").NewValue = "R"
        .Item("43").Updated = True
    End With
End Sub

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    mskFileName = gmRegistry.GetSetting("Options", "TripsImportFile", , True)
    
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
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = ID_PARAM
        If AXX.Find Then
            MaskEdBox1 = AXX.DS.Fields("16").value
            If DataClass.DS.ValidateClientField("4", MaskEdBox1) Then
                    FieldUpdateDescription "4", Me, MaskEdBox1
            End If
        End If
    End With
    
    Set OCT = New COrdine
    With OCT.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineRiga
    With OCR.DS
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
'        Set .DA = DataConnect.DA
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
    
    Set AMD = New CMagazzinoDeposito
    With AMD.DS
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
        ''Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
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
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
    With mobjInFile
            ' Testate Viaggi
            .Add "Tipo Record", 2, "1"
            .Add "Numero Viaggio", 3, "2"
            .Add "Unità", 3, "3"
            .Add "Motrice", 10, "4"
            .Add "Rimorchio", 10, "5"
            .Add "Sequenza", 2, "6"
            .Add "Data", 8, "7"
            .Add "Ora inizio carico", 4, "8"
            .Add "Ora inizio viaggio", 4, "9"
            .Add "Km previsti", 5, "10"
            .Add "Deposito", 3, "11"
            .Add "Secondo Autista", 20, "12"
            .Add "Codice Magazzino", 3, "13"
            .Add "Filler 1", 11, "14"
            .Add "Filler 2", 1, "15"
            
            ' Viaggi Dettaglio
            .Add "Tipo Record D", 2, "16"
            .Add "Numero Viaggio D", 3, "17"
            .Add "Progressivo", 3, "18"
            .Add "Societa", 2, "36"
            .Add "Destinazione", 6, "19"
            .Add "Litri prodotto 1", 5, "20", , fftNumerico
            .Add "Litri prodotto 2", 5, "21", , fftNumerico
            .Add "Litri prodotto 3", 5, "22", , fftNumerico
            .Add "Litri prodotto 4", 5, "23", , fftNumerico
            .Add "Ordine", 7, "24"
            .Add "Km parziali", 5, "25", , fftNumerico
            .Add "Orario arrivo", 4, "26"
            .Add "Codice prodotto 1", 6, "27"
            .Add "Codice prodotto 2", 6, "28"
            .Add "Codice prodotto 3", 6, "29"
            .Add "Codice prodotto 4", 6, "30"
            .Add "Pos. dog. prodotto 1", 3, "31"
            .Add "Pos. dog. prodotto 2", 3, "32"
            .Add "Pos. dog. prodotto 3", 3, "33"
            .Add "Pos. dog. prodotto 4", 3, "34"
            .Add "Filler 2", 1, "35"
        End With
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "TripsImportFile", mskFileName, True
    Set XA = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set OCR = Nothing
    Set OCT = Nothing
    Set AUT = Nothing
    Set AAM = Nothing
    Set ACL = Nothing
    Set AMD = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    Set VVP = Nothing
    Set APR = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set DataClass = Nothing
    Set frmVVTImportMover1 = Nothing
End Sub
Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub


