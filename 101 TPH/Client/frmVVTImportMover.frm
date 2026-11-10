VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTImportMover 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE VIAGGI DA MOVER"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   7680
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
         Picture         =   "frmVVTImportMover.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmVVTImportMover.frx":0102
         TabIndex        =   14
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   8475
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
            TextSave        =   "14:38"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "27/01/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   3615
         Left            =   7560
         TabIndex        =   12
         Top             =   1080
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
         Height          =   3615
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   7335
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   3255
            Left            =   120
            OleObjectBlob   =   "frmVVTImportMover.frx":0204
            TabIndex        =   5
            Top             =   240
            Width           =   7095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   4680
         Width           =   9135
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   2415
            Left            =   120
            OleObjectBlob   =   "frmVVTImportMover.frx":37EC
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
         Picture         =   "frmVVTImportMover.frx":A92E
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
         Picture         =   "frmVVTImportMover.frx":AE60
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
         Picture         =   "frmVVTImportMover.frx":AF62
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
         Picture         =   "frmVVTImportMover.frx":B494
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
         Picture         =   "frmVVTImportMover.frx":B596
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
         Picture         =   "frmVVTImportMover.frx":B698
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
               Picture         =   "frmVVTImportMover.frx":B79A
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTImportMover.frx":B8AC
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTImportMover.frx":B9BE
               Key             =   "Editing"
            EndProperty
         EndProperty
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
Attribute VB_Name = "frmVVTImportMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3
Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineriga
Private ACL As CCliente
Private AUT As CUnitaDiTrasporto
Private AAM As CAutomezzo
Private AMD As CMagazzinoDeposito
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private mobjInFile As CTextFile
Private mstrIDEnte As String
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

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
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDDestinazionePrec As String
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
    Dim strIDAutistamover As String
    Dim strBadgeAutomezzo As String
'    Dim strCodDest As String
'    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDDepositoOrdine As String
    Dim strIDMagazzino As String
    Dim strIDZona As String
    Dim strIDCliente As String
    Dim strIDUT As String
    Dim strIDOrdine As String
    Dim strIDOrdineRiga As String
    Dim strInizioViaggio As String
    Dim strDataFineViaggio As String
    Dim strOraFineViaggio As String
    Dim strIDProgressivo As String
    Dim strIDProgressivoPrec As String
    Dim strErrorMessage As String
    Dim strLocalError As String
    Dim StrSociedad As String
    Dim lngCountTrips As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim lngKmRaggio As Long
    Dim Row As Long
    Dim Col As Integer
    Dim i As Integer
    Dim intIDOrdineRiga As Integer
    Dim blnNewTrip As Boolean
    Dim blnTripOK As Boolean
    Dim varDatiCliDes(0 To 13) As Variant
'    Dim varQuantita(1 To 10) As Variant
    Dim varQuantita As Variant
    Dim varOraArrivo As Variant
    Dim intRowError As Errors
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
        strIDImballoDefault = Trim$(.Fields("1").value)
        strIDPosDefault = Trim$(.Fields("2").value)
'        strIDEnteEmittente = Trim$(.Fields("16").value)
        For i = 1 To 10
            strIDProdotti(i) = Trim$(.Fields(CStr(i + 5)).value)
        Next i
    End With

    If (strIDImballoDefault = vbNullString) Or _
       (strIDPosDefault = vbNullString) Then 'Or _
       (strIDEnteEmittente = vbNullString) Then
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
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If

    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        strLocalError = vbNullString
        blnNewTrip = False
        strBuffer = mobjInFile.GetLine(False)
        If Left$(strBuffer, 2) = "T0" Then
            '--- Testata viaggio
            mobjInFile.ParseFixedLenFields strBuffer, 1, 17 ' 16
            blnNewTrip = True
            strIDProgressivoPrec = vbNullString
            strIDAutista1 = vbNullString
            
            VVT.DS.DA.BeginTransaction
            blnTripOK = False
            
            ' --- Controllo che esistano Unità di Trasporto, Magazzino, Automezzi
            strIDUT = mobjInFile("3").value
            strIDMotrice = mobjInFile("4").value
            strIDRimorchio = mobjInFile("5").value
            strIDDeposito = mobjInFile("11").value
            strDataFineViaggio = mobjInFile("12").value
            If Len(strDataFineViaggio) > 0 Then
                strDataFineViaggio = Right(strDataFineViaggio, 2) & "/" & Mid(strDataFineViaggio, 5, 2) & "/" & Left(strDataFineViaggio, 4)
                If Not IsDate(strDataFineViaggio) Then
                    strDataFineViaggio = vbNullString
                End If
            End If
            strOraFineViaggio = mobjInFile("39").value
            If Len(strOraFineViaggio) > 0 Then
                strOraFineViaggio = Left(strOraFineViaggio, 2) & ":" & Right(strOraFineViaggio, 2)
                If Not IsDate(strOraFineViaggio) Then
                    strOraFineViaggio = vbNullString
                End If
            End If
            strIDAutista1 = mobjInFile("14").value
            strIDAutistamover = mobjInFile("14").value
            strInizioViaggio = Format$(Left$(mobjInFile("9").value, 2), "00") & ":" & Format$(Right$(mobjInFile("9").value, 2), "00")
            
            If strIDDeposito = vbNullString Then
                intRowError = NoDeposito
                GoTo LogResult
            End If
'            If strIDMagazzino = vbNullString Then
            mobjInFile("14").value = mobjInFile("11").value
            strIDMagazzino = strIDDeposito
'            End If
            If strIDUT = vbNullString Then
                intRowError = noUT
                GoTo LogResult
            End If
            
            With AMD.DS
                .Fields("0").FindValue = strIDMagazzino
                .Fields("1").FindValue = strIDDeposito
                If Not AMD.Find Then
                    intRowError = NoMagazzino
                    GoTo LogResult
                End If
                If .HowMany = 0 Then
                    intRowError = NoMagazzino
                    GoTo LogResult
                End If
            End With
            
            With AUT.DS
                .Fields("0").FindValue = strIDUT
                If Not AUT.Find Then
                    intRowError = noUT
                    GoTo LogResult
                End If
                If .HowMany = 0 Then
                    intRowError = noUT
                    GoTo LogResult
                End If
                If strIDMotrice = vbNullString Then
                    strIDMotrice = .Fields("2").value
                End If
                If strIDRimorchio = vbNullString Then
                    strIDRimorchio = .Fields("3").value
                End If
                strIDVettore = .Fields("6").value
                strIDVettoreT = .Fields("7").value
                strBadgeAutomezzo = .Fields("14").value
                If Len(strIDVettoreT) = 0 Then
                    strIDVettoreT = strIDVettore
                End If
                ' JG-SM 13/11/09 Introduce siempre el autista de Resoil seleccionando conductor 1 o 2 segun fichero mover
'                If Len(strIDAutista1) = 0 Then
                    If strIDAutistamover = "002" Then
                    strIDAutista1 = .Fields("5").value
                    Else
                    strIDAutista1 = .Fields("4").value
                    End If
'                End If
'                strIDAutista2 = .Fields("5").value
                strIDEnteEmittente = .Field("Ente Emittente").value
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
            
            ' JG 26/1/12 Busco la sociedad en la orden
            strIDOrdine = Right(mobjInFile.Item("21").value, 2) & Left(mobjInFile.Item("30").value, 5)
            
            strIDOrdine = Right(XALog.value(Row + 1, 2), 7)

            OCT.DS.Fields("0").FindValue = strIDOrdine
            If Not OCT.Find Then
                intRowError = NoOrderID
                GoTo LogResult
            End If
            If OCT.DS.HowMany = 0 Then
                intRowError = NoOrderID
                GoTo LogResult
            End If
 
            StrSociedad = OCT.DS.Fields("1").value
 
            ' ---
    
            
            ' --- Inserisco la testata del viaggio
            RecordToVVT VVT, mobjInFile, strIDMotrice, strIDRimorchio, _
                        strIDVettore, strIDVettoreT, strIDAutista1, _
                        strIDEnteEmittente, strBadgeAutomezzo, strDataFineViaggio, strOraFineViaggio, StrSociedad
                        
            '-- DP 05/08/08
            If Not VVT.CheckSequenza(VVT.DS.Fields("6").NewValue, VVT.DS.Fields("2").NewValue, VVT.DS.Fields("14").NewValue, VVT.DS.Fields("22").NewValue, "") Then
                strLocalError = "Il numero di sequenza è già stato utilizzato"
                intRowError = TripInsertError
'                GoTo ErrHandler
                GoTo LogResult
            End If
            
            If Not VVT.Insert(strLocalError) Then
                intRowError = TripInsertError
                GoTo LogResult
            End If
            strIDTrip = VVT.DS.Fields("0").NewValue
            strIDRigaConsegna = "010"
            strIDSequenza = "010"
            strIDDestinazionePrec = vbNullString
            blnTripOK = True
            ' ---
        Else
            If blnTripOK Then
'                mobjInFile.ParseFixedLenFields strBuffer, 17, 37 ' 28
                mobjInFile.ParseFixedLenFields strBuffer, 19, 39 '18, 38
                strIDDepositoOrdine = strIDDeposito
                ' --- Inserisco una nuova riga di consegna
                If strIDTrip <> vbNullString Then
                    
                    strIDProgressivo = mobjInFile.Item("18").value
                    
                    If strIDProgressivo <> strIDProgressivoPrec Then
                        strIDProgressivoPrec = strIDProgressivo
                    
                    
                        ' --- 11/08/99
                        ' --- La riga seguente è stata modificata per
                        ' --- mandare a mover i numeri ordine nello stesso
                        ' --- formato del numero ordine Agip (zznnnnnn)
                        strIDOrdine = Right$(mobjInFile.Item("30").value, 7)
    '                    strIDOrdine = Right$(mobjInFile.Item("24").value, 6)
    '                    strIDOrdine = Right$(mobjInFile.Item("30").value, 6)
                        ' ---
                        
                        If strIDOrdine = vbNullString Then
                            intRowError = NoOrderID
                            GoTo LogResult
                        End If
                        
                        ' --- 11/08/99
                        ' --- La riga seguente è stata aggiunta per
                        ' --- mandare a mover i numeri ordine nello stesso
                        ' --- formato del numero ordine Agip (zznnnnnn)
    '                    strIDOrdine = "0" & strIDOrdine
                        ' ---
                        
                        OCT.DS.Fields("0").FindValue = strIDOrdine
                        If Not OCT.Find Then
                            intRowError = NoOrderID
                            GoTo LogResult
                        End If
                        If OCT.DS.HowMany = 0 Then
                            intRowError = NoOrderID
                            GoTo LogResult
                        End If
                        
                       
                        If OCT.DS.Fields("23").value = "V" Then
                            intRowError = OrderInTrip
                            GoTo LogResult
                        End If
                        If Len(OCT.DS.Fields("37").value) > 0 Then
                            lngKmRaggio = OCT.DS.Fields("37").value
                        Else
                            lngKmRaggio = 0
                        End If
                
                        strIDTipoVendita = OCT.DS.Fields("16").value
                        strIDZona = OCT.DS.Fields("18").value
                        strIDDepositoOrdine = OCT.DS.Fields("17").value
                        ' ---
                        
                        strIDDestinazione = mobjInFile.Item("19").value
                        If IsNumeric(strIDDestinazione) Then
                            strIDDestinazione = Format$(strIDDestinazione, "00000000")
                        End If
                        
                            
                        ' --- Ricavo il cliente e la zona dalla destinazione
                        With ADZ.DS
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
    '                            strIDZona = .Fields("16").value
                            strIDCliente = .Fields("8").value
                        End With
                        ' ---
                        
                        ' --- Leggo i dati anagrafici del cliente
                        With ACL.DS
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
                        ' ---
                        
                        If strIDCliente = ID_SOC_CEPSA Then
                            If strIDDestinazione <> strIDDestinazionePrec Then
                                varOraArrivo = Format$(Left$(mobjInFile("32").value, 2), "00") & ":" & Format$(Right$(mobjInFile("32").value, 2), "00")
                                strIDDestinazionePrec = strIDDestinazione
                            End If
                        Else
                            varOraArrivo = Format$(Left$(mobjInFile("32").value, 2), "00") & ":" & Format$(Right$(mobjInFile("32").value, 2), "00")
                        End If
                        
                        RecordToVVC VVC, mobjInFile, strIDTrip, strIDRigaConsegna, _
                                    strIDSequenza, strIDCliente, strIDOrdine, strInizioViaggio, _
                                    strIDZona, strIDTipoVendita, lngKmRaggio, varOraArrivo, varDatiCliDes
                                    
                        If Not VVC.Insert Then
                            intRowError = RowTripInsertError
                            GoTo LogResult
                        End If
                        strIDConsegna = strIDRigaConsegna
                        strIDRigaConsegna = Format$(CInt(strIDRigaConsegna) + 10, "000")
                        strIDSequenza = Format$(CInt(strIDSequenza) + 10, "000")
                        strIDScarico = 10
                    End If
                    
                    ' --- Inserisco gli scarichi
'                    For i = LBound(varQuantita) To UBound(varQuantita)
'                    varQuantita(i) = mobjInFile.Item(CStr(20 + i - 1)).value
'                    strIDProdotto = strIDProdotti(i)
                    strIDProdotto = mobjInFile.Item("20").value
                    varQuantita = mobjInFile.Item("21").value
'                    If CDbl(varQuantita(i)) <> 0 Then
                    If CDbl(varQuantita) <> 0 Then
                        If strIDProdotto = vbNullString Then
                            intRowError = NoProductInParam
                            GoTo LogResult
                        End If
                        strIDOrdineRiga = OCR.RowFromProduct(strIDOrdine, strIDProdotto, varQuantita, strLocalError)
                        If strIDOrdineRiga = vbNullString Then
                            intRowError = NoRigaOrdine
                            GoTo LogResult
                        End If
                        
                        RecordToVVP VVP, mobjInFile, strIDTrip, strIDConsegna, _
                                    strIDScarico, strIDProdotto, strIDImballoDefault, _
                                    strIDPosDefault, varQuantita, strIDOrdine, strIDOrdineRiga
                        If Not VVP.Insert Then
                            intRowError = RowProductInsertError
                            GoTo LogResult
                        End If
                        strIDScarico = strIDScarico + 10
                    End If
'                    Next i
                    
                    ' --- Verifichiamo se il deposito sull'ordine è diverso rispetto a quello
                    ' --- del viaggio; in questo caso facciamo una segnalazione all'utente
                    If strIDDeposito <> strIDDepositoOrdine Then
                        intRowError = InvalidDep
                        GoTo LogResult
                    End If
                    ' ---
                End If
                ' ---
            End If
        End If
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError <> NoError) Then
            lngCountFail = lngCountFail + 1
        End If
        
        If blnNewTrip Then
            If (intRowError = NoError) Or (intRowError = OrderInTrip) Or (intRowError = InvalidDep) Then
                VVT.DS.DA.EndTrans 0
            Else
                VVT.DS.DA.EndTrans 1
            End If
            
            If intRowError = NoError Then
                lngCountOK = lngCountOK + 1
            End If
            lngCountToGo = lngCountToGo - 1
            lngCountTrips = lngCountTrips + 1
            TDBGrid1.bookmark = lngCountTrips
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
            Case InvalidDep
                strErrorMessage = "Deposito modificato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoRigaOrdine
                If Len(strLocalError) = 0 Then
                    strErrorMessage = "Riga Ordine non trovata"
                Else
                    strErrorMessage = strLocalError
                End If
            Case NoMagazzino
                strErrorMessage = "Codice Magazzino di Deposito non trovato"
            Case noUT
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
                If Len(strLocalError) = 0 Then
                    strErrorMessage = "Errore durante l'inserimento del viaggio"
                Else
                    strErrorMessage = strLocalError
                End If
            Case RowTripInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di consegna"
            Case RowProductInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di scarico"
            Case NoProductInParam
                strErrorMessage = "Prodotto non trovato tra i parametri di sistema"
            Case OrderUpdateError
                strErrorMessage = "Errore durante l'aggiornamento dello stato dell'Ordine"
            Case OrderInTrip
                strErrorMessage = "Ordine " & OCT.DS.Fields("0").value & " già inserito nel viaggio " & OCT.GetTripID
        End Select
        XALog.value(Row, 3) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = Row
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
'            If (strIDEnteEmittente = vbNullString) Then
'                strInfo = strInfo & "Ente Emittente" & vbNewLine
'            End If

        Case NoImportData
            strMessage = "Non ci sono dati da importare."
            strInfo = vbNullString
        
        Case TripInsertError
            If Len(strLocalError) = 0 Then
                strInfo = "Errore durante l'inserimento del viaggio"
            Else
                strInfo = strLocalError
            End If
        
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
        If Left$(strBuffer, 2) = "T0" Then
            ' --- Testata viaggio
            mobjInFile.ParseFixedLenFields strBuffer, 1, 16 '15
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
            ' --- Riga viaggio
            mobjInFile.ParseFixedLenFields strBuffer, 17, 37 '16, 36
            
            With XALog
                .ReDim 1, lngRows, 0, 3
                .value(.UpperBound(1), 0) = mobjInFile.Item("17").value
                .value(.UpperBound(1), 1) = Right(mobjInFile.Item("34").value, 3) & Left(mobjInFile.Item("19").value, 5)
                .value(.UpperBound(1), 2) = Right(mobjInFile.Item("21").value, 2) & Left(mobjInFile.Item("30").value, 5)
            End With
            lngRows = lngRows + 1
            ' ---
        End If
    Loop
        
    cmdImport.Enabled = True
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
    lblLog(5) = lngCountTrips - 1
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmVVTImportMover", _
                        "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmVVTImportMover", _
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
                        IDImballo As String, IDPosDoganale As String, _
                        Quantita As Variant, IDOrdine As String, _
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
        .Item("32").NewValue = IDOrdine
        .Item("32").Updated = True
        .Item("33").NewValue = IDRigaOrdine
        .Item("33").Updated = True
    End With
End Sub

Private Sub RecordToVVC(VVC As CViaggioConsegna, InFile As CTextFile, IDViaggio As String, _
                        IDConsegna As String, IDSequenza As String, IDCliente As String, _
                        IDOrdine As String, InizioViaggio As String, IDZona As String, _
                        IDTipoVendita As String, lngKmRaggio As Long, OraArrivo As Variant, ParamArray DatiClienteDest())
    Dim i As Integer
    Dim strTempoViaggio As String
    Dim strOraArrivo As String
    Dim strIDDestinazione As String
    Dim strDataViaggio As String
    
    With VVC.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDViaggio
        .Item("0").Updated = True
        .Item("1").NewValue = IDConsegna
        .Item("1").Updated = True
        .Item("2").NewValue = IDSequenza
        .Item("2").Updated = True
        .Item("3").NewValue = IIf(IDCliente = "904", "001", IDCliente) 'ID_SOC
        .Item("3").Updated = True
        .Item("4").NewValue = IDCliente
        .Item("4").Updated = True
        For i = 5 To 11
            .Item(CStr(i)).NewValue = DatiClienteDest(0)(i - 5)
            .Item(CStr(i)).Updated = True
        Next i
        strIDDestinazione = mobjInFile.Item("19").value
        If IsNumeric(strIDDestinazione) Then
            strIDDestinazione = Format$(strIDDestinazione, "00000000")
        End If
        .Item("12").NewValue = strIDDestinazione 'Trim$(Left$(InFile("19").value, .Item("12").MaxLength))
        .Item("12").Updated = True
        For i = 13 To 19
            .Item(CStr(i)).NewValue = DatiClienteDest(0)(i - 6)
            .Item(CStr(i)).Updated = True
        Next i
'        .Item("20").NewValue = Trim$(InFile("25").value)
        .Item("20").NewValue = Trim$(InFile("31").value)
        .Item("20").Updated = True
        .Item("21").NewValue = "A"
        .Item("21").Updated = True
'        .Item("22").NewValue = Trim$(InFile("25").value)
'        .Item("22").Updated = True
        .Item("23").NewValue = "A"
        .Item("23").Updated = True
'        If InFile("26").value = vbNullString Then
'            InFile("26").value = "0"
'        End If
        If InFile("32").value = vbNullString Then
            InFile("32").value = "0"
        End If
        'strOraArrivo = Format$(Left$(InFile("32").value, 2), "00") & ":" & Format$(Right$(InFile("32").value, 2), "00")
        .Item("24").NewValue = OraArrivo 'strOraArrivo
        .Item("24").Updated = True
        .Item("26").NewValue = IDZona
        .Item("26").Updated = True
        .Item("27").NewValue = IDTipoVendita
        .Item("27").Updated = True
        .Item("37").NewValue = IDOrdine
        .Item("37").Updated = True
        #If VersionID = 1 Then
            If lngKmRaggio > 0 Then
                .Item("39").NewValue = lngKmRaggio
                .Item("39").Updated = True
            Else
'                .Item("39").NewValue = Trim$(InFile("27").value)
                .Item("39").NewValue = Trim$(InFile("33").value)
                .Item("39").Updated = True
            End If
'            .Item("40").NewValue = Trim$(InFile("28").value)
            .Item("40").NewValue = Trim$(InFile("34").value)
            .Item("40").Updated = True
            strDataViaggio = Mid$(InFile("7").value, 7, 2) & "/" & Mid$(InFile("7").value, 5, 2) & "/" & Left$(InFile("7").value, 4)
            .Item("41").NewValue = strDataViaggio
            .Item("41").Updated = True
        #End If
    End With
End Sub

Private Sub RecordToVVT(VVT As CViaggio, InFile As CTextFile, _
                        IDMotrice As String, IDRimorchio As String, _
                        IDVettore As String, IDVettoreT As String, _
                        IDAutista1 As String, _
                        IDEnteEmittente As String, BadgeAutomezzo As String, _
                        DataFineViaggio As String, OraFineViaggio As String, StrSociedad As String)

    Dim strOraInizioV As String
    Dim strOraInizioC As String
    Dim strTempoCarico As String
    Dim strDataViaggio As String
    Dim strSequenza As String
    
    With VVT.DS.Fields
        .ResetFields
        If CInt(InFile("6").value) = 0 Then
            strSequenza = "1"
        Else
            strSequenza = InFile("6").value
        End If
        .Item("2").NewValue = Trim$(Left$(strSequenza, .Item("2").MaxLength))
        .Item("2").Updated = True
        .Item("3").NewValue = StrSociedad ' ID_SOC

'        .Item("3").NewValue = Format(InFile("38").value, "000") ' ID_SOC
        .Item("3").Updated = True
'        .Item("11").NewValue = Trim$(Left$(InFile("3").value, .Item("11").MaxLength))
        .Item("11").NewValue = BadgeAutomezzo
        .Item("11").Updated = True
        .Item("14").NewValue = Trim$(Left$(InFile("3").value, .Item("14").MaxLength))
        .Item("14").Updated = True
        .Item("15").NewValue = IDMotrice
        .Item("15").Updated = True
        .Item("16").NewValue = IDRimorchio
        .Item("16").Updated = True
        .Item("17").NewValue = IDVettore
        .Item("17").Updated = True
        .Item("18").NewValue = IDVettoreT
        .Item("18").Updated = True
        .Item("19").NewValue = IDAutista1
        .Item("19").Updated = True
'        .Item("20").NewValue = IDAutista2
'        .Item("20").Updated = True
        strDataViaggio = Mid$(InFile("7").value, 7, 2) & "/" & Mid$(InFile("7").value, 5, 2) & "/" & Left$(InFile("7").value, 4)
        .Item("4").NewValue = strDataViaggio
        .Item("4").Updated = True
        .Item("6").NewValue = strDataViaggio
        .Item("6").Updated = True
        .Item("8").NewValue = strDataViaggio
        .Item("8").Updated = True
        strOraInizioC = Format$(Left$(InFile("8").value, 2), "00") & ":" & Format$(Right$(InFile("8").value, 2), "00")
        strOraInizioV = Format$(Left$(InFile("9").value, 2), "00") & ":" & Format$(Right$(InFile("9").value, 2), "00")
        strTempoCarico = DateDiff("n", strOraInizioC, strOraInizioV)
        strTempoCarico = CInt(strTempoCarico) \ 60 & DataConnect.SqlDecSeparator & CInt(strTempoCarico) Mod 60
        .Item("5").NewValue = strOraInizioC
        .Item("5").Updated = True
        .Item("7").NewValue = strOraInizioV
        .Item("7").Updated = True
        .Item("30").NewValue = strTempoCarico
        .Item("30").Updated = True
        .Item("31").NewValue = ""
        .Item("31").Updated = True
        .Item("25").NewValue = InFile("10").value
        .Item("25").Updated = True
        .Item("26").NewValue = "A"
        .Item("26").Updated = True
        .Item("27").NewValue = InFile("10").value
        .Item("27").Updated = True
        .Item("28").NewValue = "A"
        .Item("28").Updated = True
        .Item("22").NewValue = Trim$(Left$(InFile("11").value, .Item("22").MaxLength))
        .Item("22").Updated = True
        .Item("32").NewValue = Trim$(Left$(InFile("11").value, .Item("32").MaxLength))
        .Item("32").Updated = True
'        .Item("29").NewValue = Format$(Left$(InFile("12").value, 2), "00") & DataConnect.SqlDecSeparator & Format$(Right$(InFile("12").value, 2), "00")
'        .Item("29").Updated = True
        .Item("23").NewValue = Trim$(Left$(InFile("14").value, .Item("23").MaxLength))
        .Item("23").Updated = True
        .Item("21").NewValue = IDEnteEmittente
        .Item("21").Updated = True
        .Item("33").NewValue = IDEnteEmittente
        .Item("33").Updated = True
        .Item("43").NewValue = "E"
        .Item("43").Updated = True
        If IsDate(DataFineViaggio) And IsDate(OraFineViaggio) Then
            .Item("8").NewValue = DataFineViaggio
            .Item("8").Updated = True
            .Item("9").NewValue = OraFineViaggio
            .Item("9").Updated = True
            .Item("29").NewValue = DateDiff("n", strDataViaggio & " " & strOraInizioV, DataFineViaggio & " " & OraFineViaggio)
            .Item("29").Updated = True
        End If
        If IsNumeric(InFile("13").value) Then
            .Item("51").NewValue = InFile("13").value
            .Item("51").Updated = True
        End If
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
    
    ConfigureObjects Me
    mskFileName = gmRegistry.GetSetting("Options", "TripsImportFile", , True)
    
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
    
    Set VVT = New CViaggio
    With VVT.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AMD = New CMagazzinoDeposito
    With AMD.DS
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
    
    Set ACL = New CCliente
    With ACL.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AAM = New CAutomezzo
    With AAM.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
    #If VersionID = 1 Then
        With mobjInFile
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
'            .Add "Durata Viaggio", 4, "12"
'            .Add "Filler 1", 16, "13"
            .Add "Data Fine Viaggio", 8, "12"
            .Add "Ora Fine Viaggio", 4, "39"
'            .Add "Filler 1", 8, "13"
            .Add "Km rientro", 4, "13"
            .Add "Km da garage 1", 4, "40"
            .Add "Codice Autista", 3, "14"
            .Add "Società Viaggio", 2, "38"
            .Add "Filler 2", 1, "15"
            .Add "Tipo Record D", 2, "16"
            .Add "Numero Viaggio D", 3, "17"
            .Add "Progressivo", 3, "18"
            .Add "Filler 3", 8, "37"
            .Add "Codice Prodotto", 20, "20"
            .Add "Litri Prodotto", 5, "21"
'            .Add "Litri prodotto 1", 5, "20"
'            .Add "Litri prodotto 2", 5, "21"
'            .Add "Litri prodotto 3", 5, "22"
'            .Add "Litri prodotto 4", 5, "23"
'            .Add "Litri prodotto 5", 5, "24"
'            .Add "Litri prodotto 6", 5, "25"
'            .Add "Litri prodotto 7", 5, "26"
'            .Add "Litri prodotto 8", 5, "27"
'            .Add "Litri prodotto 9", 5, "28"
'            .Add "Litri prodotto 10", 5, "29"
            .Add "Ordine", 8, "30" '"24"
            .Add "Km parziali", 5, "31" '"25"
            .Add "Orario arrivo", 4, "32" '"26"
            .Add "Km raggio", 5, "33" '"27"
            .Add "Km raggio stradali", 5, "34" '"28"
            .Add "Destinazione", 10, "19"
            .Add "Filler 1", 6, "35" '"29"
            .Add "Filler 2", 1, "36" '"30"
        End With
    #ElseIf VersionID = 4 Then
        With mobjInFile
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
            .Add "Filler 1", 14, "13"
            .Add "Filler 2", 1, "15"
            .Add "Tipo Record D", 2, "16"
            .Add "Numero Viaggio D", 3, "17"
            .Add "Progressivo", 3, "18"
            .Add "Destinazione", 8, "19"
            .Add "Litri prodotto 1", 5, "20", , fftNumerico
            .Add "Litri prodotto 2", 5, "21", , fftNumerico
            .Add "Litri prodotto 3", 5, "22", , fftNumerico
            .Add "Litri prodotto 4", 5, "23", , fftNumerico
            .Add "Ordine", 7, "24"
            .Add "Km parziali", 5, "25"
            .Add "Orario arrivo", 4, "26"
            .Add "Pos. dog. prodotto 1", 3, "27"
            .Add "Pos. dog. prodotto 2", 3, "28"
            .Add "Pos. dog. prodotto 3", 3, "29"
            .Add "Pos. dog. prodotto 4", 3, "30"
            .Add "Filler 2", 1, "31"
        End With
    #End If
    TDBGrid1.Columns(11).Visible = False
    TDBGrid1.Columns(13).Visible = False
    
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
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set frmVVTImportMover = Nothing
End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0)
    cmdImport.Enabled = False
End Sub


