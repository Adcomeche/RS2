VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGP1Import 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE PROPOSTA DI FATTURA"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   9405
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3120
         Picture         =   "frmGP1Import.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmGP1Import.frx":0532
         TabIndex        =   7
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
         Picture         =   "frmGP1Import.frx":0634
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmGP1Import.frx":0B66
         TabIndex        =   6
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
         Picture         =   "frmGP1Import.frx":0C68
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmGP1Import.frx":0D6A
         TabIndex        =   5
         Tag             =   "BROWSE"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   9135
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmGP1Import.frx":0E6C
            TabIndex        =   23
            Top             =   240
            Width           =   8895
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   6375
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   2055
            Left            =   120
            OleObjectBlob   =   "frmGP1Import.frx":8BD2
            TabIndex        =   21
            Top             =   240
            Width           =   6135
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   6600
         TabIndex        =   13
         Top             =   960
         Width           =   2655
         Begin VB.Label lblLog 
            BackStyle       =   0  'Transparent
            Caption         =   "Righe importate correttamente:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            BackStyle       =   0  'Transparent
            Caption         =   "Righe da importare:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1335
         End
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   600
         TabIndex        =   1
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
      Begin MSMask.MaskEdBox mskRifEst 
         Height          =   255
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskAnno 
         Height          =   255
         Left            =   7680
         TabIndex        =   27
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1800
         Top             =   960
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
               Picture         =   "frmGP1Import.frx":B5DE
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1Import.frx":B6F0
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1Import.frx":B802
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Anno prefattura"
         Height          =   255
         Index           =   13
         Left            =   6240
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Rif. est. prefattura"
         Height          =   255
         Index           =   14
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&File"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   5880
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
         Picture         =   "frmGP1Import.frx":B914
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmGP1Import.frx":BA16
         TabIndex        =   8
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6660
      Width           =   9405
      _ExtentX        =   16589
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
            Object.Width           =   11139
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.44"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "06/10/2004"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGP1Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State As FormState
Public Loading As Boolean
Private GP1 As CPrefatturaSoc
Private mobjInFile As CTextFile
Private OCT As COrdineSelezione
Private ADE As CDeposito
Private AZO As CZona
Private AUT As CUnitaDiTrasporto
Private APR As CProdotto
Private ADZ As CDestinazione
Private XA As XArrayDB
Private XALog As XArrayDB

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

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
    Dim strIDOrderSoc As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDUT As String
    Dim strIDRiga  As String
    Dim strBadge As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strData As String
    Dim strErrorMessage As String
    Dim strRifEstPrefattura As String
    Dim intAnno As Integer
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim blnNewOrder As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    strRifEstPrefattura = mskRifEst
    intAnno = mskAnno
    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        blnNewOrder = True
        intRowError = NoError
        
        For Col = 0 To 15
            mobjInFile(Col + 1).value = XA(Row, Col)
        Next Col
        
        ' --- Controllo che l'ordine non sia già stato inserito
        strIDOrderSoc = mobjInFile("3").value
        strData = mobjInFile("4").value
        strIDProdotto = APR.DecodeFirmCode(ID_SOC, mobjInFile("10").value, True)
        If GP1.CheckSocCode(strIDOrderSoc, strIDProdotto) Then
            If Not GP1.Delete(strIDOrderSoc, strIDProdotto) Then
                intRowError = DeleteError
                GoTo LogResult
            End If
        End If
        ' ---
        
        ' --- Controllo che esistano Deposito, Zona, Prodotto, UdT, Destinazione
        strIDDeposito = ADE.DecodeFirmCode(ID_SOC, mobjInFile("2").value, True)
        strIDProdotto = APR.DecodeFirmCode(ID_SOC, mobjInFile("10").value, True)
'        strIDUT = AUT.DecodeFirmCode(ID_SOC, mobjInFile("13").value, True)
        strBadge = mobjInFile("13").value
        strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC, Format$(mobjInFile("11").value, "00000000"), True)
        If Len(strBadge) > 0 Then
            strIDUT = AUT.GetIDByBadge(strBadge)
        Else
            strIDUT = OCT.GetAUT(strIDOrderSoc, strIDProdotto, strIDDestinazione, strData)
            If Len(strIDUT) = 0 Then
                strIDUT = InputBox("Inserisci il codice unità di trasporto per la consegna: " & vbCr & _
                                    "Destinazione   " & mobjInFile("11").value & vbCr & _
                                    "               " & mobjInFile("12").value & vbCr & _
                                    "Rif. Ordine    " & mobjInFile("3").value & vbCr & _
                                    "Prodotto       " & strIDProdotto & vbCr & _
                                    "Quantità       " & mobjInFile("6").value & vbCr & _
                                    "Data           " & strData & vbCr & _
                                    "Bolletta       " & mobjInFile("5").value & vbCr, _
                                    "Unità di trasporto non trovata")
                
            End If
            
        End If
        strIDZona = mobjInFile("1").value
        AZO.DS.Fields("Codice").FindValue = strIDZona
        If Not AZO.Find Then
            intRowError = NoZona
            GoTo LogResult
        End If
        If AZO.DS.HowMany = 0 Then
            intRowError = NoZona
            GoTo LogResult
        End If
'        ' --- Controllo che la zona non sia da escludere
'        ' --- dall'importazione
'        If AZO.DS.Fields("2").value <> 0 Then
'            intRowError = ZonaEsclusaImport
'            GoTo LogResult
'        End If
'        ' ---
        
        If strIDDeposito = vbNullString Then
            intRowError = NoDeposito
            GoTo LogResult
        End If
        If strIDProdotto = vbNullString Then
            intRowError = NoProdotto
            GoTo LogResult
        End If
        If strIDUT = vbNullString Then
            strIDUT = InputBox("Inserisci il codice unità di trasporto per la consegna: " & vbCr & _
                                "Destinazione   " & mobjInFile("11").value & vbCr & _
                                "               " & mobjInFile("12").value & vbCr & _
                                "Rif. Ordine    " & mobjInFile("3").value & vbCr & _
                                "Prodotto       " & strIDProdotto & vbCr & _
                                "Quantità       " & mobjInFile("6").value & vbCr & _
                                "Data           " & strData & vbCr & _
                                "Bolletta       " & mobjInFile("5").value & vbCr, _
                                "Unità di trasporto non trovata")
            If strIDUT = vbNullString Then
                intRowError = noUT
                GoTo LogResult
            End If
        End If
'        If strIDDestinazione = vbNullString Then
'            intRowError = NoDestinazione
'            GoTo LogResult
'        End If
        ' ---
        
        ' --- Inserisco la riga di prefattura
        GP1.DS.DA.BeginTransaction
        RecordToGP1 GP1, mobjInFile, strIDDeposito, strIDDestinazione, _
                    strIDUT, strIDProdotto, strIDZona, strRifEstPrefattura, _
                    intAnno
        
        If Not GP1.Insert Then
            intRowError = GP1InsertError
            GoTo LogResult
        End If
        ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If intRowError = NoError Then
            GP1.DS.DA.EndTrans 0
        Else
            GP1.DS.DA.EndTrans 1
        End If
        
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            Select Case intRowError
                Case ZonaEsclusaImport
                    lngCountExcluded = lngCountExcluded + 1
                Case Else
                    lngCountFail = lngCountFail + 1
            End Select
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
'        lblLog(6) = lngCountExcluded
        
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case GP1InsertError
                strErrorMessage = "Errore durante l'inserimento"
            Case DeleteError
                strErrorMessage = "Impossibile cancellare la proposta di fattura"
            Case NoTipoVendita
                strErrorMessage = "Codice Tipo Vendita non trovato"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case noUT
                strErrorMessage = "Codice Unità di Trasporto non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione non trovato"
            Case OrderAlreadyExisting
                strErrorMessage = "Proposta di fattura già inserita"
        End Select
        XALog.value(Row, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .Bookmark = Row
        End With
        TDBGrid1.Bookmark = Row
        ' ---
        
        DoEvents
    Next Row
    ' ---
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    OCT.DS.DA.EndTrans 1
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
    Set C = ErrH.SetError(ErrCode, "frmGP1Import", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    OCT.DS.DA.EndTrans 0
    Me.mskRifEst = vbNullString
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file di prefattura... "
'                .InitDir = App.Path
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
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
    Dim lngCountOrders As Long
    Dim lngLogRows As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
'    For i = 0 To 3
'        lblField(i) = vbNullString
'    Next i
    cmdImport.Enabled = False
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    Do While Not mobjInFile.EOF
        varBuffer = mobjInFile.GetLine
'        If varBuffer(0) = "1" Then
'            blnValidHeader = True
'            mstrIDEnte = CStr(varBuffer(1))
'            strEnte = AEE.DecodeFirmCode(ID_SOC, mstrIDEnte)
'            If strEnte <> vbNullString Then
'                lblField(1) = strEnte
'                blnIDOk = True
'            End If
'            lblField(0) = varBuffer(2)
'            lblField(2) = varBuffer(3)
'            lblField(3) = varBuffer(4)
'        Else
        
        With XA
            .ReDim 1, lngRows, 0, 15
            For i = 0 To 15
                .value(.UpperBound(1), i) = varBuffer(i)
            Next i
        End With
        
        
        XALog.ReDim 1, lngRows, 0, 1
        lngLogRows = XALog.UpperBound(1)
        XALog.value(lngLogRows, 0) = varBuffer(2)
'        With OCT.DS
'            .Fields("27").FindValue = varBuffer(2)
'            OCT.Find
'            If .HowMany = 0 Then
'                XALog.value(lngLogRows, 1) = "Ordine non trovato"
'            Else
'                XALog.value(lngLogRows, 1) = .Fields("24").value
'            End If
'        End With
                
        lngRows = lngRows + 1
        DoEvents
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
    lblLog(5) = lngRows - 1
'            lblLog(6) = 0
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmGP1Import", _
                        "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmGP1Import", _
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
    mskFileName = gmRegistry.GetSetting("Options", "GP1ImportFile", , True)
    
    Set OCT = New COrdineSelezione
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set GP1 = New CPrefatturaSoc
    With GP1.DS
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
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
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
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ErrH = New CErrorHandler
    
    Set mobjInFile = New CTextFile
    With mobjInFile
        .Add "Zona", 2, "1"
        .Add "Deposito", 2, "2"
        .Add "Numero Ordine", 8, "3"
        .Add "Data Consegna", 8, "4"
        .Add "Numero Bolletta", 15, "5"
        .Add "Quantità", 6, "6"
        .Add "Riferimento", 8, "7"
        .Add "Tariffa", 9, "8"
        .Add "Importo", 12, "9"
        .Add "Prodotto", 6, "10"
        .Add "Codice Destinazione", 8, "11"
        .Add "Ragione Sociale", 23, "12"
        .Add "Unità di trasporto", 4, "13"
        .Add "Filler", 2, "14"
        .Add "Km raggio", 3, "15"
        .Add "Importo Forzato", 1, "16"
    End With
    Loading = False
End Sub

Private Sub RecordToGP1(GP1 As CPrefatturaSoc, InFile As CTextFile, _
                        IDDeposito As String, IDDestinazione As String, _
                        IDUT As String, IDProdotto As String, _
                        IDZona As String, RifEsterno As String, _
                        Anno As Integer)
    Dim strTariffa As String
    Dim strImporto As String
    
    With GP1.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDZona
        .Item("0").Updated = True
        .Item("1").NewValue = IDDeposito
        .Item("1").Updated = True
        .Item("2").NewValue = Trim$(Left$(InFile("3").value, .Item("2").MaxLength))
        .Item("2").Updated = True
        .Item("3").NewValue = Format$(InFile("4").value, "dd/mm/yyyy")
        .Item("3").Updated = True
        .Item("4").NewValue = Trim$(Left$(InFile("5").value, .Item("4").MaxLength))
        .Item("4").Updated = True
        .Item("5").NewValue = Trim$(InFile("6").value)
        .Item("5").Updated = True
        .Item("6").NewValue = Trim$(Left$(InFile("7").value, .Item("6").MaxLength))
        .Item("6").Updated = True
        strTariffa = Trim$(InFile("8").value)
        'Mid$(strTariffa, Len(strTariffa) - 4, 1) = GP1.DS.DA.DecSeparator
        If Len(strTariffa) > 0 Then
            If (InStr(1, strTariffa, ".") <> 0) Then
                Mid(strTariffa, InStr(1, strTariffa, "."), 1) = GP1.DS.DA.DecSeparator
            End If
        End If
        strImporto = Trim$(InFile("9").value)
        If Len(strImporto) > 0 Then
            If InStr(1, strImporto, ".") <> 0 Then
                Mid(strImporto, InStr(1, strImporto, "."), 1) = GP1.DS.DA.DecSeparator
    '            Mid$(strImporto, Len(strImporto) - 2, 1) = GP1.DS.DA.DecSeparator
            End If
        End If
        .Item("7").NewValue = strTariffa
        .Item("7").Updated = True
        .Item("8").NewValue = strImporto
        .Item("8").Updated = True
        .Item("9").NewValue = Trim$(InFile("10").value)
        .Item("9").Updated = True
        .Item("10").NewValue = IDDestinazione 'Trim$(InFile("11").Value)
        .Item("10").Updated = True
        .Item("11").NewValue = Trim$(Left$(InFile("12").value, .Item("11").MaxLength))
        .Item("11").Updated = True
        .Item("12").NewValue = IDUT
        .Item("12").Updated = True
'        .Item("13").NewValue = Trim$(Left$(InFile("14").value, .Item("14").MaxLength))
'        .Item("13").Updated = True
        .Item("13").NewValue = InFile("15").value
        .Item("13").Updated = True
        .Item("14").NewValue = (Len(InFile("16").value) > 0)
        .Item("14").Updated = True
        .Item("15").NewValue = RifEsterno
        .Item("15").Updated = True
        .Item("16").NewValue = Anno
        .Item("16").Updated = True
        .Item("18").NewValue = ID_SOC_AGIP
        .Item("18").Updated = True
    End With

End Sub



Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "GP1ImportFile", mskFileName, True
    Set OCT = Nothing
    Set XA = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set frmGP1Import = Nothing
End Sub

Private Sub mskAnno_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0) And (Len(mskRifEst) > 0) And (Len(mskAnno) > 0)

End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0) And (Len(mskRifEst) > 0) And (Len(mskAnno) > 0)
    cmdImport.Enabled = False
End Sub

Private Sub mskRifEst_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0) And (Len(mskRifEst) > 0) And (Len(mskAnno) > 0)

End Sub

