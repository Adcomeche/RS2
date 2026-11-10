VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmGP1ImportCEPSA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE PREFATTURE CEPSA"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   9405
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   9
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
         Picture         =   "frmGP1ImportCEPSA.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmGP1ImportCEPSA.frx":0532
         TabIndex        =   4
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
         Picture         =   "frmGP1ImportCEPSA.frx":0634
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmGP1ImportCEPSA.frx":0B66
         TabIndex        =   3
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
         Picture         =   "frmGP1ImportCEPSA.frx":0C68
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmGP1ImportCEPSA.frx":0D6A
         TabIndex        =   2
         Tag             =   "BROWSE"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   9135
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmGP1ImportCEPSA.frx":0E6C
            TabIndex        =   20
            Top             =   240
            Width           =   8895
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   6375
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   2055
            Left            =   120
            OleObjectBlob   =   "frmGP1ImportCEPSA.frx":D346
            TabIndex        =   18
            Top             =   240
            Width           =   6135
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   6600
         TabIndex        =   10
         Top             =   960
         Width           =   2655
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            BackStyle       =   0  'Transparent
            Caption         =   "Righe importate correttamente:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
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
            TabIndex        =   15
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLog 
            BackStyle       =   0  'Transparent
            Caption         =   "Righe da importare:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   12
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
            TabIndex        =   11
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
               Picture         =   "frmGP1ImportCEPSA.frx":FD52
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1ImportCEPSA.frx":FE64
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1ImportCEPSA.frx":FF76
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
         TabIndex        =   0
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   6
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
         Picture         =   "frmGP1ImportCEPSA.frx":10088
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
         Left            =   7560
         Picture         =   "frmGP1ImportCEPSA.frx":1018A
         TabIndex        =   5
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
      Top             =   6615
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
            TextSave        =   "13:50"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/12/2017"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGP1ImportCEPSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State As FormState
Public Loading As Boolean
Private VVT  As CViaggio
Private GP1 As CPrefatturaSoc
Private VVP As CViaggioProdotto
Private mobjInFile As CTextFile
Private OCT As COrdine
Private ADE As CDeposito
Private AZO As CZona
Private AAM As CAutomezzo
Private AAU As CAutista
Private AUT As CUnitaDiTrasporto
Private APR As CProdotto
Private ADZ As CDestinazione
Private XA As XArrayDB
Private XALog As XArrayDB
Private mstrFileName As String
Private mstrRiferimentoFattura As String
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private Sub ImportRecords()
    On Error GoTo ErrHandler
    
    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strRaSocDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDOrder As String
    Dim strRifOrderSoc As String
    Dim strTipoConsegna As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDUT As String
    Dim strIDUTSoc As String
    Dim strIDUTDeposito As String
    Dim strIDVVT As String
    Dim strIDVVC As String
    Dim strIDVVP As String
    Dim strIDAUM As String
    Dim strIDOCT As String
    Dim strIDRiga  As String
    Dim strBadge As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strIDAutista As String
    Dim strIDMotrice As String
    Dim strIDRimorchio As String
    Dim strData As String
    Dim strErrorMessage As String
    Dim strRifEstPrefattura As String
    Dim strTipoRiga As String
    Dim strFile As String
    Dim sngQuantita As Single
    Dim sngQuantitaKG As Single
    Dim intAnnoFile As Integer
    Dim intMeseFile As Integer
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim blnNewOrder As Boolean
    Dim blnConcErr As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
'    strRifEstPrefattura = mskRifEst
    strFile = Dir(mstrFileName)
    Select Case UCase(Left(strFile, 5))
        Case "CPREF"
            strTipoRiga = "P"
        Case "CFACT"
            strTipoRiga = "F"
        Case "CDIFE"
            strTipoRiga = "D"
    End Select
    intAnnoFile = CInt(Mid(strFile, 12, 4))
    intMeseFile = CInt(Mid(strFile, 16, 2))
    Select Case XA(1, 0)
        Case "002881"
            strIDEnteEmittente = "004"
        Case "054320"
            strIDEnteEmittente = "624"
        Case "709704"
            strIDEnteEmittente = "004"
        Case Else
            intRowError = NoEnteEmittente
            GoTo ErrHandler
    End Select
    If Not GP1.DeletePrefCEPSA(intAnnoFile, intMeseFile, strIDEnteEmittente) Then
        GoTo ErrHandler
    End If
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
        
        For Col = 0 To 33
            mobjInFile(Col + 1).value = XA(Row, Col)
        Next Col
        
        ' --- Controllo che esistano Deposito, Zona, Prodotto, UdT, Destinazione
        strIDDeposito = ADE.DecodeFirmCode(ID_SOC_CEPSA, mobjInFile("6").value, True)
        
        
        ' -- DP 23/01/2009
'        If strIDDeposito = "022" And strIDEnteEmittente = "624" Then
'            strIDDeposito = "023"
'        End If
'        If strIDDeposito = "023" And strIDEnteEmittente = "004" Then
'            strIDDeposito = "022"
'        End If
'        If strIDDeposito = "009" And strIDEnteEmittente = "624" Then
'            strIDDeposito = "008"
'        End If
'        If strIDDeposito = "008" And strIDEnteEmittente = "004" Then
'            strIDDeposito = "009"
'        End If
'        If strIDDeposito = "001" And strIDEnteEmittente = "624" Then
'            strIDDeposito = "003"
'        End If
'        If strIDDeposito = "003" And strIDEnteEmittente = "004" Then
'            strIDDeposito = "001"
'        End If
        
        
        strIDUTSoc = mobjInFile("7").value
        If Len(strIDUTSoc) > 0 Then
            If Left(strIDUTSoc, 1) = "D" Then
                strIDUTSoc = Mid(strIDUTSoc, 2)
            End If
            strIDUTSoc = CStr(CLng(strIDUTSoc))
        End If
        strIDUT = AUT.DecodeFirmCode(ID_SOC_CEPSA, strIDUTSoc, True)
        strIDUTDeposito = AUT.DecodeFirmCode(ID_SOC_CLH, mobjInFile("8").value, True)
        strRifOrderSoc = mobjInFile("16").value
        
        Select Case UCase(Left(mobjInFile("9").value, 2))
            Case "ES"
                strIDMotrice = Mid(mobjInFile("9").value, 3)
            Case Else
                If UCase(Left(mobjInFile("9").value, 1)) = "E" Then
                    strIDMotrice = Mid(mobjInFile("9").value, 2)
                Else
                    strIDMotrice = mobjInFile("9").value
                End If
        End Select
        strIDMotrice = AAM.GetIDByExtRef(strIDMotrice)
        
        Select Case UCase(Left(mobjInFile("10").value, 2))
            Case "ES"
                strIDRimorchio = Mid(mobjInFile("10").value, 3)
            Case Else
                If UCase(Left(mobjInFile("10").value, 1)) = "E" Then
                    strIDRimorchio = Mid(mobjInFile("10").value, 2)
                Else
                    strIDRimorchio = mobjInFile("10").value
                End If
        End Select
        strIDRimorchio = AAM.GetIDByExtRef(strIDRimorchio)
        
        strIDAutista = AAU.DecodeFirmCode(ID_SOC_CEPSA, mobjInFile("11").value, True)
        strIDProdotto = APR.DecodeFirmCode(ID_SOC_CEPSA, Left(mobjInFile("19").value, 5), True)
        strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC_CLH, "CP" & Right(mobjInFile("20").value, 8), True)
        strRaSocDestinazione = ADZ.DecodeFirmCode(ID_SOC_CLH, "CP" & Right(mobjInFile("20").value, 8), False)
        ' ---
        
        strTipoConsegna = mobjInFile("15").value
        
        ' --- Inserisco la riga di prefattura
        GP1.DS.DA.BeginTransaction
        RecordToGP1 GP1, mobjInFile, strIDDeposito, strIDDestinazione, _
                    strIDUT, strIDProdotto, strRaSocDestinazione, strRifEstPrefattura, _
                    intAnnoFile, strTipoRiga, intMeseFile, strIDMotrice, strIDRimorchio, strIDAutista, mstrRiferimentoFattura, _
                    strIDEnteEmittente
        
        If Not GP1.Insert Then
            intRowError = GP1InsertError
            GoTo LogResult
        End If

        
        
        If Not GP1.GetConsegnaCEPSA(strRifOrderSoc, strTipoConsegna, strIDVVT, strIDVVC, strIDVVP, strIDAUM, strIDOCT) Then
            intRowError = BadIDOrderSoc
            GoTo LogResult
        End If
'        If strIDAUM = "K" Then
'            sngQuantita = mobjInFile("23").value
'        Else
            sngQuantita = mobjInFile("22").value
            sngQuantitaKG = mobjInFile("23").value
'        End If

        
        ' Completo el viaje si no está completado y está en estado Definitivo o Previsorio
        
        VVT.DS.Fields("0").FindValue = strIDVVT
        VVT.Find
        If VVT.DS.Fields("42").value = "D" Or VVT.DS.Fields("42").value = "P" Then
            If VVT.SetStato("C") Then
            End If
        
        End If
        
        If Not VVP.UpdateQtyDelivered(strIDVVT, strIDVVC, strIDVVP, sngQuantita, sngQuantitaKG, True) Then
            intRowError = DeliveryUpdateError
            GoTo LogResult
        End If
                
        Select Case strTipoConsegna
            Case "910151", "910152", "910153"
                With OCT.DS
                    .Fields.ResetFields
                    .Field("Numero Ordine").FindValue = strIDOCT
                    OCT.Find
                    .Field("Rif. Esterno 1").NewValue = strRifOrderSoc
                    .Field("Rif. Esterno 1").Updated = True
                    OCT.BeginSave blnConcErr
                    OCT.DS.DA.EndTrans ETCOMMIT
                End With
        End Select
                
        ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError = NoError) Or (intRowError = DeliveryUpdateError) Or (intRowError = BadIDOrderSoc) Then
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
            Case DeliveryUpdateError
                strErrorMessage = "Errore nell'aggiornamento della quantità della consegna"
            Case BadIDOrderSoc
                strErrorMessage = "Consegna non trovata"
        End Select
        XALog.value(Row, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = Row
        End With
        TDBGrid1.bookmark = Row
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
        Case NoEnteEmittente
            strErrorMessage = "Vettore non trovato"
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
'    Me.mskRifEst = vbNullString
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
    mstrFileName = mskFileName
    
    If Not mobjInFile.OpenFile(mstrFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
'    For i = 0 To 3
'        lblField(i) = vbNullString
'    Next i
    mstrRiferimentoFattura = Left(Dir(mstrFileName), InStr(1, Dir(mstrFileName), ".") - 1)
    cmdImport.Enabled = False
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    Do While Not mobjInFile.EOF
        varBuffer = mobjInFile.GetLine(False)
        If Not mobjInFile.ParseFixedLenFields(CStr(varBuffer), 1, 34) Then
            GoTo ErrHandler
        End If
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
            .ReDim 1, lngRows, 0, 33
            For i = 1 To 34
                .value(.UpperBound(1), i - 1) = mobjInFile.Item(i).value
            Next i
        End With
        
        
        XALog.ReDim 1, lngRows, 0, 1
        lngLogRows = XALog.UpperBound(1)
        XALog.value(lngLogRows, 0) = mobjInFile.Item(16).value
                
        lngRows = lngRows + 1
        DoEvents
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
    lblLog(5) = lngRows - 1
'            lblLog(6) = 0
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmGP1Import", _
                        "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & mstrFileName)
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
    mskFileName = gmRegistry.GetSetting("Options", "GP1ImportFileCEPSA", , True)
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set GP1 = New CPrefatturaSoc
    With GP1.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
   
    Set VVP = New CViaggioProdotto
    With VVP.DS
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
    
    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
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
        .Add "Vettore", 6, "1", , fftNumerico
        .Add "Data", 8, "2", , fftNumerico
        .Add "Tipo di servizio", 6, "3"
        .Add "Tipo consegna", 2, "4", " 0"
        .Add "Stato record", 1, "5"
        .Add "Deposito", 4, "6", , fftNumerico
        .Add "Riferimento UDT", 12, "7"
        .Add "Riferimento UDT Deposito", 3, "8"
        .Add "Targa Motrice", 15, "9"
        .Add "Targa Rimorchio", 15, "10"
        .Add "Riferimento autista", 7, "11"
        .Add "NIF Autista", 14, "12"
        .Add "Numero viaggio", 5, "13", , fftNumerico
        .Add "Codice operatore", 6, "14"
        .Add "Codice fornitore", 6, "15"
        .Add "Riferimento ordine", 16, "16"
        .Add "Riferimento consegna SAP", 12, "17", , fftNumerico
        .Add "Riferimento ticket", 15, "18"
        .Add "Codice Prodotto", 9, "19", , fftNumerico
        .Add "Codice Destinazione", 15, "20"
        .Add "Nome imbarcazione", 30, "21"
        .Add "Quantità Litri", 7, "22", , fftNumerico
        .Add "Quantità Kili", 7, "23", , fftNumerico
        .Add "Distanza Km", 4, "24", , fftNumerico
        .Add "Tipo tariffa", 2, "25"
        .Add "Tariffa", 9, "26", , fftNumerico
        .Add "Importo", 10, "27", , fftNumerico
        .Add "Importo pedaggio", 10, "28", , fftNumerico
        .Add "Servicios Exentos", 9, "34", 0, fftNumerico   '-- DP 20/08/08
        .Add "Filler", 54, "29"
        .Add "Servizio operario bunker", 10, "30", 0, fftNumerico
        .Add "Servizio buques a puerto", 10, "31", 0, fftNumerico
        .Add "Ore servizi speciali", 5, "32", 0, fftNumerico
        .Add "Importo servizi speciali", 10, "33", 0, fftNumerico
    End With
    Loading = False
End Sub

Private Sub RecordToGP1(GP1 As CPrefatturaSoc, InFile As CTextFile, _
                        IDDeposito As String, IDDestinazione As String, _
                        IDUT As String, IDProdotto As String, _
                        RagioneSocialeDestinazione As String, RifEsterno As String, _
                        Anno As Integer, TipoRiga As String, Mese As Integer, _
                        IDMotrice, IDRimorchio, IDAutista, RiferimentoFattura As String, _
                        EnteEmittente As String)
    Dim sngTariffa As Single
    Dim sngImporto As Single
    Dim sngPedaggio As Single
    Dim sngOperario As Single
    Dim sngBuques As Single
    Dim sngOreServiziSpeciali As Single
    Dim sngServiziSpeciali As Single
    Dim sngServiziEsenti As Single
    Dim strRifOrdine As String
    
    strRifOrdine = InFile("16").value
        
    With GP1.DS.Fields
        .ResetFields
        .Item("1").NewValue = IDDeposito
        .Item("1").Updated = True
        .Item("2").NewValue = Trim$(Left$(strRifOrdine, .Item("2").MaxLength))
        .Item("2").Updated = True
        .Item("3").NewValue = Right(InFile("2").value, 2) & "/" & Mid(InFile("2").value, 5, 2) & "/" & Left(InFile("2").value, 4)
        .Item("3").Updated = True
        .Item("5").NewValue = Trim$(InFile("22").value)
        .Item("5").Updated = True
        
        sngTariffa = CLng(InFile("26").value) / 1000000
        sngImporto = CLng(InFile("27").value) / 100
        sngPedaggio = CLng(InFile("28").value) / 100
        sngOperario = CLng(InFile("30").value) / 100
        sngBuques = CLng(InFile("31").value) / 100
        If IsNumeric((InFile("32").value)) Then
        sngOreServiziSpeciali = CLng(InFile("32").value) / 100
        Else
        sngOreServiziSpeciali = 0
        End If
        sngServiziSpeciali = CLng(InFile("33").value) / 100
        
        '-- DP 20/08/08
         If IsNumeric((InFile("34").value)) Then
        sngServiziEsenti = CLng(InFile("34").value) / 1000
        Else
        sngServiziEsenti = 0
        End If
        .Item("7").NewValue = sngTariffa
        .Item("7").Updated = True
        .Item("8").NewValue = sngImporto
        .Item("8").Updated = True
        .Item("9").NewValue = IDProdotto
        .Item("9").Updated = True
        .Item("10").NewValue = IDDestinazione
        .Item("10").Updated = True
        .Item("11").NewValue = Trim$(Left$(RagioneSocialeDestinazione, .Item("11").MaxLength))
        .Item("11").Updated = True
        .Item("12").NewValue = IDUT
        .Item("12").Updated = True
        .Item("13").NewValue = InFile("24").value
        .Item("13").Updated = True
        .Item("14").NewValue = 0
        .Item("14").Updated = True
        .Item("16").NewValue = Anno
        .Item("16").Updated = True
        .Item("18").NewValue = ID_SOC_CEPSA
        .Item("18").Updated = True
        .Item("19").NewValue = InFile("1").value
        .Item("19").Updated = True
        
        Select Case UCase(InFile("3").value)
            Case "DESVIO"
                .Item("20").NewValue = "D"
            Case Else
                .Item("20").NewValue = "N"
        End Select
        .Item("20").Updated = True
        .Item("21").NewValue = InFile("6").value
        .Item("21").Updated = True
        .Item("22").NewValue = InFile("7").value
        .Item("22").Updated = True
        .Item("23").NewValue = InFile("8").value
        .Item("23").Updated = True
        .Item("24").NewValue = InFile("9").value
        .Item("24").Updated = True
        .Item("25").NewValue = InFile("10").value
        .Item("25").Updated = True
        .Item("26").NewValue = InFile("11").value
        .Item("26").Updated = True
        .Item("27").NewValue = InFile("12").value
        .Item("27").Updated = True
        .Item("28").NewValue = InFile("13").value
        .Item("28").Updated = True
        .Item("29").NewValue = InFile("14").value
        .Item("29").Updated = True
        .Item("30").NewValue = InFile("15").value
        .Item("30").Updated = True
        .Item("31").NewValue = strRifOrdine
        .Item("31").Updated = True
        .Item("32").NewValue = InFile("17").value
        .Item("32").Updated = True
        .Item("33").NewValue = InFile("18").value
        .Item("33").Updated = True
        .Item("34").NewValue = InFile("21").value
        .Item("34").Updated = True
        .Item("35").NewValue = InFile("23").value
        .Item("35").Updated = True
        Select Case UCase(InFile("25").value)
            Case "BT"
                .Item("36").NewValue = "B"
        End Select
        .Item("36").Updated = True
        .Item("37").NewValue = sngPedaggio
        .Item("37").Updated = True
        .Item("38").NewValue = TipoRiga
        .Item("38").Updated = True
        .Item("39").NewValue = IDMotrice
        .Item("39").Updated = True
        .Item("40").NewValue = IDRimorchio
        .Item("40").Updated = True
        .Item("41").NewValue = IDAutista
        .Item("41").Updated = True
        .Item("42").NewValue = RiferimentoFattura
        .Item("42").Updated = True
        .Item("43").NewValue = sngOperario
        .Item("43").Updated = True
        .Item("44").NewValue = sngBuques
        .Item("44").Updated = True
        .Item("45").NewValue = sngOreServiziSpeciali
        .Item("45").Updated = True
        .Item("46").NewValue = sngServiziSpeciali
        .Item("46").Updated = True
        .Item("47").NewValue = EnteEmittente
        .Item("47").Updated = True
        
        '-- DP 20/08/08
        .Item("48").NewValue = sngServiziEsenti
        .Item("48").Updated = True
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "GP1ImportFileCEPSA", mskFileName, True
    Set OCT = Nothing
    Set AAU = Nothing
    Set AAM = Nothing
    Set XA = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set frmGP1ImportCEPSA = Nothing
End Sub

Private Sub mskFileName_Change()
    cmdReadFile.Enabled = (Len(mskFileName) > 0) 'And (Len(mskRifEst) > 0)
    cmdImport.Enabled = False
End Sub

'Private Sub mskRifEst_Change()
'    cmdReadFile.Enabled = (Len(mskFileName) > 0) And (Len(mskRifEst) > 0)
'
'End Sub

