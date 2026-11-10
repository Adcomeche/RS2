VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmKNOSOSViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRASMISSIONE DATI CONSEGNE"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11430
   Begin VB.CheckBox chkActiveOnly 
      Caption         =   "Solo viaggi non annullati"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   5220
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   4740
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4140
      Width           =   3555
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   5880
      Width           =   11415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   3360
         Picture         =   "frmKNOSOSViewer.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   435
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
         Left            =   9840
         Picture         =   "frmKNOSOSViewer.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9720
         Picture         =   "frmKNOSOSViewer.frx":0204
         TabIndex        =   5
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   8280
         Picture         =   "frmKNOSOSViewer.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8160
         Picture         =   "frmKNOSOSViewer.frx":0408
         TabIndex        =   4
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   6600
         Picture         =   "frmKNOSOSViewer.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         Picture         =   "frmKNOSOSViewer.frx":060C
         TabIndex        =   3
         Tag             =   "BROWSE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   5040
         Picture         =   "frmKNOSOSViewer.frx":070E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton CmdOpenTrip 
         Caption         =   "Apri &Viaggio"
         Height          =   375
         Left            =   4920
         Picture         =   "frmKNOSOSViewer.frx":0810
         TabIndex        =   2
         Tag             =   "OPEN"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Modifica"
         Height          =   375
         Left            =   3240
         Picture         =   "frmKNOSOSViewer.frx":0912
         TabIndex        =   14
         Tag             =   "PRINT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1800
         Picture         =   "frmKNOSOSViewer.frx":0A14
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - &Annulla"
         Height          =   375
         Left            =   1680
         Picture         =   "frmKNOSOSViewer.frx":0B16
         TabIndex        =   1
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmKNOSOSViewer.frx":0C18
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - &Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmKNOSOSViewer.frx":0D1A
         TabIndex        =   0
         Tag             =   "FIND"
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6750
      Width           =   11430
      _ExtentX        =   20161
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
            Object.Width           =   14711
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.55"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12/05/2005"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid2 
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmKNOSOSViewer.frx":0E1C
      TabIndex        =   16
      Top             =   2160
      Width           =   4095
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid3 
      Height          =   1815
      Left            =   4380
      OleObjectBlob   =   "frmKNOSOSViewer.frx":1309E
      TabIndex        =   17
      Top             =   2160
      Width           =   4155
   End
   Begin MSMask.MaskEdBox mskMaxRecords 
      Height          =   300
      Left            =   2160
      TabIndex        =   20
      Top             =   4860
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   "_"
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmKNOSOSViewer.frx":25320
      TabIndex        =   23
      Top             =   180
      Width           =   10815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "viaggi"
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   4860
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Restituisci al massimo"
      Height          =   255
      Index           =   16
      Left            =   480
      TabIndex        =   21
      Top             =   4860
      Width           =   1575
   End
End
Attribute VB_Name = "frmKNOSOSViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DB 21/08/2003 Rilascio
Option Explicit

'Const DISPONIBILE = "S"
'Const MOTRICE = "M"
'Const RIMORCHIO = "R"

Private XA As XArrayDB
Private XALog As XArrayDB

Public WithEvents DataClass As CViaggioSelezioneSatellitare
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
'Private mintCurFrame As Integer
'Private mvarFunction As FunctionCodes
'Private lngViaggiModificati As Long
'Private mstrUdTAppoggio As String
Private AXX As CParametro
'Private APR As CProdotto
'Private AUT As CUnitaDiTrasporto
'Private AAM As CAutomezzo
'Private AAU As CAutista
'Private OCT As COrdine
'Private VVT As CViaggio
'Private VVC As CViaggioConsegna
'Private VVP As CViaggioProdotto
'Private VPI As CViaggioEvento
'Private AEE As CEnteEmittente
'Private ADE As CDeposito
'Private ADZ As CDestinazione
'Private mobjOutFile As CTextFile
'Private mobjOutFileInc As CTextFile

Private strOraInCaricoValore As String
Private strUdTValoreValore As String
Private strMotriceValore As String
Private strRimorchioValore As String
Private strAutista1Valore As String
Private strAutista2Valore As String
Private strPedaggioValore As String
Private strOraConsegnaValore As String
Private strProgressivoValore As String
Private strKmValore As String
Private strNoteValore As String
Private strNoteEsteseValore As String
Private strFineCaricoValore As String
Private strQtaCaricataValore As String
Private strQtaScaricataValore As String
Private strDataConsegna As String
Private strDataViaggio As String

'Private Enum AzioneRichiesta
'    arModificaViaggio = 1
'    arAnnullaModifica = 2
'    arLetturaViaggi = 3
'End Enum
'
'Private mvarTripToUpd(0 To 3, 0 To 0) As Variant
'
'Private Enum SincronizzazioneRichiesta
'    srLetturaDati = 1
'    srScritturaDati = 2
'End Enum

'Private Sub ExportEventi()
'    On Error GoTo ErrHandler
'
'    Const TEMP_FILE = "INC.tmp"
'    Dim intTempFileNum As Integer
'    Dim intRowError As Integer
'    Dim intFileFromNum As Integer
'    Dim intFileToNum As Integer
'    Dim strLine As String
'    Dim strTipoEvento As String
'    Dim strDataInvio As String
'    Dim strOraInvio As String
'    Dim strProgrInvio As String
'    Dim strFileName As String
'    Dim strIDViaggio As String
'    Dim strIDViaggioPrec As String
'    Dim strIDDestinazione As String
'    Dim strIDDestinazioneDirottamento As String
'    Dim strIDDepositoRientro As String
'    Dim strIDProdotto As String
'    Dim strIDProdottoContaminato As String
'    Dim varCounter As Variant
'    Dim varBufferEventi As Variant
'    Dim i As Long
'    Dim lngRow As Long
'    Dim lngCountOK As Long
'    Dim lngNumRows As Long
'    Dim lngNumRowsEventi As Long
'    Dim ErrCode As Errors
'
'    If Not mobjOutFileInc.OpenFile(TEMP_FILE, 1) Then
'        ErrCode = FileNotFound
'        GoTo ErrHandler
'    End If
'
'    strDataInvio = Format$(Date, "YYYYMMDD")
'    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
'
'    lngNumRows = XA.UpperBound(1)
'    For lngRow = 0 To lngNumRows
'
'        strIDViaggio = XA(lngRow, 5)
'        If strIDViaggio <> strIDViaggioPrec Then
'            strIDViaggioPrec = strIDViaggio
'            If VPI.GetSelectionExport(strIDViaggio) Then
'                varBufferEventi = VPI.DS.List
'            End If
'
'            If Not IsEmpty(varBufferEventi) Then
'
'                lngNumRowsEventi = UBound(varBufferEventi, 2)
'                If Len(mskFileName) > 0 Then
'                    strFileName = mskFileName
'                Else
'                    ErrCode = FileNotFound
'                    GoTo ErrHandler
'                End If
'
'
'
'                For i = 0 To lngNumRowsEventi
'                    strTipoEvento = varBufferEventi(0, i)
'
'                    strIDDestinazione = vbNullString
'                    With ADZ.DS
'                        .Fields.ResetFields
'                        .Fields("0").FindValue = varBufferEventi(34, i)
'                        If (Not ADZ.Find) Or (.HowMany = 0) Then
'                            intRowError = NoDestinazione
'                            GoTo ErrHandler
'                        End If
'                        strIDDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
'                        strIDDestinazione = Mid(strIDDestinazione, 3)
'                    End With
'
'                    strIDProdotto = vbNullString
'                    With APR.DS
'                        .Fields.ResetFields
'                        .Fields("0").FindValue = varBufferEventi(7, i)
'                        If (Not APR.Find) Or (.HowMany = 0) Then
'                            intRowError = NoProdotto
'                            GoTo ErrHandler
'                        End If
'                        strIDProdotto = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
'                    End With
'
'                    With mobjOutFileInc
'                        .ClearFields
'            '            .Item("2").value = strProgrInvio
'                        .Item("3").value = strDataInvio
'                        .Item("4").value = strOraInvio
'            '            .Item("5").value = lngNumRowsEventi
'                        .Item("6").value = varBufferEventi(32, i) & varBufferEventi(33, i)
'                        .Item("7").value = varBufferEventi(41, i)
'                        .Item("8").value = strIDDestinazione
'                        .Item("9").value = strTipoEvento
'                        .Item("10").value = varBufferEventi(1, i)
'                        .Item("11").value = Format$(varBufferEventi(2, i), "YYYYMMDD")
'                        .Item("12").value = Format$(Hour(varBufferEventi(3, i)), "00") & ":" & Format$(Minute(varBufferEventi(3, i)), "00")
'                        .Item("13").value = Format$(Hour(varBufferEventi(4, i)), "00") & ":" & Format$(Minute(varBufferEventi(4, i)), "00")
'                        If Len(varBufferEventi(35, i)) > 0 Then
'                            .Item("14").value = Left(varBufferEventi(35, i), 13)
'                            .Item("15").value = Right(varBufferEventi(35, i), 2)
'                            .Item("16").value = varBufferEventi(36, i)
'                        End If
'                        .Item("50").value = Format$(varBufferEventi(29, i), "YYYYMMDD")
'                        .Item("51").value = varBufferEventi(30, i)
'                        .Item("52").value = varBufferEventi(31, i)
'                        Select Case strTipoEvento
'                            Case "007", "008"
'                                .Item("17").value = strIDProdotto
'                                .Item("18").value = varBufferEventi(8, i)
'                                .Item("19").value = Left(varBufferEventi(9, i), 1)
'                                strIDProdottoContaminato = vbNullString
'                                With APR.DS
'                                    .Fields.ResetFields
'                                    .Fields("0").FindValue = varBufferEventi(7, i)
'                                    If (Not APR.Find) Or (.HowMany = 0) Then
'                                        intRowError = NoProdotto
'                                        GoTo ErrHandler
'                                    End If
'                                    strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
'                                End With
'                                .Item("20").value = strIDProdottoContaminato
'                                .Item("21").value = varBufferEventi(11, i)
'                                .Item("22").value = Left(varBufferEventi(12, i), 1)
'                                If Len(varBufferEventi(13, i)) > 0 Then
'                                    strIDProdottoContaminato = vbNullString
'                                    With APR.DS
'                                        .Fields.ResetFields
'                                        .Fields("0").FindValue = varBufferEventi(13, i)
'                                        If (Not APR.Find) Or (.HowMany = 0) Then
'                                            intRowError = NoProdotto
'                                            GoTo ErrHandler
'                                        End If
'                                        strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
'                                    End With
'                                    .Item("23").value = strIDProdottoContaminato
'                                    .Item("24").value = varBufferEventi(14, i)
'                                    .Item("25").value = Left(varBufferEventi(15, i), 1)
'                                End If
'                                If Len(varBufferEventi(16, i)) > 0 Then
'                                    strIDProdottoContaminato = vbNullString
'                                    With APR.DS
'                                        .Fields.ResetFields
'                                        .Fields("0").FindValue = varBufferEventi(16, i)
'                                        If (Not APR.Find) Or (.HowMany = 0) Then
'                                            intRowError = NoProdotto
'                                            GoTo ErrHandler
'                                        End If
'                                        strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
'                                    End With
'                                    .Item("26").value = strIDProdottoContaminato
'                                    .Item("27").value = varBufferEventi(17, i)
'                                    .Item("28").value = Left(varBufferEventi(18, i), 1)
'                                End If
'                                If Len(varBufferEventi(19, i)) > 0 Then
'                                    strIDProdottoContaminato = vbNullString
'                                    With APR.DS
'                                        .Fields.ResetFields
'                                        .Fields("0").FindValue = varBufferEventi(19, i)
'                                        If (Not APR.Find) Or (.HowMany = 0) Then
'                                            intRowError = NoProdotto
'                                            GoTo ErrHandler
'                                        End If
'                                        strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
'                                    End With
'                                    .Item("29").value = strIDProdottoContaminato
'                                    .Item("30").value = varBufferEventi(20, i)
'                                    .Item("31").value = Left(varBufferEventi(21, i), 1)
'                                End If
'                            Case "009", "010"
'                                .Item("32").value = strIDProdotto
'                                .Item("33").value = varBufferEventi(8, i)
'                                .Item("34").value = Left(varBufferEventi(9, i), 1)
'                                .Item("35").value = varBufferEventi(22, i)
'                            Case "011"
'                                strIDDepositoRientro = vbNullString
'                                With ADE.DS
'                                    .Fields.ResetFields
'                                    .Fields("0").FindValue = varBufferEventi(23, i)
'                                    If (Not ADE.Find) Or (.HowMany = 0) Then
'                                        intRowError = NoDeposito
'                                        GoTo ErrHandler
'                                    End If
'                                    strIDDepositoRientro = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
'                                End With
'                                .Item("36").value = strIDProdotto
'                                .Item("37").value = varBufferEventi(8, i)
'                                .Item("38").value = Left(varBufferEventi(9, i), 1)
'                                .Item("39").value = strIDDepositoRientro
'                                .Item("40").value = Format$(varBufferEventi(24, i), "YYYYMMDD")
'                            Case "012"
'                                strIDDestinazioneDirottamento = vbNullString
'                                With ADZ.DS
'                                    .Fields.ResetFields
'                                    .Fields("0").FindValue = varBufferEventi(25, i)
'                                    If (Not ADZ.Find) Or (.HowMany = 0) Then
'                                        intRowError = NoDestinazione
'                                        GoTo ErrHandler
'                                    End If
'                                    strIDDestinazioneDirottamento = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
'                                    strIDDestinazioneDirottamento = Mid(strIDDestinazioneDirottamento, 3)
'                                End With
'                                .Item("41").value = strIDProdotto
'                                .Item("42").value = varBufferEventi(8, i)
'                                .Item("43").value = Left(varBufferEventi(9, i), 1)
'                                .Item("44").value = strIDDestinazioneDirottamento
'                                .Item("45").value = Format$(varBufferEventi(26, i), "YYYYMMDD")
'                            Case "013"
'                                .Item("46").value = Format$(Hour(varBufferEventi(27, i)), "00") & ":" & Format$(Minute(varBufferEventi(27, i)), "00")
'                                .Item("47").value = Format$(Hour(varBufferEventi(28, i)), "00") & ":" & Format$(Minute(varBufferEventi(28, i)), "00")
'                            Case "014"
'                                .Item("48").value = Format$(Hour(varBufferEventi(27, i)), "00") & ":" & Format$(Minute(varBufferEventi(27, i)), "00")
'                                .Item("49").value = Format$(Hour(varBufferEventi(28, i)), "00") & ":" & Format$(Minute(varBufferEventi(28, i)), "00")
'                        End Select
'
'                        If Not .WriteLine(True) Then
'                            intRowError = OutputError
'                            GoTo ErrHandler
'                        End If
'                        With VPI.DS
'                            .Fields.ResetFields
'                            .Fields("0").FindValue = varBufferEventi(37, i)
'                            .Fields("1").FindValue = varBufferEventi(38, i)
'                            .Fields("2").FindValue = varBufferEventi(39, i)
'                            .Fields("3").FindValue = varBufferEventi(40, i)
'                        End With
'                        VPI.Find
'                        If Not VPI.SetFlagTrasmissione("S") Then
'                            GoTo ErrHandler
'                        End If
'                        lngCountOK = lngCountOK + 1
'                    End With
'                Next i
'            End If
'        End If
'    Next lngRow
'    GoTo EndProc
'ErrHandler:
'EndProc:
'    mobjOutFileInc.CloseFile
'
'    If lngCountOK > 0 Then
'        With DataClass.DS.DA
'            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "008", 0)
'            If varCounter = -1 Then
'                ErrCode = InvalidCode
'                GoTo ErrHandler
'            End If
'            If varCounter > 9999999999# Then
'                varCounter = 1
'            End If
'        End With
'
'        strProgrInvio = Format(varCounter, String(10, "0"))
'        strFileName = strFileName & "\INC_" & strProgrInvio & ".DAT"
'        intFileFromNum = FreeFile
'        Open TEMP_FILE For Input Access Read As #intFileFromNum
'        intFileToNum = FreeFile
'        Open strFileName For Output Access Write As #intFileToNum
'        Do While Not EOF(intFileFromNum)
'            Line Input #intFileFromNum, strLine
'            Mid(strLine, 31, 4) = Format(lngCountOK, "0000")
'            Mid(strLine, 5, 10) = strProgrInvio
'            Print #intFileToNum, strLine
'        Loop
'        Close #intFileFromNum
'        Close #intFileToNum
'    End If
'
'End Sub
'
'
'Private Property Get ViaggiModificati() As Long
'    Const COUNT_VIAGGI_MODIFICATI = _
'        "select count(*) from viagvvp0 where vvpflst5 = 'S'"
'    Dim lngResult As Long
'    Dim varBuf As Variant
'    On Error GoTo ErrHandler
'
'    If Not DataClass.DS.DA.QueryData(COUNT_VIAGGI_MODIFICATI, varBuf, lngResult) Then
'        GoTo ErrHandler
'    End If
'    ViaggiModificati = varBuf(0, 0)
'Exit Property
'ErrHandler:
'    ViaggiModificati = 0
'End Property
'
'Private Property Get UnitàDiTrasportoAppoggio() As String
'    UnitàDiTrasportoAppoggio = mstrUdTAppoggio
'End Property
'
'Private Property Let UnitàDiTrasportoAppoggio(ByVal NewValue As String)
'    mstrUdTAppoggio = NewValue
'End Property

'' Dopo...di dubbia utilità, ma non lo commento
'Private Sub CreateTrips(DataViaggio As Variant)
'    On Error GoTo ErrHandler
'    Dim i As Long
'    Dim strIDAUTPrec As String
'    Dim strIDAUT As String
'    Dim strIDVVT As String
'    Dim strIDVVC As String
'    Dim strIDOCT As String
'    Dim varSeqPrec As Variant
'    Dim varSeq As Variant
'
'    XA.QuickSort 0, XA.UpperBound(1), _
'                 0, XORDER_ASCEND, XTYPE_DEFAULT, _
'                 1, XORDER_ASCEND, XTYPE_LONG
'
'    For i = 0 To XA.UpperBound(1)
'        If (Not IsNull(XA.value(i, 0))) And (Len(XA.value(i, 0)) > 0) Then
'            strIDAUT = XA.value(i, 0)
'            varSeq = XA.value(i, 1)
'            strIDOCT = XA.value(i, 15)
'            If (strIDAUT <> strIDAUTPrec) Or (varSeq <> varSeqPrec) Then
'                strIDAUTPrec = strIDAUT
'                varSeqPrec = varSeq
'                AUT.DS.Fields.ResetFields
'                AUT.DS.Fields("0").FindValue = strIDAUT
'                AUT.Find
'                VVT.DS.Fields.ResetFields
'                With VVT.DS
'
'                    .Fields("2").NewValue = varSeq
'                    .Fields("2").Updated = True
'                    .Fields("3").NewValue = XA.value(i, 23)
'                    .Fields("3").Updated = True
'                    .Fields("4").NewValue = DataViaggio
'                    .Fields("4").Updated = True
'                    .Fields("6").NewValue = DataViaggio
'                    .Fields("6").Updated = True
'                    .Fields("13").NewValue = "TD"
'                    .Fields("13").Updated = True
'                    .Fields("14").NewValue = strIDAUT
'                    .Fields("14").Updated = True
'                    .Fields("15").NewValue = AUT.DS.Fields("2").value
'                    .Fields("15").Updated = True
'                    .Fields("16").NewValue = AUT.DS.Fields("3").value
'                    .Fields("16").Updated = True
'                    .Fields("17").NewValue = AUT.DS.Fields("6").value
'                    .Fields("17").Updated = True
'                    .Fields("18").NewValue = AUT.DS.Fields("7").value
'                    .Fields("18").Updated = True
'                    .Fields("19").NewValue = XA.value(i, 2)
'                    .Fields("19").Updated = True
'                    .Fields("21").NewValue = AUT.DS.Fields("11").value
'                    .Fields("21").Updated = True
'                    .Fields("22").NewValue = XA.value(i, 9)
'                    .Fields("22").Updated = True
'                    .Fields("23").NewValue = XA.value(i, 9)
'                    .Fields("23").Updated = True
'                    .Fields("25").NewValue = 0
'                    .Fields("25").Updated = True
'                    .Fields("26").NewValue = "M"
'                    .Fields("26").Updated = True
'                    .Fields("27").NewValue = 0
'                    .Fields("27").Updated = True
'                    .Fields("28").NewValue = "M"
'                    .Fields("28").Updated = True
'                    .Fields("43").NewValue = "R"
'                    .Fields("43").Updated = True
'                End With
'                If Len(strIDVVT) > 0 Then
'                    VVT.DS.DA.EndTrans ETCOMMIT
'                End If
'                VVT.DS.DA.BeginTransaction
'                If Not VVT.Insert Then
'                    GoTo ErrHandler
'                End If
'                strIDVVC = "010"
'            End If
'            With VVC.DS
'                strIDVVT = VVT.DS.Fields("0").value
'                .Fields("0").NewValue = strIDVVT
'                .Fields("0").Updated = True
'                .Fields("1").NewValue = strIDVVC
'                .Fields("1").Updated = True
'                .Fields("2").NewValue = strIDVVC
'                .Fields("2").Updated = True
'            End With
'            If Not VVC.OrdineToConsegna(strIDOCT) Then
'                GoTo ErrHandler
'            End If
'            If Not VVC.Insert Then
'                GoTo ErrHandler
'            End If
'            If Not VVP.OrdineToProdotti(strIDOCT, strIDVVT, strIDVVC) Then
'                GoTo ErrHandler
'            End If
'            strIDVVC = Format$(CInt(strIDVVC) + 10, "000")
'        End If
'        DoEvents
'    Next i
'    VVT.DS.DA.EndTrans ETCOMMIT
'    Exit Sub
'ErrHandler:
'    VVT.DS.DA.EndTrans ETROLLBACK
'End Sub

'Public Property Let FunctionCode(NewValue As FunctionCodes)
'    mvarFunction = NewValue
'End Property

Private Sub DisplaySelection(List As Variant)
    Dim i As Long
    
    RefreshXArray XA, TDBGrid1, List
    TDBGrid1.Bookmark = 0
    
'    lblLog(5) = 0
'    lblLog(2) = 0
'    lblLog(4) = 0
'    If UBound(List, 2) >= 0 Then
'        ReDim varLog(0 To 3, 0 To UBound(List, 2))
'        For i = 0 To UBound(List, 2)
'            varLog(0, i) = List(0, i)
'            varLog(1, i) = List(1, i)
'            varLog(2, i) = List(2, i)
'        Next i
'        RefreshXArray XALog, tdbLog, varLog
'    End If
End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.Bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.Bookmark, 0)
    End If
End Function

' Dopo... per salvare i campi sul viaggio
'Private Function SaveTripData() As Boolean
'    On Error GoTo ErrHandler
'
'    Dim strData1 As String
'    Dim strData2 As String
'    Dim strData3 As String
'    Dim strAUTData As String
'    Dim i As Long
'    Dim lngUbound As Long
'    Dim blnConcErr As Boolean
'    lngUbound = XA.UpperBound(1)
'    For i = 0 To lngUbound
'        If (Not IsNull(XA.value(i, 0))) And (Not IsNull(XA.value(i, 1))) And (Not IsNull(XA.value(i, 2))) Then
'            OCT.DS.Fields("0").FindValue = XA.value(i, 15)
'            OCT.Find
'            strData1 = Space(3)
'            strData2 = Space(3)
'            strData3 = Space(3)
'            LSet strData1 = XA.value(i, 0)
'            LSet strData2 = XA.value(i, 1)
'            LSet strData3 = XA.value(i, 2)
'            strAUTData = strData1 & strData2 & strData3
'            OCT.DS.Fields("11").NewValue = strAUTData
'            OCT.DS.Fields("11").Updated = True
'            OCT.BeginSave blnConcErr
'            OCT.DS.DA.EndTrans ETCOMMIT
'        End If
'    Next i
'    SaveTripData = True
'    Exit Function
'ErrHandler:
'    OCT.DS.DA.EndTrans ETROLLBACK
'    SaveTripData = False
'End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    
    With DataClass.DS
        If .Fields.Validate Then
            If .ScanClientCond Then
                ReDim varOrderFields(2) As Variant
                Me.MousePointer = vbHourglass
                varOrderFields(0) = "37"
                varOrderFields(1) = "0"
'                varOrderFields(2) = "34"
                varOrderFields(2) = "38"
'                varOrderFields(3) = "94"
'                varOrderFields(4) = "96"
                blnActiveOnly = (chkActiveOnly = vbChecked)
                blnActiveOnly = True
                If Not DataClass.GetSelection(blnActiveOnly, varOrderFields, CLng(mskMaxRecords.Text)) Then
                    GoTo ErrHandler
                End If
                varBuffer = DataClass.DS.List
                If IsEmpty(varBuffer) Then
                    GoTo ErrHandler
                End If
                DisplaySelection varBuffer
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    SelectRecords = True
    GoTo EndProc
ErrHandler:
    SelectRecords = False
EndProc:
    Me.MousePointer = vbDefault
End Function

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strSelectionFormula As String
    
    ReDim varReportsList(0 To 2, 0 To 3) As Variant
    
    Select Case Cmd
        Case fcUndo
            If (State = fsParamInput) Or (State = fsSelection) Then
                ClearForm
                SetState fsParamInput
            End If
        Case fcFind
            If SelectRecords Then
                SetState fsSelection
            End If
        Case fcPrint
'            If TDBGrid1.Bookmark <> vbNullString Then
'                Call SincronizzaDatiViaggio(srLetturaDati)
'                Call SetInterfaceForUpdTrip(arModificaViaggio)
'            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcBrowse
'            With CommonDialog1
'                .fileName = mskFileName
'                .DialogTitle = "Seleziona il file per la trasmissione... "
'                .InitDir = App.Path
'                .Filter = "Tutti i files (*.*)|*.*| "
'                .FilterIndex = 1
'                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
'                .ShowSave
'                mskFileName = .fileName
'            End With
        Case fcExport
'            ExportRecords
'            ExportEventi
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub GetRecord(strID As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields.ResetFields
        .Fields("2").FindValue = strID
        If Not DataClass.Find Then
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

'Private Sub SetInterfaceForUpdTrip(ByVal pRichiesta As AzioneRichiesta)
'    On Error GoTo ErrHandler
'
'    Select Case pRichiesta
'        Case arModificaViaggio
'            TDBGrid1.Enabled = False
'            FraModificaViaggio.Visible = True
'            FraTrasmettiViaggi.Visible = False
'            CmdOpenTrip.Enabled = False
'            cmdOpenFile.Enabled = False
'            cmdExport.Enabled = False
'            Command4.Enabled = False
'            Command5.Enabled = False
'            Command8.Enabled = False
'            MaskEdBox1.SetFocus
'        Case arAnnullaModifica
'            TDBGrid1.Enabled = False
'            FraModificaViaggio.Visible = True
'            FraTrasmettiViaggi.Visible = False
'            CmdOpenTrip.Enabled = False
'            cmdOpenFile.Enabled = False
'            cmdExport.Enabled = False
'            Command4.Enabled = False
'            Command5.Enabled = False
'            Command8.Enabled = False
'        Case arLetturaViaggi
'            TDBGrid1.Enabled = True
'            FraModificaViaggio.Visible = False
'            FraTrasmettiViaggi.Visible = True
'            CmdOpenTrip.Enabled = True
'            cmdOpenFile.Enabled = True
'            cmdExport.Enabled = True
'            Command4.Enabled = True
'            Command5.Enabled = True
'            Command8.Enabled = True
'    End Select
'    lblLog(5).Caption = CStr(ViaggiModificati)
'Exit Sub
'ErrHandler:
'    TDBGrid1.Enabled = True
'    FraModificaViaggio.Visible = False
'End Sub

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsParamInput
            If Not Loading Then
'                TabStrip1.Tabs(1).Selected = True
'                mskField9.SetFocus
            End If
'            MaskEdBox12 = Format(Date, "dd/mm/yy")
'            MaskEdBox14 = Format(Date, "dd/mm/yy")
'            mskField9 = "038"
            'Combo3 = "ACQUISITO"
            'Combo3.Locked = True
        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Viaggi trovati: " & DataClass.DS.HowMany
'            TabStrip1.Tabs(2).Selected = True
'            If Not Loading Then
'                TDBGrid1.SetFocus
'            End If
    End Select
'    Call SetInterfaceForUpdTrip(arLetturaViaggi)
End Sub

'Private Sub cmdCalendar_Click(Index As Integer)
'    Dim strTitle As String
'    Dim M As MaskEdBox
'
'    Select Case Index
'        Case 0
'            Set M = MaskEdBox12
'            strTitle = "Dalla data inizio viaggio"
'        Case 4
'            Set M = MaskEdBox14
'            strTitle = "Alla data inizio viaggio"
'    End Select
'    GetCalendarDate M, strTitle
'End Sub

Private Sub cmdEventi_Click()
    Dim LinkForm As Form
    
    Set LinkForm = DisplayForm(fcVIAGVPI)
    If Not LinkForm Is Nothing Then
        With LinkForm
            Set .ParentForm = Me
            If .GetRecord(XA.value(TDBGrid1.Bookmark, 0), XA.value(TDBGrid1.Bookmark, 1), XA.value(TDBGrid1.Bookmark, 2)) Then
                .SetState fsEdit
            Else
                .SetState fsInsert
'                .mskLink = XA.value(TDBGrid1.Bookmark, 0)
                .DataClass.DS.Fields("0").NewValue = XA.value(TDBGrid1.Bookmark, 0)
                .DataClass.DS.Fields("0").Updated = True
                .DataClass.DS.Fields("1").NewValue = XA.value(TDBGrid1.Bookmark, 1)
                .DataClass.DS.Fields("1").Updated = True
                .DataClass.DS.Fields("2").NewValue = XA.value(TDBGrid1.Bookmark, 2)
                .DataClass.DS.Fields("2").Updated = True
                .DataClass.DS.Fields("3").NewValue = "1"
                .DataClass.DS.Fields("3").Updated = True
                .DataClass.DS.Fields("11").NewValue = XA.value(TDBGrid1.Bookmark, 28)
                .DataClass.DS.Fields("11").Updated = True
                
'                .mskLink1 = XA.value(TDBGrid1.Bookmark, 1)
'                .mskLink2 = XA.value(TDBGrid1.Bookmark, 2)
'                .mskField1 = "1"
'                .mskField11 = XA.value(TDBGrid1.Bookmark, 28)
                .MaskEdBox24 = Format(XA.value(TDBGrid1.Bookmark, 3), "dd/mm/yy")
'                .DataClass.DS.ScanClient smInsert
            End If
            Me.Enabled = False
        End With
    End If
End Sub

Private Sub CmdSalva_Click()
    Dim strErr As String
    
    ' Salva i dati eventualmente modificati sul viaggio
'    If SincronizzaDatiViaggio(srScritturaDati) Then
'        ' Chiude il frame di modifica viaggi e rinfresca il viaggio sulla griglia
'        Call RipristinaSelezione
'    End If
End Sub

Private Sub CmdAnnulla_Click()
    ' Rileggi i valori del viaggio
'    Call PulisciDatiModificatiViaggio
'    Call SincronizzaDatiViaggio(srLetturaDati)
End Sub

Private Sub CmdRitorna_Click()
'    Call RipristinaSelezione
End Sub

Private Sub RipristinaSelezione()
'    Dim BKM As Variant
    
'    BKM = TDBGrid1.Bookmark                         ' Mi salvo il Bookmark della griglia
'    Call SetInterfaceForUpdTrip(arLetturaViaggi)    ' Chiudo il frame che permette la modifca dei dati del viaggio
'    'TDBGrid1.Array = XA                             ' Riassegno alla griglia l'array con la segnalazione del record modificato...
'    'TDBGrid1.ReBind                                 ' ..in questo modo se ho modificato una testata per un viaggio che ha + prodotti però ho un disallineamento in  griglia
'    PulisciDatiModificatiViaggio                    ' Pulisco i controlli con gli eventuali dati modificati
'    ' Rinfrescare è rischioso: se il viaggio modificato non rispondesse più alla selezione?
'    DoCommand fcFind                                ' rinfresco la griglia dei viaggi selezionati
'    TDBGrid1.SetFocus                               ' Assegno il fuoco alla griglia
'    TDBGrid1.Bookmark = BKM                         ' Mi posiziono col Bookmark della griglia sul viaggio che stavo modificando
End Sub

'Private Function ValidateTestataViaggio(ByRef rErr As String) As Boolean
'    On Error GoTo ErrHandler
'
'    ' Verifico la validità dell'ora di carico
'    If MaskEdBox1.Text = vbNullString Then
'        rErr = "Occorre indicare l'ora di carico."
'        GoTo ErrHandler
'    Else
'        If MaskEdBox1.Text <> TIME_MASK Then
'            If Not IsDate(Format(MaskEdBox1.Text, "HH:MM")) Then
'                rErr = "Occore indicare un ora di carico corretta."
'                GoTo ErrHandler
'            End If
'        End If
'    End If
'
'    ' Verifico la validità dell'unità di trasporto
'    If MaskEdBox4.Text = vbNullString Then
'        rErr = "Occorre indicare l'unità di trasporto."
'        GoTo ErrHandler
'    Else
'        With AUT.DS
'            .Fields.ResetFields
'            .Fields("0").FindValue = MaskEdBox4.Text    ' Codice
'            .Fields("9").FindValue = DISPONIBILE        ' Disponibilità
'            .Fields("16").FindValue = 0                 ' Annullato
'            If Not AUT.Find Or .HowMany = 0 Then
'                rErr = "Unità di trasporto " + MaskEdBox4.Text + " non disponibile."
'                GoTo ErrHandler
'            End If
'        End With
'    End If
'
'    ' Verifico la validità della motrice
'    If MaskEdBox3.Text = vbNullString Then
'        rErr = "Occorre indicare la motrice."
'        GoTo ErrHandler
'    Else
'        With AAM.DS
'            .Fields.ResetFields
'            .Fields("0").FindValue = MaskEdBox3.Text    ' Codice
'            .Fields("1").FindValue = MOTRICE            ' Motrice/Rimorchio
'            .Fields("4").FindValue = DISPONIBILE        ' Disponibilità
'            .Fields("26").FindValue = 0                 ' Annullato
'            If Not AAM.Find Or .HowMany = 0 Then
'                rErr = "Motrice " + MaskEdBox3.Text + " non disponibile."
'                GoTo ErrHandler
'            End If
'        End With
'    End If
'
'    ' Verifico la validità del rimorchio
'    If MaskEdBox5.Text <> vbNullString Then
'        With AAM.DS
'            .Fields.ResetFields
'            .Fields("0").FindValue = MaskEdBox5.Text    ' Codice
'            .Fields("1").FindValue = RIMORCHIO          ' Motrice/Rimorchio
'            .Fields("4").FindValue = DISPONIBILE        ' Disponibilità
'            .Fields("26").FindValue = 0                 ' Annullato
'            If Not AAM.Find Or .HowMany = 0 Then
'                rErr = "Rimorchio " + MaskEdBox5.Text + " non disponibile."
'                GoTo ErrHandler
'            End If
'        End With
'    End If
'
'    ' Verifico la validità del primo autista
'    If MaskEdBox6.Text = vbNullString Then
'        rErr = "Occorre indicare il primo autista."
'        GoTo ErrHandler
'    Else
'        With AAU.DS
'            .Fields.ResetFields
'            .Fields("0").FindValue = MaskEdBox6.Text    ' Codice
'            .Fields("5").FindValue = DISPONIBILE        ' Disponibilità
'            .Fields("8").FindValue = 0                  ' Annullato
'            If Not AAU.Find Or .HowMany = 0 Then
'                rErr = "Primo autista " + MaskEdBox6.Text + " non disponibile."
'                GoTo ErrHandler
'            End If
'        End With
'    End If
'
'    ' Verifico la validità dei km viaggio
'    If MaskEdBox18.Text <> vbNullString Then
'        If Not IsNumeric(MaskEdBox18.Text) Then
'            rErr = "Controllare i km viaggio"
'            GoTo ErrHandler
'        End If
'    End If
'
'    If Not VVT.CheckSequenza(XA.value(TDBGrid1.Bookmark, 3), MaskEdBox11.Text, MaskEdBox4.Text, XA.value(TDBGrid1.Bookmark, 4), XA.value(TDBGrid1.Bookmark, 0)) Then
'        rErr = "Il numero di sequenza è già in uso"
'        GoTo ErrHandler
'    End If
'
'    ValidateTestataViaggio = True
'Exit Function
'ErrHandler:
'    ValidateTestataViaggio = False
'End Function

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim strErr As String
    
    ClientKeyUp KeyCode, Me
    Select Case KeyCode
        Case vbKeyF2
            ' Salva i dati eventualmente modificati sul viaggio
'            If Not TDBGrid1.Enabled Then
'                If SincronizzaDatiViaggio(srScritturaDati) Then
'                    ' Chiude il frame di modifica viaggi e rinfresca il viaggio sulla griglia
'                    Call RipristinaSelezione
'                End If
'            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strIDParam As String
    ReDim ColTags(1 To 2) As Variant
    Dim Col As Column
    
    Loading = True
'    mintCurFrame = 1
    ColTags(1) = "0"
    ColTags(2) = "1"
    mskMaxRecords = MAX_RECORDS
'    mskFileName = gmRegistry.GetSetting("Options", "TripChangedExport", , True)
    
'    Set mobjOutFile = New CTextFile
'    With mobjOutFile
'        .Add "Ente Emittente", 4, "1"
'        .Add "Progressivo Invio", 10, "2", , fftNumerico
'        .Add "Data Sistema", 8, "3"
'        .Add "Ora Sistema", 8, "4"
'        .Add "Numero Record", 4, "5", , fftNumerico
'        .Add "Rif. Esterno 1 Consegna", 13, "6"
'        .Add "Partizione", 2, "7", "00"
'        .Add "Destinazione", 8, "8"
'        .Add "GPS", 2, "9", , fftNumerico
'        .Add "X GPS", 8, "10", , fftNumerico
'        .Add "Y GPS", 8, "11", , fftNumerico
'        .Add "Prodotto", 7, "12"
'        .Add "Deposito", 4, "13"
'        .Add "Q.tà Concordata Ordine", 9, "14", , fftNumerico
'        .Add "Unità di Misura", 1, "15", "L"
'        .Add "Data Viaggio", 8, "16"
'        .Add "Data Consegna aggiornata", 8, "17"
'        .Add "Ora Consegna aggiornata", 5, "18"
'        .Add "Q.tà Trasportata", 9, "19", , fftNumerico
'        .Add "Q.tà Scaricata", 9, "20", , fftNumerico
'        .Add "Data Viaggio 2", 8, "21"
'        .Add "Ora Inizio Carico", 5, "22"
'        .Add "Data Consegna effettiva", 8, "23"
'        .Add "Ora Consegna effettiva", 5, "24"
'        .Add "Badge UdT", 11, "25"
'        .Add "Nr. Sequenza", 2, "26"
'        .Add "Nr. Carta Identità", 20, "27"
'        .Add "Km Definitivi", 4, "28", , fftNumerico
'        .Add "Km CEPSA", 4, "29", , fftNumerico
'        .Add "Quantità 1", 6, "30", "000.00", fftAlfa
'        .Add "Quantità 2", 6, "31", "000.00", fftAlfa
'        .Add "Quantità 3", 6, "32", "000.00", fftAlfa
'        .Add "Numero Consegna", 2, "33", , fftNumerico
'        .Add "Note Consegna", 40, "34"
'        .Add "Note Estese", 100, "35"
'    End With
'
'    Set mobjOutFileInc = New CTextFile
'    With mobjOutFileInc
'        .Add "Ente Emittente", 4, "1", "0056"
'        .Add "Progressivo Invio", 10, "2", , fftNumerico
'        .Add "Data Generazione", 8, "3"
'        .Add "Ora Generazione", 8, "4"
'        .Add "Numero Records", 4, "5", , fftNumerico
'        .Add "Riferimento Ordine", 13, "6", , fftNumerico               ' Ordine + riga
'        .Add "Partizione Ordine", 2, "7", "00", fftNumerico
'        .Add "Codice Destinazione", 8, "8"
'        .Add "Tipo evento", 3, "9", , fftNumerico
'        .Add "Motivo evento", 3, "10", , fftNumerico
'        .Add "Data evento", 8, "11"
'        .Add "Ora inizio evento", 5, "12"
'        .Add "Ora termine evento", 5, "13"
'        .Add "Ordine precedente responsabile evento", 13, "14", , fftNumerico
'        .Add "Partizione ordine precedente", 2, "15", , fftNumerico
'        .Add "Evento precedente", 3, "16", , fftNumerico
'
'        .Add "Prodotto contaminante (CLH)", 7, "17"
'        .Add "Quantità prodotto contaminante", 9, "18", , fftNumerico
'        .Add "Unità di misura prodotto contaminante", 1, "19"
'        .Add "Prodotto contaminato 1 (CLH)", 7, "20"
'        .Add "Quantità prodotto contaminato 1", 9, "21", , fftNumerico
'        .Add "Unità di misura prodotto contaminato 1", 1, "22"
'        .Add "Prodotto contaminato 2 (CLH)", 7, "23"
'        .Add "Quantità prodotto contaminato 2", 9, "24", , fftNumerico
'        .Add "Unità di misura prodotto contaminato 2", 1, "25"
'        .Add "Prodotto contaminato 3 (CLH)", 7, "26"
'        .Add "Quantità prodotto contaminato 3", 9, "27", , fftNumerico
'        .Add "Unità di misura prodotto contaminato 3", 1, "28"
'        .Add "Prodotto contaminato 4 (CLH)", 7, "29"
'        .Add "Quantità prodotto contaminato 4", 9, "30", , fftNumerico
'        .Add "Unità di misura prodotto contaminato 4", 1, "31"
'
'        .Add "Prodotto spandimento (CLH)", 7, "32"
'        .Add "Quantità prodotto spandimento", 9, "33", , fftNumerico
'        .Add "Unità di misura prodotto spandimento", 1, "34"
'        .Add "Misure adottate", 50, "35"
'
'        .Add "Prodotto rientrato (CLH)", 7, "36"
'        .Add "Quantità prodotto rientrato", 9, "37", , fftNumerico
'        .Add "Unità di misura prodotto rientrato", 1, "38"
'        .Add "Deposito rientro (CLH)", 4, "39"
'        .Add "Data rientro in deposito", 8, "40"
'
'        .Add "Prodotto dirottato (CLH)", 7, "41"
'        .Add "Quantità prodotto dirottato", 9, "42", , fftNumerico
'        .Add "Unità di misura prodotto dirottato", 1, "43"
'        .Add "Destinazione dirottamento", 8, "44"
'        .Add "Data consegna prodotto dirottato", 8, "45"
'
'        .Add "Ora inizio paralisi deposito", 5, "46"
'        .Add "Ora termine paralisi deposito", 5, "47"
'        .Add "Ora inizio paralisi destinazione", 5, "48"
'        .Add "Ora termine paralisi destinazione", 5, "49"
'
'        .Add "Data comunicazione evento", 8, "50"
'        .Add "Responsabile comunicazione evento", 40, "51"
'        .Add "Osservazioni", 200, "52"
'
'    End With
'
    Set DataClass = New CViaggioSelezioneSatellitare
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
'    Set VVT = New CViaggio
'    With VVT.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set VVC = New CViaggioConsegna
'    With VVC.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set VVP = New CViaggioProdotto
'    With VVP.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set VPI = New CViaggioEvento
'    With VPI.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set OCT = New COrdine
'    With OCT.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set APR = New CProdotto
'    With APR.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set AUT = New CUnitaDiTrasporto
'    With AUT.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set AAM = New CAutomezzo
'    With AAM.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set AAU = New CAutista
'    With AAU.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set AEE = New CEnteEmittente
'    With AEE.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set ADE = New CDeposito
'    With ADE.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
'    Set ADZ = New CDestinazione
'    With ADZ.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = strIDParam
        AXX.Find
    End With
    ' ---
    For Each Col In TDBGrid1.Columns
        Col.ButtonHeader = True
    Next Col
    
    SetState fsParamInput
    CmdOpenTrip.Enabled = False
    cmdOpenFile.Enabled = False
    cmdExport.Enabled = False
    
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
'    If lblLog(5).Caption = "0" Then
'        If Not UpdModificato(vbNullString, vbNullString, vbNullString, False, False, False, "NULL") Then
'            Cancel = 1
'        End If
'    Else
'        If MsgBox("Stai uscendo senza aver trasmesso i viaggi modificati. Vuoi uscire comunque?", vbYesNo + vbExclamation, "Attenzione") <> vbYes Then
'            Cancel = 1
'        Else
'            If Not UpdModificato(vbNullString, vbNullString, vbNullString, False, False, False, "NULL") Then
'                Cancel = 1
'            End If
'        End If
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    gmRegistry.SaveSetting "Options", "TripChangedExport", mskFileName, True
'    Set AUT = Nothing
'    Set AAU = Nothing
'    Set AAM = Nothing
'    Set AAM = Nothing
'    Set ADE = Nothing
'    Set ADZ = Nothing
    Set AXX = Nothing
'    Set VPI = Nothing
'    Set XA = Nothing
'    Set XALog = Nothing
'    Set mobjOutFile = Nothing
'    Set mobjOutFileInc = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set frmVVTSnapshotExport = Nothing
End Sub

'Private Sub DisplayScomparti(IDAutomezzo As String, L As Label)
'    On Error GoTo ErrHandler
'
'    Dim i As Integer
'    Dim strScomparti As String
'    Dim dblPortata As Double
'    Dim dblPortataM As Double
'    Dim dblPortataR As Double
'
'    L = vbNullString
'    If Len(IDAutomezzo) = 0 Then
'        dblPortata = 0
'    Else
'        With AAM.DS
'            .Fields("0").FindValue = IDAutomezzo
'            If Not AAM.Find Then
'                GoTo ErrHandler
'            End If
'            If AAM.DS.HowMany = 0 Then
'                GoTo ErrHandler
'            End If
'            For i = 1 To 10
'                If .Fields(CStr(i * 2 + 4)).value <> 0 Then
'                    strScomparti = strScomparti & .Fields(CStr(i * 2 + 4)).value & " x " & .Fields(CStr((i * 2 - 1) + 4)).value & " "
'                    dblPortata = dblPortata + .Fields(CStr(i * 2 + 4)).value * .Fields(CStr((i * 2 - 1) + 4)).value
'                End If
'            Next i
'        End With
'    End If
'    L = " " & strScomparti
'
'    Exit Sub
'
'ErrHandler:
'    L = vbNullString
'End Sub

'Private Sub TabStrip1_Click()
'    If TabStrip1.SelectedItem.Index = mintCurFrame Then
'        Exit Sub
'    End If
'    Frame1(TabStrip1.SelectedItem.Index).Visible = True
'    Frame1(mintCurFrame).Visible = False
'    mintCurFrame = TabStrip1.SelectedItem.Index
'    If mintCurFrame = 1 Then
'        CmdOpenTrip.Enabled = False
'        cmdOpenFile.Enabled = False
'        cmdExport.Enabled = False
'    Else
'        CmdOpenTrip.Enabled = True
'        cmdOpenFile.Enabled = True
'        cmdExport.Enabled = True
'    End If
'End Sub

'Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
'    Dim i As Long
'    Dim strIDAUT As String
'    Dim varMaxSeq As Variant
'
'    mblnIsDirty = True
'    Select Case ColIndex
'        Case 0
'            strIDAUT = TDBGrid1.Columns(ColIndex).value
'            AUT.DS.Fields.ResetFields
'            AUT.DS.Fields("0").FindValue = strIDAUT
'            AUT.Find
'            If AUT.DS.HowMany > 0 Then
''                For i = 0 To XA.UpperBound(1)
''                    If XA.value(i, 0) = strIDAUT Then
''                        If XA.value(i, 1) > varMaxSeq Then
''                            varMaxSeq = XA.value(i, 1)
''                        End If
''                    End If
''                Next i
''                If IsEmpty(varMaxSeq) Then
''                    varMaxSeq = 0
''                End If
''                TDBGrid1.Columns(1).value = varMaxSeq + 1
'                TDBGrid1.Columns(2).value = AUT.DS.Fields("4").value
'                DisplayScomparti AUT.DS.Fields("2").value, lblScompartiM
'                DisplayScomparti AUT.DS.Fields("3").value, lblScompartiR
'                AAU.DS.Fields("0").FindValue = AUT.DS.Fields("4").value
'                AAU.Find
'                TDBGrid1.Columns(3).value = AAU.DS.Fields("1").value
'            End If
'            If Len(strIDAUT) = 0 Then
'                TDBGrid1.Columns(1).value = vbNullString
'                TDBGrid1.Columns(3).value = vbNullString
'            End If
'        Case 2
'            AAU.DS.Fields.ResetFields
'            AAU.DS.Fields("0").FindValue = TDBGrid1.Columns(ColIndex).value
'            AAU.Find
'            If AAU.DS.HowMany > 0 Then
'                TDBGrid1.Columns(3).value = AAU.DS.Fields("1").value
'            Else
'                TDBGrid1.Columns(3).value = vbNullString
'            End If
'    End Select
'End Sub

Private Sub TDBGrid1_Change()
'    Select Case TDBGrid1.Col
'        Case 0, 2
'            If TDBGrid1.Columns(TDBGrid1.Col).value = "?" Then
'                TDBGrid1.Columns(TDBGrid1.Col).value = ""
'                Dim i As Integer
'                For i = 0 To Controls.Count - 1
'                    If Controls(i).Name = "TDBGrid1" Then
'                        frmSearch.LinkControlIndex = i
'                        Exit For
'                    End If
'                Next i
'                If TDBGrid1.Col = 0 Then
'                    ShowSearchForm Me, "", AUT, TDBGrid1.TabIndex, True
'                Else
'                    ShowSearchForm Me, "", AAU, TDBGrid1.TabIndex, True
'                End If
'                Do While IsLoaded("frmSearch")
'                    DoEvents
'                Loop
'                SendKeys "{TAB}"
'            End If
'    End Select

End Sub

Private Sub TDBGrid1_DblClick()
    DoCommand fcOpen
End Sub

' Dopo... La uso per colorare le CMP modificate
Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
'    If (XA(Bookmark, 25) = "S") Then
'        RowStyle.BackColor = vbGreen
'        RowStyle.ForeColor = vbBlack
'    Else
'        If (XA(Bookmark, 32) = "S") Then
'            RowStyle.BackColor = vbRed
'            RowStyle.ForeColor = vbWhite
'        End If
'    End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim x As XTYPE
    Dim o As XORDER
    Static blnLastSortASC  As Boolean
    Static intLastCol As Integer
    
    If ColIndex <> intLastCol Then
        blnLastSortASC = False
    Else
        blnLastSortASC = Not blnLastSortASC
    End If
    If blnLastSortASC Then
        o = XORDER_DESCEND
    Else
        o = XORDER_ASCEND
    End If
        
    Select Case ColIndex
        Case 0, 3, 4, 16, 17, 18, 20, 21
            x = XTYPE_LONG
        Case 5, 6, 7, 15
            x = XTYPE_DATE
        Case Else
            x = XTYPE_STRING
    End Select
    XA.QuickSort 0, XA.UpperBound(1), ColIndex, o, x
    TDBGrid1.ReBind
    intLastCol = ColIndex
End Sub


'Private Sub ExportRecords()
'    On Error GoTo ErrHandler
'
'    Const TEMP_FILE = "GES.tmp"
'
'    Dim blnNeedTrasmission As Boolean
'    Dim blnFineCarico As Boolean
'    Dim lngRow As Long
'    Dim lngNumRows As Long
'    Dim lngCountOK As Long
'    Dim lngCountFail As Long
'    Dim lngCountToGo As Long
'    Dim lngCountExcluded As Long
'    Dim lngProgrConsegna As Long
'    Dim lngKMCepsa As Long
'    Dim lngKM As Long
'    Dim lngPortataVeicoloKG As Long
'    Dim lngTotQuantitàKG As Long
'    Dim i As Long
'    Dim j As Long
'    Dim intRowsTrip As Integer
'    Dim intFileFromNum As Integer
'    Dim intFileToNum As Integer
'    Dim intProgressivoViaggio As Integer
'    Dim intProgressivoConsegna As Integer
'    Dim intCountSospesi As Integer
'    Dim strCodOperatore As String
'    Dim strDestinazione As String
'    Dim strCartaIdentitaAutista1 As String
'    Dim strIDZona As String
'    Dim strIDZonaPrec As String
'    Dim strIDViaggio As String
'    Dim strIDViaggioPrec As String
'    Dim strIDEnteEmittente As String
'    Dim strIDDeposito As String
'    Dim strProdotto As String
'    Dim strFineCarico As String
'    Dim strIDDestinazione As String
'    Dim strIDTipoVendita As String
'    Dim strDataViaggio As String
'    Dim strDataInvio As String
'    Dim strOraInvio As String
'    Dim strProgrInvio As String
'    Dim strOraConsegna As String
'    Dim strOraInizioCarico As String
'    Dim varBadge As Variant
'    Dim strCentroOperativo As String
'    Dim strUnitàTrasporto As String
'    Dim strDestinatario As String
'    Dim strFileName As String
'    Dim strRifAGIP As String
'    Dim strStatoOrdine As String
'    Dim strTipoVendita As String
'    Dim strErrorMessage As String
'    Dim strHeader As String
'    Dim strDummy As String
'    Dim strLine As String
'    Dim intRowError As Errors
'    Dim varCounter As Variant
'    Dim varDensitàProdotto As Variant
'    Dim varCUCarga As Variant
'    Dim varCUReparto As Variant
'    Dim varCUTotale As Variant
'    Dim varEventi As Variant
''    Dim varTotaliViaggio As Variant
'    Dim ErrCode As Errors
'    Dim C As CError
'
'    Me.MousePointer = vbHourglass
'    Me.Enabled = False
'    ErrCode = NoError
'
''    ' --- Leggo i valori di default del sistema
''    With AXX.DS
''        .Fields(.KeyFields(1)).FindValue = ID_PARAM
''        If Not AXX.Find Then
''            ErrCode = ParameterNotFound
''            GoTo ErrHandler
''        End If
''        If .HowMany = 0 Then
''            ErrCode = ParameterNotFound
''            GoTo ErrHandler
''        End If
''        strIDEnteEmittente = .Fields("16").value
''    End With
''    If Len(strIDEnteEmittente) = 0 Then
''        ErrCode = NoEnteEmittente
''        GoTo ErrHandler
''    End If
'    ' ---
'
'    strCodOperatore = vbNullString 'Format$(mskCodOperatore, "00")
'    'mobjOutFile.Delimiter = "|"
'
'    'If Len(mstrProgrInvio) = 0 Then
'        With DataClass.DS.DA
'            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "007", 0)
'            If varCounter = -1 Then
'                ErrCode = InvalidCode
'                GoTo ErrHandler
'            End If
'            If varCounter > 9999999999# Then
'                varCounter = 1
'            End If
'        End With
'    'End If
'
'    If Len(mskFileName) > 0 Then
'        strFileName = mskFileName
'    Else
'        ErrCode = FileNotFound
'        GoTo ErrHandler
'    End If
'
'    If Len(Dir(TEMP_FILE)) > 0 Then
'        Kill TEMP_FILE
'    End If
'    If Not mobjOutFile.OpenFile(TEMP_FILE, 1) Then
'        ErrCode = FileNotFound
'        GoTo ErrHandler
'    End If
'
'    strDataInvio = Format$(Date, "YYYYMMDD")
'    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
'
'    lngNumRows = XA.UpperBound(1)
'    If chkTransAll.value = vbChecked Then
'        lngCountToGo = lngNumRows + 1
'        lblLog(5).Caption = lngNumRows + 1
'    Else
'        lngCountToGo = ViaggiModificati
'    End If
'
'    ' --- Ciclo di elaborazione dei viaggi
'    For lngRow = 0 To lngNumRows
'        intRowError = NoError
'        blnNeedTrasmission = True
'
'        If (XA(lngRow, 32) = "S") Or (chkTransAll.value = vbChecked) Then
'            strIDViaggio = XA(lngRow, 5)
'            If strIDViaggio <> strIDViaggioPrec Then
'                strIDViaggioPrec = strIDViaggio
'
'                lngProgrConsegna = 0
'                lngPortataVeicoloKG = 0
''                lngTotQuantitàKG = 0
'                With VVT.DS
'                    .Fields.ResetFields
'                    .Fields("0").FindValue = strIDViaggio
'                    VVT.Find
''                    VVT.GetTotaliProdotto
''                    For i = 0 To UBound(VVT.TotaliQtaProdotto, 2)
''                        With APR.DS
''                            .Fields.ResetFields
''                            .Fields("0").FindValue = VVT.TotaliQtaProdotto(0, i)
''                            APR.Find
''                            If IsNumeric(APR.DS.Fields("9").value) Then
''                                varDensitàProdotto = APR.DS.Fields("9").value
''                            Else
''                                varDensitàProdotto = 0
''                            End If
''                        End With
''                        lngTotQuantitàKG = lngTotQuantitàKG + Round(VVT.TotaliQtaProdotto(2, i) * varDensitàProdotto)
''                    Next i
'                End With
'                With AAM.DS
'                    .Fields.ResetFields
'                    .Fields("0").FindValue = VVT.DS.Fields("15").value
'                    AAM.Find
'                    If IsNumeric(.Fields("25").value) Then
'                        lngPortataVeicoloKG = .Fields("25").value
'                    End If
'                End With
'                If Len(VVT.DS.Fields("16").value) > 0 Then
'                    With AAM.DS
'                        .Fields.ResetFields
'                        .Fields("0").FindValue = VVT.DS.Fields("16").value
'                        AAM.Find
'                        If IsNumeric(.Fields("25").value) Then
'                            lngPortataVeicoloKG = lngPortataVeicoloKG + .Fields("25").value
'                        End If
'                    End With
'                End If
'                If Not VVT.ComputeCU(strIDViaggio, lngPortataVeicoloKG, varCUCarga, varCUReparto, varCUTotale) Then
'                    ErrCode = FasciaTariffariaError
'                    GoTo LogResult
'                End If
'            End If
'            lngProgrConsegna = lngProgrConsegna + 1
'
'            With AEE.DS
'                .Fields.ResetFields
'                .Fields(.KeyFields(1)).FindValue = XA(lngRow, 7)
'                If Not AEE.Find Then
'                    ErrCode = NoEnteEmittente
'                    GoTo LogResult
'                End If
'                If .HowMany = 0 Then
'                    ErrCode = NoEnteEmittente
'                    GoTo LogResult
'                End If
'                strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(XA(lngRow, 33))
'            End With
'
'            With ADE.DS
'                .Fields.ResetFields
'                .Fields("0").FindValue = XA(lngRow, 4)
'                If (Not ADE.Find) Or (.HowMany = 0) Then
'                    intRowError = NoDeposito
'                    GoTo LogResult
'                End If
'                strIDDeposito = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
'                'strCentroOperativo = .Fields("8").value
'            End With
'
'            With ADZ.DS
'                .Fields.ResetFields
'                .Fields("0").FindValue = XA(lngRow, 20)
'                If (Not ADZ.Find) Or (.HowMany = 0) Then
'                    intRowError = NoDestinazione
'                    GoTo LogResult
'                End If
'                strDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
'                If IsNumeric(.Fields("29").value) Then
'                    lngKMCepsa = .Fields("29").value
'                Else
'                    lngKMCepsa = 0
'                End If
'            End With
'
'            With AUT.DS
'                .Fields.ResetFields
'                .Fields("0").FindValue = XA(lngRow, 11)
'                If (Not AUT.Find) Or (.HowMany = 0) Then
'                    intRowError = noUT
'                    GoTo LogResult
'                End If
'                varBadge = .MultiSoc.DS.GetFirmCode(XA(lngRow, 4))
'                If IsNull(varBadge) Or Len(varBadge) = 0 Then
'                    intRowError = InvalidCode
'                    GoTo LogResult
'                End If
''                varBadge = .Fields("14").value
'            End With
'
'            With APR.DS
'                .Fields.ResetFields
'                .Fields("0").FindValue = XA(lngRow, 28)
'                If (Not APR.Find) Or (.HowMany = 0) Then
'                    intRowError = NoProdotto
'                    GoTo LogResult
'                End If
'                strProdotto = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
'            End With
'
'            With AAU.DS
'                .Fields.ResetFields
'                .Fields("0").FindValue = XA(lngRow, 14)
'                If (Not AAU.Find) Or (.HowMany = 0) Then
'                    intRowError = NoAutista
'                    GoTo LogResult
'                End If
'                strCartaIdentitaAutista1 = "ES" & Replace(Replace(.Fields("9").value, " ", ""), "-", "")
'            End With
'
'            strDataViaggio = Format(XA(lngRow, 3), "YYYYMMDD")
'            strOraConsegna = Format(XA(lngRow, 22), "HH:MM")
'            strOraConsegna = Left(strOraConsegna, 2) & ":" & Right(strOraConsegna, 2)
'            strOraInizioCarico = Format(XA(lngRow, 10), "HH:MM")
'            strOraInizioCarico = Left(strOraInizioCarico, 2) & ":" & Right(strOraInizioCarico, 2)
'            strFineCarico = IIf(IsNull(XA(lngRow, 25)), vbNullString, XA(lngRow, 25))
'            blnFineCarico = (strFineCarico = "S")
'            If IsNumeric(XA(lngRow, 24)) Then
'                lngKM = XA(lngRow, 24)
'            Else
'                lngKM = 0
'            End If
'
'            With mobjOutFile
'                .ClearFields
'                .Item("1").value = strIDEnteEmittente
'                .Item("2").value = strProgrInvio
'                .Item("3").value = strDataInvio
'                .Item("4").value = strOraInvio
''                .Item("5").value = Format(ViaggiModificati, "0000")
'                .Item("6").value = IIf(IsNull(XA(lngRow, 19)), vbNullString, XA(lngRow, 19))    'Riferimento esterno 1
'                If Not IsNull(XA(lngRow, 34)) Then
'                    .Item("7").value = XA(lngRow, 34)                                           'Partizione
'                End If
'                .Item("8").value = Mid(strDestinazione, 3)
'                .Item("9").value = 0
'                .Item("10").value = 0
'                .Item("11").value = 0
'                .Item("12").value = strProdotto
'                .Item("13").value = strIDDeposito
'                .Item("14").value = IIf(IsNull(XA(lngRow, 29)), 0, XA(lngRow, 29))              ' Q.tà concordata ordine
'                .Item("16").value = strDataViaggio
''                .Item("17").value = strDataViaggio
''                .Item("18").value = strOraConsegna
'                .Item("19").value = IIf(IsNull(XA(lngRow, 30)), 0, XA(lngRow, 30))              ' Q.tà caricata
'                .Item("20").value = IIf(IsNull(XA(lngRow, 31)), 0, XA(lngRow, 31))              ' Q.tà scaricata
'                .Item("21").value = strDataViaggio
'                If XA(lngRow, 37) = "S" Then
'                    .Item("22").value = strOraInizioCarico
'                End If
'                .Item("17").value = Format(XA(lngRow, 36), "yyyymmdd")
'                If blnFineCarico Then
''                    .Item("17").value = vbNullString
'                    .Item("18").value = strOraConsegna
'                    .Item("23").value = Format(XA(lngRow, 36), "yyyymmdd")
'                    .Item("24").value = strOraConsegna
'                Else
'                    .Item("18").value = strOraConsegna
'                    .Item("23").value = vbNullString
'                    .Item("24").value = vbNullString
'                End If
'                .Item("25").value = varBadge
'                .Item("26").value = IIf(IsNull(XA(lngRow, 8)), vbNullString, XA(lngRow, 8))     ' Nr. sequenza viaggio
'                .Item("27").value = strCartaIdentitaAutista1
'                .Item("28").value = lngKM
'                .Item("29").value = lngKMCepsa
'                .Item("30").value = Replace(Format(varCUCarga, "000.00"), ",", ".")
'                .Item("31").value = Replace(Format(varCUReparto, "000.00"), ",", ".")
'                .Item("32").value = Replace(Format(varCUTotale, "000.00"), ",", ".")
'                .Item("33").value = lngProgrConsegna                                            ' Nr. sequenza consegna
'                .Item("34").value = IIf(IsNull(XA(lngRow, 26)), vbNullString, XA(lngRow, 26))   ' Note
'                .Item("35").value = IIf(IsNull(XA(lngRow, 27)), vbNullString, XA(lngRow, 27))   ' Note estese
'                If Not .WriteLine(True) Then
'                    intRowError = OutputError
'                    GoTo LogResult
'                End If
'            End With
'
'            '' --- Modifico lo stato del viaggio
'            'With VVT
'            '    .DS.Fields("0").FindValue = strIDViaggio
'            '    If Not .Find Then
'            '        intRowError = NoTripID
'            '        GoTo LogResult
'            '    End If
'            '    If Not .SetFlagTrasmissione("T") Then
'            '        intRowError = NoSetState
'            '        GoTo LogResult
'            '    End If
'            'End With
'            ' ---
'        Else
'            blnNeedTrasmission = False
'        End If
'LogResult:
'        ' --- Scrivo il risultato dell'elaborazione
'        If blnNeedTrasmission Then
'            If intRowError = NoError Then
'                lngCountOK = lngCountOK + 1
'            Else
'                'If (intRowError = ZonaEsclusaExport) Or _
'                '   (intRowError = OrderNotActive) Then
'                '    lngCountExcluded = lngCountExcluded + 1
'                'Else
'                '    lngCountFail = lngCountFail + 1
'                'End If
'                lngCountFail = lngCountFail + 1
'            End If
'            lngCountToGo = lngCountToGo - 1
'            lblLog(2) = lngCountOK
'            lblLog(4) = lngCountFail
'            lblLog(5) = lngCountToGo
''            lblLog(7) = lngCountExcluded
'            Select Case intRowError
'                Case NoError
'                    strErrorMessage = "OK"
'                Case NoAutista
'                    strErrorMessage = "Primo autista non trovato"
'                Case OutputError
'                    strErrorMessage = "Errore durante la scrittura sul file"
'                Case noUT
'                    strErrorMessage = "Unità di trasporto non trovata"
'                Case NoDestinazione
'                    strErrorMessage = "Destinazione non trovata"
'                Case RowsNotFound
'                    strErrorMessage = "Errore durante la lettura delle consegne"
'                Case NoSetState
'                    strErrorMessage = "Impossibile cambiare lo stato del viaggio"
'                Case NoOrderID
'                    strErrorMessage = "Ordine non trovato"
'                Case NoTipoVendita
'                    strErrorMessage = "Codice Tipo Vendita non trovato"
'                Case NoDestinazione
'                    strErrorMessage = "Codice Destinazione non trovato"
'                Case NoDeposito
'                    strErrorMessage = "Codice Deposito non trovato"
'                Case NoEnteEmittente
'                    strErrorMessage = "Ente Emittente non trovato"
'                Case InvalidCode
'                    strErrorMessage = "Badge non trovato"
'            End Select
'        Else
'            strErrorMessage = "Record da non trasmettere"
'        End If
'        XALog.value(lngRow, 3) = strErrorMessage
'        With tdbLog
'            .Refresh
'            .Bookmark = lngRow
'        End With
'        TDBGrid1.Bookmark = lngRow
'        DoEvents
'    Next lngRow
'
'    GoTo EndProc
'ErrHandler:
'    Dim strMessage As String
'    Dim strInfo As String
'    Dim lngNativeCode As Long
'
'    Select Case ErrCode
'        Case FileNotFound
'            strMessage = "Impossibile aprire il file."
'            strInfo = "Nome del file: " & strFileName
'        Case ParameterNotFound
'            strMessage = "Parametri di sistema mancanti."
'            strInfo = "Default mancanti: Prodotti"
'        Case NoEnteEmittente
'            strMessage = "Ente Emittente non trovato."
'            strInfo = "Ente emittente: " & strIDEnteEmittente
'        Case FasciaTariffariaError
'            strMessage = "Errore nel calcolo dei coefficienti di utilizzo"
'        Case InvalidCode
'            strMessage = "Impossibile leggere il progressivo dell'invio"
'        Case NoError
'            Select Case Err.Number
'                Case Else
'                    ErrCode = UnknownError
'                    strMessage = "Errore Sconosciuto."
'                    strInfo = Err.Description
'                    lngNativeCode = Err.Number
'            End Select
'    End Select
'    DataClass.DS.DA.EndTrans ETROLLBACK
'    Set C = ErrH.SetError(ErrCode, "frmVVTSnapshotExport", _
'                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
'    DisplayError ErrH.LastError
'EndProc:
'    mobjOutFile.CloseFile
'
'    If lngCountOK > 0 Then
'        strProgrInvio = Format(varCounter, String(10, "0"))
'        strFileName = strFileName & "\GES_" & strProgrInvio & ".DAT"
'        intFileFromNum = FreeFile
'        Open TEMP_FILE For Input Access Read As #intFileFromNum
'        intFileToNum = FreeFile
'        Open strFileName For Output Access Write As #intFileToNum
'        Do While Not EOF(intFileFromNum)
'            Line Input #intFileFromNum, strLine
'            Mid(strLine, 31, 4) = Format(lngCountOK, "0000")
'            Mid(strLine, 5, 10) = strProgrInvio
'            Print #intFileToNum, strLine
'        Loop
'        Close #intFileFromNum
'        Close #intFileToNum
'    End If
'    chkTransAll.value = vbUnchecked
'    Me.MousePointer = vbDefault
'    Me.Enabled = True
'    Exit Sub
'End Sub

