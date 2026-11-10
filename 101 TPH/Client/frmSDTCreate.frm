VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSDTCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAZIONE STIPENDI"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9270
   Begin TrueDBGrid60.TDBGrid tdbLog 
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "frmSDTCreate.frx":0000
      TabIndex        =   25
      Top             =   2160
      Width           =   9015
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   9015
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Autisti elaborati:"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   7560
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Autisti:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   2055
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   8280
         Picture         =   "frmSDTCreate.frx":3120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   5640
         Picture         =   "frmSDTCreate.frx":3222
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2520
         Picture         =   "frmSDTCreate.frx":3324
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
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
         Left            =   4320
         TabIndex        =   4
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Tag             =   "82"
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField11 
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Tag             =   "80"
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Tag             =   "81"
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataEmissione 
         Height          =   255
         Left            =   6960
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Emissione"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   8760
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   4560
         Tag             =   "81"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vettore"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblField10 
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
         Left            =   1800
         TabIndex        =   27
         Tag             =   "81"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   4560
         Tag             =   "80"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Autista"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblField8 
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
         Left            =   1800
         TabIndex        =   26
         Tag             =   "80"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblField3 
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
         Left            =   1800
         TabIndex        =   15
         Tag             =   "82"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Base"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   4560
         Tag             =   "82"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   6120
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3000
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   5160
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   2400
         Picture         =   "frmSDTCreate.frx":3426
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Esegui"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmSDTCreate.frx":3528
         TabIndex        =   17
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmSDTCreate.frx":362A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Seleziona Autisti"
         Height          =   375
         Left            =   120
         Picture         =   "frmSDTCreate.frx":372C
         TabIndex        =   16
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
         Index           =   2
         Left            =   7560
         Picture         =   "frmSDTCreate.frx":382E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmSDTCreate.frx":3930
         TabIndex        =   18
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6030
      Width           =   9270
      _ExtentX        =   16351
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
            Object.Width           =   10901
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17.57"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "31/05/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmSDTCreate.frx":3A32
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSDTCreate.frx":3B44
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSDTCreate.frx":3C56
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSDTCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 12 '11
Private Const COL_RESULT_LOG = 2
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Public Loading As Boolean

Public DataClass As CViaggioSelezione
Public XA As XArrayDB
Public State As FormState

Private Const CONTRATTOID_AGIP = "1"
Private Const CONTRATTOID_CEPSA = "3"


Private Function RowToPPD(RowIndex As Long, _
                          IDDettaglio As Long, _
                          PPD As CPrefatturaDettaglio, _
                          IDTestata As String, _
                          IDZonaI As String, _
                          IDFasciaTariffariaI As String, _
                          TariffaI As Single, _
                          ImportoI As Single, _
                          IDValutaI As String, _
                          TipoPercorsoI As String, _
                          KmRaggioInterni As Long, _
                          RiferimentoEsterno As String, _
                          QtaScaricata As Single, _
                          AnomaliaI As AnomaliePrefattura, _
                          QtaScaricataKg As Single, _
                          Optional Pedaggio As Single, _
                          Optional ServOperario As Single, _
                          Optional ServBuques As Single, _
                          Optional ServOreSeServBuquesrvSpeciali As Single, _
                          Optional ServServSpeciali As Single) As Errors
    On Error GoTo ErrHandler
    
    Dim intError As Errors
    Dim intCols As Integer
    Dim i As Integer
    
    intError = NoError
    
    intCols = UBound(DataClass.DS.List, 1)
    ReDim varrow(intCols, 0) As Variant
    With DataClass.DS
        For i = 0 To intCols
            varrow(i, 0) = .ListItem(i, RowIndex)
        Next i
    End With
    
    With PPD.DS
        .Fields("0").NewValue = IDTestata
        .Fields("1").NewValue = IDDettaglio
        .Fields("2").NewValue = varrow(4, 0)
        .Fields("3").NewValue = varrow(5, 0)
        .Fields("4").NewValue = varrow(6, 0)
        .Fields("5").NewValue = varrow(7, 0)
        .Fields("6").NewValue = RiferimentoEsterno 'varrow(8, 0)
'        .Fields("7").NewValue = varrow(9, 0)
        .Fields("8").NewValue = varrow(10, 0)
        If varrow(11, 0) = 0 Then
            .Fields("9").NewValue = varrow(35, 0)
        Else
            .Fields("9").NewValue = varrow(11, 0)
        End If
        .Fields("10").NewValue = varrow(12, 0)
        .Fields("11").NewValue = varrow(13, 0)
        .Fields("12").NewValue = varrow(0, 0)
        .Fields("13").NewValue = varrow(14, 0)
        .Fields("14").NewValue = varrow(15, 0)
        ' Data consegna
        If Format(varrow(16, 0), "dd/mm/yyyy") = "01/01/1900" Then
            .Fields("15").NewValue = varrow(39, 0)
       Else
            .Fields("15").NewValue = varrow(16, 0)
        End If
        .Fields("16").NewValue = varrow(17, 0)
        .Fields("17").NewValue = varrow(18, 0)
        .Fields("18").NewValue = varrow(19, 0)
        .Fields("19").NewValue = varrow(20, 0)
        .Fields("20").NewValue = varrow(21, 0)
        .Fields("21").NewValue = varrow(22, 0)
        .Fields("22").NewValue = varrow(23, 0)
        .Fields("23").NewValue = varrow(24, 0)
        .Fields("24").NewValue = varrow(25, 0)
        .Fields("25").NewValue = varrow(26, 0)
        .Fields("26").NewValue = varrow(2, 0)
        .Fields("27").NewValue = varrow(27, 0)
        .Fields("28").NewValue = varrow(28, 0)
        .Fields("29").NewValue = QtaScaricataKg 'varrow(29, 0)
        .Fields("30").NewValue = QtaScaricata 'varrow(3, 0)
        .Fields("31").NewValue = varrow(30, 0)
        .Fields("32").NewValue = varrow(31, 0)
        
        .Fields("55").NewValue = varrow(34, 0)
        .Fields("55").Updated = True
        
        ' --- Valorizzazione interna
        .Fields("41").NewValue = IDZonaI
        .Fields("42").NewValue = IDFasciaTariffariaI
        .Fields("43").NewValue = TariffaI
        .Fields("44").NewValue = ImportoI
        .Fields("45").NewValue = ""
        .Fields("46").NewValue = IDValutaI
        .Fields("47").NewValue = KmRaggioInterni 'varRow(33, 0)
        If Not IsMissing(Pedaggio) Then
            .Field("Importo Pedaggio Interno").NewValue = Pedaggio
            .Field("Importo Pedaggio Interno").Updated = True
        End If
        If Not IsMissing(ServOperario) Then
            .Field("Importo Servizio Operario Interno").NewValue = ServOperario
            .Field("Importo Servizio Operario Interno").Updated = True
        End If
        If Not IsMissing(ServBuques) Then
            .Field("Importo Servizio Buques Interno").NewValue = ServBuques
            .Field("Importo Servizio Buques Interno").Updated = True
        End If
        If Not IsMissing(ServOreSeServBuquesrvSpeciali) Then
            .Field("Ore Servizi Speciali Interno").NewValue = ServOreSeServBuquesrvSpeciali
            .Field("Ore Servizi Speciali Interno").Updated = True
        End If
        If Not IsMissing(ServServSpeciali) Then
            .Field("Importo Servizi Speciali Interno").NewValue = ServServSpeciali
            .Field("Importo Servizi Speciali Interno").Updated = True
        End If
        
        .Fields("48").NewValue = TipoPercorsoI
        .Fields("50").NewValue = .Fields("50").NewValue Or AnomaliaI
        ' ---
        
        If IsDate(varrow(44, 0)) Then
            .Field("Data Grado Servizio").NewValue = varrow(44, 0)
            .Field("Data Grado Servizio").Updated = True
        End If
        
        For i = 0 To 32
            .Fields(CStr(i)).Updated = True
        Next i
        For i = 41 To 50
            .Fields(CStr(i)).Updated = True
        Next i
    End With
    RowToPPD = intError
    Exit Function
ErrHandler:
    If intError = NoError Then
        intError = UnknownError
    End If
    RowToPPD = intError
End Function

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            DataClass.DS.DisplayMessage vbNullString, 0
            If SelectRecords Then
                cmdExport.Enabled = True
            Else
                cmdExport.Enabled = False
            End If
        Case fcExport
            CreateRecords
            cmdExport.Enabled = False
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub CreateRecords()
    On Error GoTo ErrHandler
    
    Dim varDataFrom As Variant
    Dim varDataTo As Variant
    Dim varDataEmissione As Variant
    
    Dim strIDAutista As String
    Dim strError As String
    Dim strErrorMessage As String
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    
    Dim intRowErrorLog As Errors
    
    Dim SDT As CStipendioTestata
    
    varDataFrom = mskDateFrom
    varDataTo = mskDateTo
    varDataEmissione = mskDataEmissione
    
    If Not IsDate(varDataFrom) Then
        strError = "Data iniziale non valida"
        GoTo ErrHandler
    End If
    If Not IsDate(varDataTo) Then
        strError = "Data finale non valida"
        GoTo ErrHandler
    End If
    If Not IsDate(varDataEmissione) Then
        strError = "Data emissione non valida"
        GoTo ErrHandler
    End If
    
    If DateDiff("d", varDataTo, varDataFrom) > 0 Then
        strError = "La data finale non può precedere la data iniziale"
        GoTo ErrHandler
    End If
    If DateDiff("d", varDataFrom, varDataEmissione) < 0 Then
        strError = "La data emissione non può precedere la data iniziale"
        GoTo ErrHandler
    End If
    
    Set SDT = New CStipendioTestata
    With SDT.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    
    StatusBar1.Panels(2).Text = "Elaborazione in corso... "
    
    lngNumRows = XA.Count(1) - 1
    lngCountToGo = lngNumRows + 1
    
    For lngRow = 0 To lngNumRows
        strIDAutista = XA.value(lngRow, 0)
        strError = vbNullString
        intRowErrorLog = NoError
        If Not SDT.CreateStipendio(strIDAutista, varDataEmissione, varDataFrom, varDataTo, strError) Then
            intRowErrorLog = UnknownError
            GoTo LogResult
        End If
        
LogResult:
        If intRowErrorLog <> NoError Then
            lngCountFail = lngCountFail + 1
        Else
            lngCountOK = lngCountOK + 1
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        
        Select Case intRowErrorLog
            Case NoError
                strErrorMessage = "OK"
            Case UnknownError
                strErrorMessage = strError
        End Select
        
        XA.value(lngRow, COL_RESULT_LOG) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        DoEvents
        ' ---
        
    Next lngRow
    ' ---
    
    GoTo EndProc
    
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    DataClass.DS.DisplayMessage strError, 9
EndProc:
    StatusBar1.Panels(2).Text = vbNullString
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub


Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim j As Integer
    
    Me.MousePointer = vbHourglass
    With DataClass.DS
        If .Fields.Validate Then
            If Not .ScanClientCond Then
                GoTo ErrHandler
            End If
            If Not DataClass.GetSelectionCreateSDT Then
                GoTo ErrHandler
            End If
            DisplaySelection
            GoTo EndProc
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    GoTo EndProc
ErrHandler:
    Me.MousePointer = vbDefault
    SelectRecords = False
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function


Private Sub DisplaySelection()
    Dim varLog As Variant
    Dim i As Long
    Dim lngRows As Long
    
    lngRows = DataClass.DS.HowMany

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
    If lngRows > 0 Then
        ReDim varLog(0 To 2, 0 To lngRows - 1)
        For i = 0 To lngRows - 1
            varLog(0, i) = DataClass.DS.ListItem(0, i)
            varLog(1, i) = DataClass.DS.ListItem(1, i)
        Next i
        RefreshXArray XA, tdbLog, varLog
        tdbLog.bookmark = 0
    End If
End Sub



Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Dalla data"
        Case 1
            Set M = mskDateTo
            strTitle = "Alla data"
        Case 2
            Set M = mskDataEmissione
            strTitle = "Data emissione"
    End Select
    GetCalendarDate M, strTitle

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim D As Variant
    Dim i As Long
    
    Loading = True
    Set XA = New XArrayDB
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .Fields("81").Required = True
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set DataClass = Nothing
    Set XA = Nothing
    Set frmSDTCreate = Nothing
End Sub




