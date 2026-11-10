VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2PPT0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GESTIONE PREFATTURAZIONE"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   20
      Top             =   4920
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   6000
         Picture         =   "OIL2PPT0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   795
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
         Left            =   7680
         Picture         =   "OIL2PPT0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2PPT0.frx":0204
         TabIndex        =   17
         Tag             =   "EXIT"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Gen. Fattura"
         Height          =   375
         Left            =   5880
         Picture         =   "OIL2PPT0.frx":0306
         TabIndex        =   35
         Tag             =   "EXPORT"
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   7320
         Picture         =   "OIL2PPT0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Riga"
         Height          =   375
         Left            =   7200
         Picture         =   "OIL2PPT0.frx":050A
         TabIndex        =   15
         Tag             =   "EDITLINK"
         Top             =   360
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3960
         Picture         =   "OIL2PPT0.frx":060C
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F12 - Elimina Riga"
         Height          =   375
         Left            =   3840
         Picture         =   "OIL2PPT0.frx":0B3E
         TabIndex        =   28
         Tag             =   "CUSTOM2"
         Top             =   720
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   2280
         Picture         =   "OIL2PPT0.frx":0C40
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F11 - Elimina"
         Height          =   375
         Left            =   2160
         Picture         =   "OIL2PPT0.frx":1172
         TabIndex        =   16
         Tag             =   "CUSTOM1"
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   240
         Picture         =   "OIL2PPT0.frx":1274
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F9 - Importe Cepsa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2PPT0.frx":1376
         TabIndex        =   32
         Tag             =   "FIND"
         Top             =   720
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5280
         Picture         =   "OIL2PPT0.frx":1478
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F6 - Verifica"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2PPT0.frx":157A
         TabIndex        =   30
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   3600
         Picture         =   "OIL2PPT0.frx":167C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2PPT0.frx":177E
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
         Left            =   1920
         Picture         =   "OIL2PPT0.frx":1880
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2PPT0.frx":1982
         TabIndex        =   13
         Tag             =   "UNDO"
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
         Left            =   240
         Picture         =   "OIL2PPT0.frx":1A84
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2PPT0.frx":1B86
         TabIndex        =   12
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6180
      Width           =   9375
      _ExtentX        =   16536
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
            Object.Width           =   11086
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:19"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "08/03/2018"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "6"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "4"
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Tag             =   "0"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Tag             =   "3"
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "OIL2PPT0.frx":1C88
         TabIndex        =   11
         Top             =   2100
         Width           =   9135
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Tag             =   "1"
         Top             =   600
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
         Left            =   5400
         TabIndex        =   6
         Tag             =   "2"
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8520
         Top             =   240
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
               Picture         =   "OIL2PPT0.frx":8332
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2PPT0.frx":8444
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2PPT0.frx":8556
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Tag             =   "5"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   3840
         Tag             =   "6"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Stato"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   2280
         TabIndex        =   34
         Tag             =   "5"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   5520
         Tag             =   "5"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Prefattura"
         Height          =   255
         Index           =   27
         Left            =   3960
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   7560
         Tag             =   "4"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Data Inizio Periodo"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3000
         Tag             =   "1"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data &Fine Periodo"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   6720
         Tag             =   "2"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5520
         Tag             =   "3"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3000
         Tag             =   "0"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Numero"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2PPT0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 15

Public GridArrayDetail As ArrayGridCls
Public WithEvents DataClass As CPrefattura
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CPrefatturaDettaglio
Private PPE As CPrefatturaSocSelezione
Private PPD As CPrefatturaDettaglio
Private AXX As CParametro
Private VVT As CViaggio
Private VVP As CViaggioProdotto
Private APR As CProdotto
Private GP1 As CPrefatturaSoc

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object

Private Sub CheckPrefatture(Optional ForzaImporto As Boolean = False, _
                            Optional ForzaQuantita As Boolean = False)
    On Error GoTo ErrHandler
    
    Dim lngCount As Long
    Dim i As Long
    Dim j As Long
    Dim strRifOrdineEsterno As String
    Dim blnPropostaMancante As Boolean
    Dim strDataConsegna As String
    Dim strIDProdotto As String
    Dim strIDProdottoSoc As String
    Dim strIDUdT As String
    Dim strIDDestinazione As String
    Dim strIDValutaDefault As String
    Dim varZona As Variant
    Dim varTariffa As Variant
    Dim varImporto As Variant
    Dim varFlagImpForzato As Variant
    Dim varKmRaggio As Variant
    Dim varNomeFile As Variant
    Dim varAnno As Variant
    Dim varValuta As Variant
    Dim varBollettaScarico As String
    Dim varRifEst2 As Variant
    Dim varAnomalia As Variant
    Dim varQuantita As Variant
    Dim varIDPrefattura As Variant
    Dim varIDRiga As Variant
    Dim varDetailList As Variant
    Dim varListItem As Variant
    Dim strIDVVT As String
    Dim strIDVVC As String
    Dim strIDVVP As String
    Dim blnConcErr As Boolean
    Dim intRowErrorLog As Errors
    Dim ErrCode As Errors
    Dim intRowErrorAnomaly As Errors
    Dim intAnomalia As AnomaliePrefattura
    Dim intAnomaliaTot As AnomaliePrefattura
    Dim C As CError
    
    StatusBar1.Panels(2).Text = "Elaborazione in corso..."
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        strIDValutaDefault = Trim$(.Fields("17").Value)
    End With
    
    If Len(strIDValutaDefault) = 0 Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    varIDPrefattura = DataClass.DS.Fields("0").Value
    
    With DetailClass.DS
        varDetailList = .List
        lngCount = UBound(DetailClass.DS.List, 2)
        For i = 0 To lngCount
            TDBGrid1.bookmark = i
            intAnomaliaTot = apNoAnomalie
            
            varIDRiga = TDBGrid1.Columns(0).Text
            
            .DA.BeginTransaction
            strRifOrdineEsterno = varDetailList(3, i)
            strIDDestinazione = varDetailList(12, i)
            strDataConsegna = varDetailList(5, i)
            strIDProdotto = varDetailList(8, i)
            strIDUdT = varDetailList(6, i)
            strIDVVT = varDetailList(4, i)
            strIDVVC = varDetailList(22, i)
            strIDVVP = varDetailList(23, i)
            
            ' --- Valorizzazione esterna
            blnPropostaMancante = False
            PPE.DS.Fields.ResetFields
        
            strIDProdottoSoc = vbNullString
            APR.DS.Fields.ResetFields                   ' Codice Prodotto
            APR.DS.Fields("0").FindValue = strIDProdotto
            APR.Find
            strIDProdottoSoc = APR.DS.MultiSoc.DS.GetFirmCode("001", 3)
            If strIDProdottoSoc = vbNullString Then
                ErrCode = NoProdotto
                GoTo ErrHandler
            End If
            
            ' --- Se sull'ordine manca il numero di riferimento esterno
            ' --- ricerco per Data, prodotto, destinazione, UdT
            With PPE.DS
                If Len(strRifOrdineEsterno) = 0 Then
                    ' Destinazione
                    .Fields("10").FindValue = strIDDestinazione
                    ' Data consegna
                    .Fields("14").FindValue = strDataConsegna
                    ' Prodotto
                    .Fields("9").FindValue = strIDProdottoSoc 'strIDProdotto 'PPD.DS.Fields("26").NewValue
                    ' Unità di trasporto
                    .Fields("12").FindValue = strIDUdT 'PPD.DS.Fields("34").NewValue
                    If Not PPE.GetSelection(True) Then
                        blnPropostaMancante = True
                    End If
                    If IsEmpty(.List) Then
                        blnPropostaMancante = True
                    Else
                        strRifOrdineEsterno = .List(4, 0)
                    End If
                Else
                    ' --- Eseguo la ricerca del record di prefatturazione
                    ' --- esterno per Codice Ordine esterno e Prodotto
                    .Fields("13").FindValue = strRifOrdineEsterno
                    .Fields("9").FindValue = strIDProdottoSoc 'strIDProdotto ' PPD.DS.Fields("26").NewValue
                    If Not PPE.GetSelection(True) Then
                        blnPropostaMancante = True
                    End If
                    If IsEmpty(.List) Then
                        blnPropostaMancante = True
                    End If
                    ' ---
                End If
                varListItem = .List
                If Not blnPropostaMancante Then
                    varZona = varListItem(2, 0)
                    varTariffa = varListItem(9, 0)
                    varImporto = varListItem(10, 0)
                    varFlagImpForzato = IIf(varListItem(16, 0) <> 0, "S", "N")
                    varKmRaggio = varListItem(15, 0)
                    varNomeFile = varListItem(1, 0)
                    varAnno = varListItem(0, 0)
                    If Not IsNull(varListItem(6, 0)) Then
                        varBollettaScarico = varListItem(6, 0)
                    Else
                        varBollettaScarico = vbNullString
                    End If
                    If Not IsNull(varListItem(8, 0)) Then
                        varRifEst2 = varListItem(8, 0)
                    Else
                        varRifEst2 = vbNullString
                    End If
                    varQuantita = varListItem(7, 0)
                Else
                    ' --- Se non ho trovato la prefattura esterna segnalo
                    ' --- l'anomalia
                    intAnomaliaTot = intAnomaliaTot Or apPropostaNonTrovata
                    varZona = Empty
                    varTariffa = Empty
                    varImporto = Empty
                    varFlagImpForzato = Empty
                    varKmRaggio = Empty
                    varNomeFile = Empty
                    varAnno = Empty
                    varBollettaScarico = Empty
                    varRifEst2 = Empty
                    varQuantita = Empty
                    ' ---
                End If
            End With
    
            varValuta = strIDValutaDefault
            ' --- Fine valorizzazione esterna
            
            ' --- Controllo le discordanze tra valorizzazione esterna
            ' --- e valorizzazione interna
            If Not blnPropostaMancante Then
                intAnomalia = apNoAnomalie
                intRowErrorAnomaly = AnomalyCheck(intAnomalia, i)
                If intAnomalia <> apNoAnomalie Then
                    intAnomaliaTot = intAnomaliaTot Or intAnomalia
                End If
            End If
            ' ---
            
            ' --- Aggiornamento della riga di prefattura
            If Not ForzaImporto Then
                If Not DetailClass.AddExtData(varZona, varTariffa, varImporto, varFlagImpForzato, _
                                              varKmRaggio, varNomeFile, varAnno, varValuta, _
                                              varIDPrefattura, varIDRiga, varRifEst2, intAnomaliaTot, , varQuantita) Then
                    GoTo ErrHandler
                End If
                If ForzaQuantita Then
                    If IsNumeric(varQuantita) Then
                        If varQuantita > 0 Then
                            If Not VVP.UpdateQtyDelivered(strIDVVT, strIDVVC, strIDVVP, CSng(varQuantita)) Then
                                GoTo ErrHandler
                            End If
                        End If
                    End If
                End If
            Else
                If Not DetailClass.AddExtData(varZona, varTariffa, varImporto, varFlagImpForzato, _
                                              varKmRaggio, varNomeFile, varAnno, varValuta, _
                                              varIDPrefattura, varIDRiga, varRifEst2, intAnomaliaTot, True, _
                                              varQuantita) Then
                    GoTo ErrHandler
                End If
                If IsNumeric(varQuantita) Then
                    If Not VVP.UpdateQtyDelivered(strIDVVT, strIDVVC, strIDVVP, CSng(varQuantita)) Then
                        GoTo ErrHandler
                    End If
                End If
            End If
            
            .DA.EndTrans ETCOMMIT
            ' ---
            
LogResult:
            DoEvents
            ' ---
        Next i
    End With
    GoTo EndProc
    
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    DetailClass.DS.DA.EndTrans ETROLLBACK
    Select Case ErrCode
        Case UpdateError
            strMessage = "Si è verificato un errore durante l'aggiornamento."
        Case NoProdotto
            strMessage = "Codice prodotto per società non trovato"
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Impossibile trovare la valuta di default"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "OIL2PPT0", _
                "CheckPrefatture", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    DisplayDetail
    StatusBar1.Panels(2).Text = vbNullString
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
    
End Sub


Private Sub CheckPrefattureCEPSA()
    On Error GoTo ErrHandler
    
    Dim lngCount As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lngNumProposte As Long
    Dim strRifOrdineEsterno As String
    Dim blnPropostaMancante As Boolean
    Dim strDataConsegna As String
    Dim strIDProdotto As String
    Dim strIDUdT As String
    Dim strIDDestinazione As String
    Dim strIDViaggio As String
    Dim strIDValutaDefault As String
    Dim varZona As Variant
    Dim varTariffa As Variant
    Dim varImporto As Variant
    Dim varPedaggio As Variant
    Dim varServOperarioBunker As Variant
    Dim varServBuquesPuerto As Variant
    Dim varServOreServSpeciali As Variant
    Dim varServImpServSpeciali As Variant
    Dim varFlagImpForzato As Variant
    Dim varKmRaggio As Variant
    Dim varNomeFile As Variant
    Dim varAnno As Variant
    Dim varValuta As Variant
    Dim varBollettaScarico As String
    Dim varRifEst2 As Variant
    Dim varAnomalia As Variant
    Dim varQuantita As Variant
    Dim varQuantitaKG As Single
    Dim varIDPrefattura As Variant
    Dim varIDRiga As Variant
    Dim varDetailList As Variant
    Dim varListItem As Variant
    Dim blnConcErr As Boolean
    Dim intRowErrorLog As Errors
    Dim ErrCode As Errors
    Dim intRowErrorAnomaly As Errors
    Dim intAnomalia As AnomaliePrefattura
    Dim intAnomaliaTot As AnomaliePrefattura
    Dim C As CError
    
    Dim varServImpServEsenti As Variant
    
    ' -- DP 10/01/09
    Dim strIDADE As String
    
    StatusBar1.Panels(2).Text = "Elaborazione in corso..."
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = MissingDefaults
            GoTo ErrHandler
        End If
        strIDValutaDefault = Trim$(.Fields("17").Value)
    End With
    
    If Len(strIDValutaDefault) = 0 Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    varIDPrefattura = DataClass.DS.Fields("0").Value
    
    With DetailClass.DS
        .Fields("0").FindValue = varIDPrefattura
        DetailClass.GetListCheck True
        varDetailList = .List
        lngCount = UBound(DetailClass.DS.List, 2)
        For i = 0 To lngCount
            TDBGrid1.bookmark = i
            intAnomaliaTot = apNoAnomalie
            
            varIDRiga = TDBGrid1.Columns(0).Text
            
            .DA.BeginTransaction
            
            strRifOrdineEsterno = varDetailList(3, i)
            
            strIDDestinazione = varDetailList(12, i)
            strDataConsegna = varDetailList(5, i)
            strIDViaggio = varDetailList(4, i)
            strIDProdotto = varDetailList(8, i)
            strIDUdT = varDetailList(6, i)
            
            ' -- DP 10/01/09
            strIDADE = varDetailList(13, i)
            
            ' --- Valorizzazione esterna
            blnPropostaMancante = False
            PPE.DS.Fields.ResetFields
    
            ' --- Se sull'ordine manca il numero di riferimento esterno
            ' --- ricerco per Data, prodotto, destinazione, UdT
            With PPE.DS
                If Len(strRifOrdineEsterno) = 0 Then
                    ' Destinazione
                    .Fields("10").FindValue = strIDDestinazione
                    ' Data consegna
                    .Fields("14").FindValue = strDataConsegna
                    ' Prodotto
                    .Fields("9").FindValue = strIDProdotto
                    ' Unità di trasporto
                    .Fields("12").FindValue = strIDUdT
                    
                    ' -- DP 10/01/09
                    .Fields("1").FindValue = strIDADE
                    
                    If Not PPE.GetSelectionCheck(True) Then
                        blnPropostaMancante = True
                    End If
                    If IsEmpty(.List) Then
                        blnPropostaMancante = True
                    Else
                        strRifOrdineEsterno = .List(4, 0)
                    End If
                Else
                    ' --- Eseguo la ricerca del record di prefatturazione
                    ' --- esterno per Codice Ordine esterno
                    .Fields("13").FindValue = strRifOrdineEsterno
                    
                    ' -- DP 10/01/09
                    .Fields("1").FindValue = strIDADE
                    
                    If Not PPE.GetSelectionCheck(True) Then
                        blnPropostaMancante = True
                    End If
                    If IsEmpty(.List) Then
                        blnPropostaMancante = True
                    End If
                    ' ---
                End If
                varListItem = .List
                If Not blnPropostaMancante Then
                
                'JG 02/06/2017 Recorrer diferentes cfact y cdife
                    lngNumProposte = UBound(varListItem, 2) + 1
                    If lngNumProposte > 1 Then
                        For k = 1 To lngNumProposte - 1
                            varListItem(9, 0) = varListItem(9, 0) + varListItem(9, k)
                            varListItem(10, 0) = varListItem(10, 0) + varListItem(10, k)
                            varListItem(25, 0) = varListItem(25, 0) + varListItem(25, k)
                            varListItem(27, 0) = varListItem(27, 0) + varListItem(27, k)
                            varListItem(28, 0) = varListItem(28, 0) + varListItem(28, k)
                            varListItem(29, 0) = varListItem(29, 0) + varListItem(29, k)
                            varListItem(30, 0) = varListItem(30, 0) + varListItem(30, k)
                        Next k
                    End If
                    varZona = varListItem(2, 0)
                    If IsNull(varZona) Or Len(varZona) = 0 Then
                        varZona = "00"
                    End If
                    varTariffa = varListItem(9, 0)
                    varImporto = varListItem(10, 0)
                    varPedaggio = varListItem(25, 0)
                    varServOperarioBunker = varListItem(27, 0)
                    varServBuquesPuerto = varListItem(28, 0)
                    varServOreServSpeciali = varListItem(29, 0)
                    varServImpServSpeciali = varListItem(30, 0)
                    varFlagImpForzato = IIf(varListItem(16, 0) <> 0, "S", "N")
                    varKmRaggio = varListItem(15, 0)
                    varNomeFile = "" ' verificare varListItem(1, 0)
                    varAnno = varListItem(0, 0)
                    varBollettaScarico = "" 'varListItem(6, 0)
                    varRifEst2 = "" 'varListItem(8, 0)
                    varQuantita = varListItem(7, 0)
                    varQuantitaKG = varListItem(26, 0)
                    
                    '-- DP 20/08/08
                    varServImpServEsenti = varListItem(31, 0)   ' Importo Servizi Esenti
                Else
                    ' --- Se non ho trovato la prefattura esterna segnalo
                    ' --- l'anomalia
                    intAnomaliaTot = intAnomaliaTot Or apPropostaNonTrovata
                    intAnomaliaTot = intAnomaliaTot Or apPedaggio
                    varZona = Empty
                    varTariffa = Empty
                    varImporto = Empty
                    varPedaggio = Empty
                    varServOperarioBunker = 0
                    varServBuquesPuerto = 0
                    varServOreServSpeciali = 0
                    varServImpServSpeciali = 0
                    varFlagImpForzato = Empty
                    varKmRaggio = Empty
                    varNomeFile = Empty
                    varAnno = Empty
                    varBollettaScarico = Empty
                    varRifEst2 = Empty
                    varQuantita = Empty
                    varQuantitaKG = Empty
                    
                    '-- DP 20/08/08
                    varServImpServEsenti = 0
                    ' ---
                End If
            End With
    
            varValuta = strIDValutaDefault
            ' --- Fine valorizzazione esterna
            
            ' --- Controllo le discordanze tra valorizzazione esterna
            ' --- e valorizzazione interna
            If Not blnPropostaMancante Then
                intAnomalia = apNoAnomalie
                intRowErrorAnomaly = AnomalyCheckCEPSA(intAnomalia, i)
                If intAnomalia <> apNoAnomalie Then
                    intAnomaliaTot = intAnomaliaTot Or intAnomalia
                End If
            End If
            ' ---
            
            ' --- Aggiornamento della riga di prefattura
            If Not DetailClass.AddExtData(varZona, varTariffa, varImporto, varFlagImpForzato, _
                                          varKmRaggio, varNomeFile, varAnno, varValuta, _
                                          varIDPrefattura, varIDRiga, varRifEst2, intAnomaliaTot, _
                                          False, varQuantita, varQuantitaKG, varPedaggio, _
                                          varServOperarioBunker, varServBuquesPuerto, _
                                          varServOreServSpeciali, varServImpServSpeciali, _
                                          varServImpServEsenti) Then
                GoTo ErrHandler
            End If
            
            ' --- 8/2/14 Pone el flag de VVTFLST5 a T Transmitido
            
            With VVT
                .DS.Fields("0").FindValue = strIDViaggio
                If Not .Find Then
                    GoTo LogResult
                End If
                If Not .SetFlagTrasmissione("T") Then
                    GoTo LogResult
                End If
            End With

            
            .DA.EndTrans ETCOMMIT
            ' ---
            
LogResult:
            DoEvents
            ' ---
        Next i
    End With
    GoTo EndProc
    
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    DetailClass.DS.DA.EndTrans ETROLLBACK
    Select Case ErrCode
        Case UpdateError
            strMessage = "Si è verificato un errore durante l'aggiornamento."
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Impossibile trovare la valuta di default"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "OIL2PPT0", _
                "CheckPrefattureCEPSA", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    'JG Control Bloqueos DB
    DataClass.DS.DA.EndTrans ETCOMMIT
    DisplayDetail
    StatusBar1.Panels(2).Text = vbNullString
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
    
End Sub



Private Function AnomalyCheck(ByRef AnomalyCode As AnomaliePrefattura, _
                              RowIndex As Long) As Errors
    On Error GoTo ErrHandler
    
    Const TOLLERANZA_IMPORTO = 100
    
    Dim intError As Errors
    Dim varImpDiff As Variant
    Dim varPPE As Variant
    Dim varList As Variant
    
    intError = NoError
    
    varPPE = PPE.DS.List
    varList = DetailClass.DS.List
        ' --- Deposito
        If varList(13, RowIndex) <> varPPE(3, 0) Then
            AnomalyCode = apDatiDiscordanti
        ' --- Zona
        ElseIf varList(14, RowIndex) <> varPPE(2, 0) Then
            AnomalyCode = apDatiDiscordanti
        ' --- Data di consegna
        ElseIf varList(5, RowIndex) <> varPPE(5, 0) Then
            AnomalyCode = apDatiDiscordanti
        ' --- Codice destinazione
        ElseIf varList(12, RowIndex) <> varPPE(12, 0) Then
            AnomalyCode = apDatiDiscordanti
        ' --- Unità di trasporto
        ElseIf varList(6, RowIndex) <> varPPE(14, 0) Then
            AnomalyCode = apDatiDiscordanti
        ' --- Quantità scaricata
        ElseIf varList(9, RowIndex) <> varPPE(7, 0) Then
            AnomalyCode = AnomalyCode Or apDatiDiscordanti
        End If
        
        ' --- Distanza
        If CStr(varList(15, RowIndex)) <> varPPE(15, 0) Then
            AnomalyCode = AnomalyCode Or apDiffKm
        End If
        
        ' --- Importo
        varImpDiff = CStr(varList(11, RowIndex)) - varPPE(10, 0)
        If Abs(varImpDiff) > TOLLERANZA_IMPORTO Then
            AnomalyCode = AnomalyCode Or apDiffImporto
        End If
        
        ' --- Data grado servizio
        If IsDate(varList(24, RowIndex)) Then
            If IsDate(varPPE(18, 0)) Then
                If varList(24, RowIndex) <> varPPE(18, 0) Then
                    AnomalyCode = AnomalyCode Or apDataGradoServizio
                End If
            End If
        End If
        
        
'    End With
    AnomalyCheck = intError
    Exit Function
ErrHandler:
    If intError = NoError Then
        intError = UnknownError
    End If
    AnomalyCheck = intError

End Function

Private Function AnomalyCheckCEPSA(ByRef AnomalyCode As AnomaliePrefattura, _
                                   RowIndex As Long) As Errors
    On Error GoTo ErrHandler
    
    Const TOLLERANZA_IMPORTO = 0.2
    Const ID_VETTORE = "041883"
    
    Dim intError As Errors
    Dim varImpDiff As Variant
    Dim varPedDiff As Variant
    Dim varPPE As Variant
    Dim varList As Variant
    Dim strCodUdTDeposito As String
    Dim AUT As CUnitaDiTrasporto
    Dim blnCheckPedaggio As Boolean
    Dim blnNewViaggio As Boolean
    Dim i As Long
    Dim lngNumRows As Long
    
    Static sngImportoPedaggio As Single
    Static sngImportoPedaggioEsterno As Single
        
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    intError = NoError
    
    varPPE = PPE.DS.List
    varList = DetailClass.DS.List
    
    ' --- Data di consegna
    If varList(5, RowIndex) <> varPPE(5, 0) Then
        AnomalyCode = apDatiDiscordanti Or AnomalyCode
    End If
    ' --- Codice destinazione
    If varList(12, RowIndex) <> varPPE(12, 0) Then
        AnomalyCode = apCodiceDestinatario Or AnomalyCode
    End If
    
    ' --- Unità di trasporto
    If varList(6, RowIndex) <> varPPE(14, 0) Then
        AnomalyCode = apUnitàDiTrasporto Or AnomalyCode
    End If
    
    ' --- Distanza
    If CStr(varList(15, RowIndex)) <> varPPE(15, 0) Then
        AnomalyCode = AnomalyCode Or apDiffKm
    End If
    
    ' --- Importo
    For i = 1 To (UBound(varPPE, 2))
        varPPE(10, 0) = varPPE(10, 0) + varPPE(10, i)
        varPPE(27, 0) = varPPE(27, 0) + varPPE(27, i)
        varPPE(28, 0) = varPPE(28, 0) + varPPE(28, i)
        varPPE(29, 0) = varPPE(29, 0) + varPPE(29, i)
        varPPE(30, 0) = varPPE(30, 0) + varPPE(30, i)
    Next i
    varImpDiff = CStr(varList(11, RowIndex)) - varPPE(10, 0)
    If Abs(varImpDiff) > TOLLERANZA_IMPORTO Then
        AnomalyCode = AnomalyCode Or apDiffImporto
    End If
    
    ' --- Vettore
    If ID_VETTORE <> varPPE(18, 0) Then
        AnomalyCode = AnomalyCode Or apVettore
    End If
    
    ' --- Motrice
    If UCase(Left(varPPE(19, 0), 2)) = "ES" Then
        varPPE(19, 0) = Mid(varPPE(19, 0), 3)
    End If
    If UCase(Replace(varList(16, RowIndex), "-", "")) <> UCase(varPPE(19, 0)) Then
        AnomalyCode = apTargaMotrice Or AnomalyCode
    End If
    
    ' --- Rimorchio
    If UCase(Left(varPPE(20, 0), 2)) = "ES" Then
        varPPE(20, 0) = Mid(varPPE(20, 0), 3)
    End If
    If Not IsNull(varList(17, RowIndex)) Then
        If UCase(Replace(varList(17, RowIndex), "-", "")) <> UCase(varPPE(20, 0)) Then
            AnomalyCode = apTargaRimorchio Or AnomalyCode
        End If
    End If
    ' --- Codice UdT * Deposito
    AUT.DS.Fields("0").FindValue = varList(6, RowIndex)
    AUT.Find
    strCodUdTDeposito = AUT.DS.GetFirmCode(ID_SOC_CLH)
    If Len(strCodUdTDeposito) > 0 Then
        If strCodUdTDeposito <> UCase(varPPE(21, 0)) Then
            AnomalyCode = apVeicoloPerDeposito Or AnomalyCode
        End If
    End If
    
    ' --- NIF Autista
    If Len(varList(19, RowIndex)) > 0 Then
        If UCase(Left(varPPE(23, 0), 2)) = "ES" Then
            varPPE(23, 0) = Mid(varPPE(23, 0), 3)
        End If
        If UCase(Left(varList(19, RowIndex), 2)) = "ES" Then
            varList(19, RowIndex) = Mid(varList(19, RowIndex), 3)
        End If
        If UCase(Replace(Replace(varList(19, RowIndex), "-", ""), " ", "")) <> UCase(varPPE(23, 0)) Then
            AnomalyCode = apCodFiscAutista Or AnomalyCode
        End If
    End If
    
    ' --- Sequenza Viaggio
    If IsNumeric(varList(20, RowIndex)) Then
        varList(20, RowIndex) = CInt(varList(20, RowIndex))
    Else
        AnomalyCode = apProgressivoViaggio Or AnomalyCode
    End If
    If IsNumeric(varPPE(24, 0)) Then
        varPPE(24, 0) = CInt(varPPE(24, 0))
    Else
        AnomalyCode = apProgressivoViaggio Or AnomalyCode
    End If
    If varList(20, RowIndex) <> varPPE(24, 0) Then
        AnomalyCode = apProgressivoViaggio Or AnomalyCode
    End If
    
    ' --- Tariffa
    For i = 1 To UBound(varPPE, 2)
        varPPE(9, 0) = varPPE(9, 0) + varPPE(9, i)
    Next i
    If varList(10, RowIndex) <> varPPE(9, 0) Then
        AnomalyCode = apTariffa Or AnomalyCode
    End If
    
    ' --- Servizi Speciali
    If Abs(varList(23, RowIndex) - varPPE(27, 0)) > TOLLERANZA_IMPORTO Then
        AnomalyCode = apServOperarioBunker Or AnomalyCode
    End If
    If Abs(varList(24, RowIndex) - varPPE(28, 0)) > TOLLERANZA_IMPORTO Then
        AnomalyCode = apServBuquesPuerto Or AnomalyCode
    End If
    If Abs(varList(25, RowIndex) - varPPE(29, 0)) > TOLLERANZA_IMPORTO Then
        AnomalyCode = apServOreServSpeciali Or AnomalyCode
    End If
    If Abs(varList(26, RowIndex) - varPPE(30, 0)) > TOLLERANZA_IMPORTO Then
        AnomalyCode = apServImpServSpeciali Or AnomalyCode
    End If
    
    ' --- DP 20/08/08
    ' --- Servizi Esenti
    If varList(28, RowIndex) <> varPPE(31, 0) Then
        AnomalyCode = apServImpServEsenti Or AnomalyCode
    End If
    
    
    ' --- Pedaggio
    If RowIndex = 0 Then
        blnNewViaggio = True
    Else
        blnNewViaggio = varList(4, RowIndex) <> varList(4, RowIndex - 1)
    End If
    lngNumRows = UBound(varList, 2)
    
    If Not IsNumeric(varPPE(25, 0)) Then
        varPPE(25, 0) = 0
    End If
'    For i = 1 To UBound(varPPE, 2)
'        If Not IsNumeric(varPPE(25, i)) Then
'            varPPE(25, i) = 0
'        End If
'        varPPE(25, 0) = varPPE(25, 0) + varPPE(25, i)
'    Next i
    sngImportoPedaggio = varList(22, RowIndex)
'    i = RowIndex
'    If blnNewViaggio Then
'        sngImportoPedaggio = varList(22, RowIndex)
'        sngImportoPedaggioEsterno = 0
'        Do
'            If varList(4, i) = varList(4, i + 1) Then
'                sngImportoPedaggioEsterno = sngImportoPedaggioEsterno + GetPedaggio(varList(3, i), varList(12, i), varList(5, i), varList(8, i), varList(6, i))
'            End If
'            i = i + 1
'        Loop While (i < lngNumRows) And varList(4, i) = varList(4, RowIndex)
'    End If
    varPedDiff = sngImportoPedaggio - varPPE(25, 0)
    If (Abs(varPedDiff) > TOLLERANZA_IMPORTO) And (sngImportoPedaggio <> 0) Then
        AnomalyCode = apPedaggio Or AnomalyCode
    End If
    
    AnomalyCheckCEPSA = intError
    Exit Function
ErrHandler:
    If intError = NoError Then
        intError = UnknownError
    End If
    AnomalyCheckCEPSA = intError

End Function



Private Sub ForzaImporti()
    On Error GoTo ErrHandler
         
    Me.MousePointer = vbHourglass
    Me.Enabled = False

    If Not DataClass.SetImporti Then
        GoTo ErrHandler
    End If
    
    DisplayDetail
    GoTo EndProc

ErrHandler:

EndProc:
    Me.MousePointer = vbDefault
    Me.Enabled = True
    DisplayDetail
End Sub

Private Sub FuerzaImporte()
    On Error GoTo ErrHandler
    
    Dim lngCount As Long
    Dim i As Long
    Dim varIDPrefattura As Variant
    Dim varDetailList As Variant
    Dim ErrCode As ErrCodes
    Dim strRifOrdineEsterno As String
    Dim varIDRiga As Variant
    Dim blnConcurError As Boolean
    Dim Difextint As Double
    Dim Margen As Double
    Dim Cambiar As Boolean
    Dim Cambiados As Integer
    Dim k As Integer
    Dim varListItem As Variant
    Dim blnPropostaMancante As Boolean
    
    ' establecemos el margen que no queremos discrepar
    Margen = 0.02
    
    StatusBar1.Panels(2).Text = "Forzando importes..."
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    
    ErrCode = NoError

    
        With DetailClass.DS
        DetailClass.GetListCheck True
        varDetailList = .List
        lngCount = UBound(DetailClass.DS.List, 2)
        Cambiar = False
        Cambiados = 0
        
        For i = 0 To lngCount

            TDBGrid1.bookmark = i
            
            varIDRiga = TDBGrid1.Columns(0).Text
            .Fields.ResetFields
            .DA.EndTrans ETCOMMIT
            .DA.BeginTransaction
           
            varIDPrefattura = DataClass.DS.Fields("0").Value
            strRifOrdineEsterno = varDetailList(3, i)
           
            ' Buscamos la referencia Cepsa en prefactura Cepsa
            
            GP1.DS.Fields("2").FindValue = strRifOrdineEsterno
            If GP1.Find Then
                If GP1.DS.HowMany > 0 Then
                
                    ' Buscamos en nuestra prefactura la referencia cepsa
                    .Fields("0").FindValue = varIDPrefattura
                    .Fields("1").FindValue = varIDRiga
                    If DetailClass.Find Then
                    
         
                    
                        If .HowMany > 0 Then
                        
                            '' JG 02/06/2017 Buscar en Cfact y CDIFE
                         With PPE.DS
                         .Fields("13").FindValue = strRifOrdineEsterno
                        
                        
                        If Not PPE.GetSelectionCheck(True) Then
                            blnPropostaMancante = True
                        End If
                        
                        varListItem = .List
                        For k = 1 To .HowMany - 1
                            varListItem(10, 0) = varListItem(10, 0) + varListItem(10, k)
                            varListItem(25, 0) = varListItem(25, 0) + varListItem(25, k)
                            varListItem(27, 0) = varListItem(27, 0) + varListItem(27, k)
                            varListItem(28, 0) = varListItem(28, 0) + varListItem(28, k)
                            varListItem(29, 0) = varListItem(29, 0) + varListItem(29, k)
                            varListItem(30, 0) = varListItem(30, 0) + varListItem(30, k)
                        Next k
                        End With
                        

                        
                        ' Chequeo importe
                        ' Cambiamos Margen
                        Difextint = Abs(.Fields("44").Value - varListItem(10, 0))
                        If Difextint > 0 And Difextint < Margen Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("44").NewValue = varListItem(10, 0)
                        .Fields("44").Updated = True
                        End If
                        ' Cambiamos importe
                        If .Fields("44").Value < varListItem(10, 0) Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("44").NewValue = varListItem(10, 0)
                        .Fields("44").Updated = True
                        End If
                        ' Chequeo Tarifa
                        If .Fields("43").Value < GP1.DS.Fields("7").Value Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("43").NewValue = GP1.DS.Fields("7").Value
                        .Fields("43").Updated = True
                        End If
                        ' Chequeo peaje
                        If .Fields("67").Value < varListItem(25, 0) Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("67").NewValue = varListItem(25, 0)
                        .Fields("67").Updated = True
                        End If
                        ' Chequeo Operario
                        If .Fields("69").Value < varListItem(27, 0) Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("69").NewValue = varListItem(27, 0)
                        .Fields("69").Updated = True
                        End If
                        ' Chequeo Buques
                        If .Fields("70").Value < varListItem(28, 0) Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("70").NewValue = varListItem(28, 0)
                        .Fields("70").Updated = True
                        End If
                        ' Chequeo Especiales
                        If .Fields("72").Value < varListItem(30, 0) Then
                        Cambiados = Cambiados + 1
                        Cambiar = True
                        .Fields("72").NewValue = varListItem(30, 0)
                        .Fields("72").Updated = True
                        End If
                        End If
                    End If
                End If
            
            End If
        'Cambiar = True
        If Cambiar = True Then
        DetailClass.BeginSave blnConcurError
            If blnConcurError Then
            .DA.EndTrans ETROLLBACK
            Else
            .DA.EndTrans ETCOMMIT
            End If
        End If
        .DA.EndTrans ETCOMMIT
        Cambiar = False
        DoEvents
        Next i
    End With
    GoTo EndProc

ErrHandler:


EndProc:
    'JG Control Bloqueos DB
    DataClass.DS.DA.EndTrans ETCOMMIT
    DetailClass.DS.DA.EndTrans ETCOMMIT
    MsgBox (Cambiados & " Importes forzados. Se debe hacer un VERIFICA.")
    DisplayDetail
    StatusBar1.Panels(2).Text = vbNullString
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Unload Me
End Sub
Private Function GetPedaggio(RifOrdine As Variant, _
                             IDDestinazione As Variant, _
                             DataConsegna As Variant, _
                             IDProdotto As Variant, _
                             IDUdT As Variant) As Single
    Dim sngResult As Single
    
    With PPE
        .DS.Fields.ResetFields
        If Len(RifOrdine) = 0 Then
            .DS.Fields("10").FindValue = IDDestinazione
            .DS.Fields("14").FindValue = DataConsegna
            .DS.Fields("9").FindValue = IDProdotto
            .DS.Fields("12").FindValue = IDUdT
        Else
            .DS.Fields("13").FindValue = RifOrdine
        End If
                
        .GetSelectionPedaggio
        If IsEmpty(.DS.List) Then
            sngResult = 0
        Else
            sngResult = .DS.List(0, 0)
        End If
    End With
    GetPedaggio = sngResult
End Function

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Private Sub DisplayDetail()
    Me.MousePointer = vbHourglass
    With DetailClass
        .DS.Fields("0").FindValue = DataClass.DS.Fields("0").Value
        .GetList
        RefreshUnboundGrid TDBGrid1, .DS.List, GridArrayDetail, _
                           .DS.HowMany, MAX_COLS
    End With
    Me.MousePointer = vbDefault
End Sub


Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(mskField1) Then
                        SetState fsEdit
                    Else
                        If State = fsNone Then
                            SetState fsNone
                        End If
                    End If
                End If
            End If
        End If
        

    End With
End Sub





Private Sub ClearForm()
    DataClass.DS.ClearClient
    ClearUnboundGrid TDBGrid1, GridArrayDetail, 6
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim C As Control
    Dim LinkForm As Form
    ReDim varReportsList(0 To 2, 0 To 19) As Variant
    Dim strSelectionFormula As String
    Dim varDataDoc As Variant
    
    Select Case Cmd
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
                Case Else
                    blnResult = False
            End Select
'            If blnResult Then
'                SetState fsNone
'            End If
        Case fcFind
            If MsgBox("Esta opción cambiará los importes de la prefactura que sean más elevados en la prefactura de Cepsa ¡DEPUES SE DEBE HACER UN VERIFICA!. Esta seguro?", _
                      vbYesNo + vbExclamation, "Attenzione") = vbYes Then
                FuerzaImporte
            End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcOpen
            Select Case DataClass.DS.Fields("5").Value
                Case ID_SOC_CEPSA
                    CheckPrefattureCEPSA
                Case Else
                    If MsgBox("Vuoi forzare i dati valorizzati ?", vbExclamation + vbYesNo + vbDefaultButton2, _
                              "Attenzione") = vbYes Then
                        CheckPrefatture True
                    ElseIf MsgBox("Vuoi forzare le quantita consegnate ?", vbExclamation + vbYesNo + vbDefaultButton2, _
                              "Attenzione") = vbYes Then
                        CheckPrefatture False, True
                    Else
                        CheckPrefatture
                    End If
            End Select
        Case fcPrint
            strSelectionFormula = vbNullString
            If (State = fsEdit) Then
                varReportsList(0, 0) = "Stampa di controllo per Unità di Trasporto"
                varReportsList(1, 0) = "oil2ppt1.rpt"
                varReportsList(2, 0) = 1
                varReportsList(0, 1) = "Stampa di controllo per Data"
                varReportsList(1, 1) = "oil2ppt8.rpt"
                varReportsList(2, 1) = 1
                varReportsList(0, 2) = "Stampa anomalie Prefatturazione"
                varReportsList(1, 2) = "oil2ppt2.rpt"
                varReportsList(2, 2) = 1
                varReportsList(0, 3) = "Stampa importi discordanti Prefatturazione"
                varReportsList(1, 3) = "oil2ppt4.rpt"
                varReportsList(2, 3) = 1
                varReportsList(0, 4) = "Stampa Prefatture Trasportatori"
                varReportsList(1, 4) = "oil2ppt3.rpt"
                varReportsList(2, 4) = 1
                varReportsList(0, 5) = "Riepilogo Prefatture Trasportatori"
                varReportsList(1, 5) = "oil2ppt5.rpt"
                varReportsList(2, 5) = 1
                varReportsList(0, 6) = "Prestazioni Automezzi"
                varReportsList(1, 6) = "oil2ppt6a.rpt"
                varReportsList(2, 6) = 1
                varReportsList(0, 7) = "Riepilogo Prestazioni Automezzi"
                varReportsList(1, 7) = "oil2ppt6b.rpt"
                varReportsList(2, 7) = 1
                varReportsList(0, 8) = "Prestazioni Autisti"
                varReportsList(1, 8) = "oil2pptA.rpt"
                varReportsList(2, 8) = 1
                varReportsList(0, 9) = "Controllo quantità caricate e scaricate"
                varReportsList(1, 9) = "oil2ppt7.rpt"
                varReportsList(2, 9) = 0
                varReportsList(0, 10) = "Controllo Minimi Tariffari"
                varReportsList(1, 10) = "oil2ppt9.rpt"
                varReportsList(2, 10) = 1
                varReportsList(0, 11) = "Stampa di controllo per Autista"
                varReportsList(1, 11) = "oil2pptC.rpt"
                varReportsList(2, 11) = 1
                varReportsList(0, 12) = "Prestazioni Automezzi per Deposito"
                varReportsList(1, 12) = "oil2pptD.rpt"
                varReportsList(2, 12) = 1
                varReportsList(0, 13) = "Prestazioni Automezzi per Fascia Tariffaria"
                varReportsList(1, 13) = "oil2pptE.rpt"
                varReportsList(2, 13) = 1
                varReportsList(0, 14) = "Prefatturazione per Unità di Trasporto"
                varReportsList(1, 14) = "oil2pptF.rpt"
                varReportsList(2, 14) = 1
                varReportsList(0, 15) = "Allegato fattura"
                varReportsList(1, 15) = "oil2pptG.rpt"
                varReportsList(2, 15) = 1
                varReportsList(0, 16) = "Elenco servizi trasportatori terzi"
                varReportsList(1, 16) = "oil2pptH.rpt"
                varReportsList(2, 16) = 1
                varReportsList(0, 17) = "Elenco servizi autisti"
                varReportsList(1, 17) = "oil2pptI.rpt"
                varReportsList(2, 17) = 1
                varReportsList(0, 18) = "Totale controllo fatturazione"
                varReportsList(1, 18) = "oil2pptL.rpt"
                varReportsList(2, 18) = 1
                varReportsList(0, 19) = "Resumen Facturacion por conductor"
                varReportsList(1, 19) = "oil2ppt6c.rpt"
                varReportsList(2, 19) = 1
                
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                CRPrint varReportsList, strSelectionFormula, Me
            End If
'        Case fcFind
'            DoSearch "PPT", Me, "0", 1
        Case fcCustom1
        
        Call MsgBox("Las prefacturas solo se pueden eliminar desde Cybertruck.", vbOKOnly + vbExclamation, "ATENCION")
        
'            If DataClass.DS.Field("Stato").Value <> "G" Then
'                Call MsgBox("La prefattura non può essere eliminata in quanto sono state generate le fatture.", vbOKOnly + vbExclamation, "ATTENZIONE")
'            Else
'                With DataClass.DS
'                    .DA.BeginTransaction
'                    If DataClass.Delete Then
'                        If MsgBox("La prefattura verrà eliminata definitivamente. Vuoi proseguire?", vbYesNo + vbExclamation, "ATTENZIONE") = vbYes Then
'                            .DA.EndTrans ETCOMMIT
'                            SetState fsNone
'                        Else
'                            .DA.EndTrans ETROLLBACK
'                        End If
'                    Else
'                        .DA.EndTrans ETROLLBACK
'                    End If
'                End With
'            End If
        Case fccustom2
            If DataClass.DS.Field("Stato").Value <> "G" Then
                Call MsgBox("La prefattura non può essere eliminata in quanto sono state generate le fatture.", vbOKOnly + vbExclamation, "ATTENZIONE")
            Else
                With DetailClass.DS
                    .DA.BeginTransaction
                    If DetailClass.Delete(DataClass.DS.Fields("0").Value, GridArrayDetail.Value(TDBGrid1.bookmark, 0)) Then
                        If MsgBox("La riga selezionata verrà eliminata definitivamente. Vuoi proseguire?", vbYesNo + vbExclamation, "ATTENZIONE") = vbYes Then
                            .DA.EndTrans ETCOMMIT
                            DisplayDetail
                        Else
                            .DA.EndTrans ETROLLBACK
                        End If
                    Else
                        .DA.EndTrans ETROLLBACK
                    End If
                End With
            End If
        
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcPREFPPD)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    With DetailClass.DS
'                        LinkForm.GetRecord .Fields(.KeyFields(1)).value, _
'                                           .Fields(.KeyFields(2)).value
                        LinkForm.GetRecord DataClass.DS.Fields("0").Value, _
                                           GridArrayDetail.Value(TDBGrid1.bookmark, 0)
                    End With
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcExport
            If MsgBox("Verranno generati i documenti per questa prefattura. Vuoi proseguire?", _
                      vbYesNo + vbExclamation, "Attenzione") = vbYes Then
                varDataDoc = InputBox("Inserire la data per i documenti: ", "Data Documenti", Format(Now, "dd/mm/yy"))
                Me.MousePointer = vbHourglass
                DataClass.DS.DisplayMessage "Elaborazione in corso..."
                DoEvents
                If DataClass.CreateDocumento(varDataDoc) Then
                    DataClass.DS.DisplayMessage "Documenti creati correttamente."
                Else
                    DataClass.DS.DisplayMessage "Si è verificato un errore durante la creazione dei documenti.", Error
                End If
                Me.MousePointer = vbDefault
                Call GetRecord(DataClass.DS.Field("Codice").Value)
            End If
        Case fcExit
            Unload Me
    End Select
End Sub


Public Function GetRecord(strID As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
            If .HowMany > 0 Then
'                If DataClass.DS.Fields("5").value = ID_SOC_CEPSA Then
'                    TDBGrid1.FetchRowStyle = True
'                Else
'                    TDBGrid1.FetchRowStyle = False
'                End If
                If .RefreshClient Then
                    DisplayDetail
                End If
            Else
                GoTo ErrHandler
            End If
        Else
            GoTo ErrHandler
        End If
    End With
    GetRecord = True
    Exit Function
ErrHandler:
    GetRecord = False
    Exit Function
End Function







Private Function UpdateRecord() As Boolean
    On Error GoTo ErrHandler
    
    Dim blnChkValue As Boolean
    Dim blnConcurError As Boolean
    Dim C As Control
    
    'Scansione della finestra
    DataClass.DS.ScanClient 0
    
    If DataClass.DS.Fields.Validate Then
        DataClass.BeginSave blnConcurError
        If blnConcurError Then
            If MsgBox(MSG_CONCUR_ERR, _
                       vbOKCancel + vbExclamation + vbDefaultButton2, _
                      "Attenzione") = vbOK Then
                GoTo EndProc
            Else
                DataClass.DS.DA.EndTrans ETROLLBACK
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).Value
                GoTo ErrHandler
            End If
        Else
            GoTo EndProc
        End If
        GoTo EndProc
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    
ErrHandler:
    UpdateRecord = False
    Exit Function
EndProc:
    DataClass.DS.DA.EndTrans ETCOMMIT
    DataClass.DS.EndSave
    UpdateRecord = True
End Function

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsNone
            If Not Loading Then
                mskField1.SetFocus
            End If
            ClearUnboundGrid TDBGrid1, GridArrayDetail, 6
            Command1.Enabled = False
            Command5.Enabled = False
            Command7.Enabled = False
        Case fsEdit
            Command1.Enabled = True
            Command5.Enabled = True
            Command7.Enabled = True
            mskField2.SetFocus
    End Select
End Sub


Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    ReDim ColTags(1 To MAX_COLS) As Variant
    
    Loading = True
    Set GridArrayDetail = New ArrayGridCls
    Set DataClass = New CPrefattura
    Set DetailClass = New CPrefatturaDettaglio
    
    ColTags(1) = "1"
    ColTags(2) = "2"
    ColTags(3) = "3"
    ColTags(4) = "6"
    ColTags(5) = "12"
    ColTags(6) = "13"
    ColTags(7) = "14"
    ColTags(8) = "15"
    ColTags(9) = "55"
    ColTags(10) = "56"
    ColTags(11) = "26"
    ColTags(12) = "30"
    ColTags(13) = "43"
    ColTags(14) = "44"
    ColTags(15) = "50"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
        Set .DA = DataConnect.DA
        .Field("Stato").ReadOnly = True
        .ConfigureInterface
        ConfigureObjects Me, ColTags, False
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
'        Set DetailClass.DS.DA = DA
        Set DetailClass.DS.DA = DataConnect.DA
        Set DetailClass.DS.ErrH = ErrH
    End With
    
    Set PPE = New CPrefatturaSocSelezione
    With PPE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set PPD = New CPrefatturaDettaglio
    With PPD.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
   
    Set AXX = New CParametro
    With AXX.DS
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
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set GP1 = New CPrefatturaSoc
    With GP1.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    SetState fsNone
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set DetailClass = Nothing
    Set PPE = Nothing
    Set AXX = Nothing
    Set VVT = Nothing
    Set VVP = Nothing
    Set APR = Nothing
    
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set OIL2PPT0 = Nothing
End Sub





Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub





Private Sub TDBGrid1_DblClick()
    DoCommand fcOpenLinkEdit
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
'    On Error GoTo ErrHandler
'
'    Dim varImportoI As Variant
'    Dim varImportoE As Variant
'
'    varImportoI = DetailClass.DS.List(11, Bookmark)
'    varImportoE = DetailClass.DS.List(21, Bookmark)
'    If IsNumeric(varImportoI) And IsNumeric(varImportoE) Then
'        Select Case Abs(varImportoI - varImportoE)
'            Case 0
'            Case Is <= 0.1
'                RowStyle.BackColor = vbYellow
'                RowStyle.ForeColor = vbBlack
'            Case Is > 0.1
'                RowStyle.BackColor = vbRed
'                RowStyle.ForeColor = vbWhite
'        End Select
'    End If
'ErrHandler:
End Sub


Private Sub TDBGrid1_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    GridArrayDetail.ReadData RowBuf, StartLocation, offset, ApproximatePosition
End Sub


