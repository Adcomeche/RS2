VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPPTSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONE PREFATTURE"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   9390
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   2520
         Picture         =   "frmPPTSelect.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3675
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Prefattura"
         Height          =   375
         Left            =   2400
         Picture         =   "frmPPTSelect.frx":0102
         TabIndex        =   16
         Tag             =   "OPEN"
         Top             =   3600
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmPPTSelect.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3675
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "frmPPTSelect.frx":0306
         TabIndex        =   15
         Tag             =   "PRINT"
         Top             =   3600
         Width           =   2295
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "frmPPTSelect.frx":0408
         TabIndex        =   14
         Tag             =   "SELECT"
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   4095
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Width           =   9015
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "5"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   4
         Left            =   7440
         Picture         =   "frmPPTSelect.frx":39F2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmPPTSelect.frx":3AF4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Tag             =   "2"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Tag             =   "1"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   6120
         TabIndex        =   8
         Tag             =   "3"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   2640
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
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
               Picture         =   "frmPPTSelect.frx":3BF6
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPPTSelect.frx":3D08
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPPTSelect.frx":3E1A
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Tag             =   "4"
         Top             =   1080
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
         Index           =   27
         Left            =   4080
         Tag             =   "5"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Prefattura"
         Height          =   255
         Index           =   27
         Left            =   480
         TabIndex        =   33
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   2520
         TabIndex        =   32
         Tag             =   "4"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   6240
         Tag             =   "4"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   31
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "testate di prefattura"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   7920
         Tag             =   "3"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Cons. Ric&h."
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7440
         Tag             =   "1"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla D&ata Cons. Rich."
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3720
         Tag             =   "2"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3240
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   22
      Top             =   4920
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmPPTSelect.frx":3F2C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmPPTSelect.frx":402E
         TabIndex        =   18
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
         Index           =   7
         Left            =   240
         Picture         =   "frmPPTSelect.frx":4130
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmPPTSelect.frx":4232
         TabIndex        =   17
         Tag             =   "FIND"
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
         Index           =   2
         Left            =   7680
         Picture         =   "frmPPTSelect.frx":4334
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmPPTSelect.frx":4436
         TabIndex        =   19
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Condizioni"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Risultato Ricerca"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   5805
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
            TextSave        =   "10.57"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15/10/2008"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPPTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3 '11
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As CPrefatturaSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Sub DisplaySelection()
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, DataClass.DS.HowMany, MAX_COLS

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = GridArraySelect.value(DBGrid.bookmark, 0)
    End If
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varOrderFields As Variant
    
    With DataClass.DS
        If .Fields.Validate Then
            If .ScanClientCond Then
                ReDim varOrderFields(0) As Variant
                
                varOrderFields(0) = "0"
                If Not DataClass.GetSelection(varOrderFields, CLng(mskMaxRecords)) Then
                    GoTo ErrHandler
                End If
                DisplaySelection
                GoTo EndProc
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
ErrHandler:
    SelectRecords = False
    Exit Function
EndProc:
    SelectRecords = True
End Function

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            If Not Loading Then
                TabStrip1.Tabs(1).Selected = True
                mskField1.SetFocus
            End If
        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Prefatture trovate: " & DataClass.DS.HowMany
            TabStrip1.Tabs(2).Selected = True
            If Not Loading Then
                TDBGrid1.SetFocus
            End If
    End Select
End Sub



Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strSelectionFormula As String
    ReDim varReportsList(0 To 2, 0 To 13) As Variant
    
    Select Case Cmd
        Case fcUndo
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
        Case fcFind
            If SelectRecords Then
                SetState fsSelection
            End If
        Case fcPrint
            If (State = fsSelection) Then
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
                varReportsList(0, 8) = "Controllo quantità caricate e scaricate"
                varReportsList(1, 8) = "oil2ppt7.rpt"
                varReportsList(2, 8) = 0
                varReportsList(0, 9) = "Controllo Minimi Tariffari"
                varReportsList(1, 9) = "oil2ppt9.rpt"
                varReportsList(2, 9) = 1
                varReportsList(0, 10) = "Prestazioni Automezzi per Deposito"
                varReportsList(1, 10) = "oil2pptD.rpt"
                varReportsList(2, 10) = 1
                varReportsList(0, 11) = "Prestazioni Automezzi per Fascia Tariffaria"
                varReportsList(1, 11) = "oil2pptE.rpt"
                varReportsList(2, 11) = 1
                varReportsList(0, 12) = "Prefatturazione per Unità di Trasporto"
                varReportsList(1, 12) = "oil2pptF.rpt"
                varReportsList(2, 12) = 1
                varReportsList(0, 13) = "Totale controllo fatturazione"
                varReportsList(1, 13) = "oil2pptL.rpt"
                varReportsList(2, 13) = 1
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcPREFPPT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcLinkRefresh
            If SelectRecords Then
                If TDBGrid1.Enabled Then
                    TDBGrid1.SetFocus
                End If
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub GetRecord(strID As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
        Else
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub







Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Dalla data cons. rich."
        Case 4
            Set M = MaskEdBox7
            strTitle = "Alla data cons. rich."
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    ReDim ColTags(1 To 2) As Variant
    
    Loading = True
    mintCurFrame = 1
    Set GridArraySelect = New ArrayGridCls
    
    ColTags(1) = "0"
    ColTags(2) = "1"
    
    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New CPrefatturaSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    
    SetState fsparaminput
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set frmPPTSelect = Nothing
End Sub





Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DoEvents
        DoCommand fcOpen
    End If
End Sub


