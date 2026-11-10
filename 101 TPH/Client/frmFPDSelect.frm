VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFPDSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONE FASCE PREZZI"
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
      Caption         =   "Parametri di selezione"
      Height          =   4095
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   9015
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   3
         Format          =   "0000000"
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
         _Version        =   327681
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   327681
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7560
         Tag             =   "1"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "testate di prefattura"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Codice"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero Codice"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3240
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8280
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmFPDSelect.frx":0000
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmFPDSelect.frx":0112
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmFPDSelect.frx":0224
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   2
      Left            =   240
      TabIndex        =   17
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
         Picture         =   "frmFPDSelect.frx":0336
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3675
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Dettaglio"
         Height          =   375
         Left            =   2400
         Picture         =   "frmFPDSelect.frx":0438
         TabIndex        =   7
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
         Picture         =   "frmFPDSelect.frx":053A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3675
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "frmFPDSelect.frx":063C
         TabIndex        =   6
         Tag             =   "PRINT"
         Top             =   3600
         Width           =   2295
      End
      Begin TrueDBGrid50.TDBGrid TDBGrid1 
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "frmFPDSelect.frx":073E
         TabIndex        =   5
         Tag             =   "SELECT"
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   13
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
         Picture         =   "frmFPDSelect.frx":217F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmFPDSelect.frx":2281
         TabIndex        =   9
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
         Picture         =   "frmFPDSelect.frx":2383
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmFPDSelect.frx":2485
         TabIndex        =   8
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
         Picture         =   "frmFPDSelect.frx":2587
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmFPDSelect.frx":2689
         TabIndex        =   10
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Condizioni"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Risultato Ricerca"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   5805
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11113
            Key             =   "Msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17.45"
            Key             =   "Time"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/11/99"
            Key             =   "Date"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFPDSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 2 '11
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As CFasciaPrezziSelezione
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
    If IsNull(DBGrid.Bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = GridArraySelect.value(DBGrid.Bookmark, 0)
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
        Case fsParamInput
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
    ReDim varReportsList(0 To 2, 0 To 2) As Variant
    
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
            If (State = fsSelection) Then
                varReportsList(0, 0) = "Stampa di controllo Prefatturazione"
                varReportsList(1, 0) = "oil2ppt1.rpt"
                varReportsList(2, 0) = 1
                varReportsList(0, 1) = "Stampa anomalie Prefatturazione"
                varReportsList(1, 1) = "oil2ppt2.rpt"
                varReportsList(2, 1) = 1
                varReportsList(0, 2) = "Stampa Prefatture Trasportatori"
                varReportsList(1, 2) = "oil2ppt3.rpt"
                varReportsList(2, 2) = 1
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
    
    ColTags(1) = "Codice"
    ColTags(2) = "Descrizione"
    
    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New CFasciaPrezziSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    SetState fsParamInput
    Loading = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set frmFPDSelect = Nothing
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


