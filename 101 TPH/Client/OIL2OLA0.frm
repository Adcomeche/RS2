VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2OLA0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADDEBITI STIPENDIO"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   10845
   Begin VB.Frame Frame1 
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10815
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   1020
         TabIndex        =   8
         Tag             =   "0"
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   6000
         TabIndex        =   9
         Tag             =   "1"
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9600
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.txt"
      End
      Begin VB.Label lblField1 
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
         Left            =   1740
         TabIndex        =   12
         Tag             =   "0"
         Top             =   300
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Autista"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   7320
         Tag             =   "1"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "&Data"
         Height          =   255
         Index           =   12
         Left            =   5280
         TabIndex        =   10
         Top             =   300
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   4560
         Tag             =   "0"
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   915
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   10815
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   9120
         Picture         =   "OIL2OLA0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
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
         Picture         =   "OIL2OLA0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9000
         Picture         =   "OIL2OLA0.frx":0204
         TabIndex        =   1
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva e Chiudi"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2OLA0.frx":0306
         TabIndex        =   0
         Tag             =   "SAVE"
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4860
      Width           =   10845
      _ExtentX        =   19129
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
            Object.Width           =   13679
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9.32"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "13/04/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Addebiti"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   10815
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6480
         Top             =   2880
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
               Picture         =   "OIL2OLA0.frx":0408
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2OLA0.frx":051A
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2OLA0.frx":062C
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin TrueDBGrid60.TDBGrid TDBgrid1 
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "OIL2OLA0.frx":073E
         TabIndex        =   13
         Tag             =   "SELECT"
         Top             =   240
         Width           =   10575
      End
   End
End
Attribute VB_Name = "OIL2OLA0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public XAAddebiti As XArrayDB
Public WithEvents DataClass As CStipendioDettaglioAddebito
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mobjParentForm As Object
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Public Property Get ParentForm() As Form
    Set ParentForm = mobjParentForm
End Property


Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
                Case Else
                    blnResult = False
            End Select
            If blnResult Then
                DoCommand fcExit
            End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsEdit
            End If
        Case fcExit
            Unload Me
    End Select
End Sub
Public Function GetRecord(IDAutista As Variant, _
                          Data As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields.ResetFields
'        .Field("Codice Autista").FindValue(1) = IDAutista
'        .Field("Codice Autista").FindValue(2) = "#NULL#"
'        .Field("Data").FindValue(1) = Data
'        .Field("Data").FindValue(2) = "#NULL#"
        
        If DataClass.GetSelection Then
            DisplaySelection DataClass.DS.List
'            SetState fsEdit
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


Private Sub DisplaySelection(ListAddebiti As Variant)
    Dim i As Long
    Dim j As Long
    Dim varList As Variant
    
    ReDim varList(UBound(ListAddebiti, 1) + 1, UBound(ListAddebiti, 2))
    For i = 0 To UBound(ListAddebiti, 2)
    
        With DataClass.DS
            .Fields.ResetFields
            .Field("Codice Autista").FindValue = mskField10
            .Field("Data").FindValue = mskDateFrom
            .Field("Codice Addebito").FindValue = ListAddebiti(0, i)
            DataClass.Find
            If .HowMany > 0 Then
                varList(0, i) = 1
                ListAddebiti(2, i) = .Field("Quantità").value
                ListAddebiti(4, i) = .Field("Descrizione").value
                ListAddebiti(5, i) = .Field("Importo").value
            Else
                varList(0, i) = 0
            End If
        End With
            
        For j = 1 To UBound(ListAddebiti, 1) + 1
            varList(j, i) = ListAddebiti(j - 1, i)
        Next j
    Next i
    
    If Not IsEmpty(varList) Then
        RefreshXArray XAAddebiti, TDBgrid1, varList
        TDBgrid1.Enabled = True
    Else
        ResetGrid TDBgrid1, XAAddebiti
    End If
End Sub

Private Sub ResetGrid(Grid As TDBGrid, XA As XArrayDB)
    
    If XA.Count(1) > 0 Then
        XA.DeleteRows XA.LowerBound(1), XA.Count(1)
    End If
    Set Grid.Array = XA
    Grid.ReBind

End Sub


Private Function InsertRecord() As Boolean
    On Error GoTo ErrHandler
    Dim C As Control
    
    Dim blnChkValue As Boolean
    DataClass.DS.ScanClient 1
    
    If DataClass.DS.Fields.Validate Then
        If Not DataClass.Insert Then
            GoTo ErrHandler
        End If
        GoTo EndProc
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
ErrHandler:
    InsertRecord = False
    Exit Function
EndProc:
    InsertRecord = True
End Function

Private Function UpdateRecord() As Boolean
    On Error GoTo ErrHandler
    
    Dim blnConcurError As Boolean
    Dim i As Long
    
    
    DataClass.DS.DA.BeginTransaction
    
    Me.MousePointer = vbHourglass
    For i = 0 To XAAddebiti.UpperBound(1)
        TDBgrid1.bookmark = i
        If Not DataClass.Delete(mskField10, mskDateFrom, TDBgrid1.Columns(1).Text) Then
            GoTo ErrHandler
        End If
        If TDBgrid1.Columns(0).value <> 0 Then
            With DataClass.DS
                .Fields.ResetFields
                .Field("Codice Autista").NewValue = mskField10
                .Field("Codice Autista").Updated = True
                .Field("Data").NewValue = mskDateFrom
                .Field("Data").Updated = True
                .Field("Codice Addebito").NewValue = TDBgrid1.Columns(1).Text
                .Field("Codice Addebito").Updated = True
                If IsNumeric(TDBgrid1.Columns(3).value) Then
                    .Field("Quantità").NewValue = TDBgrid1.Columns(3).value
                Else
                    .Field("Quantità").NewValue = 1
'                    .Field("Unità di Misura").NewValue = TDBgrid1.Columns(4).value
'                    .Field("Unità di Misura").Updated = True
                End If
                .Field("Quantità").Updated = True
                .Field("Descrizione").NewValue = TDBgrid1.Columns(5).Text
                .Field("Descrizione").Updated = True
                If IsNumeric(TDBgrid1.Columns(6).value) Then
                    .Field("Importo").NewValue = TDBgrid1.Columns(6).value
                Else
                    .Field("Importo").NewValue = 0
                End If
                .Field("Importo").Updated = True
            End With
            If Not DataClass.Insert Then
                GoTo ErrHandler
            End If
        End If
    Next i
    
    GoTo EndProc
    
ErrHandler:
    Me.MousePointer = vbNormal
    DataClass.DS.DA.EndTrans ETROLLBACK
    UpdateRecord = False
    Exit Function
EndProc:
    Me.MousePointer = vbNormal
    DataClass.DS.DA.EndTrans ETCOMMIT
    UpdateRecord = True
End Function

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsNone
        Case fsEdit
            Me.TDBgrid1.SetFocus
        Case fsInsert
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
    ReDim ColTags(1 To 5) As Variant
    
    Loading = True
    Set XAAddebiti = New XArrayDB
    Set DataClass = New CStipendioDettaglioAddebito
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags, True
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .TDBgrid1.SetFocus
'            .DoCommand fcLinkRefresh
        End With
    End If
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2OLA0 = Nothing
End Sub

