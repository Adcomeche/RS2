VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form OIL2ACD0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIGA CONTRATTO"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   9630
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9615
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7500
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "6"
         Top             =   1740
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1740
         Width           =   3435
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1320
         Width           =   1935
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7320
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
               Picture         =   "OIL2ACD0.frx":0000
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ACD0.frx":0112
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ACD0.frx":0224
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1980
         TabIndex        =   1
         Tag             =   "1"
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   1980
         TabIndex        =   0
         Tag             =   "0"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1980
         TabIndex        =   2
         Tag             =   "2"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1980
         TabIndex        =   7
         Tag             =   "7"
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   7500
         TabIndex        =   4
         Tag             =   "4"
         Top             =   1380
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   9300
         Tag             =   "4"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   9300
         Tag             =   "6"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   5460
         Tag             =   "5"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   4020
         Tag             =   "3"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5340
         Tag             =   "7"
         Top             =   2160
         Width           =   240
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
         Left            =   2580
         TabIndex        =   29
         Tag             =   "7"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Calcolo"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   28
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Separazione Fattura"
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   27
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Parametri tassazione"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimo tassabile"
         Height          =   255
         Index           =   2
         Left            =   5940
         TabIndex        =   25
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Flag UDM"
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   24
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5340
         Tag             =   "0"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5340
         Tag             =   "2"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Addebito"
         Height          =   255
         Index           =   8
         Left            =   780
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblField0 
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
         Left            =   2580
         TabIndex        =   22
         Tag             =   "2"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2700
         Tag             =   "1"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Contratto"
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente"
         Height          =   255
         Index           =   9
         Left            =   1020
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   2580
         TabIndex        =   19
         Tag             =   "0"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   2640
      Width           =   9615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   5280
         Picture         =   "OIL2ACD0.frx":0336
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "OIL2ACD0.frx":0438
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
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
         Left            =   1920
         Picture         =   "OIL2ACD0.frx":053A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
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
         Index           =   2
         Left            =   7920
         Picture         =   "OIL2ACD0.frx":063C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
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
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2ACD0.frx":073E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7800
         Picture         =   "OIL2ACD0.frx":0840
         TabIndex        =   11
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Tariffe Contratto"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2ACD0.frx":0942
         TabIndex        =   31
         Tag             =   "CUSTOM1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2ACD0.frx":0A44
         TabIndex        =   10
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2ACD0.frx":0B46
         TabIndex        =   9
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ACD0.frx":0C48
         TabIndex        =   8
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3525
      Width           =   9630
      _ExtentX        =   16986
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
            Object.Width           =   11536
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14.50"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09/03/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2ACD0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DataClass As CContrattoDettaglio
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mobjParentForm As Object

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcAddNew
            If (State = fsEdit) Or (State = fsNone) Then
                DataClass.DS.Fields.ResetFields
                ClearForm
                SetState fsInsert
            End If
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
                Case fsInsert
                    blnResult = InsertRecord
                Case Else
                    blnResult = False
            End Select
            If blnResult Then
                If Not mobjParentForm Is Nothing Then
                    Select Case mobjParentForm.Name
                        Case "OIL2ACN0"
                            SetState fsNone
                            DoCommand fcExit
                    End Select
                Else
                    SetState fsNone
                End If
            End If

        Case fcCustom1
            If (State = fsEdit) Then
                Set LinkForm = DisplayForm(fcTARITFC)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .IDACL = DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
                        .IDACN = DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value
                        .IDADD = DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value
'                        .SetState fsInsert
                        .DisplayList
                        Me.Enabled = False
                    End With
                End If
            End If

        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcFind
            DoSearch "AAM", Me, "0", 1
        Case fcExit
            Unload Me
    End Select
End Sub

Public Function GetRecord(varIDACL As Variant, _
    varIDACN As Variant, _
    varIDADD As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = varIDACL
        .Fields(.KeyFields(2)).FindValue = varIDACN
        .Fields(.KeyFields(3)).FindValue = varIDADD
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    SetState fsEdit
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

Private Function InsertRecord() As Boolean
    On Error GoTo ErrHandler
    
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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value
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
                MaskEdBox7.SetFocus
            End If
        Case fsEdit
            Combo1.SetFocus
        Case fsInsert
            MaskEdBox7.SetFocus
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
    
    Loading = True
    
    Set DataClass = New CContrattoDettaglio
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With

    SetState fsNone
    Loading = False
End Sub

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(3) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(MaskEdBox7.Text, mskField1.Text, mskField4.Text) Then
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            Select Case .Name
                Case "OIL2ACN0"
                    .DoCommand fcLinkRefresh
            End Select
        End With
    End If

    Set ErrH = Nothing
    Set DataClass.DS.FormHandle = Nothing
    Set DataClass = Nothing
    Set OIL2ACD0 = Nothing
End Sub

Private Sub mskField4_LostFocus()
    TrackLostFocus mskField4
End Sub
