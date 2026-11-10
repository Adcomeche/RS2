VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2PAD0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PREFATTURA ADDEBITO"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   915
      Left            =   0
      TabIndex        =   12
      Top             =   2700
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
         Picture         =   "OIL2PAD0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
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
         Picture         =   "OIL2PAD0.frx":0102
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
         Left            =   7560
         Picture         =   "OIL2PAD0.frx":0204
         TabIndex        =   9
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2PAD0.frx":0306
         TabIndex        =   8
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3660
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
            TextSave        =   "11.06"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "29/04/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   " Dettaglio"
      Height          =   1455
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   9375
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8700
         Top             =   600
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
               Picture         =   "OIL2PAD0.frx":0408
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2PAD0.frx":051A
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2PAD0.frx":062C
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Tag             =   "5"
         Top             =   300
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Tag             =   "6"
         Top             =   660
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   6180
         TabIndex        =   5
         Tag             =   "7"
         Top             =   660
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Tag             =   "8"
         Top             =   1020
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   6180
         TabIndex        =   7
         Tag             =   "9"
         Top             =   1020
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Importo"
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   19
         Top             =   1020
         Width           =   915
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   8040
         Tag             =   "9"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valore Tariffa"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   3780
         Tag             =   "8"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Litri"
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   17
         Top             =   660
         Width           =   675
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   8040
         Tag             =   "7"
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Chilometri"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   660
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   3780
         Tag             =   "6"
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero Consegne"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   3780
         Tag             =   "5"
         Top             =   300
         Width           =   240
      End
   End
   Begin MSMask.MaskEdBox mskLink1 
      Height          =   255
      Left            =   7560
      TabIndex        =   20
      Tag             =   "0"
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskField3 
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Tag             =   "4"
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   25
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskField1 
      Height          =   255
      Left            =   1620
      TabIndex        =   0
      Tag             =   "2"
      Top             =   120
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
      Left            =   1620
      TabIndex        =   1
      Tag             =   "3"
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   3
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLink2 
      Height          =   255
      Left            =   7560
      TabIndex        =   28
      Tag             =   "1"
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero Riga Pref."
      Height          =   255
      Index           =   9
      Left            =   5760
      TabIndex        =   29
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   8
      Left            =   5400
      Tag             =   "4"
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cod. Addebito"
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   27
      Top             =   840
      Width           =   1455
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
      Left            =   2280
      TabIndex        =   26
      Tag             =   "4"
      Top             =   840
      Width           =   2955
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   46
      Left            =   5400
      Tag             =   "2"
      Top             =   120
      Width           =   240
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
      Left            =   2280
      TabIndex        =   25
      Tag             =   "2"
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente"
      Height          =   255
      Index           =   50
      Left            =   780
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   2
      Left            =   5400
      Tag             =   "3"
      Top             =   480
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
      Left            =   2280
      TabIndex        =   23
      Tag             =   "3"
      Top             =   480
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Contratto"
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   22
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero Testata Pref."
      Height          =   255
      Index           =   28
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "OIL2PAD0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GridArray As ArrayGridCls
Public WithEvents DataClass As CPrefatturaDettaglioAddebito
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

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(3) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(mskLink1, mskLink2, mskField1, mskField2, mskField3) Then
                        SetState fsEdit
                    Else
'                        If State = fsNone Then
'                            SetState fsNone
'                        End If
                    End If
                End If
            End If
        End If
        

    End With
End Sub

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
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
                SetState fsNone
                
                If Not mobjParentForm Is Nothing Then
                    Select Case mobjParentForm.Name
                        Case "OIL2PPD0"
                            DoCommand fcExit
                        Case Else
                            SetState fsNone
                    End Select
                End If
            End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcPrint
            If (State = fsEdit) Or (State = fsNone) Then
            End If
        Case fcFind
            ShowSearchForm Me, "0", DataClass, 1
        Case fcExit
            Unload Me
    End Select
End Sub
Public Function GetRecord(varIDPPT As Variant, _
                          varIDPPD As Variant, _
                          varIDACL As Variant, _
                          varIDACN As Variant, _
                          varIDADD As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = varIDPPT
        .Fields(.KeyFields(2)).FindValue = varIDPPD
        .Fields(.KeyFields(3)).FindValue = varIDACL
        .Fields(.KeyFields(4)).FindValue = varIDACN
        .Fields(.KeyFields(5)).FindValue = varIDADD
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(4)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(5)).value
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
            If Not Loading And mskField1.Enabled Then
                mskField1.SetFocus
            End If
        Case fsEdit
'            MaskEdBox2.SetFocus
        Case fsInsert
            If mskField1.Enabled Then
                mskField1.SetFocus
            End If
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
    Set GridArray = New ArrayGridCls
    Set DataClass = New CPrefatturaDettaglioAddebito
    
'    ColTags(1) = "Società"
'    ColTags(2) = "Imballo"
'    ColTags(3) = "PosizioneDoganale"
'    ColTags(4) = "CodPerSoc"
'    ColTags(5) = "Annullato"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags, True
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2PAD0 = Nothing
End Sub

Private Sub mskField3_LostFocus()
    TrackLostFocus mskField1
End Sub

