VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2NDA0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NUMERATORI DOCUMENTI AMMINISTRATIVI"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   8580
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   60
      TabIndex        =   20
      Top             =   3060
      Width           =   8475
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "OIL2NDA0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
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
         Picture         =   "OIL2NDA0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
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
         Left            =   6780
         Picture         =   "OIL2NDA0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
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
         Picture         =   "OIL2NDA0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2NDA0.frx":0408
         TabIndex        =   25
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   6660
         Picture         =   "OIL2NDA0.frx":050A
         TabIndex        =   26
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2NDA0.frx":060C
         TabIndex        =   27
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2NDA0.frx":070E
         TabIndex        =   28
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraKey 
      Caption         =   "Tipo Documento"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   3600
      TabIndex        =   9
      Tag             =   "*"
      Top             =   60
      Width           =   4935
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Tag             =   "1"
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Tag             =   "0"
         Top             =   360
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
         Index           =   1
         Left            =   4620
         Tag             =   "2"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   4620
         Tag             =   "1"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   4620
         Tag             =   "0"
         Top             =   360
         Width           =   240
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
         Left            =   2040
         TabIndex        =   15
         Tag             =   "0"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Società"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   615
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
         Left            =   2040
         TabIndex        =   13
         Tag             =   "2"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ente emittente"
         Height          =   255
         Index           =   47
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1095
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
         Left            =   2040
         TabIndex        =   11
         Tag             =   "1"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Doc."
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3975
      Width           =   8580
      _ExtentX        =   15134
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
            Object.Width           =   9684
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17.10"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "24/08/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   3600
      TabIndex        =   7
      Top             =   1620
      Width           =   4935
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "OIL2NDA0.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "CAL"
         Top             =   960
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Tag             =   "3"
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Tag             =   "4"
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Tag             =   "5"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   3660
         Tag             =   "5"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   3180
         Tag             =   "4"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   4620
         Tag             =   "4"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrizione"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Progressivo"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Agg."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   2835
      Left            =   0
      OleObjectBlob   =   "OIL2NDA0.frx":0912
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1260
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
            Picture         =   "OIL2NDA0.frx":3304
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2NDA0.frx":3416
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2NDA0.frx":3528
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2NDA0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3

Private XA As XArrayDB
Public WithEvents DataClass As CNumeratoreDocumentiAmm
Attribute DataClass.VB_VarHelpID = -1
Public GridArray As ArrayGridCls
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
'Private mstrIDAEE As String
'Private mstrIDASC As String
Private mobjParentForm As Object

Private AXX As CParametro

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mvarFunctionCode As FunctionCodes

Public Property Let FunctionCode(varFunctionCode As FunctionCodes)
    mvarFunctionCode = varFunctionCode
End Property

Public Property Get FunctionCode() As FunctionCodes
    FunctionCode = mvarFunctionCode
End Property
'Public Property Let IDSocietà(NewValue As String)
'    mstrIDASC = NewValue
'End Property
'
Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
        mobjParentForm.Enabled = False
    End If
End Property

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Private Sub DisplayList()
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    
    If (Len(MaskEdBox4) > 0) And (Len(MaskEdBox6) > 0) Then
        With DataClass.DS
            .Field("Ente Emittente").FindValue = MaskEdBox4
            .Field("Società").FindValue = MaskEdBox6
            If DataClass.GetList Then
                If .HowMany > 0 Then
                    RefreshXArray XA, TDBGrid1, .List
                End If
            End If
        End With
    End If
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
'        Case fcAddNew
'            If State = fsEdit Then
'                DataClass.DS.Fields.ResetFields
'                ClearForm
'                SetState fsInsert
'            End If
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
'                Case fsInsert
'                    blnResult = InsertRecord
            End Select
            If blnResult Then
                DisplayList
                If State = fsInsert Then
                    SyncList DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
                End If
                SetState fsEdit
            End If
        Case fcUndo
'            ClearForm
            SetState fsNone
        Case fcPrint
            If State = fsEdit Then
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub GetRecord(IDATO As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    
    With DataClass.DS
        .Field("Società").FindValue = MaskEdBox6
        .Field("Ente Emittente").FindValue = MaskEdBox4
        .Field("Tipo Documento").FindValue = IDATO
        If DataClass.Find Then
            .RefreshClient
            SetState fsEdit
        Else
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub




Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    On Error GoTo ErrHandler
    If Not IsNull(DBGrid.bookmark) Then
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
    End If
    Exit Function
ErrHandler:
End Function

'Public Property Let IDEnteEmittente(NewValue As String)
'    mstrIDAEE = NewValue
'End Property
'
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


Private Sub SyncList(ByVal ID As String)
    Dim i As Integer
    Dim blnFound As Boolean
    Dim varList As Variant
    
    i = 0
    blnFound = False
    varList = DataClass.DS.List
    Do While i <= UBound(varList, 2) And (Not blnFound)
        If ID = varList(0, i) Then
            blnFound = True
        Else
            i = i + 1
        End If
    Loop
    If blnFound Then
        TDBGrid1.bookmark = i
    End If
End Sub

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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
                DisplayList
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
        Case fsEdit
            If Not Loading Then
                TDBGrid1.SetFocus
            End If
        Case fsNone
            ClearXArray XA, TDBGrid1
            If Not Loading Then
                With MaskEdBox6
                    .Enabled = True
                    .SetFocus
                End With
                 With MaskEdBox4
                    .Enabled = True
                End With
           End If
    End Select
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskField4
            strTitle = "Data aggiornamento"
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
    
    Loading = True
    
    Set GridArray = New ArrayGridCls
    
    Set DataClass = New CNumeratoreDocumentiAmm
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtFil1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        AXX.Find
'        mstrIDASC = AXX.DS.Field("Società Default").value
'        mstrIDAEE = AXX.DS.Field("Ente Emittente").value
    End With
    
    MaskEdBox6 = AXX.DS.Field("Società Default").value
    FieldUpdateDescription "0", Me, MaskEdBox6
    FieldSetDescription "0", Me, MaskEdBox6, MaskEdBox6

    MaskEdBox4 = AXX.DS.Field("Ente Emittente").value
    FieldUpdateDescription "1", Me, MaskEdBox4
    FieldSetDescription "1", Me, MaskEdBox4, MaskEdBox4

    DisplayList
    SetState fsEdit
    
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParentForm Is Nothing Then
        mobjParentForm.Enabled = True
    End If
    Set AXX = Nothing
    Set XA = Nothing
    Set ErrH = Nothing
    Set GridArray = Nothing
    Set DataClass = Nothing
    Set OIL2NDA0 = Nothing
End Sub





Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control

    With DataClass.DS
        If .ValidateClientField(mskField.Tag, mskField.Text) Then

'            mstrIDADE = mskField.Text
            FieldUpdateDescription "1", Me, mskField
            FieldSetDescription "1", Me, mskField.Text, mskField

            SetState fsEdit
            DisplayList
            TDBGrid1.bookmark = 0
        End If
    End With
End Sub

Private Sub MaskEdBox4_LostFocus()
    TrackLostFocus MaskEdBox4
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strID As String
    
    strID = GetSelectedID(TDBGrid1)
    GetRecord strID
End Sub





