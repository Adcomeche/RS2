VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2FPZ0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FASCE PREZZI"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9180
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   9135
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5280
         Picture         =   "OIL2FPZ0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
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
         Index           =   5
         Left            =   5280
         Picture         =   "OIL2FPZ0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Riga"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2FPZ0.frx":0204
         TabIndex        =   6
         Tag             =   "EDITLINK"
         Top             =   360
         Width           =   1815
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
         Picture         =   "OIL2FPZ0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2FPZ0.frx":0408
         TabIndex        =   5
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
         Left            =   1920
         Picture         =   "OIL2FPZ0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2FPZ0.frx":060C
         TabIndex        =   4
         Tag             =   "SAVE"
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2FPZ0.frx":070E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2FPZ0.frx":0810
         TabIndex        =   15
         Tag             =   "ADD"
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
         Left            =   7440
         Picture         =   "OIL2FPZ0.frx":0912
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7320
         Picture         =   "OIL2FPZ0.frx":0A14
         TabIndex        =   7
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2FPZ0.frx":0B16
         TabIndex        =   21
         Tag             =   "FIND"
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5340
      Width           =   9180
      _ExtentX        =   16193
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
            Object.Width           =   10742
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.38"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "05/01/00"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9135
      Begin VB.Frame Frame3 
         Caption         =   "Dettaglio"
         Height          =   3255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   8895
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   2655
            Left            =   240
            OleObjectBlob   =   "OIL2FPZ0.frx":0C18
            TabIndex        =   19
            Tag             =   "DETAIL"
            Top             =   360
            Width           =   8415
         End
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Annullato"
         Height          =   195
         Index           =   5
         Left            =   7680
         TabIndex        =   17
         Tag             =   "Annullato"
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Tag             =   "Codice"
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
         Left            =   3840
         TabIndex        =   3
         Tag             =   "Descrizione"
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8760
         Top             =   120
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
               Picture         =   "OIL2FPZ0.frx":5DCA
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2FPZ0.frx":5EDC
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2FPZ0.frx":5FEE
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7080
         Tag             =   "Descrizione"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2400
         Tag             =   "Codice"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2FPZ0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Nome Form      : OILFPZ0
' Descrizione    : Gestione Fasce di Prezzi
' Data Creazione : 22/11/99
' Autore         : Fabio Mozzi
'
Option Explicit

Private Const MAX_COLS = 14

Public GridArrayDetail As ArrayGridCls
Public WithEvents DataClass As CFasciaPrezzi
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CFasciaPrezziDettaglio
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property
Private Sub DisplayDetail()
    Me.MousePointer = vbHourglass
    With DetailClass
        .DS.Fields("0").FindValue = DataClass.DS.Fields("Codice").value
        .GetList
        RefreshUnboundGrid TDBGrid1, .DS.List, GridArrayDetail, _
                           .DS.HowMany, 6
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
                    DetailClass.DS.Fields.ResetFields
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
    ReDim varReportsList(0 To 2, 0 To 2) As Variant
    Dim strSelectionFormula As String
    
    Select Case Cmd
    
        Case fcLinkRefresh
            DetailClass.DS.Fields.ResetFields
            DisplayDetail
        
        Case fcAddNew
'            If State = fsedit Then
                DataClass.DS.Fields.ResetFields
                ClearForm
                ClearUnboundGrid TDBGrid1, GridArrayDetail, 6
                SetState fsInsert
'            End If
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
            End If
            
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
            
        Case fcFind
             
            ShowSearchForm Me, "Codice", DataClass, 1
            'DoSearch "ASC", Me, "0", 1
        Case fcPrint
            
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcPREZFPD)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                        .SetIdFasciaPrezzi = DataClass.DS.Fields(1).value
                        If DetailClass.DS.HowMany > 0 Then
                            LinkForm.GetRecord DataClass.DS.Fields(1).value, _
                                           GetSelectedID(TDBGrid1, 0), _
                                           GetSelectedID(TDBGrid1, 1), _
                                           GetSelectedID(TDBGrid1, 2)
                            .SetState fsEdit
                         Else
                            .SetState fsInsert
                         End If
                    Me.Enabled = False
                End With
            End If
         
        Case fcExit
            Unload Me
    End Select
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
Public Function GetRecord(strID As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
            If .HowMany > 0 Then
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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
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
        Case fsEdit
            If Not Loading Then
                mskField2.SetFocus
            End If
            'mskField2.SetFocus
        Case fsInsert
            If Not Loading Then
                
                Command6.Enabled = False
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
    
    ReDim ColTags(1 To 6) As Variant
    
    Loading = True
    Set GridArrayDetail = New ArrayGridCls
    Set DataClass = New CFasciaPrezzi
    Set DetailClass = New CFasciaPrezziDettaglio
    
    ColTags(1) = "1"
    ColTags(2) = "2"
    ColTags(3) = "3"
    ColTags(4) = "4"
    ColTags(5) = "5"
    ColTags(6) = "15"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags, False
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
'        Set DetailClass.DS.DA = DA
        Set DetailClass.DS.DA = DataConnect.DA
        Set DetailClass.DS.ErrH = ErrH
    End With
    SetState fsNone
    Loading = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not CheckExit(State)
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
        End With
    End If
    Set OIL2FPZ0 = Nothing
End Sub
Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub
Private Function GetSelectedID(DBGrid As TDBGrid, Columns As Integer) As Variant
    If IsNull(DBGrid.Bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = GridArrayDetail.value(DBGrid.Bookmark, Columns)
    End If
End Function
