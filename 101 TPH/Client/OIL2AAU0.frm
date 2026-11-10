VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2AAU0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AUTISTI"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   22
      Top             =   4320
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1920
         Picture         =   "OIL2AAU0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2AAU0.frx":0102
         TabIndex        =   34
         Tag             =   "FIND"
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
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2AAU0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2AAU0.frx":0306
         TabIndex        =   28
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
         Index           =   3
         Left            =   240
         Picture         =   "OIL2AAU0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2AAU0.frx":050A
         TabIndex        =   29
         Tag             =   "PRINT"
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
         Index           =   1
         Left            =   1920
         Picture         =   "OIL2AAU0.frx":060C
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
         Left            =   1800
         Picture         =   "OIL2AAU0.frx":070E
         TabIndex        =   31
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
         Index           =   2
         Left            =   7680
         Picture         =   "OIL2AAU0.frx":0810
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2AAU0.frx":0912
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2AAU0.frx":0A14
         TabIndex        =   32
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2AAU0.frx":0B16
         TabIndex        =   30
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   5610
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
            TextSave        =   "16:21"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15/03/2013"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton btn_cancelarpalm 
         Caption         =   "Cancelar Viajes Palm "
         Height          =   855
         Left            =   7320
         TabIndex        =   39
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "5"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   18
         Tag             =   "8"
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Tag             =   "1"
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Tag             =   "2"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField5 
         Height          =   255
         Left            =   6840
         TabIndex        =   11
         Tag             =   "4"
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Tag             =   "7"
         Top             =   2520
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField6 
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Tag             =   "6"
         Top             =   2160
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
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
               Picture         =   "OIL2AAU0.frx":0C18
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2AAU0.frx":0D2A
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2AAU0.frx":0E3C
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Tag             =   "9"
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   1215
         Left            =   1920
         OleObjectBlob   =   "OIL2AAU0.frx":0F4E
         TabIndex        =   36
         Tag             =   "MULTISOC"
         Top             =   2880
         Width           =   4935
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Tag             =   "10"
         Top             =   1800
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
         Index           =   9
         Left            =   5280
         Tag             =   "10"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Base"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   38
         Top             =   1800
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
         Left            =   2520
         TabIndex        =   37
         Tag             =   "10"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   4080
         Tag             =   "9"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero N.I.F."
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   35
         Top             =   1080
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
         Left            =   2520
         TabIndex        =   21
         Tag             =   "3"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Vettore"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5280
         Tag             =   "3"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefono"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   6960
         Tag             =   "6"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefono interno"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   6960
         Tag             =   "7"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Disponibilità"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   9000
         Tag             =   "5"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N&umero Patente"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   9000
         Tag             =   "4"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N&ome abbr."
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   8160
         Tag             =   "2"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5280
         Tag             =   "1"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2640
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Nome"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2AAU0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3

Public WithEvents DataClass As CAutista
Attribute DataClass.VB_VarHelpID = -1
Public GridArray As ArrayGridCls
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
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
    ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
End Sub


Private Sub DisplayMultiSoc()
    RefreshUnboundGrid TDBGrid2, DataClass.DS.MultiSoc.DS.List, GridArray, _
                       DataClass.DS.MultiSoc.DS.HowMany, MAX_COLS
    
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    
    Select Case Cmd
        Case fcAddNew
        '    If (State = fsEdit) Or (State = fsNone) Then
        '        DataClass.DS.Fields.ResetFields True
        '        ClearForm
        '        SetState fsInsert
        '    End If
            MsgBox ("Esta opción se realiza desde Resoil WEB")
        Case fcSave
        '    Select Case State
        '        Case fsEdit
        '            blnResult = UpdateRecord
        '        Case fsInsert
        '            blnResult = InsertRecord
        '        Case Else
        '            blnResult = False
        '    End Select
        '    If blnResult Then
        '        SetState fsNone
        '    End If
            MsgBox ("Esta opción se realiza desde Resoil WEB")
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcFind
            DoSearch "AAU", Me, "0", 1
        Case fcPrint
            If (State = fsEdit) Or (State = fsNone) Then
                varReportsList(0, 0) = "Stampa elenco Autisti"
                varReportsList(1, 0) = "oil2aau1.rpt"
                varReportsList(2, 0) = 1
                CRPrint varReportsList, "", Me
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
                If .RefreshClient Then
                    DisplayMultiSoc
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


'Private Sub SyncList(ByVal ID As String)
'    Dim I As Integer
'    Dim blnFound As Boolean
'    Dim varList As Variant
'
'    I = 0
'    blnFound = False
'    varList = DataClass.DS.List
'    Do While I <= UBound(varList, 2) And (Not blnFound)
'        If ID = varList(0, I) Then
'            blnFound = True
'        Else
'            I = I + 1
'        End If
'    Loop
'    If blnFound Then
'        TDBGrid1.Bookmark = I
'    End If
'End Sub

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
'                DisplayList
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
            ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
        Case fsEdit
            mskField2.SetFocus
        Case fsInsert
            mskField1.SetFocus
    End Select
End Sub



Private Sub btn_cancelarpalm_Click()

Dim rowbuff As Variant
Dim lng As Long
Dim SQL_id As String
Dim Service As VLocalize.VLocalize
Set Service = New VLocalize.VLocalize
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim COD As String
  
 
COD = InputBox(" Introduce código para proseguir.", "Cancelar Viajes Palm")
If ((COD) = mskField1.Text) Then
        SQL_id = "SELECT AAUNRCEL From ANAGAAU0 WHERE (AAUIDAAU = ?)"
        
        Service.serviceLocation = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
        Service.serviceTimeOut = 60
        Service.serviceUsername = "wstph"
        Service.servicePassword = "swtph"
        Service.setServerID = DataConnect.ConnServer
        Service.setServerPassword = DataConnect.ConnPWD
        Service.setServerUser = DataConnect.ConnUserName
        
        If Not DA.QueryCommand(SQL_id, rowbuff, lng, mskField1.Text) Then
        End If

        
        If lng > 0 And Not IsNull(rowbuff(0, 0)) And rowbuff(0, 0) <> "" Then
        Service.Open
            If (Service.AnulaToaLaPam(rowbuff(0, 0))) Then
                If (MsgBox("Servicios eliminados correctamente", vbOKOnly, "")) Then
                End If
            Else
                If (MsgBox("No se ha podido eliminar el viaje", vbOKOnly, "")) Then
                End If
            End If
        Else
                If (MsgBox("Esta Ud. intentando eliminar servicios de la PDA un Conductor que no tiene assignada ninguna PDA ", vbOKOnly, "")) Then
                End If
        End If
        
        If Service.isOpen Then
           Service.Close
        End If
        
Else
        MsgBox ("Codigo Incorrecto")

End If
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
    Set DataClass = New CAutista
    With DataClass.DS
        Set .FormHandle = Me
        Set .DA = DataConnect.DA
        Set Soc = New CSocieta
        Set Soc.DS.DA = DataConnect.DA
        Soc.GetList True
        .ConfigureInterface
        ConfigureObjects Me
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
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2AAU0 = Nothing
End Sub





Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1

End Sub


