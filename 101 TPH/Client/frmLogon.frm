VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGON"
   ClientHeight    =   2880
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlOpenDB 
      Left            =   5280
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleziona il database..."
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5880
      Top             =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   570
      Left            =   5040
      Picture         =   "frmLogon.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Connetti"
      Default         =   -1  'True
      Height          =   570
      Left            =   5040
      Picture         =   "frmLogon.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   1320
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Opzioni di connessione:"
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdSelectDB 
         Height          =   300
         Left            =   4440
         Picture         =   "frmLogon.frx":0210
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1365
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox cmbProvider 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1365
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1365
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label lblStep3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P&rovider:"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblStep3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome Utente:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblStep3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   0
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblStep3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data&base:"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblStep3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Server:"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2505
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "RES Data"
            TextSave        =   "RES Data"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8837
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private MSG_FIRM  As String

Private Sub CheckFields()
    If cmbProvider.ListIndex = 0 Then
        cmdOK.Enabled = (txtUID <> vbNullString) And (txtServer <> vbNullString) And (txtDatabase <> vbNullString)
    Else
        cmdOK.Enabled = (txtUID <> vbNullString) And (txtDatabase <> vbNullString)
    End If
End Sub

Private Sub cmbProvider_Click()
    Select Case cmbProvider.ListIndex
        Case 0, 2 'SQL Server
            cmdSelectDB.Visible = False
            txtServer.Enabled = True
        Case 1  'Jet
            cmdSelectDB.Visible = True
            txtServer.Enabled = False
    End Select
End Sub

Private Sub cmdCancel_Click()
    gmblnCancelConnect = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    DoConnect
End Sub

Public Function DoConnect() As Boolean
'    Set DA.ErrH = ErrH
    Set DataConnect.ErrH = ErrH
'    DA.DataEngine = CStr(cmbProvider.ListIndex)
    DataConnect.DataEngine = CStr(cmbProvider.ListIndex)
'    If DA.Connect(txtUID, txtPWD, "", txtServer, txtDatabase) Then
    If DataConnect.Connect(txtUID, txtPWD, "", txtServer, txtDatabase) Then
        gmblnConnectOk = True
        DoConnect = True
        Unload Me
    Else
        DoConnect = False
    End If
End Function

Private Sub cmdSelectDB_Click()
    With cdlOpenDB
        .fileName = txtDatabase
        .DefaultExt = ".mdb"
        .DialogTitle = "Seleziona il database... "
        .Filter = "JET Database (*.mdb)|*.mdb|Tutti i files (*.*)|*.*| "
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
        .ShowOpen
        txtDatabase = .fileName
    End With
End Sub

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_Load()
    Set ErrH = New CErrorHandler
    
    StatusBar1.Panels(1).Bevel = sbrNoBevel
    Me.MousePointer = vbDefault
    cmbProvider.AddItem "SQL Server"
    cmbProvider.AddItem "Jet"
    cmbProvider.AddItem "SQL Server 2000"
    MSG_FIRM = App.CompanyName & " "
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
End Sub


Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static blnInCapture As Boolean
    Dim lngReturn As Long
    If (x <= StatusBar1.Panels(1).Width) And (y <= (StatusBar1.Height) And (y > 0)) Then
        If Not blnInCapture Then
            lngReturn = SetCapture(StatusBar1.hwnd)
            blnInCapture = True
            StatusBar1.Panels(1).Bevel = sbrRaised
            Timer1.Enabled = True
        End If
    Else
        If blnInCapture Then
            lngReturn = ReleaseCapture
            blnInCapture = False
            Timer1.Enabled = False
            StatusBar1.Panels(1).Bevel = sbrNoBevel
            StatusBar1.Panels(1).Text = MSG_FIRM
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Static strMsg As String
    Static InCycle As Boolean
    Static i As Integer
    Static intMsgLen As Integer
    Dim strOut As String
    
    If Not InCycle Then
        strMsg = MSG_FIRM
        intMsgLen = Len(strMsg)
        InCycle = True
        i = 1
    Else
        If i < intMsgLen Then
            strOut = Mid$(strMsg, i) & Left$(strMsg, i - 1)
        Else
            i = 1
            strOut = MSG_FIRM
        End If
        i = i + 1
        StatusBar1.Panels(1).Text = strOut
        DoEvents
    End If
    
End Sub

Private Sub txtDatabase_Change()
    CheckFields
End Sub

Private Sub txtServer_Change()
    CheckFields
End Sub


Private Sub txtUID_Change()
    CheckFields
End Sub


