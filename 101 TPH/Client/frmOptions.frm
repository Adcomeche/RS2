VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Height          =   3585
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5415
         Begin VB.TextBox txtOption 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   3855
         End
         Begin VB.CommandButton cmdOpenFile 
            Height          =   300
            Index           =   1
            Left            =   4080
            Picture         =   "frmOptions.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdOpenFile 
            Height          =   300
            Index           =   0
            Left            =   4080
            Picture         =   "frmOptions.frx":010E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtOption 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cartella dei reports:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Immagine di sfondo:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2775
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample1 
         Height          =   3585
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5415
         Begin VB.CheckBox chkOption 
            Caption         =   "Permetti l'apertura contemporanea di più finestre"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Applica"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Generale"
            Key             =   "Group1"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Opzioni generali"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Files e cartelle"
            Key             =   "Group2"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Opzioni relative alle posizioni dei files di lavoro"
            ImageVarType    =   2
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
   Begin MSComDlg.CommonDialog cdlOpenFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Seleziona il database..."
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadOptions()
    chkOption(0).value = IIf(gmblnMultipleWindows, vbChecked, vbUnchecked)
    txtOption(0) = gmstrPicFileName
    txtOption(1) = gmstrReportsPath
End Sub

Private Sub SetOptions()
    Me.MousePointer = vbHourglass
    gmblnMultipleWindows = (chkOption(0).value = vbChecked)
    gmstrPicFileName = txtOption(0)
    If Right$(txtOption(1), 1) = "\" Then
        txtOption(1) = Left$(txtOption(1), Len(txtOption(1)) - 1)
    End If
    gmstrReportsPath = txtOption(1)
    SetBackPic
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdApply_Click()
    SetOptions
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SetOptions
    Unload Me
End Sub

Private Sub cmdOpenFile_Click(Index As Integer)
    With cdlOpenFile
'        .filename = txtOption(Index)
        Select Case Index
            Case 0
                .DialogTitle = "Seleziona l'immagine da usare come sfondo... "
                .Filter = "Immagini JPEG (*.jpg)|*.jpg|Immagini Bitmap (*.bmp)|*.bmp|Tutti i files (*.*)|*.*| "
        End Select
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
        .ShowOpen
         txtOption(Index) = .filename
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
'    'center the form
'    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    LoadOptions
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub
