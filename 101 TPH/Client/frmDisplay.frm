VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GESTIONE ERRORI"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2775
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   5535
      Begin VB.TextBox txtInfo 
         Height          =   885
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codice:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblErrCode 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrizione:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblErrDesc 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Livello di errore:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblSeverity 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Informazioni:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtSuppInfo 
         Height          =   1245
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label lblNative 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codice Nativo:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblRoutine 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Informazioni supplementari:"
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Modulo:"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblModule 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Routine:"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Chiudi"
      Default         =   -1  'True
      Height          =   975
      Left            =   6000
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Descrizione dell'errore"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Informazioni aggiuntive"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3450
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10795
            Key             =   "Msg"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintCurFrame As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Load()
'    Call ForceForegroundWindow(Me.hwnd)
    mintCurFrame = 1
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = mintCurFrame Then Exit Sub
    ' Otherwise, hide old frame, show new.
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    Frame1(mintCurFrame).Visible = False
    ' Set mintCurFrame to new value.
    mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


