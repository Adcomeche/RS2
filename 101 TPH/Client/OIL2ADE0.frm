VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2ADE0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPOSITI"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9495
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   5520
      Width           =   9495
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5280
         Picture         =   "OIL2ADE0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2ADE0.frx":0102
         TabIndex        =   25
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
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2ADE0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
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
         Picture         =   "OIL2ADE0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   795
         Visible         =   0   'False
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
         Left            =   7680
         Picture         =   "OIL2ADE0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
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
         Picture         =   "OIL2ADE0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
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
         Picture         =   "OIL2ADE0.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2ADE0.frx":070E
         TabIndex        =   23
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ADE0.frx":0810
         TabIndex        =   24
         Tag             =   "PRINT"
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2ADE0.frx":0912
         TabIndex        =   26
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2ADE0.frx":0A14
         TabIndex        =   22
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ADE0.frx":0B16
         TabIndex        =   21
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   6375
      Width           =   9495
      _ExtentX        =   16748
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
            Object.Width           =   11298
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11.33"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10/07/2009"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9495
      Begin VB.Frame Frame3 
         Height          =   3855
         Index           =   1
         Left            =   360
         TabIndex        =   37
         Top             =   1200
         Width           =   8895
         Begin VB.ComboBox comboInvioKMCepsa 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "11"
            Top             =   2700
            Width           =   3315
         End
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   1740
            TabIndex        =   3
            Tag             =   "1"
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField3 
            Height          =   255
            Left            =   1740
            TabIndex        =   5
            Tag             =   "2"
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   3780
            TabIndex        =   9
            Tag             =   "3"
            Top             =   960
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField6 
            Height          =   255
            Left            =   8040
            TabIndex        =   11
            Tag             =   "4"
            Top             =   960
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField4 
            Height          =   255
            Left            =   1740
            TabIndex        =   7
            Tag             =   "5"
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField9 
            Height          =   255
            Left            =   1740
            TabIndex        =   15
            Tag             =   "6"
            Top             =   1680
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField7 
            Height          =   255
            Left            =   1740
            TabIndex        =   17
            Tag             =   "7"
            Top             =   2040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField8 
            Height          =   255
            Left            =   1740
            TabIndex        =   19
            Tag             =   "8"
            Top             =   2400
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
            Left            =   1740
            TabIndex        =   13
            Tag             =   "10"
            Top             =   1320
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   7320
            TabIndex        =   47
            Tag             =   "12"
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   255
            Left            =   7320
            TabIndex        =   49
            Tag             =   "13"
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Coordinata Y"
            Height          =   255
            Index           =   12
            Left            =   5640
            TabIndex        =   50
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   8520
            Tag             =   "13"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Coordinata X"
            Height          =   255
            Index           =   11
            Left            =   5640
            TabIndex        =   48
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   8520
            Tag             =   "12"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   10
            Left            =   5100
            Tag             =   "11"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Trasm. KM CEPSA"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   45
            Top             =   2760
            Width           =   1515
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   5100
            Tag             =   "10"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Presidio"
            Height          =   255
            Index           =   9
            Left            =   780
            TabIndex        =   12
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Descrizione"
            Height          =   255
            Index           =   1
            Left            =   660
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   5100
            Tag             =   "1"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   6780
            Tag             =   "2"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Indirizzo"
            Height          =   255
            Index           =   2
            Left            =   900
            TabIndex        =   4
            Top             =   600
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   6780
            Tag             =   "3"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Località"
            Height          =   255
            Index           =   3
            Left            =   3060
            TabIndex        =   8
            Top             =   960
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   8520
            Tag             =   "4"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pro&vincia"
            Height          =   255
            Index           =   4
            Left            =   7200
            TabIndex        =   10
            Top             =   960
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   2700
            Tag             =   "5"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "C&AP"
            Height          =   255
            Index           =   5
            Left            =   1140
            TabIndex        =   6
            Top             =   960
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   5100
            Tag             =   "6"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Società"
            Height          =   255
            Index           =   8
            Left            =   900
            TabIndex        =   14
            Top             =   1680
            Width           =   735
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
            Left            =   2340
            TabIndex        =   43
            Tag             =   "6"
            Top             =   1680
            Width           =   2655
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
            Left            =   2340
            TabIndex        =   42
            Tag             =   "7"
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Zona"
            Height          =   255
            Index           =   6
            Left            =   780
            TabIndex        =   16
            Top             =   2040
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   5100
            Tag             =   "7"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label lblField2 
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
            Left            =   2340
            TabIndex        =   41
            Tag             =   "8"
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Raggruppamento"
            Height          =   255
            Index           =   7
            Left            =   300
            TabIndex        =   18
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   5100
            Tag             =   "8"
            Top             =   2400
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   8895
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   3495
            Left            =   120
            OleObjectBlob   =   "OIL2ADE0.frx":0C18
            TabIndex        =   36
            Tag             =   "DETAIL"
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Index           =   3
         Left            =   360
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   8895
         Begin TrueDBGrid60.TDBGrid TDBGrid2 
            Height          =   3375
            Left            =   1320
            OleObjectBlob   =   "OIL2ADE0.frx":6216
            TabIndex        =   40
            Tag             =   "MULTISOC"
            Top             =   360
            Width           =   5415
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   20
         Tag             =   "9"
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1440
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
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4335
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7646
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati &Deposito"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Magazzini di Deposito"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Codici &Per Società"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
               Picture         =   "OIL2ADE0.frx":A490
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ADE0.frx":A5A2
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ADE0.frx":A6B4
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2160
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2ADE0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GridArray As ArrayGridCls
Public GridArrayDetail As ArrayGridCls
Public WithEvents DataClass As CDeposito
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta
Public DetailClass As CMagazzinoDeposito
Private Tve As CTipoVendita
Private mintCurFrame As Integer
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Public Function InsertDetail() As Boolean
    InsertDetail = DetailClass.Insert
End Function

Private Sub DisplayDetail()
    DetailClass.DS.Fields("1").FindValue = DataClass.DS.Fields(DataClass.DS.KeyFields("1")).value
    DetailClass.GetList
    RefreshUnboundGrid TDBGrid1, DetailClass.DS.List, GridArrayDetail, _
                       DetailClass.DS.HowMany, 6

End Sub


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
    ClearUnboundGrid TDBGrid1, GridArrayDetail, 6
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim C As Control
    
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
            End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcPrint
            If (State = fsEdit) Or (State = fsNone) Then
            End If
        Case fcFind
            'ShowSearchForm Me, "0", DataClass, 1
            DoSearch "ADE", Me, "0", 1
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





Private Sub DisplayMultiSoc()
    RefreshUnboundGrid TDBGrid2, DataClass.DS.MultiSoc.DS.List, GridArray, _
                       DataClass.DS.MultiSoc.DS.HowMany, MAX_COLS
    
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
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
            mintCurFrame = 1
            TabStrip1.Enabled = False
            ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
            ClearUnboundGrid TDBGrid1, GridArrayDetail, 6
        Case fsEdit
            TabStrip1.Enabled = True
            If TabStrip1.SelectedItem.Index = 1 Then
                mskField2.SetFocus
            End If
        Case fsInsert
            TabStrip1.Enabled = True
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
            mskField1.SetFocus
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
    mintCurFrame = 1
    Set GridArray = New ArrayGridCls
    Set GridArrayDetail = New ArrayGridCls
    Set DataClass = New CDeposito
    Set DetailClass = New CMagazzinoDeposito
    
    ColTags(1) = "0"
    ColTags(2) = "2"
    ColTags(3) = "3"
    ColTags(4) = "4"
    ColTags(5) = "5"
    ColTags(6) = "6"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        Set Soc = New CSocieta
        Set Soc.DS.DA = DataConnect.DA
        Soc.GetList True
        .ConfigureInterface
        ConfigureObjects Me, ColTags, False
        'Set .MultiSoc.DS.DA = DA
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
'        Set DetailClass.DS.DA = DA
        Set DetailClass.DS.DA = DataConnect.DA
        Set DetailClass.DS.ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
        Set Tve = New CTipoVendita
'        Set Tve.DS.DA = DA
        Set Tve.DS.DA = DataConnect.DA
    End With
    
    SetValueItems TDBGrid1.Columns(2), Soc
    SetValueItems TDBGrid1.Columns(3), Tve
    
    SetState fsNone
    Loading = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set Tve = Nothing
    Set OIL2ADE0 = Nothing
End Sub

Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub



Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame3(TabStrip1.SelectedItem.Index).Visible = True
        Frame3(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


