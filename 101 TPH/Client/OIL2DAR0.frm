VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2DAR0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIGA DOCUMENTO AMMINISTRATIVO"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   10110
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   33
      Top             =   4560
      Width           =   10095
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "OIL2DAR0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   315
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
         Picture         =   "OIL2DAR0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   315
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
         Picture         =   "OIL2DAR0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2DAR0.frx":0306
         TabIndex        =   37
         Tag             =   "UNDO"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7800
         Picture         =   "OIL2DAR0.frx":0408
         TabIndex        =   38
         Tag             =   "EXIT"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2DAR0.frx":050A
         TabIndex        =   39
         Tag             =   "SAVE"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10095
      Begin VB.Frame fraMaster 
         Caption         =   "Dati Riga 1"
         ForeColor       =   &H00FF0000&
         Height          =   2955
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1440
         Width           =   8415
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   255
            Left            =   1560
            TabIndex        =   4
            Tag             =   "4"
            Top             =   720
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField4 
            Height          =   255
            Left            =   1560
            TabIndex        =   40
            Tag             =   "10"
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField10 
            Height          =   255
            Left            =   1560
            TabIndex        =   41
            Tag             =   "11"
            Top             =   1800
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
            Left            =   1560
            TabIndex        =   45
            Tag             =   "7"
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   1560
            TabIndex        =   47
            Tag             =   "22"
            Top             =   2160
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   1560
            TabIndex        =   49
            Tag             =   "31"
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
            Index           =   0
            Left            =   5160
            Tag             =   "11"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   7920
            Tag             =   "31"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Addebito"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   51
            Top             =   360
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
            Left            =   2280
            TabIndex        =   50
            Tag             =   "31"
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label1 
            Caption         =   "Imponibile"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   48
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Quantità"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   46
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Prezzo Unitario"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   44
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Aliquota IVA"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   43
            Top             =   1800
            Width           =   1095
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
            Left            =   2280
            TabIndex        =   42
            Tag             =   "11"
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Descrizione riga"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame fraMaster 
         Caption         =   "Dati Riga 2"
         ForeColor       =   &H00FF0000&
         Height          =   2955
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   8415
         Begin MSMask.MaskEdBox mskField6 
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Tag             =   "12"
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox14 
            Height          =   255
            Left            =   1440
            TabIndex        =   8
            Tag             =   "15"
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Tag             =   "16"
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   1440
            TabIndex        =   10
            Tag             =   "17"
            Top             =   2160
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   255
            Left            =   1440
            TabIndex        =   6
            Tag             =   "13"
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox18 
            Height          =   255
            Left            =   1440
            TabIndex        =   7
            Tag             =   "14"
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Tag             =   "23"
            Top             =   2520
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Importo Sconto"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sconto % 3"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Sconto % 2"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Sconto Valore 3"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   21
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Sconto Valore 2"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Sconto % 1"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Sconto Valore 1"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame fraKey 
         Caption         =   "Documento Amministrativo"
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   8415
         Begin VB.CheckBox Check1 
            Caption         =   "Annullato"
            Height          =   255
            Index           =   1
            Left            =   7200
            TabIndex        =   26
            Tag             =   "24"
            Top             =   360
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mskField1 
            Height          =   255
            Left            =   4080
            TabIndex        =   2
            Tag             =   "2"
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            Height          =   255
            Left            =   1560
            TabIndex        =   1
            Tag             =   "1"
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox15 
            Height          =   255
            Left            =   1560
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
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   6240
            TabIndex        =   3
            Tag             =   "3"
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Riga"
            Height          =   255
            Index           =   19
            Left            =   5640
            TabIndex        =   31
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Progressivo"
            Height          =   255
            Index           =   18
            Left            =   3000
            TabIndex        =   30
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Anno"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Ente Emittente"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   1215
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
            TabIndex        =   27
            Tag             =   "0"
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame fraSelect 
         Caption         =   "Selezione Vista"
         ForeColor       =   &H00FF0000&
         Height          =   2955
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
         Begin MSComctlLib.TabStrip tspSelect 
            CausesValidation=   0   'False
            Height          =   1215
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   2143
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dati Riga &1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dati Riga &2"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9120
         Top             =   4680
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
               Picture         =   "OIL2DAR0.frx":060C
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2DAR0.frx":071E
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2DAR0.frx":0830
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   5295
      Width           =   10110
      _ExtentX        =   17833
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
            Object.Width           =   12383
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18.19"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "04/02/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0942
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0A54
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0B66
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0C78
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0D8A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0E9C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":0FAE
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":10C0
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":11D2
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":12E4
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":13F6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAR0.frx":1508
            Key             =   "Camera"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2DAR0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3

Public WithEvents DataClass As CRigaDocumentoAmministrativo
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public mstrIDADZ As String
Public mvarData As Variant
'Public mstrTipoOrdine As String
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private mintCurFrame As Integer
Private mblnDataChange As Boolean

Private ASZ As CStoricoPrezzi
Private APR As CProdotto
Private ACL As CCliente
' (PM) 03/01/06
'Private ACF As CConfezioneProdotto
Private DAM As CDocumentoAmministrativo

Private Sub CalcolaImponibile()
    Dim varImponibile As Variant
    Dim varQuantità As Variant
    
'    Select Case UCase(MaskEdBox19)
'        Case "KG"
'            varQuantità = MaskEdBox4
'        Case "L"
'            varQuantità = MaskEdBox5
'        Case "PZ"
            varQuantità = MaskEdBox6
'        Case "CNF"
'            varQuantità = MaskEdBox7
'        Case Else
'            varQuantità = 0
'    End Select
    If Not IsNumeric(MaskEdBox1) Then
        MaskEdBox1 = 0
    End If
    If Not IsNumeric(MaskEdBox14) Then
        MaskEdBox14 = 0
    End If
    If Not IsNumeric(MaskEdBox16) Then
        MaskEdBox16 = 0
    End If
    varImponibile = DataClass.ComputeImponibileRiga(mskField4, varQuantità, MaskEdBox14, MaskEdBox1, MaskEdBox16)
    If Not IsEmpty(varImponibile) Then
        MaskEdBox8 = varImponibile
    End If
End Sub

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(MaskEdBox15, MaskEdBox13, mskField1, MaskEdBox20) Then
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
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
                    If blnResult Then
                        mblnDataChange = True
'                        SetState fsEdit
                        DoCommand fcExit
                    End If
                Case fsInsert
                    blnResult = InsertRecord
                    If blnResult Then
                        SetState fsInsert
                        mblnDataChange = True
                    End If
            End Select
        Case fcUndo
            Select Case State
                Case fsEdit
                    SetState fsEdit
                Case fsInsert
                    SetState fsInsert
            End Select
        Case fcExit
            If mblnDataChange Then
                If mobjParentForm.DataClass.ComputeTotali Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
    End Select
End Sub

Public Function GetRecord(strIDEnteEmit As String, AnnoDoc As String, ProgDoc As String, intIDRiga As Integer) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .FieldS(.KeyFields(1)).FindValue = strIDEnteEmit
        .FieldS(.KeyFields(2)).FindValue = AnnoDoc
        .FieldS(.KeyFields(3)).FindValue = ProgDoc
        .FieldS(.KeyFields(4)).FindValue = intIDRiga
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                
                Else
                    MsgBox "Errore"
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
    
    CalcolaImponibile
    
    DataClass.DS.ScanClient 1
    
    DAM.DS.FieldS("0").FindValue = DataClass.DS.FieldS("0").NewValue
    DAM.DS.FieldS("1").FindValue = DataClass.DS.FieldS("1").NewValue
    DAM.DS.FieldS("2").FindValue = DataClass.DS.FieldS("2").NewValue
    DataClass.DS.FieldS("0").Updated = True
    DataClass.DS.FieldS("1").Updated = True
    DataClass.DS.FieldS("2").Updated = True
    
    DAM.Find
    If Not DAM.IsUpdatable Then
        DataClass.DS.DisplayMessage "il Documento non è modificabile", Warning
        GoTo ErrHandler
    End If
    
    If DataClass.DS.FieldS.Validate Then
            
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
    Dim strMsg As String
    
    Dim C As Control
    
    'Scansione della finestra
    If Not DataClass.DS.ScanClient(0) Then
        GoTo ErrHandler
    End If
    
    DAM.DS.FieldS("0").FindValue = DataClass.DS.FieldS("0").value
    DAM.DS.FieldS("1").FindValue = DataClass.DS.FieldS("1").value
    DAM.DS.FieldS("2").FindValue = DataClass.DS.FieldS("2").value
    DAM.Find
    If Not DAM.IsUpdatable Then
        DataClass.DS.DisplayMessage "il Documento non è modificabile", Warning
        GoTo ErrHandler
    End If
'    If DataClass.DS.Field("UM Acquisto").NewValue = "CNF" Then
'        DataClass.DS.Field("Quantita Richiesta Kg").NewValue = ACF.GetTotalPzz(DataClass.DS.Field("Prodotto").NewValue, DataClass.DS.Field("Imballo").NewValue, DataClass.DS.Field("Quantita Richiesta Cnf").NewValue)
'        DataClass.DS.Field("Quantita Richiesta L").NewValue = ACF.GetTotalPzz(DataClass.DS.Field("Prodotto").NewValue, DataClass.DS.Field("Imballo").NewValue, DataClass.DS.Field("Quantita Richiesta Cnf").NewValue)
'        DataClass.DS.Field("Quantita Richiesta Pz").NewValue = ACF.GetTotalPzz(DataClass.DS.Field("Prodotto").NewValue, DataClass.DS.Field("Imballo").NewValue, DataClass.DS.Field("Quantita Richiesta Cnf").NewValue)
'    End If
    
    If DataClass.DS.FieldS.Validate Then
        DataClass.BeginSave blnConcurError
            If blnConcurError Then
                If MsgBox(MSG_CONCUR_ERR, _
                           vbOKCancel + vbExclamation + vbDefaultButton2, _
                          "Attenzione") = vbOK Then
                    GoTo EndProc
                Else
                    DataClass.DS.DA.EndTrans ETROLLBACK
                    GetRecord DataClass.DS.FieldS("0").value, _
                              DataClass.DS.FieldS("1").value, _
                              DataClass.DS.FieldS("2").value, _
                              DataClass.DS.FieldS("3").value
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
    Dim lngIDRiga As Long
    Dim varIDEnteEmittente As Variant
    Dim varAnno As Variant
    Dim varIDDocumento As Variant
    
    DataClass.DS.SetClientState NewState
    Me.tspSelect.Enabled = True
    Me.tspSelect.Tabs(1).Selected = True
    Select Case NewState
        Case fsEdit
            If MaskEdBox12.Enabled Then
                MaskEdBox12.SetFocus
            End If
        Case fsInsert
            If Len(MaskEdBox20) > 0 Then
                lngIDRiga = MaskEdBox20
            End If
            varIDEnteEmittente = MaskEdBox15
            varAnno = MaskEdBox13
            varIDDocumento = mskField1
            ClearForm
            MaskEdBox15 = varIDEnteEmittente
            MaskEdBox13 = varAnno
            mskField1 = varIDDocumento
            MaskEdBox20 = lngIDRiga + 1
            If MaskEdBox12.Enabled Then
                MaskEdBox12.SetFocus
            End If
'            Combo2.ListIndex = 0
    End Select
End Sub


'Private Sub DataClass_Command(Cmd As RESLIB05.FormCommand)
'    DoCommand Cmd
'
'End Sub


Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Loading = True
    
    mintCurFrame = 1
    Set DataClass = New CRigaDocumentoAmministrativo
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set ASZ = New CStoricoPrezzi
    With ASZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
' (PM) 03/01/06
'    Set ACF = New CConfezioneProdotto
'    With ACF.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set ASZ = Nothing
    Set APR = Nothing
' (PM) 03/01/06
'    Set ACF = Nothing
    Set DAM = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
'            Me.Visible = False
'            DoEvents
            .DoCommand fcLinkRefresh
'            If State = fsInsert Then
                .SetState fsEdit
'            End If
        End With
    End If
    Set OIL2DAR0 = Nothing
End Sub





Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property


Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    CalcolaImponibile
End Sub


Private Sub MaskEdBox14_Validate(Cancel As Boolean)
    CalcolaImponibile
End Sub


Private Sub MaskEdBox16_Validate(Cancel As Boolean)
    CalcolaImponibile
End Sub


Private Sub MaskEdBox2_LostFocus()
'    GetProductData
End Sub




Private Sub MaskEdBox6_Validate(Cancel As Boolean)
    CalcolaImponibile
End Sub




Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub



Private Sub mskField4_Validate(Cancel As Boolean)
    CalcolaImponibile
End Sub


Private Sub tspSelect_Click()
    If Me.tspSelect.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
    fraMaster(Me.tspSelect.SelectedItem.Index).Visible = True
    fraMaster(mintCurFrame).Visible = False
    mintCurFrame = Me.tspSelect.SelectedItem.Index
' (PM) 03/01/06 - <
'    Me.Controls("cmdLookup").Visible = False
    Select Case mintCurFrame
        Case 1
            MaskEdBox2.SetFocus
        Case 2
            'mskField3.SetFocus
    End Select
End Sub



