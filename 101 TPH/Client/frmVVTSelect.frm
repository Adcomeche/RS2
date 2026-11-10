VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONE VIAGGI"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   9510
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   60
      TabIndex        =   54
      Top             =   6360
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmVVTSelect.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla condizioni"
         Height          =   375
         Left            =   1800
         Picture         =   "frmVVTSelect.frx":0102
         TabIndex        =   41
         Tag             =   "UNDO"
         Top             =   360
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmVVTSelect.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSelect.frx":0306
         TabIndex        =   40
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
         Index           =   2
         Left            =   7680
         Picture         =   "frmVVTSelect.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmVVTSelect.frx":050A
         TabIndex        =   42
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   53
      Top             =   7215
      Width           =   9510
      _ExtentX        =   16775
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
            Object.Width           =   11324
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:57"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11/02/2014"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   5715
      Index           =   1
      Left            =   240
      TabIndex        =   52
      Top             =   480
      Width           =   9015
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   7680
         Picture         =   "frmVVTSelect.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3480
         Picture         =   "frmVVTSelect.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo viaggi non annullati"
         Height          =   255
         Left            =   6000
         TabIndex        =   74
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3480
         Picture         =   "frmVVTSelect.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7680
         Picture         =   "frmVVTSelect.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4560
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "9"
         Top             =   5280
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "10"
         Top             =   5280
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Tag             =   "0"
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Tag             =   "3"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Tag             =   "4"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Tag             =   "5"
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Tag             =   "1"
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Tag             =   "6"
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Tag             =   "11"
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   6240
         TabIndex        =   25
         Tag             =   "12"
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Tag             =   "13"
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Tag             =   "14"
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Tag             =   "15"
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   7680
         TabIndex        =   71
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8100
         Top             =   2640
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
               Picture         =   "frmVVTSelect.frx":0A14
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSelect.frx":0B26
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSelect.frx":0C38
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Tag             =   "92"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Tag             =   "89"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField15 
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Tag             =   "7"
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
         Left            =   840
         TabIndex        =   7
         Tag             =   "17"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox12 
         Height          =   255
         Left            =   2040
         TabIndex        =   78
         Tag             =   "34"
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   2040
         TabIndex        =   81
         Tag             =   "36"
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Tag             =   "40"
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   255
         Left            =   6240
         TabIndex        =   34
         Tag             =   "41"
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   2040
         TabIndex        =   82
         Tag             =   "42"
         Top             =   3480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
         Left            =   1440
         TabIndex        =   75
         Tag             =   "92"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblField3 
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
         Left            =   2640
         TabIndex        =   63
         Tag             =   "13"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label6 
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
         Left            =   2640
         TabIndex        =   89
         Tag             =   "42"
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vettore Motrice"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   88
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   19
         Left            =   7320
         Tag             =   "42"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   17
         Left            =   3960
         Tag             =   "40"
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data Riferimento"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   87
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Riferimento"
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   86
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   8160
         Tag             =   "41"
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   7320
         Tag             =   "36"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Autista 2"
         Height          =   255
         Index           =   15
         Left            =   1200
         TabIndex        =   84
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label5 
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
         Left            =   3480
         TabIndex        =   83
         Tag             =   "36"
         Top             =   3120
         Width           =   3735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   7320
         Tag             =   "34"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Autista"
         Height          =   255
         Index           =   14
         Left            =   1200
         TabIndex        =   80
         Top             =   2760
         Width           =   735
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
         Left            =   3480
         TabIndex        =   79
         Tag             =   "34"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblField6 
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
         Left            =   1440
         TabIndex        =   77
         Tag             =   "7"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Dep."
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   4200
         Tag             =   "7"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   4200
         Tag             =   "17"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Zona"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   495
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
         Left            =   1440
         TabIndex        =   76
         Tag             =   "17"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   2280
         Tag             =   "90"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Presidio"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   4200
         Tag             =   "92"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Raggr."
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "viaggi"
         Height          =   255
         Left            =   8280
         TabIndex        =   73
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   72
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblField5 
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
         Left            =   3480
         TabIndex        =   65
         Tag             =   "15"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Rimorchio"
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   16
         Top             =   2400
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   7320
         Tag             =   "15"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   7320
         Tag             =   "14"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Motrice"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5400
         Tag             =   "13"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unità di Trasporto"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3480
         Tag             =   "11"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal Num. &Viaggio Def."
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   22
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   7680
         Tag             =   "12"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio Def."
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   24
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   8160
         Tag             =   "6"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Ini&zio Viaggio"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   29
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Viaggio"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   20
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7680
         Tag             =   "1"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Inizio Viaggio"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   26
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3960
         Tag             =   "5"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Tipo Viaggio"
         Height          =   255
         Index           =   26
         Left            =   720
         TabIndex        =   36
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   4200
         Tag             =   "9"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Stato Viaggio"
         Height          =   255
         Index           =   28
         Left            =   5040
         TabIndex        =   38
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   8400
         Tag             =   "10"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   8640
         Tag             =   "3"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   8
         Top             =   960
         Width           =   615
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
         Left            =   5880
         TabIndex        =   59
         Tag             =   "3"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   8640
         Tag             =   "4"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Em."
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   10
         Top             =   1320
         Width           =   735
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
         Left            =   5880
         TabIndex        =   58
         Tag             =   "4"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero Viaggio"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3480
         Tag             =   "0"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label lblField4 
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
         Left            =   3480
         TabIndex        =   64
         Tag             =   "14"
         Top             =   2040
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5715
      Index           =   2
      Left            =   240
      TabIndex        =   60
      Top             =   480
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton Command11 
         Caption         =   "Palms"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         Picture         =   "frmVVTSelect.frx":0D4A
         TabIndex        =   90
         Tag             =   "EXPPALM"
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Vehiculos"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         Picture         =   "frmVVTSelect.frx":0E4C
         TabIndex        =   85
         Tag             =   "EXPORT"
         Top             =   5160
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   4920
         Picture         =   "frmVVTSelect.frx":0F4E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   5235
         Width           =   240
      End
      Begin VB.CommandButton Command10 
         Caption         =   "F12 - Copia"
         Height          =   375
         Left            =   4800
         Picture         =   "frmVVTSelect.frx":1050
         TabIndex        =   51
         Tag             =   "CUSTOM2"
         Top             =   5160
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   2520
         Picture         =   "frmVVTSelect.frx":1152
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   5235
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F11 - Rendi provvisori"
         Height          =   375
         Left            =   2400
         Picture         =   "frmVVTSelect.frx":1254
         TabIndex        =   50
         Tag             =   "CUSTOM1"
         Top             =   5160
         Width           =   2415
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   240
         Picture         =   "frmVVTSelect.frx":1356
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4875
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
         Left            =   6960
         Picture         =   "frmVVTSelect.frx":1458
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4875
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Annulla"
         Height          =   375
         Left            =   6840
         Picture         =   "frmVVTSelect.frx":155A
         TabIndex        =   49
         Tag             =   "EDITLINK"
         Top             =   4800
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmVVTSelect.frx":165C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   4875
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri"
         Height          =   375
         Left            =   4800
         Picture         =   "frmVVTSelect.frx":175E
         TabIndex        =   47
         Tag             =   "OPEN"
         Top             =   4800
         Width           =   2055
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
         Picture         =   "frmVVTSelect.frx":1860
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   5235
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F8 - Rendi Definitivi"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSelect.frx":1962
         TabIndex        =   48
         Tag             =   "REFRESHLIST"
         Top             =   5160
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   2520
         Picture         =   "frmVVTSelect.frx":1A64
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   4875
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   2400
         Picture         =   "frmVVTSelect.frx":1B66
         TabIndex        =   46
         Tag             =   "PRINT"
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F2 - Seleziona Tutti"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSelect.frx":1C68
         TabIndex        =   45
         Tag             =   "SAVE"
         Top             =   4800
         Width           =   2295
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   4455
         Left            =   0
         OleObjectBlob   =   "frmVVTSelect.frx":1D6A
         TabIndex        =   44
         Tag             =   "SELECT"
         Top             =   360
         Width           =   8775
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10821
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Condi&zioni"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Risu&ltato Ricerca"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVVTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 10
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As CViaggioSelezione
Attribute DataClass.VB_VarHelpID = -1
Private VVT As CViaggio
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes
Private mstrPrevValueARO As String

Public Sub CopyTrips()
    On Error GoTo ErrHandler
    Dim Rows As Long
    Dim i As Long
    Dim strIDViaggio As String
    Dim intError As Errors
    
    Me.MousePointer = vbHourglass
    intError = NoError
    Rows = TDBGrid1.SelBookmarks.Count - 1
    If Rows >= 0 Then
        For i = 0 To Rows
            strIDViaggio = GetSelectedID(TDBGrid1, TDBGrid1.SelBookmarks(i))
            With VVT
                .DS.Fields("0").FindValue = strIDViaggio
                If .Find Then
                    ' --- Escludo i viaggi annullati
                    If .DS.Fields("44").value <> True Then
                        If Not .Copia Then
                            intError = TripCopyError
                            GoTo ErrHandler
                        End If
                    End If
                End If
            End With
        Next i
    End If
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    Dim C As CError
    
    Select Case intError
        Case TripDefinitionError
            strMessage = "Errore durante la copia del viaggio"
            strInfo = "Viaggio numero: " & strIDViaggio
        Case NoError
            Select Case Err.Number
                Case Else
                    intError = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(intError, "frmVVTSelect", _
                "CopyTrips", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    DoCommand fcFind
    Me.MousePointer = vbDefault
End Sub

Public Sub DeleteTrips()
    On Error GoTo ErrHandler
    
    Dim Rows As Long
    Dim i As Long
    Dim strIDViaggio As String
    Dim intError As Errors
    Dim blnConcurError As Boolean
    
    Rows = TDBGrid1.SelBookmarks.Count - 1
    
    #If VersionID = 1 Then
        If Rows >= 0 Then
            If MsgBox("Verranno eliminati definitivamente " & Rows + 1 & " viaggi. " & _
                      "Confermi l'eliminazione?", vbYesNo + vbDefaultButton2 + vbExclamation, _
                      "Attenzione") <> vbYes Then
                GoTo EndProc
            End If
        End If
        DoEvents
    #End If
    
    Me.MousePointer = vbHourglass
    If Rows >= 0 Then
        For i = 0 To Rows
            intError = NoError
            strIDViaggio = GetSelectedID(TDBGrid1, TDBGrid1.SelBookmarks(i))
            With VVT.DS
                intError = TripDeleteError
                .Fields("0").FindValue = strIDViaggio
                If VVT.Find Then
                    #If VersionID = 1 Then
                        VVT.Delete
                    #Else
                        .Fields("44").NewValue = True
                        .Fields("44").Updated = True
                        VVT.BeginSave blnConcurError
                        If blnConcurError Then
                            If MsgBox(MSG_CONCUR_ERR, _
                                       vbOKCancel + vbExclamation + vbDefaultButton2, _
                                      "Attenzione") = vbOK Then
                                GoTo EndProc
                            Else
                                .DA.EndTrans ETROLLBACK
                            End If
                        Else
                            .DA.EndTrans ETCOMMIT
                            .EndSave
                        End If
                    #End If
                End If
            End With
        Next i
    End If
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    Dim C As CError
    
    Select Case intError
        Case TripDeleteError
            strMessage = "Errore durante l'annullamento del viaggio"
            strInfo = "Viaggio numero: " & strIDViaggio
        Case NoError
            Select Case Err.Number
                Case Else
                    intError = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(intError, "frmVVTSelect", _
                "DeleteTrips", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    DoCommand fcFind
    Me.MousePointer = vbDefault

End Sub

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Sub SetButtonState()
    Dim lngCount As Long
    
    lngCount = TDBGrid1.SelBookmarks.Count
    Command1.Enabled = (lngCount > 0)
    Command9.Enabled = (lngCount > 0)
    #If VersionID <> 1 Then
        Command10.Enabled = (lngCount > 0)
    #End If
    Command6.Enabled = (lngCount > 0)
    Command5.Enabled = (lngCount = 0)
    Command2.Enabled = (lngCount = 0)
    cmdExport.Enabled = (lngCount > 0)
    Command11.Enabled = (lngCount > 0)
End Sub

Public Sub SetTripsState(NewState As String)
    On Error GoTo ErrHandler
    Dim Rows As Long
    Dim i As Long
    Dim strIDViaggio As String
    Dim strMessage As String
    Dim intError As Errors
    
    Me.MousePointer = vbHourglass
    intError = NoError
    Rows = TDBGrid1.SelBookmarks.Count - 1
    If Rows >= 0 Then
        For i = 0 To Rows
            strIDViaggio = GetSelectedID(TDBGrid1, TDBGrid1.SelBookmarks(i))
            With VVT
                .DS.Fields("0").FindValue = strIDViaggio
                If .Find Then
                    ' --- Escludo i viaggi annullati
                    If .DS.Fields("44").value <> True Then
                        If Not .SetStato(NewState, strMessage) Then
                            intError = TripDefinitionError
                            GoTo ErrHandler
                        End If
                    End If
                    ' ---
                End If
            End With
        Next i
    End If
    GoTo EndProc
ErrHandler:
    Dim strInfo As String
    Dim lngNativeCode As Long
    Dim C As CError
    
    Select Case intError
        Case TripDefinitionError
            If Len(strMessage) = 0 Then
                strMessage = "Errore durante la modifica dello stato del viaggio"
            End If
            strInfo = "Viaggio numero: " & strIDViaggio
        Case NoError
            Select Case Err.Number
                Case Else
                    intError = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(intError, "frmVVTSelect", _
                "SetTripsState", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    DoCommand fcFind
    Me.MousePointer = vbDefault
End Sub


Private Sub DisplaySelection()
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, DataClass.DS.HowMany, MAX_COLS

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid, Optional GridBookmark As String = vbNullString) As Variant
    If GridBookmark = vbNullString Then
        If IsNull(DBGrid.bookmark) Then
            GetSelectedID = vbNullString
        Else
            GetSelectedID = GridArraySelect.value(DBGrid.bookmark, 0)
        End If
    Else
        GetSelectedID = GridArraySelect.value(GridBookmark, 0)
    End If
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    ReDim varOrderFields(0) As Variant
    Dim blnActiveOnly As Boolean
    
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            varOrderFields(0) = "0"
            blnActiveOnly = (chkActiveOnly = vbChecked)

            If Not DataClass.GetSelection(blnActiveOnly, varOrderFields, CLng(mskMaxRecords), (UCase(MaskEdBox10) = "GC")) Then
                GoTo ErrHandler
            End If
            DisplaySelection
            GoTo EndProc
        End If
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
ErrHandler:
    SelectRecords = False
    Exit Function
EndProc:
    SelectRecords = True
End Function

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim LinkForm As Form
    Dim i As Long
    Dim Rows As Long
    Dim strSelectionFormula As String
    Dim strError As String
    Dim varIDViaggio As Variant
    Dim Respok, Respuesta As String

    
    Dim Service As VLocalize.VLocalize
    ReDim varReportsList(0 To 4, 0 To 2) As Variant
    
    Select Case Cmd
        Case fcExport
            Set Service = New VLocalize.VLocalize

            ' -- DP 10/02/2009
            'Service.serviceLocation = "http://www.tphweb.es/mbtph/mbws/mbws.asmx"
            Service.serviceLocation = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
            Service.serviceTimeOut = 60

            Service.serviceUsername = "wstph"
            Service.servicePassword = "swtph"

            DataClass.DS.DisplayMessage "Trasmissione in corso... "
            Rows = TDBGrid1.SelBookmarks.Count - 1
            If Rows >= 0 Then
                For i = 0 To Rows
                    varIDViaggio = GetSelectedID(TDBGrid1, TDBGrid1.SelBookmarks(i))
                    If Not DataClass.DEUSTO_InsertViaggio(varIDViaggio, Service, strError) Then
                        MsgBox strError
                    End If
                Next i
            End If

            If Service.isOpen Then
                Service.Close
            End If

            DataClass.DS.DisplayMessage "Trasmissione completata"
            
        Case fcSave
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Esta opción solo esta disponible desde Rs3 Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then

                Rows = GridArraySelect.RowCount - 1
                With TDBGrid1.SelBookmarks
                    For i = 0 To Rows
                        .Add " " & CStr(i)
                    Next i
                End With
                SetButtonState
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
            
        Case fcCustom1
            SetTripsState "P"
            
        Case fccustom2
            #If VersionID = 1 Then
                If mvarFunction = fcVIACOMP Or mvarFunction = fcVIASELC Then
                    Set LinkForm = DisplayForm(fcVIAGCOM)
                    If Not LinkForm Is Nothing Then
                        With LinkForm
                            Set .ParentForm = Me
                            .GetRecord GetSelectedID(TDBGrid1)
                            .SetState fsEdit
                            Me.Enabled = False
                        End With
                    End If
                Else
                    CopyTrips
                End If
            #Else
                CopyTrips
            #End If
            
        Case fcUndo
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
            
        Case fcFind
            If SelectRecords Then
                SetState fsSelection
            End If
            
        Case fcPrint
            If (State = fsSelection) Then
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                varReportsList(0, 0) = "Hojas de ruta"
                varReportsList(1, 0) = "oil2vvt1_gc.rpt"
                varReportsList(2, 0) = 1
                varReportsList(3, 0) = 0
                varReportsList(4, 0) = 0
                varReportsList(0, 1) = "Detalle Programacion"
                varReportsList(1, 1) = "oil2vvt2.rpt"
                varReportsList(2, 1) = 0
                varReportsList(0, 2) = "Riepilogo Consumi e Pedaggi"
                varReportsList(1, 2) = "oil2vvt3.rpt"
                varReportsList(2, 2) = 1
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcexppalm
            Set Service = New VLocalize.VLocalize
    
            
            'Service.serviceLocation = "http://www.tphweb.es/mbtph/mbws/mbws.asmx
            DataClass.DS.DisplayMessage "Transmisión en curso.... "
            Rows = TDBGrid1.SelBookmarks.Count - 1

            Service.serviceLocation = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
            Service.serviceTimeOut = 60
    
            Service.serviceUsername = "wstph"
            Service.servicePassword = "swtph"
            Service.setServerID = DataConnect.ConnServer
            Service.setServerPassword = DataConnect.ConnPWD
            Service.setServerUser = DataConnect.ConnUserName
            
            

            If Rows >= 0 Then
                For i = 0 To Rows
                    varIDViaggio = GetSelectedID(TDBGrid1, TDBGrid1.SelBookmarks(i))
                    If Not DataClass.oesia_InsertaViajepalm(varIDViaggio, Service, strError) Then
                        MsgBox strError
                    End If
                Next i
            End If
            
            If Service.isOpen Then
                Service.Close
            End If
            
            DataClass.DS.DisplayMessage "Transmision completa"
            Command11.Enabled = False
                 
      
    
        Case fcOpenLinkEdit
            DeleteTrips
            
        Case fcLinkRefresh
            SetTripsState "D"
            
        Case fcExit
            Unload Me
               
            
    End Select
End Sub

Private Sub GetRecord(strID As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
        Else
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub







Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            If Not Loading Then
                TabStrip1.Tabs(1).Selected = True
                MaskEdBox10.SetFocus
            End If
            Select Case mvarFunction
                Case fcVIASELE
                Case fcVIASELP
                    Combo3 = "PROVVISORIO"
                    Combo3.Locked = True
                Case fcVIASELD
                    Combo3 = "DEFINITIVO"
                    Combo3.Locked = True
                Case fcVIASELT
                    Combo3 = "DEFINITIVO"
                    Combo3.Locked = True
                    DataClass.DS.Fields("16").FindValue = "T"
                Case fcVIASELC
                    Combo3 = "COMPLETATO"
                    Combo3.Locked = True
                Case fcVIACOMP
                    Combo3 = "DEFINITIVO"
                    Combo3.Locked = True
            End Select
        Case fsSelection
            Command1.Enabled = False
            Command9.Enabled = False
            Command6.Enabled = False
            #If VersionID <> 1 Then
                Command10.Enabled = False
            #End If
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Viaggi trovati: " & DataClass.DS.HowMany
            TabStrip1.Tabs(2).Selected = True
            If Not Loading Then
                TDBGrid1.SetFocus
            End If
    End Select
End Sub


Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Dalla data inizio viaggio"
            
        Case 1
            Set M = MaskEdBox3
            strTitle = "Alla data inizio viaggio"
            
        Case 2
            Set M = MaskEdBox14
            strTitle = "Dalla data riferimento"
            
        Case 3
            Set M = MaskEdBox15
            strTitle = "Alla data riferimento"
            
    End Select
    
    GetCalendarDate M, strTitle
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
    ReDim ColTags(1 To 2) As Variant
    
    Loading = True
    mintCurFrame = 1
    Set GridArraySelect = New ArrayGridCls
    
    ColTags(1) = "0"
    ColTags(2) = "1"
    
    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set Soc = New CSocieta
    Set Soc.DS.DA = DataConnect.DA
    SetValueItems TDBGrid1.Columns(2), Soc
    
    SetState fsparaminput
    
'    If IsUserGrandesCuentas Then
'        MaskEdBox10 = "GC"
'        MaskEdBox11.Tag = "93"
'        imgState(11).Tag = "93"
'    Else
'        MaskEdBox10 = "RP"
'        MaskEdBox11.Tag = "90"
'        imgState(11).Tag = "90"
'    End If
    MaskEdBox11 = DataConnect.ConnUserName
    
    Select Case mvarFunction
        Case fcVIASELE
        Case fcVIASELC
        Case Else
            DoCommand fcFind
    End Select
    
    #If VersionID = 1 Then
        Command6.Caption = "F7 - Elimina"
        If mvarFunction = fcVIACOMP Or mvarFunction = fcVIASELC Then
            Command10.Caption = "F12 - Completa"
'            Command1.Caption = "F8 - Completa"
        End If
    #End If
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set VVT = Nothing
    Set GridArraySelect = Nothing
    Set Soc = Nothing
    Set frmVVTSelect = Nothing
End Sub





Private Sub MaskEdBox10_GotFocus()
     mstrPrevValueARO = MaskEdBox10
End Sub

Private Sub MaskEdBox10_Validate(Cancel As Boolean)
    Select Case UCase(MaskEdBox10)
        Case "GC"
            MaskEdBox11.Tag = "93"
            imgState(11).Tag = "93"
            DataClass.DS.Fields("90").FindValue = vbNullString
        Case "RP"
            MaskEdBox11.Tag = "90"
            imgState(11).Tag = "90"
            DataClass.DS.Fields("93").FindValue = vbNullString
        Case Else
            Exit Sub
    End Select
    If MaskEdBox10 <> mstrPrevValueARO Then
        MaskEdBox11 = vbNullString
    End If
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    #If VersionID <> 1 Then
        If KeyCode = vbKeyReturn Then
            DoCommand fcOpen
        End If
    #End If
End Sub

Private Sub TDBGrid1_SelChange(Cancel As Integer)
    SetButtonState
End Sub


