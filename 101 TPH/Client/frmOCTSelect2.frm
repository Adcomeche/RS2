VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTSelect2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EVASIONE ORDINI"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11445
   Begin VB.Frame Frame1 
      Height          =   5175
      Index           =   2
      Left            =   240
      TabIndex        =   67
      Top             =   480
      Visible         =   0   'False
      Width           =   11055
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   8
         Left            =   10560
         Picture         =   "frmOCTSelect2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   4680
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   6360
         Picture         =   "frmOCTSelect2.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   4755
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F8 - Crea Viaggi"
         Height          =   375
         Left            =   6240
         Picture         =   "frmOCTSelect2.frx":0204
         TabIndex        =   79
         Tag             =   "REFRESHLIST"
         Top             =   4680
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
         Left            =   4320
         Picture         =   "frmOCTSelect2.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   4755
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   4200
         Picture         =   "frmOCTSelect2.frx":0408
         TabIndex        =   77
         Tag             =   "SAVE"
         Top             =   4680
         Width           =   2055
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
         Picture         =   "frmOCTSelect2.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4755
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
         Left            =   2280
         Picture         =   "frmOCTSelect2.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   4755
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Ordine"
         Height          =   375
         Left            =   2160
         Picture         =   "frmOCTSelect2.frx":070E
         TabIndex        =   73
         Tag             =   "OPEN"
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTSelect2.frx":0810
         TabIndex        =   75
         Tag             =   "PRINT"
         Top             =   4680
         Width           =   2055
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "frmOCTSelect2.frx":0912
         TabIndex        =   51
         Top             =   240
         Width           =   10815
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   255
         Left            =   9240
         TabIndex        =   81
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scomparti R"
         Height          =   255
         Index           =   49
         Left            =   5640
         TabIndex        =   86
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblScompartiR 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6720
         TabIndex        =   85
         Top             =   4320
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scomparti M"
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   84
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblScompartiM 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   83
         Top             =   4320
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Data:"
         Height          =   255
         Index           =   19
         Left            =   8640
         TabIndex        =   82
         Top             =   4680
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   5175
      Index           =   1
      Left            =   240
      TabIndex        =   55
      Top             =   480
      Width           =   11055
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo ordini non annullati"
         Height          =   255
         Left            =   6000
         TabIndex        =   71
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   7
         Left            =   7680
         Picture         =   "frmOCTSelect2.frx":16114
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   6
         Left            =   7680
         Picture         =   "frmOCTSelect2.frx":16216
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   5
         Left            =   7680
         Picture         =   "frmOCTSelect2.frx":16318
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   4
         Left            =   7680
         Picture         =   "frmOCTSelect2.frx":1641A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   3240
         Picture         =   "frmOCTSelect2.frx":1651C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3240
         Picture         =   "frmOCTSelect2.frx":1661E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frmOCTSelect2.frx":16720
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmOCTSelect2.frx":16822
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2760
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "21"
         Top             =   4560
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Tag             =   "23"
         Top             =   4560
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Tag             =   "0"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "1"
         Top             =   240
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
         Left            =   1920
         TabIndex        =   3
         Tag             =   "2"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField15 
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Tag             =   "17"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField6 
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Tag             =   "12"
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Tag             =   "13"
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField12 
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         Tag             =   "18"
         Top             =   1320
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
         Left            =   1920
         TabIndex        =   31
         Tag             =   "3"
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Tag             =   "5"
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Tag             =   "7"
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Tag             =   "25"
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Tag             =   "20"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   1920
         TabIndex        =   43
         Tag             =   "27"
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   6360
         TabIndex        =   45
         Tag             =   "28"
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Tag             =   "29"
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Tag             =   "26"
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   6360
         TabIndex        =   34
         Tag             =   "30"
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   6360
         TabIndex        =   40
         Tag             =   "31"
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   7680
         TabIndex        =   68
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
         Left            =   10320
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
               Picture         =   "frmOCTSelect2.frx":16924
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTSelect2.frx":16A36
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTSelect2.frx":16B48
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField11 
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Tag             =   "32"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label lblField8 
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
         TabIndex        =   87
         Tag             =   "32"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ra&ggr."
         Height          =   255
         Index           =   25
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   5280
         Tag             =   "32"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   70
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ordini"
         Height          =   255
         Left            =   8280
         TabIndex        =   69
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   8160
         Tag             =   "31"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Inseri&mento"
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   39
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   8160
         Tag             =   "30"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A&lla Data Ricezione"
         Height          =   255
         Index           =   14
         Left            =   4680
         TabIndex        =   33
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   8160
         Tag             =   "26"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Cons. &Pian."
         Height          =   255
         Index           =   13
         Left            =   4560
         TabIndex        =   27
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   8160
         Tag             =   "29"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Cons. Ric&h."
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   21
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   4080
         Tag             =   "27"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Riferim. Esterno &1"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   42
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   8520
         Tag             =   "28"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Riferim. Esterno &2"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   44
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Ordine"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7680
         Tag             =   "20"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla D&ata Cons. Rich."
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3720
         Tag             =   "7"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data C&ons. Pian."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3720
         Tag             =   "25"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Inserimento"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3720
         Tag             =   "5"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Ricezione"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   3720
         Tag             =   "3"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ti&po Ordine"
         Height          =   255
         Index           =   26
         Left            =   960
         TabIndex        =   46
         Top             =   4560
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   4080
         Tag             =   "21"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A&vanzamento"
         Height          =   255
         Index           =   28
         Left            =   5160
         TabIndex        =   48
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   8520
         Tag             =   "23"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   17
         Left            =   10440
         Tag             =   "18"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Zona"
         Height          =   255
         Index           =   17
         Left            =   6480
         TabIndex        =   8
         Top             =   1320
         Width           =   495
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
         Left            =   7680
         TabIndex        =   66
         Tag             =   "18"
         Top             =   1320
         Width           =   2655
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
         Left            =   3840
         TabIndex        =   65
         Tag             =   "12"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   1680
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   7680
         Tag             =   "12"
         Top             =   1680
         Width           =   240
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
         Left            =   3840
         TabIndex        =   64
         Tag             =   "13"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Des&tinazione"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   7680
         Tag             =   "13"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   5280
         Tag             =   "17"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Deposito"
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   735
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
         Left            =   2520
         TabIndex        =   63
         Tag             =   "17"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5280
         Tag             =   "1"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   0
         Top             =   240
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
         Left            =   2520
         TabIndex        =   62
         Tag             =   "1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   5280
         Tag             =   "2"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emitt."
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   855
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
         Left            =   2520
         TabIndex        =   61
         Tag             =   "2"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero Ordine"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3240
         Tag             =   "0"
         Top             =   2400
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   57
      Top             =   5880
      Width           =   11415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmOCTSelect2.frx":16C5A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmOCTSelect2.frx":16D5C
         TabIndex        =   53
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
         Index           =   7
         Left            =   240
         Picture         =   "frmOCTSelect2.frx":16E5E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTSelect2.frx":16F60
         TabIndex        =   52
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
         Left            =   9720
         Picture         =   "frmOCTSelect2.frx":17062
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9600
         Picture         =   "frmOCTSelect2.frx":17164
         TabIndex        =   54
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5655
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Condizioni"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Risultato Ricerca"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   56
      Top             =   6795
      Width           =   11445
      _ExtentX        =   20188
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
            Object.Width           =   14737
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18:00"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "03/01/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTSelect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 23 '11
Public XA As XArrayDB

Public WithEvents DataClass As COrdineSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes
Private mblnIsDirty As Boolean
Private AXX As CParametro
Private APR As CProdotto
Private AUT As CUnitaDiTrasporto
Private AAM As CAutomezzo
Private AAU As CAutista
Private OCT As COrdine
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto


Private Sub CreateTrips(DataViaggio As Variant)
    On Error GoTo ErrHandler
    Dim i As Long
    Dim strIDAUTPrec As String
    Dim strIDAUT As String
    Dim strIDVVT As String
    Dim strIDVVC As String
    Dim strIDOCT As String
    Dim varSeqPrec As Variant
    Dim varSeq As Variant
    
    XA.QuickSort 0, XA.UpperBound(1), _
                 0, XORDER_ASCEND, XTYPE_DEFAULT, _
                 1, XORDER_ASCEND, XTYPE_LONG

    For i = 0 To XA.UpperBound(1)
        If (Not IsNull(XA.value(i, 0))) And (Len(XA.value(i, 0)) > 0) Then
            strIDAUT = XA.value(i, 0)
            varSeq = XA.value(i, 1)
            strIDOCT = XA.value(i, 15)
            If (strIDAUT <> strIDAUTPrec) Or (varSeq <> varSeqPrec) Then
                strIDAUTPrec = strIDAUT
                varSeqPrec = varSeq
                AUT.DS.Fields.ResetFields
                AUT.DS.Fields("0").FindValue = strIDAUT
                AUT.Find
                VVT.DS.Fields.ResetFields
                With VVT.DS
                    
                    .Fields("2").NewValue = varSeq
                    .Fields("2").Updated = True
                    .Fields("3").NewValue = XA.value(i, 23)
                    .Fields("3").Updated = True
                    .Fields("4").NewValue = DataViaggio
                    .Fields("4").Updated = True
                    .Fields("6").NewValue = DataViaggio
                    .Fields("6").Updated = True
                    .Fields("13").NewValue = "TD"
                    .Fields("13").Updated = True
                    .Fields("14").NewValue = strIDAUT
                    .Fields("14").Updated = True
                    .Fields("15").NewValue = AUT.DS.Fields("2").value
                    .Fields("15").Updated = True
                    .Fields("16").NewValue = AUT.DS.Fields("3").value
                    .Fields("16").Updated = True
                    .Fields("17").NewValue = AUT.DS.Fields("6").value
                    .Fields("17").Updated = True
                    .Fields("18").NewValue = AUT.DS.Fields("7").value
                    .Fields("18").Updated = True
                    .Fields("19").NewValue = XA.value(i, 2)
                    .Fields("19").Updated = True
                    .Fields("21").NewValue = AUT.DS.Fields("11").value
                    .Fields("21").Updated = True
                    .Fields("22").NewValue = XA.value(i, 9)
                    .Fields("22").Updated = True
                    .Fields("23").NewValue = XA.value(i, 9)
                    .Fields("23").Updated = True
                    .Fields("25").NewValue = 0
                    .Fields("25").Updated = True
                    .Fields("26").NewValue = "M"
                    .Fields("26").Updated = True
                    .Fields("27").NewValue = 0
                    .Fields("27").Updated = True
                    .Fields("28").NewValue = "M"
                    .Fields("28").Updated = True
                    .Fields("43").NewValue = "R"
                    .Fields("43").Updated = True
                End With
                If Len(strIDVVT) > 0 Then
                    VVT.DS.DA.EndTrans ETCOMMIT
                End If
            
                '-- DP 05/08/08
                If Not VVT.CheckSequenza(VVT.DS.Fields("6").NewValue, VVT.DS.Fields("2").NewValue, VVT.DS.Fields("14").NewValue, VVT.DS.Fields("22").NewValue, "") Then
                    MsgBox "Il numero di sequenza è già stato utilizzato", vbCritical, "ResOil.2"
                    GoTo ErrHandler
                End If
                
                VVT.DS.DA.BeginTransaction
                If Not VVT.Insert Then
                    GoTo ErrHandler
                End If
                strIDVVC = "010"
            End If
            With VVC.DS
                strIDVVT = VVT.DS.Fields("0").value
                .Fields("0").NewValue = strIDVVT
                .Fields("0").Updated = True
                .Fields("1").NewValue = strIDVVC
                .Fields("1").Updated = True
                .Fields("2").NewValue = strIDVVC
                .Fields("2").Updated = True
            End With
            If Not VVC.OrdineToConsegna(strIDOCT) Then
                GoTo ErrHandler
            End If
            If Not VVC.Insert Then
                GoTo ErrHandler
            End If
            If Not VVP.OrdineToProdotti(strIDOCT, strIDVVT, strIDVVC) Then
                GoTo ErrHandler
            End If
            strIDVVC = Format$(CInt(strIDVVC) + 10, "000")
        End If
        DoEvents
    Next i
    VVT.DS.DA.EndTrans ETCOMMIT
    Exit Sub
ErrHandler:
    VVT.DS.DA.EndTrans ETROLLBACK
End Sub


Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Sub DisplaySelection(List As Variant)
'    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
'                       GridArraySelect, DataClass.DS.HowMany, MAX_COLS
    RefreshXArray XA, TDBGrid1, List
    TDBGrid1.bookmark = 0
End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.bookmark, 15)
    End If
End Function

Private Function SaveTripData() As Boolean
    On Error GoTo ErrHandler
    
    Dim strData1 As String
    Dim strData2 As String
    Dim strData3 As String
    Dim strAUTData As String
    Dim i As Long
    Dim lngUbound As Long
    Dim blnConcErr As Boolean
    lngUbound = XA.UpperBound(1)
    For i = 0 To lngUbound
        If (Not IsNull(XA.value(i, 0))) And (Not IsNull(XA.value(i, 1))) And (Not IsNull(XA.value(i, 2))) Then
            OCT.DS.Fields("0").FindValue = XA.value(i, 15)
            OCT.Find
            strData1 = Space(3)
            strData2 = Space(3)
            strData3 = Space(3)
            LSet strData1 = XA.value(i, 0)
            LSet strData2 = XA.value(i, 1)
            LSet strData3 = XA.value(i, 2)
            strAUTData = strData1 & strData2 & strData3
            OCT.DS.Fields("11").NewValue = strAUTData
            OCT.DS.Fields("11").Updated = True
            OCT.BeginSave blnConcErr
            OCT.DS.DA.EndTrans ETCOMMIT
        End If
    Next i
    SaveTripData = True
    Exit Function
ErrHandler:
    OCT.DS.DA.EndTrans ETROLLBACK
    SaveTripData = False
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    With DataClass.DS
        If .Fields.Validate Then
            If .ScanClientCond Then
                ReDim varOrderFields(0) As Variant
                Me.MousePointer = vbHourglass
                varOrderFields(0) = "0"
                blnActiveOnly = (chkActiveOnly = vbChecked)
                varBuffer = DataClass.GetSelectionEvasione '(blnActiveOnly, varOrderFields, CLng(mskMaxRecords)) Then
                If IsEmpty(varBuffer) Then
                    GoTo ErrHandler
                End If
                DisplaySelection varBuffer
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    SelectRecords = True
    GoTo EndProc
ErrHandler:
    SelectRecords = False
EndProc:
    Me.MousePointer = vbDefault
End Function


Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strSelectionFormula As String
    
    #If VersionID = 1 Then
        ReDim varReportsList(0 To 2, 0 To 3) As Variant
    #Else
        ReDim varReportsList(0 To 2, 0 To 1) As Variant
    #End If
    
    Select Case Cmd
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
                varReportsList(0, 0) = "Stampa elenco Ordini"
                varReportsList(1, 0) = "oil2oct1.rpt"
                varReportsList(2, 0) = 1
                varReportsList(0, 1) = "Stampa elenco Ordini sintetico"
                varReportsList(1, 1) = "oil2oct2.rpt"
                varReportsList(2, 1) = 1
                #If VersionID = 1 Then
                    varReportsList(0, 2) = "Stampa elenco Ordini con Annotazioni"
                    varReportsList(1, 2) = "oil2oct3.rpt"
                    varReportsList(2, 2) = 1
                    varReportsList(0, 3) = "Stampa elenco Ordini con date discordanti"
                    varReportsList(1, 3) = "oil2oct4.rpt"
                    varReportsList(2, 3) = 1
                #End If
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcORDIOCT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcSave
            TDBGrid1.Update
            Me.MousePointer = vbHourglass
            If SaveTripData Then
                mblnIsDirty = False
            End If
            Me.MousePointer = vbDefault
        Case fcLinkRefresh
            If IsDate(MaskEdBox11) Then
                CreateTrips MaskEdBox11
                DoCommand fcFind
            End If
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
                mskField9.SetFocus
            End If
            Combo3 = "ADQUIRIDO"
            Combo3.Locked = True
        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Ordini trovati: " & DataClass.DS.HowMany
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
            strTitle = "Dalla data cons. rich."
        Case 1
            Set M = MaskEdBox3
            strTitle = "Dalla data cons. pian."
        Case 2
            Set M = MaskEdBox1
            strTitle = "Dalla data ricezione"
        Case 3
            Set M = mskField4
            strTitle = "Dalla data inserimento"
        Case 4
            Set M = MaskEdBox7
            strTitle = "Alla data cons. rich."
        Case 5
            Set M = MaskEdBox8
            strTitle = "Alla data cons. pian."
        Case 6
            Set M = MaskEdBox9
            strTitle = "Alla data ricezione"
        Case 7
            Set M = MaskEdBox10
            strTitle = "Alla data inserimento"
        Case 8
            Set M = MaskEdBox11
            strTitle = "Data viaggi"
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
    Dim strIDParam As String
    ReDim ColTags(1 To 2) As Variant
    Dim Col As Column
    
    Loading = True
    mintCurFrame = 1
'    Set GridArraySelect = New ArrayGridCls
    
    ColTags(1) = "0"
    ColTags(2) = "1"
    
    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New COrdineSelezione
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
    Set VVC = New CViaggioConsegna
    With VVC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = strIDParam
        AXX.Find
        For i = 1 To 6
            APR.DS.Fields("0").FindValue = .Fields(CStr(5 + i)).value
            APR.Find
            TDBGrid1.Columns(15 + i).Caption = APR.DS.Fields("3").value
        Next i
    End With
    ' ---
    For Each Col In TDBGrid1.Columns
        Col.ButtonHeader = True
    Next Col
    
    Set AUT = New CUnitaDiTrasporto
    Set AUT.DS.DA = DataConnect.DA
'    SetValueItems TDBGrid1.Columns(0), AUT, False
'    Dim vItem As ValueItem
'    Set vItem = New ValueItem
'    vItem.value = ""
'    vItem.DisplayValue = ""
'    TDBGrid1.Columns(0).ValueItems.Add vItem
    
    Set AAU = New CAutista
    Set AAU.DS.DA = DataConnect.DA
'    SetValueItems TDBGrid1.Columns(2), AAU, False
    MaskEdBox11 = Format(DateAdd("d", 1, Date), "dd/mm/yy")
    mskField11 = "GC"
    
    SetState fsparaminput
    
    Select Case mvarFunction
        Case fcORDSELE
        Case Else
            DoCommand fcFind
    End Select
    
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnIsDirty Then
        Cancel = (MsgBox("Le modifiche apportate non verranno salvate. Vuoi proseguire ?", vbOKCancel + vbExclamation, "Attenzione") = vbCancel)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set XA = Nothing
    Set Soc = Nothing
    Set frmOCTSelect2 = Nothing
End Sub





Private Sub DisplayScomparti(IDAutomezzo As String, L As Label)
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim strScomparti As String
    Dim dblPortata As Double
    Dim dblPortataM As Double
    Dim dblPortataR As Double
    
    L = vbNullString
    If Len(IDAutomezzo) = 0 Then
        dblPortata = 0
    Else
        With AAM.DS
            .Fields("0").FindValue = IDAutomezzo
            If Not AAM.Find Then
                GoTo ErrHandler
            End If
            If AAM.DS.HowMany = 0 Then
                GoTo ErrHandler
            End If
            For i = 1 To 10
                If .Fields(CStr(i * 2 + 4)).value <> 0 Then
                    strScomparti = strScomparti & .Fields(CStr(i * 2 + 4)).value & " x " & .Fields(CStr((i * 2 - 1) + 4)).value & " "
                    dblPortata = dblPortata + .Fields(CStr(i * 2 + 4)).value * .Fields(CStr((i * 2 - 1) + 4)).value
                End If
            Next i
        End With
    End If
    L = " " & strScomparti
    
    Exit Sub

ErrHandler:
    L = vbNullString
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    Dim i As Long
    Dim strIDAUT As String
    Dim varMaxSeq As Variant
    
    mblnIsDirty = True
    Select Case ColIndex
        Case 0
            strIDAUT = TDBGrid1.Columns(ColIndex).value
            AUT.DS.Fields.ResetFields
            AUT.DS.Fields("0").FindValue = strIDAUT
            AUT.Find
            If AUT.DS.HowMany > 0 Then
'                For i = 0 To XA.UpperBound(1)
'                    If XA.value(i, 0) = strIDAUT Then
'                        If XA.value(i, 1) > varMaxSeq Then
'                            varMaxSeq = XA.value(i, 1)
'                        End If
'                    End If
'                Next i
'                If IsEmpty(varMaxSeq) Then
'                    varMaxSeq = 0
'                End If
'                TDBGrid1.Columns(1).value = varMaxSeq + 1
                TDBGrid1.Columns(2).value = AUT.DS.Fields("4").value
                DisplayScomparti AUT.DS.Fields("2").value, lblScompartiM
                DisplayScomparti AUT.DS.Fields("3").value, lblScompartiR
                AAU.DS.Fields("0").FindValue = AUT.DS.Fields("4").value
                AAU.Find
                TDBGrid1.Columns(3).value = AAU.DS.Fields("1").value
            End If
            If Len(strIDAUT) = 0 Then
                TDBGrid1.Columns(1).value = vbNullString
                TDBGrid1.Columns(3).value = vbNullString
            End If
        Case 2
            AAU.DS.Fields.ResetFields
            AAU.DS.Fields("0").FindValue = TDBGrid1.Columns(ColIndex).value
            AAU.Find
            If AAU.DS.HowMany > 0 Then
                TDBGrid1.Columns(3).value = AAU.DS.Fields("1").value
            Else
                TDBGrid1.Columns(3).value = vbNullString
            End If
    End Select
End Sub

Private Sub TDBGrid1_Change()
    Select Case TDBGrid1.Col
        Case 0, 2
            If TDBGrid1.Columns(TDBGrid1.Col).value = "?" Then
                TDBGrid1.Columns(TDBGrid1.Col).value = ""
                Dim i As Integer
                For i = 0 To Controls.Count - 1
                    If Controls(i).Name = "TDBGrid1" Then
                        frmSearch.LinkControlIndex = i
                        Exit For
                    End If
                Next i
                If TDBGrid1.Col = 0 Then
                    ShowSearchForm Me, "", AUT, TDBGrid1.TabIndex, True
                Else
                    ShowSearchForm Me, "", AAU, TDBGrid1.TabIndex, True
                End If
                Do While IsLoaded("frmSearch")
                    DoEvents
                Loop
                SendKeys "{TAB}"
            End If
    End Select
        
End Sub


Private Sub TDBGrid1_DblClick()
    DoCommand fcOpen
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If IsNull(XA(bookmark, 13)) Or Len(XA(bookmark, 13)) = 0 Then
        RowStyle.BackColor = vbRed
        RowStyle.ForeColor = vbWhite
    End If
End Sub


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim x As XTYPE
    Dim o As XORDER
    Static blnLastSortASC  As Boolean
    Static intLastCol As Integer
    
    If ColIndex <> intLastCol Then
        blnLastSortASC = False
    Else
        blnLastSortASC = Not blnLastSortASC
    End If
    If blnLastSortASC Then
        o = XORDER_DESCEND
    Else
        o = XORDER_ASCEND
    End If
        
    Select Case ColIndex
        Case 1
            x = XTYPE_LONG
        Case 11
            x = XTYPE_DATE
        Case 12
            x = XTYPE_LONG
        Case Is > 15
            x = XTYPE_LONG
        Case Else
            x = XTYPE_STRING
    End Select
    XA.QuickSort 0, XA.UpperBound(1), _
                 ColIndex, o, x
    TDBGrid1.ReBind
    intLastCol = ColIndex
End Sub

Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DoEvents
        DoCommand fcOpen
    End If
End Sub


Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strIDAUT As String
    
    lblScompartiM = vbNullString
    lblScompartiR = vbNullString
    If Not IsNull(TDBGrid1.Columns(0).value) Then
        strIDAUT = TDBGrid1.Columns(0).value
        If (Len(strIDAUT) > 0) Then
            AUT.DS.Fields("0").FindValue = strIDAUT
            AUT.Find
            If AUT.DS.HowMany > 0 Then
                DisplayScomparti AUT.DS.Fields("2").value, lblScompartiM
                DisplayScomparti AUT.DS.Fields("3").value, lblScompartiR
            End If
        End If
    End If
End Sub


