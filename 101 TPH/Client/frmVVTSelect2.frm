VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTSelect2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTROLLO VIAGGI"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11070
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   5715
      Index           =   1
      Left            =   120
      TabIndex        =   41
      Top             =   480
      Width           =   10815
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar incidencias"
         Height          =   255
         Left            =   840
         TabIndex        =   76
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   8640
         Picture         =   "frmVVTSelect2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3480
         Picture         =   "frmVVTSelect2.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo viaggi non annullati"
         Height          =   255
         Left            =   6000
         TabIndex        =   56
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3480
         Picture         =   "frmVVTSelect2.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   8640
         Picture         =   "frmVVTSelect2.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4560
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "9"
         Top             =   5280
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Tag             =   "10"
         Top             =   5280
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
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
         Left            =   7080
         TabIndex        =   5
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
         Left            =   7080
         TabIndex        =   7
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
         TabIndex        =   23
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
         Left            =   7200
         TabIndex        =   17
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
         Left            =   7200
         TabIndex        =   26
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
         TabIndex        =   19
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
         Left            =   7200
         TabIndex        =   21
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
         TabIndex        =   9
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
         TabIndex        =   11
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
         TabIndex        =   13
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
         TabIndex        =   53
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
               Picture         =   "frmVVTSelect2.frx":0408
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSelect2.frx":051A
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSelect2.frx":062C
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField15 
         Height          =   255
         Left            =   840
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   59
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
         TabIndex        =   62
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
         TabIndex        =   28
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
         Left            =   7200
         TabIndex        =   30
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
         TabIndex        =   63
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Riferimento"
         Height          =   255
         Index           =   17
         Left            =   5400
         TabIndex        =   66
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   9120
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   58
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
         TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   57
         Tag             =   "17"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "viaggi"
         Height          =   255
         Left            =   8280
         TabIndex        =   55
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   54
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
         TabIndex        =   52
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
         TabIndex        =   12
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
         TabIndex        =   10
         Top             =   2040
         Width           =   615
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
         TabIndex        =   51
         Tag             =   "14"
         Top             =   2040
         Width           =   3735
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
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
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
         TabIndex        =   50
         Tag             =   "13"
         Top             =   1680
         Width           =   2655
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
         TabIndex        =   18
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   8640
         Tag             =   "12"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio Def."
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   20
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   9120
         Tag             =   "6"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Ini&zio Viaggio"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   25
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Viaggio"
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   16
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   8640
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
         TabIndex        =   22
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
         TabIndex        =   32
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
         Left            =   6000
         TabIndex        =   34
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   9360
         Tag             =   "10"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   10440
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
         Left            =   6360
         TabIndex        =   4
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
         Left            =   7680
         TabIndex        =   48
         Tag             =   "3"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   10440
         Tag             =   "4"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Em."
         Height          =   255
         Index           =   9
         Left            =   6240
         TabIndex        =   6
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
         Left            =   7680
         TabIndex        =   47
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
         TabIndex        =   14
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
   End
   Begin VB.Frame Frame1 
      Height          =   5715
      Index           =   2
      Left            =   120
      TabIndex        =   49
      Top             =   480
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   7920
         Picture         =   "frmVVTSelect2.frx":073E
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7560
         Picture         =   "frmVVTSelect2.frx":0840
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   8280
         Picture         =   "frmVVTSelect2.frx":091A
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   5280
         Width           =   375
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frmVVTSelect2.frx":0A1C
         TabIndex        =   40
         Tag             =   "SELECT"
         Top             =   240
         Width           =   10575
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Top             =   5280
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   60
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   8880
         TabIndex        =   74
         Top             =   5280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Caption         =   "File Esportazione (.csv)"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   5280
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   43
      Top             =   6360
      Width           =   11055
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmVVTSelect2.frx":E55E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla condizioni"
         Height          =   375
         Left            =   1800
         Picture         =   "frmVVTSelect2.frx":E660
         TabIndex        =   37
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
         Picture         =   "frmVVTSelect2.frx":E762
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSelect2.frx":E864
         TabIndex        =   36
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
         Left            =   9360
         Picture         =   "frmVVTSelect2.frx":E966
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9240
         Picture         =   "frmVVTSelect2.frx":EA68
         TabIndex        =   38
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6135
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   7215
      Width           =   11070
      _ExtentX        =   19526
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
            Object.Width           =   14076
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14.46"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15/10/2008"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVVTSelect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 31
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As CViaggioSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private varResults As Variant

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer

Private Function SelectRecords() As Boolean
    Dim blnActiveOnly As Boolean
    
    On Error GoTo ErrHandler
    
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            ReDim varOrderFields(0) As Variant
            
            varOrderFields(0) = "0"
            blnActiveOnly = (chkActiveOnly = vbChecked)

            If Not DataClass.GetSelection_ControlloViaggi(blnActiveOnly, varOrderFields, CLng(mskMaxRecords)) Then
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

Private Sub DisplaySelection()
    Dim iRow, iCol, i As Integer
    Dim varEventi As Variant

    varResults = Empty
    ClearUnboundGrid TDBGrid1, GridArraySelect, MAX_COLS
    
    ReDim varResults(UBound(DataClass.DS.List, 1), UBound(DataClass.DS.List, 2))
    
    For iRow = 0 To UBound(DataClass.DS.List, 2)
        For iCol = 0 To UBound(DataClass.DS.List, 1)
            Select Case iCol
                Case 1
                    If Check1.value Then
                        '-- Calcolo gli eventi
                        varEventi = DataClass.varGetEventiConsegna(DataClass.DS.List(29, iRow), DataClass.DS.List(30, iRow))
                    End If
                    
                    varResults(iCol, iRow) = DataClass.DS.List(iCol, iRow)

                Case 29
                    '-- Codice Evento
                    If Check1.value Then
                        If Not IsEmpty(varEventi) Then
                            If UBound(varEventi, 2) > 0 Then
                                For i = 0 To UBound(varEventi, 2)
                                    varResults(iCol, iRow) = varResults(iCol, iRow) & varEventi(0, i) & "; "
                                Next i
                                
                                varResults(iCol, iRow) = Left(varResults(iCol, iRow), Len(varResults(iCol, iRow)) - 2)
                            Else
                                varResults(iCol, iRow) = varEventi(0, 0)
                            End If
                        Else
                            varResults(iCol, iRow) = DataClass.DS.List(iCol, iRow)
                        End If
                    End If
                
                Case 30
                    '-- Descrizione del motivo dell'evento
                    If Check1.value Then
                        If Not IsEmpty(varEventi) Then
                            If UBound(varEventi, 2) > 0 Then
                                For i = 0 To UBound(varEventi, 2)
                                    varResults(iCol, iRow) = varResults(iCol, iRow) & varEventi(1, i) & "; "
                                Next i
                                
                                varResults(iCol, iRow) = Left(varResults(iCol, iRow), Len(varResults(iCol, iRow)) - 2)
                            Else
                                varResults(iCol, iRow) = varEventi(1, 0)
                            End If
                        Else
                            varResults(iCol, iRow) = DataClass.DS.List(iCol, iRow)
                        End If
                    End If
                
                Case Else
                    varResults(iCol, iRow) = DataClass.DS.List(iCol, iRow)
                    
            End Select
            
            DoEvents
        Next iCol
    Next iRow

    RefreshUnboundGrid TDBGrid1, varResults, GridArraySelect, DataClass.DS.HowMany, MAX_COLS
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

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim i As Long
    Dim Rows As Long
   
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
                mskField15.SetFocus
            End If
        
        Case fsSelection
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

Private Sub Command2_Click()
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    MyExportGridToDelimitedFile mskFileName.Text, TDBGrid1, varResults, ProgressBar1, 0
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command1_Click()
    Dim sFolder As String
    
    sFolder = GetFolder(hwnd)
    mskFileName = sFolder
    mskFileName.Enabled = True
End Sub

Private Sub Command5_Click()
    mskFileName.Enabled = True
End Sub

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
    
    mskFileName.Text = gmRegistry.GetSetting("Options", "RS2 - Controllo_Viaggi", "c:\export.csv", True)
    
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
    
    SetState fsparaminput
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "RS2 - Controllo_Viaggi", mskFileName.Text, True
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set frmVVTSelect2 = Nothing
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    Frame1(mintCurFrame).Visible = False
    mintCurFrame = TabStrip1.SelectedItem.Index
End Sub
