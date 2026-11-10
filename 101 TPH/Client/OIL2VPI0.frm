VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form OIL2VPI0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EVENTI E INCIDENTI"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   9630
   Begin VB.Frame Frame1 
      Caption         =   "Paralisi"
      Height          =   3735
      Index           =   5
      Left            =   0
      TabIndex        =   87
      Top             =   2880
      Visible         =   0   'False
      Width           =   9615
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1200
         TabIndex        =   88
         Tag             =   "31"
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   3240
         TabIndex        =   89
         Tag             =   "32"
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   1920
         Tag             =   "31"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora Inizio"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   91
         Top             =   480
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   3960
         Tag             =   "32"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora Fine"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   90
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dirottamento"
      Height          =   3735
      Index           =   4
      Left            =   0
      TabIndex        =   73
      Top             =   2880
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3000
         Picture         =   "OIL2VPI0.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   83
         Tag             =   "30"
         Top             =   960
         Width           =   375
      End
      Begin MSMask.MaskEdBox MaskEdBox33 
         Height          =   255
         Left            =   1680
         TabIndex        =   80
         Tag             =   "29"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox34 
         Height          =   255
         Left            =   1680
         TabIndex        =   84
         Tag             =   "30"
         Top             =   960
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
         Caption         =   "Data consegna"
         Height          =   255
         Index           =   45
         Left            =   240
         TabIndex        =   85
         Top             =   960
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   45
         Left            =   3480
         Tag             =   "30"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   44
         Left            =   8280
         Tag             =   "29"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Destinazione"
         Height          =   255
         Index           =   44
         Left            =   480
         TabIndex        =   82
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
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
         Left            =   3120
         TabIndex        =   81
         Tag             =   "29"
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spandimento"
      Height          =   3735
      Index           =   3
      Left            =   0
      TabIndex        =   69
      Top             =   2880
      Visible         =   0   'False
      Width           =   9615
      Begin MSMask.MaskEdBox MaskEdBox30 
         Height          =   255
         Left            =   1440
         TabIndex        =   70
         Tag             =   "26"
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   32
         Left            =   8760
         Tag             =   "26"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Misure adottate"
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rientro"
      Height          =   3735
      Index           =   2
      Left            =   0
      TabIndex        =   72
      Top             =   2880
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3000
         Picture         =   "OIL2VPI0.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   840
         Width           =   375
      End
      Begin MSMask.MaskEdBox MaskEdBox31 
         Height          =   255
         Left            =   1680
         TabIndex        =   74
         Tag             =   "27"
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox32 
         Height          =   255
         Left            =   1680
         TabIndex        =   78
         Tag             =   "28"
         Top             =   840
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
         Caption         =   "Data rientro"
         Height          =   255
         Index           =   43
         Left            =   600
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   43
         Left            =   3480
         Tag             =   "28"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   42
         Left            =   7200
         Tag             =   "27"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deposito Rientro"
         Height          =   255
         Index           =   42
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   1335
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
         Left            =   2280
         TabIndex        =   75
         Tag             =   "27"
         Top             =   480
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contaminazione"
      Height          =   3735
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   9615
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   1080
         TabIndex        =   37
         Tag             =   "14"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Tag             =   "15"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   4440
         TabIndex        =   42
         Tag             =   "16"
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   1080
         TabIndex        =   45
         Tag             =   "17"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Tag             =   "18"
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   4440
         TabIndex        =   50
         Tag             =   "19"
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   1080
         TabIndex        =   53
         Tag             =   "20"
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   255
         Left            =   1080
         TabIndex        =   56
         Tag             =   "21"
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox12 
         Height          =   255
         Left            =   4440
         TabIndex        =   58
         Tag             =   "22"
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Tag             =   "23"
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   1080
         TabIndex        =   64
         Tag             =   "24"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox17 
         Height          =   255
         Left            =   4440
         TabIndex        =   66
         Tag             =   "25"
         Top             =   3240
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
         Index           =   22
         Left            =   7800
         Tag             =   "25"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Unità di misura 4"
         Height          =   255
         Index           =   26
         Left            =   3000
         TabIndex        =   68
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   5040
         TabIndex        =   67
         Tag             =   "25"
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   21
         Left            =   2520
         Tag             =   "24"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità 4"
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   65
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label9 
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
         Tag             =   "23"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prodotto 4"
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   62
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   20
         Left            =   6000
         Tag             =   "23"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   19
         Left            =   7800
         Tag             =   "22"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Unità di misura 3"
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   60
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Left            =   5040
         TabIndex        =   59
         Tag             =   "22"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   2520
         Tag             =   "21"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità 3"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   57
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label7 
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
         TabIndex        =   55
         Tag             =   "20"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prodotto 3"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   54
         Top             =   2040
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   17
         Left            =   6000
         Tag             =   "20"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   7800
         Tag             =   "19"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Unità di misura 2"
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   52
         Top             =   1560
         Width           =   1335
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
         Left            =   5040
         TabIndex        =   51
         Tag             =   "19"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   2520
         Tag             =   "18"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità 2"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   49
         Top             =   1560
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
         Left            =   2520
         TabIndex        =   47
         Tag             =   "17"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prodotto 2"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   6000
         Tag             =   "17"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   7800
         Tag             =   "16"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Unità di misura 1"
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   44
         Top             =   720
         Width           =   1335
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
         Left            =   5040
         TabIndex        =   43
         Tag             =   "16"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   2520
         Tag             =   "15"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità 1"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   735
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
         Left            =   2520
         TabIndex        =   39
         Tag             =   "14"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prodotto 1"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   6000
         Tag             =   "14"
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9615
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Tag             =   "38"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtNoteEstese 
         Height          =   765
         Left            =   5160
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Top             =   1920
         Width           =   4335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "36"
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   2640
         Picture         =   "OIL2VPI0.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Annullato"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   10
         Tag             =   "39"
         Top             =   600
         Width           =   1935
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Tag             =   "4"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Tag             =   "5"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox24 
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Tag             =   "6"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox25 
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Tag             =   "7"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox26 
         Height          =   255
         Left            =   6480
         TabIndex        =   22
         Tag             =   "8"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox27 
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Tag             =   "9"
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox28 
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Tag             =   "10"
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox29 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Tag             =   "12"
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Tag             =   "13"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   255
         Left            =   8880
         TabIndex        =   94
         Tag             =   "40"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Società"
         Height          =   255
         Index           =   2
         Left            =   7800
         TabIndex        =   95
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   9240
         Tag             =   "38"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Stato evento"
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   93
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Osservazioni"
         Height          =   255
         Index           =   46
         Left            =   5160
         TabIndex        =   86
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsabile"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   4680
         Tag             =   "36"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   7200
         Tag             =   "13"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Unità di misura"
         Height          =   255
         Index           =   9
         Left            =   2640
         TabIndex        =   35
         Top             =   1320
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
         Left            =   4440
         TabIndex        =   34
         Tag             =   "13"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   41
         Left            =   2280
         Tag             =   "12"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità"
         Height          =   255
         Index           =   41
         Left            =   480
         TabIndex        =   33
         Top             =   1320
         Width           =   735
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
         Left            =   1920
         TabIndex        =   32
         Tag             =   "10"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo ev. prec."
         Height          =   255
         Index           =   40
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   40
         Left            =   4680
         Tag             =   "10"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordine prec."
         Height          =   255
         Index           =   39
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   39
         Left            =   4680
         Tag             =   "9"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora Fine"
         Height          =   255
         Index           =   38
         Left            =   5640
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   38
         Left            =   7200
         Tag             =   "8"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data evento"
         Height          =   255
         Index           =   37
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   37
         Left            =   3120
         Tag             =   "6"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora Inizio"
         Height          =   255
         Index           =   36
         Left            =   3600
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   36
         Left            =   5160
         Tag             =   "7"
         Top             =   960
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
         Left            =   1920
         TabIndex        =   16
         Tag             =   "4"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo evento"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   4680
         Tag             =   "4"
         Top             =   240
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
         Left            =   1920
         TabIndex        =   14
         Tag             =   "5"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo evento"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   4680
         Tag             =   "5"
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   6600
      Width           =   9615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "OIL2VPI0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
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
         Index           =   2
         Left            =   7920
         Picture         =   "OIL2VPI0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
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
         Left            =   240
         Picture         =   "OIL2VPI0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2VPI0.frx":060C
         TabIndex        =   1
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7800
         Picture         =   "OIL2VPI0.frx":070E
         TabIndex        =   2
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2VPI0.frx":0810
         TabIndex        =   0
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7470
      Width           =   9630
      _ExtentX        =   16986
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
            Object.Width           =   11536
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14.27"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15/06/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
      Top             =   360
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
            Picture         =   "OIL2VPI0.frx":0912
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2VPI0.frx":0A24
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2VPI0.frx":0B36
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2VPI0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DataClass As CViaggioEvento
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private mintCurFrame As Integer
Private mstrIDSocieta As String
Private mOnChanging As Boolean

Private ATE As CTipoEvento
Private AME As CMotivoEvento

Public Property Get IDSocieta() As String
    IDSocieta = mstrIDSocieta
End Property

Public Property Let IDSocieta(strIDSocieta As String)
    mstrIDSocieta = strIDSocieta
End Property

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
'                    If GetRecord(mskLink, mskLink1, mskField1) Then
'                        SetState fsEdit
'                    Else
'                        If State = fsNone Then
'                            SetState fsNone
'                        End If
'                    End If
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
                Case fsInsert
                    blnResult = InsertRecord
                Case Else
                    blnResult = False
            End Select
            DoCommand fcExit
        Case fcUndo
            If State = fsEdit Then
                SetState fsEdit
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

Public Function GetRecord(strIDViaggio As String, strIDConsegna As String, strIDRiga As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strIDViaggio
        .Fields(.KeyFields(2)).FindValue = strIDConsegna
        .Fields(.KeyFields(3)).FindValue = strIDRiga
        
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    TxtNoteEstese = DataClass.DS.Fields("37").value
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
        If Len(TxtNoteEstese) > 0 Then
            DataClass.DS.Fields("37").NewValue = UCase(Replace(TxtNoteEstese, Chr(13), " "))
            DataClass.DS.Fields("37").Updated = True
        End If
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
    
    If DataClass.DS.Fields.Validate Then
        If Len(TxtNoteEstese) > 0 Then
            DataClass.DS.Fields("37").NewValue = UCase(Replace(TxtNoteEstese, Chr(13), " "))
            DataClass.DS.Fields("37").Updated = True
        End If
        If DataClass.BeginSave(blnConcurError) Then
            If blnConcurError Then
                If MsgBox(MSG_CONCUR_ERR, _
                           vbOKCancel + vbExclamation + vbDefaultButton2, _
                          "Attenzione") = vbOK Then
                    GoTo EndProc
                Else
                    DataClass.DS.DA.EndTrans ETROLLBACK
                    GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, _
                              DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value, _
                              DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value
                    GoTo ErrHandler
                End If
            Else
                GoTo EndProc
            End If
            GoTo EndProc
        Else
            strMsg = "Lo scarico non può venire modificato"
            MsgBox strMsg, vbCritical + vbOKOnly, "IMPOSSIBILE SALVARE"
            DataClass.DS.RefreshClient
            GoTo ErrHandler
        End If
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
        Case fsEdit
    End Select
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 1
            Set M = MaskEdBox24
            strTitle = "Data evento"
        Case 2
            Set M = MaskEdBox34
            strTitle = "Data di consegna"
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
    
    Loading = True
    
    mintCurFrame = 1
    Set DataClass = New CViaggioEvento
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set AME = New CMotivoEvento
    With AME.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ATE = New CTipoEvento
    With ATE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set OIL2VPI0 = Nothing
End Sub





Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property


Private Sub MaskEdBox2_Change()
    Dim i As Integer
    Dim intFrameIndex As Integer
    
    If Len(MaskEdBox2.Text) = 3 Then
    
'        ATE.DS.Fields.ResetFields
'        ATE.DS.Field("Società").FindValue = MaskEdBox15.Text
'        ATE.DS.Field("Codice").FindValue = MaskEdBox2.Text
'
'        If Not ATE.Find() Or ATE.DS.HowMany = 0 Then
'            MaskEdBox2.Text = ""
'            lblField3.Caption = ""
'            Exit Sub
'        End If
'        lblField3.Caption = ATE.DS.Field("Descrizione").value

        For i = 1 To Frame1.UBound
            Frame1(i).Visible = False
        Next i
        
        intFrameIndex = 0
        Select Case MaskEdBox2.Text
            Case "007", "008"
                intFrameIndex = 1
            Case "009", "010"
                intFrameIndex = 3
            Case "011"
                intFrameIndex = 2
            Case "012"
                intFrameIndex = 4
            Case "013", "014"
                intFrameIndex = 5
        End Select
        If intFrameIndex > 0 Then
            Frame1(intFrameIndex).Visible = True
        End If
    End If
End Sub

Private Sub mskField1_LostFocus()
'    TrackLostFocus mskField1
End Sub



Private Sub MaskEdBox3_Change()
'On Error GoTo ErrHandler
'    If Not mOnChanging And Len(MaskEdBox3.Text) = 3 Then
'
'        AME.DS.Fields.ResetFields
'        AME.DS.Field("Società").FindValue = MaskEdBox15.Text
'        AME.DS.Field("Codice").FindValue = MaskEdBox3.Text
'
'        If Not AME.Find() Or AME.DS.HowMany = 0 Then
'            mOnChanging = True
'            MaskEdBox3.Text = ""
'            mOnChanging = False
'            lblField4.Caption = ""
'            Exit Sub
'        End If
'        lblField4.Caption = AME.DS.Field("Descrizione").value
'    End If
'
'ErrHandler:
'    mOnChanging = False
End Sub
