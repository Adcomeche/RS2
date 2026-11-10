VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2VVP0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODOTTO CONSEGNATO"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   9810
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   1
      Left            =   240
      TabIndex        =   72
      Top             =   960
      Width           =   9375
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "26"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         Caption         =   "Quantità Conse&gnata"
         Height          =   1815
         Left            =   6840
         TabIndex        =   27
         Top             =   2160
         Width           =   2295
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Tag             =   "10"
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   480
            TabIndex        =   29
            Tag             =   "11"
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   480
            TabIndex        =   30
            Tag             =   "12"
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   255
            Left            =   480
            TabIndex        =   31
            Tag             =   "13"
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Kg"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   20
            Left            =   1920
            Tag             =   "10"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "L"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   19
            Left            =   1920
            Tag             =   "11"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pz"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   1920
            Tag             =   "12"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cnf"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   1920
            Tag             =   "13"
            Top             =   1440
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "&Quantità Trasportata"
         Height          =   1815
         Left            =   6840
         TabIndex        =   22
         Top             =   240
         Width           =   2295
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Tag             =   "6"
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Tag             =   "7"
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   480
            TabIndex        =   25
            Tag             =   "8"
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Tag             =   "9"
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   1920
            Tag             =   "9"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cnf"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   64
            Top             =   1440
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   15
            Left            =   1920
            Tag             =   "8"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pz"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   1920
            Tag             =   "7"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "L"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   255
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   1920
            Tag             =   "6"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Kg"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   255
         End
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Tag             =   "22"
         Top             =   1440
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
         Left            =   1440
         TabIndex        =   11
         Tag             =   "21"
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField11 
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Tag             =   "3"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Tag             =   "4"
         Top             =   720
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
         Left            =   1440
         TabIndex        =   7
         Tag             =   "5"
         Top             =   1080
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
         Left            =   1440
         TabIndex        =   13
         Tag             =   "19"
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Tag             =   "20"
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Tag             =   "14"
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox17 
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Tag             =   "15"
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   32
         Left            =   6120
         Tag             =   "15"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   2880
         Tag             =   "14"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fine &Contalitri"
         Height          =   255
         Index           =   27
         Left            =   3480
         TabIndex        =   20
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ini&zio Contalitri"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   3600
         Tag             =   "26"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Tipo"
         Height          =   255
         Index           =   25
         Left            =   600
         TabIndex        =   16
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   22
         Left            =   3600
         Tag             =   "20"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ri&f. Esterno 2"
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   21
         Left            =   3600
         Tag             =   "19"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Rif. Esterno 1"
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   4800
         Tag             =   "5"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P&os. Doganale"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
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
         Left            =   2040
         TabIndex        =   82
         Tag             =   "5"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   4800
         Tag             =   "4"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Imballo"
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   615
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
         Left            =   2040
         TabIndex        =   81
         Tag             =   "4"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   6360
         Tag             =   "3"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Prodotto"
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   735
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
         Left            =   2880
         TabIndex        =   80
         Tag             =   "3"
         Top             =   360
         Width           =   3375
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
         Left            =   2040
         TabIndex        =   75
         Tag             =   "21"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A&liquota IVA"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   4800
         Tag             =   "21"
         Top             =   1800
         Width           =   240
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
         Left            =   2040
         TabIndex        =   74
         Tag             =   "22"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mo&d. Pagamento"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   4800
         Tag             =   "22"
         Top             =   1440
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   2
      Left            =   240
      TabIndex        =   83
      Top             =   960
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Frame Frame6 
         Caption         =   "S&tato"
         Height          =   2295
         Left            =   6720
         TabIndex        =   51
         Top             =   240
         Width           =   2415
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   1200
            TabIndex        =   52
            Tag             =   "27"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   1200
            TabIndex        =   54
            Tag             =   "29"
            Top             =   1080
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField8 
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Tag             =   "28"
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox19 
            Height          =   255
            Left            =   1200
            TabIndex        =   55
            Tag             =   "30"
            Top             =   1440
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   1200
            TabIndex        =   56
            Tag             =   "31"
            Top             =   1800
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   28
            Left            =   1800
            Tag             =   "31"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 5"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   89
            Top             =   1800
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   27
            Left            =   1800
            Tag             =   "30"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 4"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   88
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 3"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   1800
            Tag             =   "29"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   1800
            Tag             =   "28"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 2"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   1800
            Tag             =   "27"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 1"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Tag             =   "16"
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Tag             =   "17"
         Top             =   1800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Tag             =   "23"
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField6 
         Height          =   255
         Left            =   1440
         TabIndex        =   38
         Tag             =   "24"
         Top             =   1080
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
         TabIndex        =   36
         Tag             =   "25"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox18 
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Tag             =   "18"
         Top             =   2160
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox21 
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Tag             =   "32"
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox22 
         Height          =   255
         Left            =   1440
         TabIndex        =   48
         Tag             =   "33"
         Top             =   2880
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox23 
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         Tag             =   "34"
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dettaglio Pre&fatt."
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   31
         Left            =   3240
         Tag             =   "34"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Or&dine Cliente"
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   30
         Left            =   3240
         Tag             =   "32"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ri&ga Ordine"
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   47
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   29
         Left            =   2280
         Tag             =   "33"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "S&carico Rif."
         Height          =   255
         Index           =   30
         Left            =   480
         TabIndex        =   43
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   2280
         Tag             =   "18"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   23
         Left            =   3120
         Tag             =   "25"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sconto &Valore"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   2280
         Tag             =   "24"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Sconto %"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3120
         Tag             =   "23"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Prezzo"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   24
         Left            =   2280
         Tag             =   "17"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Consegna  Ri&f."
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   41
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   3240
         Tag             =   "16"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Viaggio Rif."
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   39
         Top             =   1440
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4575
      Left            =   120
      TabIndex        =   32
      Top             =   600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dati Prodotto &1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dati Prodotto &2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Annullato"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   60
      Tag             =   "35"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   120
      TabIndex        =   76
      Top             =   5280
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
         Picture         =   "OIL2VVP0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   77
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
         Picture         =   "OIL2VVP0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   78
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
         Picture         =   "OIL2VVP0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2VVP0.frx":0306
         TabIndex        =   58
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7800
         Picture         =   "OIL2VVP0.frx":0408
         TabIndex        =   59
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2VVP0.frx":050A
         TabIndex        =   57
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Top             =   6225
      Width           =   9810
      _ExtentX        =   17304
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
            Object.Width           =   11853
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
   Begin MSMask.MaskEdBox mskField1 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Tag             =   "2"
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   12648447
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLink 
      Height          =   255
      Left            =   5880
      TabIndex        =   85
      Tag             =   "0"
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLink1 
      Height          =   255
      Left            =   8520
      TabIndex        =   86
      Tag             =   "1"
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
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
            Picture         =   "OIL2VVP0.frx":060C
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2VVP0.frx":071E
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2VVP0.frx":0830
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   34
      Left            =   9240
      Tag             =   "1"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   33
      Left            =   7320
      Tag             =   "0"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Consegna"
      Height          =   255
      Index           =   29
      Left            =   7680
      TabIndex        =   87
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero Viaggio"
      Height          =   255
      Index           =   28
      Left            =   4560
      TabIndex        =   84
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Numero"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   0
      Left            =   2160
      Tag             =   "2"
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "OIL2VVP0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DataClass As CViaggioProdotto
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private mintCurFrame As Integer

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(mskLink, mskLink1, mskField1) Then
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
    Dim Respok, Respuesta As String
    Select Case Cmd
        Case fcSave
        ' Control RS3
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Esta opción solo esta disponible desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
                Select Case State
                    Case fsEdit
                        blnResult = UpdateRecord
                    Case Else
                        blnResult = False
                End Select
                SetState fsEdit
            End If
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
            TabStrip1.Enabled = True
            TabStrip1.Tabs(1).Selected = True
            MaskEdBox9.SetFocus
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
    Set DataClass = New CViaggioProdotto
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
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
    Set OIL2VVP0 = Nothing
End Sub





Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property


Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index

End Sub


