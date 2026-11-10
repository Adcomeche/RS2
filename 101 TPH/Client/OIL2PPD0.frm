VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2PPD0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIGA PREFATTURA"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   10230
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   1
      Left            =   120
      TabIndex        =   68
      Top             =   840
      Width           =   9975
      Begin VB.Frame Frame9 
         Caption         =   "Valorizzazione Interna"
         Height          =   3255
         Left            =   5400
         TabIndex        =   73
         Top             =   240
         Width           =   4455
         Begin MSMask.MaskEdBox MaskEdBox38 
            Height          =   255
            Left            =   1200
            TabIndex        =   44
            Tag             =   "41"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox39 
            Height          =   255
            Left            =   1200
            TabIndex        =   46
            Tag             =   "42"
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox40 
            Height          =   255
            Left            =   3000
            TabIndex        =   50
            Tag             =   "43"
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox42 
            Height          =   255
            Left            =   1200
            TabIndex        =   48
            Tag             =   "47"
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox43 
            Height          =   255
            Left            =   1200
            TabIndex        =   52
            Tag             =   "44"
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox44 
            Height          =   255
            Left            =   3720
            TabIndex        =   54
            Tag             =   "45"
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   1200
            TabIndex        =   55
            Tag             =   "67"
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox19 
            Height          =   255
            Left            =   1200
            TabIndex        =   162
            Tag             =   "69"
            Top             =   2160
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox23 
            Height          =   255
            Left            =   3240
            TabIndex        =   164
            Tag             =   "70"
            Top             =   2160
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox24 
            Height          =   255
            Left            =   1200
            TabIndex        =   166
            Tag             =   "71"
            Top             =   2520
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox41 
            Height          =   255
            Left            =   3240
            TabIndex        =   168
            Tag             =   "72"
            Top             =   2520
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox48 
            Height          =   255
            Left            =   1200
            TabIndex        =   181
            Tag             =   "77"
            Top             =   2880
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   55
            Left            =   2040
            Tag             =   "78"
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Serv. Exent."
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   182
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Serv Sp."
            Height          =   255
            Index           =   43
            Left            =   2400
            TabIndex        =   169
            Top             =   2520
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   42
            Left            =   4080
            Tag             =   "72"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ore Serv Sp."
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   167
            Top             =   2520
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   37
            Left            =   2040
            Tag             =   "71"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Op/Bomb"
            Height          =   255
            Index           =   36
            Left            =   2400
            TabIndex        =   165
            Top             =   2160
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   35
            Left            =   4080
            Tag             =   "70"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Canon Puerto"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   163
            Top             =   2160
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   23
            Left            =   2040
            Tag             =   "69"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   20
            Left            =   2520
            Tag             =   "67"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pedaggio"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   154
            Top             =   1800
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   52
            Left            =   4080
            Tag             =   "42"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label13 
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
            Left            =   1800
            TabIndex        =   77
            Tag             =   "42"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fascia Tar."
            Height          =   255
            Index           =   53
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona Fascia"
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label12 
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
            Left            =   1800
            TabIndex        =   74
            Tag             =   "41"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   51
            Left            =   4080
            Tag             =   "41"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   50
            Left            =   4080
            Tag             =   "43"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tariffa"
            Height          =   255
            Index           =   51
            Left            =   2400
            TabIndex        =   49
            Top             =   1080
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   41
            Left            =   2040
            Tag             =   "47"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Km Raggio"
            Height          =   255
            Index           =   42
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   39
            Left            =   2520
            Tag             =   "44"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Importo"
            Height          =   255
            Index           =   40
            Left            =   480
            TabIndex        =   51
            Top             =   1440
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   4080
            Tag             =   "45"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Forzato"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   53
            Top             =   1440
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Valorizzazione Esterna"
         Height          =   3255
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   4455
         Begin MSMask.MaskEdBox MaskEdBox15 
            Height          =   255
            Left            =   1200
            TabIndex        =   32
            Tag             =   "33"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   255
            Left            =   1200
            TabIndex        =   34
            Tag             =   "34"
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox18 
            Height          =   255
            Left            =   3000
            TabIndex        =   38
            Tag             =   "35"
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox32 
            Height          =   255
            Left            =   1200
            TabIndex        =   36
            Tag             =   "39"
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox33 
            Height          =   255
            Left            =   1200
            TabIndex        =   40
            Tag             =   "36"
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox14 
            Height          =   255
            Left            =   3720
            TabIndex        =   42
            Tag             =   "37"
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   255
            Left            =   1200
            TabIndex        =   118
            Tag             =   "66"
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   1200
            TabIndex        =   170
            Tag             =   "73"
            Top             =   2160
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   3240
            TabIndex        =   172
            Tag             =   "74"
            Top             =   2160
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox31 
            Height          =   255
            Left            =   1200
            TabIndex        =   174
            Tag             =   "75"
            Top             =   2520
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox46 
            Height          =   255
            Left            =   3240
            TabIndex        =   176
            Tag             =   "76"
            Top             =   2520
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox47 
            Height          =   255
            Left            =   1200
            TabIndex        =   179
            Tag             =   "78"
            Top             =   2880
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   54
            Left            =   2040
            Tag             =   "77"
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Serv. Exent."
            Height          =   255
            Index           =   55
            Left            =   120
            TabIndex        =   180
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Serv Sp."
            Height          =   255
            Index           =   44
            Left            =   2400
            TabIndex        =   177
            Top             =   2520
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   43
            Left            =   4080
            Tag             =   "76"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ore Serv Sp."
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   175
            Top             =   2520
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   32
            Left            =   2040
            Tag             =   "75"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Op/Bomb"
            Height          =   255
            Index           =   14
            Left            =   2400
            TabIndex        =   173
            Top             =   2160
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   4080
            Tag             =   "74"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Canon Puerto"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   171
            Top             =   2160
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   2040
            Tag             =   "73"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   2520
            Tag             =   "66"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pedaggio"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   153
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Forzato"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   41
            Top             =   1440
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   4080
            Tag             =   "37"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Importo"
            Height          =   255
            Index           =   39
            Left            =   480
            TabIndex        =   39
            Top             =   1440
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   38
            Left            =   2520
            Tag             =   "36"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Km Raggio"
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   35
            Top             =   1080
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   33
            Left            =   2040
            Tag             =   "39"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tariffa"
            Height          =   255
            Index           =   24
            Left            =   2400
            TabIndex        =   37
            Top             =   1080
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   24
            Left            =   4080
            Tag             =   "35"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   4080
            Tag             =   "33"
            Top             =   360
            Width           =   240
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
            Left            =   1800
            TabIndex        =   72
            Tag             =   "33"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona Fascia"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   4080
            Tag             =   "34"
            Top             =   720
            Width           =   240
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
            Left            =   1800
            TabIndex        =   71
            Tag             =   "34"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fascia Tar."
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.CommandButton Command6 
         Height          =   495
         Left            =   4680
         Picture         =   "OIL2PPD0.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   1560
         Width           =   615
      End
      Begin VB.Frame Frame7 
         Caption         =   "Quantità"
         Height          =   975
         Left            =   120
         TabIndex        =   78
         Top             =   3480
         Width           =   9735
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   2160
            TabIndex        =   57
            Tag             =   "30"
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   2160
            TabIndex        =   61
            Tag             =   "9"
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   5280
            TabIndex        =   59
            Tag             =   "29"
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField6 
            Height          =   255
            Left            =   8040
            TabIndex        =   141
            Tag             =   "60"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   8040
            TabIndex        =   142
            Tag             =   "59"
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Q.ta esterna L"
            Height          =   255
            Index           =   17
            Left            =   6720
            TabIndex        =   144
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   9360
            Tag             =   "60"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Q.ta esterna Kg"
            Height          =   255
            Index           =   16
            Left            =   6480
            TabIndex        =   143
            Top             =   600
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   9360
            Tag             =   "59"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantità Scaricata Kg"
            Height          =   255
            Index           =   15
            Left            =   3600
            TabIndex        =   58
            Top             =   240
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   15
            Left            =   6240
            Tag             =   "29"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   19
            Left            =   3120
            Tag             =   "9"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantità Ordine iniziale L"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   1935
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   36
            Left            =   3120
            Tag             =   "30"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantità Scaricata L"
            Height          =   255
            Index           =   37
            Left            =   480
            TabIndex        =   56
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   3
      Left            =   120
      TabIndex        =   119
      Top             =   840
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Importo Servicios Exentos errato"
         Height          =   195
         Index           =   26
         Left            =   4200
         TabIndex        =   178
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Importo Servicios Especiales errato"
         Height          =   195
         Index           =   25
         Left            =   4200
         TabIndex        =   161
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Numero Ore Servicios Especiales errato"
         Height          =   195
         Index           =   24
         Left            =   4200
         TabIndex        =   160
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Importo Servicio Buques a Puerto errato"
         Height          =   195
         Index           =   23
         Left            =   4200
         TabIndex        =   159
         Top             =   2280
         Width           =   3495
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Importo Servicio Operario Bunker errato"
         Height          =   195
         Index           =   22
         Left            =   4200
         TabIndex        =   158
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Data Grado Servizio discordante"
         Height          =   195
         Index           =   21
         Left            =   4200
         TabIndex        =   155
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Importo pedaggio errato"
         Height          =   195
         Index           =   20
         Left            =   4200
         TabIndex        =   140
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Tariffa applicata errata"
         Height          =   195
         Index           =   19
         Left            =   4200
         TabIndex        =   139
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Tipo di Tariffa errato"
         Height          =   195
         Index           =   18
         Left            =   4200
         TabIndex        =   138
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Nome imbarcazione errato"
         Height          =   195
         Index           =   17
         Left            =   4200
         TabIndex        =   137
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Destinatario errato"
         Height          =   195
         Index           =   16
         Left            =   4200
         TabIndex        =   136
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Progressivo Viaggio errato"
         Height          =   195
         Index           =   15
         Left            =   4200
         TabIndex        =   135
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "NIF/DNI Autista errato"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   134
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Autista errato"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   133
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Targa Rimorchio errata"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   132
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Targa Motrice errata"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   131
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Codice UdT per Deposito errata"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   130
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Unità di Trasporto errata"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   129
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Tipo Servizio"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   128
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Data Servizio"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   127
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Vettore errato"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   126
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Importi discordanti"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   125
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Diff. Quantità scar./ord."
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   124
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Diff. Km"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   123
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Tariffa Minima"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   122
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Dati discordanti"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   121
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkAnomalie 
         Caption         =   "Proposta non trovata"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   120
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   4
      Left            =   120
      TabIndex        =   145
      Top             =   840
      Visible         =   0   'False
      Width           =   9975
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3060
         Picture         =   "OIL2PPD0.frx":0442
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   4140
         Width           =   240
      End
      Begin VB.CommandButton cmdDeleteAddebito 
         Caption         =   "Elimina"
         Height          =   375
         Left            =   3000
         Picture         =   "OIL2PPD0.frx":0974
         TabIndex        =   151
         Tag             =   "IMPORT"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "OIL2PPD0.frx":0A76
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   4155
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   1680
         Picture         =   "OIL2PPD0.frx":0B78
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   4140
         Width           =   240
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   3675
         Left            =   120
         OleObjectBlob   =   "OIL2PPD0.frx":10AA
         TabIndex        =   146
         Top             =   240
         Width           =   9735
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Apri"
         Height          =   375
         Left            =   1560
         Picture         =   "OIL2PPD0.frx":7D4C
         TabIndex        =   148
         Tag             =   "EDITLINK"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2PPD0.frx":7E4E
         TabIndex        =   150
         Tag             =   "ADDLINK"
         Top             =   4080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Visible         =   0   'False
      Width           =   9975
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   114
         TabStop         =   0   'False
         Tag             =   "49"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         Caption         =   "Flag Stato"
         Height          =   1935
         Left            =   7080
         TabIndex        =   82
         Top             =   240
         Width           =   2175
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   960
            TabIndex        =   84
            Tag             =   "51"
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
            Left            =   960
            TabIndex        =   86
            Tag             =   "53"
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
            Left            =   960
            TabIndex        =   88
            Tag             =   "52"
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   960
            TabIndex        =   89
            Tag             =   "54"
            Top             =   1440
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stato 4"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   93
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   1560
            Tag             =   "53"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   1560
            Tag             =   "52"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stato 3"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   92
            Top             =   720
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   1560
            Tag             =   "51"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stato 2"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   91
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stato 5"
            Height          =   255
            Index           =   30
            Left            =   240
            TabIndex        =   90
            Top             =   1440
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   31
            Left            =   1560
            Tag             =   "54"
            Top             =   1440
            Width           =   240
         End
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   6600
         TabIndex        =   27
         Tag             =   "20"
         Top             =   2760
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
         Left            =   6600
         TabIndex        =   29
         Tag             =   "21"
         Top             =   3120
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
         Left            =   4800
         TabIndex        =   11
         Tag             =   "6"
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Tag             =   "7"
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox21 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Tag             =   "2"
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox22 
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Tag             =   "3"
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Tag             =   "4"
         Top             =   1320
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
         Left            =   1440
         TabIndex        =   13
         Tag             =   "5"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox34 
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Tag             =   "12"
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox35 
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Tag             =   "13"
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox36 
         Height          =   255
         Left            =   6000
         TabIndex        =   23
         Tag             =   "14"
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox37 
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Tag             =   "15"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   6600
         TabIndex        =   25
         Tag             =   "55"
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox27 
         Height          =   255
         Left            =   6600
         TabIndex        =   94
         Tag             =   "22"
         Top             =   3480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox28 
         Height          =   255
         Left            =   1440
         TabIndex        =   97
         Tag             =   "23"
         Top             =   600
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
         Left            =   1440
         TabIndex        =   98
         Tag             =   "24"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox30 
         Height          =   255
         Left            =   1440
         TabIndex        =   103
         Tag             =   "25"
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox25 
         Height          =   255
         Left            =   1440
         TabIndex        =   105
         Tag             =   "27"
         Top             =   3480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox26 
         Height          =   255
         Left            =   1440
         TabIndex        =   106
         Tag             =   "28"
         Top             =   3840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox45 
         Height          =   255
         Left            =   1440
         TabIndex        =   107
         Tag             =   "26"
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   255
         Left            =   4200
         TabIndex        =   156
         Tag             =   "68"
         Top             =   2400
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
         Caption         =   "Data G. Serv."
         Height          =   255
         Index           =   35
         Left            =   3000
         TabIndex        =   157
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   34
         Left            =   5520
         Tag             =   "68"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   40
         Left            =   9240
         Tag             =   "49"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Stato Riga"
         Height          =   255
         Index           =   41
         Left            =   6240
         TabIndex        =   115
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P&os. Doganale"
         Height          =   255
         Index           =   46
         Left            =   240
         TabIndex        =   113
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Imballo"
         Height          =   255
         Index           =   47
         Left            =   720
         TabIndex        =   112
         Top             =   3480
         Width           =   615
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
         Left            =   2040
         TabIndex        =   111
         Tag             =   "28"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   48
         Left            =   4800
         Tag             =   "28"
         Top             =   3840
         Width           =   240
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
         Left            =   2040
         TabIndex        =   110
         Tag             =   "27"
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   49
         Left            =   4800
         Tag             =   "27"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prodotto"
         Height          =   255
         Index           =   54
         Left            =   600
         TabIndex        =   109
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label14 
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
         TabIndex        =   108
         Tag             =   "26"
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   53
         Left            =   4800
         Tag             =   "26"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   44
         Left            =   2760
         Tag             =   "25"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Rif. Est. Dest."
         Height          =   255
         Index           =   45
         Left            =   240
         TabIndex        =   104
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   45
         Left            =   6600
         Tag             =   "24"
         Top             =   960
         Width           =   240
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
         Left            =   3000
         TabIndex        =   102
         Tag             =   "24"
         Top             =   960
         Width           =   3495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   46
         Left            =   5280
         Tag             =   "23"
         Top             =   600
         Width           =   240
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
         Left            =   2040
         TabIndex        =   101
         Tag             =   "23"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Destinazione"
         Height          =   255
         Index           =   49
         Left            =   360
         TabIndex        =   100
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente"
         Height          =   255
         Index           =   50
         Left            =   600
         TabIndex        =   99
         Top             =   600
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   47
         Left            =   9600
         Tag             =   "22"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deposito"
         Height          =   255
         Index           =   48
         Left            =   5760
         TabIndex        =   96
         Top             =   3480
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
         Left            =   7200
         TabIndex        =   95
         Tag             =   "22"
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   9600
         Tag             =   "55"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "U.d.T."
         Height          =   255
         Index           =   29
         Left            =   5760
         TabIndex        =   24
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
         Left            =   7200
         TabIndex        =   69
         Tag             =   "55"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   2760
         Tag             =   "7"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Viaggio"
         Height          =   255
         Index           =   27
         Left            =   360
         TabIndex        =   16
         Top             =   2400
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   6600
         Tag             =   "14"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Scarico"
         Height          =   255
         Index           =   26
         Left            =   5280
         TabIndex        =   22
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   4920
         Tag             =   "13"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Consegna"
         Height          =   255
         Index           =   25
         Left            =   3360
         TabIndex        =   20
         Top             =   2040
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   2760
         Tag             =   "12"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Viaggio"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   2760
         Tag             =   "5"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Cons. &Rich."
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   2760
         Tag             =   "4"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Invio Ord."
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   29
         Left            =   5280
         Tag             =   "3"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ri&ga Ordine"
         Height          =   255
         Index           =   33
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   30
         Left            =   3120
         Tag             =   "2"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Or&dine"
         Height          =   255
         Index           =   34
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   22
         Left            =   6600
         Tag             =   "7"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ri&f. Esterno 2"
         Height          =   255
         Index           =   22
         Left            =   3600
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   21
         Left            =   6600
         Tag             =   "6"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Rif. Esterno 1"
         Height          =   255
         Index           =   21
         Left            =   3600
         TabIndex        =   10
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
         Left            =   7200
         TabIndex        =   65
         Tag             =   "21"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ente Emitt."
         Height          =   255
         Index           =   9
         Left            =   5640
         TabIndex        =   28
         Top             =   3120
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   9600
         Tag             =   "21"
         Top             =   3120
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
         Left            =   7200
         TabIndex        =   64
         Tag             =   "20"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Società"
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   26
         Top             =   2760
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   9600
         Tag             =   "20"
         Top             =   2760
         Width           =   240
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   0
      TabIndex        =   62
      Top             =   480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dati Valorizzazioni"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dati Viaggio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Anomalie"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Addebiti"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   66
      Top             =   5640
      Width           =   10215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "OIL2PPD0.frx":7F50
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   116
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
         Index           =   2
         Left            =   8640
         Picture         =   "OIL2PPD0.frx":8052
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "OIL2PPD0.frx":8154
         TabIndex        =   87
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1575
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
         Picture         =   "OIL2PPD0.frx":8256
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   79
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
         Index           =   0
         Left            =   5760
         Picture         =   "OIL2PPD0.frx":8358
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   75
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
         Index           =   1
         Left            =   3840
         Picture         =   "OIL2PPD0.frx":845A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F12 - Viaggio"
         Height          =   375
         Left            =   5640
         Picture         =   "OIL2PPD0.frx":855C
         TabIndex        =   85
         Tag             =   "CUSTOM2"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F11 - Ordine"
         Height          =   375
         Left            =   3720
         Picture         =   "OIL2PPD0.frx":865E
         TabIndex        =   83
         Tag             =   "CUSTOM1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F10 - Proposta"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2PPD0.frx":8760
         TabIndex        =   81
         Tag             =   "BROWSE"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2PPD0.frx":8862
         TabIndex        =   80
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   63
      Top             =   6495
      Width           =   10230
      _ExtentX        =   18045
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
            Object.Width           =   12594
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:19"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10/01/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mskField1 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Tag             =   "1"
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
      Left            =   4320
      TabIndex        =   3
      Tag             =   "0"
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   0
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
            Picture         =   "OIL2PPD0.frx":8964
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2PPD0.frx":8A76
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2PPD0.frx":8B88
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero Testata Pref."
      Height          =   255
      Index           =   28
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Numero Riga"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   0
      Left            =   2160
      Tag             =   "1"
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "OIL2PPD0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DataClass As CPrefatturaDettaglio
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private PAD As CPrefatturaDettaglioAddebito
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private mintCurFrame As Integer
Private XAADD As XArrayDB

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(mskLink, mskField1) Then
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
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcCustom1
            Set LinkForm = DisplayForm(fcORDIOCT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord DataClass.DS.Fields("2").value
                    .SetState fsEdit
                    .Command1.Enabled = False
                    Me.Enabled = False
                End With
            End If
        Case fccustom2
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord DataClass.DS.Fields("12").value
                    .SetState fsEdit
                    .Command1.Enabled = False
                    Me.Enabled = False
                End With
            End If
        Case fcBrowse
            Set LinkForm = DisplayForm(fcPREFGP1)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord DataClass.DS.Fields("6").value, DataClass.DS.Fields("26").value
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
                Case Else
                    blnResult = False
            End Select
            If blnResult Then
                SetState fsEdit
            End If
        Case fcUndo
            If (State = fsEdit) Then
                SetState fsEdit
            End If
        
        Case fcOpenLinkAdd
            If (State = fsEdit) Then
                Set LinkForm = DisplayForm(fcPREFPAD)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .DataClass.DS.Fields("0").ReadOnly = False
                        .DataClass.DS.Fields("1").ReadOnly = False
                        .DataClass.DS.Fields("2").ReadOnly = False
                        .DataClass.DS.Fields("3").ReadOnly = False
                        .DataClass.DS.Fields("4").ReadOnly = False
                        .SetState fsInsert
                        .DataClass.DS.Fields("0").NewValue = mskLink.Text
                        .DataClass.DS.Fields("0").Updated = True
                        .DataClass.DS.Fields("1").NewValue = mskField1.Text
                        .DataClass.DS.Fields("1").Updated = True
                        .mskLink1.Text = mskLink.Text
                        .mskLink1.Enabled = False
                        .mskLink2.Text = mskField1.Text
                        .mskLink2.Enabled = False
                        Me.Enabled = False
                    End With
                End If
            End If
        
        Case fcOpenLinkEdit
            If (State = fsEdit) And Not IsNull(TDBGrid2.bookmark) Then
                Set LinkForm = DisplayForm(fcPREFPAD)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .GetRecord XAADD.value(TDBGrid2.bookmark, 0), _
                                    XAADD.value(TDBGrid2.bookmark, 1), _
                                    XAADD.value(TDBGrid2.bookmark, 2), _
                                    XAADD.value(TDBGrid2.bookmark, 4), _
                                    XAADD.value(TDBGrid2.bookmark, 6)
                        .SetState fsEdit
                        Me.Enabled = False
                    End With
                End If
            End If
        
        Case fcLinkRefresh
            DisplayDetail
            
        Case fcImport
            If (State = fsEdit) And Not IsNull(TDBGrid2.bookmark) Then
                If MsgBox("Sei sicuro di voler eliminare l'addebito?", _
                           vbYesNo + vbExclamation + vbDefaultButton2, _
                          "Attenzione") = vbYes Then

                    If Not PAD.Delete(XAADD.value(TDBGrid2.bookmark, 0), _
                                    XAADD.value(TDBGrid2.bookmark, 1), _
                                    XAADD.value(TDBGrid2.bookmark, 2), _
                                    XAADD.value(TDBGrid2.bookmark, 4), _
                                    XAADD.value(TDBGrid2.bookmark, 6)) Then
                    End If
                    DisplayDetail
                End If
            End If

        
        Case fcExit
            Unload Me
    End Select
End Sub

Public Function GetRecord(strIDOrdine As String, intIDRiga As Integer) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strIDOrdine
        .Fields(.KeyFields(2)).FindValue = intIDRiga
        
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    For i = 0 To chkAnomalie.UBound
                        chkAnomalie(i).value = _
                            IIf(((.Fields("50").value) And (2 ^ i)) <> 0, vbChecked, vbUnchecked)
                    Next i
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
    Dim i As Integer
    
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
                              DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value
                    GoTo ErrHandler
                End If
            Else
                GoTo EndProc
            End If
            GoTo EndProc
        Else
            strMsg = "La riga di prefattura non può venire modificata"
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
    For i = 0 To chkAnomalie.UBound
        chkAnomalie(i).value = _
            IIf(((DataClass.DS.Fields("50").value) And (2 ^ i)) <> 0, vbChecked, vbUnchecked)
    Next i
    UpdateRecord = True
End Function

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsEdit
            TabStrip1.Enabled = True
            TabStrip1.Tabs(1).Selected = True
'            MaskEdBox21.SetFocus
    End Select
End Sub

Private Sub Command6_Click()
    MaskEdBox43 = MaskEdBox33
    MaskEdBox40 = MaskEdBox18
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
    Set DataClass = New CPrefatturaDettaglio
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With

    Set XAADD = New XArrayDB

    Set PAD = New CPrefatturaDettaglioAddebito
    With PAD.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set XAADD = Nothing
    Set PAD = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set OIL2PPD0 = Nothing
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

Private Sub DisplayDetail()
    ' TODO: Order
    PAD.DS.Fields.ResetFields
    PAD.DS.Fields("0").FindValue = DataClass.DS.Fields("0").value
    PAD.DS.Fields("1").FindValue = DataClass.DS.Fields("1").value
        
    If Not PAD.GetSelection() Then
    End If
    
    If PAD.DS.HowMany > 0 Then
        RefreshXArray XAADD, TDBGrid2, PAD.DS.List
    Else
        ClearXArray XAADD, TDBGrid2
    End If
End Sub

Private Sub TDBGrid2_DblClick()
    DoCommand fcOpenLinkEdit
End Sub
