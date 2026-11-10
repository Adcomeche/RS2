VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2DAM1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DOCUMENTI AMMINISTRATIVI"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   10830
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1035
      Left            =   0
      TabIndex        =   124
      Top             =   5460
      Width           =   10815
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "OIL2DAM1Z30.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   135
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
         Left            =   8640
         Picture         =   "OIL2DAM1Z30.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   134
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
         Index           =   4
         Left            =   1920
         Picture         =   "OIL2DAM1Z30.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   133
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
         Index           =   3
         Left            =   3600
         Picture         =   "OIL2DAM1Z30.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   131
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2DAM1Z30.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   675
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
         Left            =   2400
         Picture         =   "OIL2DAM1Z30.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   4680
         Picture         =   "OIL2DAM1Z30.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F8 - ..."
         Height          =   375
         Left            =   4560
         Picture         =   "OIL2DAM1Z30.frx":070E
         TabIndex        =   126
         Tag             =   "EXPORT"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Riga Dettaglio"
         Height          =   375
         Left            =   2280
         Picture         =   "OIL2DAM1Z30.frx":0810
         TabIndex        =   128
         Tag             =   "EDITLINK"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Aggiungi Riga"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2DAM1Z30.frx":0912
         TabIndex        =   130
         Tag             =   "OPEN"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2DAM1Z30.frx":0A14
         TabIndex        =   132
         Tag             =   "PRINT"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2DAM1Z30.frx":0B16
         TabIndex        =   136
         Tag             =   "UNDO"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "OIL2DAM1Z30.frx":0C18
         TabIndex        =   137
         Tag             =   "EXIT"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2DAM1Z30.frx":0D1A
         TabIndex        =   138
         Tag             =   "SAVE"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   52
      Tag             =   "*"
      Top             =   0
      Width           =   10815
      Begin VB.Frame fraMaster 
         Caption         =   "Dati Documento 3"
         ForeColor       =   &H00FF0000&
         Height          =   3135
         Index           =   3
         Left            =   1560
         TabIndex        =   68
         Tag             =   "*"
         Top             =   840
         Visible         =   0   'False
         Width           =   9135
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "48"
            Top             =   1740
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            ItemData        =   "OIL2DAM1Z30.frx":0E1C
            Left            =   7200
            List            =   "OIL2DAM1Z30.frx":0E1E
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Tag             =   "10"
            Top             =   1020
            Width           =   1815
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Tag             =   "28"
            Top             =   1380
            Width           =   1815
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Tag             =   "29"
            Top             =   2100
            Width           =   1815
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Tag             =   "30"
            Top             =   2460
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   1320
            TabIndex        =   39
            Tag             =   "39"
            Top             =   2100
            Visible         =   0   'False
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   1320
            TabIndex        =   40
            Tag             =   "40"
            Top             =   2460
            Visible         =   0   'False
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Tag             =   "41"
            Top             =   2820
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   1320
            TabIndex        =   41
            Tag             =   "43"
            Top             =   2820
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   255
            Left            =   5880
            TabIndex        =   43
            Tag             =   "42"
            Top             =   2820
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField6 
            Height          =   255
            Left            =   1320
            TabIndex        =   32
            Tag             =   "4"
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox33 
            Height          =   255
            Left            =   1320
            TabIndex        =   33
            Tag             =   "34"
            Top             =   660
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox40 
            Height          =   255
            Left            =   1320
            TabIndex        =   35
            Tag             =   "35"
            Top             =   1380
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox50 
            Height          =   255
            Left            =   3000
            TabIndex        =   37
            Tag             =   "36"
            Top             =   1740
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox52 
            Height          =   255
            Left            =   1320
            TabIndex        =   36
            Tag             =   "38"
            Top             =   1740
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox54 
            Height          =   255
            Left            =   5880
            TabIndex        =   38
            Tag             =   "37"
            Top             =   1740
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   255
            Left            =   1320
            TabIndex        =   34
            Tag             =   "53"
            Top             =   1020
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Rag.s. agg."
            Height          =   255
            Index           =   34
            Left            =   240
            TabIndex        =   123
            Top             =   1020
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo"
            Height          =   255
            Index           =   22
            Left            =   6360
            TabIndex        =   118
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Destinazione"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   107
            Top             =   300
            Width           =   975
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
            Left            =   2760
            TabIndex        =   106
            Tag             =   "4"
            Top             =   300
            Width           =   4455
         End
         Begin VB.Label Label1 
            Caption         =   "CAP"
            Height          =   255
            Index           =   50
            Left            =   240
            TabIndex        =   105
            Top             =   1740
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Località"
            Height          =   255
            Index           =   65
            Left            =   2280
            TabIndex        =   104
            Tag             =   "L36"
            Top             =   1740
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Indirizzo"
            Height          =   255
            Index           =   66
            Left            =   240
            TabIndex        =   103
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Rag. sociale"
            Height          =   255
            Index           =   67
            Left            =   240
            TabIndex        =   102
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pr."
            Height          =   255
            Index           =   69
            Left            =   5520
            TabIndex        =   101
            Top             =   1740
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "CAP"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   100
            Top             =   2820
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Località"
            Height          =   255
            Index           =   10
            Left            =   2280
            TabIndex        =   99
            Top             =   2820
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Indirizzo"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   98
            Top             =   2460
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Destinatario"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   97
            Top             =   2100
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pr."
            Height          =   255
            Index           =   14
            Left            =   5520
            TabIndex        =   96
            Top             =   2820
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Esito"
            Height          =   255
            Index           =   20
            Left            =   6360
            TabIndex        =   95
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Gestione"
            Height          =   255
            Index           =   28
            Left            =   6360
            TabIndex        =   94
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Incasso"
            Height          =   255
            Index           =   41
            Left            =   6360
            TabIndex        =   93
            Top             =   2100
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Saldo"
            Height          =   255
            Index           =   31
            Left            =   6360
            TabIndex        =   92
            Top             =   2460
            Width           =   615
         End
      End
      Begin VB.Frame fraMaster 
         Caption         =   "Dati Documento 1"
         ForeColor       =   &H00FF0000&
         Height          =   3135
         Index           =   1
         Left            =   1560
         TabIndex        =   54
         Tag             =   "*"
         Top             =   840
         Width           =   9135
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   1
            Left            =   8520
            Picture         =   "OIL2DAM1Z30.frx":0E20
            Style           =   1  'Graphical
            TabIndex        =   51
            TabStop         =   0   'False
            Tag             =   "CAL"
            Top             =   2760
            Width           =   375
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   0
            Left            =   8520
            Picture         =   "OIL2DAM1Z30.frx":0F22
            Style           =   1  'Graphical
            TabIndex        =   50
            TabStop         =   0   'False
            Tag             =   "CAL"
            Top             =   2400
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskField7 
            Height          =   255
            Left            =   960
            TabIndex        =   7
            Tag             =   "5"
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox34 
            Height          =   255
            Left            =   960
            TabIndex        =   8
            Tag             =   "19"
            Top             =   1320
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox35 
            Height          =   255
            Left            =   960
            TabIndex        =   10
            Tag             =   "20"
            Top             =   2040
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox36 
            Height          =   255
            Left            =   2760
            TabIndex        =   12
            Tag             =   "21"
            Top             =   2400
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox37 
            Height          =   255
            Left            =   5760
            TabIndex        =   13
            Tag             =   "22"
            Top             =   2400
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox38 
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Tag             =   "23"
            Top             =   2400
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox28 
            Height          =   255
            Left            =   960
            TabIndex        =   15
            Tag             =   "3"
            Top             =   2760
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
            Left            =   960
            TabIndex        =   5
            Tag             =   "14"
            Top             =   600
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
            Left            =   5880
            TabIndex        =   6
            Tag             =   "15"
            Top             =   600
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
            Left            =   7200
            TabIndex        =   14
            Tag             =   "6"
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
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Tag             =   "8"
            Top             =   240
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
            Left            =   5880
            TabIndex        =   4
            Tag             =   "9"
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox29 
            Height          =   255
            Left            =   7200
            TabIndex        =   17
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
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   6720
            TabIndex        =   18
            Tag             =   "24"
            Top             =   2040
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField20 
            Height          =   255
            Left            =   6720
            TabIndex        =   16
            Tag             =   "25"
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   960
            TabIndex        =   9
            Tag             =   "52"
            Top             =   1680
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Rag. agg."
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   122
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. Fisc."
            Height          =   255
            Index           =   17
            Left            =   5880
            TabIndex        =   111
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "P. IVA"
            Height          =   255
            Index           =   5
            Left            =   5940
            TabIndex        =   110
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Data Cons."
            Height          =   255
            Index           =   40
            Left            =   6240
            TabIndex        =   87
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "N° Doc."
            Height          =   255
            Index           =   36
            Left            =   5040
            TabIndex        =   86
            Top             =   240
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
            Left            =   1680
            TabIndex        =   85
            Tag             =   "8"
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Doc."
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Data Doc."
            Height          =   255
            Index           =   11
            Left            =   6240
            TabIndex        =   83
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Deposito"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   82
            Top             =   600
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
            Left            =   1680
            TabIndex        =   81
            Tag             =   "14"
            Top             =   600
            Width           =   2895
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
            Left            =   6600
            TabIndex        =   80
            Tag             =   "15"
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Magazzino"
            Height          =   255
            Index           =   33
            Left            =   4920
            TabIndex        =   79
            Top             =   600
            Width           =   855
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
            Left            =   1680
            TabIndex        =   78
            Tag             =   "3"
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Società"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   77
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "CAP"
            Height          =   255
            Index           =   56
            Left            =   120
            TabIndex        =   67
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Prov."
            Height          =   255
            Index           =   55
            Left            =   5280
            TabIndex        =   66
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Località"
            Height          =   255
            Index           =   54
            Left            =   2040
            TabIndex        =   65
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Indirizzo"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   64
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Rag. soc."
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   63
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   960
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
            Left            =   2400
            TabIndex        =   60
            Tag             =   "5"
            Top             =   960
            Width           =   6615
         End
      End
      Begin VB.Frame fraMaster 
         Caption         =   "Dati Documento  2"
         ForeColor       =   &H00FF0000&
         Height          =   3135
         Index           =   2
         Left            =   1560
         TabIndex        =   62
         Tag             =   "*"
         Top             =   840
         Visible         =   0   'False
         Width           =   9135
         Begin MSMask.MaskEdBox mskField4 
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Tag             =   "11"
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox23 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Tag             =   "44"
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField14 
            Height          =   255
            Left            =   5280
            TabIndex        =   22
            Tag             =   "13"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox41 
            Height          =   255
            Left            =   5280
            TabIndex        =   23
            Tag             =   "18"
            Top             =   720
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
            Left            =   1680
            TabIndex        =   21
            Tag             =   "12"
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField9 
            Height          =   255
            Left            =   5280
            TabIndex        =   24
            Tag             =   "26"
            Top             =   1080
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
            Left            =   7560
            TabIndex        =   26
            Tag             =   "27"
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox53 
            Height          =   255
            Left            =   1680
            TabIndex        =   25
            Tag             =   "16"
            Top             =   1440
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLink 
            Height          =   255
            Left            =   7560
            TabIndex        =   30
            Tag             =   "45"
            Top             =   2160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   255
            Left            =   1680
            TabIndex        =   28
            Tag             =   "31"
            Top             =   2160
            Visible         =   0   'False
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Tag             =   "32"
            Top             =   2520
            Visible         =   0   'False
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLink1 
            Height          =   255
            Left            =   7560
            TabIndex        =   31
            Tag             =   "46"
            Top             =   2520
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox14 
            Height          =   255
            Left            =   5280
            TabIndex        =   27
            Tag             =   "47"
            Top             =   1800
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Banca"
            Height          =   255
            Index           =   21
            Left            =   3960
            TabIndex        =   117
            Top             =   1800
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
            Left            =   6000
            TabIndex        =   116
            Tag             =   "47"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Consegna"
            Height          =   255
            Index           =   29
            Left            =   6000
            TabIndex        =   115
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Numero Viaggio"
            Height          =   255
            Index           =   2
            Left            =   6000
            TabIndex        =   114
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Note 1"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   113
            Top             =   2160
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Note 2"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   112
            Top             =   2520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Mod. Pag."
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   109
            Top             =   1440
            Width           =   855
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
            Left            =   2400
            TabIndex        =   108
            Tag             =   "16"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Numero Ordine"
            Height          =   255
            Index           =   6
            Left            =   6000
            TabIndex        =   91
            Top             =   1440
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
            Left            =   6000
            TabIndex        =   90
            Tag             =   "26"
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Vettore"
            Height          =   255
            Index           =   8
            Left            =   3960
            TabIndex        =   89
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Iva Documento"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   1080
            Width           =   1455
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
            Left            =   6000
            TabIndex        =   74
            Tag             =   "18"
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Vendita"
            Height          =   255
            Index           =   59
            Left            =   3960
            TabIndex        =   73
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Valuta"
            Height          =   255
            Index           =   19
            Left            =   3960
            TabIndex        =   72
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblField7 
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
            Left            =   6000
            TabIndex        =   71
            Tag             =   "13"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Importo Documento"
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. già Incassato"
            Height          =   255
            Index           =   45
            Left            =   120
            TabIndex        =   69
            Top             =   720
            Width           =   1455
         End
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   975
         Left            =   1680
         OleObjectBlob   =   "OIL2DAM1Z30.frx":1024
         TabIndex        =   120
         Tag             =   "DETAIL"
         Top             =   4380
         Width           =   8895
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   975
         Left            =   1680
         OleObjectBlob   =   "OIL2DAM1Z30.frx":9DAE
         TabIndex        =   121
         Tag             =   "DETAIL"
         Top             =   4380
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.Frame fraKey 
         Caption         =   "Documento Amministrativo"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1560
         TabIndex        =   55
         Top             =   120
         Width           =   9135
         Begin VB.CheckBox Check1 
            Caption         =   "Annullato"
            Height          =   255
            Left            =   8040
            TabIndex        =   49
            Tag             =   "33"
            Top             =   360
            Width           =   975
         End
         Begin MSMask.MaskEdBox mskField1 
            Height          =   255
            Left            =   6840
            TabIndex        =   2
            Tag             =   "2"
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   5400
            TabIndex        =   1
            Tag             =   "1"
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField10 
            Height          =   255
            Left            =   1200
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
            Left            =   1920
            TabIndex        =   76
            Tag             =   "0"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Ente Emitt."
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Anno"
            Height          =   255
            Index           =   32
            Left            =   4800
            TabIndex        =   59
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Progr."
            Height          =   255
            Index           =   0
            Left            =   6240
            TabIndex        =   56
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame fraSelect 
         Caption         =   "Selezione Vista"
         ForeColor       =   &H00FF0000&
         Height          =   4515
         Left            =   120
         TabIndex        =   57
         Tag             =   "*"
         Top             =   900
         Width           =   1335
         Begin MSComctlLib.TabStrip tspSelect 
            CausesValidation=   0   'False
            Height          =   1695
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   2990
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dati Doc. &1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dati Doc. &2"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Dati Doc. &3"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
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
               Picture         =   "OIL2DAM1Z30.frx":FF94
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2DAM1Z30.frx":100A6
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2DAM1Z30.frx":101B8
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         CausesValidation=   0   'False
         Height          =   1395
         Left            =   1560
         TabIndex        =   119
         Top             =   4020
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2461
         MultiRow        =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dettaglio Documento"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "IVA Documento"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   53
      Top             =   6495
      Width           =   10830
      _ExtentX        =   19103
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
            Object.Width           =   13653
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15.17"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "03/02/2006"
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
            Picture         =   "OIL2DAM1Z30.frx":102CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":103DC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":104EE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10600
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10712
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10824
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10936
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10A48
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10B5A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10C6C
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10D7E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2DAM1Z30.frx":10E90
            Key             =   "Camera"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2DAM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_COLS = 14
Public WithEvents DataClass As CDocumentoAmministrativo
Attribute DataClass.VB_VarHelpID = -1

Public GridArrayDetail As ArrayGridCls
Public DetailClass As CRigaDocumentoAmministrativo

Private XAIva As XArrayDB

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private mintCurFrame As Integer
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private mvarFunction As FunctionCodes
Private mstrIDADEDef As String
Private mstrIDAEEDef As String
Private mstrIDAVADef As String
Private mstrIDATVDef As String

' (PM) 03/01/06 - <
'Private AXU As CUtente

Private AXX As CParametro
Dim strIDValuta As String
Private ADZ As CDestinazione
Private ACL As CCliente

' (PM) 03/01/06 - <
'Private TAQ As CAliquotaImposta

Private AMP As CPagamento

Private Sub DisplayADZ(Optional IDADZ As Variant, _
                       Optional IDACL As Variant)

    Dim lngSaveMousePointer As Long
    
    lngSaveMousePointer = Me.MousePointer
    
    With ADZ
        Me.MousePointer = vbHourglass
        If Not IsMissing(IDADZ) Then
            .DS.Fields("0").FindValue = IDADZ
            .Find
            If .DS.HowMany > 0 Then
                MaskEdBox33 = .DS.Fields("1").value
                MaskEdBox17 = .DS.Fields("2").value
                MaskEdBox40 = .DS.Fields("4").value
                MaskEdBox52 = .DS.Fields("7").value
                MaskEdBox50 = .DS.Fields("5").value
                MaskEdBox54 = .DS.Fields("6").value
                Select Case .DS.Fields("11").value
                    Case "002"
                        Combo3 = "EXTRARETE"
                    Case "001"
                        Combo3 = "RETE"
                    Case Else
                        Combo3 = "EXTRARETE"
                End Select
            End If
        Else
            If Not IsMissing(IDACL) Then
                .DS.Fields("8").FindValue = IDACL
                .GetList
                If .DS.HowMany = 1 Then
                    .DS.Fields("0").FindValue = .DS.List(0, 0)
                    .Find
                    mskField6 = .DS.List(0, 0)
                    DisplayADZ .DS.List(0, 0)
                End If
            End If
        End If
        Me.MousePointer = lngSaveMousePointer
    End With
End Sub


Private Sub DisplayDetail()
    Dim i As Integer
    Dim BufferIva As Variant
    
    With DetailClass
        .DS.Fields("0").FindValue = DataClass.DS.Fields("0").value
        .DS.Fields("1").FindValue = DataClass.DS.Fields("1").value
        .DS.Fields("2").FindValue = DataClass.DS.Fields("2").value
        
        .GetList2
        RefreshUnboundGrid TDBgrid1, .DS.List, GridArrayDetail, _
                           .DS.HowMany, MAX_COLS
        TDBgrid1.Refresh
    End With
    
    BufferIva = DataClass.GetIvaDetail(DataClass.DS.Fields("0").value, _
                                        DataClass.DS.Fields("1").value, _
                                        DataClass.DS.Fields("2").value)
    
    If Not IsEmpty(BufferIva) Then
        RefreshXArray XAIva, TDBGrid2, BufferIva
    
    Else
        ClearXArray XAIva, TDBGrid2
    End If
End Sub

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property


Private Sub ClearForm()
    DataClass.DS.ClearClient
    ClearUnboundGrid TDBgrid1, GridArrayDetail, MAX_COLS
    ClearXArray XAIva, TDBGrid2
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strIDViaggio As String
    Dim strIDATO As String
    Dim strSelectionFormula As String
    Dim strReportName As String
    ReDim varReportsList(0 To 4, 0 To 1) As Variant
    Dim i As Long
    Dim lngRowCount As Long
    Dim lngMax As Long
    
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
            If blnResult Then
                If State = fsInsert Then
                    DoCommand fcOpen
                Else
                    SetState fsEdit
                End If
            End If
        Case fcUndo
            Select Case State
                Case fsEdit
                    SetState fsEdit
                Case fsInsert
                    SetState fsInsert
            End Select
        Case fcOpenLinkEdit
            If (State = fsEdit) Then
                Set LinkForm = DisplayForm(fcDAMDAR0)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .GetRecord GridArrayDetail.value(TDBgrid1.bookmark, 0), _
                                   GridArrayDetail.value(TDBgrid1.bookmark, 1), _
                                   GridArrayDetail.value(TDBgrid1.bookmark, 2), _
                                   GridArrayDetail.value(TDBgrid1.bookmark, 3)
                        .SetState fsEdit
                        Me.Enabled = False
                    End With
                End If
            End If
        Case fcPrint
            strIDATO = UCase(DataClass.DS.Fields("8").value)
            
            DataClass.DS.Fields("0").FindValue = DataClass.DS.Fields("0").value
            DataClass.DS.Fields("1").FindValue = DataClass.DS.Fields("1").value
            DataClass.DS.Fields("2").FindValue = DataClass.DS.Fields("2").value
            
            strSelectionFormula = DataClass.DS.MakeKeyCond(True)
            Select Case strIDATO
                Case "FAT", "NCR"
                    strReportName = "OIL2FAT0.rpt"
                Case "NTP"
                     If mstrIDAVADef = "E" Then
                        strReportName = "OIL2NTPE.rpt"
                     Else
                        strReportName = "OIL2NTP0.rpt"
                     End If
                Case "NTE"
                    strReportName = "OIL2NTP0.rpt"
                Case "MCE"
                    strReportName = "OIL2MCE0.rpt"
            End Select
            varReportsList(0, 0) = "Stampa Documento"
            varReportsList(1, 0) = strReportName
            varReportsList(2, 0) = 0
' (PM) 03/01/06 - <
            ' *** GR 16/09/2001
'            If SysVersions32 = "95" Then
'                varReportsList(3, 0) = 2500
'                varReportsList(4, 0) = 2800
'            Else
                varReportsList(3, 0) = 0
                varReportsList(4, 0) = 0
'            End If
            ' ***
            varReportsList(0, 1) = "Stampa Bolla permuta"
            varReportsList(1, 1) = "OIL2NTP1.rpt"
            varReportsList(2, 1) = 0
' (PM) 03/01/06 - <
            ' *** GR 16/09/2001
'            If SysVersions32 = "95" Then
'                varReportsList(3, 1) = 2500
'                varReportsList(4, 1) = 2800
'            Else
                varReportsList(3, 1) = 0
                varReportsList(4, 1) = 0
'            End If
            ' ***
            CRPrint varReportsList, strSelectionFormula, Me
        Case fcOpen
            lngRowCount = GridArrayDetail.RowCount
            If lngRowCount > 0 Then
                For i = 0 To lngRowCount - 1
                    If GridArrayDetail.value(i, 3) > lngMax Then
                        lngMax = GridArrayDetail.value(i, 3)
                    End If
                Next i
            End If
        
        
            Set LinkForm = DisplayForm(fcDAMDAR0)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .SetState fsInsert
                    .MaskEdBox15.Text = DataClass.DS.Fields("0").value
                    .MaskEdBox15.Enabled = False
                    .MaskEdBox13.Text = DataClass.DS.Fields("1").value
                    .MaskEdBox13.Enabled = False
                    .mskField1.Text = DataClass.DS.Fields("2").value
                    .mskField1.Enabled = False
                    .MaskEdBox20 = lngMax + 1
                    .MaskEdBox20.Enabled = False
                    Me.Enabled = False
                    
                End With
            End If
        Case fcLinkRefresh
            DisplayDetail
        Case fcExport
            DataClass.SetStato "S"
        Case fcExit
            Unload Me
    End Select
End Sub

Public Function GetRecord(IDDeposito As Variant, _
                          Anno As Variant, _
                          IDMovimento As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields("0").FindValue = IDDeposito
        .Fields("1").FindValue = Anno
        .Fields("2").FindValue = IDMovimento
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
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
    Dim varNumeroDoc As Variant
    Dim strTipoDoc As String
    
    DataClass.DS.ScanClient 1
    
    If Len(MaskEdBox15) = 0 Then
        strTipoDoc = DataClass.DS.Fields("8").NewValue
        If strTipoDoc = "NCR" Then
            strTipoDoc = "FAT"
        End If
        varNumeroDoc = DataClass.DS.DA.GetCounter("docunda0", "ndanrprg", _
                                        "ndaidasc", DataClass.DS.Fields("3").NewValue, 0, _
                                        "ndaidaee", DataClass.DS.Fields("0").NewValue, 0, _
                                        "ndaidato", strTipoDoc, 0)
        If varNumeroDoc = -1 Then
    '                strErrDesc = "Impossibile leggere il numeratore per il documento." & vbCrLf & "Società: " & DataClass.DS.Field("Società").NewValue & vbCrLf & "Tipo Documento: " & strTipoDoc
            GoTo ErrHandler
        End If
        DataClass.DS.Fields("9").NewValue = varNumeroDoc
    End If
    
    If DataClass.DS.Fields("16").NewValue <> vbNullString Then
        AMP.DS.Fields("Codice").FindValue = DataClass.DS.Fields("16").NewValue
        If AMP.Find Then
            If AMP.DS.HowMany > 0 Then
                DataClass.DS.Fields("17").NewValue = AMP.DS.Fields("GiorniDil").value
            Else
                DataClass.DS.Fields("17").NewValue = 0
            End If
        End If
    End If
    
    If DataClass.DS.Fields.Validate Then
        If Not DataClass.Insert Then
            GoTo ErrHandler
        End If
        MaskEdBox15 = varNumeroDoc
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
    Dim CErr As CError
    Dim i As Long
    
    'Scansione della finestra
    If Not DataClass.DS.ScanClient(0) Then
        GoTo ErrHandler
    End If
    
    ' --- Impediamo la modifica se il documento proviene da un viaggio
    If Not DataClass.IsUpdatable Then
        DataClass.DS.DisplayMessage "il documento non è modificabile", Warning
        GoTo ErrHandler
    End If
    ' ---
    
    If DataClass.DS.Fields("16").NewValue <> vbNullString Then
        AMP.DS.Fields("0").FindValue = DataClass.DS.Fields("16").NewValue
        If AMP.Find Then
            If AMP.DS.HowMany > 0 Then
                DataClass.DS.Fields("17").NewValue = AMP.DS.Fields("GiorniDil").value
            Else
                DataClass.DS.Fields("17").NewValue = 0
            End If
        End If
    End If
    
    
    If DataClass.DS.Fields.Validate Then
        DataClass.BeginSave blnConcurError
            If blnConcurError Then
                If MsgBox(MSG_CONCUR_ERR, _
                           vbOKCancel + vbExclamation + vbDefaultButton2, _
                          "Attenzione") = vbOK Then
                    GoTo EndProc
                Else
                    DataClass.DS.DA.EndTrans ETROLLBACK
'                    GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
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
    Dim strIDTipoDoc As String
    Dim strIDDeposito As String
    Dim strIDMagazzino As String
    Dim strIDEnteEmittente As String
    Dim strDataDoc As String
    Dim strDataCon As String
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsEdit
'            If MaskEdBox20.Enabled Then
'                MaskEdBox20.SetFocus
'            End If
            tspSelect.Enabled = True
            tspSelect.Tabs(1).Selected = True

        Case fsInsert
            strIDTipoDoc = MaskEdBox20
            strIDDeposito = mskField15
            strIDEnteEmittente = mskField10
            strIDMagazzino = MaskEdBox7
            strDataDoc = MaskEdBox2
            strDataCon = MaskEdBox29
            
            ClearForm
            
            MaskEdBox20 = strIDTipoDoc
            mskField15 = strIDDeposito
            mskField10 = strIDEnteEmittente
            MaskEdBox7 = strIDMagazzino
            MaskEdBox2 = strDataDoc
            MaskEdBox29 = strDataCon
            
            Combo1.Text = "GENERATO"
'            Combo3 = "EXTRARETE"
            Combo6 = "NON DA INCASSARE"
            Combo4 = "SALDATO"
            
            MaskEdBox6 = Year(Date)
'            MaskEdBox2 = Date
            
'            mskField15.Text = mstrIDADEDef
'            DataClass.DS.ValidateClientField "14", mstrIDADEDef
'            FieldUpdateDescription "14", Me, mskField15
'            FieldSetDescription "14", Me, mstrIDADEDef, mskField15
'
'
'            mskField10.Text = mstrIDAEEDef
'            DataClass.DS.ValidateClientField "0", mstrIDAEEDef
'            FieldUpdateDescription "0", Me, mskField10
'            FieldSetDescription "0", Me, mstrIDAEEDef, mskField10
            
            mskField14.Text = mstrIDAVADef
            DataClass.DS.ValidateClientField "13", mstrIDAVADef
            FieldUpdateDescription "13", Me, mskField14
            FieldSetDescription "13", Me, mstrIDAVADef, mskField14
            
            If Not Loading Then
                If mskField10.Enabled Then
                    mskField10.SetFocus
                End If
            End If
    End Select
End Sub



Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Data Documento"
        Case 1
            Set M = MaskEdBox29
            strTitle = "Data Consegna"
    End Select
    GetCalendarDate M, strTitle
End Sub
'

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strIDParam As String
    ReDim ColTags(1 To MAX_COLS) As Variant
    
    Loading = True
    
    Set DataClass = New CDocumentoAmministrativo
    Set ErrH = New CErrorHandler
    Set XAIva = New XArrayDB
    
    mintCurFrame = 1
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .Fields("45").ReadOnly = True
        .Fields("46").ReadOnly = True
        .Fields("27").ReadOnly = True
' (PM) 03/01/06 - <
'        .Fields("11").Default = 0
'        .Fields("12").Default = 0
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set .ErrH = ErrH
    End With
    
    
    
    Set DetailClass = New CRigaDocumentoAmministrativo
    Set ErrH = New CErrorHandler
    Set GridArrayDetail = New ArrayGridCls
    
    mintCurFrame = 1
    
    ColTags(1) = "0"
    ColTags(2) = "1"
    ColTags(3) = "2"
    ColTags(4) = "3"
    ColTags(5) = "18"
    ColTags(6) = "4"
    ColTags(7) = "5"
    ColTags(8) = "6"
    ColTags(9) = "7"
    ColTags(10) = "8"
    ColTags(11) = "9"
    ColTags(12) = "10"
    ColTags(13) = "11"
    ColTags(14) = "27"
    
    With DetailClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
' (PM) 03/01/06 - <
'    Set TAQ = New CAliquotaImposta
'    With TAQ.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'    End With
    
    Set AMP = New CPagamento
    With AMP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    ' --- Leggiamo il deposito e l'ente emittente di default
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        AXX.Find
'        mstrIDADEDef = Trim$(.Field("Deposito Default").value)
        mstrIDAEEDef = Trim$(.Fields("16").value)
        mstrIDAVADef = Trim$(.Fields("17").value)

' (PM) 03/01/06 - <
'        mstrIDATVDef = .Field("Tipo Vendita").value
    End With
    Set AXX = Nothing
    ' ---

' (PM) 03/01/06 - <
    ' --- Leggiamo il deposito di default
'    Set AXU = New CUtente
'    With AXU.DS
'        Set .DA = DataConnect.DA
'        Set .ErrH = ErrH
'
'        .Field("Login").FindValue = gmstrPrivateUserName ' DataClass.DS.DA.ConnUserName
'        AXU.Find
'    End With
'
'    mstrIDADEDef = IIf(IsNull(AXU.DS.Field("Deposito").value), vbNullString, AXU.DS.Field("Deposito").value)
'
'    Set AXU = Nothing
    ' ---

    Select Case mvarFunction
        Case fcDAMINSE
            mskField15 = mstrIDADEDef
            mskField10 = mstrIDAEEDef
            MaskEdBox41 = mstrIDATVDef
            MaskEdBox2 = Format(Date, "dd/mm/yy")
            MaskEdBox29 = Format(Date, "dd/mm/yy")
            SetState fsInsert
        Case fcDAMDAM1
            SetState fsEdit
        Case Else
            'SetState fsedit
    End Select
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AXX = Nothing
    Set ADZ = Nothing
    Set ACL = Nothing
' (PM) 03/01/06 - <
'    Set TAQ = Nothing
    Set AMP = Nothing
' (PM) 03/01/06 - <
'    Set AXU = Nothing
    Set DetailClass = Nothing
    Set GridArrayDetail = Nothing
    Set XAIva = Nothing


    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set OIL2DAM1 = Nothing
End Sub


Private Sub mskField6_LostFocus()

    ADZ.DS.Fields("0").FindValue = mskField6
    If ADZ.Find Then
        If ADZ.DS.HowMany > 0 Then
            MaskEdBox33 = ADZ.DS.Fields("1").value
            MaskEdBox17 = ADZ.DS.Fields("2").value
            MaskEdBox40 = ADZ.DS.Fields("4").value
            MaskEdBox50 = ADZ.DS.Fields("5").value
            MaskEdBox54 = ADZ.DS.Fields("6").value
            MaskEdBox52 = ADZ.DS.Fields("7").value
            Select Case ADZ.DS.Fields("11").value
                Case "002"
                    Combo3 = "EXTRARETE"
                Case "001"
                    Combo3 = "RETE"
                Case Else
                    Combo3 = "EXTRARETE"
            End Select

        End If
    End If
End Sub


Private Sub mskField6_Validate(Cancel As Boolean)
'    Dim strText As String
'    Dim blnValidate As Boolean
'    Dim DC As Object
'    Dim strFieldKey As String
'    Dim strFormatString As Variant
'
'    strFieldKey = mskField6.Tag
'    strText = mskField6.Text
'    With DataClass.DS
'        blnValidate = .ValidateClientField(strFieldKey, strText)
'        If Not blnValidate Then
'            .Fields(strFieldKey).ForeignTableFldValue = vbNullString
'        End If
'        If .OmonimySearch Then
'            .OmonimySearch = False
'            DoSearch .Fields(strFieldKey).ForeignTableEntity, _
'                     Me, _
'                     strFieldKey, _
'                     mskField6.TabIndex
'
'' 03/02/06 (PM)
''            DoSearch .Fields(strFieldKey).ForeignTableEntity, _
''                     Me, _
''                     strFieldKey, _
''                     mskField6.TabIndex, , strText
'
'        End If
'
'        If strText <> "?" Then
'            ' --- Aggiorno l'eventuale descrizione collegata al campo ---
'            FieldUpdateDescription strFieldKey, Me, mskField6
'            strText = mskField6.Text
'            FieldSetDescription strFieldKey, Me, strText, mskField6
'        End If
'    End With
End Sub


Private Sub mskField7_LostFocus()
    Dim lngSaveMousePointer As Long
    
    lngSaveMousePointer = Me.MousePointer
    
    If mskField7 <> vbNullString Then
        Me.MousePointer = vbHourglass
'        DataClass.DS.ValidateClientField "5", mskField7
        ACL.DS.Fields("0").FindValue = mskField7
        If ACL.Find Then
            MaskEdBox34 = ACL.DS.Fields("1").value
            MaskEdBox16 = ACL.DS.Fields("2").value
            MaskEdBox35 = ACL.DS.Fields("4").value
            MaskEdBox36 = ACL.DS.Fields("5").value
            MaskEdBox37 = ACL.DS.Fields("6").value
            MaskEdBox38 = ACL.DS.Fields("7").value
            MaskEdBox1 = ACL.DS.Fields("10").value
            mskField20 = ACL.DS.Fields("11").value
            If Len(MaskEdBox53) = 0 Then
                MaskEdBox53 = ACL.DS.Fields("14").value
                DataClass.DS.ValidateClientField "16", MaskEdBox53
                FieldUpdateDescription "16", Me, MaskEdBox53
                FieldSetDescription "16", Me, ACL.DS.Fields("14").value, MaskEdBox53
            End If
            DisplayADZ , mskField7
        End If
        Me.MousePointer = lngSaveMousePointer
    End If
End Sub

Private Sub mskField7_Validate(Cancel As Boolean)
'    Dim strText As String
'    Dim blnValidate As Boolean
'    Dim DC As Object
'    Dim strFieldKey As String
'    Dim strFormatString As Variant
'
'    strFieldKey = mskField7.Tag
'    strText = mskField7.Text
'    With DataClass.DS
'        blnValidate = .ValidateClientField(strFieldKey, strText)
'        If Not blnValidate Then
'            .Fields(strFieldKey).ForeignTableFldValue = vbNullString
'        End If
'        If .OmonimySearch Then
'            .OmonimySearch = False
'            DoSearch .Fields(strFieldKey).ForeignTableEntity, _
'                     Me, _
'                     strFieldKey, _
'                     MaskEdBox7.TabIndex
'
''            DoSearch .Fields(strFieldKey).ForeignTableEntity, _
''                     Me, _
''                     strFieldKey, _
''                     MaskEdBox7.TabIndex, , strText
'
'        End If
'
'        If strText <> "?" Then
'            ' --- Aggiorno l'eventuale descrizione collegata al campo ---
'            FieldUpdateDescription strFieldKey, Me, mskField7
'            strText = mskField7.Text
'            FieldSetDescription strFieldKey, Me, strText, mskField7
'        End If
'    End With
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Index
        Case 1
            TDBgrid1.Visible = True
            TDBGrid2.Visible = False
        Case 2
            TDBGrid2.Visible = True
            TDBgrid1.Visible = False
    End Select
End Sub

Private Sub tspSelect_Click()
    If tspSelect.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
    fraMaster(tspSelect.SelectedItem.Index).Visible = True
    fraMaster(mintCurFrame).Visible = False
    mintCurFrame = tspSelect.SelectedItem.Index
' (PM) 03/01/06 - <
'    Me.Controls("cmdLookup").Visible = False
    Select Case mintCurFrame
        Case 1
            If MaskEdBox20.Enabled Then
                MaskEdBox20.SetFocus
            End If
        Case 2
            'WriteAliquota MaskEdBox16, MaskEdBox18, MaskEdBox19, MaskEdBox2
            If mskField4.Enabled Then
                mskField4.SetFocus
            End If
        Case 3
            If mskField6.Enabled Then
                mskField6.SetFocus
            End If
        Case 4
    End Select
    

End Sub
'Private Sub WriteAliquota(strIDProdotto As String, strIDPosDogEst As String, strIDPosDogUsc As String, varDataMov As Variant)
'Dim varBufferAliquota As Variant
'
'Dim varImportoAliquota As Variant
'Dim varImportoImposta As Variant
'Dim varIDUMAliquota  As Variant
'Dim varQtaAliquota As Variant
'If Not (strIDProdotto = vbNullString Or strIDPosDogEst = vbNullString Or strIDPosDogUsc = vbNullString Or Not IsDate(varDataMov)) Then
'    varBufferAliquota = TAQ.FindByDate(strIDProdotto, strIDPosDogEst, strIDPosDogUsc, varDataMov)
'    If Not IsEmpty(varBufferAliquota) Then
'        varImportoAliquota = varBufferAliquota(0, 0)
'        varIDUMAliquota = varBufferAliquota(1, 0)
'        Select Case varIDUMAliquota
'            Case "L"
'                varQtaAliquota = MaskEdBox8.Text    'Litri
'            Case "L15"
'                varQtaAliquota = MaskEdBox27.Text   'Litri a 15
'            Case "KG"
'                varQtaAliquota = MaskEdBox1.Text    'Kg
'            Case "PZ"
'                varQtaAliquota = MaskEdBox9.Text    'Numero di pezzi
'            Case "CNF"
'                varQtaAliquota = MaskEdBox21.Text    'Numero di confezioni
'        End Select
'        If varQtaAliquota = vbNullString Then
'            varQtaAliquota = 0
'        End If
'        varImportoImposta = Round(varImportoAliquota * varQtaAliquota, 3)
'
'        MaskEdBox25 = varImportoAliquota
'        MaskEdBox26 = varImportoImposta
'    End If
'End If
'End Sub


