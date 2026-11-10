VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2VVT0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIAGGI - VISUALIZZAZIONE E MODIFICA "
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11340
   Begin VB.CommandButton Command10 
      Caption         =   "Cancelar Viaje PALM"
      Height          =   615
      Left            =   960
      Picture         =   "OIL2VVT0.frx":0000
      TabIndex        =   172
      Tag             =   "DELLSERV"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   112
      Top             =   5520
      Width           =   11295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   9600
         Picture         =   "OIL2VVT0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9480
         Picture         =   "OIL2VVT0.frx":0204
         TabIndex        =   104
         Tag             =   "EXIT"
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   8
         Left            =   4920
         Picture         =   "OIL2VVT0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   795
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
         Left            =   7080
         Picture         =   "OIL2VVT0.frx":0408
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F11 - Elimina Consegna"
         Height          =   375
         Left            =   6960
         Picture         =   "OIL2VVT0.frx":093A
         TabIndex        =   103
         Tag             =   "CUSTOM1"
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F9 - Mod. Scarico"
         Height          =   375
         Left            =   4800
         Picture         =   "OIL2VVT0.frx":0A3C
         TabIndex        =   102
         Tag             =   "FIND"
         Top             =   720
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   2640
         Picture         =   "OIL2VVT0.frx":0B3E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F8 - Agg. Dettaglio"
         Height          =   375
         Left            =   2520
         Picture         =   "OIL2VVT0.frx":0C40
         TabIndex        =   101
         Tag             =   "REFRESHLIST"
         Top             =   720
         Width           =   2295
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
         Picture         =   "OIL2VVT0.frx":0D42
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Mod. Consegna"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2VVT0.frx":0E44
         TabIndex        =   100
         Tag             =   "EDITLINK"
         Top             =   720
         Width           =   2415
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   5280
         Picture         =   "OIL2VVT0.frx":0F46
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F6 - Apri Ordine"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2VVT0.frx":1048
         TabIndex        =   99
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2175
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
         Picture         =   "OIL2VVT0.frx":114A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2VVT0.frx":124C
         TabIndex        =   98
         Tag             =   "PRINT"
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
         Left            =   1920
         Picture         =   "OIL2VVT0.frx":134E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   113
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
         Picture         =   "OIL2VVT0.frx":1450
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2VVT0.frx":1552
         TabIndex        =   97
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2VVT0.frx":1654
         TabIndex        =   96
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   111
      Top             =   7935
      Width           =   11340
      _ExtentX        =   20003
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
            Object.Width           =   14552
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
      Height          =   5415
      Left            =   0
      TabIndex        =   110
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame6 
         Height          =   2775
         Index           =   1
         Left            =   240
         TabIndex        =   121
         Top             =   960
         Width           =   10815
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   3
            Left            =   9960
            Picture         =   "OIL2VVT0.frx":1756
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   2400
            Width           =   375
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6840
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Tag             =   "50"
            Top             =   2340
            Width           =   975
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   0
            Left            =   8160
            Picture         =   "OIL2VVT0.frx":1858
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskField9 
            Height          =   255
            Left            =   6840
            TabIndex        =   26
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
            Left            =   6840
            TabIndex        =   28
            Tag             =   "21"
            Top             =   1320
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
            Left            =   1200
            TabIndex        =   13
            Tag             =   "15"
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField7 
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Tag             =   "14"
            Top             =   240
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
            Left            =   6840
            TabIndex        =   32
            Tag             =   "23"
            Top             =   2040
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
            Left            =   1200
            TabIndex        =   15
            Tag             =   "16"
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   6840
            TabIndex        =   22
            Tag             =   "25"
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   255
            Left            =   9720
            TabIndex        =   24
            Tag             =   "27"
            Top             =   600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   3
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   6840
            TabIndex        =   17
            Tag             =   "6"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   255
            Left            =   9720
            TabIndex        =   20
            Tag             =   "7"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:nn"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   6840
            TabIndex        =   30
            Tag             =   "22"
            Top             =   1680
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ProgressBar pgbPortata 
            Height          =   255
            Left            =   1200
            TabIndex        =   154
            Top             =   2400
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Min             =   1e-4
            Scrolling       =   1
         End
         Begin MSMask.MaskEdBox MaskEdBox29 
            Height          =   255
            Left            =   8760
            TabIndex        =   181
            Tag             =   "53"
            Top             =   2400
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "F. Ref."
            Height          =   255
            Index           =   62
            Left            =   8040
            TabIndex        =   182
            Top             =   2400
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   50
            Left            =   10440
            Tag             =   "53"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   47
            Left            =   7800
            Tag             =   "50"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Usa dati GPS"
            Height          =   255
            Index           =   55
            Left            =   5580
            TabIndex        =   167
            Top             =   2400
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Riempimento"
            Height          =   255
            Index           =   52
            Left            =   120
            TabIndex        =   158
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblPortataU 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4680
            TabIndex        =   157
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Vuoto Lt"
            Height          =   255
            Index           =   39
            Left            =   3960
            TabIndex        =   156
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblTotVuoto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4680
            TabIndex        =   155
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblScompartiM 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   153
            Top             =   960
            Width           =   4215
         End
         Begin VB.Label lblScompartiR 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   152
            Top             =   1680
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Scomparti M"
            Height          =   255
            Index           =   50
            Left            =   120
            TabIndex        =   151
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Scomparti R"
            Height          =   255
            Index           =   49
            Left            =   120
            TabIndex        =   150
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblPortataM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   149
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Portata M Lt"
            Height          =   255
            Index           =   48
            Left            =   120
            TabIndex        =   148
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label lblPortataR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3000
            TabIndex        =   147
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Portata R Lt"
            Height          =   255
            Index           =   46
            Left            =   2040
            TabIndex        =   146
            Top             =   2040
            Width           =   855
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
            Left            =   7440
            TabIndex        =   139
            Tag             =   "22"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "De&posito"
            Height          =   255
            Index           =   47
            Left            =   5880
            TabIndex        =   29
            Top             =   1680
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   10440
            Tag             =   "22"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Data inizio"
            Height          =   255
            Index           =   36
            Left            =   6000
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   36
            Left            =   8640
            Tag             =   "6"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ora ini&zio"
            Height          =   255
            Index           =   35
            Left            =   8880
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   35
            Left            =   10440
            Tag             =   "7"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Km d&efinit."
            Height          =   195
            Index           =   12
            Left            =   8880
            TabIndex        =   23
            Top             =   600
            Width           =   765
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   10440
            Tag             =   "27"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Km prev."
            Height          =   195
            Index           =   11
            Left            =   6090
            TabIndex        =   21
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   8160
            Tag             =   "25"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label lblField9 
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
            TabIndex        =   133
            Tag             =   "16"
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Rimorchio"
            Height          =   255
            Index           =   32
            Left            =   360
            TabIndex        =   14
            Top             =   1320
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   32
            Left            =   5520
            Tag             =   "16"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   10440
            Tag             =   "23"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ma&gazzino"
            Height          =   255
            Index           =   18
            Left            =   5880
            TabIndex        =   31
            Top             =   2040
            Width           =   855
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
            Left            =   7440
            TabIndex        =   129
            Tag             =   "23"
            Top             =   2040
            Width           =   2895
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
            Left            =   7440
            TabIndex        =   125
            Tag             =   "21"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ente E&mitt."
            Height          =   195
            Index           =   9
            Left            =   5970
            TabIndex        =   27
            Top             =   1320
            Width           =   765
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   10440
            Tag             =   "21"
            Top             =   1320
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
            Left            =   7440
            TabIndex        =   124
            Tag             =   "3"
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "S&ocietà"
            Height          =   255
            Index           =   8
            Left            =   6120
            TabIndex        =   25
            Top             =   960
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   10440
            Tag             =   "3"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   5520
            Tag             =   "14"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Unità Trasp."
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   975
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
            Left            =   1800
            TabIndex        =   123
            Tag             =   "14"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   5520
            Tag             =   "15"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motri&ce"
            Height          =   195
            Index           =   3
            Left            =   570
            TabIndex        =   12
            Top             =   600
            Width           =   525
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
            TabIndex        =   122
            Tag             =   "15"
            Top             =   600
            Width           =   2655
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2775
         Index           =   2
         Left            =   240
         TabIndex        =   117
         Top             =   960
         Visible         =   0   'False
         Width           =   10815
         Begin VB.Frame Frame5 
            Caption         =   "Inizio Carico"
            Height          =   1095
            Left            =   5040
            TabIndex        =   128
            Top             =   1560
            Width           =   2775
            Begin VB.CommandButton cmdCalendar 
               Height          =   300
               Index           =   1
               Left            =   1920
               Picture         =   "OIL2VVT0.frx":195A
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   360
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskField4 
               Height          =   255
               Left            =   600
               TabIndex        =   54
               Tag             =   "4"
               Top             =   360
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
               Left            =   600
               TabIndex        =   57
               Tag             =   "5"
               Top             =   720
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
               Index           =   6
               Left            =   1320
               Tag             =   "5"
               Top             =   720
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "&Ora"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   56
               Top             =   720
               Width           =   255
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   5
               Left            =   2400
               Tag             =   "4"
               Top             =   360
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "&Data"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Fine Viaggio"
            Height          =   1095
            Left            =   7920
            TabIndex        =   127
            Top             =   1560
            Width           =   2775
            Begin VB.CommandButton cmdCalendar 
               Height          =   300
               Index           =   2
               Left            =   1920
               Picture         =   "OIL2VVT0.frx":1A5C
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   360
               Width           =   375
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   255
               Left            =   600
               TabIndex        =   59
               Tag             =   "8"
               Top             =   360
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
               Left            =   600
               TabIndex        =   62
               Tag             =   "9"
               Top             =   720
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
               Index           =   7
               Left            =   1320
               Tag             =   "9"
               Top             =   720
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "O&ra"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   61
               Top             =   720
               Width           =   255
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   10
               Left            =   2400
               Tag             =   "8"
               Top             =   360
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Da&ta"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   375
            End
         End
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   6480
            TabIndex        =   44
            Tag             =   "35"
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   6480
            TabIndex        =   46
            Tag             =   "36"
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField11 
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Tag             =   "32"
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField14 
            Height          =   255
            Left            =   1560
            TabIndex        =   34
            Tag             =   "13"
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField13 
            Height          =   255
            Left            =   1560
            TabIndex        =   38
            Tag             =   "33"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField12 
            Height          =   255
            Left            =   1560
            TabIndex        =   42
            Tag             =   "10"
            Top             =   1680
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
            Left            =   1560
            TabIndex        =   40
            Tag             =   "34"
            Top             =   1320
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
            Left            =   1560
            TabIndex        =   159
            Tag             =   "45"
            Top             =   2040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox19 
            Height          =   255
            Left            =   6480
            TabIndex        =   48
            Tag             =   "46"
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox25 
            Height          =   255
            Left            =   9720
            TabIndex        =   50
            Tag             =   "49"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox26 
            Height          =   255
            Left            =   9720
            TabIndex        =   52
            Tag             =   "51"
            Top             =   600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox27 
            Height          =   255
            Left            =   1560
            TabIndex        =   168
            Tag             =   "55"
            Top             =   2400
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox28 
            Height          =   255
            Left            =   9720
            TabIndex        =   170
            Tag             =   "57"
            Top             =   960
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   49
            Left            =   10440
            Tag             =   "57"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Km rientro reali"
            Height          =   195
            Index           =   58
            Left            =   8520
            TabIndex        =   171
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sequenza di invio"
            Height          =   195
            Index           =   57
            Left            =   120
            TabIndex        =   169
            Top             =   2400
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Km rientro"
            Height          =   195
            Index           =   56
            Left            =   8640
            TabIndex        =   51
            Top             =   600
            Width           =   1005
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   48
            Left            =   10440
            Tag             =   "51"
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   46
            Left            =   10440
            Tag             =   "25"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Km a vuoto"
            Height          =   195
            Index           =   54
            Left            =   8640
            TabIndex        =   49
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Importo pedaggi"
            Height          =   255
            Index           =   40
            Left            =   5160
            TabIndex        =   47
            Top             =   960
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   45
            Left            =   8280
            Tag             =   "46"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   39
            Left            =   4680
            Tag             =   "45"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Contratto"
            Height          =   255
            Index           =   53
            Left            =   600
            TabIndex        =   161
            Top             =   2040
            Width           =   855
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
            Left            =   2160
            TabIndex        =   160
            Tag             =   "45"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   40
            Left            =   4680
            Tag             =   "34"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Destinaz&ione FV"
            Height          =   255
            Index           =   41
            Left            =   240
            TabIndex        =   39
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblField14 
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
            Left            =   2160
            TabIndex        =   135
            Tag             =   "34"
            Top             =   1320
            Width           =   2415
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
            Left            =   2160
            TabIndex        =   134
            Tag             =   "10"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Zona"
            Height          =   255
            Index           =   17
            Left            =   960
            TabIndex        =   41
            Top             =   1680
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   4680
            Tag             =   "10"
            Top             =   1680
            Width           =   240
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
            Left            =   2160
            TabIndex        =   120
            Tag             =   "32"
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "De&posito FV"
            Height          =   255
            Index           =   25
            Left            =   480
            TabIndex        =   35
            Top             =   600
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   25
            Left            =   4680
            Tag             =   "32"
            Top             =   600
            Width           =   240
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
            Left            =   2160
            TabIndex        =   119
            Tag             =   "13"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo &Vendita"
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   19
            Left            =   4680
            Tag             =   "13"
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
            Left            =   2160
            TabIndex        =   118
            Tag             =   "33"
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ente Emittente FV"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   4680
            Tag             =   "33"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. &Est. 1"
            Height          =   255
            Index           =   1
            Left            =   5640
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   8280
            Tag             =   "35"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   15
            Left            =   8280
            Tag             =   "36"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. E&st. 2"
            Height          =   255
            Index           =   15
            Left            =   5640
            TabIndex        =   45
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2775
         Index           =   3
         Left            =   240
         TabIndex        =   126
         Top             =   960
         Visible         =   0   'False
         Width           =   10815
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Tag             =   "28"
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Tag             =   "26"
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Tag             =   "31"
            Top             =   720
            Width           =   2295
         End
         Begin VB.Frame Frame7 
            Height          =   2055
            Left            =   8880
            TabIndex        =   130
            Top             =   120
            Width           =   1815
            Begin MSMask.MaskEdBox mskField22 
               Height          =   255
               Left            =   840
               TabIndex        =   89
               Tag             =   "37"
               Top             =   240
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField23 
               Height          =   255
               Left            =   840
               TabIndex        =   90
               Tag             =   "38"
               Top             =   600
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField24 
               Height          =   255
               Left            =   840
               TabIndex        =   91
               Tag             =   "39"
               Top             =   960
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField25 
               Height          =   255
               Left            =   840
               TabIndex        =   92
               Tag             =   "40"
               Top             =   1320
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField26 
               Height          =   255
               Left            =   840
               TabIndex        =   93
               Tag             =   "41"
               Top             =   1680
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stato 1"
               Height          =   255
               Index           =   20
               Left            =   120
               TabIndex        =   105
               Top             =   240
               Width           =   615
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   20
               Left            =   1320
               Tag             =   "37"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Festivo"
               Height          =   255
               Index           =   21
               Left            =   120
               TabIndex        =   106
               Top             =   600
               Width           =   615
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   21
               Left            =   1320
               Tag             =   "38"
               Top             =   600
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stato 3"
               Height          =   255
               Index           =   22
               Left            =   120
               TabIndex        =   107
               Top             =   960
               Width           =   615
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   22
               Left            =   1320
               Tag             =   "39"
               Top             =   960
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stato 4"
               Height          =   255
               Index           =   23
               Left            =   120
               TabIndex        =   108
               Top             =   1320
               Width           =   615
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   23
               Left            =   1320
               Tag             =   "40"
               Top             =   1320
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stato 5"
               Height          =   255
               Index           =   24
               Left            =   120
               TabIndex        =   109
               Top             =   1680
               Width           =   615
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   24
               Left            =   1320
               Tag             =   "41"
               Top             =   1680
               Width           =   240
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Tag             =   "43"
            Top             =   240
            Width           =   2295
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            Height          =   255
            Left            =   6120
            TabIndex        =   80
            Tag             =   "24"
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox21 
            Height          =   255
            Left            =   6120
            TabIndex        =   82
            Tag             =   "11"
            Top             =   1560
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox22 
            Height          =   255
            Left            =   6120
            TabIndex        =   84
            Tag             =   "12"
            Top             =   1920
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox23 
            Height          =   255
            Left            =   8880
            TabIndex        =   88
            Tag             =   "29"
            Top             =   2280
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   3
            Format          =   "##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox24 
            Height          =   255
            Left            =   6120
            TabIndex        =   86
            Tag             =   "30"
            Top             =   2280
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   3
            Format          =   "##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   1050
            TabIndex        =   72
            Tag             =   "17"
            Top             =   1920
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
            Left            =   1050
            TabIndex        =   74
            Tag             =   "18"
            Top             =   2280
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
            Left            =   1050
            TabIndex        =   68
            Tag             =   "19"
            Top             =   1200
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
            Left            =   1050
            TabIndex        =   70
            Tag             =   "20"
            Top             =   1560
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
            Index           =   2
            Left            =   4170
            Tag             =   "17"
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vett&ore"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   71
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblField10 
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
            Left            =   1650
            TabIndex        =   145
            Tag             =   "17"
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   4170
            Tag             =   "18"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vett. &Trasp."
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   73
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label lblField11 
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
            Left            =   1650
            TabIndex        =   144
            Tag             =   "18"
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   33
            Left            =   4170
            Tag             =   "19"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au&tista 1"
            Height          =   255
            Index           =   33
            Left            =   360
            TabIndex        =   67
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblField12 
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
            Left            =   1650
            TabIndex        =   143
            Tag             =   "19"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   34
            Left            =   4170
            Tag             =   "20"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Aut&ista 2"
            Height          =   255
            Index           =   34
            Left            =   360
            TabIndex        =   69
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label lblField13 
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
            Left            =   1650
            TabIndex        =   142
            Tag             =   "20"
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "T&empo carico"
            Height          =   255
            Index           =   45
            Left            =   5040
            TabIndex        =   85
            Top             =   2280
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   44
            Left            =   7080
            Tag             =   "30"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Te&mpo viaggio"
            Height          =   255
            Index           =   44
            Left            =   7680
            TabIndex        =   87
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   43
            Left            =   9840
            Tag             =   "29"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Badge A&utista"
            Height          =   255
            Index           =   43
            Left            =   4920
            TabIndex        =   83
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   42
            Left            =   7080
            Tag             =   "12"
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Badge Automezzo"
            Height          =   255
            Index           =   42
            Left            =   4680
            TabIndex        =   81
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   41
            Left            =   7080
            Tag             =   "11"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Calcolo Km def."
            Height          =   255
            Index           =   38
            Left            =   4800
            TabIndex        =   77
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   38
            Left            =   7800
            Tag             =   "28"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Calcolo &Km prev."
            Height          =   255
            Index           =   37
            Left            =   4800
            TabIndex        =   75
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   37
            Left            =   7800
            Tag             =   "26"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rientro/&Dirottamento"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   31
            Left            =   4200
            Tag             =   "31"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Viaggio Rif."
            Height          =   255
            Index           =   29
            Left            =   5040
            TabIndex        =   79
            Top             =   1200
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   29
            Left            =   7800
            Tag             =   "24"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   26
            Left            =   4200
            Tag             =   "43"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Tipo Viaggio"
            Height          =   255
            Index           =   26
            Left            =   720
            TabIndex        =   63
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2775
         Index           =   4
         Left            =   240
         TabIndex        =   164
         Top             =   960
         Width           =   10815
         Begin TrueDBGrid60.TDBGrid TDBGrid3 
            Height          =   2415
            Left            =   120
            OleObjectBlob   =   "OIL2VVT0.frx":1B5E
            TabIndex        =   165
            Top             =   240
            Width           =   10575
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "42"
         Top             =   240
         Width           =   2055
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5741
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Viaggio &1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Viaggio &2"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Viaggio &3"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Carico"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   10080
         TabIndex        =   8
         Tag             =   "44"
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Tag             =   "0"
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   8640
         TabIndex        =   7
         Tag             =   "1"
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "OIL2VVT0.frx":7888
         TabIndex        =   94
         Tag             =   "DETAIL"
         Top             =   3960
         Width           =   4335
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   975
         Left            =   4560
         OleObjectBlob   =   "OIL2VVT0.frx":C6DA
         TabIndex        =   95
         Tag             =   "DETAIL1"
         Top             =   3960
         Width           =   6615
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Tag             =   "2"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   2
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9720
         Top             =   600
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
               Picture         =   "OIL2VVT0.frx":12114
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2VVT0.frx":12226
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2VVT0.frx":12338
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tot. Ordinato Viaggio Lt"
         Height          =   255
         Index           =   51
         Left            =   4560
         TabIndex        =   163
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lblTotTrasportato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6480
         TabIndex        =   162
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblTotConsegnato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10080
         TabIndex        =   141
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tot. consegnato Viaggio Lt"
         Height          =   255
         Index           =   13
         Left            =   8040
         TabIndex        =   140
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Stato"
         Height          =   255
         Index           =   27
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   7200
         Tag             =   "42"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N. Se&quenza"
         Height          =   255
         Index           =   30
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   30
         Left            =   4080
         Tag             =   "2"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero de&f."
         Height          =   255
         Index           =   28
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   9600
         Tag             =   "1"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2040
         Tag             =   "0"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Numero"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Generación y Cancelación de eventos"
      Height          =   1215
      Left            =   0
      TabIndex        =   173
      Top             =   6720
      Width           =   11295
      Begin VB.ComboBox cmb_cges 
         Height          =   315
         Left            =   4560
         TabIndex        =   176
         Text            =   "Selecionar Tipo Evento:"
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton btn_generaevento 
         Caption         =   "Genera Evento"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         Picture         =   "OIL2VVT0.frx":1244A
         TabIndex        =   175
         Top             =   600
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txt_dataev 
         Height          =   255
         Left            =   9120
         TabIndex        =   179
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_hrev 
         Height          =   255
         Left            =   9120
         TabIndex        =   180
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora evento"
         Height          =   255
         Index           =   61
         Left            =   7920
         TabIndex        =   178
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data evento"
         Height          =   255
         Index           =   60
         Left            =   7920
         TabIndex        =   177
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generar evento:"
         Height          =   255
         Index           =   59
         Left            =   2640
         TabIndex        =   174
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "OIL2VVT0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 5
Private Const MAX_COLS_1 = 7 '8
Public GridArrayDetail As ArrayGridCls
Public GridArrayDetail1 As ArrayGridCls
Private XALoad As XArrayDB

Public WithEvents DataClass As CViaggio
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CViaggioConsegna
Public DetailClass1 As CViaggioProdotto
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private mstrIDImballoDefault As String
Private mstrIDPosDefault As String

Private mintCurFrame As Integer
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object

Private APR As CProdotto
Private OCT As COrdine
Private AAM As CAutomezzo
Private AUT As CUnitaDiTrasporto
'SM Generación de eventos GES en tabla VIAGES0
Private DA As CDataAccess


Public Sub ClearDetail1()
    ClearUnboundGrid TDBGrid2, GridArrayDetail1, MAX_COLS_1
End Sub


Public Function DeleteDetail()
    On Error GoTo ErrHandler
    
    If Not DetailClass.Delete Then
        GoTo ErrHandler
    End If
    DisplayDetail
    Exit Function
ErrHandler:

End Function

Public Sub DisplayDetail1()
    Dim strFieldKey As String
    Dim strFieldKey1 As String
'    Dim varTotConsegna As Variant
    Dim objDS As CDataStruct
    
    Set objDS = DetailClass.DS
    strFieldKey = objDS.KeyFields("1")
    strFieldKey1 = objDS.KeyFields("2")
    
    With DetailClass1.DS
        .Fields("0").FindValue = objDS.Fields(strFieldKey).value
        .Fields("1").FindValue = objDS.Fields(strFieldKey1).value
        DetailClass1.GetList
        RefreshUnboundGrid TDBGrid2, .List, GridArrayDetail1, _
                           .HowMany, MAX_COLS_1
    End With
    
'    With DetailClass
'        .GetTotali
'        varTotConsegna = .TotaliQta(5)
'    End With
'
'    If Not IsNull(varTotConsegna) Then
'        lblTotConsegna = varTotConsegna
'    End If
End Sub



Private Sub DisplayLoadData()
    On Error GoTo ErrHandler
    
    Dim varLoadList As Variant
    Dim VCA As CViaggioCarico
    
    ' --- Visualizziamo i dati di carico
    Set VCA = New CViaggioCarico
    With VCA
        Set .DS.DA = DataClass.DS.DA
    
        varLoadList = .GetLoadData(DataClass.DS.Fields("0").value)
    End With
    
    If IsEmpty(varLoadList) Then
        GoTo ErrHandler
    End If
    RefreshXArray XALoad, TDBGrid3, varLoadList
    ' ---
    
    Set VCA = Nothing
    Exit Sub

ErrHandler:
    Set VCA = Nothing
End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
'    If IsNull(DBGrid.Bookmark) Then
'        GetSelectedID = vbNullString
'    Else
'        GetSelectedID = GridArrayDetail.value(DBGrid.Bookmark, 2)
'    End If
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = GridArrayDetail1.value(DBGrid.bookmark, 5)
    End If
End Function

Private Sub DisplayDetail()
    With DetailClass
        .DS.Fields.ResetFields
        .DS.Fields("0").FindValue = DataClass.DS.Fields(DataClass.DS.KeyFields("1")).value
        .GetListProgConsegna
        RefreshUnboundGrid TDBGrid1, .DS.List, GridArrayDetail, _
                           .DS.HowMany, MAX_COLS
    End With
    DataClass.GetTotali
    lblTotConsegnato = DataClass.TotaliQta(5)
    lblTotTrasportato = DataClass.TotaliQta(1)
End Sub

Public Function InsertDetail() As Boolean
    On Error GoTo ErrHandler
    
    #If VersionID = 1 Then
        ' --- Imposto la data di consegna uguale alla data del viaggio
        DetailClass.DS.Fields("41").NewValue = DataClass.DS.Fields("6").value
        DetailClass.DS.Fields("41").Updated = True
        ' ---
    #End If
    
    If Not DetailClass.Insert Then
        GoTo ErrHandler
    End If
    With DetailClass.DS.Fields
        If Not DetailClass1.OrdineToProdotti(.Item("37").value, _
                                             .Item("0").value, _
                                             .Item("1").value) Then
            GoTo ErrHandler
        End If
    End With
    InsertDetail = True
    Exit Function
ErrHandler:
    InsertDetail = False
End Function

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

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
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strSelectionFormula As String
    Dim Respok, Respuesta As String
'    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    ReDim varReportsList(0 To 4, 0 To 0) As Variant
     Dim Service As VLocalize.VLocalize
    
    Select Case Cmd
        Case fcSave
                ' Control RS3
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            'Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * 3
            Respuesta = InputBox("Esta opción solo esta disponible desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
                Select Case State
                    Case fsEdit
                        blnResult = UpdateRecord
                    Case Else
                        blnResult = False
                End Select
                If blnResult Then
                    SetState fsEdit
                End If
            End If
        Case fcUndo
            If (State = fsEdit) Then
                TDBGrid1.CurrentCellModified = False
                TDBGrid1.EditActive = False
                SetState fsEdit
            End If
        Case fcPrint
            If (State = fsEdit) Then
                varReportsList(0, 0) = "Hoja de Ruta"
                varReportsList(1, 0) = "oil2vvt1_gc.rpt"
                varReportsList(2, 0) = 1
                varReportsList(3, 0) = 0
                varReportsList(4, 0) = 0
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcORDIOCT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
'                    .GetRecord GetSelectedID(TDBGrid1)
                    .GetRecord GetSelectedID(TDBGrid2)
                    .SetState fsEdit
                    .Command1.Enabled = False
                    Me.Enabled = False
                End With
            End If
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcVIAGVVC)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord DetailClass.DS.Fields(DetailClass.DS.KeyFields(1)).value, _
                               DetailClass.DS.Fields(DetailClass.DS.KeyFields(2)).value
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
            DisplayTotViaggio
        Case fcLinkRefresh
            DisplayDetail
            DisplayTotViaggio
        Case fcCustom1
            DeleteDetail
            DisplayTotViaggio
        Case fcFind
            Set LinkForm = DisplayForm(fcVIAGVVP)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord DetailClass.DS.Fields(DetailClass.DS.KeyFields(1)).value, _
                               DetailClass.DS.Fields(DetailClass.DS.KeyFields(2)).value, _
                               DetailClass1.DS.Fields(DetailClass1.DS.KeyFields(3)).value
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcdellserv
         Dim COD As String
         Dim Dia As Integer
         Dim Mes As Integer
         Dim Año As Integer
         Dim Hora As Integer
         ' NO DEBUGADO
         Dia = Day(Now)
         Mes = Month(Now)
         Año = Year(Now)
         Hora = Hour(Now)
         
         COD = InputBox("Se va a Cancelar el servicio de la Palm, Esta Seguro de querer continuar con la operación? Si es así introduce el codigo de desbloqueo.", "Eliminar Viaje")
            If ((COD) = mskField1.Text) Then
               
              
          Dim KVE As CSatellitareEvento
    
    
            Set KVE = New CSatellitareEvento
            With KVE
                Set .DS.DA = DataConnect.DA
                Set .DS.ErrH = DataConnect.DA.ErrH
                            If Not KVE.EliminaServicioPalm(mskField1.Text, DataConnect) Then
                                
                            End If
            End With
            Else
              MsgBox ("Codigo Incorrecto")
            End If
                
        Case fcExit
            Unload Me
'            Set OIL2VVT0 = Nothing
    End Select
End Sub

Public Function GetRecord(strID As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    
                    DisplayDetail
                    DisplayScomparti mskField6.Text, lblScompartiM, lblPortataM
                    DisplayScomparti MaskEdBox7.Text, lblScompartiR, lblPortataR
                    DisplayTotViaggio
                    DisplayLoadData
'                    If (Trim$(.Fields("42").value) <> "P") Then 'Or (.Fields("41").value = "T") Then
                    If (Trim$(.Fields("42").value) = "E") Or _
                       ((.Fields("41").value <> "T") And (Trim$(.Fields("42").value) = "D")) Then
                        With TDBGrid1
                            .AllowAddNew = False
                            .AllowDelete = False
                            .AllowUpdate = False
                        End With
                        With TDBGrid2
                            .AllowAddNew = False
                            .AllowDelete = False
                            .AllowUpdate = False
                        End With
                        Command8.Enabled = False
                    End If
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





Private Function UpdateRecord() As Boolean
    On Error GoTo ErrHandler
    
    Dim blnChkValue As Boolean
    Dim blnConcurError As Boolean
    Dim strMsg As String
    Dim C As Control
    Dim CErr As CError
    
    'Scansione della finestra
    With DataClass.DS
        If Not .ScanClient(0) Then
            GoTo ErrHandler
        End If
        
        If .Fields.Validate Then
            If Not DataClass.CheckScadenze(strMsg) Then
                MsgBox strMsg, vbCritical
                GoTo ErrHandler
            End If
            
            '-- DP 05/08/08
            If Not DataClass.CheckSequenza(MaskEdBox16.Text, _
                                            MaskEdBox14.Text, _
                                            mskField7.Text, _
                                            MaskEdBox4.Text, _
                                            mskField1.Text) Then
                MsgBox "Il numero di sequenza è già stato utilizzato", vbCritical, "ResOil.2"
                GoTo ErrHandler
            End If
            
            If DataClass.BeginSave(blnConcurError) Then
                If blnConcurError Then
                    If MsgBox(MSG_CONCUR_ERR, _
                               vbOKCancel + vbExclamation + vbDefaultButton2, _
                              "Attenzione") = vbOK Then
                        GoTo EndProc
                    Else
                        .DA.EndTrans ETROLLBACK
                        GetRecord .Fields(.KeyFields(1)).value
                        GoTo ErrHandler
                    End If
                Else
                    GoTo EndProc
                End If
                GoTo EndProc
            Else
                .DA.EndTrans ETROLLBACK
                strMsg = "Il viaggio non può venire modificato."
                MsgBox strMsg, vbCritical + vbOKOnly, "IMPOSSIBILE SALVARE"
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
        .DA.EndTrans ETCOMMIT
        .EndSave
        UpdateRecord = True
    End With
End Function

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsEdit
            TabStrip1.Enabled = True
            TabStrip1.Tabs(1).Selected = True
'            mskField7.SetFocus
            MaskEdBox14.SetFocus
    End Select
End Sub



Private Sub btn_generaevento_Click()

Set DA = DataConnect.DA
Dim CodEv As String
Dim COD As String
Dim SQL_GEN_EV As String
Dim SQL_GFIND_NPALM As String
Dim Conductor As String
Dim NPDA As Integer
Dim HREV As String
Dim DYEV As String
Dim NPalm As Variant
Dim lng As Long
  
  CodEv = cmb_cges.ItemData(cmb_cges.ListIndex)
  COD = mskField1.Text
  HREV = txt_hrev.Text
  DYEV = txt_dataev.Text
  Conductor = MaskEdBox12.Text
  
    Do While (Len(CodEv) < 3)
     CodEv = "0" + CodEv
  Loop

  
  If (Len(CodEv) = 3) Then
    If (Len(txt_hrev.Text) > 0 And Len(txt_dataev.Text) > 0 And txt_hrev.Text <> "__:__" And txt_dataev <> "__/__/__") Then
       SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + COD + "','" + CodEv + "','" + DYEV + " " + HREV + "','" + DYEV + " " + HREV + "','0')"
       If DA.ExecuteQuery(SQL_GEN_EV) Then
           MsgBox ("Evento '" + CodEv + "' generado con exito para el conductor : " + Conductor + "  El: " + CStr(DYEV) + " A las :" + CStr(HREV))
       End If
    Else
       MsgBox ("Hay que introducir una fecha y una hora para crear el evento")
    End If
  Else
       MsgBox ("Numero de evento no indicado.")
  End If
End Sub

Private Sub cmb_cges_Change()
btn_generaevento.Enabled = True
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox16
            strTitle = "Data inizio"
        Case 1
            Set M = mskField4
            strTitle = "Data inizio carico"
        Case 2
            Set M = MaskEdBox1
            strTitle = "Data fine viaggio"
        Case 3
            Set M = MaskEdBox29
            strTitle = "Fecha Referencia viaje"

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
    ReDim ColTags(1 To MAX_COLS) As Variant
    ReDim ColTags1(1 To MAX_COLS_1) As Variant
    
    Loading = True
    
    Set DataClass = New CViaggio
    Set DetailClass = New CViaggioConsegna
    Set DetailClass1 = New CViaggioProdotto
    Set ErrH = New CErrorHandler
    Set GridArrayDetail = New ArrayGridCls
    Set GridArrayDetail1 = New ArrayGridCls
    
'SM Combo de generación de Eventos GES
'No se porque al añadir items es necesario devolver un booleano.
Dim nose As Boolean
     
     
    
    
   cmb_cges.AddItem "002-Entrada en Factoria"
   cmb_cges.ItemData(0) = "002"
   cmb_cges.AddItem "003-Salida de Factoria"
   cmb_cges.ItemData(1) = "003"
    


    
    
    mintCurFrame = 1
    ColTags(1) = "1"
    ColTags(2) = "2"
    ColTags(3) = "37"
    ColTags(4) = ""
    ColTags(5) = "38"
    
    ColTags1(1) = "2"
    ColTags1(2) = "3"
    ColTags1(3) = "7" '"11"
    ColTags1(4) = "11" '"7"
    ColTags1(5) = "19" ' "14"
    ColTags1(6) = "32" ' "15"
    ColTags1(7) = "35"
'    ColTags1(8) = "35"
'    ColTags1(4) = "10"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags, , ColTags1
        Set .ErrH = ErrH
    End With
    
    With DetailClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    With DetailClass1.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        With .MultiSoc.DS
            Set .DA = DataConnect.DA
            Set .ErrH = ErrH
        End With
    End With
    
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    SetValueItems TDBGrid2.Columns(1), APR
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set APR = Nothing
    Set OCT = Nothing
    Set AAM = Nothing
    Set AUT = Nothing
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set DetailClass = Nothing
    Set DetailClass1 = Nothing
    Set GridArrayDetail = Nothing
    Set GridArrayDetail1 = Nothing
    Set XALoad = Nothing
    
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set OIL2VVT0 = Nothing
End Sub







Private Sub MaskEdBox7_LostFocus()
    DisplayScomparti MaskEdBox7.Text, lblScompartiR, lblPortataR
    DisplayTotViaggio
End Sub


Private Sub DisplayScomparti(IDAutomezzo As String, L As Label, LTot As Label)
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim strScomparti As String
    Dim dblPortata As Double
    Dim dblPortataM As Double
    Dim dblPortataR As Double
    
    L = vbNullString
    LTot = vbNullString
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
    LTot = dblPortata
    
    If Len(lblPortataM) > 0 Then
        dblPortataM = CDbl(lblPortataM)
    End If
    If Len(lblPortataR) > 0 Then
        dblPortataR = CDbl(lblPortataR)
    End If
    lblPortataU = dblPortataM + dblPortataR
    If lblPortataU > 0 Then
        pgbPortata.Max = lblPortataU
    End If
    Exit Sub

ErrHandler:
    L = vbNullString
End Sub

Private Sub DisplayTotViaggio()
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim varTotConsegna As Variant
    Dim strIDOrdine As String
    Dim intRows As Integer
    
    Me.MousePointer = vbHourglass
    intRows = GridArrayDetail.RowCount - 1
    strIDOrdine = GridArrayDetail.value(i, 2)
    If Len(strIDOrdine) > 0 Then
        With DataClass
            .GetTotali
            varTotConsegna = .TotaliQta(1)
        End With
    End If
    lblTotTrasportato = varTotConsegna
    If Len(lblPortataU) = 0 Then
        lblPortataU = 0
    End If
    If Len(lblTotTrasportato) = 0 Then
        lblTotTrasportato = 0
    End If
    lblTotVuoto = Me.lblPortataU - lblTotTrasportato
    If lblTotVuoto < 0 Then
        lblTotVuoto.ForeColor = scRed
    Else
        lblTotVuoto.ForeColor = scBlue
    End If
    If lblTotTrasportato > pgbPortata.Max Then
        pgbPortata.value = pgbPortata.Max
    Else
        pgbPortata.value = lblTotTrasportato
    End If
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
End Sub

Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub


Private Sub mskField6_LostFocus()
    DisplayScomparti mskField6.Text, lblScompartiM, lblPortataM
    DisplayTotViaggio
End Sub


Private Sub mskField7_LostFocus()
    If mskField7 <> vbNullString Then
        With AUT.DS
            If mskField7 <> DataClass.DS.Field("unità di trasporto").value Then
                .Fields("0").FindValue = mskField7
                If AUT.Find Then
                    
    '                If Len(mskField6.Text) = 0 Then
                        mskField6.Text = .Fields("2").value
                        DataClass.DS.ValidateClientField "15", mskField6.Text
                        FieldUpdateDescription "15", Me, mskField6
                        DisplayScomparti mskField6.Text, lblScompartiM, lblPortataM
    '                End If
                    
    '                If Len(MaskEdBox7.Text) = 0 Then
                        MaskEdBox7.Text = .Fields("3").value
                        DataClass.DS.ValidateClientField "16", MaskEdBox7.Text
                        FieldUpdateDescription "16", Me, MaskEdBox7
                        DisplayScomparti MaskEdBox7.Text, lblScompartiR, lblPortataR
    '                End If
                    
                    DisplayTotViaggio
    
                    
    '                If Len(MaskEdBox12.Text) = 0 Then
                        MaskEdBox12.Text = .Fields("4").value
                        DataClass.DS.ValidateClientField "19", MaskEdBox12.Text
                        FieldUpdateDescription "19", Me, MaskEdBox12
    '                End If
                    
    '                If Len(MaskEdBox15.Text) = 0 Then
                        MaskEdBox15.Text = .Fields("5").value
                        DataClass.DS.ValidateClientField "20", MaskEdBox15.Text
                        FieldUpdateDescription "20", Me, MaskEdBox15
    '                End If
                    
    '                If Len(MaskEdBox18.Text) = 0 Then
                        MaskEdBox18.Text = .Fields("15").value
                        DataClass.DS.ValidateClientField "45", MaskEdBox18.Text
                        FieldUpdateDescription "45", Me, MaskEdBox18
    '                End If
                        
                        MaskEdBox5.Text = .Fields("6").value
                        DataClass.DS.ValidateClientField "17", MaskEdBox5.Text
                        FieldUpdateDescription "17", Me, MaskEdBox5
                        
                        MaskEdBox11.Text = .Fields("7").value
                        DataClass.DS.ValidateClientField "18", MaskEdBox11.Text
                        FieldUpdateDescription "18", Me, MaskEdBox11
                End If
            End If
        End With
    End If
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame6(TabStrip1.SelectedItem.Index).Visible = True
        Frame6(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index

End Sub



Private Sub TDBGrid1_AfterDelete()
    DisplayTotViaggio
End Sub

Private Sub TDBGrid1_AfterInsert()
    DataClass.GetTotali
    lblTotConsegnato = DataClass.TotaliQta(5)
    lblTotTrasportato = DataClass.TotaliQta(1)
    DisplayTotViaggio
End Sub

Private Sub TDBgrid1_AfterUpdate()
    DisplayTotViaggio
End Sub


Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Dim blnDataOk As Boolean
    
    blnDataOk = True
    Select Case ColIndex
        Case 0, 2
            If TDBGrid1.AddNewMode = dbgNoAddNew Then
                blnDataOk = False
            End If
    End Select
    Cancel = Not blnDataOk
End Sub

Private Sub TDBgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim blnDataOk As Boolean
    Dim i As Integer
    
    blnDataOk = True
    
    With TDBGrid1.Columns(ColIndex)
        Select Case ColIndex
            Case 0
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                Else
                    .Text = Format$(.Text, "000")
                    For i = 0 To GridArrayDetail.RowCount - 1
                        If .Text = GridArrayDetail.value(i, ColIndex) Then
                            blnDataOk = False
                            Exit For
                        End If
                    Next i
                End If
            Case 1
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                Else
                    .Text = Format$(.Text, "000")
                    For i = 0 To GridArrayDetail.RowCount - 1
                        If .Text = GridArrayDetail.value(i, ColIndex) Then
                            blnDataOk = False
                            Exit For
                        End If
                    Next i
                End If
            Case 2
                For i = 0 To GridArrayDetail.RowCount - 1
                    If .Text = GridArrayDetail.value(i, 2) Then
                        blnDataOk = False
                        GoTo EndProc
                    End If
                Next i
                OCT.DS.Fields("0").FindValue = .Text
                With OCT.DS
                    ' --- Controllo che il codice dell'ordine sia valido
                    If .HowMany = 0 Then
                        blnDataOk = False
                        GoTo EndProc
                    End If
                    ' ---
                    ' --- Controllo che l'ordine sia attivo
'                    If .Fields("22").value <> "A" Then
'                        blnDataOk = False
'                        GoTo EndProc
'                    End If
                    ' ---
                    ' --- Controllo lo stato di avanzamento dell'ordine
                    If (.Fields("23").value <> "A") And (.Fields("23").value <> "M") Then
                        blnDataOk = False
                        GoTo EndProc
                    End If
                    ' ---
                    ' --- Controllo che l'ordine non sia stato annullato
                    If .Fields("32").value <> False Then
                        blnDataOk = False
                        GoTo EndProc
                    End If
                    ' ---
                End With
        End Select
    End With
EndProc:
    Cancel = Not blnDataOk
End Sub

Private Sub TDBGrid1_BeforeUpdate(Cancel As Integer)
    Dim blnDataOk As Boolean
    
    blnDataOk = True
    With TDBGrid1.Columns
        If Not IsNumeric(.Item(0).Text) Then
            blnDataOk = False
        End If
        If Not IsNumeric(.Item(1).Text) Then
            blnDataOk = False
        End If
        If .Item(2).Text = vbNullString Then
            blnDataOk = False
        End If
    End With
    Cancel = Not blnDataOk
End Sub


Private Sub TDBGrid1_Change()
    With TDBGrid1
        If .Col = 2 Then
            If .Text <> vbNullString Then
                With OCT.DS
                    If Len(TDBGrid1.Columns(2).Text) = 7 Then
                        .Fields("0").FindValue = TDBGrid1.Columns(2).Text
                        OCT.Find
                        If .HowMany = 0 Then
                            GoTo EndProc
                        End If
                        If .Fields("32").value = True Then
                            GoTo EndProc
                        End If
                        If (.Fields("23").value = "V") Or (.Fields("23").value = "F") Then 'Or (.Fields("22").value = "S") Then
                            GoTo EndProc
                        End If
                        If Not DetailClass.OrdineToConsegna(.Fields("0").value) Then
                            GoTo EndProc
                        End If
                        TDBGrid1.Columns(3).Text = DetailClass.DS.Fields("13").NewValue
                        Exit Sub
                    End If
                End With
                GoTo EndProc
            Else
                GoTo EndProc
            End If
        End If
    End With
    Exit Sub
EndProc:
    TDBGrid1.Columns(3).Text = vbNullString
    ClearDetail1
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    If TDBGrid1.Col = 2 Then
        If Chr(KeyAscii) = "?" Then
            KeyAscii = 0
            TDBGrid1.Text = vbNullString
            OCT.DS.Fields("23").FindValue = "A"
            OCT.DS.Fields("23").FindValue(2) = "M"
'            OCT.DS.Fields("22").FindValue = "A"
            OCT.DS.Fields("32").FindValue = 0
            
            ShowSearchForm Me, "", OCT, TDBGrid1.TabIndex, True
            
            Dim C As Control
            Dim i As Integer
            For i = 0 To Controls.Count - 1
                If Controls(i).Name = "TDBGrid1" Then
                    frmSearch.LinkControlIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    DoEvents
End Sub

Private Sub TDBGrid2_BeforeInsert(Cancel As Integer)
    Cancel = True
End Sub

