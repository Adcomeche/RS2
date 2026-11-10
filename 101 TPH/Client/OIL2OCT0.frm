VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2OCT0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDINI - VISUALIZZAZIONE E MODIFICA "
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   87
      Top             =   5640
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   7300
         Picture         =   "OIL2OCT0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   149
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
         Left            =   6960
         Picture         =   "OIL2OCT0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   6840
         Picture         =   "OIL2OCT0.frx":0204
         TabIndex        =   84
         Tag             =   "EXIT"
         Top             =   720
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   4680
         Picture         =   "OIL2OCT0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F8 - Agg. Dettaglio"
         Height          =   375
         Left            =   4560
         Picture         =   "OIL2OCT0.frx":0408
         TabIndex        =   110
         Tag             =   "REFRESHLIST"
         Top             =   720
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   8
         Left            =   5160
         Picture         =   "OIL2OCT0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   140
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
         Index           =   5
         Left            =   2400
         Picture         =   "OIL2OCT0.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Riga"
         Height          =   375
         Left            =   2280
         Picture         =   "OIL2OCT0.frx":070E
         TabIndex        =   108
         Tag             =   "EDITLINK"
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2OCT0.frx":0810
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F6 - Apri Viaggio"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2OCT0.frx":0912
         TabIndex        =   112
         Tag             =   "OPEN"
         Top             =   720
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
         Left            =   3480
         Picture         =   "OIL2OCT0.frx":0A14
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   89
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
         Index           =   4
         Left            =   1800
         Picture         =   "OIL2OCT0.frx":0B16
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   88
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
         Picture         =   "OIL2OCT0.frx":0C18
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Annulla Riga"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         Picture         =   "OIL2OCT0.frx":0D1A
         TabIndex        =   148
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F9 - Copia Ordine"
         Height          =   375
         Left            =   5040
         Picture         =   "OIL2OCT0.frx":0E1C
         TabIndex        =   141
         Tag             =   "FIND"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   3360
         Picture         =   "OIL2OCT0.frx":0F1E
         TabIndex        =   83
         Tag             =   "PRINT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1680
         Picture         =   "OIL2OCT0.frx":1020
         TabIndex        =   82
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2OCT0.frx":1122
         TabIndex        =   81
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   86
      Top             =   6900
      Width           =   9375
      _ExtentX        =   16536
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
            Object.Width           =   11086
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
      Height          =   5595
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame6 
         Height          =   3075
         Index           =   1
         Left            =   240
         TabIndex        =   96
         Top             =   1080
         Width           =   8895
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   5
            Left            =   8040
            Picture         =   "OIL2OCT0.frx":1224
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2760
            Width           =   375
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   4
            Left            =   8040
            Picture         =   "OIL2OCT0.frx":1326
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   1
            Left            =   8040
            Picture         =   "OIL2OCT0.frx":1428
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   2040
            Width           =   375
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   0
            Left            =   8040
            Picture         =   "OIL2OCT0.frx":152A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1680
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskField9 
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Tag             =   "1"
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
            Left            =   5400
            TabIndex        =   13
            Tag             =   "2"
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
            Left            =   1560
            TabIndex        =   9
            Tag             =   "12"
            Top             =   600
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
            Left            =   1560
            TabIndex        =   7
            Tag             =   "13"
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField3 
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Tag             =   "14"
            Top             =   1680
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   960
            TabIndex        =   21
            Tag             =   "15"
            Top             =   2040
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   6720
            TabIndex        =   23
            Tag             =   "7"
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
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   255
            Left            =   6720
            TabIndex        =   26
            Tag             =   "8"
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField15 
            Height          =   255
            Left            =   960
            TabIndex        =   15
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
         Begin MSMask.MaskEdBox mskField12 
            Height          =   255
            Left            =   5400
            TabIndex        =   17
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
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   6720
            TabIndex        =   30
            Tag             =   "9"
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
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Tag             =   "36"
            Top             =   2400
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   6720
            TabIndex        =   33
            Tag             =   "52"
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
         Begin MSMask.MaskEdBox MaskEdBox21 
            Height          =   255
            Left            =   960
            TabIndex        =   32
            Tag             =   "51"
            Top             =   2760
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Barco"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   144
            Top             =   2760
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   42
            Left            =   5400
            Tag             =   "51"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Grado"
            Height          =   255
            Index           =   41
            Left            =   5640
            TabIndex        =   143
            Top             =   2760
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   41
            Left            =   8520
            Tag             =   "52"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   32
            Left            =   5400
            Tag             =   "36"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nota 3"
            Height          =   255
            Index           =   32
            Left            =   360
            TabIndex        =   113
            Top             =   2400
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   8520
            Tag             =   "9"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data E&ff."
            Height          =   255
            Index           =   13
            Left            =   5880
            TabIndex        =   29
            Top             =   2400
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   8520
            Tag             =   "18"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Z&ona"
            Height          =   255
            Index           =   17
            Left            =   4800
            TabIndex        =   16
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
            Left            =   6000
            TabIndex        =   105
            Tag             =   "18"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   4080
            Tag             =   "17"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Deposito"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   14
            Top             =   1320
            Width           =   615
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
            Left            =   1560
            TabIndex        =   104
            Tag             =   "17"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data &Rich."
            Height          =   255
            Index           =   11
            Left            =   5760
            TabIndex        =   22
            Top             =   1680
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   8520
            Tag             =   "7"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data &Pianif."
            Height          =   255
            Index           =   12
            Left            =   5760
            TabIndex        =   25
            Top             =   2040
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   8520
            Tag             =   "8"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   5400
            Tag             =   "14"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "No&ta 1"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   18
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "N&ota 2"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   20
            Top             =   2040
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   5400
            Tag             =   "15"
            Top             =   2040
            Width           =   240
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
            Left            =   6000
            TabIndex        =   100
            Tag             =   "2"
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ente Emitt."
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   12
            Top             =   960
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   8520
            Tag             =   "2"
            Top             =   960
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
            Left            =   1560
            TabIndex        =   99
            Tag             =   "1"
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Società"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   4080
            Tag             =   "1"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   8520
            Tag             =   "13"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dest&inazione"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Width           =   1095
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
            Left            =   3480
            TabIndex        =   98
            Tag             =   "13"
            Top             =   240
            Width           =   4935
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   8520
            Tag             =   "12"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Cliente"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   8
            Top             =   600
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
            Left            =   3480
            TabIndex        =   97
            Tag             =   "12"
            Top             =   600
            Width           =   4935
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3075
         Index           =   2
         Left            =   240
         TabIndex        =   92
         Top             =   1080
         Width           =   8895
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Tag             =   "56"
            Top             =   2640
            Width           =   2055
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "55"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Frame Frame5 
            Caption         =   "Inserimento"
            Height          =   1095
            Left            =   5760
            TabIndex        =   103
            Top             =   240
            Width           =   2895
            Begin VB.CommandButton cmdCalendar 
               Height          =   300
               Index           =   2
               Left            =   1920
               Picture         =   "OIL2OCT0.frx":162C
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   360
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskField4 
               Height          =   255
               Left            =   600
               TabIndex        =   49
               Tag             =   "5"
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
               TabIndex        =   52
               Tag             =   "6"
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
               Tag             =   "6"
               Top             =   720
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "&Ora"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   51
               Top             =   720
               Width           =   255
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   5
               Left            =   2400
               Tag             =   "5"
               Top             =   360
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "&Data"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   48
               Top             =   360
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ricezione"
            Height          =   1095
            Left            =   5760
            TabIndex        =   102
            Top             =   1440
            Width           =   2895
            Begin VB.CommandButton cmdCalendar 
               Height          =   300
               Index           =   3
               Left            =   1920
               Picture         =   "OIL2OCT0.frx":172E
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   360
               Width           =   375
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   255
               Left            =   600
               TabIndex        =   54
               Tag             =   "3"
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
               TabIndex        =   57
               Tag             =   "4"
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
               Tag             =   "4"
               Top             =   720
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "O&ra"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   56
               Top             =   720
               Width           =   255
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   10
               Left            =   2400
               Tag             =   "3"
               Top             =   360
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Da&ta"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   375
            End
         End
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   1920
            TabIndex        =   42
            Tag             =   "10"
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   1920
            TabIndex        =   44
            Tag             =   "11"
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField11 
            Height          =   255
            Left            =   1920
            TabIndex        =   38
            Tag             =   "20"
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
            Left            =   1920
            TabIndex        =   36
            Tag             =   "16"
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
            Left            =   1920
            TabIndex        =   40
            Tag             =   "19"
            Top             =   960
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
            Left            =   1920
            TabIndex        =   45
            Tag             =   "46"
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   44
            Left            =   7920
            Tag             =   "56"
            Top             =   2640
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Venta En Ruta"
            Height          =   255
            Index           =   44
            Left            =   4560
            TabIndex        =   146
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Responsabilità Riduz."
            Height          =   255
            Index           =   43
            Left            =   120
            TabIndex        =   145
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   43
            Left            =   4200
            Tag             =   "55"
            Top             =   2640
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cond. Tecnico"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   142
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   40
            Left            =   3480
            Tag             =   "46"
            Top             =   2040
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
            Left            =   2520
            TabIndex        =   95
            Tag             =   "20"
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ra&ggrupp. Ordine"
            Height          =   255
            Index           =   25
            Left            =   480
            TabIndex        =   37
            Top             =   600
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   25
            Left            =   5280
            Tag             =   "20"
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
            Left            =   2520
            TabIndex        =   94
            Tag             =   "16"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo &Vendita"
            Height          =   255
            Index           =   19
            Left            =   840
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   19
            Left            =   5280
            Tag             =   "16"
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
            Left            =   2520
            TabIndex        =   93
            Tag             =   "19"
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Modalità di &Pagamento"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   1695
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   5280
            Tag             =   "19"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. &Esterno 1"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   41
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   4200
            Tag             =   "10"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   15
            Left            =   4200
            Tag             =   "11"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. E&sterno 2"
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   43
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3075
         Index           =   3
         Left            =   240
         TabIndex        =   101
         Top             =   1080
         Width           =   8895
         Begin VB.ComboBox Combo10 
            Height          =   315
            Left            =   7800
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Tag             =   "26"
            Top             =   600
            Width           =   735
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Tag             =   "31"
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Frame Frame7 
            Height          =   2295
            Left            =   5640
            TabIndex        =   106
            Top             =   240
            Width           =   3135
            Begin MSMask.MaskEdBox mskField23 
               Height          =   255
               Left            =   2160
               TabIndex        =   73
               Tag             =   "27"
               Top             =   720
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField24 
               Height          =   255
               Left            =   2160
               TabIndex        =   75
               Tag             =   "28"
               Top             =   1080
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField25 
               Height          =   255
               Left            =   2160
               TabIndex        =   77
               Tag             =   "29"
               Top             =   1440
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               Format          =   "00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskField26 
               Height          =   255
               Left            =   2160
               TabIndex        =   79
               Tag             =   "30"
               Top             =   1800
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
               Caption         =   "Pedido Eventual"
               Height          =   255
               Index           =   20
               Left            =   120
               TabIndex        =   70
               Top             =   360
               Width           =   1935
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   20
               Left            =   2880
               Tag             =   "26"
               Top             =   360
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Momento"
               Height          =   255
               Index           =   21
               Left            =   120
               TabIndex        =   72
               Top             =   720
               Width           =   1935
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   21
               Left            =   2880
               Tag             =   "27"
               Top             =   720
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ordine Programmato"
               Height          =   255
               Index           =   22
               Left            =   120
               TabIndex        =   74
               Top             =   1080
               Width           =   1935
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   22
               Left            =   2880
               Tag             =   "28"
               Top             =   1080
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Partizione"
               Height          =   255
               Index           =   23
               Left            =   120
               TabIndex        =   76
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   23
               Left            =   2880
               Tag             =   "29"
               Top             =   1440
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ordine Aperto"
               Height          =   255
               Index           =   24
               Left            =   120
               TabIndex        =   78
               Top             =   1800
               Width           =   1935
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   24
               Left            =   2880
               Tag             =   "30"
               Top             =   1800
               Width           =   240
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Tag             =   "22"
            Top             =   720
            Width           =   2055
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Tag             =   "21"
            Top             =   240
            Width           =   2055
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            Height          =   255
            Left            =   1800
            TabIndex        =   65
            Tag             =   "24"
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox14 
            Height          =   255
            Left            =   1800
            TabIndex        =   67
            Tag             =   "25"
            Top             =   2040
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   255
            Left            =   1800
            TabIndex        =   69
            Tag             =   "37"
            Top             =   2400
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox22 
            Height          =   255
            Left            =   4200
            TabIndex        =   68
            Tag             =   "57"
            Top             =   2040
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   45
            Left            =   4920
            Tag             =   "57"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Servizio"
            Height          =   255
            Index           =   45
            Left            =   3000
            TabIndex        =   147
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Km Raggio"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   114
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   33
            Left            =   3120
            Tag             =   "37"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat&o Modifiche"
            Height          =   255
            Index           =   31
            Left            =   240
            TabIndex        =   62
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   31
            Left            =   3960
            Tag             =   "31"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Priorità"
            Height          =   255
            Index           =   30
            Left            =   1080
            TabIndex        =   66
            Top             =   2040
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   30
            Left            =   2280
            Tag             =   "25"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Or&dine di Riferimento"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   64
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   29
            Left            =   3240
            Tag             =   "24"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   27
            Left            =   3960
            Tag             =   "22"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Stato Ordine"
            Height          =   255
            Index           =   27
            Left            =   600
            TabIndex        =   60
            Top             =   720
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   26
            Left            =   3960
            Tag             =   "21"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Tipo Ordine"
            Height          =   255
            Index           =   26
            Left            =   840
            TabIndex        =   58
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3075
         Index           =   4
         Left            =   240
         TabIndex        =   115
         Top             =   1080
         Width           =   8895
         Begin VB.Frame Frame12 
            Caption         =   "Orario Apertura 1"
            Height          =   615
            Left            =   4560
            TabIndex        =   150
            Top             =   240
            Width           =   4215
            Begin MSMask.MaskEdBox MaskEdBox27 
               Height          =   255
               Left            =   2760
               TabIndex        =   124
               Tag             =   "62"
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
            Begin MSMask.MaskEdBox MaskEdBox28 
               Height          =   255
               Left            =   840
               TabIndex        =   123
               Tag             =   "61"
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
            Begin VB.Image imgState 
               Height          =   240
               Index           =   50
               Left            =   3480
               Tag             =   "62"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Alle"
               Height          =   255
               Index           =   50
               Left            =   2040
               TabIndex        =   152
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   51
               Left            =   1560
               Tag             =   "61"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dalle"
               Height          =   255
               Index           =   51
               Left            =   120
               TabIndex        =   151
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Orario Apertura 2"
            Height          =   615
            Left            =   4560
            TabIndex        =   153
            Top             =   840
            Width           =   4215
            Begin MSMask.MaskEdBox MaskEdBox25 
               Height          =   255
               Left            =   2760
               TabIndex        =   126
               Tag             =   "64"
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
            Begin MSMask.MaskEdBox MaskEdBox26 
               Height          =   255
               Left            =   840
               TabIndex        =   125
               Tag             =   "63"
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dalle"
               Height          =   255
               Index           =   48
               Left            =   120
               TabIndex        =   155
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   48
               Left            =   1560
               Tag             =   "63"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Alle"
               Height          =   255
               Index           =   49
               Left            =   2040
               TabIndex        =   154
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   49
               Left            =   3480
               Tag             =   "64"
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Orario Apertura 3"
            Height          =   615
            Left            =   4560
            TabIndex        =   156
            Top             =   1440
            Width           =   4215
            Begin MSMask.MaskEdBox MaskEdBox23 
               Height          =   255
               Left            =   2760
               TabIndex        =   128
               Tag             =   "66"
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
            Begin MSMask.MaskEdBox MaskEdBox24 
               Height          =   255
               Left            =   840
               TabIndex        =   127
               Tag             =   "65"
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dalle"
               Height          =   255
               Index           =   46
               Left            =   120
               TabIndex        =   158
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   46
               Left            =   1560
               Tag             =   "65"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Alle"
               Height          =   255
               Index           =   47
               Left            =   2040
               TabIndex        =   157
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   47
               Left            =   3480
               Tag             =   "66"
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Tag             =   "58"
            Top             =   2280
            Width           =   2655
         End
         Begin VB.ComboBox Combo8 
            Height          =   315
            Left            =   6120
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Tag             =   "59"
            Top             =   2280
            Width           =   2295
         End
         Begin VB.ComboBox Combo9 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Tag             =   "60"
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Frame Frame8 
            Caption         =   "Orario Consegna 1"
            Height          =   615
            Left            =   240
            TabIndex        =   116
            Top             =   240
            Width           =   4215
            Begin MSMask.MaskEdBox MaskEdBox12 
               Height          =   255
               Left            =   2760
               TabIndex        =   118
               Tag             =   "39"
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
            Begin MSMask.MaskEdBox MaskEdBox11 
               Height          =   255
               Left            =   840
               TabIndex        =   117
               Tag             =   "38"
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
            Begin VB.Image imgState 
               Height          =   240
               Index           =   34
               Left            =   3480
               Tag             =   "39"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Alle"
               Height          =   255
               Index           =   34
               Left            =   2040
               TabIndex        =   133
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   35
               Left            =   1560
               Tag             =   "38"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dalle"
               Height          =   255
               Index           =   35
               Left            =   120
               TabIndex        =   132
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Orario Consegna 2"
            Height          =   615
            Left            =   240
            TabIndex        =   134
            Top             =   840
            Width           =   4215
            Begin MSMask.MaskEdBox MaskEdBox15 
               Height          =   255
               Left            =   2760
               TabIndex        =   120
               Tag             =   "41"
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
            Begin MSMask.MaskEdBox MaskEdBox16 
               Height          =   255
               Left            =   840
               TabIndex        =   119
               Tag             =   "40"
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dalle"
               Height          =   255
               Index           =   37
               Left            =   120
               TabIndex        =   136
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   37
               Left            =   1560
               Tag             =   "40"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Alle"
               Height          =   255
               Index           =   36
               Left            =   2040
               TabIndex        =   135
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   36
               Left            =   3480
               Tag             =   "41"
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Orario Consegna 3"
            Height          =   615
            Left            =   240
            TabIndex        =   137
            Top             =   1440
            Width           =   4215
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   255
               Left            =   2760
               TabIndex        =   122
               Tag             =   "43"
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
            Begin MSMask.MaskEdBox MaskEdBox18 
               Height          =   255
               Left            =   840
               TabIndex        =   121
               Tag             =   "42"
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dalle"
               Height          =   255
               Index           =   39
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   39
               Left            =   1560
               Tag             =   "42"
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Alle"
               Height          =   255
               Index           =   38
               Left            =   2040
               TabIndex        =   138
               Top             =   240
               Width           =   495
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   38
               Left            =   3480
               Tag             =   "43"
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Azione consentita"
            Height          =   255
            Index           =   52
            Left            =   240
            TabIndex        =   161
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   52
            Left            =   4440
            Tag             =   "58"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Contattare cliente"
            Height          =   255
            Index           =   53
            Left            =   4680
            TabIndex        =   160
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   53
            Left            =   8520
            Tag             =   "59"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Accordi con cliente"
            Height          =   255
            Index           =   54
            Left            =   120
            TabIndex        =   159
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   54
            Left            =   4440
            Tag             =   "60"
            Top             =   2640
            Width           =   240
         End
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "23"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   2
         Tag             =   "32"
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1215
         Left            =   120
         OleObjectBlob   =   "OIL2OCT0.frx":1830
         TabIndex        =   80
         Tag             =   "DETAIL"
         Top             =   4320
         Width           =   9135
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8760
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
               Picture         =   "OIL2OCT0.frx":7E42
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2OCT0.frx":7F54
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2OCT0.frx":8066
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3555
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6271
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Ordine &1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Ordine &2"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Ordine &3"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Ordine &4"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Avan&zamento"
         Height          =   255
         Index           =   28
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   8400
         Tag             =   "23"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3120
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Numero Ordine"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "OIL2OCT0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 9
Public GridArrayDetail As ArrayGridCls

Public WithEvents DataClass As COrdine
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As COrdineriga
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private mstrIDImballoDefault As String
Private mstrIDPosDefault As String
Private mstrIDPagamento As String
Private mstrIDIva As String

Private mintCurFrame As Integer
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object

Private APR As CProdotto
Private AIV As CIVA
Private AMP As CPagamento
Private APD As CPosDoganale
Private AIM As CImballo
Private AXX As CParametro
Private ADZ As CDestinazione

Public Function InsertDetail() As Boolean
    InsertDetail = DetailClass.Insert
End Function

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = GridArrayDetail.value(DBGrid.bookmark, 0)
    End If
End Function

Private Sub DisplayDetail()
    With DetailClass
        .DS.Fields("0").FindValue = DataClass.DS.Fields(DataClass.DS.KeyFields("1")).value
        .GetList
        RefreshUnboundGrid TDBGrid1, .DS.List, GridArrayDetail, _
                           .DS.HowMany, MAX_COLS
    End With
End Sub

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
    Dim strIDViaggio As String
    Dim strSelectionFormula As String
    Dim strNewOrder As String
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
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
                    Case fsInsert
                        blnResult = InsertRecord
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
                varReportsList(0, 0) = "Stampa Ordine"
                varReportsList(1, 0) = "oil2oct1.rpt"
                varReportsList(2, 0) = 1
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                CRPrint varReportsList, strSelectionFormula, Me
            End If
            
        Case fcOpen
            strIDViaggio = DataClass.GetTripID
            If strIDViaggio <> vbNullString Then
                Set LinkForm = DisplayForm(fcVIAGVVT)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .GetRecord strIDViaggio
                        .SetState fsEdit
                        .Command1.Enabled = False
                        Me.Enabled = False
                    End With
                End If
            End If
            
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcORDIOCR)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
            
        Case fcLinkRefresh
            DisplayDetail
            
        Case fcFind
            If DataClass.CopyOrder(DataClass.DS.Fields("0").value, , strNewOrder, False, , , DataClass.DS.Field("Data Consegna Richiesta").value) Then
                DataClass.DS.DisplayMessage "E' stato inserito l'ordine n. " & strNewOrder, 0
            Else
                DataClass.DS.DisplayMessage "Si è verificato un errore durante la copia dell'ordine", 9
            End If
            
        Case fcExport
            If Not IsNull(TDBGrid1.bookmark) Then
                AnnullaRiga
            End If
            
        Case fcExit
            Unload Me
            
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
                    
                    If (.Fields("23").value <> "A") Then
                        With TDBGrid1
                            .AllowAddNew = False
                            .AllowDelete = False
                            .AllowUpdate = False
                        End With
                        If (.Fields("23").value <> "M") Then
                            Command1.Enabled = True
                        Else
                            Command1.Enabled = False
                        End If
                    Else
                        Command1.Enabled = False
                        
                        '-- DP 06/08/08
                        cmdExport.Enabled = True
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
    Dim strMsgDepRaggio As String
    Dim C As Control
    Dim CErr As CError
    
    'Scansione della finestra
    If Not DataClass.DS.ScanClient(0) Then
        GoTo ErrHandler
    End If
    
    If DataClass.DS.Fields.Validate Then
    
        ' --- Controlliamo di non aver modificato il deposito senza aver aggiornato il raggio
        ' --- o viceversa
        If (DataClass.DS.Fields("17").Updated And (Not DataClass.DS.Fields("37").Updated)) Or _
           (DataClass.DS.Fields("37").Updated And (Not DataClass.DS.Fields("17").Updated)) Then
        
            If DataClass.DS.Fields("17").Updated Then
                strMsgDepRaggio = "Il deposito è stato modificato. Vuoi proseguire?"
            Else
                strMsgDepRaggio = "Il raggio in Km è stato modificato. Vuoi proseguire?"
            End If
            If MsgBox(strMsgDepRaggio, vbYesNo + vbExclamation, "Attenzione") = vbNo Then
                GoTo ErrHandler
            End If
        End If
        
        If DataClass.BeginSave(blnConcurError) Then
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
            strMsg = "L'ordine non può venire modificato"
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
            mskField7.SetFocus
            If DataClass.DS.Fields("23").value = "V" Then
                Check1(1).Enabled = False
            End If
    End Select
End Sub



Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Data consegna richiesta"
        Case 1
            Set M = MaskEdBox3
            strTitle = "Data consegna pianificata"
        Case 2
            Set M = mskField4
            strTitle = "Data inserimento"
        Case 3
            Set M = MaskEdBox1
            strTitle = "Data ricezione"
        Case 4
            Set M = MaskEdBox4
            strTitle = "Data consegna effettiva"
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
    
    Loading = True
    
    Set DataClass = New COrdine
    Set DetailClass = New COrdineriga
    Set ErrH = New CErrorHandler
    Set GridArrayDetail = New ArrayGridCls
    
    mintCurFrame = 1
    ColTags(1) = "1"
    ColTags(2) = "2"
    ColTags(3) = "4"
    ColTags(4) = "3"
    ColTags(5) = "7"
    ColTags(6) = "11"
    ColTags(7) = "14"
    ColTags(8) = "15"
    ColTags(9) = "29"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set .ErrH = ErrH
    End With
    
    With DetailClass.DS
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
    
    Set AIV = New CIVA
    With AIV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AMP = New CPagamento
    With AMP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APD = New CPosDoganale
    With APD.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AIM = New CImballo
    With AIM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        With .MultiSoc.DS
            Set .DA = DataConnect.DA
            Set .ErrH = ErrH
        End With
    End With
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = strIDParam
        AXX.Find
        mstrIDImballoDefault = Trim$(.Fields("1").value)
        mstrIDPosDefault = Trim$(.Fields("2").value)
    End With
    ' ---
    
    SetValueItems TDBGrid1.Columns(1), APR
    TDBGrid1.Columns(1).ValueItems.Translate = False
    SetValueItems TDBGrid1.Columns(2), APD
    SetValueItems TDBGrid1.Columns(3), AIM
    SetValueItems TDBGrid1.Columns(6), AIV
    SetValueItems TDBGrid1.Columns(7), AMP
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AXX = Nothing
    Set AIM = Nothing
    Set APD = Nothing
    Set AXX = Nothing
    Set AMP = Nothing
    Set AIV = Nothing
    Set APR = Nothing
    Set ADZ = Nothing
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set DetailClass = Nothing
    Set GridArrayDetail = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            
            Select Case UCase(.Name)
                Case "FRMOCTSELECT2"
                    .DoCommand fcLinkRefresh
                
            End Select
        End With
    End If
    Set OIL2OCT0 = Nothing
End Sub











Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub


Private Sub mskField7_LostFocus()
    With ADZ
        .DS.Fields("0").FindValue = mskField7
        .Find
        mstrIDIva = .DS.Fields("14").value
    End With
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame6(TabStrip1.SelectedItem.Index).Visible = True
        Frame6(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index

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
                    For i = 0 To GridArrayDetail.RowCount - 1
                        If .Text = GridArrayDetail.value(i, 0) Then
                            blnDataOk = False
                        End If
                    Next i
                End If
            Case 4
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                Else
                    If CDbl(.Text) <= 0 Then
                        blnDataOk = False
                    End If
                End If
        End Select
    End With
    Cancel = Not blnDataOk
End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    With TDBGrid1
        .Columns(2).value = mstrIDPosDefault
        .Columns(3).value = mstrIDImballoDefault
        .Columns(5).value = mstrIDIva
        .Columns(6).value = CStr(mskField13)
    End With

End Sub


Private Sub TDBGrid1_BeforeUpdate(Cancel As Integer)
    Dim blnDataOk As Boolean
    
    blnDataOk = True
    With TDBGrid1.Columns
        If Not IsNumeric(.Item(0).Text) Then
            blnDataOk = False
        End If
        If .Item(1).Text = vbNullString Then
            blnDataOk = False
        End If
        If .Item(2).Text = vbNullString Then
            blnDataOk = False
        End If
        If .Item(3).Text = vbNullString Then
            blnDataOk = False
        End If
        If Not IsNumeric(.Item(4).Text) Then
            blnDataOk = False
        Else
            If CDbl(.Item(4).Text) <= 0 Then
                blnDataOk = False
            End If
        End If
    End With
    Cancel = Not blnDataOk

End Sub

'-- DP 16/09/08
Private Function AnnullaRiga() As Boolean
    Dim blnConcurError As Boolean
    
    If MsgBox("ATENCIÓN: ¿Se desea anular la línea seleccionada?", vbYesNo, "ResOil.2") = vbYes Then
        If DetailClass.SetAnnullato(CVar(mskField1.Text), CVar(GridArrayDetail.value(TDBGrid1.bookmark, 0)), True) Then
            DisplayDetail
        Else
            MsgBox "ATENCIÓN: Han pasado errores durante la anulación de la línea seleccionada", vbCritical + vbOKOnly, "ResOil.2"
            GoTo ErrHandler
        End If
    End If

ErrHandler:
    AnnullaRiga = False
    Exit Function
    
EndProc:
    AnnullaRiga = True
End Function
