VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2VVC0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSEGNA"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   9810
   Begin VB.Frame Frame5 
      Caption         =   "Generación y Cancelación de eventos"
      Height          =   1095
      Left            =   120
      TabIndex        =   107
      Top             =   6360
      Width           =   9615
      Begin VB.ComboBox cmb_cges 
         Height          =   315
         Left            =   3840
         TabIndex        =   110
         Text            =   "Selecionar Tipo Evento:"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton btn_generaevento 
         Caption         =   "Genera Evento"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         Picture         =   "OIL2VVC0.frx":0000
         TabIndex        =   109
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Anular Orden Palm"
         Height          =   615
         Left            =   600
         Picture         =   "OIL2VVC0.frx":0102
         TabIndex        =   108
         Tag             =   "DELLTR"
         Top             =   360
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txt_dataev 
         Height          =   255
         Left            =   8160
         TabIndex        =   114
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
      Begin MSMask.MaskEdBox txt_hrev 
         Height          =   255
         Left            =   8160
         TabIndex        =   115
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
      Begin VB.Label Label5 
         Caption         =   "Hora evento:"
         Height          =   255
         Left            =   7080
         TabIndex        =   113
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Data evento:"
         Height          =   255
         Left            =   7080
         TabIndex        =   112
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Generar evento:"
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   120
      TabIndex        =   84
      Top             =   5520
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
         Picture         =   "OIL2VVC0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   85
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
         Picture         =   "OIL2VVC0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   86
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
         Picture         =   "OIL2VVC0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2VVC0.frx":050A
         TabIndex        =   79
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7800
         Picture         =   "OIL2VVC0.frx":060C
         TabIndex        =   80
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2VVC0.frx":070E
         TabIndex        =   78
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Annullata"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   81
      Tag             =   "38"
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   83
      Top             =   7455
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
      Left            =   1680
      TabIndex        =   1
      Tag             =   "1"
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   12648447
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLink 
      Height          =   255
      Left            =   6600
      TabIndex        =   90
      Tag             =   "0"
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
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
            Picture         =   "OIL2VVC0.frx":0810
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2VVC0.frx":0922
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2VVC0.frx":0A34
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   1
      Left            =   240
      TabIndex        =   82
      Top             =   840
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   120
         TabIndex        =   91
         Top             =   120
         Width           =   9135
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   255
            Left            =   1440
            TabIndex        =   3
            Tag             =   "4"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Tag             =   "5"
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   1440
            TabIndex        =   7
            Tag             =   "6"
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   3960
            TabIndex        =   11
            Tag             =   "7"
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   7920
            TabIndex        =   13
            Tag             =   "8"
            Top             =   1320
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Tag             =   "9"
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField15 
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Tag             =   "10"
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField20 
            Height          =   255
            Left            =   5640
            TabIndex        =   17
            Tag             =   "11"
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   4200
            Tag             =   "10"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codice &Fiscale"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   8400
            Tag             =   "11"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Partita IVA"
            Height          =   255
            Index           =   18
            Left            =   4680
            TabIndex        =   16
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&CAP"
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   8
            Top             =   1320
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   15
            Left            =   2760
            Tag             =   "9"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Prov."
            Height          =   255
            Index           =   14
            Left            =   7440
            TabIndex        =   12
            Top             =   1320
            Width           =   375
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   8400
            Tag             =   "8"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "L&ocalità"
            Height          =   255
            Index           =   13
            Left            =   3120
            TabIndex        =   10
            Top             =   1320
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   6720
            Tag             =   "7"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "In&dirizzo"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   6
            Top             =   960
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   5640
            Tag             =   "6"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   5640
            Tag             =   "5"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ragione sociale"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   1215
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
            Left            =   2880
            TabIndex        =   92
            Tag             =   "4"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "C&liente"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   10
            Left            =   5640
            Tag             =   "4"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   120
         TabIndex        =   93
         Top             =   2280
         Width           =   9135
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   255
            Left            =   1440
            TabIndex        =   19
            Tag             =   "12"
            Top             =   240
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
            Left            =   1440
            TabIndex        =   21
            Tag             =   "13"
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   1440
            TabIndex        =   23
            Tag             =   "14"
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   3960
            TabIndex        =   27
            Tag             =   "15"
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   255
            Left            =   7920
            TabIndex        =   29
            Tag             =   "16"
            Top             =   1320
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox15 
            Height          =   255
            Left            =   1440
            TabIndex        =   25
            Tag             =   "17"
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   1440
            TabIndex        =   31
            Tag             =   "18"
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   255
            Left            =   5640
            TabIndex        =   33
            Tag             =   "19"
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   27
            Left            =   4200
            Tag             =   "18"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codice Fi&scale"
            Height          =   255
            Index           =   27
            Left            =   240
            TabIndex        =   30
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   26
            Left            =   8400
            Tag             =   "19"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Partita &IVA"
            Height          =   255
            Index           =   26
            Left            =   4680
            TabIndex        =   32
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "CA&P"
            Height          =   255
            Index           =   25
            Left            =   840
            TabIndex        =   24
            Top             =   1320
            Width           =   495
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   25
            Left            =   2760
            Tag             =   "17"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pro&v."
            Height          =   255
            Index           =   24
            Left            =   7440
            TabIndex        =   28
            Top             =   1320
            Width           =   375
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   24
            Left            =   8400
            Tag             =   "16"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Locali&tà"
            Height          =   255
            Index           =   20
            Left            =   3240
            TabIndex        =   26
            Top             =   1320
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   20
            Left            =   6720
            Tag             =   "15"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Indiri&zzo"
            Height          =   255
            Index           =   19
            Left            =   600
            TabIndex        =   22
            Top             =   960
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   19
            Left            =   5640
            Tag             =   "14"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   5640
            Tag             =   "13"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ra&gione sociale"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1215
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
            Left            =   2880
            TabIndex        =   94
            Tag             =   "12"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "D&estinazione"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   5640
            Tag             =   "12"
            Top             =   240
            Width           =   240
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   2
      Left            =   240
      TabIndex        =   88
      Top             =   840
      Visible         =   0   'False
      Width           =   9375
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   1440
         TabIndex        =   47
         Tag             =   "26"
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtNoteEstese 
         Height          =   885
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   61
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2760
         Picture         =   "OIL2VVC0.frx":0B46
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2400
         Width           =   375
      End
      Begin VB.Frame Frame6 
         Caption         =   "S&tato"
         Height          =   2055
         Left            =   6840
         TabIndex        =   77
         Top             =   120
         Width           =   2295
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   1200
            TabIndex        =   62
            Tag             =   "32"
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox21 
            Height          =   255
            Left            =   1200
            TabIndex        =   64
            Tag             =   "34"
            Top             =   960
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
            TabIndex        =   63
            Tag             =   "33"
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox22 
            Height          =   255
            Left            =   1200
            TabIndex        =   65
            Tag             =   "35"
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox23 
            Height          =   255
            Left            =   1200
            TabIndex        =   66
            Tag             =   "36"
            Top             =   1680
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 1"
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   34
            Left            =   1800
            Tag             =   "32"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 2"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   101
            Top             =   600
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   33
            Left            =   1800
            Tag             =   "33"
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   32
            Left            =   1800
            Tag             =   "34"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 3"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   100
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 4"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   99
            Top             =   1320
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   31
            Left            =   1800
            Tag             =   "35"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Flag Stato 5"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   98
            Top             =   1680
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   30
            Left            =   1800
            Tag             =   "36"
            Top             =   1680
            Width           =   240
         End
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Tag             =   "21"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Tag             =   "23"
         Top             =   960
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   7320
         TabIndex        =   57
         Tag             =   "37"
         Top             =   2760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1440
         TabIndex        =   55
         Tag             =   "28"
         Top             =   2760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   4800
         TabIndex        =   56
         Tag             =   "29"
         Top             =   2760
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   5160
         TabIndex        =   53
         Tag             =   "24"
         Top             =   2400
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   8160
         TabIndex        =   54
         Tag             =   "25"
         Top             =   2400
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   1440
         TabIndex        =   48
         Tag             =   "27"
         Top             =   1680
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
         Left            =   3360
         TabIndex        =   38
         Tag             =   "3"
         Top             =   240
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
         TabIndex        =   58
         Tag             =   "30"
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   6720
         TabIndex        =   59
         Tag             =   "31"
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField6 
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Tag             =   "2"
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox18 
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Tag             =   "20"
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox19 
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Tag             =   "22"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   3
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox24 
         Height          =   255
         Left            =   1440
         TabIndex        =   49
         Tag             =   "39"
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox25 
         Height          =   255
         Left            =   4200
         TabIndex        =   50
         Tag             =   "40"
         Top             =   2040
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
         Left            =   1440
         TabIndex        =   51
         Tag             =   "41"
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
      Begin MSMask.MaskEdBox MaskEdBox28 
         Height          =   255
         Left            =   5760
         TabIndex        =   60
         Tag             =   "42"
         Top             =   3480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   42
         Left            =   8880
         Tag             =   "42"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Note 1"
         Height          =   255
         Index           =   42
         Left            =   5160
         TabIndex        =   106
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Note Estese"
         Height          =   255
         Index           =   41
         Left            =   360
         TabIndex        =   105
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Data Consegna"
         Height          =   255
         Index           =   40
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   40
         Left            =   3240
         Tag             =   "41"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Km raggio stradali"
         Height          =   195
         Index           =   39
         Left            =   2760
         TabIndex        =   104
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   39
         Left            =   4920
         Tag             =   "40"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Km raggio"
         Height          =   195
         Index           =   36
         Left            =   600
         TabIndex        =   103
         Top             =   2040
         Width           =   765
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   36
         Left            =   2160
         Tag             =   "39"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   37
         Left            =   5880
         Tag             =   "21"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ca&lcolo Km prev."
         Height          =   255
         Index           =   37
         Left            =   2880
         TabIndex        =   41
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   38
         Left            =   5880
         Tag             =   "23"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Calcolo Km defin."
         Height          =   255
         Index           =   38
         Left            =   2760
         TabIndex        =   45
         Top             =   960
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   29
         Left            =   2160
         Tag             =   "20"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Km parziali prev."
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1245
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   2160
         Tag             =   "22"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Km parziali d&efin."
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Se&quenza"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   2160
         Tag             =   "2"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5280
         Tag             =   "26"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Zona"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   70
         Top             =   1320
         Width           =   495
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
         TabIndex        =   97
         Tag             =   "26"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   5280
         Tag             =   "27"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Ven&dita"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   71
         Top             =   1680
         Width           =   975
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
         TabIndex        =   96
         Tag             =   "27"
         Top             =   1680
         Width           =   3135
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
         Left            =   3960
         TabIndex        =   95
         Tag             =   "3"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Società"
         Height          =   255
         Index           =   11
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   6480
         Tag             =   "3"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Rif. Esterno 1"
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   74
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   21
         Left            =   3600
         Tag             =   "30"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ri&f. Esterno 2"
         Height          =   255
         Index           =   22
         Left            =   5520
         TabIndex        =   75
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   22
         Left            =   8880
         Tag             =   "31"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Te&mpo Scarico"
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   69
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   8880
         Tag             =   "25"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5880
         Tag             =   "24"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora arr&ivo"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   68
         Top             =   2400
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   23
         Left            =   5520
         Tag             =   "29"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Conse&gna Rif."
         Height          =   255
         Index           =   23
         Left            =   3600
         TabIndex        =   73
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3000
         Tag             =   "28"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Viaggio Rif."
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   72
         Top             =   2760
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   8880
         Tag             =   "37"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ordine Cliente"
         Height          =   255
         Index           =   6
         Left            =   6120
         TabIndex        =   76
         Top             =   2760
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   120
      TabIndex        =   34
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dati Consegna &1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dati Consegna &2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   35
      Left            =   8040
      Tag             =   "0"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero Viaggio"
      Height          =   255
      Index           =   28
      Left            =   5160
      TabIndex        =   89
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Numero Consegna"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   0
      Left            =   2400
      Tag             =   "1"
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "OIL2VVC0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3

Public WithEvents DataClass As CViaggioConsegna
Attribute DataClass.VB_VarHelpID = -1
Private VVP As CViaggioProdotto
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mstrNoteEstese As String

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
        Case fcdelltr
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
                
                COD = InputBox("Se va a Cancelar la Orden de la Palm, Esta Seguro de querer continuar con la operación? Si es así introduce el codigo de desbloqueo.", "Eliminar Viaje")
                   If ((COD) = mskField7.Text) Then
                      
                     Dim KVE As CSatellitareEvento
                          Set KVE = New CSatellitareEvento
                            With KVE
                            Set .DS.DA = DataConnect.DA
                                Set .DS.ErrH = DataConnect.DA.ErrH
                                If Not KVE.EliminaTareaPalm(mskField7.Text, DataConnect) Then
                                End If
                            End With
                   
                   Else
                     MsgBox ("Codigo Incorrecto")
                   End If
                   
        Case fcExit
            Unload Me
            
    End Select
End Sub

Public Function GetRecord(strIDOrdine As String, strIDRiga As String) As Boolean
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strIDOrdine
        .Fields(.KeyFields(2)).FindValue = strIDRiga
        
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                End If
                TxtNoteEstese.Text = DataClass.DS.Fields("43").value
                mstrNoteEstese = TxtNoteEstese.Text
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
    
    DataClass.DS.Fields("47").NewValue = mskField3.Text
    DataClass.DS.Fields("47").Updated = True
   
    DataClass.DS.Fields("49").NewValue = MaskEdBox26.Text
    DataClass.DS.Fields("49").Updated = True
   
    If mstrNoteEstese <> TxtNoteEstese.Text Then
        DataClass.DS.Fields("43").NewValue = TxtNoteEstese.Text
        DataClass.DS.Fields("43").Updated = True
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
            strMsg = "La consegna non può venire modificata"
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
            mskField2.SetFocus
                    
            '-- DP 16/10/08
            '-- Se i dati della consegna sono già stati inviati in un GES, data e ora non devono essere modificabili
            If VVP.blnGESAlreadyTransmitted(DataClass.DS.Fields("0").value, DataClass.DS.Fields("1").value) Then
                MaskEdBox26.Enabled = False
                cmdCalendar(0).Enabled = False
                mskField3.Enabled = False
            End If
            
    End Select
End Sub

Private Sub btn_generaevento_Click()
Dim DA As CDataAccess
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
Dim codentrega  As String
Dim CODINC As String
Dim h As Integer

  CodEv = cmb_cges.ItemData(cmb_cges.ListIndex)
  COD = mskLink.Text
  codentrega = mskField7.Text
  HREV = txt_hrev.Text
  DYEV = txt_dataev.Text
  
  Do While (Len(CodEv) < 3)
     CodEv = "0" + CodEv
  Loop

  'Comprobaciones de integridad de los datos
  If (Len(CodEv) = 3) Then
    If (Len(txt_hrev.Text) > 0 And Len(txt_dataev.Text) > 0 And txt_hrev.Text <> "__:__" And txt_dataev <> "__/__/__") Then
    
    If (CInt(CodEv) > 8 And CInt(CodEv) < 16) Then
    
      CODINC = InputBox("Se enviará una incidencia a CEPSA, Introduce código para proseguir.", "Generar Incidencia")
      If ((CODINC) = "incpass2008") Then
        'Generar el evento en la tabla
        SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + COD + "','" + codentrega + "','" + CodEv + "','" + DYEV + " " + HREV + "','" + DYEV + " " + HREV + "','0')"
        If DA.ExecuteQuery(SQL_GEN_EV) Then
           MsgBox ("Evento '" + CodEv + "' generado con exito   El: " + CStr(DYEV) + " A las :" + CStr(HREV))
        End If
    
      Else
        MsgBox ("Codigo Incorrecto")
      End If
    
    Else
        'Generar el evento en la tabla
        SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + COD + "','" + codentrega + "','" + CodEv + "','" + DYEV + " " + HREV + "','" + DYEV + " " + HREV + "','0')"
        If DA.ExecuteQuery(SQL_GEN_EV) Then
           MsgBox ("Evento '" + CodEv + "' generado con exito   El: " + CStr(DYEV) + " A las :" + CStr(HREV))
        End If


  End If
    
    
    Else
       MsgBox ("Hay que introducir una fecha y una hora para crear el evento")
    End If
  Else
       MsgBox ("Numero de evento no indicado.")
  End If
  
  
  
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox26
            strTitle = "Data consegna"
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
    
    Loading = True
    
    mintCurFrame = 1
    Set DataClass = New CViaggioConsegna
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
'    cmb_cges.AddItem "005"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 0
'    cmb_cges.AddItem "006"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 1
'    cmb_cges.AddItem "009"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 2
'    cmb_cges.AddItem "010"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 3
'    cmb_cges.AddItem "011"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 4
'    cmb_cges.AddItem "012"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 5
'    cmb_cges.AddItem "013"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 6
'    cmb_cges.AddItem "014"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 7
'    cmb_cges.AddItem "015"
'    cmb_cges.ItemData(cmb_cges.NewIndex) = 8
    
    cmb_cges.AddItem "005-Llegada a destino"
    cmb_cges.ItemData(0) = "005"
    cmb_cges.AddItem "006-Salida de destino"
    cmb_cges.ItemData(1) = "006"
    cmb_cges.AddItem "009-Imposible descarga total"
    cmb_cges.ItemData(2) = "009"
    cmb_cges.AddItem "010-Imposible descarga parcial"
    cmb_cges.ItemData(3) = "010"
    cmb_cges.AddItem "011-Contaminacion"
    cmb_cges.ItemData(4) = "011"
    cmb_cges.AddItem "012-Derrame"
    cmb_cges.ItemData(5) = "012"
    cmb_cges.AddItem "013-Discrepancia cantidad"
    cmb_cges.ItemData(6) = "013"
    cmb_cges.AddItem "014-Accidente"
    cmb_cges.ItemData(7) = "014"
    cmb_cges.AddItem "015-Averia"
    cmb_cges.ItemData(8) = "015"
    

    
    
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set VVP = Nothing
    
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    
    Set OIL2VVC0 = Nothing
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
