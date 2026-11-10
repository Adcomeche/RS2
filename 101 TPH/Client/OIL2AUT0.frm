VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2AUT0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UNITA' DI TRASPORTO"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   44
      Top             =   4920
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   3600
         Picture         =   "OIL2AUT0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F11 - Esporta"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2AUT0.frx":0102
         TabIndex        =   62
         Tag             =   "CUSTOM1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1920
         Picture         =   "OIL2AUT0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2AUT0.frx":0306
         TabIndex        =   59
         Tag             =   "FIND"
         Top             =   720
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
         Left            =   3600
         Picture         =   "OIL2AUT0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   45
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
         Index           =   3
         Left            =   240
         Picture         =   "OIL2AUT0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   46
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
         Height          =   240
         Index           =   2
         Left            =   7680
         Picture         =   "OIL2AUT0.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   47
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
         Left            =   1920
         Picture         =   "OIL2AUT0.frx":070E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   48
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2AUT0.frx":0810
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2AUT0.frx":0912
         TabIndex        =   38
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2AUT0.frx":0A14
         TabIndex        =   39
         Tag             =   "PRINT"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2AUT0.frx":0B16
         TabIndex        =   40
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2AUT0.frx":0C18
         TabIndex        =   37
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2AUT0.frx":0D1A
         TabIndex        =   36
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   6225
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
      Height          =   4815
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Tag             =   "9"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "13"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   34
         Tag             =   "16"
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Tag             =   "1"
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   7920
         TabIndex        =   9
         Tag             =   "8"
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Tag             =   "6"
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   1215
         Left            =   4200
         OleObjectBlob   =   "OIL2AUT0.frx":0E1C
         TabIndex        =   35
         Tag             =   "MULTISOC"
         Top             =   3480
         Width           =   4995
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Tag             =   "11"
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField8 
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Tag             =   "7"
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Tag             =   "2"
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField6 
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Tag             =   "12"
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField5 
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField11 
         Height          =   255
         Left            =   5520
         TabIndex        =   27
         Tag             =   "4"
         Top             =   2520
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
         Left            =   1080
         TabIndex        =   19
         Tag             =   "10"
         Top             =   2520
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
         Left            =   5520
         TabIndex        =   29
         Tag             =   "5"
         Top             =   2880
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
         Left            =   7440
         TabIndex        =   11
         Tag             =   "14"
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Tag             =   "15"
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6240
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5640
         Top             =   120
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
               Picture         =   "OIL2AUT0.frx":5096
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2AUT0.frx":51A8
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2AUT0.frx":52BA
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   7440
         TabIndex        =   13
         Tag             =   "17"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Presidio"
         Height          =   255
         Index           =   16
         Left            =   6240
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   8880
         Tag             =   "17"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Con&tratto"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   14
         Top             =   1800
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
         Left            =   1680
         TabIndex        =   60
         Tag             =   "15"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   4320
         Tag             =   "15"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   8880
         Tag             =   "14"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Badge"
         Height          =   255
         Index           =   14
         Left            =   6720
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Disp&onibilità"
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   32
         Top             =   3840
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   3720
         Tag             =   "9"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Soc&iale/Terzo"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3720
         Tag             =   "13"
         Top             =   3480
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
         Left            =   6120
         TabIndex        =   57
         Tag             =   "5"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E&nte Em."
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   8880
         Tag             =   "5"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   4320
         Tag             =   "10"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Vettore"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   22
         Top             =   1800
         Width           =   735
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
         Left            =   1680
         TabIndex        =   56
         Tag             =   "10"
         Top             =   2520
         Width           =   2535
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
         Left            =   6120
         TabIndex        =   55
         Tag             =   "4"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vett. T&r."
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   8880
         Tag             =   "4"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   5160
         Tag             =   "3"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rimorc&hio"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   735
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
         Left            =   2520
         TabIndex        =   54
         Tag             =   "3"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   4320
         Tag             =   "12"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A&utista 1"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   26
         Top             =   2520
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
         Left            =   1680
         TabIndex        =   53
         Tag             =   "12"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5160
         Tag             =   "2"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Motrice"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1080
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
         Left            =   2520
         TabIndex        =   52
         Tag             =   "2"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   8880
         Tag             =   "7"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Au&tista 2"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   28
         Top             =   2880
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
         Left            =   6120
         TabIndex        =   51
         Tag             =   "7"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   4320
         Tag             =   "11"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "D&eposito"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   615
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
         Left            =   1680
         TabIndex        =   50
         Tag             =   "11"
         Top             =   2880
         Width           =   2535
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
         Left            =   6120
         TabIndex        =   43
         Tag             =   "6"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   16
         Top             =   2160
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   8880
         Tag             =   "6"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Portata Tot. Kg"
         Height          =   255
         Index           =   5
         Left            =   6720
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   8880
         Tag             =   "8"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5160
         Tag             =   "1"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   1800
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "OIL2AUT0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GridArray As ArrayGridCls
Public WithEvents DataClass As CUnitaDiTrasporto
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public mstrExportFileName As String
Public Soc As CSocieta
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjOutFile01 As CTextFile
Private mobjOutFile02 As CTextFile
Private mobjOutFile04 As CTextFile
Private mobjOutFile05 As CTextFile
Private ATN As CTipoContratto
Private ATM As CTipoAutomezzo
Private ATS As CTipoCisterna
Private AVE As CVettore
Private AEE As CEnteEmittente
Private ASC As CSocieta

Private Function ExportData() As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDSocieta As String
    Dim strIDAutista1 As String
    Dim strIDAutista2 As String
    Dim varExportData As Variant
    Dim blnResult As Boolean
    Dim intFileNum As Integer
    Dim intRows As Integer
    Dim intSpacePos As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ErrCode As Errors
    Dim C As CError
    
    DoEvents
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    
    blnResult = DataClass.GetExportData(varExportData)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If Not mobjOutFile01.OpenFile(mstrExportFileName, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    intFileNum = mobjOutFile01.FileNum
    mobjOutFile02.FileNum = intFileNum
    mobjOutFile04.FileNum = intFileNum
    mobjOutFile05.FileNum = intFileNum
    
    intRows = UBound(varExportData, 2)
    
    
'    ' --- Esporta i dati del contratto
'    For i = 0 To intRows
'        strIDSocieta = varExportData(5, i)
'        With mobjOutFile01
'            If Len(varExportData(0, i)) > 0 Then
'                .Item("3").value = varExportData(0, i)
'                .Item("4").value = varExportData(1, i)
'                .Item("5").value = varExportData(2, i)
'                .Item("6").value = varExportData(3, i)
'                ATN.DS.Fields("Codice").FindValue = varExportData(77, i)
'                ATN.Find
'                .Item("8").value = ATN.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
'                .Item("9").value = varExportData(13, i)
'                .Item("10").value = varExportData(42, i)
'                AVE.DS.Fields("0").FindValue = varExportData(4, i)
'                AVE.Find
'                .Item("11").value = AVE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
'                AEE.DS.Fields("0").FindValue = varExportData(6, i)
'                AEE.Find
'                .Item("13").value = AEE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
'                If Not .WriteLine(True) Then
'                    ErrCode = OutputError
'                    GoTo ErrHandler
'                End If
'            End If
'        End With
'    Next i
'    ' ---
    
    ' --- Esporta i dati del vettore
    For i = 0 To intRows
        With mobjOutFile02
            strIDSocieta = varExportData(5, i)
            AVE.DS.Fields("0").FindValue = varExportData(4, i)
            AVE.Find
            .Item("3").value = AVE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
            .Item("5").value = varExportData(7, i)
            .Item("6").value = varExportData(8, i)
            .Item("7").value = varExportData(9, i)
            .Item("8").value = varExportData(10, i)
            .Item("9").value = varExportData(11, i)
            .Item("11").value = Left$(varExportData(7, i), 15)
            .Item("12").value = varExportData(12, i)
            If Not .WriteLine(True) Then
                ErrCode = OutputError
                GoTo ErrHandler
            End If
        End With
    Next i
    ' ---
    
    ' --- Esporto i dati dei veicoli
    For i = 0 To intRows
        With mobjOutFile04
            strIDSocieta = varExportData(5, i)
            ' --- Esporto il primo veicolo
            If Len(varExportData(13, i)) > 0 Then
                .Item("3").value = varExportData(13, i)
                ATM.DS.Fields("Codice").FindValue = varExportData(14, i)
                ATM.Find
                .Item("5").value = ATM.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                If Len(varExportData(15, i)) > 0 Then
                    ATS.DS.Fields("Codice").FindValue = varExportData(15, i)
                    ATS.Find
                    .Item("6").value = ATS.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                End If
                .Item("7").value = varExportData(16, i)
                .Item("8").value = varExportData(17, i)
                .Item("9").value = varExportData(18, i)
                .Item("10").value = varExportData(19, i)
                .Item("11").value = varExportData(20, i)
                For j = 12 To 30 Step 2
                    If Len(varExportData(j + 9, i)) > 0 Then
                        ' --- Numero scomparti da
                        .Item(CStr(j)).value = varExportData(j + 10, i)
                        ' --- Capacità in Kilolitri
                        .Item(CStr(j + 1)).value = varExportData(j + 9, i) / 1000
                    End If
                Next j
                .Item("32").value = varExportData(41, i)
                If Not .WriteLine(True) Then
                    ErrCode = OutputError
                    GoTo ErrHandler
                End If
            End If
            ' ---
            ' --- Esporto il secondo veicolo
            If Len(varExportData(42, i)) > 0 Then
                .Item("3").value = varExportData(42, i)
                ATM.DS.Fields("Codice").FindValue = varExportData(43, i)
                ATM.Find
                .Item("5").value = ATM.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                If Len(varExportData(44, i)) > 0 Then
                    ATS.DS.Fields("Codice").FindValue = varExportData(44, i)
                    ATS.Find
                    .Item("6").value = ATS.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                End If
                .Item("7").value = varExportData(45, i)
                .Item("8").value = varExportData(46, i)
                .Item("9").value = varExportData(47, i)
                .Item("10").value = varExportData(48, i)
                .Item("11").value = varExportData(49, i)
                For j = 12 To 30 Step 2
                    If Len(varExportData(j + 38, i)) > 0 Then
                        ' --- Numero scomparti da
                        .Item(CStr(j)).value = varExportData(j + 39, i)
                        ' --- Capacità in Kilolitri
                        .Item(CStr(j + 1)).value = varExportData(j + 38, i) / 1000
                    End If
                Next j
                .Item("32").value = varExportData(70, i)
                If Not .WriteLine(True) Then
                    ErrCode = OutputError
                    GoTo ErrHandler
                End If
            End If
            ' ---
        End With
    Next i
    ' ---
    
    ' --- Esporto i dati degli autisti
    For i = 0 To intRows
        With mobjOutFile05
            strIDSocieta = varExportData(5, i)
            ' --- Esporto il primo autista
            If Len(varExportData(71, i)) > 0 Then
                ' --- 13/09/99 Provvisorio: il codice autista può essere più lungo di 3
'                .Item("2").value = varExportData(71, i)
                .Item("2").value = "1" & varExportData(71, i)
                ' ---
                
                intSpacePos = InStr(4, varExportData(72, i), " ")
                .Item("3").value = Left$(varExportData(72, i), intSpacePos)
                .Item("4").value = Mid$(varExportData(72, i), intSpacePos + 1)
                AVE.DS.Fields("0").FindValue = varExportData(73, i)
                AVE.Find
                .Item("5").value = AVE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                AEE.DS.Fields("0").FindValue = varExportData(6, i)
                AEE.Find
                .Item("11").value = AEE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                If Not .WriteLine(True) Then
                    ErrCode = OutputError
                    GoTo ErrHandler
                End If
            End If
            ' ---
            
            ' --- Esporto il secondo autista, se presente e diverso dal primo autista
            If (Len(varExportData(74, i)) > 0) Then
                If (varExportData(74, i) <> varExportData(71, i)) Then
                    .Item("2").value = varExportData(74, i)
                    intSpacePos = InStr(4, varExportData(75, i), " ")
                    .Item("3").value = Left$(varExportData(75, i), intSpacePos)
                    .Item("4").value = Mid$(varExportData(75, i), intSpacePos + 1)
                    AVE.DS.Fields("0").FindValue = varExportData(76, i)
                    AVE.Find
                    .Item("5").value = AVE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                    AEE.DS.Fields("0").FindValue = varExportData(6, i)
                    AEE.Find
                    .Item("11").value = AEE.DS.MultiSoc.DS.GetFirmCode(strIDSocieta)
                    If Not .WriteLine(True) Then
                        ErrCode = OutputError
                        GoTo ErrHandler
                    End If
                End If
            End If
            ' ---
        End With
    Next i
    ' ---
    
    MsgBox "Esportazione terminata correttamente", vbOKOnly + vbInformation, "Elaborazione terminata"
    ExportData = True
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & mstrExportFileName
        Case OutputError
            strMessage = "Si è verificato un errore durante la scrittura del file"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "OIL2AUT0", _
                "ExportData", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    ExportData = False
EndProc:
    mobjOutFile01.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
End Function

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    .Fields(.KeyFields(1)).FindValue = mskField1
                    If DataClass.Find = True Then
                        If .HowMany > 0 Then
                            If .RefreshClient Then
                                DisplayMultiSoc
                            End If
                            SetState fsEdit
                        Else
                            If State = fsNone Then
                                SetState fsNone
                            End If
                        End If
                    End If
                End If
            End If
        End If
        

    End With
End Sub





Private Sub ClearForm()
    DataClass.DS.ClearClient
    ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim Respok, Respuesta As String
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    
    Select Case Cmd
        Case fcAddNew

            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Solo se pueden modificar estos datos desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
            MsgBox "Password Correcta, datos guardados", vbOKOnly
        
                If (State = fsEdit) Or (State = fsNone) Then
                    DataClass.DS.Fields.ResetFields True
                    ClearForm
                    SetState fsInsert
                End If
            End If
        Case fcSave

            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Solo se pueden modificar estos datos desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
            MsgBox "Password Correcta, datos guardados", vbOKOnly
               Select Case State
                   Case fsEdit
                       blnResult = UpdateRecord
                   Case fsInsert
                       blnResult = InsertRecord
                   Case Else
                       blnResult = False
               End Select
            End If
            If blnResult Then
                SetState fsNone
            End If
        Case fcCustom1
'            If gmstrCodInstallazione = ciInst2 Then
            #If (VersionID = 2) Then 'Or (VersionID = 5) Then
                With CommonDialog1
                    mstrExportFileName = gmRegistry.GetSetting("Options", "AUTExportFile", , True)
                    .fileName = mstrExportFileName
                    .DialogTitle = "Seleziona il file per l'esportazione... "
                    .InitDir = App.Path
                    .Filter = "Tutti i files (*.*)|*.*| "
                    .FilterIndex = 1
                    .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
                    .CancelError = True
                    On Error Resume Next
                    .ShowSave
                    If Err.Number <> cdlCancel Then
                        mstrExportFileName = .fileName
                        If ExportData Then
                        End If
                    End If
                End With
            #End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcPrint
            If (State = fsEdit) Or (State = fsNone) Then
                varReportsList(0, 0) = "Stampa elenco Unità di Trasporto"
                varReportsList(1, 0) = "oil2aut1.rpt"
                varReportsList(2, 0) = 1
                CRPrint varReportsList, "", Me
            End If
        Case fcFind
            DoSearch "AUT", Me, "0", 1
        Case fcExit
            Unload Me
    End Select
End Sub

Public Function GetRecord(strID As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    DisplayMultiSoc
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




Private Sub DisplayMultiSoc()
    RefreshUnboundGrid TDBGrid2, DataClass.DS.MultiSoc.DS.List, GridArray, _
                       DataClass.DS.MultiSoc.DS.HowMany, MAX_COLS
    
End Sub


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
    Dim C As Control
    
    'Scansione della finestra
    DataClass.DS.ScanClient 0
    
    If DataClass.DS.Fields.Validate Then
        DataClass.BeginSave blnConcurError
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
        Case fsNone
            If Not Loading Then
                mskField1.SetFocus
            End If
            ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
            Command6.Enabled = False
        Case fsEdit
            Command6.Enabled = True
            mskField2.SetFocus
        Case fsInsert
            Command6.Enabled = False
            mskField1.SetFocus
    End Select
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
    Set GridArray = New ArrayGridCls
    Set DataClass = New CUnitaDiTrasporto
    
'    Select Case gmstrCodInstallazione
'        Case ciInst1
    #If VersionID = 1 Then
            Command6.Visible = False
            Picture1(5).Visible = False
'        Case ciInst2
    #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
            Command6.Visible = True
            Picture1(5).Visible = True
            Set mobjOutFile01 = New CTextFile
            With mobjOutFile01
                .Add "Tipo Record", 4, "1", "MI01"
                .Add "Filler", 6, "2", Space$(6)
                .Add "Numero contratto", 8, "3"
                .Add "Numero contratto sec.", 8, "4"
                .Add "Note", 50, "5"
                .Add "Km a contratto", 6, "6", , fftNumerico
                .Add "Km percorsi", 44, "7", 0, fftNumerico
                .Add "Tipo contratto", 2, "8"
                .Add "Targa 1", 8, "9"
                .Add "Targa 2", 8, "10"
                .Add "Vettore", 3, "11"
                .Add "Unità Emittente", 3, "12", , fftNumerico
                .Add "Presidio Logistico", 3, "13", , fftNumerico
                .Add "Filler", 12, "14", Space$(12)
            End With
            Set mobjOutFile02 = New CTextFile
            With mobjOutFile02
                .Add "Tipo Record", 4, "1", "MI02"
                .Add "Filler", 6, "2", Space$(6)
                .Add "Cod. Vettore", 3, "3"
                .Add "Filler", 5, "4", Space$(5)
                .Add "Ragione Sociale", 30, "5"
                .Add "Indirizzo", 30, "6"
                .Add "Località", 30, "7"
                .Add "Provincia", 2, "8"
                .Add "CAP", 5, "9", , fftNumerico
                .Add "Cod. Nazione", 3, "10", 1, fftNumerico
                .Add "Rag. Sociale abbrebiata", 15, "11"
                .Add "Telefono", 15, "12"
                .Add "Fax", 15, "13"
                .Add "Filler", 2, "14", Space$(2)
            End With
            Set mobjOutFile04 = New CTextFile
            With mobjOutFile04
                .Add "Tipo Record", 4, "1", "MI04"
                .Add "Filler", 6, "2", Space$(6)
                .Add "Targa", 8, "3"
                .Add "Filler", 2, "4", Space$(2)
                .Add "Tipo veicolo", 3, "5"
                .Add "Tipo cisterna", 3, "6"
                .Add "Portata lorda", 6, "7", , fftNumerico
                .Add "Tara pesata", 6, "8", , fftNumerico
                .Add "Tara libretto", 6, "9", , fftNumerico
                .Add "Portata rimorchiabile", 6, "10", , fftNumerico
                .Add "Cubatura costruttore", 6, "11", , fftNumerico
                .Add "Scomparto 01", 2, "12", , fftNumerico
                .Add "Portata 01", 2, "13", , fftNumerico
                .Add "Scomparto 02", 2, "14", , fftNumerico
                .Add "Portata 02", 2, "15", , fftNumerico
                .Add "Scomparto 03", 2, "16", , fftNumerico
                .Add "Portata 03", 2, "17", , fftNumerico
                .Add "Scomparto 04", 2, "18", , fftNumerico
                .Add "Portata 04", 2, "19", , fftNumerico
                .Add "Scomparto 05", 2, "20", , fftNumerico
                .Add "Portata 05", 2, "21", , fftNumerico
                .Add "Scomparto 06", 2, "22", , fftNumerico
                .Add "Portata 06", 2, "23", , fftNumerico
                .Add "Scomparto 07", 2, "24", , fftNumerico
                .Add "Portata 07", 2, "25", , fftNumerico
                .Add "Scomparto 08", 2, "26", , fftNumerico
                .Add "Portata 08", 2, "27", , fftNumerico
                .Add "Scomparto 09", 2, "28", , fftNumerico
                .Add "Portata 09", 2, "29", , fftNumerico
                .Add "Scomparto 10", 2, "30", , fftNumerico
                .Add "Portata 10", 2, "31", , fftNumerico
                .Add "Carica", 1, "32"
                .Add "Filler", 68, "33", Space$(68)
            End With
            Set mobjOutFile05 = New CTextFile
            With mobjOutFile05
                .Add "Tipo Record", 4, "1", "MI05"
                .Add "Cod. Autista", 6, "2", , fftNumerico
                .Add "Cognome", 20, "3"
                .Add "Nome", 20, "4", Space$(5)
                .Add "Cod. Vettore", 3, "5"
                .Add "Disponibilità", 2, "6", "SI"
                .Add "Dal", 8, "7", , fftNumerico
                .Add "Al", 8, "8", , fftNumerico
                .Add "Telefono", 15, "9", , fftNumerico
                .Add "Unità Emittente", 3, "10", , fftNumerico
                .Add "Presidio logistico", 3, "11"
                .Add "Sospensione", 68, "12", Space$(68)
                .Add "Deposito OL", 3, "13", , fftNumerico
                .Add "Filler", 2, "14", Space$(2)
            End With
'    End Select
        #End If
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
        Set .DA = DataConnect.DA
        Set Soc = New CSocieta
        Set Soc.DS.DA = DataConnect.DA
        Soc.GetList True
        .ConfigureInterface
        ConfigureObjects Me
        'Set .MultiSoc.DS.DA = DA
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set ATN = New CTipoContratto
    With ATN.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AEE = New CEnteEmittente
    With AEE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ATM = New CTipoAutomezzo
    With ATM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ATS = New CTipoCisterna
    With ATS.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AVE = New CVettore
    With AVE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ASC = New CSocieta
    With ASC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    SetValueItemsSocieta TDBGrid2.Columns(0), ASC, True
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Select Case gmstrCodInstallazione
'        Case ciInst2
    #If (VersionID = 2) Then 'Or (VersionID = 5) Then
            gmRegistry.SaveSetting "Options", "AUTExportFile", mstrExportFileName, True
'    End Select
    #End If
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set ATN = Nothing
    Set ATS = Nothing
    Set ATM = Nothing
    Set AVE = Nothing
    Set AEE = Nothing
    Set OIL2AUT0 = Nothing
End Sub

Private Sub SetValueItemsSocieta(Col As Column, Entity As Object, _
                         Optional Translate As Boolean = True)
    Dim vItem As ValueItem
    Dim DS As CDataStruct
    Dim i As Long
    Dim intRows As Integer
    
    With Col.ValueItems
        .Clear
        .Presentation = dbgSortedComboBox
        .Translate = Translate
        .Validate = True
        Set DS = Entity.DS
        Entity.GetList True
        If DS.HowMany > 0 Then
            .MaxComboItems = 20
            intRows = UBound(DS.List, 2)
            For i = 0 To intRows 'UBound(DS.List, 2)
                 Set vItem = New ValueItem
                 vItem.value = DS.ListItem(0, i)
                 vItem.DisplayValue = DS.ListItem(1, i) & " (" & DS.ListItem(0, i) & ")"
                .Add vItem
            Next i
        End If
    End With

End Sub

Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub



