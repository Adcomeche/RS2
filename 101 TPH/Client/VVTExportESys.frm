VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form VVTExportESys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRASMISSIONE VIAGGI"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9270
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      ForeColor       =   &H00FF0000&
      Height          =   6435
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox chkFirstTime 
         Caption         =   "Solo Viaggi non trasmessi"
         Height          =   255
         Left            =   6600
         TabIndex        =   51
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Tag             =   "33"
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileName2 
         Height          =   255
         Left            =   6600
         TabIndex        =   23
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame5 
         Height          =   1995
         Left            =   120
         TabIndex        =   26
         Top             =   2700
         Width           =   9015
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "VVTExportESys.frx":0000
            TabIndex        =   27
            Tag             =   "SELECT"
            Top             =   240
            Width           =   8775
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   5400
         TabIndex        =   30
         Top             =   4680
         Width           =   3735
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   240
            Picture         =   "VVTExportESys.frx":799E
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Stampa elenco viaggi"
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            Caption         =   "Viaggi trasmessi correttamente:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   35
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   34
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   33
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            Caption         =   "Viaggi da trasmettere:"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   4680
         Width           =   5175
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1335
            Left            =   120
            OleObjectBlob   =   "VVTExportESys.frx":7AA0
            TabIndex        =   29
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3960
         Picture         =   "VVTExportESys.frx":A4A8
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "CAL"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7920
         Picture         =   "VVTExportESys.frx":A5AA
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "CAL"
         Top             =   240
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Tag             =   "5"
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
      Begin MSMask.MaskEdBox mskDateTo 
         Height          =   255
         Left            =   6600
         TabIndex        =   4
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Tag             =   "11"
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   6600
         TabIndex        =   9
         Tag             =   "12"
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8880
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Tag             =   "32"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":A6AC
               Key             =   "New"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":A7BE
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":A8D0
               Key             =   "Undo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":A9E2
               Key             =   "Find"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":AAF4
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":AC06
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":AD18
               Key             =   "Copy"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":AE2A
               Key             =   "Cut"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":AF3C
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":B04E
               Key             =   "Prop"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":B160
               Key             =   "Paste"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":B272
               Key             =   "Camera"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":B384
               Key             =   "Preview"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Tag             =   "3"
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
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
               Picture         =   "VVTExportESys.frx":B8C6
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":B9D8
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "VVTExportESys.frx":BAEA
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Tag             =   "13"
         Top             =   960
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
         Left            =   1080
         TabIndex        =   17
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
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Tag             =   "92"
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
         Left            =   2280
         TabIndex        =   55
         Tag             =   "92"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Raggr."
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   54
         Top             =   2400
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   5040
         Tag             =   "92"
         Top             =   2400
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
         Left            =   2400
         TabIndex        =   37
         Tag             =   "33"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   4800
         Tag             =   "34"
         Top             =   1320
         Width           =   240
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
         Left            =   1800
         TabIndex        =   50
         Tag             =   "34"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Autista"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   16
         Top             =   1320
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
         Left            =   6360
         TabIndex        =   49
         Tag             =   "13"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&U.d.T."
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   8880
         Tag             =   "13"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File Dettaglio:"
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   48
         Top             =   1680
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   4440
         Tag             =   "5"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   8400
         Tag             =   "6"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   7920
         Tag             =   "12"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   3960
         Tag             =   "11"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   4800
         Tag             =   "32"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   5160
         Tag             =   "33"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   5160
         Tag             =   "3"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Soc. Comp. Viaggio"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
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
         Left            =   2400
         TabIndex        =   38
         Tag             =   "3"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Soc. Dest. Trasm."
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vettore"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   25
         Top             =   960
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
         Left            =   1800
         TabIndex        =   24
         Tag             =   "32"
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al Num. Viaggio Def."
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal Num. Viaggio Def."
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data"
         Height          =   255
         Index           =   12
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File Testata:"
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   39
      Top             =   6480
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "VVTExportESys.frx":BBFC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   2160
         Picture         =   "VVTExportESys.frx":BCFE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   42
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
         Index           =   6
         Left            =   4080
         Picture         =   "VVTExportESys.frx":BE00
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   40
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
         Left            =   7560
         Picture         =   "VVTExportESys.frx":BF02
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         Picture         =   "VVTExportESys.frx":C004
         TabIndex        =   41
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Viaggi"
         Height          =   375
         Left            =   2040
         Picture         =   "VVTExportESys.frx":C106
         TabIndex        =   43
         Tag             =   "BROWSE"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "VVTExportESys.frx":C208
         TabIndex        =   47
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Seleziona File..."
         Height          =   375
         Left            =   120
         Picture         =   "VVTExportESys.frx":C30A
         TabIndex        =   45
         Tag             =   "OPEN"
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   7395
      Width           =   9270
      _ExtentX        =   16351
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
            Object.Width           =   10901
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11.05"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "07/07/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "VVTExportESys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 30

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mstrProgrInvio As String

Public GridArraySelect As ArrayGridCls
Public XA As XArrayDB
Public Loading As Boolean

Public DataClass As CViaggioSelezione
Public State As FormState

Private mvarFunction As FunctionCodes
Private VVT As CViaggio
Private ExViagDetail As CTextFile
Private ExViagProd As CTextFile
Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property
Private Sub DisplaySelection()
    Dim varLog As Variant
    Dim i As Long
    Dim lngRows As Long
    Dim blnFlagTrasm As Boolean
    
    lngRows = DataClass.DS.HowMany
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
    If lngRows > 0 Then
        ReDim varLog(0 To 1, 0 To lngRows - 1)
        For i = 0 To lngRows - 1
            varLog(0, i) = DataClass.DS.ListItem(0, i)
            If DataClass.DS.ListItem(9, i) = "T" Then
                blnFlagTrasm = True
            End If
        Next i
        RefreshXArray XA, tdbLog, varLog
    End If
        If blnFlagTrasm Then
            DataClass.DS.DisplayMessage "sono stati selezionati alcuni viaggi già trasmessi", 1
        End If
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
                cmdExport.Enabled = True

            End If
        Case fcExport
            ExportRecords
            cmdExport.Enabled = False

        Case fcOpen
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per la trasmissione... "
                .InitDir = App.Path
                .Filter = "Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                .ShowSave
                mskFileName = .fileName
            End With
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(tdbLog)
                    '.GetRecord GetSelectedID(Me.TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
    End If
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    
    Dim blnFlagFirstTime As Boolean
    
    Me.MousePointer = vbHourglass
    DataClass.DS.DisplayMessage vbNullString, 0
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            blnFlagFirstTime = (chkFirstTime.value = vbChecked)
            If Not DataClass.GetSelectionExportExViagDetail(blnFlagFirstTime) Then
                GoTo ErrHandler
            End If
            DisplaySelection
            GoTo EndProc
        End If
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    GoTo EndProc
ErrHandler:
    Me.MousePointer = vbDefault
    SelectRecords = False
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Dalla data inizio viaggio"
        Case 1
            Set M = mskDateTo
            strTitle = "Alla data inizio viaggio"
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub cmdPrint_Click()
    TDBGrid1.PrintInfo.PageHeader = "Trasmissione del " & Format$(Now, "dd/mm/yyyy")
    TDBGrid1.PrintInfo.SettingsOrientation = 2
    TDBGrid1.PrintInfo.PrintData
End Sub


Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim D As Variant
    
    Loading = True
    Set GridArraySelect = New ArrayGridCls
    Set XA = New XArrayDB
    
    cmdBrowse.Enabled = True
    mskFileName = gmRegistry.GetSetting("Options", "TripExportFileEsys", , True)
    mskFileName2 = gmRegistry.GetSetting("Options", "TripExportFileEsysDett", , True)
    MaskEdBox1 = gmRegistry.GetSetting("Options", "TripExportVetViag", , True)
    MaskEdBox3 = gmRegistry.GetSetting("Options", "TripExportASCViag", , True)
    MaskEdBox2 = gmRegistry.GetSetting("Options", "TripExportASCExp", , True)
    
    
'    D = Format$(Date, "dd/mm/yy")
'    mskDateFrom.SelText = D
'    mskDateTo.SelText = D
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .Fields("33").Required = True
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    If IsUserGrandesCuentas Then
        MaskEdBox10 = "GC"
    Else
        MaskEdBox10 = "RP"
    End If
    DataClass.DS.ValidateClientField MaskEdBox10.Tag, MaskEdBox10.Text
    FieldUpdateDescription MaskEdBox10.Tag, Me, MaskEdBox10
    FieldSetDescription MaskEdBox10.Tag, Me, MaskEdBox10.Text, MaskEdBox10
    
    Set ExViagDetail = New CTextFile
    With ExViagDetail
        .Add "Anno viaggio", 4, "1", , fftNumerico
        .Add "Filiale viaggio", 2, "2", "01", fftAlfa
        .Add "Key viaggio", 7, "3", , fftNumerico
        .Add "Ctr viaggio", 3, "4", , fftNumerico
        .Add "Fil. cliente", 2, "5", "01", fftAlfa
        .Add "Codice cliente", 6, "6", , fftNumerico
        .Add "Contratto cliente", 3, "7", , fftNumerico
        .Add "Codice partenza", 5, "8", , fftAlfa
        .Add "Località partenza", 30, "9", , fftAlfa
        .Add "Codice destino", 5, "10", , fftAlfa
        .Add "Località destino", 30, "11", , fftAlfa
        .Add "Fil. vettore", 2, "12", "01", fftAlfa
        .Add "Codice vettore", 6, "13", , fftNumerico
        .Add "Contratto vettore", 3, "14", , fftNumerico
        .Add "Data servizio", 8, "15", , fftNumerico
        .Add "Trattore", 10, "16", , fftAlfa
        .Add "Rimorchio", 10, "17", , fftAlfa
        .Add "Tipo merce", 3, "18", , fftAlfa
        .Add "Maggiorazione sabato cliente", 1, "19", "S", fftAlfa
        .Add "Maggiorazione festivo cliente", 1, "20", "F", fftAlfa
        .Add "Km tachigrafo", 6, "21", "0", fftNumerico
        .Add "Km viaggio", 6, "22", , fftNumerico
        .Add "Peso", 7, "23", , fftNumerico
        .Add "N° ddt", 15, "24", , fftAlfa
        .Add "Data ddt", 8, "25", , fftNumerico
        .Add "Unità viaggiante", 4, "26", , fftAlfa
        .Add "Maggiorazione bomba vettore", 1, "27", " ", fftAlfa
        .Add "Maggiorazione festivo vettore", 1, "28", "F", fftAlfa
        .Add "Sconto ads", 1, "29", "1", fftAlfa
        .Add "Consegne", 2, "30", , fftNumerico
        .Add "Filiale deposito", 2, "31", "01"
        .Add "Codice deposito", 6, "32", , fftNumerico
        .Add "Progressivo", 3, "33", "0", fftNumerico
        .Add "Filiale destinatario", 2, "34", "01"
        .Add "Codice destinatario", 6, "35", , fftNumerico
        .Add "Progressivo", 3, "36", "0", fftNumerico
        .Add "Numero ordine", 15, "37"
        .Add "Numero ordine di carico", 15, "38"
        .Add "Area AGIP", 3, "39"
        .Add "Terzo riferimento", 15, "40"
        .Add "Tipo servizio", 3, "41"
        .Add "s/compte", 30, "42"
        .Add "Mittente", 10, "43"
        .Add "Destinazione", 10, "44"
        .Add "Raggruppamento", 3, "45"
        ' (PM) - Km Passivi
        .Add "Km passivi", 6, "46", , fftNumerico
    End With
    
    Set ExViagProd = New CTextFile
    With ExViagProd
        .Add "Anno viaggio", 4, "1", , fftNumerico
        .Add "Filiale viaggio", 2, "2", "01", fftAlfa
        .Add "Key viaggio", 7, "3", , fftNumerico
        .Add "Ctr viaggio", 3, "4", , fftNumerico
        .Add "Progressivo", 3, "5", "0", fftNumerico
        .Add "Tipo prodotto", 3, "6"
        .Add "Quantità", 6, "7", , fftNumerico
    End With
    
    Loading = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim strFileName As String
    Dim strFileName2 As String
    
    strFileName = mskFileName ' Left$(mskFileName, InStr(1, mskFileName, "."))
    strFileName2 = mskFileName2
    gmRegistry.SaveSetting "Options", "TripExportFileESys", strFileName, True
    gmRegistry.SaveSetting "Options", "TripExportFileESysDett", strFileName2, True
   
    gmRegistry.SaveSetting "Options", "TripExportVetViag", MaskEdBox1.Text, True
    gmRegistry.SaveSetting "Options", "TripExportASCViag", MaskEdBox3.Text, True
    gmRegistry.SaveSetting "Options", "TripExportASCExp", MaskEdBox2, True
    
    Set DataClass = Nothing
    Set ExViagDetail = Nothing
    Set ExViagProd = Nothing
    Set VVT = Nothing
    Set XA = Nothing
    Set VVTExportESys = Nothing
End Sub

Private Sub ExportRecords()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim i As Long
    Dim j As Long
    Dim intFileNum As Integer
    Dim strIDViaggio As String
    Dim strIDViaggioDef As String
    Dim strFileName As String
    Dim strFileName2 As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim StringWarning As String
    Dim ErrDescr As Variant
    Dim ErrCode As Errors
    Dim strErrorMessage As String
    Dim blnFound As Boolean
    Dim C As CError
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    Dim intFileMode As Integer
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    If (Len(Dir(mskFileName)) > 0) Then
        If MsgBox("Almeno un file di trasmissione contiene dati. Si desidera sovrascrivere i files?" & vbCrLf & _
                  "Premendo Sì il contenuto attuale dei files verrà eliminato." & vbCrLf & _
                  "Premendo No i dati verranno accodati ai dati esistenti", _
                  vbYesNo + vbInformation, _
                  "Attenzione") = vbYes Then
            intFileMode = 1
        Else
            intFileMode = 2
        End If
    Else
        intFileMode = 1
    End If
    strFileName = mskFileName
    If Not ExViagDetail.OpenFile(strFileName, intFileMode) Then
        ErrCode = FileNotFound
        strInfo = "Nome del file: " & strFileName
        GoTo ErrHandler
    End If
    strFileName2 = mskFileName2
    If Not ExViagProd.OpenFile(strFileName2, intFileMode) Then
        ErrCode = FileNotFound
        strInfo = "Nome del file: " & strFileName2
        GoTo ErrHandler
    End If
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione dei documenti
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        With VVT.DS
            .Fields.ResetFields
            .Fields(1).FindValue = GridArraySelect.value(lngRow, 0)
            If Not VVT.Find Then
                GoTo ErrHandler
            End If
        End With
        StringWarning = ""
        If CreateRowsExViagg(VVT, MaskEdBox2.Text, ErrDescr) Then
'            If Not VVT.SetFlagTrasmissioneESys("T") Then
'                intRowError = NoSetState
'                GoTo LogResult
'            End If
        End If
            
LogResult:
        
        Select Case ErrDescr
          Case ""
            strErrorMessage = "OK"
            lngCountOK = lngCountOK + 1
          Case Else
            strErrorMessage = ErrDescr
            lngCountFail = lngCountFail + 1
        End Select
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        ErrDescr = ""
        XA.value(lngRow, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
        ' ---
        DoEvents
    Next lngRow
    GoTo EndProc

ErrHandler:

    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
        Case OutputError
            strMessage = "Impossibile Trasferire i Documenti"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
                    
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "VVTExport2.frm", _
                "ExportRecord", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    
EndProc:
    ExViagDetail.CloseFile
    ExViagProd.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub
Public Function CreateRowsExViagg(VVT As CViaggio, _
                                  IdSoc As String, _
                                  Optional ByRef ErrDesc As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim strError As String
    Dim tempStr As String
    Dim intNumConsegne As Integer
    Dim strIDViaggio  As String
    Dim strIDViaggioDef As String
    Dim strIDSocietàEstrap As String
    Dim blnCons As Boolean
    Dim blnConsegnaAnnullata As Boolean
    Dim i As Integer
    Dim j As Long
    Dim k As Long
    Dim lngProgConsegna As Long
    Dim lngProgProdotti As Long
    Dim dblPeso As Double
    Dim blnResult As Boolean
    Dim blnConcErr As Boolean
    Dim varBuffer As Variant
    Dim varIDConsegna As Variant
    Dim varIDConsegnaPrec As Variant
    Dim varIDVettoreSoc As Variant
    Dim varBufferDistanza As Variant
    Dim lngHowMany As Long
    
    Dim VVC As CViaggioConsegna
    Dim AXX As CParametro
    Dim ACN As CContratto
    Dim ACL As CCliente
    Dim ADE As CDeposito
    Dim ADZ As CDestinazione
    Dim AVE As CVettore
    Dim AUT As CUnitaDiTrasporto
    Dim APR As CProdotto
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
        .Fields("0").FindValue = "001"
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            GoTo ErrHandler
        End If
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set ACN = New CContratto
    With ACN.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set AVE = New CVettore
    With AVE.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
    Set APR = New CProdotto
    With APR.DS
        Set .DA = VVT.DS.DA
        Set .ErrH = VVT.DS.ErrH
    End With
         
    strIDViaggio = VVT.DS.Fields("0").value
    strIDViaggioDef = VVT.DS.Fields("1").value
    strIDSocietàEstrap = IdSoc
    varIDConsegnaPrec = "ZZZZ"
    
    varBuffer = VVT.GetExportEsysData(strIDSocietàEstrap)
'    blnResult = VVT.DS.DA.QueryCommand(SQL_GET_EXVIAGDETAIL, varBuffer, lngHowMany, _
'                                   strIDViaggio, strIDSocietàEstrap, strIDSocietàEstrap, _
'                                   strIDSocietàEstrap)
    If IsEmpty(varBuffer) Then
'    If Not blnResult Then
        strError = "Si è verificato un errore durante la lettura del Viaggio. "
        GoTo ErrHandler
    End If
'    If lngHowMany > 0 Then
    For j = 0 To UBound(varBuffer, 2)
        If varBuffer(27, j) <> varIDConsegnaPrec Then
            blnConsegnaAnnullata = False
            If (VVT.DS.Fields("44").value <> 0) Or (varBuffer(29, j) <> 0) Then
                blnConsegnaAnnullata = True
            End If
            varIDConsegnaPrec = varBuffer(27, j)
            lngProgConsegna = lngProgConsegna + 1
            lngProgProdotti = 1
            With ExViagDetail
                .ClearFields
                .Item("1").value = Year(varBuffer(10, j))   ' Anno Viaggio
    '            .Item("2").value = varBuffer(8, 0)         ' Filiale Viaggio ' Cod Per Soc Ente Emit
                .Item("3").value = strIDViaggio             ' Numero Provvisorio
                .Item("4").value = lngProgConsegna          ' Ctr viaggio = Progressivo consegna
                
                ACL.DS.Fields.ResetFields                   ' Codice Cliente
                ACL.DS.Fields("0").FindValue = varBuffer(24, j)
                ACL.Find
                .Item("6").value = ACL.DS.MultiSoc.DS.GetFirmCode(IdSoc)
                If .Item("6").value = vbNullString Then
                    strError = "Codice Cliente per Società non trovato"
                    GoTo ErrHandler
                End If
                
'                If IsNull(Len(varBuffer(25, j))) Or (Len(varBuffer(25, j)) = 0) Then           ' Contratto Cliente
'                    strError = "Contratto Cliente non trovato"
'                    GoTo ErrHandler
'                End If
                .Item("7").value = varBuffer(25, j)
                
                ADE.DS.Fields.ResetFields                   ' Codice Partenza = Codice Dep x Soc
                ADE.DS.Fields("0").FindValue = varBuffer(7, j)
                ADE.Find
                .Item("8").value = ADE.DS.MultiSoc.DS.GetFirmCode(IdSoc)
                If .Item("8").value = vbNullString Then
                    strError = "Codice Deposito per Società non trovato"
                    GoTo ErrHandler
                End If
                
                .Item("9").value = varBuffer(13, 0)         ' Localita Partenza = Localita Deposito
                
                ADZ.DS.Fields.ResetFields                   ' Codice destino = Codice Destinazione x Soc
                ADZ.DS.Fields("0").FindValue = varBuffer(17, j)
                ADZ.Find
                .Item("10").value = ADZ.DS.MultiSoc.DS.GetFirmCode(IdSoc)
'                If .Item("10").value = vbNullString Then
'                    strError = "Codice Destinazione per Società non trovato"
'                    GoTo ErrHandler
'                End If
                
                .Item("11").value = varBuffer(26, j)        ' Località destino
                
                AVE.DS.Fields.ResetFields                   ' Codice / Contratto Vettore
                AVE.DS.Fields("0").FindValue = varBuffer(3, j)
                AVE.Find
                varIDVettoreSoc = AVE.DS.MultiSoc.DS.GetFirmCode(IdSoc)
                If varIDVettoreSoc <> "900000" Then
                    ' Vettore terzo
                    .Item("13").value = varIDVettoreSoc
                    If Len(.Item("13").value) = 0 Then
                        strError = "Codice Vettore non trovato"
                        GoTo ErrHandler
                    End If
                    .Item("14").value = varBuffer(28, j)
'                    If Len(.Item("14").value) = 0 Then
'                        strError = "Contratto Vettore non trovato"
'                        GoTo ErrHandler
'                    End If
                Else
                    ' Vettore Logimovil
                    .Item("13").value = varBuffer(28, j)
                    If Len(.Item("13").value) = 0 Then
                        strError = "Codice Autista non trovato"
                        GoTo ErrHandler
                    End If
                    .Item("14").value = "0"
                End If
'                If .Item("13").value = vbNullString Then
'                    strError = "Codice Vettore per Società non trovato"
'                    GoTo ErrHandler
'                End If
                
'                .Item("14").value = AVE.DS.Fields("14").value                   ' Contratto vettore
                
                .Item("15").value = Format(varBuffer(10, j), "ddmmyyyy")        ' Data Servizio = Data Inizio Viaggio

                
                ' (PM) 17/03/06 - Passiamo sempre le targhe Motrice e Rimorchio
                ' -------------
                .Item("16").value = varBuffer(4, j)         ' Targa Motrice
                .Item("17").value = varBuffer(5, j)         ' Targa Rimorchio
'                If (Len(varBuffer(15, j)) = 0) Or (IsNull(varBuffer(15, j))) Then
'                    .Item("16").value = varBuffer(4, j)         ' Targa Motrice
'                Else
'                    .Item("16").value = varBuffer(15, j)
'                End If
'                If (Len(varBuffer(16, j)) = 0) Or (IsNull(varBuffer(16, j))) Then
'                    .Item("17").value = varBuffer(5, j)         ' Targa Rimorchio
'                Else
'                    .Item("17").value = varBuffer(16, j)
'                End If
                ' -------------
                
                ' Nel caso di Agip Reparto preleva i km della consegna
                If varBuffer(2, j) = ID_SOC_AGIP And Not IsNull(varBuffer(30, j)) And _
                    Not IsEmpty(varBuffer(30, j)) And varBuffer(30, j) = "RP" Then
                    .Item("21").value = varBuffer(32, j)
                Else
                    ' Altrimenti preleva i km dal deposito di carico alla destinazione
                    varBufferDistanza = ADZ.GetDistanza(CVar(varBuffer(17, j)), CVar(varBuffer(7, j)))
    
                    If IsEmpty(varBufferDistanza) Then
                        strError = "Km Deposito di carico per destinazione " & varBuffer(17, j) & " non trovati."
                        GoTo ErrHandler
                    End If
                    .Item("21").value = varBufferDistanza(0, 0)           ' KM attivi
                    .Item("46").value = varBufferDistanza(1, 0)           ' KM passivi
                End If
                
                ' *** GR 01/03/02 - Passiamo i kilometri sempre a zero in modo che vengano usati
                ' ***               i kilometri Eurosystem
                ' *** PM 22/03/06 - Cambio di rotta, adesso passiamo i km dal deposito alla destinazione
'                .Item("22").value = 0

                ' Se il deposito è CEPSA, controlla se trasmettere i km dal deposito o quelli del viaggio
'                If ADE.DS.Field("Società").value = "038" And _
'                    Not IsNull(ADE.DS.Field("Trasmissione KM CEPSA").value) And _
'                    Not IsEmpty(ADE.DS.Field("Trasmissione KM CEPSA").value) And _
'                    ADE.DS.Field("Trasmissione KM CEPSA").value = "V" Then
'
'                    ' Invia KM del viaggio
'                    .Item("22").value = varBuffer(6, j)
'                Else
'                    varBufferDistanza = ADZ.GetDistanza(CVar(varBuffer(17, j)), CVar(varBuffer(7, j)))
'
'                    If IsEmpty(varBufferDistanza) Then
'                        strError = "Km Deposito di carico per destinazione " & varBuffer(17, j) & " non trovati."
'                        GoTo ErrHandler
'                    End If
'                    .Item("22").value = varBufferDistanza(0, 0)           ' KM consegna
'                End If
                .Item("22").value = varBuffer(6, j)
                ' ***
                
                VVC.DS.Fields("0").FindValue = strIDViaggio     ' Volume
                blnCons = VVC.GetList(True)
                If blnCons Then
                    If Not IsEmpty(VVC.DS.List) Then
                        intNumConsegne = UBound(VVC.DS.List, 2) + 1
                        For i = 0 To UBound(VVC.DS.List, 2)
                            VVC.DS.Fields("0").FindValue = strIDViaggio
                            VVC.DS.Fields("1").FindValue = VVC.DS.List(0, i)
                            If VVC.Find Then
                                If VVC.DS.HowMany > 0 Then
                                    VVC.GetTotali
                                    If IsNull(VVC.TotaliQta(1)) Then
                                    Else
                                        dblPeso = dblPeso + VVC.TotaliQta(1)
                                    End If
                                End If
                                If VVC.DS.List(0, i) = varBuffer(27, j) Then
                                    VVC.DS.Field("Flag stato 1").NewValue = lngProgConsegna
                                    VVC.DS.Field("Flag stato 1").Updated = True
                                    VVC.BeginSave blnConcErr, False
                                    VVC.DS.DA.EndTrans ETCOMMIT
                                End If
                            End If
                        Next i
                    End If
                Else
                    strError = "Impossibile Leggere Le Consegne "
                End If
                If dblPeso > 0 Then
                    dblPeso = (dblPeso / 1000)
                End If
                .Item("23").value = dblPeso * 100
                
                k = j
                Do While varBuffer(27, j) = varBuffer(27, k)                ' N ddt
                    If Len(varBuffer(21, k)) > 0 Then
                        .Item("24").value = varBuffer(21, k)
                    End If
                    k = k + 1
                    If k > UBound(varBuffer, 2) Then Exit Do
                Loop
'                .Item("24").value = varBuffer(1, j)
                
                .Item("25").value = Format(varBuffer(10, j), "ddmmyyyy")    ' data ddt = Data inizio Viaggio
                
                AUT.DS.Fields.ResetFields                                   ' Codice Unità di Trasporto
                AUT.DS.Fields("0").FindValue = varBuffer(9, j)
                AUT.Find
                .Item("26").value = AUT.DS.MultiSoc.DS.GetFirmCode(IdSoc)
                If .Item("26").value = vbNullString Then
                    strError = "Codice UdT per Società non trovato"
                    GoTo ErrHandler
                End If
                .Item("30").value = intNumConsegne                          ' Num Consegne
                
                .Item("32").value = .Item("8").value                        ' Codice deposito
'                .Item("33").value = lngProgConsegna                         ' Progressivo
                .Item("35").value = .Item("10").value                       ' Codice destinatario
'                .Item("36").value = lngProgConsegna                         ' Progressivo
                .Item("37").value = varBuffer(18, j)                        ' Rif. esterno ordine
                .Item("38").value = varBuffer(19, j)                        ' Numero ordine di carico
                .Item("39").value = varBuffer(20, j)                        ' Codice Zona AGIP
                .Item("40").value = .Item("24").value                       ' 3° Riferimento
                If blnConsegnaAnnullata Then
                    .Item("41").value = "AAA"                               ' Tipo servizio
                End If
                .Item("42").value = varBuffer(17, j)                        ' Codice destinazione
                .Item("43").value = varBuffer(7, j)
                .Item("44").value = Trim(varBuffer(17, j))

                ' (PM) 21/04/06 - Supplemento Bomba
                .Item("27").value = IIf(varBuffer(31, j) = "B" Or varBuffer(31, j) = "FB", "1", "0")
                
                ' (PM) 21/04/06 - Raggruppamento (GC o RP)
                If IsNull(varBuffer(30, j)) Or IsEmpty(varBuffer(30, j)) Then
                    .Item("45").value = ""
                Else
                    .Item("45").value = varBuffer(30, j)
                End If

                .WriteLine
            End With
        End If
        With ExViagProd
            .ClearFields
            .Item("1").value = Year(varBuffer(10, j))   ' Anno Viaggio
'            .Item("2").value = varBuffer(8, 0)         ' Filiale Viaggio ' Cod Per Soc Ente Emit
            .Item("3").value = strIDViaggio             ' Numero Provvisorio
            .Item("4").value = lngProgConsegna          ' Ctr viaggio
'            .Item("5").value = lngProgProdotti          ' Progressivo
            
            APR.DS.Fields.ResetFields                   ' Codice Prodotto
            APR.DS.Fields("0").FindValue = varBuffer(22, j)
            APR.Find
            .Item("6").value = APR.DS.MultiSoc.DS.GetFirmCode(IdSoc, 3)
            If .Item("6").value = vbNullString Then
                strError = "Codice Prodotto per Società non trovato"
                GoTo ErrHandler
            End If
            .Item("7").value = varBuffer(23, j)         ' Quantità consegnata
            .WriteLine
            lngProgProdotti = lngProgProdotti + 1
        End With
    Next j
'    Else
'        strError = "Impossibile Leggere i Codici Per Società. "
'        GoTo ErrHandler
'    End If
    CreateRowsExViagg = True
    GoTo EndFunction
    
    
    
ErrHandler:

    If Not IsMissing(ErrDesc) Then
        If Len(strError) > 0 Then
            ErrDesc = strError & vbCrLf & _
                  "Viaggio Definitivo: " & strIDViaggioDef & vbCrLf
                  
        Else
            ErrDesc = "Errore Sconosciuto" & vbCrLf
        End If
    End If
    CreateRowsExViagg = False
    
EndFunction:
    
    Set ACN = Nothing
    Set VVC = Nothing
    
End Function






