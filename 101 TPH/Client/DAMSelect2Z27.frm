VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form DAMSelect2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GESTIONE FATTURE"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10770
   Begin VB.Frame Frame3 
      Caption         =   " Addebiti "
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   0
      TabIndex        =   39
      Top             =   3360
      Width           =   10755
      Begin VB.CommandButton cmdBrowseExport 
         Height          =   300
         Index           =   0
         Left            =   4980
         Picture         =   "DAMSelect2Z27.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1620
         Width           =   375
      End
      Begin VB.TextBox txtFileNameAddebiti 
         Height          =   285
         Left            =   1260
         TabIndex        =   9
         Top             =   1620
         Width           =   3675
      End
      Begin VB.TextBox txtFileNameViag 
         Height          =   285
         Left            =   1260
         TabIndex        =   11
         Top             =   1980
         Width           =   3675
      End
      Begin VB.CommandButton cmdBrowseExport 
         Height          =   300
         Index           =   1
         Left            =   4980
         Picture         =   "DAMSelect2Z27.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1980
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseExport 
         Height          =   300
         Index           =   2
         Left            =   10260
         Picture         =   "DAMSelect2Z27.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox txtFileNameViagProd 
         Height          =   285
         Left            =   6540
         TabIndex        =   14
         Top             =   1980
         Width           =   3675
      End
      Begin VB.CheckBox ckExportTrasmessi 
         Caption         =   "Solo non trasmessi"
         Height          =   255
         Left            =   9000
         TabIndex        =   13
         Top             =   1620
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin TrueDBGrid60.TDBGrid TDBGridExport 
         Height          =   1275
         Left            =   120
         OleObjectBlob   =   "DAMSelect2Z27.frx":0306
         TabIndex        =   8
         Top             =   240
         Width           =   10515
      End
      Begin VB.Label Label1 
         Caption         =   "File addebiti"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "File Viag. Test."
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "File Viag. Det."
         Height          =   195
         Index           =   6
         Left            =   5460
         TabIndex        =   40
         Top             =   2040
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Risultato Della Ricerca"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   25
      Top             =   1440
      Width           =   10755
      Begin TrueDBGrid60.TDBGrid TDBgrid1 
         Height          =   1575
         Left            =   120
         OleObjectBlob   =   "DAMSelect2Z27.frx":12E4D
         TabIndex        =   7
         Tag             =   "SELECT"
         Top             =   240
         Width           =   10515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1035
      Left            =   0
      TabIndex        =   31
      Top             =   5820
      Width           =   10755
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   2460
         Picture         =   "DAMSelect2Z27.frx":1E003
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   660
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
         Left            =   240
         Picture         =   "DAMSelect2Z27.frx":1E105
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   660
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   4680
         Picture         =   "DAMSelect2Z27.frx":1E207
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   660
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Esporta addebiti"
         Height          =   375
         Left            =   4560
         Picture         =   "DAMSelect2Z27.frx":1E739
         TabIndex        =   22
         Tag             =   "EXPORT"
         Top             =   600
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   3600
         Picture         =   "DAMSelect2Z27.frx":1E83B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   35
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
         Picture         =   "DAMSelect2Z27.frx":1E93D
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
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
         Left            =   8580
         Picture         =   "DAMSelect2Z27.frx":1EA3F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
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
         Index           =   1
         Left            =   240
         Picture         =   "DAMSelect2Z27.frx":1EB41
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8460
         Picture         =   "DAMSelect2Z27.frx":1EC43
         TabIndex        =   23
         Tag             =   "EXIT"
         Top             =   240
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5760
         Picture         =   "DAMSelect2Z27.frx":1ED45
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Annulla fatture"
         Height          =   375
         Left            =   2340
         Picture         =   "DAMSelect2Z27.frx":1EE47
         TabIndex        =   21
         Tag             =   "CUSTOM2"
         Top             =   600
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Rendi Definitive"
         Height          =   375
         Left            =   120
         Picture         =   "DAMSelect2Z27.frx":1EF49
         TabIndex        =   20
         Tag             =   "CUSTOM1"
         Top             =   600
         Width           =   2235
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   5640
         Picture         =   "DAMSelect2Z27.frx":1F04B
         TabIndex        =   19
         Tag             =   "PRINT"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Documento"
         Height          =   375
         Left            =   3480
         Picture         =   "DAMSelect2Z27.frx":1F14D
         TabIndex        =   18
         Tag             =   "OPEN"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "DAMSelect2Z27.frx":1F24F
         TabIndex        =   17
         Tag             =   "UNDO"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "DAMSelect2Z27.frx":1F351
         TabIndex        =   16
         Tag             =   "FIND"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7320
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1F453
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1F565
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1F677
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1F789
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1F89B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1F9AD
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1FABF
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1FBD1
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1FCE3
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1FDF5
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":1FF07
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":20019
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":2012B
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":2066D
            Key             =   "CreateDocs"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":20777
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":20889
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":20E2F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   6870
      Width           =   10770
      _ExtentX        =   18997
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
            Object.Width           =   13547
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11.20"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "25/08/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1260
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
            Picture         =   "DAMSelect2Z27.frx":20F41
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":21053
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DAMSelect2Z27.frx":21165
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      ForeColor       =   &H00FF0000&
      Height          =   1395
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10755
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo Documenti  non annullati"
         Height          =   255
         Left            =   8220
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7080
         Picture         =   "DAMSelect2Z27.frx":21277
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "CAL"
         Top             =   1020
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3360
         Picture         =   "DAMSelect2Z27.frx":21379
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "CAL"
         Top             =   1020
         Width           =   375
      End
      Begin MSMask.MaskEdBox MaskEdBox28 
         Height          =   255
         Left            =   2100
         TabIndex        =   1
         Tag             =   "64"
         Top             =   660
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
         Left            =   2100
         TabIndex        =   0
         Tag             =   "63"
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   2100
         TabIndex        =   2
         Tag             =   "65"
         Top             =   1020
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
         Left            =   5820
         TabIndex        =   4
         Tag             =   "66"
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   1260
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.txt"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   7380
         TabIndex        =   46
         Tag             =   "67"
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
         Left            =   8700
         TabIndex        =   48
         Tag             =   "67"
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Numero Pref."
         Height          =   255
         Index           =   0
         Left            =   6300
         TabIndex        =   47
         Top             =   660
         Width           =   1035
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   10380
         Tag             =   "67"
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5640
         Tag             =   "64"
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   5640
         Tag             =   "63"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inizio Periodo"
         Height          =   255
         Index           =   12
         Left            =   300
         TabIndex        =   38
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3840
         Tag             =   "65"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data &Fine Periodo"
         Height          =   255
         Index           =   11
         Left            =   4380
         TabIndex        =   37
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   7500
         Tag             =   "66"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ente Emitt."
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   30
         Top             =   300
         Width           =   855
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
         Left            =   2820
         TabIndex        =   29
         Tag             =   "63"
         Top             =   300
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Società"
         Height          =   255
         Index           =   49
         Left            =   300
         TabIndex        =   28
         Top             =   660
         Width           =   735
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
         Left            =   2820
         TabIndex        =   27
         Tag             =   "64"
         Top             =   660
         Width           =   2775
      End
   End
End
Attribute VB_Name = "DAMSelect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_COLS = 19
Public GridArraySelect As ArrayGridCls
Public WithEvents DataClass As CDocumentoSelezione
Attribute DataClass.VB_VarHelpID = -1
Public PPDSelect As CPrefatturaDettaglioSelezione
Public XAADD As XArrayDB
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes
Private mstrIDADEDef As String

Private Const EUROSYSTEM_ASC_CODE = "900"
Private Const REGISTRY_KEY_NAME_ADD_FILE_NAME = "ExportADDFileName"
Private Const REGISTRY_KEY_NAME_VIA_FILE_NAME = "ExportVIAFileName"
Private Const REGISTRY_KEY_NAME_VPROD_FILE_NAME = "ExportVPRDOFileName"

Private ColTags(1 To MAX_COLS) As Variant
Private XADetail As XArrayDB

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property
Private Sub SetButtonState()
    Dim lngCount As Long
    
    lngCount = GridArraySelect.RowCount
    
    Command1.Enabled = (lngCount > 0) And (State = fsSelection)
    Command5.Enabled = (lngCount > 0) And (State = fsSelection)
    Command6.Enabled = (XAADD.Count(1) > 0) And (State = fsSelection)
    Command8.Enabled = (lngCount > 0) And (State = fsSelection)
    Command2.Enabled = (lngCount > 0) And (State = fsSelection)
End Sub

Private Sub ResetGrid(Grid As TDBGrid, XA As XArrayDB)
    
    If XA.Count(1) > 0 Then
        XA.DeleteRows XA.LowerBound(1), XA.Count(1)
    End If
    Set Grid.Array = XA
    Grid.ReBind

End Sub

Private Sub DisplaySelection()
    If DataClass.DS.HowMany > 0 Then
        RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                           GridArraySelect, DataClass.DS.HowMany, MAX_COLS
    Else
        ClearUnboundGrid TDBGrid1, GridArraySelect, MAX_COLS
    End If
    
    If PPDSelect.DS.HowMany = 0 Then
        ResetGrid TDBGridExport, XAADD
    Else
        RefreshXArray XAADD, TDBGridExport, PPDSelect.DS.List
        TDBGridExport.Enabled = True
    End If

End Sub
Private Function GetSelectedID(DBGrid As TDBGrid, Optional GridBookmark As String = vbNullString) As Variant
    If GridBookmark = vbNullString Then
        If IsNull(DBGrid.bookmark) Then
            GetSelectedID = vbNullString
        Else
            GetSelectedID = GridArraySelect.value(DBGrid.bookmark, 2)
        End If
    Else
        GetSelectedID = GridArraySelect.value(GridBookmark, 2)
    End If
End Function

Private Function SelectRecords(Optional varSortIndex As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim blnActiveOnly As Boolean
    
    Me.MousePointer = vbHourglass

    If Not DataClass.DS.Fields.Validate Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    If Not DataClass.DS.ScanClientCond Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If

    If IsMissing(varSortIndex) Then
        ReDim varOrderFields(2) As Variant
        ' FM 19/06/2002 Cambiato l'ordinamento nella selezione (Ente, Data, Numero)
        varOrderFields(0) = "17"
        varOrderFields(1) = "2"
        varOrderFields(2) = "7"
    Else
        ReDim varOrderFields(0) As Variant
        
        varOrderFields(0) = varSortIndex
    End If

    blnActiveOnly = (chkActiveOnly = vbChecked)

    If Not DataClass.GetSelectionDAM_PAD(blnActiveOnly, varOrderFields, 0) Then
        GoTo ErrHandler
    End If

    ' Carica Addebiti
    PPDSelect.DS.Fields.ResetFields
    If ckExportTrasmessi.value = 1 Then
        PPDSelect.DS.Field("Prefattura Stato Diverso").FindValue = "T"
    End If

    If Not IsEmpty(DataClass.DS.Field("Pref. Ente Emittente").FindValue) And _
        Len(DataClass.DS.Field("Pref. Ente Emittente").FindValue) > 0 Then
        PPDSelect.DS.Field("Ente Emittente").FindValue = DataClass.DS.Field("Pref. Ente Emittente").FindValue
    End If

    If Not IsEmpty(DataClass.DS.Field("Pref. Società").FindValue) And _
        Len(DataClass.DS.Field("Pref. Società").FindValue) > 0 Then
        PPDSelect.DS.Field("Società").FindValue = DataClass.DS.Field("Pref. Società").FindValue
    End If

    If Not IsEmpty(DataClass.DS.Field("Pref. Data inizio periodo").FindValue) And _
        Len(DataClass.DS.Field("Pref. Data inizio periodo").FindValue) > 0 Then
        PPDSelect.DS.Field("Dalla Data inizio periodo").FindValue = DataClass.DS.Field("Pref. Data inizio periodo").FindValue
    End If

    If Not IsEmpty(DataClass.DS.Field("Pref. Data fine periodo").FindValue) And _
        Len(DataClass.DS.Field("Pref. Data fine periodo").FindValue) > 0 Then
        PPDSelect.DS.Field("Alla Data fine periodo").FindValue = DataClass.DS.Field("Pref. Data fine periodo").FindValue
    End If

    varOrderFields(0) = "66"
    If Not PPDSelect.GetSelectionE2K(True, varOrderFields) Then
        GoTo ErrHandler
    End If

    DisplaySelection
    GoTo EndProc
ErrHandler:
    SelectRecords = False
    Me.MousePointer = vbDefault
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function
Private Sub ClearForm()
    DataClass.DS.ClearClient
'    mskField15 = mstrIDADEDef
End Sub
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim i As Long
    Dim Rows As Long
    Dim strSelectionFormula As String
    Dim strError As String
    Dim varDocsList As Variant
    ReDim varReportsList(0 To 4, 0 To 2) As Variant
    Dim DAM As CDocumentoAmministrativo
    Dim strErrorMessage As String
    
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    Select Case Cmd
        Case fcUndo
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
        Case fcFind
            If SelectRecords Then
                SetState fsSelection
                'AutoSizeGrid TDBgrid1
            End If
        
        Case fcCustom1
            If (State = fsSelection) Then
                If MsgBox("Si vuole rendere definitive tutte le fatture selezionate?", vbYesNo + vbDefaultButton2, "Conferma") = vbYes Then
                    If Not RendiDefinitiveFatture(strErrorMessage) Then
                        MsgBox "Errore nel rendere definitive le fatture: " & strErrorMessage, vbOKOnly Or vbCritical, "Errore"
                    Else
                        DataClass.DS.DisplayMessage "Le fatture sono state rese definitive in maniera corretta."
                    End If
                    DoCommand fcFind
                End If
            End If
                
        Case fccustom2
            If (State = fsSelection) Then
                If MsgBox("Si vuole annullare tutte le fatture selezionate?", vbYesNo + vbDefaultButton2, "Conferma annullamento") = vbYes Then
                    If Not AnnullaFatture(strErrorMessage) Then
                        MsgBox "Errore nell'annullamento delle fatture: " & strErrorMessage, vbOKOnly Or vbCritical, "Errore"
                    Else
                        DataClass.DS.DisplayMessage "Le fatture sono state annullate correttamente."
                    End If
                    DoCommand fcFind
                End If
            End If

        Case fcOpen
            Set LinkForm = DisplayForm(fcDAMDAM1)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GridArraySelect.value(TDBGrid1.bookmark, 0), GridArraySelect.value(TDBGrid1.bookmark, 1), GridArraySelect.value(TDBGrid1.bookmark, 2)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If

        Case fcPrint
            Select Case mvarFunction
                Case Else
                    If (State = fsSelection) Then
                        strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                        varReportsList(0, 0) = "Stampa Documenti in Zona F"
                        varReportsList(1, 0) = "oil2dam1.rpt"
                        varReportsList(2, 0) = 1
                        varReportsList(0, 1) = "Stampa Documenti in Zona E"
                        varReportsList(1, 1) = "oil2dam2.rpt"
                        varReportsList(2, 1) = 1
                        
                        
                        varReportsList(0, 2) = "Stampa elenco Documenti Dettagliato"
                        varReportsList(1, 2) = "oil2dam3.rpt"
                        varReportsList(2, 2) = 1
                        CRPrint varReportsList, strSelectionFormula, Me
                    End If
            End Select

        Case fcAddNew
'            If ("Il documento " & GridArraySelect.value(TDBgrid1.bookmark, 5) & " verrà stornato. Si desidera proseguire?", vbYesNo + vbDefaultButton2, "Conferma eliminazione") = vbYes Then
'                If Not DAM.AnnullaDocumento(GridArraySelect.value(TDBgrid1.bookmark, 0), _
'                                           GridArraySelect.value(TDBgrid1.bookmark, 1), _
'                                           GridArraySelect.value(TDBgrid1.bookmark, 2), _
'                                           strError) Then
'                    DataClass.DS.DisplayMessage strError, Error
'                Else
'                    DataClass.DS.DisplayMessage "Il documento selezionato è stato stornato", info
'                End If
'            End If
        
        Case fcExport
            If (State = fsSelection) Then
                If MsgBox("Si vuole esportare tutti gli addebiti selezionati?", vbYesNo + vbDefaultButton2, "Conferma esportazione") = vbYes Then
                    If Not ExportAddebiti(txtFileNameAddebiti.Text, _
                                    txtFileNameViag.Text, _
                                    txtFileNameViagProd.Text, _
                                    strErrorMessage) Then
                        MsgBox "Errore nell'esportazione degli addebiti: " & strErrorMessage, vbOKOnly Or vbCritical, "Errore"
                    Else
                        MsgBox "L'esportazione degli addebiti si è conclusa correttamente.", vbOKOnly, "Messaggio"
                    End If
                    DoCommand fcFind
                End If
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
    Dim lngHowMany As Long
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            
            SetButtonState
            
            Select Case mvarFunction

            End Select
            
            mskField10.Text = "056"
            DataClass.DS.ValidateClientField mskField10.Tag, mskField10.Text
            FieldUpdateDescription mskField10.Tag, Me, mskField10
            FieldSetDescription mskField10.Tag, Me, mskField10.Text, mskField10
            
            If Not Loading Then
               MaskEdBox28.SetFocus
            End If
            
        Case fsSelection
            lngHowMany = DataClass.DS.HowMany
            SetButtonState
            
            ' FM 13/08/02 ' Aggiunta Visualizzazione Record Trovati
            Select Case lngHowMany
                Case 1
                    DataClass.DS.DisplayMessage lngHowMany & " Documento selezionato"
                Case Else
                    DataClass.DS.DisplayMessage lngHowMany & " Documenti selezionati"
            End Select
            
            If Not Loading Then
                If TDBGrid1.Enabled Then
                    TDBGrid1.SetFocus
                End If
            End If
    End Select
End Sub

Private Sub cmdBrowseExport_Click(Index As Integer)
    With CommonDialog1
    
        Select Case Index
            Case 0
                CommonDialog1.DialogTitle = "Seleziona il file degli addebiti."
                .fileName = txtFileNameAddebiti.Text
        
            Case 1
                CommonDialog1.DialogTitle = "Seleziona il file testata dei viaggi."
                .fileName = txtFileNameViag.Text
        
            Case 2
                CommonDialog1.DialogTitle = "Seleziona il file dettaglio dei viaggi."
                .fileName = txtFileNameViagProd.Text
        
        End Select
        
        .DefaultExt = ".txt"
        .Filter = "Text file (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
        .ShowOpen
        If Not .CancelError Then

            Select Case Index
                Case 0
                    txtFileNameAddebiti.Text = .fileName
                
                Case 1
                    txtFileNameViag.Text = .fileName
            
                Case 2
                    txtFileNameViagProd.Text = .fileName
            
            End Select
        End If
    End With

End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Data Inizio Periodo"
        Case 1
            Set M = mskDateTo
            strTitle = "Data Fine Periodo"
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
    Set GridArraySelect = New ArrayGridCls
    
    ColTags(1) = "17"
    ColTags(2) = "1"
    ColTags(3) = "4"
    ColTags(4) = "22"
    ColTags(5) = "23"
    ColTags(6) = "6"
    ColTags(7) = "2"
    ColTags(8) = "10"
    ColTags(9) = "55"
    ColTags(10) = "24"
    ColTags(11) = "25"
    
    ColTags(12) = "26"
    ColTags(13) = "13"
    ColTags(14) = "21"
    ColTags(15) = "30"
    
    ColTags(16) = "56"
    ColTags(17) = ""
    ColTags(18) = ""
    
    ColTags(19) = "31"
    
    For i = 1 To UBound(ColTags)
        TDBGrid1.Columns(i - 1).ButtonHeader = (ColTags(i) <> vbNullString)
    Next i
    
'    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New CDocumentoSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        .Field("Pref. Ente Emittente").Required = True
        .Field("Pref. Società").Required = True
        .Field("Pref. Data inizio periodo").Required = True
        .Field("Pref. Data fine periodo").Required = True
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    TDBGrid1.Columns(12).FetchStyle = True

    Set PPDSelect = New CPrefatturaDettaglioSelezione
    With PPDSelect.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    Set XAADD = New XArrayDB
    LoadRegistryVar

    SetState fsparaminput

    Loading = False
    
    If MaskEdBox28.Visible Then
        MaskEdBox28.SetFocus
    End If

End Sub

Private Sub Form_Resize() '    Dim i As Integer
'
'    If Me.WindowState <> vbMinimized Then
'        For i = Frame1.LBound To Frame1.UBound
'            Frame1(i).Move 0, _
'                           Frame1(i).Top, _
'                           Me.Width - 105, _
'                           Me.Height - Frame1(i).Top - StatusBar1.Height - Frame2.Height - 360
'            With TDBgrid1
'                .Move .Left, _
'                      .Top, _
'                      Frame1(2).Width - 240, _
'                      Frame1(2).Height - .Top - 120
'            End With
'        Next i
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegistryVar
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set XAADD = Nothing
    Set XADetail = Nothing
    Set PPDSelect = Nothing
    Set DAMSelect2 = Nothing
End Sub


Private Sub TDBGrid1_DblClick()
    DoCommand fcOpen
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Static blnDescending As Boolean
    Static intColIndex As Integer
    
    If intColIndex <> ColIndex Then
        blnDescending = False
    End If
    
    If blnDescending Then
        SelectRecords "D" & ColTags(ColIndex + 1)
    Else
        SelectRecords ColTags(ColIndex + 1)
    End If
    
    blnDescending = Not blnDescending
    intColIndex = ColIndex
End Sub


Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'    #If VersionID <> 1 Then
'        If KeyCode = vbKeyReturn Then
'            DoCommand fcOpen
'        End If
'    #End If
End Sub

Private Sub LoadRegistryVar()

    txtFileNameAddebiti.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_ADD_FILE_NAME, "", True)
    txtFileNameViag.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_VIA_FILE_NAME, "", True)
    txtFileNameViagProd.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_VPROD_FILE_NAME, "", True)

End Sub

Private Sub SaveRegistryVar()
    Dim blnDummy As Boolean
    
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_ADD_FILE_NAME, txtFileNameAddebiti.Text, True)
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_VIA_FILE_NAME, txtFileNameViag.Text, True)
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_VPROD_FILE_NAME, txtFileNameViagProd.Text, True)

End Sub

Private Function RendiDefinitiveFatture(Optional ByRef strErrorMessage As String)
On Error GoTo ErrHandler
    
    Dim DAM As CDocumentoAmministrativo
    Dim lngCounter As Long
    Dim strErrMessage As String
    
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    DataClass.DS.DA.BeginTransaction
    
    For lngCounter = 0 To GridArraySelect.RowCount - 1
        
        With DAM.DS
        
            .Fields.ResetFields
            .Field("Ente Emittente").FindValue = GridArraySelect.value(lngCounter, 0)
            .Field("Anno Documento").FindValue = GridArraySelect.value(lngCounter, 1)
            .Field("Progressivo Documento").FindValue = GridArraySelect.value(lngCounter, 2)

            If Not DAM.Find() Or .HowMany = 0 Then
                strErrMessage = "Errore nella ricerca del documento amministrativo!"
                GoTo ErrHandler
            End If

            If Not DAM.SetStato("S", strErrMessage) Then
                GoTo ErrHandler
            End If
        End With

    Next lngCounter

    DataClass.DS.DA.EndTrans 0
    Set DAM = Nothing
    RendiDefinitiveFatture = True
    Exit Function

ErrHandler:
    DataClass.DS.DA.EndTrans 1
    Set DAM = Nothing
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = IIf(strErrMessage = "", "Errore sconosciuto.", strErrMessage)
    End If
    RendiDefinitiveFatture = False
    
End Function

Private Function AnnullaFatture(Optional ByRef strErrorMessage As String)
On Error GoTo ErrHandler
    
    Dim PPT As CPrefattura
    Dim lngCounter As Long
    Dim strErrMessage As String

    Set PPT = New CPrefattura
    With PPT.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    If Not DataClass.GetSelectionPrefatture() Then
        strErrMessage = "Errore durante la selezione delle fatture."
        GoTo ErrHandler
    End If

    If DataClass.DS.HowMany = 0 Then
        AnnullaFatture = True
        Exit Function
    End If

    DataClass.DS.DA.BeginTransaction

    For lngCounter = LBound(DataClass.DS.List, 2) To UBound(DataClass.DS.List, 2)

        With PPT.DS
            .Fields.ResetFields
            .Field("Codice").FindValue = DataClass.DS.List(0, lngCounter)
            
            If Not PPT.Find() Or .HowMany = 0 Then
                strErrMessage = "Errore nella ricerca della prefattura."
                GoTo ErrHandler
            End If
            
            If Not PPT.SetStato("G", strErrMessage) Then
                GoTo ErrHandler
            End If
        End With
    Next lngCounter
    
    DataClass.DS.DA.EndTrans 0
    AnnullaFatture = True
    Exit Function

ErrHandler:
    DataClass.DS.DA.EndTrans 1
    AnnullaFatture = False
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = IIf(strErrMessage = "", "Errore sconosciuto.", strErrMessage)
    End If
    
End Function

Private Function IsEmptyVariantString(varValue As Variant) As Variant
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(varValue) = 0 Then
        IsEmptyVariantString = True
    Else
        IsEmptyVariantString = False
    End If
End Function

Private Function GetCodMultiSoc(ByRef strCodSoc As String, varID As Variant, strIDSoc As String, objDataClass As Object) As Boolean

    strCodSoc = ""

    If Not IsEmptyVariantString(varID) Then
        With objDataClass
            .DS.Fields(.DS.KeyFields(1)).FindValue = varID
            .DS.Fields(.DS.KeyFields(2)).FindValue = strIDSoc
            If Not .Find Then
                GetCodMultiSoc = False
                Exit Function
            End If
    
            If .DS.HowMany = 0 Then
                GetCodMultiSoc = False
                Exit Function
            End If
            
            strCodSoc = .DS.Fields("CodPerSoc").value
        End With
    End If

    GetCodMultiSoc = True

End Function

Private Function E2KFormatNumber(varValue As Variant, lngSize As Long, Optional lngDecimal As Long = 0) As String

    Dim dblValue As Double
    Dim strFormat As String
    Dim dblDecimal As Double

    dblDecimal = CDbl("1" & String(lngDecimal, "0"))
    If IsEmpty(varValue) Or IsNull(varValue) Then
        dblValue = 0
    Else
        If IsNumeric(varValue) Then
            dblValue = CDbl(varValue)
        Else
            dblValue = 0
        End If
    End If
    
    strFormat = String(lngSize, "0")

    E2KFormatNumber = Format(dblValue * dblDecimal, strFormat)

End Function

Private Function ExportAddebiti(strFileName As String, _
                    strViagFileName As String, _
                    strViagProdFileName As String, _
                    ByRef strErrorMessage As String) As Boolean
On Error GoTo ErrHandler

    Dim lngCounter As Long
    Dim TextFile As CTextFile
    Dim ViagTextFile As CTextFile
    Dim ViagProdTextFile As CTextFile
    Dim strIDACL As String
    Dim strIDAIV As String
    Dim strIDAMP As String
    Dim objTrasmAnagrafiche As CTrasmAnagrafiche
    Dim strPrefatturaIDPrec As String
    Dim blnFattureDefinitive As Boolean
    Dim XADAM As XArrayDB
    
    Dim VVT As CViaggio
    Dim SCL As CClienteSoc
    Dim SIV As CIVASoc
    Dim SMP As CPagamentoSoc
    Dim PPT As CPrefattura

    Set objTrasmAnagrafiche = New CTrasmAnagrafiche
    Set XADAM = New XArrayDB
    
    DataClass.DS.DA.BeginTransaction
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set SCL = New CClienteSoc
    With SCL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set SIV = New CIVASoc
    With SIV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set SMP = New CPagamentoSoc
    With SMP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set PPT = New CPrefattura
    With PPT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set TextFile = objTrasmAnagrafiche.TRAddebiti
    Set ViagTextFile = objTrasmAnagrafiche.TRExViagT
    Set ViagProdTextFile = objTrasmAnagrafiche.TRExViagProd
    
    Me.MousePointer = vbHourglass

    strErrorMessage = ""
'    lngRecordCount = XAADD.Count(1)
    
    If XAADD.Count(1) = 0 Then
        GoTo OKHandler
    End If
    
    If Not TextFile.OpenFile(strFileName, 1) Then
        strErrorMessage = "Impossibile aprire il file " & strFileName & " per la scrittura!"
        GoTo ErrHandler
    End If
    
    If Not ViagTextFile.OpenFile(strViagFileName, 1) Then
        strErrorMessage = "Impossibile aprire il file " & strViagFileName & " per la scrittura!"
        GoTo ErrHandler
    End If
    
    If Not ViagProdTextFile.OpenFile(strViagProdFileName, 1) Then
        strErrorMessage = "Impossibile aprire il file " & strViagProdFileName & " per la scrittura!"
        GoTo ErrHandler
    End If
    
    For lngCounter = 0 To XAADD.Count(1) - 1

        DoEvents
        
        ' Cliente
        If Not GetCodMultiSoc(strIDACL, XAADD.value(lngCounter, 7), EUROSYSTEM_ASC_CODE, SCL) Then
            strErrorMessage = "Codice per Società del Cliente '" & XAADD.value(lngCounter, 7) & "' non trovato!"
            GoTo ErrHandler
        End If
            
        ' Iva
        If Not GetCodMultiSoc(strIDAIV, XAADD.value(lngCounter, 13), EUROSYSTEM_ASC_CODE, SIV) Then
            strErrorMessage = "Codice per Società per l'IVA '" & XAADD.value(lngCounter, 13) & "' non trovato!"
            GoTo ErrHandler
        End If
        
        ' Modalità di pagamento
        If Not GetCodMultiSoc(strIDAMP, XAADD.value(lngCounter, 22), EUROSYSTEM_ASC_CODE, SMP) Then
            strErrorMessage = "Codice per Società per Modalità di Pagamento '" & XAADD.value(lngCounter, 22) & "' non trovato!"
            GoTo ErrHandler
        End If

        ' Anno
        TextFile.Item("1").value = E2KFormatNumber(XAADD.value(lngCounter, 0), TextFile.Item("1").Size)
        
        ' Filiale
        TextFile.Item("2").value = XAADD.value(lngCounter, 1)
        
        ' Chiave
        TextFile.Item("3").value = XAADD.value(lngCounter, 2)
        
        ' Ctr
        TextFile.Item("4").value = E2KFormatNumber(XAADD.value(lngCounter, 3), TextFile.Item("4").Size)

        ' Voce Addebito Cliente
        If IsNumeric(XAADD.value(lngCounter, 4)) Then
            TextFile.Item("5").value = CLng(XAADD.value(lngCounter, 4))
        Else
            TextFile.Item("5").value = XAADD.value(lngCounter, 4)
        End If
        
        ' Descrizione Addebito
        TextFile.Item("6").value = XAADD.value(lngCounter, 5)
        
        ' Filiale Cliente
        TextFile.Item("7").value = XAADD.value(lngCounter, 6)
        
        ' Codice Cliente
        TextFile.Item("8").value = E2KFormatNumber(strIDACL, TextFile.Item("8").Size)
        
        ' Contr. Fatturazione
        TextFile.Item("9").value = E2KFormatNumber(XAADD.value(lngCounter, 8), TextFile.Item("9").Size)
        
        ' Unità Misura
        TextFile.Item("10").value = XAADD.value(lngCounter, 9)
        
        ' Quantità
        TextFile.Item("11").value = E2KFormatNumber(XAADD.value(lngCounter, 10), TextFile.Item("11").Size, 2)

        ' Tariffa
        TextFile.Item("12").value = E2KFormatNumber(XAADD.value(lngCounter, 11), TextFile.Item("12").Size, 4)

        ' Importo
        TextFile.Item("13").value = E2KFormatNumber(XAADD.value(lngCounter, 12), TextFile.Item("13").Size, 4)
        
        ' Codice IVA
        TextFile.Item("14").value = strIDAIV
        
        ' Anno Fattura
        TextFile.Item("15").value = E2KFormatNumber(XAADD.value(lngCounter, 14), TextFile.Item("15").Size)
        
        ' Filiale Emissione
        TextFile.Item("16").value = XAADD.value(lngCounter, 15)
        
        ' Numero Fattura
        If IsNumeric(XAADD.value(lngCounter, 16)) Then
            TextFile.Item("17").value = E2KFormatNumber(XAADD.value(lngCounter, 16), TextFile.Item("17").Size)
        Else
            TextFile.Item("17").value = XAADD.value(lngCounter, 16)
        End If
        
        ' Data Fattura
        If IsDate(XAADD.value(lngCounter, 17)) Then
            TextFile.Item("18").value = Format(XAADD.value(lngCounter, 17), "yyyymmdd")
        Else
            TextFile.Item("18").value = "00000000"
        End If
        
        ' Rif. Anno Documento
        TextFile.Item("19").value = E2KFormatNumber(XAADD.value(lngCounter, 18), TextFile.Item("19").Size)

        ' Rif. Filiale Documento
        TextFile.Item("20").value = XAADD.value(lngCounter, 19)
        
        ' Rif. Numero Documento
        TextFile.Item("21").value = E2KFormatNumber(XAADD.value(lngCounter, 20), TextFile.Item("21").Size)

        ' Causale Documento
        Select Case XAADD.value(lngCounter, 21)
            Case "FAT"
                TextFile.Item("22").value = "001"
            
            Case "NCR"
                TextFile.Item("22").value = "002"
            
            Case Else
                TextFile.Item("22").value = XAADD.value(lngCounter, 21)
        End Select
        
        ' Pagamento
        TextFile.Item("23").value = strIDAMP
        
        ' Codice Divisa
        If XAADD.value(lngCounter, 23) = "E" Then
            TextFile.Item("24").value = "EUR"
        Else
            TextFile.Item("24").value = XAADD.value(lngCounter, 23)
        End If
        
        ' Cambio Euro
        TextFile.Item("25").value = E2KFormatNumber(XAADD.value(lngCounter, 24), TextFile.Item("25").Size, 6)
        
        ' Data Viaggio
        If IsDate(XAADD.value(lngCounter, 25)) Then
            TextFile.Item("26").value = Format(XAADD.value(lngCounter, 25), "yyyymmdd")
        Else
            TextFile.Item("26").value = "00000000"
        End If
        
        ' Filiale Fatturazione
        TextFile.Item("27").value = XAADD.value(lngCounter, 26)
        
        ' Tipo Documento
        TextFile.Item("28").value = XAADD.value(lngCounter, 27)

        ' Conto
        TextFile.Item("29").value = XAADD.value(lngCounter, 28)

        ' Resoil Key 1
        TextFile.Item("30").value = XAADD.value(lngCounter, 29)

        ' Resoil Key 2
        TextFile.Item("31").value = E2KFormatNumber(XAADD.value(lngCounter, 30), TextFile.Item("31").Size)

        ' Annullamento
        TextFile.Item("32").value = XAADD.value(lngCounter, 31)

        If Not TextFile.WriteLine() Then
            strErrorMessage = "Errore durante la scrittura del file " & strFileName & "!"
            GoTo ErrHandler
        End If

        ' Scrive il file testata dei viaggi
        With VVT.DS
            .Fields.ResetFields
            .Fields("0").FindValue = XAADD.value(lngCounter, 2)
            If Not VVT.Find Then
                strErrorMessage = "Errore nella selezione del viaggio!"
                GoTo ErrHandler
            End If
            
            If .HowMany = 0 Then
                strErrorMessage = "Errore nella selezione del viaggio!"
                GoTo ErrHandler
            End If
        
        End With

        If Not CreateRowsExViagg(VVT, EUROSYSTEM_ASC_CODE, ViagTextFile, ViagProdTextFile, strErrorMessage) Then
            GoTo ErrHandler
        End If

    Next lngCounter

    strPrefatturaIDPrec = ""
    blnFattureDefinitive = True
    
    ReDim varPPTIDBuffer(0 To 0, 0 To 1) As Variant
    Dim lngPPTIDCount As Long
    Dim lngCounter2 As Long
    
    lngPPTIDCount = 0
    
    If GridArraySelect.RowCount >= 1 And Not IsEmptyVariantString(GridArraySelect.value(0, 0)) Then

        ReDim varPPTIDBuffer(0 To GridArraySelect.RowCount - 1, 0 To 1)
        For lngCounter = 0 To GridArraySelect.RowCount - 1
            
            If lngPPTIDCount = 0 Then
                varPPTIDBuffer(lngPPTIDCount, 0) = GridArraySelect.value(lngCounter, 19)
                varPPTIDBuffer(lngPPTIDCount, 1) = (GridArraySelect.value(lngCounter, 20) = "S")
                lngPPTIDCount = lngPPTIDCount + 1
            Else
                For lngCounter2 = 0 To lngPPTIDCount - 1
                    
                    If varPPTIDBuffer(lngCounter2, 0) = GridArraySelect.value(lngCounter, 19) Then
                        varPPTIDBuffer(lngPPTIDCount - 1, 1) = varPPTIDBuffer(lngPPTIDCount - 1, 1) And GridArraySelect.value(lngCounter, 20) = "S"
                        Exit For
                    End If
                    
                    If lngCounter2 = lngPPTIDCount - 1 Then
                        varPPTIDBuffer(lngPPTIDCount, 0) = GridArraySelect.value(lngCounter, 19)
                        varPPTIDBuffer(lngPPTIDCount, 1) = (GridArraySelect.value(lngCounter, 20) = "S")
                        lngPPTIDCount = lngPPTIDCount + 1
                        Exit For
                    End If
                
                Next lngCounter2
            End If
        
        Next lngCounter
        
        For lngCounter = 0 To lngPPTIDCount - 1
    
            With PPT
            
                .DS.Fields.ResetFields
                .DS.Field("Codice").FindValue = varPPTIDBuffer(lngCounter, 0)
            
                If Not .Find() Or .DS.HowMany = 0 Then
                    strErrorMessage = "Errore nella ricerca della prefattura."
                    GoTo ErrHandler
                End If
                
                If varPPTIDBuffer(lngCounter, 1) And .DS.Field("Stato").value = "E" Then
                    If Not PPT.SetStato("T", strErrorMessage) Then
                        GoTo ErrHandler
                    End If
                End If
            End With
            
        Next lngCounter
    End If

OKHandler:
    
    DataClass.DS.DA.EndTrans ETCOMMIT
    TextFile.CloseFile
    ViagTextFile.CloseFile
    ViagProdTextFile.CloseFile
    Me.MousePointer = vbDefault
    
    Set TextFile = Nothing
    Set ViagTextFile = Nothing
    Set ViagProdTextFile = Nothing
    Set XADAM = Nothing
    Set SCL = Nothing
    Set SIV = Nothing
    Set SMP = Nothing
    Set PPT = Nothing
    Set VVT = Nothing
    Set objTrasmAnagrafiche = Nothing
    
    ExportAddebiti = True
    Exit Function

ErrHandler:
   
    TextFile.CloseFile
    ViagTextFile.CloseFile
    ViagProdTextFile.CloseFile
    
    DataClass.DS.DA.EndTrans ETROLLBACK
    If Len(strErrorMessage) = 0 Then
        strErrorMessage = "Errore sconosciuto durante l'esportazione degli addebiti!"
    End If
    
    Me.MousePointer = vbDefault
    Set ViagTextFile = Nothing
    Set ViagProdTextFile = Nothing
    Set TextFile = Nothing
    Set XADAM = Nothing
    Set SCL = Nothing
    Set SIV = Nothing
    Set SMP = Nothing
    Set PPT = Nothing
    Set VVT = Nothing
    Set objTrasmAnagrafiche = Nothing
    ExportAddebiti = False
End Function


Public Function CreateRowsExViagg(VVT As CViaggio, _
                                    IdSoc As String, _
                                    ViagTextFile As CTextFile, _
                                    ViagProdTextFile As CTextFile, _
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
            With ViagTextFile
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
                .Item("27").value = IIf(varBuffer(31, j) = "B" Or varBuffer(31, j) = "FB", "1", " ")
                
                ' (PM) 21/04/06 - Raggruppamento (GC o RP)
                If IsNull(varBuffer(30, j)) Or IsEmpty(varBuffer(30, j)) Then
                    .Item("45").value = ""
                Else
                    .Item("45").value = varBuffer(30, j)
                End If
                
                .WriteLine
            End With
        End If
        With ViagProdTextFile
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
