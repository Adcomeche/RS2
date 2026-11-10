VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIAGGI - INSERIMENTO"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10590
   Begin VB.Frame Frame6 
      Height          =   5775
      Index           =   1
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   2640
         Picture         =   "frmVVTInsert.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2640
         Picture         =   "frmVVTInsert.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   375
      End
      Begin MSComctlLib.ProgressBar pgbPortata 
         Height          =   255
         Left            =   1320
         TabIndex        =   76
         Top             =   3600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   1335
         Left            =   4680
         OleObjectBlob   =   "frmVVTInsert.frx":0204
         TabIndex        =   62
         Tag             =   "INSDETAIL1"
         Top             =   4320
         Width           =   5775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "43"
         Top             =   240
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Tag             =   "3"
         Top             =   720
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
         Left            =   1320
         TabIndex        =   5
         Tag             =   "21"
         Top             =   1080
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
         Left            =   6120
         TabIndex        =   29
         Tag             =   "15"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   6120
         TabIndex        =   27
         Tag             =   "14"
         Top             =   360
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
         Left            =   1320
         TabIndex        =   9
         Tag             =   "23"
         Top             =   1800
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
         Left            =   6120
         TabIndex        =   31
         Tag             =   "16"
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   6120
         TabIndex        =   37
         Tag             =   "17"
         Top             =   3240
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
         Left            =   6120
         TabIndex        =   33
         Tag             =   "19"
         Top             =   2520
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
         Left            =   6120
         TabIndex        =   35
         Tag             =   "20"
         Top             =   2880
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
         Left            =   3960
         TabIndex        =   25
         Tag             =   "25"
         Top             =   3240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Tag             =   "6"
         Top             =   2880
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
         Left            =   3960
         TabIndex        =   21
         Tag             =   "7"
         Top             =   2880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "frmVVTInsert.frx":499E
         TabIndex        =   38
         Tag             =   "INSDETAIL"
         Top             =   4320
         Width           =   4455
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Tag             =   "2"
         Top             =   3240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1320
         TabIndex        =   7
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
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Tag             =   "4"
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Tag             =   "5"
         Top             =   2520
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:nn"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Tag             =   "13"
         Top             =   2160
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
         Left            =   6120
         TabIndex        =   77
         Tag             =   "18"
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   6120
         TabIndex        =   79
         Tag             =   "45"
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4200
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
               Picture         =   "frmVVTInsert.frx":7F70
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTInsert.frx":8082
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTInsert.frx":8194
               Key             =   "Editing"
            EndProperty
         EndProperty
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
         Left            =   6720
         TabIndex        =   82
         Tag             =   "45"
         Top             =   3960
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contratto"
         Height          =   255
         Index           =   17
         Left            =   5160
         TabIndex        =   78
         Top             =   3960
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   10170
         Tag             =   "45"
         Top             =   3960
         Width           =   240
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
         Left            =   6720
         TabIndex        =   81
         Tag             =   "18"
         Top             =   3600
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vett. Tra&sp."
         Height          =   255
         Index           =   14
         Left            =   5160
         TabIndex        =   80
         Top             =   3600
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   10170
         Tag             =   "18"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label lblTotVuoto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3600
         TabIndex        =   75
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vuoto Lt"
         Height          =   255
         Index           =   16
         Left            =   2760
         TabIndex        =   74
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Portata R Lt"
         Height          =   255
         Index           =   15
         Left            =   8040
         TabIndex        =   73
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblPortataR 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   9120
         TabIndex        =   72
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Portata M Lt"
         Height          =   255
         Index           =   12
         Left            =   5040
         TabIndex        =   71
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblPortataM 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6120
         TabIndex        =   70
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scomparti R"
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   69
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scomparti M"
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   68
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblScompartiR 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6120
         TabIndex        =   67
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label lblScompartiM 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6120
         TabIndex        =   66
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   4680
         Tag             =   "13"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo Vendita"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   975
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
         Left            =   1920
         TabIndex        =   63
         Tag             =   "13"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Via&ggio"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   3480
         Tag             =   "43"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Data Carico"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   2640
         Tag             =   "4"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ora Carico"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   4680
         Tag             =   "5"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   4680
         Tag             =   "22"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "De&posito"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1440
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
         Left            =   1920
         TabIndex        =   61
         Tag             =   "22"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   30
         Left            =   2040
         Tag             =   "2"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N. Se&quenza"
         Height          =   255
         Index           =   30
         Left            =   240
         TabIndex        =   22
         Top             =   3240
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
         Left            =   7440
         TabIndex        =   60
         Tag             =   "15"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motric&e"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   10200
         Tag             =   "15"
         Top             =   720
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
         Left            =   6720
         TabIndex        =   59
         Tag             =   "14"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U&nità Trasp."
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   10200
         Tag             =   "14"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   4680
         Tag             =   "3"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "So&cietà"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   615
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
         Left            =   1920
         TabIndex        =   58
         Tag             =   "3"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   4680
         Tag             =   "21"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ente E&mitt."
         Height          =   195
         Index           =   9
         Left            =   450
         TabIndex        =   4
         Top             =   1080
         Width           =   765
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
         TabIndex        =   57
         Tag             =   "21"
         Top             =   1080
         Width           =   2655
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
         Left            =   1920
         TabIndex        =   56
         Tag             =   "23"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maga&zzino"
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   4680
         Tag             =   "23"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   32
         Left            =   10200
         Tag             =   "16"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Rimorchio"
         Height          =   255
         Index           =   32
         Left            =   5280
         TabIndex        =   30
         Top             =   1440
         Width           =   735
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
         Left            =   7440
         TabIndex        =   55
         Tag             =   "16"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   10200
         Tag             =   "17"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Vettore"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   36
         Top             =   3240
         Width           =   615
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
         Left            =   6720
         TabIndex        =   54
         Tag             =   "17"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   33
         Left            =   10200
         Tag             =   "19"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A&utista 1"
         Height          =   195
         Index           =   33
         Left            =   5400
         TabIndex        =   32
         Top             =   2520
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
         Left            =   6720
         TabIndex        =   53
         Tag             =   "19"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   34
         Left            =   10200
         Tag             =   "20"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aut&ista 2"
         Height          =   255
         Index           =   34
         Left            =   5400
         TabIndex        =   34
         Top             =   2880
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
         Left            =   6720
         TabIndex        =   52
         Tag             =   "20"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   4680
         Tag             =   "25"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Km prev."
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   24
         Top             =   3240
         Width           =   645
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   35
         Left            =   4680
         Tag             =   "7"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora &Part."
         Height          =   255
         Index           =   35
         Left            =   3120
         TabIndex        =   20
         Top             =   2880
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   36
         Left            =   2640
         Tag             =   "6"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "D&ata Part."
         Height          =   255
         Index           =   36
         Left            =   360
         TabIndex        =   17
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tot. Carico Lt"
         Height          =   255
         Index           =   39
         Left            =   240
         TabIndex        =   51
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblTotTrasportato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lblPortataU 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4080
         TabIndex        =   49
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Riempimento Lt"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   48
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   43
      Top             =   5880
      Width           =   10575
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   3600
         Picture         =   "frmVVTInsert.frx":82A6
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F11 - Elimina Consegna"
         Height          =   375
         Left            =   3480
         Picture         =   "frmVVTInsert.frx":87D8
         TabIndex        =   65
         Tag             =   "CUSTOM1"
         Top             =   360
         Width           =   2535
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
         Picture         =   "frmVVTInsert.frx":88DA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
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
         Left            =   8880
         Picture         =   "frmVVTInsert.frx":89DC
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
         Index           =   1
         Left            =   240
         Picture         =   "frmVVTInsert.frx":8ADE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmVVTInsert.frx":8BE0
         TabIndex        =   40
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8760
         Picture         =   "frmVVTInsert.frx":8CE2
         TabIndex        =   41
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTInsert.frx":8DE4
         TabIndex        =   39
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   6795
      Width           =   10590
      _ExtentX        =   18680
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
            Object.Width           =   13229
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
End
Attribute VB_Name = "frmVVTInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 4
Private Const MAX_COLS_1 = 7
Public GridArrayDetail As ArrayGridCls
Public GridArrayDetail1 As ArrayGridCls
Public WithEvents DataClass As CViaggio
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CViaggioConsegna
Public DetailClass1 As CViaggioProdotto
Public DetailOrder As COrdineriga
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mblnDoSearch As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private ColTags As Variant

Private APR As CProdotto
Private OCT As COrdine
Private AUT As CUnitaDiTrasporto
Private APD As CPosDoganale
Private AIM As CImballo
Private AAM As CAutomezzo

Public Sub ClearDetail1()
    ClearUnboundGrid TDBGrid2, GridArrayDetail1, MAX_COLS_1
'    lblTotOrdinato = vbNullString
End Sub







Private Sub ClearForm(Optional blnInsResult As Boolean = True)
    Dim intTipoViaggio As Integer
    Dim strNewID As String
    
    intTipoViaggio = Combo2.ListIndex
    DataClass.DS.ClearClient
    If blnInsResult Then
        strNewID = DataClass.DS.Fields("0").NewValue
        If strNewID <> vbNullString Then
            StatusBar1.Panels(2).Text = "Ultimo viaggio inserito: N." & strNewID
        End If
    End If
    Combo2.ListIndex = intTipoViaggio
    ClearUnboundGrid TDBGrid1, GridArrayDetail, MAX_COLS
    ClearUnboundGrid TDBGrid2, GridArrayDetail1, MAX_COLS_1
    lblPortataU = vbNullString
    lblTotTrasportato = vbNullString
    lblScompartiM = vbNullString
    lblPortataM = vbNullString
    lblScompartiR = vbNullString
    lblPortataR = vbNullString
    lblTotVuoto = vbNullString
    pgbPortata.value = 0
End Sub



Public Function DeleteDetail()
    On Error GoTo ErrHandler
    
    GridArrayDetail.Delete TDBGrid1.bookmark
    TDBGrid1.bookmark = 0
    TDBGrid1.Refresh
    
    Exit Function
ErrHandler:

End Function

Public Sub DisplayDetail1()
    On Error GoTo ErrHandler
    
    Dim strFieldKey As String
    Dim strFieldKey1 As String
    Dim strIDOrdine As String
    Dim varTotConsegna As Variant
    
    Me.MousePointer = vbHourglass
    strFieldKey = DetailClass.DS.KeyFields("1")
    strFieldKey1 = DetailClass.DS.KeyFields("2")
    If IsNull(TDBGrid1.bookmark) Then
        TDBGrid1.bookmark = 0
    End If
    strIDOrdine = TDBGrid1.Columns(2).Text
    
    With DetailOrder
        .DS.Fields("0").FindValue = strIDOrdine
        .GetList True
        RefreshUnboundGrid TDBGrid2, .DS.List, GridArrayDetail1, _
                           .DS.HowMany, MAX_COLS_1
    End With
    
    OCT.DS.Fields("0").FindValue = strIDOrdine
    If strIDOrdine = vbNullString Then
        GoTo ErrHandler
    End If
    If Not OCT.Find Then
        GoTo ErrHandler
    End If
'    OCT.GetTotali
'    varTotConsegna = OCT.TotaliQta(1)
'    If Not IsNull(varTotConsegna) Then
'        lblTotOrdinato = varTotConsegna
'    End If
    Me.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Me.MousePointer = vbDefault
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
    pgbPortata.Max = lblPortataU
    Exit Sub

ErrHandler:

End Sub

Private Sub DisplayTotViaggio()
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim varTotConsegna As Variant
    Dim strIDOrdine As String
    Dim intRows As Integer
    
    Me.MousePointer = vbHourglass
    intRows = GridArrayDetail.RowCount - 1
    For i = 0 To intRows
        strIDOrdine = GridArrayDetail.value(i, 2)
        If Len(strIDOrdine) > 0 Then
            With OCT
                .DS.Fields("0").FindValue = strIDOrdine
                If .Find Then
                    .GetTotali
                    varTotConsegna = varTotConsegna + .TotaliQta(1)
                End If
            End With
        End If
    Next i
    lblTotTrasportato = varTotConsegna
    If Len(lblPortataU) = 0 Then
        lblPortataU = 0
    End If
    If Len(lblTotTrasportato) = 0 Then
        lblTotTrasportato = 0
    End If
    lblTotVuoto = lblPortataU - lblTotTrasportato
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

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim Respok, Respuesta As String
    
    Select Case Cmd
        Case fcSave
                ' Control RS3
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Esta opción solo esta disponible desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
               Select Case State
                    Case fsInsert
                        blnResult = InsertRecord
                    Case Else
                        blnResult = False
                End Select
                If blnResult Then
                    SetState fsInsert, blnResult
                End If
            End If
        Case fcUndo
            If (State = fsInsert) Then
                SetState fsInsert, False
            End If
        Case fcCustom1
            DeleteDetail
            DisplayTotViaggio
        Case fcExit
            Unload Me
    End Select
End Sub





Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Private Function InsertRecord() As Boolean
    On Error GoTo ErrHandler
    Dim C As Control
    Dim blnChkValue As Boolean
    Dim i As Integer
    Dim strTempoCarico As String
    Dim strError As String
    Dim strOraInizioV As String
    Dim strOraInizioC As String
   
    
    With DataClass.DS
        .ScanClient 1
        
       
'        If Not DataClass.CheckSequenza(.Fields("6").NewValue, .Fields("2").NewValue, .Fields("14").NewValue, .Fields("22").NewValue, "") Then
'            MsgBox "Il numero di sequenza è già stato utilizzato", vbOKOnly + vbCritical, "Impossibile inserire il viaggio"
'            GoTo ErrHandler
'        End If
        
        '--JG 31/3/9 No controlar depósito
        
        If Not DataClass.CheckSequenza(.Fields("6").NewValue, .Fields("2").NewValue, .Fields("14").NewValue, "", "") Then
            MsgBox "Il numero di sequenza è già stato utilizzato", vbOKOnly + vbCritical, "Impossibile inserire il viaggio"
            GoTo ErrHandler
        End If

        '--- JG 2/4/9 Control depósito
        
         If DataClass.DS.Fields("22").NewValue <> OCT.DS.Fields("17").value Then
        
            MsgBox "El centro de carga de alguno de los pedidos no coincide con el del viaje. ", vbOKOnly + vbCritical, "Centro de carga erroneo"
            GoTo ErrHandler
            End If
        If MsgBox("¿Desea guardar?", vbYesNo + vbCritical, "Control") = vbNo Then GoTo ErrHandler
        
        ' --- Imposto il flag Calcolo km
        .Fields("26").NewValue = "M"
        .Fields("26").Updated = True
        .Fields("28").NewValue = "M"
        .Fields("28").Updated = True
        ' ---
        
        If .Fields.Validate Then
        
            .DA.BeginTransaction
            
            ' --- Imposto i km definitivi uguali ai provvisori
            .Fields("27").NewValue = .Fields("25").NewValue
            .Fields("27").Updated = True
            ' ---
            
            ' --- Calcolo il tempo di carico (presuppongo di caricare e partire
            ' --- lo stesso giorno)
            strOraInizioC = .Fields("5").NewValue
            strOraInizioV = .Fields("7").NewValue
            If (strOraInizioC <> vbNullString) And (strOraInizioV <> vbNullString) Then
                strTempoCarico = DateDiff("n", strOraInizioC, strOraInizioV)
                strTempoCarico = CInt(strTempoCarico) \ 60 & .DA.SqlDecSeparator & CInt(strTempoCarico) Mod 60
                
                .Fields("30").NewValue = strTempoCarico
                .Fields("30").Updated = True
            End If
            ' ---
            
            If Not DataClass.Insert(strError) Then
                GoTo ErrHandler
            End If
            
            With DetailClass.DS
                For i = 1 To GridArrayDetail.RowCount
                    
                    RowToRecord i - 1, DetailClass.DS, GridArrayDetail, Me, ColTags
                    If Not DetailClass.OrdineToConsegna(DetailClass.DS.Fields("37").NewValue) Then
                        GoTo ErrHandler
                    End If
                    
                    ' --- Imposto il flag Calcolo km
                    .Fields("21").NewValue = "M"
                    .Fields("21").Updated = True
                    .Fields("23").NewValue = "M"
                    .Fields("23").Updated = True
                    ' ---
                    
                    If .Fields.Validate Then
                        ' --- Imposto i km definitivi uguali ai provvisori
                        .Fields("22").NewValue = .Fields("20").NewValue
                        .Fields("22").Updated = True
                        ' ---
                        
                        #If VersionID = 1 Then
                            ' --- Imposto la data di consegna uguale alla data del viaggio
                            .Fields("41").NewValue = DataClass.DS.Fields("6").NewValue
                            .Fields("41").Updated = True
                            ' ---
                        #End If
                        
                        If Not DetailClass.Insert Then
                            GoTo ErrHandler
                        End If
                        If Not DetailClass1.OrdineToProdotti(.Fields("37").value, _
                                                             .Fields("0").value, _
                                                             .Fields("1").value) Then
                            GoTo ErrHandler
                        End If
                    End If
                    
                Next i
            End With
            GoTo EndProc
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
ErrHandler:
    InsertRecord = False
    DataClass.DS.DA.EndTrans 1
    If Len(strError) > 0 Then
        DataClass.DS.DisplayMessage strError, 9
    End If
    Exit Function
EndProc:
    InsertRecord = True
    DataClass.DS.DA.EndTrans 0
End Function




Private Sub SearchOrder()
    Dim C As Control
    Dim i As Integer
    
    With OCT.DS
        .Fields("23").FindValue = "A"
        .Fields("23").FindValue(2) = "M"
'        .Fields("22").FindValue = "A"
        .Fields("32").FindValue = 0
        If IsUserGrandesCuentas Then
            .Fields("91").FindValue = DataConnect.ConnUserName
        Else
            .Fields("90").FindValue = DataConnect.ConnUserName

        End If
    End With
    
    ShowSearchForm Me, "", OCT, TDBGrid1.TabIndex, True
    
    For i = 0 To Controls.Count - 1
        If Controls(i).Name = "TDBGrid1" Then
            frmSearch.LinkControlIndex = i
            Exit For
        End If
    Next i

End Sub

Public Sub SetState(NewState As FormState, Optional blnInsResult As Boolean = True)
    Dim strDate As String
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsInsert
            ClearForm blnInsResult
            DataClass.DS.Fields.ResetFields
            strDate = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
            MaskEdBox16.SelText = strDate
            MaskEdBox3.SelText = strDate
            If Not Loading Then
                mskField9.SetFocus
            End If
    End Select
End Sub




Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox3
            strTitle = "Data di carico"
        Case 1
            Set M = MaskEdBox16
            strTitle = "Data di partenza"
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
    Set DetailOrder = New COrdineriga
    Set ErrH = New CErrorHandler
    Set GridArrayDetail = New ArrayGridCls
    Set GridArrayDetail1 = New ArrayGridCls
    
    ColTags(1) = "1"
    ColTags(2) = "2"
    ColTags(3) = "37"
    ColTags(4) = ""
    
    ColTags1(1) = "2"
    ColTags1(2) = "3"
    ColTags1(3) = "7"
    ColTags1(4) = "6"
    ColTags1(5) = "23"
    
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
    
    With DetailOrder.DS
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
    SetValueItems TDBGrid2.Columns(2), APD
    SetValueItems TDBGrid2.Columns(3), AIM
    SetState fsInsert
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set APR = Nothing
    Set OCT = Nothing
    Set AUT = Nothing
    Set APD = Nothing
    Set AIM = Nothing
    Set AAM = Nothing
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set DetailClass = Nothing
    Set DetailClass1 = Nothing
    Set DetailOrder = Nothing
    Set GridArrayDetail = Nothing
    Set GridArrayDetail1 = Nothing
    
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set frmVVTInsert = Nothing
End Sub

Private Sub MaskEdBox7_LostFocus()
    DisplayScomparti MaskEdBox7.Text, lblScompartiR, lblPortataR
    DisplayTotViaggio
End Sub


Private Sub mskField6_LostFocus()
    DisplayScomparti mskField6.Text, lblScompartiM, lblPortataM
    DisplayTotViaggio
End Sub


Private Sub mskField7_LostFocus()
    
    If mskField7 <> vbNullString Then
        With AUT.DS
            .Fields("0").FindValue = mskField7
            If AUT.Find Then
'                mskPortata.Text = .Fields("8").value
                
                mskField6.Text = .Fields("2").value
                DataClass.DS.ValidateClientField "15", mskField6.Text
                FieldUpdateDescription "15", Me, mskField6
                DisplayScomparti mskField6.Text, lblScompartiM, lblPortataM
                
                MaskEdBox7.Text = .Fields("3").value
                DataClass.DS.ValidateClientField "16", MaskEdBox7.Text
                FieldUpdateDescription "16", Me, MaskEdBox7
                DisplayScomparti MaskEdBox7.Text, lblScompartiR, lblPortataR
                
                DisplayTotViaggio
                
                MaskEdBox5.Text = .Fields("6").value
                DataClass.DS.ValidateClientField "17", MaskEdBox5.Text
                FieldUpdateDescription "17", Me, MaskEdBox5
                
                MaskEdBox11.Text = .Fields("7").value
                DataClass.DS.ValidateClientField "18", MaskEdBox11.Text
                FieldUpdateDescription "18", Me, MaskEdBox11
                
                MaskEdBox12.Text = .Fields("4").value
                DataClass.DS.ValidateClientField "19", MaskEdBox12.Text
                FieldUpdateDescription "19", Me, MaskEdBox12
                
                MaskEdBox15.Text = .Fields("5").value
                DataClass.DS.ValidateClientField "20", MaskEdBox15.Text
                FieldUpdateDescription "20", Me, MaskEdBox15
                
                MaskEdBox8.Text = .Fields("15").value
                DataClass.DS.ValidateClientField "45", MaskEdBox8.Text
                FieldUpdateDescription "45", Me, MaskEdBox8
            End If
        End With
    End If
End Sub


Private Sub TDBGrid1_AfterDelete()
    DisplayTotViaggio
End Sub

Private Sub TDBGrid1_AfterInsert()
    DisplayTotViaggio
End Sub

Private Sub TDBgrid1_AfterUpdate()
    DisplayTotViaggio
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
                    If GridArrayDetail.value(i, 2) = .Text Then
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
    If Not blnDataOk Then
        If ColIndex = 2 Then
            TDBGrid1.Columns(3).Text = vbNullString
            ClearDetail1
        End If
    End If
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
                Me.MousePointer = vbHourglass
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
                        DisplayDetail1
                        Exit Sub
                    End If
                End With
                GoTo EndProc
            Else
                GoTo EndProc
            End If
        End If
    End With
    Me.MousePointer = vbDefault
    Exit Sub
EndProc:
    TDBGrid1.Columns(3).Text = vbNullString
    ClearDetail1
    Me.MousePointer = vbDefault
End Sub


Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    If TDBGrid1.Col = 2 Then
        If Chr(KeyAscii) = "?" Then
            KeyAscii = 0
            TDBGrid1.Text = vbNullString
            SearchOrder
        End If
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Select Case TDBGrid1.AddNewMode
        Case dbgNoAddNew
            ' --- Non modificare: LastRow può essere Null
            If LastRow = TDBGrid1.bookmark Then
            Else
                DisplayDetail1
            End If
            ' ---
        Case dbgAddNewCurrent
            ClearDetail1
    End Select
End Sub


