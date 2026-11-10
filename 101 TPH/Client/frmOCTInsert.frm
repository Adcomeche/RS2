VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDINI - INSERIMENTO"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9375
   Begin VB.Frame Frame6 
      Height          =   5295
      Index           =   1
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   2880
         Picture         =   "frmOCTInsert.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3600
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "22"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2880
         Picture         =   "frmOCTInsert.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2160
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Tag             =   "1"
         Top             =   1080
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
         Left            =   6000
         TabIndex        =   7
         Tag             =   "2"
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
         Left            =   1560
         TabIndex        =   3
         Tag             =   "12"
         Top             =   720
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
         TabIndex        =   1
         Tag             =   "13"
         Top             =   360
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
         Left            =   1560
         TabIndex        =   22
         Tag             =   "14"
         Top             =   2520
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Tag             =   "15"
         Top             =   2880
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Tag             =   "7"
         Top             =   2160
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
         TabIndex        =   9
         Tag             =   "17"
         Top             =   1440
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
         Left            =   6000
         TabIndex        =   11
         Tag             =   "18"
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   6600
         TabIndex        =   24
         Tag             =   "25"
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1215
         Left            =   120
         OleObjectBlob   =   "frmOCTInsert.frx":0204
         TabIndex        =   37
         Tag             =   "INSDETAIL"
         Top             =   3960
         Width           =   9135
      End
      Begin MSMask.MaskEdBox mskField13 
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Tag             =   "20"
         Top             =   1800
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
         Left            =   960
         TabIndex        =   13
         Tag             =   "16"
         Top             =   1800
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
         Left            =   6600
         TabIndex        =   20
         Tag             =   "10"
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   480
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
               Picture         =   "frmOCTInsert.frx":499E
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTInsert.frx":4AB0
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTInsert.frx":4BC2
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Tag             =   "5"
         Top             =   3600
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
         Left            =   4440
         TabIndex        =   35
         Tag             =   "6"
         Top             =   3600
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
         Left            =   6600
         TabIndex        =   36
         Tag             =   "37"
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Tag             =   "36"
         Top             =   3240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nota 3"
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   29
         Top             =   3240
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   5160
         Tag             =   "36"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Km Raggio"
         Height          =   255
         Index           =   26
         Left            =   5520
         TabIndex        =   55
         Top             =   3240
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   7920
         Tag             =   "37"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Or&a"
         Height          =   255
         Index           =   6
         Left            =   4080
         TabIndex        =   34
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5160
         Tag             =   "6"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Inseriment&o"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   31
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3360
         Tag             =   "5"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "St&ato Ordine"
         Height          =   255
         Index           =   27
         Left            =   5520
         TabIndex        =   27
         Top             =   2880
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   8880
         Tag             =   "22"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   8280
         Tag             =   "10"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ri&f. Esterno"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
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
         Left            =   1560
         TabIndex        =   54
         Tag             =   "16"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo &Vend."
         Height          =   255
         Index           =   0
         Left            =   -840
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   4080
         Tag             =   "16"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   8880
         Tag             =   "20"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Raggruppamento"
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
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
         Left            =   6600
         TabIndex        =   53
         Tag             =   "20"
         Top             =   1800
         Width           =   2175
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
         TabIndex        =   52
         Tag             =   "12"
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   8880
         Tag             =   "12"
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
         Left            =   3480
         TabIndex        =   51
         Tag             =   "13"
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dest&inazione"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   8880
         Tag             =   "13"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   4080
         Tag             =   "1"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   4
         Top             =   1080
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
         Left            =   1560
         TabIndex        =   50
         Tag             =   "1"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   8880
         Tag             =   "2"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Emitt."
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   6
         Top             =   1080
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
         Left            =   6600
         TabIndex        =   49
         Tag             =   "2"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   5160
         Tag             =   "15"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N&ota 2"
         Height          =   255
         Index           =   14
         Left            =   960
         TabIndex        =   25
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "No&ta 1"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   21
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5160
         Tag             =   "14"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3360
         Tag             =   "7"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Cons. &Rich."
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
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
         TabIndex        =   48
         Tag             =   "17"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Deposito"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   4080
         Tag             =   "17"
         Top             =   1440
         Width           =   240
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
         Left            =   6600
         TabIndex        =   47
         Tag             =   "18"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Zona"
         Height          =   255
         Index           =   17
         Left            =   5400
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   17
         Left            =   8880
         Tag             =   "18"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Priorità"
         Height          =   255
         Index           =   30
         Left            =   5880
         TabIndex        =   23
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   30
         Left            =   7080
         Tag             =   "25"
         Top             =   2520
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   42
      Top             =   5400
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmOCTInsert.frx":4CD4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   43
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
         Left            =   7680
         Picture         =   "frmOCTInsert.frx":4DD6
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
         Index           =   1
         Left            =   240
         Picture         =   "frmOCTInsert.frx":4ED8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmOCTInsert.frx":4FDA
         TabIndex        =   39
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmOCTInsert.frx":50DC
         TabIndex        =   40
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTInsert.frx":51DE
         TabIndex        =   38
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   41
      Top             =   6270
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
            TextSave        =   "11:04"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "26/03/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 7
Public GridArrayDetail As ArrayGridCls

Public WithEvents DataClass As COrdine
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As COrdineriga
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private mintCurFrame As Integer
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private mstrIDImballoDefault As String
Private mstrIDPosDefault As String
Private mstrIDPagamento As String
Private mstrIDTipoDestinazione As String
Private mstrIDIva As String
Private ColTags As Variant

Private strKeepDate As String

Private APR As CProdotto
Private AIV As CIVA
Private AMP As CPagamento
Private APD As CPosDoganale
Private AIM As CImballo
Private ACL As CCliente
Private ADZ As CDestinazione
Private AXX As CParametro

Public Sub InsertDetail()
    DetailClass.Insert
End Sub
Private Sub ClearForm(Optional blnInsResult As Boolean = True)
    Dim strNewID As String
    strKeepDate = mskField2
    DataClass.DS.ClearClient
    strNewID = DataClass.DS.Fields("0").NewValue
    If strNewID <> vbNullString Then
        StatusBar1.Panels(2).Text = "Ultimo ordine inserito: N." & strNewID
    End If
    ClearUnboundGrid TDBGrid1, GridArrayDetail, MAX_COLS
End Sub
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcSave
'            Select Case State
'                Case fsInsert
'                    blnResult = InsertRecord
'                Case Else
'                    blnResult = False
'            End Select
'            If blnResult Then
'                SetState fsInsert, blnResult
'            End If
        MsgBox ("Esta opción se realiza desde Resoil 3 Web")
        Case fcUndo
            If (State = fsInsert) Then
                SetState fsInsert
            End If
        Case fcExit
            Unload Me
    End Select
End Sub
Private Function InsertRecord() As Boolean
    On Error GoTo ErrHandler
    Dim C As Control
    Dim blnChkValue As Boolean
    Dim i As Integer
    
    With DataClass.DS
        .ScanClient 1
        
        #If VersionID = 1 Then
            If Len(MaskEdBox2.Text) > 0 Then
                If DataClass.CheckSocCode(MaskEdBox2.Text) Then
                    MsgBox "Numero ordine già inserito nel sistema", vbExclamation, "Attenzione!"
                    GoTo ErrHandler
                End If
            End If
        #End If
        
        'Data e Ora inserimento
        
'        .Fields("5").NewValue = Format$(Date, "dd/mm/yyyy")
'        .Fields("5").Updated = True
        If Not .Fields("6").Updated Then
            .Fields("6").NewValue = Format$(Time, "hh:nn")
            .Fields("6").Updated = True
        End If
        '--- Data consegna pianificata
        If Not .Fields("8").Updated Then
            .Fields("8").NewValue = .Fields("7").NewValue
            .Fields("8").Updated = True
        End If
        ' ---
        
        #If VersionID = 4 Then
            .Fields("26").NewValue = "MM"
            .Fields("26").Updated = True
        #End If
        
        .DA.BeginTransaction
        
        If .Fields.Validate Then
            ACL.DS.Fields("0").FindValue = .Fields("12").NewValue
            ACL.Find
            
            'Tipo Ordine
            If Len(mstrIDTipoDestinazione) > 0 Then
                .Fields("21").NewValue = mstrIDTipoDestinazione
            Else
                .Fields("21").NewValue = ACL.DS.Fields("13").value
            End If
            .Fields("21").Updated = True
            
            If Not DataClass.Insert Then
                GoTo ErrHandler
            End If
            
'            #If VersionID = 4 Then
'                If Not (ReOrderArray(GridArrayDetail)) Then
'                    GoTo ErrHandler
'                End If
'            #End If
            
            For i = 1 To GridArrayDetail.RowCount
                RowToRecord i - 1, DetailClass.DS, GridArrayDetail, Me, ColTags
                If DetailClass.DS.Fields.Validate Then
                    If Not DetailClass.Insert Then
                        GoTo ErrHandler
                    End If
                End If
            Next i

            GoTo EndProc
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    GoTo EndProc
ErrHandler:
    InsertRecord = False
    DataClass.DS.DA.EndTrans 1
    Exit Function
EndProc:
    DataClass.DS.DA.EndTrans 0
    InsertRecord = True
End Function
Public Sub SetState(NewState As FormState, Optional blnInsResult As Boolean = True)
    Dim strNewID As String
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsInsert
            ClearForm blnInsResult
'            strNewID = DataClass.DS.Fields("0").NewValue
'            If strNewID <> vbNullString Then
'                StatusBar1.Panels(2).Text = "Ultimo ordine inserito: N." & strNewID
'            End If
            DataClass.DS.Fields.ResetFields
            DataClass.DS.Fields("13").ForeignTableFldValue = vbNullString
            
            MaskEdBox3.SelText = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
            
            If strKeepDate = DATE_MASK Then
                mskField2.SelText = Format$(DateAdd("d", 1, Date), "dd/mm/yy")
            Else
                mskField2.SelText = strKeepDate
            End If
            Combo1.ListIndex = 0
            If Not Loading Then
                mskField7.SetFocus
            End If
    End Select
End Sub
Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskField2
            strTitle = "Data cons. richiesta"
        Case 1
            Set M = MaskEdBox3
            strTitle = "Data inserimento"
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
    
    Loading = True
    
    Set DataClass = New COrdine
    Set DetailClass = New COrdineriga
    Set ErrH = New CErrorHandler
    Set GridArrayDetail = New ArrayGridCls

    mintCurFrame = 1
    ReDim ColTags(1 To 7) As Variant
    ColTags(1) = "1"
    ColTags(2) = "2"
    ColTags(3) = "4"
    ColTags(4) = "3"
    ColTags(5) = "7"
    ColTags(6) = "14"
    ColTags(7) = "15"
    
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
            
            RefreshUnboundGrid TDBGrid1, .List, GridArrayDetail, _
                          0, MAX_COLS
        
        
        
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
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        With .MultiSoc.DS
            Set .DA = DataConnect.DA
            Set .ErrH = ErrH
        End With
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
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
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
    SetValueItems TDBGrid1.Columns(2), APD
    SetValueItems TDBGrid1.Columns(3), AIM
    SetValueItems TDBGrid1.Columns(5), AIV
    SetValueItems TDBGrid1.Columns(6), AMP
    
'    Select Case gmstrCodInstallazione
'        Case ciInst2
    #If (VersionID = 2) Then ' Or (VersionID = 5) Then
            Label1(27).Visible = False
            Combo1.Visible = False
            imgState(27).Visible = False
'    End Select
    #End If
        
    SetState fsInsert
    
    Loading = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set APR = Nothing
    Set AIV = Nothing
    Set AMP = Nothing
    Set APD = Nothing
    Set AIM = Nothing
    Set ACL = Nothing
    Set ADZ = Nothing
    Set AXX = Nothing
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set DetailClass = Nothing
    Set GridArrayDetail = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set frmOCTInsert = Nothing
End Sub

Private Sub MaskEdBox2_LostFocus()
    #If VersionID = 1 Then
        If Len(MaskEdBox2.Text) > 0 Then
            If DataClass.CheckSocCode(MaskEdBox2.Text) Then
                MsgBox "Numero ordine già inserito nel sistema", vbExclamation, "Attenzione!"
            End If
        End If
    #End If
End Sub


Private Sub mskField7_LostFocus()
    With ADZ
        If mskField7 <> vbNullString Then
            .DS.Fields("0").FindValue = mskField7
            .Find
            mstrIDIva = .DS.Fields("14").value
            mstrIDPagamento = .DS.Fields("13").value
            mstrIDTipoDestinazione = .DS.Fields("12").value
            ' --- Km Raggio
            MaskEdBox4 = .DS.Fields("29").value
            ' ---
            #If VersionID = 4 Then
               ProponiOrdine mskField7
               
            #End If
        End If
    End With
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
            Case 1
                For i = 0 To GridArrayDetail.RowCount - 1
                    If GridArrayDetail.value(i, 1) = .Text Then
                        blnDataOk = False
                        Exit For
                    End If
                Next i
            Case 4
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                Else
                    If CDbl(.Text) <= 0 Then
                        #If VersionID <> 4 Then
                             blnDataOk = False
                        #End If
                    End If
                End If
'                If CDbl(.Text) > 0 And CDbl(.Text) <= 99 Then
'                        .Text = CStr(CDbl(.Text) * 1000)
'                End If
'                ElseIf CDbl(.Text) > 9 And CDbl(.Text) <= 99 Then
'                        .Text = CStr(CDbl(.Text) * 100)
'                End If
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
                #If VersionID <> 4 Then
                    blnDataOk = False
                #End If
            End If
        End If
    End With
    Cancel = Not blnDataOk
End Sub
Private Sub ProponiOrdine(CodDest As String)

' Routine che propone, per una destinazione, le quantità ordinabili
' Prendendole dalla tabella PREZAPZ0

Dim RowList As Variant
Dim intRes As Integer
Dim HowManyProd As Integer
Dim NewRowList() As Variant
Dim intRowCount As Integer

    intRes = DetailClass.GetOrderList(CodDest, RowList, HowManyProd)
        If intRes > 0 Then
            ReDim Preserve NewRowList(0 To 6, 0 To HowManyProd)
            For intRowCount = 0 To HowManyProd
                NewRowList(0, intRowCount) = (intRowCount + 1) * 10 ' Contatore
                NewRowList(1, intRowCount) = RowList(0, intRowCount)  'Prodotto da APZ
                NewRowList(2, intRowCount) = RowList(2, intRowCount) ' PDog
                NewRowList(3, intRowCount) = RowList(3, intRowCount) 'mstrIDImballoDefault
                NewRowList(4, intRowCount) = 0
                NewRowList(5, intRowCount) = mstrIDIva
                NewRowList(6, intRowCount) = mstrIDPagamento
            Next intRowCount
            RefreshUnboundGrid TDBGrid1, NewRowList, GridArrayDetail, _
                HowManyProd, MAX_COLS
        End If
End Sub
Private Function ReOrderArray(ArrayDetail As ArrayGridCls) As Boolean

' Con questa funzione Elimino le righe di ordine con quantità 0 e
' rinumero l'array
' se ritorna false significa che l'array contiene solo righe con quantità 0
' e quindi annullo l'inserimento

Dim intCont As Integer
Dim intInizRows As Integer

intInizRows = ArrayDetail.RowCount - 1

    For intCont = intInizRows To 0 Step -1
        If ArrayDetail.value(intCont, 4) <= 0 Then
           ArrayDetail.DeleteRow (intCont)
        End If
    Next intCont
    For intCont = 0 To ArrayDetail.RowCount - 1
        ArrayDetail.value(intCont, 0) = (intCont + 1) * 10
    Next intCont
    If ArrayDetail.RowCount = 0 Then
        ReOrderArray = False
    Else
        ReOrderArray = True
    End If
End Function

