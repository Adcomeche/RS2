VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmGP1Select 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONE PROPOSTE DI FATTURA"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   4335
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Tag             =   "2"
      Top             =   480
      Width           =   9015
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3480
         Picture         =   "frmGP1Select.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7320
         Picture         =   "frmGP1Select.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3240
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Tag             =   "4"
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Tag             =   "9"
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Tag             =   "10"
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField15 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Tag             =   "1"
         Top             =   1440
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
         Left            =   2040
         TabIndex        =   19
         Tag             =   "3"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   5880
         TabIndex        =   17
         Tag             =   "6"
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Tag             =   "0"
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
         Left            =   5880
         TabIndex        =   22
         Tag             =   "5"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Tag             =   "12"
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Tag             =   "2"
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Tag             =   "7"
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   7680
         TabIndex        =   47
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   2040
         TabIndex        =   50
         Tag             =   "16"
         Top             =   3960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   5880
         TabIndex        =   52
         Tag             =   "15"
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8280
         Top             =   2640
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
               Picture         =   "frmGP1Select.frx":0204
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1Select.frx":0316
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGP1Select.frx":0428
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Tag             =   "18"
         Top             =   360
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
         Left            =   2040
         TabIndex        =   3
         Tag             =   "47"
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
         Left            =   2640
         TabIndex        =   59
         Tag             =   "47"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Ente emittente"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   5400
         Tag             =   "47"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   5400
         Tag             =   "18"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   12
         Left            =   960
         TabIndex        =   0
         Top             =   360
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
         Left            =   2640
         TabIndex        =   58
         Tag             =   "18"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Riferimento Ordine"
         Height          =   255
         Index           =   10
         Left            =   4080
         TabIndex        =   53
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   7320
         Tag             =   "15"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " &Anno"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   51
         Top             =   3960
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   2880
         Tag             =   "16"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "records"
         Height          =   255
         Left            =   8280
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   48
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3480
         Tag             =   "2"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Riferimento Ordine"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   7320
         Tag             =   "7"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al Ri&ferimento Ordine"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   26
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5400
         Tag             =   "12"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unità di Trasporto"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
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
         Left            =   2640
         TabIndex        =   45
         Tag             =   "12"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   7800
         Tag             =   "5"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Co&nsegna"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   21
         Top             =   3240
         Width           =   1455
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
         Left            =   2640
         TabIndex        =   44
         Tag             =   "0"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Zona"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5400
         Tag             =   "0"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al Nu&mero Bolletta"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7320
         Tag             =   "6"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Consegna"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3960
         Tag             =   "3"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   5400
         Tag             =   "1"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Deposito"
         Height          =   255
         Index           =   18
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
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
         Left            =   2640
         TabIndex        =   41
         Tag             =   "1"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   7320
         Tag             =   "9"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Prodotto"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   12
         Top             =   2520
         Width           =   735
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
         Left            =   3480
         TabIndex        =   40
         Tag             =   "9"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   7320
         Tag             =   "10"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "D&estinazione"
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
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
         Left            =   3480
         TabIndex        =   39
         Tag             =   "10"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal Numero &Bolletta"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3480
         Tag             =   "4"
         Top             =   2880
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   2
      Left            =   120
      TabIndex        =   42
      Top             =   480
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   4560
         Picture         =   "frmGP1Select.frx":053A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3915
         Width           =   240
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F11 - Elimina"
         Height          =   375
         Left            =   4440
         Picture         =   "frmGP1Select.frx":063C
         TabIndex        =   57
         Tag             =   "CUSTOM1"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   2640
         Picture         =   "frmGP1Select.frx":073E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3915
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   2520
         Picture         =   "frmGP1Select.frx":0840
         TabIndex        =   32
         Tag             =   "PRINT"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   240
         Picture         =   "frmGP1Select.frx":0942
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3915
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F2 - Seleziona Tutte"
         Height          =   375
         Left            =   120
         Picture         =   "frmGP1Select.frx":0A44
         TabIndex        =   55
         Tag             =   "SAVE"
         Top             =   3840
         Width           =   2415
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   3495
         Left            =   120
         OleObjectBlob   =   "frmGP1Select.frx":0B46
         TabIndex        =   46
         Tag             =   "SELECT"
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   35
      Top             =   5040
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmGP1Select.frx":B47C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmGP1Select.frx":B57E
         TabIndex        =   29
         Tag             =   "UNDO"
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
         Index           =   7
         Left            =   240
         Picture         =   "frmGP1Select.frx":B680
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmGP1Select.frx":B782
         TabIndex        =   28
         Tag             =   "FIND"
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
         Index           =   2
         Left            =   7560
         Picture         =   "frmGP1Select.frx":B884
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmGP1Select.frx":B986
         TabIndex        =   30
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   0
      TabIndex        =   31
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Condizioni"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Risultato Ricerca"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   5910
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
            TextSave        =   "15.23"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "21/03/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGP1Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 15
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As CPrefatturaSocSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta
Public GP1 As CPrefatturaSoc

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Sub SetButtonState()
    Dim lngCount As Long
    
    lngCount = TDBGrid1.SelBookmarks.Count
    Command5.Enabled = (lngCount = 0)
    Command7.Enabled = (lngCount = 0)
    Command9.Enabled = (lngCount > 0)
End Sub

Private Sub DisplaySelection()
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, DataClass.DS.HowMany, MAX_COLS

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid, Optional GridBookmark As String = vbNullString) As Variant
    If GridBookmark = vbNullString Then
        If IsNull(DBGrid.bookmark) Then
            GetSelectedID = vbNullString
        Else
            GetSelectedID = GridArraySelect.value(DBGrid.bookmark, 0)
        End If
    Else
        GetSelectedID = GridArraySelect.value(GridBookmark, 0)
    End If
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    ReDim varOrderFields(0) As Variant
    
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            varOrderFields(0) = "0"
            If Not DataClass.GetSelection(False, varOrderFields, CLng(mskMaxRecords)) Then
                GoTo ErrHandler
            End If
            DisplaySelection
            GoTo EndProc
        End If
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
ErrHandler:
    SelectRecords = False
    Exit Function
EndProc:
    SelectRecords = True
End Function







Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim i As Long
    Dim Rows As Long
    Dim strSelectionFormula As String
    ReDim varReportsList(0 To 2, 0 To 4) As Variant
    
    Select Case Cmd
        Case fcUndo
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
        Case fcSave
            Rows = GridArraySelect.RowCount - 1
            With TDBGrid1.SelBookmarks
                For i = 0 To Rows
                    .Add " " & CStr(i)
                Next i
            End With
            SetButtonState
        Case fcCustom1
            DeleteGP1
        Case fcFind
            If SelectRecords Then
                SetState fsSelection
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcPREFGP1)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GridArraySelect.value(TDBGrid1.bookmark, 4), _
                               GridArraySelect.value(TDBGrid1.bookmark, 11)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcPrint
            If (State = fsSelection) Then
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                varReportsList(0, 0) = "Stampa elenco proposte di fattura"
                varReportsList(1, 0) = "oil2gp11.rpt"
                varReportsList(2, 0) = 0
                varReportsList(0, 1) = "Stampa elenco proposte non valorizzate"
                varReportsList(1, 1) = "oil2gp12.rpt"
                varReportsList(2, 1) = 0
                varReportsList(0, 2) = "Stampa riepilogo per proposta di fattura"
                varReportsList(1, 2) = "oil2gp15.rpt"
                varReportsList(2, 2) = 0
                varReportsList(0, 3) = "Stampa riepilogo per automezzo"
                varReportsList(1, 3) = "oil2gp14.rpt"
                varReportsList(2, 3) = 0
                varReportsList(0, 4) = "CEPSA - Importi fatture"
                varReportsList(1, 4) = "oil2gp16.rpt"
                varReportsList(2, 4) = 0
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcExit
            Unload Me
    End Select
    
End Sub

Public Sub DeleteGP1()
    On Error GoTo ErrHandler
    Dim Rows As Long
    Dim i As Long
    Dim strIDViaggio As String
    Dim intError As Errors
    Dim strIDOrdine As String
    Dim strIDProdotto As String
    
    Me.MousePointer = vbHourglass
    intError = NoError
    Rows = TDBGrid1.SelBookmarks.Count - 1
    If Rows >= 0 Then
        If MsgBox("Verranno eliminate " & Rows + 1 & " prefatture. Vuoi proseguire? ", vbExclamation + vbYesNo, "Attenzione") = vbYes Then
            For i = 0 To Rows
                strIDOrdine = GridArraySelect.value(TDBGrid1.SelBookmarks(i), 4)
                strIDProdotto = GridArraySelect.value(TDBGrid1.SelBookmarks(i), 11)
                If Not GP1.Delete(strIDOrdine, strIDProdotto) Then
                    intError = GP1InsertError
                    GoTo ErrHandler
                End If
            Next i
        End If
    End If
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    Dim C As CError
    
    Select Case intError
        Case GP1InsertError
            strMessage = "Errore durante l'eliminazione della prefattura"
            strInfo = "Ordine numero: " & strIDOrdine & " - Prodotto: " & strIDOrdine
        Case NoError
            Select Case Err.Number
                Case Else
                    intError = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(intError, "frmGP1Select", _
                "DeleteGP1", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    DoCommand fcFind
    Me.MousePointer = vbDefault
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
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            If Not Loading Then
                TabStrip1.Tabs(1).Selected = True
                MaskEdBox1.SetFocus
            End If
        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Viaggi trovati: " & DataClass.DS.HowMany
            TabStrip1.Tabs(2).Selected = True
            If Not Loading Then
                TDBGrid1.SetFocus
            End If
            SetButtonState
    End Select
End Sub


Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Dalla data consegna"
        Case 1
            Set M = MaskEdBox3
            strTitle = "Alla data consegna"
    End Select
    GetCalendarDate M, strTitle
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
    ReDim ColTags(1 To 2) As Variant
    
    Loading = True
    mintCurFrame = 1
    Set GridArraySelect = New ArrayGridCls
    
    ColTags(1) = "0"
    ColTags(2) = "1"
    
    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New CPrefatturaSocSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set GP1 = New CPrefatturaSoc
    With GP1.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    SetState fsparaminput
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set Soc = Nothing
    Set frmGP1Select = Nothing
End Sub





Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


Private Sub TDBGrid1_SelChange(Cancel As Integer)
    SetButtonState
End Sub


