VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTSelect4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIDUZIONI"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11310
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   5055
      Index           =   1
      Left            =   240
      TabIndex        =   53
      Top             =   480
      Width           =   10935
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo ordini non annullati"
         Height          =   255
         Left            =   6000
         TabIndex        =   69
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   7
         Left            =   7680
         Picture         =   "frmOCTSelect4.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   6
         Left            =   7680
         Picture         =   "frmOCTSelect4.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   5
         Left            =   7680
         Picture         =   "frmOCTSelect4.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   4
         Left            =   7680
         Picture         =   "frmOCTSelect4.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   3240
         Picture         =   "frmOCTSelect4.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3240
         Picture         =   "frmOCTSelect4.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frmOCTSelect4.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmOCTSelect4.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2760
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Tag             =   "21"
         Top             =   4560
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "23"
         Top             =   4560
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Tag             =   "0"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   5280
         TabIndex        =   5
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
         Left            =   5280
         TabIndex        =   7
         Tag             =   "2"
         Top             =   1320
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
         Left            =   840
         TabIndex        =   1
         Tag             =   "17"
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
         Left            =   1920
         TabIndex        =   9
         Tag             =   "12"
         Top             =   1680
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
         Left            =   1920
         TabIndex        =   11
         Tag             =   "13"
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField12 
         Height          =   255
         Left            =   840
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Tag             =   "3"
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1920
         TabIndex        =   35
         Tag             =   "5"
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1920
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
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Tag             =   "25"
         Top             =   3120
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
         Left            =   6360
         TabIndex        =   15
         Tag             =   "20"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         Tag             =   "27"
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   6360
         TabIndex        =   43
         Tag             =   "28"
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Tag             =   "29"
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
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   6360
         TabIndex        =   26
         Tag             =   "26"
         Top             =   3120
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
         Left            =   6360
         TabIndex        =   32
         Tag             =   "30"
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   6360
         TabIndex        =   38
         Tag             =   "31"
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   7680
         TabIndex        =   66
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
         Top             =   5040
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
               Picture         =   "frmOCTSelect4.frx":0810
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTSelect4.frx":0922
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTSelect4.frx":0A34
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   68
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ordini"
         Height          =   255
         Left            =   8280
         TabIndex        =   67
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   8160
         Tag             =   "31"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Inseri&mento"
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   37
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   8160
         Tag             =   "30"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A&lla Data Ricezione"
         Height          =   255
         Index           =   14
         Left            =   4680
         TabIndex        =   31
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   8160
         Tag             =   "26"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Cons. &Pian."
         Height          =   255
         Index           =   13
         Left            =   4560
         TabIndex        =   25
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   8160
         Tag             =   "29"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Cons. Ric&h."
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   19
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   4080
         Tag             =   "27"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Riferim. Esterno &1"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   40
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   8520
         Tag             =   "28"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Riferim. Esterno &2"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   42
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Ordine"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7680
         Tag             =   "20"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla D&ata Cons. Rich."
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3720
         Tag             =   "7"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data C&ons. Pian."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3720
         Tag             =   "25"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Inserimento"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3720
         Tag             =   "5"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Ricezione"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   28
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   3720
         Tag             =   "3"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ti&po Ordine"
         Height          =   255
         Index           =   26
         Left            =   960
         TabIndex        =   44
         Top             =   4560
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   26
         Left            =   4080
         Tag             =   "21"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A&vanzamento"
         Height          =   255
         Index           =   28
         Left            =   5160
         TabIndex        =   46
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   28
         Left            =   8520
         Tag             =   "23"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   17
         Left            =   4200
         Tag             =   "18"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Zona"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   2
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
         Left            =   1440
         TabIndex        =   64
         Tag             =   "18"
         Top             =   1320
         Width           =   2655
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
         Left            =   3840
         TabIndex        =   63
         Tag             =   "12"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   8
         Top             =   1680
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   7680
         Tag             =   "12"
         Top             =   1680
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
         Left            =   3840
         TabIndex        =   62
         Tag             =   "13"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Des&tinazione"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   7680
         Tag             =   "13"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   4200
         Tag             =   "17"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Dep."
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   0
         Top             =   960
         Width           =   375
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
         Left            =   1440
         TabIndex        =   61
         Tag             =   "17"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   8640
         Tag             =   "1"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   4
         Top             =   960
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
         Left            =   5880
         TabIndex        =   60
         Tag             =   "1"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   8640
         Tag             =   "2"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Ente Em."
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   6
         Top             =   1320
         Width           =   735
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
         Left            =   5880
         TabIndex        =   59
         Tag             =   "2"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero Ordine"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3240
         Tag             =   "0"
         Top             =   2400
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Index           =   2
      Left            =   240
      TabIndex        =   65
      Top             =   480
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   8280
         Picture         =   "frmOCTSelect4.frx":0B46
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7560
         Picture         =   "frmOCTSelect4.frx":0C48
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   7920
         Picture         =   "frmOCTSelect4.frx":0D22
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   4680
         Width           =   375
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   4215
         Left            =   120
         OleObjectBlob   =   "frmOCTSelect4.frx":0E24
         TabIndex        =   49
         Tag             =   "SELECT"
         Top             =   240
         Width           =   10695
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Top             =   4680
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   60
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   8880
         TabIndex        =   74
         Top             =   4680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Caption         =   "File Esportazione (.csv)"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   4680
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   55
      Top             =   5760
      Width           =   11295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmOCTSelect4.frx":79A6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmOCTSelect4.frx":7AA8
         TabIndex        =   51
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
         Picture         =   "frmOCTSelect4.frx":7BAA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTSelect4.frx":7CAC
         TabIndex        =   50
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
         Left            =   9600
         Picture         =   "frmOCTSelect4.frx":7DAE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9480
         Picture         =   "frmOCTSelect4.frx":7EB0
         TabIndex        =   52
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5535
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9763
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
      TabIndex        =   54
      Top             =   6615
      Width           =   11310
      _ExtentX        =   19950
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
            Object.Width           =   14499
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.25"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "21/08/2008"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTSelect4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 12
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As COrdineSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private varResults As Variant

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer

Private Sub DisplaySelection()
    ClearUnboundGrid TDBGrid1, GridArraySelect, MAX_COLS
    
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, GridArraySelect, DataClass.DS.HowMany, MAX_COLS
End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = GridArraySelect.value(DBGrid.bookmark, 0)
    End If
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    
    With DataClass.DS
        If .Fields.Validate Then
            If .ScanClientCond Then
                .Fields("35").FindValue = 0
                .Fields("35").Updated = True
            
                ReDim varOrderFields(0) As Variant
                
                varOrderFields(0) = "0"
                blnActiveOnly = (chkActiveOnly = vbChecked)
                If Not DataClass.GetSelection_Riduzioni(blnActiveOnly, varOrderFields, CLng(mskMaxRecords)) Then
                    GoTo ErrHandler
                End If
                DisplaySelection
                GoTo EndProc
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    
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
    Select Case Cmd
        Case fcUndo
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
        
        Case fcFind
            If SelectRecords Then
                SetState fsSelection
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
    DataClass.DS.SetClientState NewState
    
    Select Case NewState
        Case fsparaminput
            If Not Loading Then
                TabStrip1.Tabs(1).Selected = True
            End If
        
        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Ordini trovati: " & DataClass.DS.HowMany
            TabStrip1.Tabs(2).Selected = True
            If Not Loading Then
                TDBGrid1.SetFocus
            End If
    End Select
End Sub


Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Dalla data cons. rich."
        Case 1
            Set M = MaskEdBox3
            strTitle = "Dalla data cons. pian."
        Case 2
            Set M = MaskEdBox1
            strTitle = "Dalla data ricezione"
        Case 3
            Set M = mskField4
            strTitle = "Dalla data inserimento"
        Case 4
            Set M = MaskEdBox7
            strTitle = "Alla data cons. rich."
        Case 5
            Set M = MaskEdBox8
            strTitle = "Alla data cons. pian."
        Case 6
            Set M = MaskEdBox9
            strTitle = "Alla data ricezione"
        Case 7
            Set M = MaskEdBox10
            strTitle = "Alla data inserimento"
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    MyExportGridToDelimitedFile mskFileName.Text, TDBGrid1, DataClass.DS.List, ProgressBar1, 0
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrHandler:
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command1_Click()
    Dim sFolder As String
    
    sFolder = GetFolder(hwnd)
    mskFileName = sFolder
    mskFileName.Enabled = True
End Sub

Private Sub Command5_Click()
    mskFileName.Enabled = True
End Sub

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
    
    mskFileName.Text = gmRegistry.GetSetting("Options", "RS2 - Riduzioni", "c:\export_riduzioni.csv", True)
    
    Set DataClass = New COrdineSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    SetState fsparaminput
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "RS2 - Riduzioni", mskFileName.Text, True
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set frmOCTSelect4 = Nothing
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub
