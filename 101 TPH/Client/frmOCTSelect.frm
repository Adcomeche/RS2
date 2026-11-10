VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONE ORDINI"
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
      TabIndex        =   60
      Top             =   480
      Width           =   10935
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "59"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo ordini non annullati"
         Height          =   255
         Left            =   8160
         TabIndex        =   78
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   7
         Left            =   7680
         Picture         =   "frmOCTSelect.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   6
         Left            =   7680
         Picture         =   "frmOCTSelect.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   5
         Left            =   7680
         Picture         =   "frmOCTSelect.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   4
         Left            =   7680
         Picture         =   "frmOCTSelect.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   3240
         Picture         =   "frmOCTSelect.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3240
         Picture         =   "frmOCTSelect.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   3240
         Picture         =   "frmOCTSelect.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3240
         Picture         =   "frmOCTSelect.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2760
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Tag             =   "21"
         Top             =   4560
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Tag             =   "23"
         Top             =   4560
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   18
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
         TabIndex        =   10
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
         TabIndex        =   12
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
         TabIndex        =   6
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
         TabIndex        =   14
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
         TabIndex        =   16
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
         TabIndex        =   8
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
         TabIndex        =   34
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
         TabIndex        =   40
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
         TabIndex        =   22
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
         TabIndex        =   28
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
         TabIndex        =   20
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
         TabIndex        =   46
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
         TabIndex        =   48
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
         TabIndex        =   25
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
         TabIndex        =   31
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
         TabIndex        =   37
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
         TabIndex        =   43
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
         TabIndex        =   75
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
               Picture         =   "frmOCTSelect.frx":0810
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTSelect.frx":0922
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTSelect.frx":0A34
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField11 
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Tag             =   "32"
         Top             =   240
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
         Left            =   840
         TabIndex        =   3
         Tag             =   "89"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   53
         Left            =   7680
         Tag             =   "59"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Contattare cliente"
         Height          =   255
         Index           =   53
         Left            =   3840
         TabIndex        =   80
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Presidio"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   2280
         Tag             =   "90"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   25
         Left            =   4200
         Tag             =   "32"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ra&ggr."
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   495
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
         Left            =   1440
         TabIndex        =   79
         Tag             =   "32"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   6000
         TabIndex        =   77
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ordini"
         Height          =   255
         Left            =   8280
         TabIndex        =   76
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
         TabIndex        =   42
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
         TabIndex        =   36
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
         TabIndex        =   30
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
         TabIndex        =   24
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
         TabIndex        =   45
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
         TabIndex        =   47
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Ordine"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   19
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
         TabIndex        =   21
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
         TabIndex        =   27
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
         TabIndex        =   39
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
         TabIndex        =   33
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
         TabIndex        =   49
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
         TabIndex        =   51
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
         TabIndex        =   7
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   13
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
         TabIndex        =   69
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
         TabIndex        =   15
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
         TabIndex        =   5
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
         TabIndex        =   68
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
         TabIndex        =   9
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
         TabIndex        =   67
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
         TabIndex        =   11
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
         TabIndex        =   66
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
         TabIndex        =   17
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
      TabIndex        =   72
      Top             =   480
      Visible         =   0   'False
      Width           =   10935
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   2520
         Picture         =   "frmOCTSelect.frx":0B46
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   4635
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Ordine"
         Height          =   375
         Left            =   2400
         Picture         =   "frmOCTSelect.frx":0C48
         TabIndex        =   56
         Tag             =   "OPEN"
         Top             =   4560
         Width           =   2295
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
         Picture         =   "frmOCTSelect.frx":0D4A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4635
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTSelect.frx":0E4C
         TabIndex        =   55
         Tag             =   "PRINT"
         Top             =   4560
         Width           =   2295
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   4215
         Left            =   120
         OleObjectBlob   =   "frmOCTSelect.frx":0F4E
         TabIndex        =   54
         Tag             =   "SELECT"
         Top             =   240
         Width           =   10695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   62
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
         Picture         =   "frmOCTSelect.frx":9AC4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmOCTSelect.frx":9BC6
         TabIndex        =   58
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
         Picture         =   "frmOCTSelect.frx":9CC8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTSelect.frx":9DCA
         TabIndex        =   57
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
         Picture         =   "frmOCTSelect.frx":9ECC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9480
         Picture         =   "frmOCTSelect.frx":9FCE
         TabIndex        =   59
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5535
      Left            =   120
      TabIndex        =   53
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
      TabIndex        =   61
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
            TextSave        =   "11:02"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23/04/2012"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 13 '11
Public GridArraySelect As ArrayGridCls

Public WithEvents DataClass As COrdineSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes
Private mstrPrevValueARO As String

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Sub DisplaySelection()
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, DataClass.DS.HowMany, MAX_COLS

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
                ReDim varOrderFields(0) As Variant
                
                varOrderFields(0) = "0"
                blnActiveOnly = (chkActiveOnly = vbChecked)
                If Not DataClass.GetSelection(blnActiveOnly, varOrderFields, CLng(mskMaxRecords), True) Then
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
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strSelectionFormula As String
    #If VersionID = 1 Then
        ReDim varReportsList(0 To 2, 0 To 4) As Variant
    #Else
        ReDim varReportsList(0 To 2, 0 To 1) As Variant
    #End If
    
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
        Case fcPrint
            If (State = fsSelection) Then
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                varReportsList(0, 0) = "Informe Pedidos"
                varReportsList(1, 0) = "oil2oct1.rpt"
                varReportsList(2, 0) = 0
                varReportsList(0, 1) = "Resumen Diario Pedidos Previstos"
                varReportsList(1, 1) = "oil2oct2.rpt"
                varReportsList(2, 1) = 1
'                varReportsList(0, 2) = "Stampa elenco Ordini con Annotazioni"
                varReportsList(0, 2) = "Informe pedidos solicitados"
                varReportsList(1, 2) = "oil2oct3.rpt"
                varReportsList(2, 2) = 1
'                varReportsList(0, 3) = "Stampa elenco Ordini con date discordanti"
                varReportsList(0, 3) = "Incumplimientos en la Programacion"
                varReportsList(1, 3) = "oil2oct4.rpt"
                varReportsList(2, 3) = 1
                varReportsList(0, 4) = "Stampa elenco Ordini non fatturati"
                varReportsList(1, 4) = "oil2oct5.rpt"
                varReportsList(2, 4) = 1
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcORDIOCT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcLinkRefresh
            If TDBGrid1.Enabled Then
                TDBGrid1.SetFocus
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
                mskField11.SetFocus
            End If
            Select Case mvarFunction
                Case fcORDSELE
                Case fcORDSELA
                    Combo3 = "ADQUIRIDO"
                    Combo3.Locked = True
                Case fcORDSELV
                    Combo3 = "VIAJE"
                    Combo3.Locked = True
                Case fcORDSELM
                    Combo3 = "MOVER"
                    Combo3.Locked = True
            End Select
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
    
    
'    Set Soc = New CSocieta
'    Set Soc.DS.DA = DataConnect.DA
'    SetValueItems TDBGrid1.Columns(1), Soc
    
    SetState fsparaminput
    
'    If IsUserGrandesCuentas Then
'        mskField11 = "GC"
'        MaskEdBox11.Tag = "91"
'        imgState(11).Tag = "91"
'    Else
'        mskField11 = "RP"
'        MaskEdBox11.Tag = "90"
'        imgState(11).Tag = "90"
'    End If
    MaskEdBox11 = DataConnect.ConnUserName
    
    Select Case mvarFunction
        Case fcORDSELE
        Case Else
            DoCommand fcFind
    End Select
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set Soc = Nothing
    Set frmOCTSelect = Nothing
End Sub





Private Sub mskField11_GotFocus()
     mstrPrevValueARO = mskField11
End Sub

Private Sub mskField11_Validate(Cancel As Boolean)
    Select Case UCase(mskField11)
        Case "GC"
            MaskEdBox11.Tag = "91"
            imgState(11).Tag = "91"
            DataClass.DS.Fields("90").FindValue = vbNullString
        Case "RP"
            MaskEdBox11.Tag = "90"
            imgState(11).Tag = "90"
            DataClass.DS.Fields("91").FindValue = vbNullString
        Case Else
            Exit Sub
    End Select
    If mskField11 <> mstrPrevValueARO Then
        MaskEdBox11 = vbNullString
    End If
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DoEvents
        DoCommand fcOpen
    End If
End Sub


