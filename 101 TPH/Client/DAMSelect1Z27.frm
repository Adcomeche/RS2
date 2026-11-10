VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form DAMSelect1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONE DOCUMENTI"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   10815
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      ForeColor       =   &H00FF0000&
      Height          =   5595
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10815
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   7680
         Picture         =   "DAMSelect1Z27.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "CAL"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   7680
         Picture         =   "DAMSelect1Z27.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "CAL"
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3720
         Picture         =   "DAMSelect1Z27.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "CAL"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3720
         Picture         =   "DAMSelect1Z27.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "CAL"
         Top             =   4560
         Width           =   375
      End
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo Documenti  non annullati"
         Height          =   255
         Left            =   7380
         TabIndex        =   26
         Top             =   480
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   6000
         TabIndex        =   18
         Tag             =   "12"
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Tag             =   "10"
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Tag             =   "9"
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Tag             =   "11"
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox12 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Tag             =   "6"
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Tag             =   "3"
         Top             =   4560
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
         Left            =   2040
         TabIndex        =   15
         Tag             =   "2"
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Tag             =   "7"
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox28 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Tag             =   "22"
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
         Left            =   2040
         TabIndex        =   4
         Tag             =   "24"
         Top             =   1680
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
         Left            =   2040
         TabIndex        =   5
         Tag             =   "25"
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox20 
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Tag             =   "23"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox53 
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Tag             =   "26"
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Tag             =   "1"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   2040
         TabIndex        =   0
         Tag             =   "17"
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
         Left            =   2040
         TabIndex        =   7
         Tag             =   "27"
         Top             =   2760
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Tag             =   "21"
         Top             =   5280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   7560
         TabIndex        =   12
         Tag             =   "56"
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Tag             =   "57"
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Tag             =   "61"
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   255
         Left            =   9060
         TabIndex        =   75
         Top             =   180
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
         Left            =   0
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
               Picture         =   "DAMSelect1Z27.frx":0408
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DAMSelect1Z27.frx":051A
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DAMSelect1Z27.frx":062C
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   20
         Left            =   3900
         Tag             =   "21"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   18
         Left            =   7380
         Tag             =   "12"
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   17
         Left            =   7380
         Tag             =   "63"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   3420
         Tag             =   "11"
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   3420
         Tag             =   "63"
         Top             =   4620
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   14
         Left            =   8100
         Tag             =   "7"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   13
         Left            =   3900
         Tag             =   "6"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   8580
         Tag             =   "56"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   5820
         Tag             =   "57"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   3180
         Tag             =   "61"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   3420
         Tag             =   "9"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   3420
         Tag             =   "10"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   2580
         Tag             =   "27"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   2580
         Tag             =   "26"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   2580
         Tag             =   "25"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   2580
         Tag             =   "24"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   2580
         Tag             =   "22"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   2580
         Tag             =   "23"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   2580
         Tag             =   "1"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2580
         Tag             =   "17"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   7380
         TabIndex        =   77
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "records"
         Height          =   255
         Left            =   9660
         TabIndex        =   76
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Num Viaggio Def."
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Num. Viaggio"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   56
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Num. Ordine"
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   55
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Importo Documento"
         Height          =   255
         Index           =   43
         Left            =   240
         TabIndex        =   54
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Vendita"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   53
         Top             =   2760
         Width           =   1095
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
         Left            =   2940
         TabIndex        =   52
         Tag             =   "27"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Anno"
         Height          =   255
         Index           =   32
         Left            =   240
         TabIndex        =   51
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Ente Emitt."
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   50
         Top             =   240
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
         Left            =   2940
         TabIndex        =   49
         Tag             =   "17"
         Top             =   240
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "Società"
         Height          =   255
         Index           =   49
         Left            =   240
         TabIndex        =   48
         Top             =   1320
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
         Left            =   2940
         TabIndex        =   47
         Tag             =   "22"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Magazzino"
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   46
         Top             =   2040
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
         Left            =   2940
         TabIndex        =   45
         Tag             =   "25"
         Top             =   2040
         Width           =   2775
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
         Left            =   2940
         TabIndex        =   44
         Tag             =   "24"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
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
         Left            =   2940
         TabIndex        =   41
         Tag             =   "23"
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mod. Pag."
         Height          =   255
         Index           =   68
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label9 
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
         Left            =   2940
         TabIndex        =   39
         Tag             =   "26"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Dalla Data Documento"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Alla Data Documento"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   37
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Al Num. Documento"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   36
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Dal Num. Documento"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Alla Data Consegna"
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   34
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Dalla Data Consegna"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   33
         Top             =   4920
         Width           =   1575
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
         Left            =   3720
         TabIndex        =   32
         Tag             =   "10"
         Top             =   3120
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destinazione"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   975
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
         Left            =   3720
         TabIndex        =   29
         Tag             =   "9"
         Top             =   3480
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Risultato Della Ricerca"
      ForeColor       =   &H00FF0000&
      Height          =   5595
      Index           =   2
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   10815
      Begin TrueDBGrid60.TDBGrid TDBgrid1 
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "DAMSelect1Z27.frx":073E
         TabIndex        =   25
         Tag             =   "SELECT"
         Top             =   240
         Width           =   10575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1035
      Left            =   0
      TabIndex        =   58
      Top             =   5640
      Width           =   10815
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   3600
         Picture         =   "DAMSelect1Z27.frx":B8F4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   4980
         Picture         =   "DAMSelect1Z27.frx":B9F6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   64
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
         Left            =   2400
         Picture         =   "DAMSelect1Z27.frx":BAF8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   675
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
         Picture         =   "DAMSelect1Z27.frx":BBFA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   675
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
         Picture         =   "DAMSelect1Z27.frx":BCFC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   61
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
         Left            =   8640
         Picture         =   "DAMSelect1Z27.frx":BDFE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   60
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
         Picture         =   "DAMSelect1Z27.frx":BF00
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Risultati Della Selezione"
         Height          =   375
         Left            =   4860
         Picture         =   "DAMSelect1Z27.frx":C002
         TabIndex        =   65
         Tag             =   "EXPORT"
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Parametri di selezione"
         Height          =   375
         Left            =   2280
         Picture         =   "DAMSelect1Z27.frx":C104
         TabIndex        =   66
         Tag             =   "IMPORT"
         Top             =   600
         Width           =   2595
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Documento"
         Height          =   375
         Left            =   120
         Picture         =   "DAMSelect1Z27.frx":C206
         TabIndex        =   67
         Tag             =   "OPEN"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "DAMSelect1Z27.frx":C308
         TabIndex        =   70
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
         Left            =   6180
         Picture         =   "DAMSelect1Z27.frx":C40A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   6060
         Picture         =   "DAMSelect1Z27.frx":C50C
         TabIndex        =   74
         Tag             =   "PRINT"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F1 - Storna il documento"
         Height          =   375
         Left            =   3480
         Picture         =   "DAMSelect1Z27.frx":C60E
         TabIndex        =   68
         Tag             =   "ADD"
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "DAMSelect1Z27.frx":C710
         TabIndex        =   69
         Tag             =   "UNDO"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "DAMSelect1Z27.frx":C812
         TabIndex        =   71
         Tag             =   "FIND"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   6690
      Width           =   10815
      _ExtentX        =   19076
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
            Object.Width           =   13626
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17.28"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "24/08/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DAMSelect1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_COLS = 19
Public GridArraySelect As ArrayGridCls
Public WithEvents DataClass As CDocumentoSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes
Private mstrIDADEDef As String

' (PM) 03/01/06 - <
' Private AXU As CUtente

Private ColTags(1 To MAX_COLS) As Variant
Private XADetail As XArrayDB

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property
Private Sub SetButtonState()
    Dim lngCount As Long
    
    lngCount = TDBgrid1.SelBookmarks.Count
    
'    With Toolbar1.Buttons
'        .Item("REFRESH").Enabled = (lngCount > 0)
'        .Item("CUSTOM1").Enabled = (lngCount > 0)
'        .Item("CUSTOM2").Enabled = (lngCount > 0)
'        .Item("EDIT").Enabled = (lngCount > 0)
'        .Item("Print").Enabled = (lngCount = 0)
'        .Item("Open").Enabled = (lngCount = 0)
'    End With
End Sub

Private Sub DisplaySelection()
    RefreshUnboundGrid TDBgrid1, DataClass.DS.List, _
                       GridArraySelect, DataClass.DS.HowMany, MAX_COLS
    
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

    If Not DataClass.GetSelectionDAM(blnActiveOnly, varOrderFields, CLng(mskMaxRecords)) Then
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
    
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    Select Case Cmd
        Case fcSave
            Select Case mvarFunction
                Case Else
                    Rows = GridArraySelect.RowCount - 1
                    With TDBgrid1.SelBookmarks
                        For i = 0 To Rows
                            .Add " " & CStr(i)
                        Next i
                    End With
                    SetButtonState
            End Select
        Case fcOpen
            Set LinkForm = DisplayForm(fcDAMDAM1)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GridArraySelect.value(TDBgrid1.bookmark, 0), GridArraySelect.value(TDBgrid1.bookmark, 1), GridArraySelect.value(TDBgrid1.bookmark, 2)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcCustom1

        Case fccustom2
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
        Case fcOpenLinkEdit
            
        Case fcAddNew
            If MsgBox("Il documento " & GridArraySelect.value(TDBgrid1.bookmark, 5) & " verrà stornato. Si desidera proseguire?", vbYesNo + vbDefaultButton2, "Conferma eliminazione") = vbYes Then
                If Not DAM.AnnullaDocumento(GridArraySelect.value(TDBgrid1.bookmark, 0), _
                                           GridArraySelect.value(TDBgrid1.bookmark, 1), _
                                           GridArraySelect.value(TDBgrid1.bookmark, 2), _
                                           strError) Then
                    DataClass.DS.DisplayMessage strError, Error
                Else
                    DataClass.DS.DisplayMessage "Il documento selezionato è stato stornato", info
                End If
            End If
        Case fcLinkRefresh
            If SelectRecords Then
                SetState fsSelection
            End If
         Case fcImport
            Frame1(1).Visible = True
            Frame1(2).Visible = False
'            Toolbar1.Buttons("EXPORT").value = tbrUnpressed
'            Toolbar1.Buttons("IMPORT").value = tbrPressed
        Case fcExport
            Frame1(1).Visible = False
            Frame1(2).Visible = True
'            Toolbar1.Buttons("EXPORT").value = tbrPressed
'            Toolbar1.Buttons("IMPORT").value = tbrUnpressed
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
            
            
'            With Toolbar1.Buttons
'                .Item("EXPORT").value = tbrUnpressed
'                .Item("IMPORT").value = tbrPressed
'                .Item("Insert").Enabled = False
'                .Item("Open").Enabled = False
'                .Item("EDIT").Enabled = False
'                .Item("CUSTOM2").Enabled = False
'            End With
            MaskEdBox6 = Year(Date)
            Frame1(1).Visible = True
            Frame1(2).Visible = False
            
            Select Case mvarFunction

            End Select
            
            mskField15.Text = mstrIDADEDef
            DataClass.DS.ValidateClientField mskField15.Tag, mstrIDADEDef
            FieldUpdateDescription mskField15.Tag, Me, mskField15
            FieldSetDescription mskField15.Tag, Me, mstrIDADEDef, mskField15
            
            If Not Loading Then
               ' mskField9.SetFocus
            End If
            
        Case fsSelection
            lngHowMany = DataClass.DS.HowMany
            Frame1(2).Visible = True
            Frame1(1).Visible = False
            
'            With Toolbar1.Buttons
'                .Item("EXPORT").value = tbrPressed
'                .Item("IMPORT").value = tbrUnpressed
'                .Item("REFRESH").Enabled = False
'                .Item("CUSTOM1").Enabled = False
'                .Item("EDIT").Enabled = False
'                .Item("Insert").Enabled = (lngHowMany > 0)
'                .Item("CUSTOM2").Enabled = False
'                .Item("Open").Enabled = (lngHowMany > 0)
'            End With
            
            ' FM 13/08/02 ' Aggiunta Visualizzazione Record Trovati
            Select Case lngHowMany
                Case 1
                    DataClass.DS.DisplayMessage lngHowMany & " Documento selezionato"
                Case Else
                    DataClass.DS.DisplayMessage lngHowMany & " Documenti selezionati"
            End Select
            
            If Not Loading Then
                If TDBgrid1.Enabled Then
                    TDBgrid1.SetFocus
                End If
            End If
    End Select
End Sub
Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox4
            strTitle = "Dalla data Documento"
        Case 1
            Set M = MaskEdBox3
            strTitle = "Alla data Documento"
        Case 2
            Set M = MaskEdBox10
            strTitle = "Dalla data Consegna"
        Case 3
            Set M = MaskEdBox2
            strTitle = "Alla data Consegna"
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
        TDBgrid1.Columns(i - 1).ButtonHeader = (ColTags(i) <> vbNullString)
    Next i
    
    
    mskMaxRecords = MAX_RECORDS
    
    Set DataClass = New CDocumentoSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    TDBgrid1.Columns(12).FetchStyle = True
    MaskEdBox3 = Format(Date, "dd/mm/yy")
    MaskEdBox4 = Format(Date, "dd/mm/yy")
    
    ' TEST---
    mskField10 = "056"
    DataClass.DS.ValidateClientField mskField10.Tag, mskField10.Text
    FieldUpdateDescription mskField10.Tag, Me, mskField10
    FieldSetDescription mskField10.Tag, Me, mskField10.Text, mskField10
    ' END TEST ---

    ' --- Leggiamo il deposito di default
' (PM) 03/01/06 - <
'    Set AXU = New CUtente
'    With AXU.DS
'        Set .da = DataConnect.da
'        Set .ErrH = ErrH
'
'        .Field("Login").FindValue = gmstrPrivateUserName ' DataClass.DS.DA.ConnUserName
'        AXU.Find
'    End With
'    mstrIDADEDef = IIf(IsNull(AXU.DS.Field("Deposito").value), vbNullString, AXU.DS.Field("Deposito").value)
'    Set AXU = Nothing
    ' ---

    SetState fsparaminput

    Loading = False
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    If Me.WindowState <> vbMinimized Then
        For i = Frame1.LBound To Frame1.UBound
            Frame1(i).Move 0, _
                           Frame1(i).Top, _
                           Me.Width - 105, _
                           Me.Height - Frame1(i).Top - StatusBar1.Height - Frame2.Height - 360
            With TDBgrid1
                .Move .Left, _
                      .Top, _
                      Frame1(2).Width - 240, _
                      Frame1(2).Height - .Top - 120
            End With
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set GridArraySelect = Nothing
    Set XADetail = Nothing
    Set DAMSelect1 = Nothing
' (PM) 03/01/06 - <
'    Set AXU = Nothing
End Sub

Private Sub TDBGrid1_DblClick()
'    DoCommand fcOpen

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





