VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2ACL0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIENTI"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   48
      Top             =   5040
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5280
         Picture         =   "OIL2ACL0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2ACL0.frx":0102
         TabIndex        =   44
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
         Index           =   3
         Left            =   240
         Picture         =   "OIL2ACL0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   795
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ACL0.frx":0306
         TabIndex        =   43
         Tag             =   "PRINT"
         Top             =   720
         Visible         =   0   'False
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
         Picture         =   "OIL2ACL0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   49
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
         Picture         =   "OIL2ACL0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   51
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
         Picture         =   "OIL2ACL0.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   52
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
         Picture         =   "OIL2ACL0.frx":070E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2ACL0.frx":0810
         TabIndex        =   42
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2ACL0.frx":0912
         TabIndex        =   45
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2ACL0.frx":0A14
         TabIndex        =   41
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ACL0.frx":0B16
         TabIndex        =   40
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   47
      Top             =   5970
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
            TextSave        =   "16.38"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "05/01/00"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   3495
         Index           =   1
         Left            =   360
         TabIndex        =   55
         Top             =   1200
         Width           =   8655
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   2040
            TabIndex        =   3
            Tag             =   "1"
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField3 
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Tag             =   "4"
            Top             =   1320
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   2040
            TabIndex        =   13
            Tag             =   "5"
            Top             =   2040
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField6 
            Height          =   255
            Left            =   6840
            TabIndex        =   15
            Tag             =   "6"
            Top             =   2040
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField4 
            Height          =   255
            Left            =   2040
            TabIndex        =   11
            Tag             =   "7"
            Top             =   1680
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField7 
            Height          =   255
            Left            =   2040
            TabIndex        =   5
            Tag             =   "2"
            Top             =   600
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField8 
            Height          =   255
            Left            =   2040
            TabIndex        =   7
            Tag             =   "3"
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField16 
            Height          =   255
            Left            =   2040
            TabIndex        =   17
            Tag             =   "16"
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField17 
            Height          =   255
            Left            =   2040
            TabIndex        =   19
            Tag             =   "17"
            Top             =   2760
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField19 
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Tag             =   "18"
            Top             =   3120
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   7320
            Tag             =   "18"
            Top             =   3120
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nota &2"
            Height          =   255
            Index           =   16
            Left            =   1200
            TabIndex        =   20
            Top             =   3120
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   15
            Left            =   7320
            Tag             =   "17"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nota &1"
            Height          =   255
            Index           =   15
            Left            =   1200
            TabIndex        =   18
            Top             =   2760
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   3720
            Tag             =   "16"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Telefono"
            Height          =   255
            Index           =   14
            Left            =   1200
            TabIndex        =   16
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ragione sociale Abbr."
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   1695
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   3360
            Tag             =   "3"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ragione sociale Agg."
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   1695
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   7320
            Tag             =   "2"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ragione sociale"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   7320
            Tag             =   "1"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   7320
            Tag             =   "4"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Indirizzo"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   8
            Top             =   1320
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   5400
            Tag             =   "5"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Località"
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   12
            Top             =   2040
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   7320
            Tag             =   "6"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Prov."
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   14
            Top             =   2040
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   3360
            Tag             =   "7"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "C&AP"
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   10
            Top             =   1680
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Index           =   2
         Left            =   360
         TabIndex        =   56
         Top             =   1200
         Visible         =   0   'False
         Width           =   8655
         Begin MSMask.MaskEdBox mskField9 
            Height          =   255
            Left            =   2160
            TabIndex        =   27
            Tag             =   "8"
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
            Left            =   2160
            TabIndex        =   29
            Tag             =   "9"
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField11 
            Height          =   255
            Left            =   2160
            TabIndex        =   31
            Tag             =   "12"
            Top             =   1680
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
            Left            =   2160
            TabIndex        =   33
            Tag             =   "13"
            Top             =   2040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField13 
            Height          =   255
            Left            =   2160
            TabIndex        =   35
            Tag             =   "14"
            Top             =   2400
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField14 
            Height          =   255
            Left            =   2160
            TabIndex        =   37
            Tag             =   "15"
            Top             =   2760
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
            Left            =   2160
            TabIndex        =   23
            Tag             =   "10"
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField20 
            Height          =   255
            Left            =   2160
            TabIndex        =   25
            Tag             =   "11"
            Top             =   600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Partita IVA"
            Height          =   255
            Index           =   18
            Left            =   480
            TabIndex        =   24
            Top             =   600
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   5160
            Tag             =   "11"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codice &Fiscale"
            Height          =   255
            Index           =   17
            Left            =   480
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   5160
            Tag             =   "10"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   5520
            Tag             =   "15"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&IVA"
            Height          =   255
            Index           =   13
            Left            =   960
            TabIndex        =   36
            Top             =   2760
            Width           =   1095
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
            Left            =   2760
            TabIndex        =   62
            Tag             =   "15"
            Top             =   2760
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   5520
            Tag             =   "14"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Modalità di Pagamento"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   34
            Top             =   2400
            Width           =   1815
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
            Left            =   2760
            TabIndex        =   61
            Tag             =   "14"
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   5520
            Tag             =   "13"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Tipo Cliente"
            Height          =   255
            Index           =   11
            Left            =   1080
            TabIndex        =   32
            Top             =   2040
            Width           =   975
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
            Left            =   2760
            TabIndex        =   60
            Tag             =   "13"
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   10
            Left            =   5520
            Tag             =   "12"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Ragrupp. Cliente"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   30
            Top             =   1680
            Width           =   1335
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
            Left            =   2760
            TabIndex        =   59
            Tag             =   "12"
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   5520
            Tag             =   "9"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Ente Emittente"
            Height          =   255
            Index           =   9
            Left            =   960
            TabIndex        =   28
            Top             =   1320
            Width           =   1095
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
            Left            =   2760
            TabIndex        =   58
            Tag             =   "9"
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   5520
            Tag             =   "8"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Società"
            Height          =   255
            Index           =   8
            Left            =   1320
            TabIndex        =   26
            Top             =   960
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
            Left            =   2760
            TabIndex        =   57
            Tag             =   "8"
            Top             =   960
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Index           =   4
         Left            =   360
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   8655
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   2520
            Picture         =   "OIL2ACL0.frx":0C18
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   3075
            Width           =   240
         End
         Begin VB.CommandButton Command9 
            Caption         =   "F8 - Agg. Dettaglio"
            Height          =   375
            Left            =   2400
            Picture         =   "OIL2ACL0.frx":0D1A
            TabIndex        =   68
            Tag             =   "REFRESHLIST"
            Top             =   3000
            Width           =   2295
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   240
            Picture         =   "OIL2ACL0.frx":0E1C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   3075
            Width           =   240
         End
         Begin VB.CommandButton Command10 
            Caption         =   "F7 - Apri Destinaz."
            Height          =   375
            Left            =   120
            Picture         =   "OIL2ACL0.frx":0F1E
            TabIndex        =   70
            Tag             =   "EDITLINK"
            Top             =   3000
            Width           =   2295
         End
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   2655
            Left            =   120
            OleObjectBlob   =   "OIL2ACL0.frx":1020
            TabIndex        =   65
            Top             =   240
            Width           =   8415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Index           =   3
         Left            =   360
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   8655
         Begin TrueDBGrid60.TDBGrid TDBGrid2 
            Height          =   3135
            Left            =   1080
            OleObjectBlob   =   "OIL2ACL0.frx":646A
            TabIndex        =   38
            Tag             =   "MULTISOC"
            Top             =   240
            Width           =   5415
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3975
         Left            =   240
         TabIndex        =   54
         Top             =   840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati A&nagrafici"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati A&mministrativi"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Codici &Per Società"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Destina&zioni"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   39
         Tag             =   "19"
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8520
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
               Picture         =   "OIL2ACL0.frx":A6E4
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ACL0.frx":A7F6
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ACL0.frx":A908
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3480
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2ACL0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GridArray As ArrayGridCls
Public WithEvents DataClass As CCliente
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private XA As XArrayDB
Private DetailClass As CDestinazione
Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.Bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.Bookmark, 0)
    End If
End Function

Public Sub DisplayDetail()
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    Dim strSQL As String
    
    With DetailClass
        strSQL = .SQL_GETLIST_CLIENT & .DS.DA.MakeSQL(DataClass.DS.Fields("0").value, DataClass.DS.Fields("0").FldType)
        If .DS.GetSelection(strSQL, False, , , False) Then
'        If .GetList Then
            If .DS.HowMany > 0 Then
                RefreshXArray XA, TDBGrid1, .DS.List
            End If
        End If
    End With
End Sub

Private Sub TrackLostFocus(mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(1) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsNone) Or (State = fsInsert) Then
                    If GetRecord(mskField1) Then
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
    ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcAddNew
            If (State = fsEdit) Or (State = fsNone) Then
                DataClass.DS.Fields.ResetFields
                ClearForm
                SetState fsInsert
            End If
        Case fcSave
            Select Case State
                Case fsEdit
                    blnResult = UpdateRecord
                Case fsInsert
                    blnResult = InsertRecord
                Case Else
                    blnResult = False
            End Select
            If blnResult Then
                SetState fsNone
            End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcPrint
            If (State = fsEdit) Or (State = fsNone) Then
            End If
        Case fcFind
            DoSearch "ACL", Me, "0", 1
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcANAGADZ)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .DataClass.DS.Fields("1").FindValue = DataClass.DS.Fields(DataClass.DS.KeyFields("1")).value
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcLinkRefresh
            DisplayDetail
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
                    DisplayDetail
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


'Private Sub SyncList(ByVal ID As String)
'    Dim I As Integer
'    Dim blnFound As Boolean
'    Dim varList As Variant
'
'    I = 0
'    blnFound = False
'    varList = DataClass.DS.List
'    Do While I <= UBound(varList, 2) And (Not blnFound)
'        If ID = varList(0, I) Then
'            blnFound = True
'        Else
'            I = I + 1
'        End If
'    Loop
'    If blnFound Then
'        TDBGrid1.Bookmark = I
'    End If
'End Sub

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
'                DisplayList
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
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
            mintCurFrame = 1
            TabStrip1.Enabled = False
            ClearUnboundGrid TDBGrid2, GridArray, MAX_COLS
        Case fsEdit
            TabStrip1.Enabled = True
            If TabStrip1.SelectedItem.Index = 1 Then
                mskField2.SetFocus
            End If
        Case fsInsert
            TabStrip1.Enabled = True
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
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
    mintCurFrame = 1
    Set GridArray = New ArrayGridCls
    Set DataClass = New CCliente
    Set DetailClass = New CDestinazione
    
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
'        Set DetailClass.DS.DA = DA
        Set DetailClass.DS.DA = DataConnect.DA
        Set DetailClass.DS.ErrH = ErrH
    End With
    
    TDBGrid1.Columns(5).Alignment = dbgCenter
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2ACL0 = Nothing
End Sub





Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub



Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame3(TabStrip1.SelectedItem.Index).Visible = True
        Frame3(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index

End Sub


Private Sub TDBGrid1_DblClick()
    DoCommand fcOpenLinkEdit
End Sub


