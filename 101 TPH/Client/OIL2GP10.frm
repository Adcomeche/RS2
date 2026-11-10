VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2GP10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PREFATTURA SOCIETA'"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   10470
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      Begin VB.Frame Frame3 
         Caption         =   "Riferimenti CEPSA"
         Height          =   3615
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   9975
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5760
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Tag             =   "38"
            Top             =   2880
            Width           =   2055
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5760
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Tag             =   "20"
            Top             =   360
            Width           =   2055
         End
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   255
            Left            =   1560
            TabIndex        =   48
            Tag             =   "19"
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   255
            Left            =   1560
            TabIndex        =   52
            Tag             =   "21"
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            Height          =   255
            Left            =   1560
            TabIndex        =   54
            Tag             =   "22"
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox14 
            Height          =   255
            Left            =   1560
            TabIndex        =   56
            Tag             =   "23"
            Top             =   1440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox15 
            Height          =   255
            Left            =   1560
            TabIndex        =   58
            Tag             =   "24"
            Top             =   1800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   255
            Left            =   1560
            TabIndex        =   60
            Tag             =   "25"
            Top             =   2160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox18 
            Height          =   255
            Left            =   1560
            TabIndex        =   62
            Tag             =   "26"
            Top             =   2520
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox19 
            Height          =   255
            Left            =   1560
            TabIndex        =   64
            Tag             =   "27"
            Top             =   2880
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   5760
            TabIndex        =   66
            Tag             =   "28"
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox21 
            Height          =   255
            Left            =   5760
            TabIndex        =   68
            Tag             =   "29"
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox22 
            Height          =   255
            Left            =   5760
            TabIndex        =   70
            Tag             =   "30"
            Top             =   1440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox24 
            Height          =   255
            Left            =   5760
            TabIndex        =   72
            Tag             =   "32"
            Top             =   1800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox25 
            Height          =   255
            Left            =   5760
            TabIndex        =   74
            Tag             =   "33"
            Top             =   2160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox26 
            Height          =   255
            Left            =   1560
            TabIndex        =   76
            Tag             =   "34"
            Top             =   3240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox23 
            Height          =   255
            Left            =   5760
            TabIndex        =   78
            Tag             =   "42"
            Top             =   2520
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Fattura"
            Height          =   255
            Index           =   30
            Left            =   4440
            TabIndex        =   77
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   30
            Left            =   7920
            Tag             =   "42"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   35
            Left            =   7920
            Tag             =   "38"
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Riga"
            Height          =   255
            Index           =   35
            Left            =   4680
            TabIndex        =   83
            Top             =   2880
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   33
            Left            =   7920
            Tag             =   "34"
            Top             =   3240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Imbarcazione"
            Height          =   255
            Index           =   33
            Left            =   240
            TabIndex        =   79
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   32
            Left            =   7920
            Tag             =   "33"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Ticket"
            Height          =   255
            Index           =   32
            Left            =   4440
            TabIndex        =   75
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   31
            Left            =   7920
            Tag             =   "32"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Consegna"
            Height          =   255
            Index           =   31
            Left            =   4440
            TabIndex        =   73
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   29
            Left            =   7920
            Tag             =   "30"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cod. Fornitore"
            Height          =   255
            Index           =   29
            Left            =   4440
            TabIndex        =   71
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   28
            Left            =   7920
            Tag             =   "29"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cod. Operatore"
            Height          =   255
            Index           =   28
            Left            =   4440
            TabIndex        =   69
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   26
            Left            =   7920
            Tag             =   "28"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Num. Viaggio"
            Height          =   255
            Index           =   26
            Left            =   4440
            TabIndex        =   67
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   25
            Left            =   3720
            Tag             =   "27"
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "NIF Autista"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   65
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   24
            Left            =   3720
            Tag             =   "26"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Autista"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   63
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   23
            Left            =   3720
            Tag             =   "25"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Targa Rimorchio"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   61
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   22
            Left            =   3720
            Tag             =   "24"
            Top             =   1800
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Targa Motrice"
            Height          =   255
            Index           =   22
            Left            =   360
            TabIndex        =   59
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   21
            Left            =   3720
            Tag             =   "23"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. UdT (Dep.)"
            Height          =   255
            Index           =   21
            Left            =   360
            TabIndex        =   57
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   20
            Left            =   3720
            Tag             =   "22"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. UdT"
            Height          =   255
            Index           =   20
            Left            =   480
            TabIndex        =   55
            Top             =   1080
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   19
            Left            =   3720
            Tag             =   "21"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Deposito"
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   53
            Top             =   720
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   27
            Left            =   7920
            Tag             =   "20"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Servizio"
            Height          =   255
            Index           =   27
            Left            =   4680
            TabIndex        =   51
            Top             =   360
            Width           =   975
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   3720
            Tag             =   "19"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Vettore"
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Riferimenti AGIP"
         Height          =   3615
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   9975
         Begin MSMask.MaskEdBox mskField7 
            Height          =   255
            Left            =   1800
            TabIndex        =   10
            Tag             =   "6"
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imp&orto Forzato"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   12
            Tag             =   "14"
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   1800
            TabIndex        =   94
            Tag             =   "4"
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   1800
            TabIndex        =   96
            Tag             =   "15"
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   3840
            Tag             =   "15"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Riferimento esterno"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Bolletta di scarico"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   95
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   3840
            Tag             =   "4"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   3840
            Tag             =   "6"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Riferimento esterno 2"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3615
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   9975
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   7560
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Tag             =   "36"
            Top             =   240
            Width           =   1935
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Tag             =   "9"
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Tag             =   "10"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Tag             =   "11"
            Top             =   600
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField9 
            Height          =   255
            Left            =   1560
            TabIndex        =   22
            Tag             =   "1"
            Top             =   1320
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
            Left            =   1560
            TabIndex        =   23
            Tag             =   "0"
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
            Left            =   1560
            TabIndex        =   28
            Tag             =   "12"
            Top             =   2040
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
            Left            =   4680
            TabIndex        =   31
            Tag             =   "5"
            Top             =   2400
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField4 
            Height          =   255
            Left            =   7800
            TabIndex        =   32
            Tag             =   "8"
            Top             =   2760
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField8 
            Height          =   255
            Left            =   7800
            TabIndex        =   33
            Tag             =   "7"
            Top             =   2400
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            Format          =   "0.000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   1560
            TabIndex        =   34
            Tag             =   "3"
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
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   1560
            TabIndex        =   35
            Tag             =   "13"
            Top             =   3120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Tag             =   "16"
            Top             =   2400
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   4680
            TabIndex        =   43
            Tag             =   "35"
            Top             =   2760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   7800
            TabIndex        =   45
            Tag             =   "37"
            Top             =   3120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox27 
            Height          =   255
            Left            =   6120
            TabIndex        =   85
            Tag             =   "39"
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox28 
            Height          =   255
            Left            =   6120
            TabIndex        =   88
            Tag             =   "40"
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox29 
            Height          =   255
            Left            =   6120
            TabIndex        =   91
            Tag             =   "41"
            Top             =   2040
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
            Left            =   7560
            TabIndex        =   98
            Tag             =   "2"
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rif. Ordine"
            Height          =   255
            Index           =   2
            Left            =   6600
            TabIndex        =   99
            Top             =   600
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   9600
            Tag             =   "2"
            Top             =   600
            Width           =   240
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
            Index           =   4
            Left            =   6720
            TabIndex        =   93
            Tag             =   "41"
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Autista"
            Height          =   255
            Index           =   39
            Left            =   5280
            TabIndex        =   92
            Top             =   2040
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   39
            Left            =   9600
            Tag             =   "41"
            Top             =   2040
            Width           =   240
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
            Index           =   3
            Left            =   7560
            TabIndex        =   90
            Tag             =   "40"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rimorchio"
            Height          =   255
            Index           =   38
            Left            =   5280
            TabIndex        =   89
            Top             =   1680
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   38
            Left            =   9600
            Tag             =   "40"
            Top             =   1680
            Width           =   240
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
            Index           =   2
            Left            =   7560
            TabIndex        =   87
            Tag             =   "39"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Motrice"
            Height          =   255
            Index           =   37
            Left            =   5280
            TabIndex        =   86
            Top             =   1320
            Width           =   735
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   37
            Left            =   9600
            Tag             =   "39"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   34
            Left            =   9600
            Tag             =   "36"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo Tariffa"
            Height          =   255
            Index           =   34
            Left            =   6600
            TabIndex        =   81
            Top             =   240
            Width           =   855
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   17
            Left            =   9600
            Tag             =   "37"
            Top             =   3120
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   9600
            Tag             =   "8"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Importo Pedaggi"
            Height          =   255
            Index           =   17
            Left            =   6360
            TabIndex        =   46
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   16
            Left            =   6000
            Tag             =   "35"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Q.ta scaricata (Kg)"
            Height          =   255
            Index           =   16
            Left            =   3120
            TabIndex        =   44
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   6000
            Tag             =   "5"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Q.ta scaricata (L)"
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   42
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Importo Trasporto"
            Height          =   255
            Index           =   5
            Left            =   6360
            TabIndex        =   41
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   9600
            Tag             =   "7"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tariffa Appllicata"
            Height          =   255
            Index           =   7
            Left            =   6480
            TabIndex        =   40
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Data Consegna"
            Height          =   255
            Index           =   36
            Left            =   240
            TabIndex        =   39
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   36
            Left            =   2880
            Tag             =   "3"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Distanza Km"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   38
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   2880
            Tag             =   "13"
            Top             =   3120
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Anno prefattura"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   37
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   10
            Left            =   2880
            Tag             =   "16"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   0
            Left            =   4920
            Tag             =   "12"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Unità di Trasporto"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   2040
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
            Left            =   2160
            TabIndex        =   29
            Tag             =   "12"
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   4920
            Tag             =   "1"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Deposito"
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   27
            Top             =   1320
            Width           =   855
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
            Left            =   2160
            TabIndex        =   26
            Tag             =   "1"
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   4920
            Tag             =   "0"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Zona"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   25
            Top             =   1680
            Width           =   975
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
            Index           =   1
            Left            =   2160
            TabIndex        =   24
            Tag             =   "0"
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Prodotto"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   21
            Top             =   960
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
            Left            =   2760
            TabIndex        =   20
            Tag             =   "9"
            Top             =   960
            Width           =   3375
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   6240
            Tag             =   "10"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "D&estinazione"
            Height          =   255
            Index           =   11
            Left            =   480
            TabIndex        =   19
            Top             =   240
            Width           =   975
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
            Index           =   0
            Left            =   3000
            TabIndex        =   18
            Tag             =   "10"
            Top             =   240
            Width           =   3135
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   6240
            Tag             =   "9"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ragione Sociale"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   6240
            Tag             =   "11"
            Top             =   600
            Width           =   240
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   0
         Tag             =   "17"
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
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
               Picture         =   "OIL2GP10.frx":0000
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2GP10.frx":0112
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2GP10.frx":0224
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Tag             =   "18"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4095
         Left            =   120
         TabIndex        =   84
         Top             =   600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7223
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Prefattura"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dati Prefattura 2"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
         TabIndex        =   8
         Tag             =   "18"
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   15
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   15
         Left            =   6720
         Tag             =   "18"
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   8760
         Picture         =   "OIL2GP10.frx":0336
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8640
         Picture         =   "OIL2GP10.frx":0438
         TabIndex        =   5
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   10470
      _ExtentX        =   18468
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
            Object.Width           =   13018
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23.31"
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
Attribute VB_Name = "OIL2GP10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DataClass As CPrefatturaSoc
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mobjParentForm As Object
Private mintCurFrame As Integer

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1





Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcExit
            Unload Me
    End Select
End Sub

Public Sub GetRecord(strID As String, strIDProdotto As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        .Fields(.KeyFields(2)).FindValue = strIDProdotto
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
'                    For i = 1 To Frame3.UBound
'                        Frame3(i).Visible = False
'                    Next i
'                    Select Case .Fields("18").value
'                        Case "001"  ' AGIP
'                            Frame3(2).Visible = True
'                        Case "038"  ' CEPSA
'                            Frame3(3).Visible = True
'                    End Select
                End If
            Else
                GoTo ErrHandler
            End If
        Else
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
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

'Private Function UpdateRecord() As Boolean
'    On Error GoTo ErrHandler
'
'    Dim blnChkValue As Boolean
'    Dim blnConcurError As Boolean
'    Dim C As Control
'
'    'Scansione della finestra
'    DataClass.DS.ScanClient 0
'
'    If DataClass.DS.Fields.Validate Then
'        DataClass.BeginSave blnConcurError
'        If blnConcurError Then
'            If MsgBox(MSG_CONCUR_ERR, _
'                       vbOKCancel + vbExclamation + vbDefaultButton2, _
'                      "Attenzione") = vbOK Then
'                GoTo EndProc
'            Else
'                DataClass.DS.DA.EndTrans etRollback
'                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
''                DisplayList
'                GoTo ErrHandler
'            End If
'        Else
'            GoTo EndProc
'        End If
'        GoTo EndProc
'    Else
'        ShowValidationError DataClass.DS
'        GoTo ErrHandler
'    End If
'
'ErrHandler:
'    UpdateRecord = False
'    Exit Function
'EndProc:
'    DataClass.DS.DA.EndTrans etCommit
'    DataClass.DS.EndSave
'    UpdateRecord = True
'End Function
'
Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsNone
            If Not Loading Then
            End If
        Case fsEdit
        Case fsInsert
    End Select
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
    
    Loading = True
    
    mintCurFrame = 1
    Set DataClass = New CPrefatturaSoc
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
        End With
    End If
    Set OIL2GP10 = Nothing
End Sub







Private Sub TabStrip1_Click()
    Dim intCurrentASC As Integer
    Dim i As Integer
    
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        For i = 1 To Frame3.UBound
            Frame3(i).Visible = False
        Next i
        If TabStrip1.SelectedItem.Index = 1 Then
            Frame3(TabStrip1.SelectedItem.Index).Visible = True
        Else
            Select Case DataClass.DS.Fields("18").value
                Case "001"  ' AGIP
                    intCurrentASC = 2
                Case "038"  ' CEPSA
                    intCurrentASC = 3
            End Select
            Frame3(intCurrentASC).Visible = True
        End If
        mintCurFrame = TabStrip1.SelectedItem.Index
End Sub


