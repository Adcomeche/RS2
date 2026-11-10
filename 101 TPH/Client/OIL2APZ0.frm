VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form OIL2APZ0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PREZZI"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   10470
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   66
      Top             =   4560
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   2040
         Picture         =   "OIL2APZ0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1920
         Picture         =   "OIL2APZ0.frx":0102
         TabIndex        =   3
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
         Index           =   1
         Left            =   360
         Picture         =   "OIL2APZ0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   67
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
         Left            =   8640
         Picture         =   "OIL2APZ0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "OIL2APZ0.frx":0408
         TabIndex        =   4
         Tag             =   "EXIT"
         Top             =   345
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   240
         Picture         =   "OIL2APZ0.frx":050A
         TabIndex        =   2
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.Frame Frame6 
         Height          =   3255
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Width           =   9975
         Begin VB.CheckBox Check1 
            Caption         =   "&Sospeso"
            Height          =   255
            Index           =   0
            Left            =   6720
            TabIndex        =   24
            Tag             =   "16"
            Top             =   1320
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mskField2 
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Tag             =   "1"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField4 
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Tag             =   "3"
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField3 
            Height          =   255
            Left            =   1440
            TabIndex        =   7
            Tag             =   "2"
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField5 
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Tag             =   "4"
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField6 
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            Tag             =   "5"
            Top             =   1680
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskField7 
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Tag             =   "6"
            Top             =   2040
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox15 
            Height          =   255
            Left            =   6720
            TabIndex        =   19
            Tag             =   "9"
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16 
            Height          =   255
            Left            =   6720
            TabIndex        =   21
            Tag             =   "10"
            Top             =   600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   327681
            BackColor       =   12648447
            MaxLength       =   25
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Tag             =   "7"
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   1440
            TabIndex        =   18
            Tag             =   "8"
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox23 
            Height          =   255
            Left            =   6720
            TabIndex        =   23
            Tag             =   "15"
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0.0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   6720
            TabIndex        =   25
            Tag             =   "11"
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   6720
            TabIndex        =   26
            Tag             =   "12"
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   6720
            TabIndex        =   27
            Tag             =   "13"
            Top             =   2400
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   255
            Left            =   6720
            TabIndex        =   28
            Tag             =   "14"
            Top             =   2760
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   327681
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   14
            Left            =   7920
            Tag             =   "14"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   7920
            Tag             =   "13"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   7920
            Tag             =   "12"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   10
            Left            =   7920
            Tag             =   "11"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantità cisterne"
            Height          =   255
            Index           =   23
            Left            =   5160
            TabIndex        =   74
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Quantità Drop Medio"
            Height          =   255
            Index           =   24
            Left            =   5160
            TabIndex        =   73
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Quantità Minima"
            Height          =   255
            Index           =   25
            Left            =   5160
            TabIndex        =   72
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Capacità cisterne"
            Height          =   255
            Index           =   26
            Left            =   5160
            TabIndex        =   71
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   7920
            Tag             =   "15"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   9600
            Tag             =   "10"
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   9600
            Tag             =   "9"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   2880
            Tag             =   "8"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   2880
            Tag             =   "7"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   4560
            Tag             =   "6"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   4560
            Tag             =   "5"
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   3
            Left            =   4560
            Tag             =   "4"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   5040
            Tag             =   "3"
            Top             =   960
            Width           =   240
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   3960
            Top             =   2400
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   3
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "OIL2APZ0.frx":060C
                  Key             =   "GreenCheck"
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "OIL2APZ0.frx":071E
                  Key             =   "RedCross"
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "OIL2APZ0.frx":0830
                  Key             =   "Editing"
               EndProperty
            EndProperty
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   5040
            Tag             =   "2"
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   0
            Left            =   5040
            Tag             =   "1"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Giorni rifornimento"
            Height          =   255
            Index           =   20
            Left            =   5280
            TabIndex        =   70
            Top             =   960
            Width           =   1335
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
            Left            =   7440
            TabIndex        =   22
            Tag             =   "10"
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "UM Vendita"
            Height          =   255
            Index           =   19
            Left            =   5520
            TabIndex        =   64
            Top             =   600
            Width           =   1095
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
            Left            =   7440
            TabIndex        =   20
            Tag             =   "9"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "UM Acquisto"
            Height          =   255
            Index           =   18
            Left            =   5400
            TabIndex        =   63
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label6 
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
            TabIndex        =   16
            Tag             =   "6"
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fascia Prezzo"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Prezzo vendita"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   61
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Prezzo acquisto"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   60
            Top             =   2400
            Width           =   1215
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
            Left            =   2160
            TabIndex        =   14
            Tag             =   "5"
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "M. Pagamento"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   59
            Top             =   1680
            Width           =   1095
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
            Left            =   2160
            TabIndex        =   12
            Tag             =   "4"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Iva"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   58
            Top             =   1320
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
            TabIndex        =   8
            Tag             =   "2"
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Imballo"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   57
            Top             =   600
            Width           =   975
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
            Left            =   2640
            TabIndex        =   10
            Tag             =   "3"
            Top             =   960
            Width           =   2295
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
            Left            =   2640
            TabIndex        =   6
            Tag             =   "1"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "P. Doganale"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   54
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Prodotto"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3255
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   9975
         Begin VB.Frame Frame3 
            Caption         =   "Sconto in valuta"
            Height          =   1695
            Left            =   360
            TabIndex        =   48
            Top             =   480
            Width           =   2775
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   255
               Left            =   840
               TabIndex        =   30
               Tag             =   "20"
               Top             =   480
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox18 
               Height          =   255
               Left            =   840
               TabIndex        =   31
               Tag             =   "21"
               Top             =   840
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox19 
               Height          =   255
               Left            =   840
               TabIndex        =   32
               Tag             =   "22"
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.000"
               PromptChar      =   "_"
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   17
               Left            =   2400
               Tag             =   "22"
               Top             =   1200
               Width           =   240
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   16
               Left            =   2400
               Tag             =   "21"
               Top             =   840
               Width           =   240
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   15
               Left            =   2400
               Tag             =   "20"
               Top             =   480
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   51
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   50
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   49
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Sconto in Percentuale"
            Height          =   1695
            Left            =   3600
            TabIndex        =   44
            Top             =   480
            Width           =   2775
            Begin MSMask.MaskEdBox MaskEdBox12 
               Height          =   255
               Left            =   840
               TabIndex        =   33
               Tag             =   "17"
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox13 
               Height          =   255
               Left            =   840
               TabIndex        =   34
               Tag             =   "18"
               Top             =   840
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox14 
               Height          =   255
               Left            =   840
               TabIndex        =   35
               Tag             =   "19"
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.00"
               PromptChar      =   "_"
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   20
               Left            =   2280
               Tag             =   "19"
               Top             =   1200
               Width           =   240
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   19
               Left            =   2280
               Tag             =   "18"
               Top             =   840
               Width           =   240
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   18
               Left            =   2280
               Tag             =   "17"
               Top             =   480
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   47
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   46
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   45
               Top             =   1200
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Margine"
            Height          =   1695
            Left            =   6840
            TabIndex        =   38
            Top             =   480
            Width           =   2775
            Begin MSMask.MaskEdBox MaskEdBox20 
               Height          =   255
               Left            =   840
               TabIndex        =   37
               Tag             =   "23"
               Top             =   480
               Width           =   1445
               _ExtentX        =   2540
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox21 
               Height          =   255
               Left            =   840
               TabIndex        =   39
               Tag             =   "24"
               Top             =   840
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox22 
               Height          =   255
               Left            =   840
               TabIndex        =   41
               Tag             =   "25"
               Top             =   1200
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   450
               _Version        =   327681
               Format          =   "0.000"
               PromptChar      =   "_"
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   23
               Left            =   2400
               Tag             =   "25"
               Top             =   1200
               Width           =   240
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   22
               Left            =   2400
               Tag             =   "24"
               Top             =   840
               Width           =   240
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   21
               Left            =   2400
               Tag             =   "23"
               Top             =   480
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   43
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   42
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valore 1"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   40
               Top             =   480
               Width           =   615
            End
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   29
         Tag             =   "26"
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6588
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "    Dati Prezzi&1    "
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "    Dati Prezzi &2    "
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1680
         TabIndex        =   55
         Tag             =   "0"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   327681
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Destinazione"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   65
      Top             =   5475
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13018
            Key             =   "Msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10.00"
            Key             =   "Time"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16/12/99"
            Key             =   "Date"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2APZ0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Nome Form: OILAPZ0
' Descrizione: Form Prezzi
' Data Creazione: 17/11/99
' Autore: Fabio Mozzi

Option Explicit
Public WithEvents DataClass As Cprezzi
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private mintCurFrame As Integer
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private StrIDDest As String
Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub
Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    mskField1.Enabled = False
    Select Case NewState
        Case fsNone
               
              mskField1.Text = StrIDDest
              'mobjParentForm.RefreshDetail
              'mskField1.SetFocus
        Case fsEdit
             TabStrip1.Enabled = True
             'mskField1.SetFocus
        Case fsInsert
             ClearForm
             TabStrip1.Enabled = True
             mskField1.Text = StrIDDest
             mskField2.SetFocus
    End Select
End Sub
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
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
                SetState fsEdit
            End If
                        
        Case fcUndo
            If State = fsEdit Then
                SetState fsEdit
            ElseIf State = fsInsert Then
                SetState fsInsert
            End If
        Case fcExit
            Unload Me
    End Select
End Sub
Public Function GetRecord(StrIDADZ As Variant, StrIDAPR As Variant, StrIDAIM As Variant, StrIDAPD As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
         
        .Fields(.KeyFields(1)).FindValue = StrIDADZ
        .Fields(.KeyFields(2)).FindValue = StrIDAPR
        .Fields(.KeyFields(3)).FindValue = StrIDAIM
        .Fields(.KeyFields(4)).FindValue = StrIDAPD
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(0)).value, DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value, DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value
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

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub
Private Sub Form_Load()
    Dim i As Integer
    mintCurFrame = 1
    Loading = True
    
    Set DataClass = New Cprezzi
    With DataClass.DS
        Set .FormHandle = Me
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    Loading = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
    mobjParentForm.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mobjParentForm.DoCommand fcLinkRefresh
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2APZ0 = Nothing
End Sub
Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame _
            Then Exit Sub
        Frame6(TabStrip1.SelectedItem.Index).Visible = True
        Frame6(mintCurFrame).Visible = False
        mintCurFrame = TabStrip1.SelectedItem.Index
        'If mintCurFrame = 2 Then MaskEdBox2.SetFocus
End Sub
Public Property Let IDDestinazione(DEST As String)
     StrIDDest = DEST
End Property

