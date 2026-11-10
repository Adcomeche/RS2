VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTSnapshotExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRASMISSIONE DATI CONSEGNE"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   12105
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   101
      Top             =   9120
      Width           =   12135
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   2160
         Picture         =   "frmVVTSnapshotExport.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   780
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Elabora Anomalie"
         Height          =   375
         Left            =   1800
         Picture         =   "frmVVTSnapshotExport.frx":0102
         TabIndex        =   66
         Tag             =   "SAVE"
         Top             =   720
         Width           =   3370
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   10440
         Picture         =   "frmVVTSnapshotExport.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   780
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton cmdStartStopAutomatic 
         Caption         =   "Stop"
         Height          =   375
         Left            =   10320
         Picture         =   "frmVVTSnapshotExport.frx":0306
         TabIndex        =   67
         Tag             =   "CUSTOM2"
         Top             =   720
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
         Picture         =   "frmVVTSnapshotExport.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSnapshotExport.frx":050A
         TabIndex        =   65
         Tag             =   "PRINT"
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   3480
         Picture         =   "frmVVTSnapshotExport.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   435
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
         Left            =   10440
         Picture         =   "frmVVTSnapshotExport.frx":070E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   10320
         Picture         =   "frmVVTSnapshotExport.frx":0810
         TabIndex        =   64
         Tag             =   "EXIT"
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
         Index           =   6
         Left            =   8760
         Picture         =   "frmVVTSnapshotExport.frx":0912
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Transf. INC."
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         Picture         =   "frmVVTSnapshotExport.frx":0A14
         TabIndex        =   63
         Tag             =   "EXPORT"
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
         Index           =   0
         Left            =   6960
         Picture         =   "frmVVTSnapshotExport.frx":0B16
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         Picture         =   "frmVVTSnapshotExport.frx":0C18
         TabIndex        =   62
         Tag             =   "BROWSE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   5280
         Picture         =   "frmVVTSnapshotExport.frx":0D1A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton CmdOpenTrip 
         Caption         =   "Apri &Viaggio"
         Height          =   375
         Left            =   5160
         Picture         =   "frmVVTSnapshotExport.frx":0E1C
         TabIndex        =   61
         Tag             =   "OPEN"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F11 - Modifica"
         Height          =   375
         Left            =   3360
         Picture         =   "frmVVTSnapshotExport.frx":0F1E
         TabIndex        =   60
         Tag             =   "CUSTOM1"
         Top             =   360
         Width           =   1815
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
         Picture         =   "frmVVTSnapshotExport.frx":1020
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - &Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmVVTSnapshotExport.frx":1122
         TabIndex        =   59
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1575
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
         Picture         =   "frmVVTSnapshotExport.frx":1224
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - &Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSnapshotExport.frx":1326
         TabIndex        =   58
         Tag             =   "FIND"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Timer TimerAutomaticSend 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   10380
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   100
      Top             =   10365
      Width           =   12105
      _ExtentX        =   21352
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
            Object.Width           =   15901
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
   Begin VB.Frame Frame1 
      Caption         =   "Elenco consegne"
      Height          =   8535
      Index           =   2
      Left            =   120
      TabIndex        =   105
      Top             =   480
      Visible         =   0   'False
      Width           =   11895
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1575
         Left            =   0
         OleObjectBlob   =   "frmVVTSnapshotExport.frx":1428
         TabIndex        =   95
         Top             =   240
         Width           =   11775
      End
      Begin VB.Frame FraModificaViaggio 
         Caption         =   "Modifica Viaggio"
         Height          =   6675
         Left            =   0
         MouseIcon       =   "frmVVTSnapshotExport.frx":19DF6
         TabIndex        =   120
         Top             =   1800
         Width           =   11655
         Begin VB.TextBox txt_idviag 
            Height          =   375
            Left            =   120
            TabIndex        =   212
            Text            =   "Text1"
            Top             =   5280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txt_cod 
            Height          =   375
            Left            =   720
            TabIndex        =   211
            Text            =   "Text2"
            Top             =   5280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txt_CodEv 
            Height          =   315
            Left            =   1320
            TabIndex        =   210
            Text            =   "Text1"
            Top             =   5280
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   3
            Left            =   6960
            Picture         =   "frmVVTSnapshotExport.frx":1A238
            Style           =   1  'Graphical
            TabIndex        =   209
            Top             =   5760
            Width           =   375
         End
         Begin TrueDBGrid60.TDBGrid TDBGrid2 
            Height          =   2055
            Left            =   7440
            OleObjectBlob   =   "frmVVTSnapshotExport.frx":1A33A
            TabIndex        =   208
            Tag             =   "SELECT"
            Top             =   4560
            Width           =   4215
         End
         Begin VB.CommandButton btn_generaevento 
            Caption         =   "Genera Evento"
            Height          =   375
            Left            =   2400
            Picture         =   "frmVVTSnapshotExport.frx":227D4
            TabIndex        =   201
            Top             =   6120
            Width           =   2055
         End
         Begin VB.ComboBox cmb_cges 
            Height          =   315
            Left            =   1440
            TabIndex        =   200
            Text            =   "Selecionar Tipo Evento:"
            Top             =   5760
            Width           =   3015
         End
         Begin VB.CheckBox chkNoGPSData 
            Caption         =   "Parar datos PDA/SMS"
            Height          =   255
            Left            =   9360
            TabIndex        =   51
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox chkFlagEventi 
            Caption         =   "Eventi presenti"
            Height          =   255
            Left            =   9360
            TabIndex        =   52
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox TxtNoteEstese 
            Height          =   735
            Left            =   9360
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton CmdRitorna 
            Height          =   375
            Left            =   10920
            Picture         =   "frmVVTSnapshotExport.frx":228D6
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Ritorna"
            Top             =   2640
            Width           =   615
         End
         Begin VB.CommandButton cmdEventi 
            Height          =   375
            Left            =   10320
            Picture         =   "frmVVTSnapshotExport.frx":229D8
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Gestione Eventi"
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton CmdAnnulla 
            Height          =   375
            Left            =   9840
            Picture         =   "frmVVTSnapshotExport.frx":22ADA
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Annulla"
            Top             =   2640
            Width           =   495
         End
         Begin VB.CommandButton CmdSalva 
            Height          =   375
            Left            =   9360
            Picture         =   "frmVVTSnapshotExport.frx":2301C
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "F2 - Salva"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Frame Frame5 
            Caption         =   "Consegna"
            Height          =   5415
            Left            =   4440
            TabIndex        =   122
            Top             =   240
            Width           =   4815
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               Caption         =   "Cons. Terminata"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1680
               TabIndex        =   23
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Cons. Iniziata"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   1335
            End
            Begin MSMask.MaskEdBox MaskEdBox9 
               Height          =   255
               Left            =   4080
               TabIndex        =   30
               Top             =   1440
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox10 
               Height          =   255
               Left            =   4080
               TabIndex        =   24
               Top             =   360
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox24 
               Height          =   255
               Left            =   4080
               TabIndex        =   33
               Top             =   2160
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox25 
               Height          =   255
               Left            =   1080
               TabIndex        =   34
               Top             =   2520
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   255
               Left            =   1080
               TabIndex        =   32
               Top             =   2160
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox27 
               Height          =   255
               Left            =   1080
               TabIndex        =   25
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox30 
               Height          =   255
               Left            =   4080
               TabIndex        =   28
               Top             =   1080
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox35 
               Height          =   255
               Left            =   1080
               TabIndex        =   37
               Top             =   3240
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox36 
               Height          =   255
               Left            =   1800
               TabIndex        =   39
               Top             =   3600
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox37 
               Height          =   255
               Left            =   1080
               TabIndex        =   38
               Top             =   3600
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox38 
               Height          =   255
               Left            =   4080
               TabIndex        =   41
               Top             =   3600
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox39 
               Height          =   255
               Left            =   3360
               TabIndex        =   40
               Top             =   3600
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox31 
               Height          =   255
               Left            =   1080
               TabIndex        =   42
               Top             =   3960
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox40 
               Height          =   255
               Left            =   4080
               TabIndex        =   43
               Top             =   3960
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox41 
               Height          =   255
               Left            =   1080
               TabIndex        =   35
               Top             =   2880
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox43 
               Height          =   255
               Left            =   4080
               TabIndex        =   31
               Top             =   1800
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox47 
               Height          =   255
               Left            =   4080
               TabIndex        =   26
               Top             =   720
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox48 
               Height          =   255
               Left            =   1080
               TabIndex        =   27
               Top             =   1080
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox49 
               Height          =   255
               Left            =   1080
               TabIndex        =   29
               Top             =   1440
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox50 
               Height          =   255
               Left            =   4080
               TabIndex        =   36
               Top             =   2880
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox51 
               Height          =   255
               Left            =   1080
               TabIndex        =   44
               Top             =   4320
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox52 
               Height          =   255
               Left            =   1080
               TabIndex        =   45
               Top             =   4680
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox53 
               Height          =   255
               Left            =   1080
               TabIndex        =   46
               Top             =   5040
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tel. Movil"
               Height          =   255
               Index           =   56
               Left            =   240
               TabIndex        =   196
               Top             =   5040
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tel. 2"
               Height          =   255
               Index           =   55
               Left            =   240
               TabIndex        =   195
               Top             =   4680
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tel. 1"
               Height          =   255
               Index           =   54
               Left            =   240
               TabIndex        =   194
               Top             =   4320
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo Servicio"
               Height          =   255
               Index           =   53
               Left            =   2520
               TabIndex        =   193
               Top             =   2880
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Data arr."
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   192
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Data prev."
               Height          =   255
               Index           =   51
               Left            =   120
               TabIndex        =   191
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ora progr."
               Height          =   255
               Index           =   50
               Left            =   3240
               TabIndex        =   190
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Note"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   132
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ora fine"
               Height          =   255
               Index           =   45
               Left            =   3360
               TabIndex        =   181
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "NIS D"
               Height          =   255
               Index           =   42
               Left            =   3240
               TabIndex        =   178
               Top             =   3960
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Note Ord."
               Height          =   255
               Index           =   41
               Left            =   240
               TabIndex        =   175
               Top             =   3240
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ora prev."
               Height          =   255
               Index           =   26
               Left            =   3240
               TabIndex        =   160
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Data progr."
               Height          =   255
               Index           =   25
               Left            =   120
               TabIndex        =   157
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Km "
               Height          =   255
               Index           =   21
               Left            =   720
               TabIndex        =   152
               Top             =   2160
               Width           =   255
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Km Soc."
               Height          =   255
               Index           =   10
               Left            =   3120
               TabIndex        =   131
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Progr."
               Height          =   255
               Index           =   9
               Left            =   3480
               TabIndex        =   130
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ora arr."
               Height          =   255
               Index           =   6
               Left            =   3360
               TabIndex        =   129
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "NIS C"
               Height          =   255
               Index           =   28
               Left            =   240
               TabIndex        =   177
               Top             =   3960
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Rif. Ord."
               Height          =   255
               Index           =   43
               Left            =   240
               TabIndex        =   179
               Top             =   2880
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Viaggio"
            Height          =   5055
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Width           =   4215
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   "Carico Terminato"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2520
               TabIndex        =   2
               Top             =   240
               Width           =   1575
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               Caption         =   "Carico Iniziato"
               Enabled         =   0   'False
               Height          =   255
               Left            =   360
               TabIndex        =   1
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkFestivo 
               Caption         =   "Festivo"
               Height          =   255
               Left            =   3240
               TabIndex        =   20
               Top             =   3960
               Width           =   855
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   270
               Left            =   1080
               TabIndex        =   5
               Top             =   960
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   476
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox3 
               Height          =   255
               Left            =   1080
               TabIndex        =   10
               Top             =   2040
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox4 
               Height          =   255
               Left            =   1080
               TabIndex        =   8
               Top             =   1680
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox5 
               Height          =   255
               Left            =   3120
               TabIndex        =   11
               Top             =   2040
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox6 
               Height          =   255
               Left            =   1080
               TabIndex        =   12
               Tag             =   "34"
               Top             =   2400
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
               Left            =   3120
               TabIndex        =   9
               Top             =   1680
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox8 
               Height          =   255
               Left            =   1080
               TabIndex        =   16
               Top             =   3660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   18
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox11 
               Height          =   255
               Left            =   2400
               TabIndex        =   19
               Top             =   3960
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox18 
               Height          =   255
               Left            =   1080
               TabIndex        =   18
               Top             =   3960
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox19 
               Height          =   255
               Left            =   3360
               TabIndex        =   17
               Top             =   3660
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox32 
               Height          =   255
               Left            =   1080
               TabIndex        =   14
               Tag             =   "32"
               Top             =   3000
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   12648447
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox33 
               Height          =   255
               Left            =   1080
               TabIndex        =   15
               Tag             =   "39"
               Top             =   3300
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   12648447
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox34 
               Height          =   255
               Left            =   1080
               TabIndex        =   13
               Top             =   2700
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   40
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox42 
               Height          =   270
               Left            =   3480
               TabIndex        =   6
               Top             =   960
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   476
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox46 
               Height          =   255
               Left            =   2880
               TabIndex        =   7
               Top             =   1320
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox54 
               Height          =   255
               Left            =   3600
               TabIndex        =   21
               Top             =   4320
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox55 
               Height          =   255
               Left            =   1080
               TabIndex        =   3
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox56 
               Height          =   255
               Left            =   3480
               TabIndex        =   4
               Top             =   600
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               MaxLength       =   5
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Ora progr."
               Height          =   255
               Index           =   59
               Left            =   2640
               TabIndex        =   199
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Data progr."
               Height          =   255
               Index           =   58
               Left            =   120
               TabIndex        =   198
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Sequenza di Invio"
               Height          =   255
               Index           =   57
               Left            =   2040
               TabIndex        =   197
               Top             =   4320
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Data Fine Carico"
               Height          =   255
               Index           =   49
               Left            =   1560
               TabIndex        =   189
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Ora Fine Carico"
               Height          =   255
               Index           =   44
               Left            =   2280
               TabIndex        =   180
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tel."
               Height          =   255
               Index           =   39
               Left            =   240
               TabIndex        =   176
               Top             =   2700
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Vet.Tras."
               Height          =   255
               Index           =   36
               Left            =   180
               TabIndex        =   164
               Top             =   3300
               Width           =   795
            End
            Begin VB.Label Label10 
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
               Left            =   1680
               TabIndex        =   163
               Tag             =   "39"
               Top             =   3300
               Width           =   2055
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   4
               Left            =   3840
               Tag             =   "39"
               Top             =   3300
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Vettore"
               Height          =   255
               Index           =   34
               Left            =   360
               TabIndex        =   162
               Top             =   3000
               Width           =   615
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
               Left            =   1680
               TabIndex        =   161
               Tag             =   "32"
               Top             =   3000
               Width           =   2055
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   3
               Left            =   3840
               Tag             =   "32"
               Top             =   3000
               Width           =   240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Pedaggi"
               Height          =   255
               Index           =   40
               Left            =   2520
               TabIndex        =   155
               Top             =   3660
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Km "
               Height          =   255
               Index           =   22
               Left            =   720
               TabIndex        =   154
               Top             =   3960
               Width           =   255
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Progr."
               Height          =   255
               Index           =   19
               Left            =   1800
               TabIndex        =   151
               Top             =   3960
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "N.I.F."
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   150
               Top             =   3660
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Badge"
               Height          =   255
               Index           =   18
               Left            =   2400
               TabIndex        =   149
               Top             =   1680
               Width           =   615
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   0
               Left            =   3840
               Tag             =   "34"
               Top             =   2400
               Width           =   240
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
               Left            =   1680
               TabIndex        =   145
               Tag             =   "34"
               Top             =   2400
               Width           =   2055
            End
            Begin VB.Image imgState 
               Height          =   240
               Index           =   21
               Left            =   4800
               Tag             =   "13"
               Top             =   0
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "Autista 1"
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   128
               Top             =   2400
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Rimorchio"
               Height          =   255
               Index           =   3
               Left            =   2280
               TabIndex        =   127
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "UdT"
               Height          =   255
               Index           =   2
               Left            =   600
               TabIndex        =   126
               Top             =   1680
               Width           =   375
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Motrice"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   125
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Ora Carico"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   124
               Top             =   960
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Quantità"
            Height          =   1695
            Left            =   9360
            TabIndex        =   123
            Top             =   240
            Width           =   2175
            Begin VB.CheckBox chkNoConforme 
               Alignment       =   1  'Right Justify
               Caption         =   "Non conformità cliente"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CheckBox chkDesigillato 
               Alignment       =   1  'Right Justify
               Caption         =   "Desigillato"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   1320
               Width           =   1935
            End
            Begin MSMask.MaskEdBox MaskEdBox28 
               Height          =   255
               Left            =   1200
               TabIndex        =   47
               Top             =   360
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   15
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox29 
               Height          =   255
               Left            =   1200
               TabIndex        =   48
               Top             =   720
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   15
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Scaricata Litri"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   134
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Caricata Litri"
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   133
               Top             =   360
               Width           =   975
            End
         End
         Begin MSMask.MaskEdBox txt_dataev 
            Height          =   255
            Left            =   5760
            TabIndex        =   202
            Top             =   5760
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_hrev 
            Height          =   255
            Left            =   5760
            TabIndex        =   203
            Top             =   6120
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:nn"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Generar evento:"
            Height          =   255
            Index           =   62
            Left            =   -600
            TabIndex        =   206
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data evento"
            Height          =   255
            Index           =   60
            Left            =   4680
            TabIndex        =   205
            Top             =   5760
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora evento"
            Height          =   255
            Index           =   61
            Left            =   4680
            TabIndex        =   204
            Top             =   6120
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Note Est."
            Height          =   255
            Index           =   13
            Left            =   9360
            TabIndex        =   174
            Top             =   3120
            Width           =   735
         End
      End
      Begin VB.Frame FraTrasmettiViaggi 
         Caption         =   "Trasmetti"
         Height          =   5715
         Left            =   120
         TabIndex        =   135
         Top             =   1800
         Width           =   11655
         Begin VB.Frame fraTimer 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   315
            Left            =   9360
            TabIndex        =   167
            Top             =   1080
            Width           =   2175
            Begin MSMask.MaskEdBox mskMinutes 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "99"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1040
                  SubFormatType   =   0
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   168
               Top             =   0
               Visible         =   0   'False
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Minuti"
               Height          =   255
               Index           =   37
               Left            =   60
               TabIndex        =   170
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lbTimer 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   780
               TabIndex        =   169
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1335
            Index           =   0
            Left            =   7560
            TabIndex        =   136
            Top             =   2040
            Visible         =   0   'False
            Width           =   3975
            Begin VB.Label lblLog 
               Alignment       =   1  'Right Justify
               Caption         =   "Record trasmessi correttamente:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   142
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label lblLog 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   141
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblLog 
               Alignment       =   1  'Right Justify
               Caption         =   "Errori riscontrati:"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   140
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblLog 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   2520
               TabIndex        =   139
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblLog 
               Alignment       =   1  'Right Justify
               Caption         =   "Record da trasmettere:"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   138
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label lblLog 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   2520
               TabIndex        =   137
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox tboxLog 
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   165
            Top             =   3480
            Visible         =   0   'False
            Width           =   10575
         End
         Begin VB.CheckBox chkTransAll 
            Caption         =   "Trasmetti tutti"
            Height          =   255
            Left            =   7560
            TabIndex        =   158
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   10680
            Top             =   3720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame Frame3 
            Height          =   3135
            Left            =   120
            TabIndex        =   143
            Top             =   240
            Visible         =   0   'False
            Width           =   7335
            Begin TrueDBGrid60.TDBGrid tdbLog 
               Height          =   2775
               Left            =   120
               OleObjectBlob   =   "frmVVTSnapshotExport.frx":2311E
               TabIndex        =   96
               Top             =   240
               Visible         =   0   'False
               Width           =   7095
            End
         End
         Begin MSMask.MaskEdBox mskFileName 
            Height          =   255
            Left            =   8760
            TabIndex        =   97
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOrario 
            Height          =   255
            Index           =   4
            Left            =   10920
            TabIndex        =   182
            Top             =   720
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOrario 
            Height          =   255
            Index           =   1
            Left            =   8760
            TabIndex        =   183
            Top             =   720
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOrario 
            Height          =   255
            Index           =   2
            Left            =   9480
            TabIndex        =   184
            Top             =   720
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOrario 
            Height          =   255
            Index           =   3
            Left            =   10200
            TabIndex        =   185
            Top             =   720
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Orari:"
            Height          =   255
            Index           =   46
            Left            =   7560
            TabIndex        =   186
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Percorso &File:"
            Height          =   255
            Index           =   14
            Left            =   7560
            TabIndex        =   144
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      Height          =   7575
      Index           =   1
      Left            =   120
      TabIndex        =   99
      Top             =   480
      Width           =   11895
      Begin VB.CommandButton btn_trtutt 
         Caption         =   "Trasmetti Tutti"
         Height          =   375
         Left            =   6360
         TabIndex        =   207
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   3480
         Picture         =   "frmVVTSnapshotExport.frx":2670E
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   9000
         Picture         =   "frmVVTSnapshotExport.frx":26810
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3720
         Width           =   375
      End
      Begin VB.CheckBox chkOnlyChangedTrip 
         Caption         =   "Solo viaggi modificati"
         Height          =   255
         Left            =   9240
         TabIndex        =   85
         Top             =   1020
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Tag             =   "9"
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   4
         Left            =   9000
         Picture         =   "frmVVTSnapshotExport.frx":26912
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   4080
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   3480
         Picture         =   "frmVVTSnapshotExport.frx":26A14
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   4080
         Width           =   375
      End
      Begin VB.CheckBox chkActiveOnly 
         Caption         =   "Solo viaggi non annullati"
         Height          =   255
         Left            =   9240
         TabIndex        =   84
         Top             =   720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskMaxRecords 
         Height          =   300
         Left            =   10200
         TabIndex        =   83
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10080
         Top             =   3840
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
               Picture         =   "frmVVTSnapshotExport.frx":26B16
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSnapshotExport.frx":26C28
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSnapshotExport.frx":26D3A
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   2040
         TabIndex        =   74
         Tag             =   "0"
         Top             =   3000
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
         TabIndex        =   71
         Tag             =   "3"
         Top             =   1920
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
         Left            =   2040
         TabIndex        =   80
         Tag             =   "101"
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   7560
         TabIndex        =   75
         Tag             =   "1"
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   7560
         TabIndex        =   81
         Tag             =   "100"
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   255
         Left            =   2040
         TabIndex        =   76
         Tag             =   "11"
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   7560
         TabIndex        =   77
         Tag             =   "12"
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox20 
         Height          =   255
         Left            =   2040
         TabIndex        =   0
         Tag             =   "92"
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox21 
         Height          =   255
         Left            =   2040
         TabIndex        =   68
         Tag             =   "93"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox22 
         Height          =   255
         Left            =   2040
         TabIndex        =   69
         Tag             =   "7"
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox23 
         Height          =   255
         Left            =   2040
         TabIndex        =   70
         Tag             =   "17"
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox26 
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Tag             =   "13"
         Top             =   2640
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
         Left            =   2040
         TabIndex        =   72
         Tag             =   "4"
         Top             =   2280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox44 
         Height          =   255
         Left            =   2040
         TabIndex        =   78
         Tag             =   "97"
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox45 
         Height          =   255
         Left            =   7560
         TabIndex        =   79
         Tag             =   "98"
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   9480
         Tag             =   "98"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Ini&zio Carico"
         Height          =   255
         Index           =   48
         Left            =   5760
         TabIndex        =   188
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Inizio Carico"
         Height          =   255
         Index           =   47
         Left            =   120
         TabIndex        =   187
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   3960
         Tag             =   "97"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   5400
         Tag             =   "4"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ente E&mitt."
         Height          =   195
         Index           =   38
         Left            =   1200
         TabIndex        =   173
         Top             =   2280
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
         Left            =   2640
         TabIndex        =   172
         Tag             =   "4"
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5400
         Tag             =   "13"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Unità di Trasporto"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   93
         Top             =   2640
         Width           =   1695
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
         Left            =   2640
         TabIndex        =   156
         Tag             =   "13"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   33
         Left            =   3480
         Tag             =   "0"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero Viaggio"
         Height          =   255
         Index           =   35
         Left            =   480
         TabIndex        =   119
         Top             =   3000
         Width           =   1455
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
         TabIndex        =   94
         Tag             =   "3"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Società"
         Height          =   255
         Index           =   33
         Left            =   1320
         TabIndex        =   118
         Top             =   1920
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   31
         Left            =   5400
         Tag             =   "3"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   30
         Left            =   4200
         Tag             =   "9"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Tipo Viaggio"
         Height          =   255
         Index           =   32
         Left            =   720
         TabIndex        =   117
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   29
         Left            =   3960
         Tag             =   "101"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data Fine Viaggio"
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   116
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   9000
         Tag             =   "1"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Viaggio"
         Height          =   255
         Index           =   30
         Left            =   6120
         TabIndex        =   115
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Fine Viaggio"
         Height          =   255
         Index           =   29
         Left            =   5760
         TabIndex        =   114
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   24
         Left            =   9480
         Tag             =   "100"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Al Num. Viaggio Def."
         Height          =   255
         Index           =   27
         Left            =   5880
         TabIndex        =   113
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   23
         Left            =   9000
         Tag             =   "12"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal Num. &Viaggio Def."
         Height          =   255
         Index           =   24
         Left            =   360
         TabIndex        =   112
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   22
         Left            =   3480
         Tag             =   "11"
         Top             =   3360
         Width           =   240
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
         Left            =   2640
         TabIndex        =   90
         Tag             =   "92"
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Raggr. Ordine"
         Height          =   255
         Index           =   20
         Left            =   720
         TabIndex        =   111
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   16
         Left            =   5400
         Tag             =   "92"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Presidio"
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   110
         Top             =   840
         Width           =   615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3480
         Tag             =   "93"
         Top             =   840
         Width           =   240
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
         TabIndex        =   92
         Tag             =   "17"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Zona"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   109
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5400
         Tag             =   "17"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   5400
         Tag             =   "7"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Deposito"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   108
         Top             =   1200
         Width           =   735
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
         TabIndex        =   91
         Tag             =   "7"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Restituisci al massimo"
         Height          =   255
         Index           =   16
         Left            =   8520
         TabIndex        =   107
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "viaggi"
         Height          =   255
         Left            =   10800
         TabIndex        =   106
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8955
      Left            =   0
      TabIndex        =   98
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   15796
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
End
Attribute VB_Name = "frmVVTSnapshotExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DB 21/08/2003 Rilascio
Option Explicit

Const DISPONIBILE = "S"
Const MOTRICE = "M"
Const RIMORCHIO = "R"

Private XA As XArrayDB
Private XALog As XArrayDB
Private XAEv As XArrayDB

Public WithEvents DataClass As CViaggioSelezione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mintCurFrame As Integer
Private mvarFunction As FunctionCodes
Private mstrUdTAppoggio As String
Private AXX As CParametro
Private APR As CProdotto
Private AUT As CUnitaDiTrasporto
Private AAM As CAutomezzo
Private AAU As CAutista
Private OCT As COrdine
Private VVT As CViaggio
Private VVC As CViaggioConsegna
Private VVP As CViaggioProdotto
Private VPI As CViaggioEvento
Private AEE As CEnteEmittente
Private ADE As CDeposito
Private AME As CMotivoEvento
Private ADZ As CDestinazione
Private AZO As CZona
Private ARP As CRagDeposito
Private ATV As CTipoVendita
Private SEV As CSatellitareEvento
Private mobjOutFile As CTextFile
Private mobjOutFileInc As CTextFile
Private mblnAutomaticSend As Boolean
Private mlngTimerMinutes As Long
Private mlngTickTimer As Long

Private strUdTValoreValore As String
Private strMotriceValore As String
Private strRimorchioValore As String
Private strAutista1Valore As String
Private strAutista2Valore As String
Private strVettore1Valore As String
Private strVettore2Valore As String
Private strPedaggioValore As String
Private strDataConsegnaValore As String
Private strOraConsegnaValore As String
Private strOraPrevistaValore As String
Private strDataPrevistaValore As String
Private strProgressivoValore As String
Private strKmValore As String
Private strKmViaggioValore As String
Private strNoteValore As String
Private strNoteEsteseValore As String
Private blnFineCaricoValore As Boolean
Private strQtaCaricataValore As String
Private strQtaScaricataValore As String
Private blnNoGPSDataValore As Boolean
Private strSequenzaViaggioValore As String
Private strDataConsegna As String
Private strDataViaggio As String
Private strSocietaViaggio As String
Private strOraInizioCaricoValore As String
Private strOraFineCaricoValore As String
Private strDataFineCarico As String
Private strInizioScaricoValore As String
Private strInizioCaricoValore As String
Private strFineScaricoValore As String
Private strOraFineScaricoValore As String
Private strFlagNoConforme As String
Private strDesigillato As String

Private Enum AzioneRichiesta
    arModificaViaggio = 1
    arAnnullaModifica = 2
    arLetturaViaggi = 3
    arTrasmettiAnomalieAgip = 4
End Enum

Private mvarTripToUpd(0 To 3, 0 To 0) As Variant
Private mblnTrasmAll(1 To 4) As Boolean

Private Enum SincronizzazioneRichiesta
    srLetturaDati = 1
    srScritturaDati = 2
End Enum

'-- DP 24/07/08
'-- Aggiorna data e ora di trasmissione su GES del record
Private Function blnSetDeliveryTransmitted(strIDVVT As String, strIDVVC As String, strIDVVP As String) As Boolean
    On Error GoTo ErrHandler
    
    Const UPD_DELIVERY_DATA = "UPDATE VIAGVVP0 SET VVPDTGES = ?, VVPHRGES = ? WHERE VVPIDVVT = ? AND VVPIDVVC = ? AND VVPIDVVP = ?"
    
    blnSetDeliveryTransmitted = DataClass.DS.DA.ExecuteCommand(UPD_DELIVERY_DATA, Format(Date, "DD/MM/YY"), "01/01/1900 " & Format(Now, "HH:MM"), strIDVVT, strIDVVC, strIDVVP)
    Exit Function
    
ErrHandler:
    blnSetDeliveryTransmitted = False
End Function

Private Sub ExportEventi()
    On Error GoTo ErrHandler
    
    Const TEMP_FILE = "INC.tmp"
    Dim intTempFileNum As Integer
    Dim intRowError As Integer
    Dim intFileFromNum As Integer
    Dim intFileToNum As Integer
    Dim strLine As String
    Dim strTipoEvento As String
    Dim strDataInvio As String
    Dim strOraInvio As String
    Dim strProgrInvio As String
    Dim strFileName As String
    Dim strIDViaggio As String
    Dim strIDViaggioPrec As String
    Dim strIDDestinazione As String
    Dim strIDDestinazioneDirottamento As String
    Dim strIDDepositoRientro As String
    Dim strIDProdotto As String
    Dim strIDProdottoContaminato As String
    Dim strIDEnteEmittente As String
    Dim strCodNumeratore As String
    Dim varCounter As Variant
    Dim varBufferEventi As Variant
    Dim i As Long
    Dim lngRow As Long
    Dim lngCountOK As Long
    Dim lngNumRows As Long
    Dim lngNumRowsEventi As Long
    Dim ErrCode As Errors
    
    Select Case mvarFunction
        Case fcEXPSNPA, fcEXPSNPA2
            tboxLog.Text = tboxLog.Text & Format(Now(), "dd/mm/yy Hh:Nn.Ss") & " === ExportEventi === " & vbCrLf
    End Select
    
    If Not mobjOutFileInc.OpenFile(TEMP_FILE, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    
    strDataInvio = Format$(Date, "YYYYMMDD")
    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
    
    lngNumRows = XA.UpperBound(1)
    For lngRow = 0 To lngNumRows
    
        strIDViaggio = XA(lngRow, 5)
        If strIDViaggio <> strIDViaggioPrec Then
            strIDViaggioPrec = strIDViaggio
            If VPI.GetSelectionExport(strIDViaggio) Then
                varBufferEventi = VPI.DS.List
            End If
            
            If Not IsEmpty(varBufferEventi) Then
            
                lngNumRowsEventi = UBound(varBufferEventi, 2)
                If Len(mskFileName) > 0 Then
                    strFileName = mskFileName
                Else
                    ErrCode = FileNotFound
                    GoTo ErrHandler
                End If
                
            
                
                For i = 0 To lngNumRowsEventi
                    strTipoEvento = varBufferEventi(0, i)
                    
                    strIDDestinazione = vbNullString
                    With ADZ.DS
                        .Fields.ResetFields
                        .Fields("0").FindValue = varBufferEventi(34, i)
                        If (Not ADZ.Find) Or (.HowMany = 0) Then
                            intRowError = NoDestinazione
                            GoTo ErrHandler
                        End If
                        strIDDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                        strIDDestinazione = Mid(strIDDestinazione, 3)
                    End With
                    
                    strIDProdotto = vbNullString
                    With APR.DS
                        .Fields.ResetFields
                        .Fields("0").FindValue = varBufferEventi(7, i)
                        If (Not APR.Find) Or (.HowMany = 0) Then
                            intRowError = NoProdotto
                            GoTo ErrHandler
                        End If
                        strIDProdotto = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
                    End With
                    
                    With AEE.DS
                        .Fields.ResetFields
                        .Fields(.KeyFields(1)).FindValue = varBufferEventi(42, i)
                        If Not AEE.Find Then
                            ErrCode = NoEnteEmittente
                            GoTo ErrHandler
                        End If
                        If .HowMany = 0 Then
                            ErrCode = NoEnteEmittente
                            GoTo ErrHandler
                        End If
                        strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(XA(lngRow, 33))
                    End With
                    
                    With mobjOutFileInc
                        .ClearFields
                        .Item("1").value = strIDEnteEmittente
                        .Item("3").value = strDataInvio
                        .Item("4").value = strOraInvio
                        .Item("6").value = varBufferEventi(32, i) & varBufferEventi(33, i)
                        .Item("7").value = varBufferEventi(41, i)
                        .Item("8").value = strIDDestinazione
                        .Item("9").value = strTipoEvento
                        .Item("10").value = varBufferEventi(1, i)
                        .Item("11").value = Format$(varBufferEventi(2, i), "YYYYMMDD")
                        .Item("12").value = Format$(Hour(varBufferEventi(3, i)), "00") & ":" & Format$(Minute(varBufferEventi(3, i)), "00")
                        .Item("13").value = Format$(Hour(varBufferEventi(4, i)), "00") & ":" & Format$(Minute(varBufferEventi(4, i)), "00")
                        If Len(varBufferEventi(35, i)) > 0 Then
                            .Item("14").value = Left(varBufferEventi(35, i), 13)
                            .Item("15").value = Right(varBufferEventi(35, i), 2)
                            .Item("16").value = varBufferEventi(36, i)
                        End If
                        .Item("50").value = Format$(varBufferEventi(29, i), "YYYYMMDD")
                        .Item("51").value = varBufferEventi(30, i)
                        .Item("52").value = varBufferEventi(31, i)
                        Select Case strTipoEvento
                            Case "007", "008"
                                .Item("17").value = strIDProdotto
                                .Item("18").value = varBufferEventi(8, i)
                                .Item("19").value = Left(varBufferEventi(9, i), 1)
                                strIDProdottoContaminato = vbNullString
                                With APR.DS
                                    .Fields.ResetFields
                                    .Fields("0").FindValue = varBufferEventi(7, i)
                                    If (Not APR.Find) Or (.HowMany = 0) Then
                                        intRowError = NoProdotto
                                        GoTo ErrHandler
                                    End If
                                    strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
                                End With
                                .Item("20").value = strIDProdottoContaminato
                                .Item("21").value = varBufferEventi(11, i)
                                .Item("22").value = Left(varBufferEventi(12, i), 1)
                                If Len(varBufferEventi(13, i)) > 0 Then
                                    strIDProdottoContaminato = vbNullString
                                    With APR.DS
                                        .Fields.ResetFields
                                        .Fields("0").FindValue = varBufferEventi(13, i)
                                        If (Not APR.Find) Or (.HowMany = 0) Then
                                            intRowError = NoProdotto
                                            GoTo ErrHandler
                                        End If
                                        strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
                                    End With
                                    .Item("23").value = strIDProdottoContaminato
                                    .Item("24").value = varBufferEventi(14, i)
                                    .Item("25").value = Left(varBufferEventi(15, i), 1)
                                End If
                                If Len(varBufferEventi(16, i)) > 0 Then
                                    strIDProdottoContaminato = vbNullString
                                    With APR.DS
                                        .Fields.ResetFields
                                        .Fields("0").FindValue = varBufferEventi(16, i)
                                        If (Not APR.Find) Or (.HowMany = 0) Then
                                            intRowError = NoProdotto
                                            GoTo ErrHandler
                                        End If
                                        strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
                                    End With
                                    .Item("26").value = strIDProdottoContaminato
                                    .Item("27").value = varBufferEventi(17, i)
                                    .Item("28").value = Left(varBufferEventi(18, i), 1)
                                End If
                                If Len(varBufferEventi(19, i)) > 0 Then
                                    strIDProdottoContaminato = vbNullString
                                    With APR.DS
                                        .Fields.ResetFields
                                        .Fields("0").FindValue = varBufferEventi(19, i)
                                        If (Not APR.Find) Or (.HowMany = 0) Then
                                            intRowError = NoProdotto
                                            GoTo ErrHandler
                                        End If
                                        strIDProdottoContaminato = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
                                    End With
                                    .Item("29").value = strIDProdottoContaminato
                                    .Item("30").value = varBufferEventi(20, i)
                                    .Item("31").value = Left(varBufferEventi(21, i), 1)
                                End If
                            Case "005", "010" '--JG  30/09/09 Cambio 009 A 005
                                .Item("32").value = strIDProdotto
                                .Item("33").value = varBufferEventi(8, i)
                                .Item("34").value = Left(varBufferEventi(9, i), 1)
                                .Item("35").value = varBufferEventi(22, i)
                            Case "011"
                                strIDDepositoRientro = vbNullString
                                With ADE.DS
                                    .Fields.ResetFields
                                    .Fields("0").FindValue = varBufferEventi(23, i)
                                    If (Not ADE.Find) Or (.HowMany = 0) Then
                                        intRowError = NoDeposito
                                        GoTo ErrHandler
                                    End If
                                    strIDDepositoRientro = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                                End With
                                .Item("36").value = strIDProdotto
                                .Item("37").value = varBufferEventi(8, i)
                                .Item("38").value = Left(varBufferEventi(9, i), 1)
                                .Item("39").value = strIDDepositoRientro
                                .Item("40").value = Format$(varBufferEventi(24, i), "YYYYMMDD")
                            Case "012"
                                strIDDestinazioneDirottamento = vbNullString
                                With ADZ.DS
                                    .Fields.ResetFields
                                    .Fields("0").FindValue = varBufferEventi(25, i)
                                    If (Not ADZ.Find) Or (.HowMany = 0) Then
                                        intRowError = NoDestinazione
                                        GoTo ErrHandler
                                    End If
                                    strIDDestinazioneDirottamento = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                                    strIDDestinazioneDirottamento = Mid(strIDDestinazioneDirottamento, 3)
                                End With
                                .Item("41").value = strIDProdotto
                                .Item("42").value = varBufferEventi(8, i)
                                .Item("43").value = Left(varBufferEventi(9, i), 1)
                                .Item("44").value = strIDDestinazioneDirottamento
                                .Item("45").value = Format$(varBufferEventi(26, i), "YYYYMMDD")
                            Case "013"
                                .Item("46").value = Format$(Hour(varBufferEventi(27, i)), "00") & ":" & Format$(Minute(varBufferEventi(27, i)), "00")
                                .Item("47").value = Format$(Hour(varBufferEventi(28, i)), "00") & ":" & Format$(Minute(varBufferEventi(28, i)), "00")
                            Case "014"
                                .Item("48").value = Format$(Hour(varBufferEventi(27, i)), "00") & ":" & Format$(Minute(varBufferEventi(27, i)), "00")
                                .Item("49").value = Format$(Hour(varBufferEventi(28, i)), "00") & ":" & Format$(Minute(varBufferEventi(28, i)), "00")
                        End Select
                        
                        If Not .WriteLine(True) Then
                            intRowError = OutputError
                            GoTo ErrHandler
                        End If
                        With VPI.DS
                            .Fields.ResetFields
                            .Fields("0").FindValue = varBufferEventi(37, i)
                            .Fields("1").FindValue = varBufferEventi(38, i)
                            .Fields("2").FindValue = varBufferEventi(39, i)
                            .Fields("3").FindValue = varBufferEventi(40, i)
                        End With
                        VPI.Find
                        If Not VPI.SetFlagTrasmissione("S") Then
                            GoTo ErrHandler
                        End If
                        lngCountOK = lngCountOK + 1
                    End With
                Next i
            End If
        End If
    Next lngRow
    GoTo EndProc
ErrHandler:
EndProc:
    mobjOutFileInc.CloseFile
    
    If lngCountOK > 0 Then
        With DataClass.DS.DA
            Select Case Right(strIDEnteEmittente, 3)
                Case "004"
                    strCodNumeratore = "108"
                Case "624"
                    strCodNumeratore = "008"
                Case Else
            End Select
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", strCodNumeratore, 0)
            If varCounter = -1 Then
                ErrCode = InvalidCode
                GoTo ErrHandler
            End If
            If varCounter > 9999999999# Then
                varCounter = 1
            End If
        End With
    
        strProgrInvio = Format(varCounter, String(10, "0"))
        strFileName = strFileName & "\INC_" & strProgrInvio & ".DAT"
        intFileFromNum = FreeFile
        Open TEMP_FILE For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Mid(strLine, 31, 4) = Format(lngCountOK, "0000")
            Mid(strLine, 5, 10) = strProgrInvio
            Print #intFileToNum, strLine
        Loop
        Close #intFileFromNum
        Close #intFileToNum
        
        Select Case mvarFunction
            Case fcEXPSNPA, fcEXPSNPA2
                tboxLog.Text = tboxLog.Text & " File: " & strFileName & vbCrLf & " Records inviati: " & lngCountOK & vbCrLf
        End Select
    Else
        Select Case mvarFunction
            Case fcEXPSNPA, fcEXPSNPA2
                tboxLog.Text = tboxLog.Text & " Nessun record inviato." & vbCrLf
        End Select
    End If

End Sub


Private Property Get ViaggiModificati() As Long
    Const COUNT_VIAGGI_MODIFICATI = _
        "select count(*) from viagvvp0 where vvpflst5 = 'S'"
    Dim lngResult As Long
    Dim varBuf As Variant
    On Error GoTo ErrHandler
    
    If Not DataClass.DS.DA.QueryData(COUNT_VIAGGI_MODIFICATI, varBuf, lngResult) Then
        GoTo ErrHandler
    End If
    ViaggiModificati = varBuf(0, 0)
Exit Property
ErrHandler:
    ViaggiModificati = 0
End Property

Private Property Get UnitàDiTrasportoAppoggio() As String
    UnitàDiTrasportoAppoggio = mstrUdTAppoggio
End Property

Private Property Let UnitàDiTrasportoAppoggio(ByVal NewValue As String)
    mstrUdTAppoggio = NewValue
End Property

Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property

Private Sub DisplaySelection(List As Variant)
    Dim i As Long
    
    If IsEmpty(List) Then
        Set XA = New XArrayDB
        XA.ReDim 0, 0, 0, 0
        XA.DeleteRows 0
        TDBGrid1.Array = XA
        TDBGrid1.ReBind
        XALog.ReDim 0, 0, 0, 0
        XALog.DeleteRows 0
        tdbLog.ReBind
    Else
        RefreshXArray XA, TDBGrid1, List
        TDBGrid1.bookmark = 0
    
        lblLog(5) = 0
        lblLog(2) = 0
        lblLog(4) = 0
    End If
    
    Select Case mvarFunction
        Case fcEXPSNPV, fcEXPSNP2, fcEXPSNPA, fcEXPSNPA2
            If UBound(List, 2) >= 0 Then
            ReDim varLog(0 To 3, 0 To UBound(List, 2))
                For i = 0 To UBound(List, 2)
                    varLog(0, i) = List(0, i)
                    varLog(1, i) = List(1, i)
                    varLog(2, i) = List(2, i)
                Next i
                RefreshXArray XALog, tdbLog, varLog
            End If

        Case fcEXPSNPVA
            ClearXArray XALog, tdbLog
    
    End Select

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
    End If
End Function


Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    Dim i As Long
    
    Select Case mvarFunction
        Case fcEXPSNPA, fcEXPSNPA2
            MaskEdBox12 = Format(Date, "dd/mm/yy")
            MaskEdBox14 = Format(DateAdd("d", 1, Date), "dd/mm/yy")
            MaskEdBox45 = Format(Date, "dd/mm/yy")
            MaskEdBox44 = Format(DateAdd("d", -1, Date), "dd/mm/yy")
    End Select

    If Len(mskField10) = 0 Then
        GoTo ErrHandler
    End If
    
    With DataClass.DS
        If .Fields.Validate Then
            If .ScanClientCond Then
                ReDim varOrderFields(4) As Variant
                Me.MousePointer = vbHourglass
                varOrderFields(0) = "37"
                varOrderFields(1) = "0"
                varOrderFields(2) = "38"
                varOrderFields(3) = "94"
                varOrderFields(4) = "96"
                blnActiveOnly = (chkActiveOnly = vbChecked)
                
                If (chkOnlyChangedTrip.value = vbChecked) And (Me.chkTransAll = vbUnchecked) Then
                    .Field("Prodotto Stato Riga 5").FindValue = "S"
                Else
                    .Field("Prodotto Stato Riga 5").FindValue = Empty
                End If
                
                If Not DataClass.GetSelectionViaggiAlloScarico(blnActiveOnly, varOrderFields, CLng(mskMaxRecords.Text)) Then
                    GoTo ErrHandler
                End If
                varBuffer = DataClass.DS.List
                If Not IsEmpty(varBuffer) Then
                    For i = 0 To UBound(varBuffer, 2)
'-- DP 27/05/08
'-- Ora che dispongo di 3 diversi campi data ora il seguente if non è più necessario
'                        If Format(varBuffer(22, i), "hh:nn") = "00.00" Then
'                            varBuffer(22, i) = varBuffer(41, i)
'                        End If
                    Next i
                End If
                DisplaySelection varBuffer
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    SelectRecords = True
    GoTo EndProc
ErrHandler:
    SelectRecords = False
EndProc:
    Me.MousePointer = vbDefault
End Function

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim strSelectionFormula As String
    Dim dtMod As Date
    Dim Respok, Respuesta As String
    
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    
    Select Case Cmd
        Case fcUndo
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
        Case fcFind
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Solo se pueden modificar estos datos desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
            MsgBox "Password Correcta, datos guardados", vbOKOnly
            
                If SelectRecords Then
                    SetState fsSelection
                End If
            End If
        Case fcCustom1
            If TDBGrid1.bookmark <> vbNullString Then
                '-- DP 24/07/08
                '-- Controllo se il record è già stato esportato su file GES
                dtMod = VVP.dtGetDataInvioGES(XA.value(TDBGrid1.bookmark, 0), XA.value(TDBGrid1.bookmark, 1), XA.value(TDBGrid1.bookmark, 2))
                If dtMod <> CDate("01/01/1900") Then
                    If MsgBox("ATENCIÓN: Ya se ha enviado a CEPSA los datos de esta descarga a las " & Format(dtMod, "DD/MM/YYYY HH:MM") & " ¿se desea proseguir?", vbYesNo, "ResOil.2") = vbNo Then
                        Exit Sub
                    End If
                End If
                
                Call SincronizzaDatiViaggio(srLetturaDati)
                Call SetInterfaceForUpdTrip(arModificaViaggio)
            End If
        Case fcSave
            Select Case mvarFunction
                Case fcEXPSNPVA
                    VerificaAnomalie
            End Select
        Case fcPrint
            If (State = fsSelection) Then
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                varReportsList(0, 0) = "Stampa di Controllo Dati Gestione"
                varReportsList(1, 0) = "oil2vvt3.rpt"
                varReportsList(2, 0) = 1
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcOpen
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(TDBGrid1)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per la trasmissione... "
                .InitDir = App.Path
                .Filter = "Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                .ShowSave
                mskFileName = .fileName
            End With
        Case fcExport
        
            Select Case mvarFunction
                Case fcEXPSNPVA
                    ExportEventiAgip

                Case fcEXPSNPV, fcEXPSNP2, fcEXPSNPA, fcEXPSNPA2
                    'ExportRecords
                    ExportEventi

            End Select
        
        Case fcExit
            Unload Me
        Case fccustom2
            If TimerAutomaticSend.Enabled Then
                TimerAutomaticSend.Enabled = False
                cmdStartStopAutomatic.Caption = "Avvio"
            Else
                TimerAutomaticSend.Enabled = True
                cmdStartStopAutomatic.Caption = "Stop"
            End If
    
    End Select
End Sub

Private Sub GetRecord(strID As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields.ResetFields
        .Fields("2").FindValue = strID
        If Not DataClass.Find Then
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub SetInterfaceForUpdTrip(ByVal pRichiesta As AzioneRichiesta)
    On Error GoTo ErrHandler
    
    Select Case pRichiesta
        Case arModificaViaggio

            TDBGrid1.Enabled = False
            FraModificaViaggio.Visible = True
            FraTrasmettiViaggi.Visible = False
            CmdOpenTrip.Enabled = False
            cmdOpenFile.Enabled = False
            cmdExport.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command8.Enabled = False
     
            
'            MaskEdBox1.SetFocus
            
        Case arAnnullaModifica
            TDBGrid1.Enabled = False
            FraModificaViaggio.Visible = True
            FraTrasmettiViaggi.Visible = False
            CmdOpenTrip.Enabled = False
            cmdOpenFile.Enabled = False
            cmdExport.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command8.Enabled = False
        
        Case arLetturaViaggi
            TDBGrid1.Enabled = True
            FraModificaViaggio.Visible = False
            FraTrasmettiViaggi.Visible = True
            CmdOpenTrip.Enabled = True
            cmdOpenFile.Enabled = True
            If mblnAutomaticSend Then
                Command8.Enabled = False
                Command4.Enabled = False
                Command5.Enabled = False
                CmdOpenTrip.Enabled = False
                cmdExport.Enabled = False
                Command1.Enabled = False
            Else
                cmdExport.Enabled = True
                CmdOpenTrip.Enabled = True
                Command4.Enabled = True
                Command5.Enabled = True
                Command8.Enabled = True
            End If
    
    End Select
    lblLog(5).Caption = CStr(ViaggiModificati)
Exit Sub
ErrHandler:
    TDBGrid1.Enabled = True
    FraModificaViaggio.Visible = False
End Sub

Public Sub SetState(NewState As FormState)
    Dim strIDAEE As String
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            If Not Loading Then
                TabStrip1.Tabs(1).Selected = True
                mskField9.SetFocus
            End If
            MaskEdBox12 = Format(Date, "dd/mm/yy")
            MaskEdBox14 = Format(Date, "dd/mm/yy")
            
            Select Case mvarFunction
                Case fcEXPSNP2
                    mskField9 = ID_SOC_GUARNER
                    Call DataClass.DS.ValidateClientField(mskField9.Tag, mskField9.Text)
                    FieldUpdateDescription mskField9.Tag, Me, mskField9
                    FieldSetDescription mskField9.Tag, Me, mskField9.Text, mskField9
                    mskField9.Enabled = False
                    tboxLog.Visible = False
                    fraTimer.Visible = False

                Case fcEXPSNPA, fcEXPSNPA2
                    mskField9 = ID_SOC_CEPSA
                    Call DataClass.DS.ValidateClientField(mskField9.Tag, mskField9.Text)
                    FieldUpdateDescription mskField9.Tag, Me, mskField9
                    FieldSetDescription mskField9.Tag, Me, mskField9.Text, mskField9
                    mskField9.Enabled = False
                    
                    
                    Select Case mvarFunction
                        Case fcEXPSNPA  ' TPH
                            strIDAEE = "624"
                        Case fcEXPSNPA2 ' Eutrasur
                            strIDAEE = "004"
                    End Select
                    mskField10 = strIDAEE
                    Call DataClass.DS.ValidateClientField(mskField10.Tag, mskField10.Text)
                    FieldUpdateDescription mskField10.Tag, Me, mskField10
                    FieldSetDescription mskField10.Tag, Me, mskField10.Text, mskField10
                    mskField10.Enabled = False
                    
                    chkTransAll.value = vbUnchecked 'vbChecked
                    chkTransAll.Enabled = False
                    chkOnlyChangedTrip.value = vbChecked ' vbUnchecked
                    chkOnlyChangedTrip.Enabled = False
                    tboxLog.Visible = True
                    tboxLog.Enabled = True
                    fraTimer.Visible = True
                    MaskEdBox12 = Format(DateAdd("d", -2, Date), "dd/mm/yy")
                    MaskEdBox14 = Format(Date, "dd/mm/yy")

                Case fcEXPSNPVA
                    mskField9.Text = ID_SOC_AGIP
                    DataClass.DS.ValidateClientField mskField9.Tag, mskField9.Text
                    FieldUpdateDescription mskField9.Tag, Me, mskField9
                    FieldSetDescription mskField9.Tag, Me, mskField9.Text, mskField9
                    mskField9.Enabled = False

                Case Else
                    mskField9 = ID_SOC_CEPSA
                    Call DataClass.DS.ValidateClientField(mskField9.Tag, mskField9.Text)
                    FieldUpdateDescription mskField9.Tag, Me, mskField9
                    FieldSetDescription mskField9.Tag, Me, mskField9.Text, mskField9
                    tboxLog.Visible = False
                    fraTimer.Visible = False
            
            End Select

        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Viaggi trovati: " & DataClass.DS.HowMany
            TabStrip1.Tabs(2).Selected = True
            If Not Loading Then
                If TDBGrid1.Visible Then
                    TDBGrid1.SetFocus
                End If
            End If
            
    End Select
    Call SetInterfaceForUpdTrip(arLetturaViaggi)
End Sub



Private Sub btn_generaevento_Click() 'Generamos un registro en EventosMBflotasManual para que este siga su curso normal.
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim HREV, DYEV As String
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim CodEv As String
    Dim COD, CODINC, codentrega, SQL_GEN_EV As String
    Dim BKM As Variant
    Dim Respuesta As VbMsgBoxResult
    BKM = TDBGrid1.bookmark
    strIDViaggio = XA.value(BKM, 0)
    txt_idviag = strIDViaggio
    COD = XA.value(BKM, 45)
    txt_cod = COD
    
    Dim exito, FechaOk As Boolean
    exito = False
    FechaOk = True
    HREV = txt_hrev.Text
    DYEV = txt_dataev.Text
     CodEv = cmb_cges.ItemData(cmb_cges.ListIndex)
     Do While (Len(CodEv) < 3)
        CodEv = "0" + CodEv
     Loop
     txt_CodEv.Text = CodEv
'SM Comprobaciones de integridad de los datos

    If (DYEV) <> Format(CStr(Date), "DD/MM/YY") Then
        Respuesta = MsgBox("Chequeo de hora:." + Chr(13) + " La fecha del evento no coincide con la del día de hoy, ¿Está seguro de enviar el evento?", vbYesNo, "Confirmación")
        If Respuesta = vbNo Then FechaOk = False
        Else
        FechaOk = True
    End If
    If FechaOk = True Then
         If (Len(CodEv) = 3) Then
        
             If (Len(txt_hrev.Text) > 0 And Len(txt_dataev.Text) > 0 And txt_hrev.Text <> "__:__" And txt_dataev <> "__/__/__") Then
                 
               
               
                   If (CInt(CodEv) > 7 And CInt(CodEv) < 21) Then
                   'INCIDENCIAS
'                           Respuesta = MsgBox("Confirmación." + Chr(13) + " Este evento genera una incidencia en Cepsa, " + Chr(13) + "¿Está seguro de enviar la incidencia?", vbYesNo, "Confirmación")
'                         If Respuesta = vbYes Then
                             'Generar el evento en la tabla
                                    LockWindowUpdate frmIncidencias.hwnd
                                    frmIncidencias.ZOrder
                                    frmIncidencias.Icon = RESOIL2MAIN.Icon
                                    frmIncidencias.Show
                                    LockWindowUpdate 0

                             
'                             SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + COD + "','" + CodEv + "','" + DYEV + " " + HREV + "',GETDATE (),'0')"
'                             If DA.ExecuteQuery(SQL_GEN_EV) Then
'                                MsgBox ("Evento '" + CodEv + "' generado con exito   El: " + CStr(DYEV) + " A las :" + CStr(HREV))
'                                 txt_hrev.Text = "__:__"
'                                 txt_dataev.Text = "__/__/__"
'                                 exito = True
'                             End If
                       
'                         Else
'                              MsgBox ("Incidencia NO generada")
'                         End If
'
                   Else
                   
                         If (CodEv = "001" Or CodEv = "002" Or CodEv = "003") Then
                         
                             SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + CodEv + "','" + DYEV + " " + HREV + "',GETDATE (),'0')"
                             If DA.ExecuteQuery(SQL_GEN_EV) Then
                                 MsgBox ("Evento '" + CodEv + "' generado con exito ")
                                 txt_hrev.Text = "__:__"
                                 txt_dataev.Text = "__/__/__"
                                 exito = True
                             End If
                         ElseIf (CodEv = "007") Then
                             
                             Dim buffpda As Variant
                             Dim lngpda As Long
                             If Not DA.QueryCommand("SELECT AAUNRCEL From ANAGAAU0 WHERE  (AAUIDAAU = ?)", buffpda, lngpda, MaskEdBox6.Text) Then
                             End If
                             
                             If (lngpda > 0) Then
                                     SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdUsuarioMB,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + buffpda(0, 0) + "','" + CodEv + "','" + DYEV + " " + HREV + "',GETDATE (),'0')"
                                     If DA.ExecuteQuery(SQL_GEN_EV) Then
                                         MsgBox ("Evento '" + CodEv + "' generado con exito ")
                                         txt_hrev.Text = "__:__"
                                         txt_dataev.Text = "__/__/__"
                                         exito = True
                                     End If
                             Else
                                    MsgBox ("No es posible Generar el evento ya que el conductor no tiene asignada ninguna PDA")
                             End If
                         Else
                                'Generar el evento en la tabla
                             SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + COD + "','" + CodEv + "','" + DYEV + " " + HREV + "',GETDATE (),'0')"
                             If DA.ExecuteQuery(SQL_GEN_EV) Then
                                MsgBox ("Evento '" + CodEv + "' generado con exito ")
                                    txt_hrev.Text = "__:__"
                                    txt_dataev.Text = "__/__/__"
                                    exito = True
                             End If
                         End If
                   End If
             Else
                MsgBox ("Hay que introducir una fecha y una hora para crear el evento")
             End If
         Else
              MsgBox ("Numero de evento no indicado.")
         End If
    End If
    If (exito) Then
    Call RipristinaSelezione
    End If
  


End Sub

Private Sub btn_trtutt_Click()
Dim Respuesta As String
Dim DA As CDataAccess
Set DA = DataConnect.DA

        Respuesta = MsgBox("Confirmación." + Chr(13) + " Se van a volver a enviar los eventos sin respuesta de Cepsa de hace más de 45 minutos " + Chr(13) + "¿Está seguro?", vbYesNo, "Confirmación")
        If Respuesta = vbYes Then
            'Generar el evento en la tabla
             If DA.ExecuteQuery("UPDATE VIAGGES0 Set GESSEND0 = 0 Where (GESRECEP Is Null) And (GESDAYDY > CONVERT(VarChar, GetDate() - 3, 112)) And (GESDATHR < Left(CONVERT(VarChar, GetDate() - 0.03, 24), 5))AND (GESSEND0 <> 0)") Then
              MsgBox ("Eventos Reenviados")
            End If
        End If
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox12
            strTitle = "Dalla data fine viaggio"
        Case 1
            Set M = MaskEdBox45
            strTitle = "Alla data inizio carico"
        Case 2
            Set M = MaskEdBox44
            strTitle = "Dalla data inizio carico"
        Case 3
            Set M = txt_dataev
            strTitle = "Fecha de Evento"
        Case 4
            Set M = MaskEdBox14
            strTitle = "Alla data fine viaggio"
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub cmdEventi_Click()
    Dim LinkForm As Form
    
    Set LinkForm = DisplayForm(fcVIAGVPI)
    If Not LinkForm Is Nothing Then
        With LinkForm
            Set .ParentForm = Me
            If .GetRecord(XA.value(TDBGrid1.bookmark, 0), XA.value(TDBGrid1.bookmark, 1), XA.value(TDBGrid1.bookmark, 2)) Then
                .SetState fsEdit
            Else
                .SetState fsInsert
                .DataClass.DS.Fields("0").NewValue = XA.value(TDBGrid1.bookmark, 0)
                .DataClass.DS.Fields("0").Updated = True
                .DataClass.DS.Fields("1").NewValue = XA.value(TDBGrid1.bookmark, 1)
                .DataClass.DS.Fields("1").Updated = True
                .DataClass.DS.Fields("2").NewValue = XA.value(TDBGrid1.bookmark, 2)
                .DataClass.DS.Fields("2").Updated = True
                .DataClass.DS.Fields("3").NewValue = "1"
                .DataClass.DS.Fields("3").Updated = True
                .DataClass.DS.Fields("11").NewValue = XA.value(TDBGrid1.bookmark, 28)
                .DataClass.DS.Fields("11").Updated = True
                .DataClass.DS.Fields("40").NewValue = mskField9.Text
                .DataClass.DS.Fields("40").Updated = True
                .MaskEdBox15.Text = mskField9.Text
                
                Select Case mvarFunction
                    
                    Case fcEXPSNPVA
                        .MaskEdBox2.Text = "018"
                        Call .DataClass.DS.ValidateClientField(.MaskEdBox2.Tag, .MaskEdBox2.Text)
                        FieldUpdateDescription .MaskEdBox2.Tag, LinkForm, .MaskEdBox2
                        FieldSetDescription .MaskEdBox2.Tag, LinkForm, .MaskEdBox2.Text, .MaskEdBox2

                End Select
                .MaskEdBox24 = Format(XA.value(TDBGrid1.bookmark, 3), "dd/mm/yy")
            End If
            Me.Enabled = False
        End With
    End If
End Sub

Private Sub CmdSalva_Click()
    
    Dim strErr As String
    '--SM modificación para enviar los cambios de horarios a CEPSA
    Dim BKM As Variant
    BKM = TDBGrid1.bookmark
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim strIDProdotto As String
    Dim IDAAU As String
    Dim HREV, DYEV, HRLLE, DYLLE, HRPROG, DYPROG As String
    Dim Autorizacion As String
    Dim SQL_GFIND_NPALM As String
    Dim SQL_GEN_EV As String
    Dim NPalm As Variant
    Dim lng As Long
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim Respuesta As VbMsgBoxResult
    Dim FechaOk As Boolean
    
    strIDViaggio = XA.value(BKM, 0)
    strIDConsegna = XA.value(BKM, 45)
    IDAAU = MaskEdBox3.Text
    HREV = MaskEdBox30.Text
    DYEV = MaskEdBox48.Text
    HRLLE = MaskEdBox9.Text
    DYLLE = MaskEdBox49.Text
    HRPROG = MaskEdBox47.Text
    DYPROG = MaskEdBox27.Text

    Dim fechaact As Date
    Dim Estadoviaj As Variant
    Dim lnghm As Long
    FechaOk = True
'fechaact = CDate(DYEV + " " + HREV)


'Cambiar hora/fecha arribo en el caso de modificar la hora prevista
If (MaskEdBox48.Text <> "__/__/__" And MaskEdBox30.Text <> "__:__") Then
    MaskEdBox49.Text = MaskEdBox48.Text
    MaskEdBox9.Text = MaskEdBox30.Text
End If





If DYPROG <> DYEV Then
Respuesta = MsgBox("Control de fechas:" + Chr(13) + " La fecha programada es distinta a la actualizada. ¿Desea Enviar el evento igualmente?", vbYesNo, "Confirmación")
If Respuesta = vbNo Then FechaOk = False
End If
If FechaOk = True Then


        If (HRLLE = "__:__" And DYLLE = "__/__/__") Then
        If DA.QueryCommand("SELECT  VVCFLST5 From VIAGVVC0 WHERE     (VVCIDVVT = ?) AND (VVCIDOCT = ?)", Estadoviaj, lnghm, strIDViaggio, strIDConsegna) Then
        End If
          ''" + DYEV + " " + HREV + "'
        If (Estadoviaj(0, 0) <> "S") Or IsNull(Estadoviaj(0, 0)) Then
                   
                    If (HREV <> "__:__" And DYEV <> "__/__/__") Then
                    
                        If DA.QueryCommand("SELECT  VVCFLST5 From VIAGVVC0 WHERE     (VVCIDVVT = ?) AND (VVCIDOCT = ?)", Estadoviaj, lnghm, strIDViaggio, strIDConsegna) Then
                        End If
                          ''" + DYEV + " " + HREV + "'
                        If (Estadoviaj(0, 0) <> "S") Or IsNull(Estadoviaj(0, 0)) Then
                                Respuesta = MsgBox("Registro Guardado." + Chr(13) + " Desea enviar la actualización de hora a Cepsa?", vbYesNo, "Confirmación")
                                If Respuesta = vbYes Then
                                            SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + strIDConsegna + "','999', Getdate() ,Getdate(),'0')"
                                            If Not (DA.ExecuteQuery(SQL_GEN_EV)) Then
                                                 MsgBox ("ERROR, vuelva a intentarlo pasado unos minutos")
                                                 'MsgBox ("Evento '" + CodEv + "' generado con exito para el conductor : " + Conductor(0, 0) + "  El: " + CStr(DYEV) + " A las :" + CStr(HREV))
                                            End If
                                MsgBox ("Actualización de hora enviada correctamente a Cepsa para el servicio seleccionado")
                                
                                ' Salva i dati eventualmente modificati sul viaggio
                                If SincronizzaDatiViaggio(srScritturaDati) Then
                                    ' Chiude il frame di modifica viaggi e rinfresca il via<ggio sulla griglia
                                    Call RipristinaSelezione
                                End If
                        Else
                                ' Salva i dati eventualmente modificati sul viaggio
                                If SincronizzaDatiViaggio(srScritturaDati) Then
                                ' Chiude il frame di modifica viaggi e rinfresca il via<ggio sulla griglia
                                Call RipristinaSelezione
                                 
                                End If
                        End If
                        
                        End If
                     Else
                    ' Salva i dati eventualmente modificati sul viaggio
                    If SincronizzaDatiViaggio(srScritturaDati) Then
                        ' Chiude il frame di modifica viaggi e rinfresca il via<ggio sulla griglia
                        Call RipristinaSelezione
                    End If
                       
                    End If
                 
                   
        Else
        
                Autorizacion = InputBox("Este viaje ya ha sido entregado y reportado a Cepsa." + Chr(13) + Chr(13) + " Desea borrar esta información y borrar en Cepsa la entrega actualizando la hora programada?. Introduzca numero de viaje:", "Confirmacion")
                If Autorizacion = strIDViaggio Then
                        SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + strIDConsegna + "','999', Getdate() ,Getdate(),'0')"
                        If Not (DA.ExecuteQuery(SQL_GEN_EV)) Then
                             MsgBox ("ERROR, vuelva a intentarlo pasado unos minutos")
                             'MsgBox ("Evento '" + CodEv + "' generado con exito para el conductor : " + Conductor(0, 0) + "  El: " + CStr(DYEV) + " A las :" + CStr(HREV))
                        End If
                        MsgBox ("Actualización de hora enviada correctamente a Cepsa para el servicio seleccionado")
                      
                        ' Salva i dati eventualmente modificati sul viaggio
                        If SincronizzaDatiViaggio(srScritturaDati) Then
                        ' Chiude il frame di modifica viaggi e rinfresca il via<ggio sulla griglia
                        Call RipristinaSelezione
                
                        End If
                End If
        End If
        Else
                If (HREV <> "__:__" And DYEV <> "__/__/__") Then
                    Autorizacion = InputBox("Este viaje ya ha sido entregado y reportado a Cepsa." + Chr(13) + Chr(13) + " Desea borrar esta información y borrar en Cepsa la entrega actualizando la hora programada?. Introduzca numero de viaje:", "Confirmacion")
                    If Autorizacion = strIDViaggio Then
                            SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + strIDConsegna + "','999', Getdate() ,Getdate(),'0')"
                            If Not (DA.ExecuteQuery(SQL_GEN_EV)) Then
                                 MsgBox ("ERROR, vuelva a intentarlo pasado unos minutos")
                                 'MsgBox ("Evento '" + CodEv + "' generado con exito para el conductor : " + Conductor(0, 0) + "  El: " + CStr(DYEV) + " A las :" + CStr(HREV))
                            End If
                            MsgBox ("Actualización de hora enviada correctamente a Cepsa para el servicio seleccionado")
                            ' Salva i dati eventualmente modificati sul viaggio
                            If SincronizzaDatiViaggio(srScritturaDati) Then
                            ' Chiude il frame di modifica viaggi e rinfresca il via<ggio sulla griglia
                            Call RipristinaSelezione
                    
                            End If
                    End If
                Else
                Call RipristinaSelezione
                End If
          
        End If
End If
End Sub

Private Sub CmdAnnulla_Click()
    ' Rileggi i valori del viaggio
    Call PulisciDatiModificatiViaggio
    Call SincronizzaDatiViaggio(srLetturaDati)
End Sub

Private Sub CmdRitorna_Click()
    Call RipristinaSelezione
End Sub

Private Sub RipristinaSelezione()
    Dim BKM As Variant
    
    BKM = TDBGrid1.bookmark                         ' Mi salvo il Bookmark della griglia
    Call SetInterfaceForUpdTrip(arLetturaViaggi)    ' Chiudo il frame che permette la modifca dei dati del viaggio
    PulisciDatiModificatiViaggio                    ' Pulisco i controlli con gli eventuali dati modificati
    ' Rinfrescare è rischioso: se il viaggio modificato non rispondesse più alla selezione?
    DoCommand fcFind                                ' rinfresco la griglia dei viaggi selezionati
   TDBGrid1.SetFocus                               ' Assegno il fuoco alla griglia
    TDBGrid1.bookmark = BKM                         ' Mi posiziono col Bookmark della griglia sul viaggio che stavo modificando
End Sub

Private Function ValidateTestataViaggio(ByRef rErr As String) As Boolean
    On Error GoTo ErrHandler
    
    ' Verifico la validità dell'ora di fine carico (VVTHRINV - inizio viaggio)
    If MaskEdBox42.Text = vbNullString Then
        rErr = "Occorre indicare l'ora di carico."
        GoTo ErrHandler
    Else
        If MaskEdBox42.Text <> TIME_MASK Then
            If Not IsDate(Format(MaskEdBox42.Text, "HH:MM")) Then
                rErr = "Occore indicare un ora di fine carico corretta."
                GoTo ErrHandler
            End If
        End If
    End If
    
    ' Verifico la validità dell'unità di trasporto
    If MaskEdBox4.Text = vbNullString Then
        rErr = "Occorre indicare l'unità di trasporto."
        GoTo ErrHandler
    Else
        With AUT.DS
            .Fields.ResetFields
            .Fields("0").FindValue = MaskEdBox4.Text    ' Codice
            .Fields("9").FindValue = DISPONIBILE        ' Disponibilità
            .Fields("16").FindValue = 0                 ' Annullato
            If Not AUT.Find Or .HowMany = 0 Then
                rErr = "Unità di trasporto " + MaskEdBox4.Text + " non disponibile."
                GoTo ErrHandler
            End If
        End With
    End If
    
    ' Verifico la validità della motrice
    If MaskEdBox3.Text = vbNullString Then
        rErr = "Occorre indicare la motrice."
        GoTo ErrHandler
    Else
        With AAM.DS
            .Fields.ResetFields
            .Fields("0").FindValue = MaskEdBox3.Text    ' Codice
            .Fields("1").FindValue = MOTRICE            ' Motrice/Rimorchio
            .Fields("4").FindValue = DISPONIBILE        ' Disponibilità
            .Fields("26").FindValue = 0                 ' Annullato
            If Not AAM.Find Or .HowMany = 0 Then
                rErr = "Motrice " + MaskEdBox3.Text + " non disponibile."
                GoTo ErrHandler
            End If
        End With
    End If
    
    ' Verifico la validità del rimorchio
    If MaskEdBox5.Text <> vbNullString Then
        With AAM.DS
            .Fields.ResetFields
            .Fields("0").FindValue = MaskEdBox5.Text    ' Codice
            .Fields("1").FindValue = RIMORCHIO          ' Motrice/Rimorchio
            .Fields("4").FindValue = DISPONIBILE        ' Disponibilità
            .Fields("26").FindValue = 0                 ' Annullato
            If Not AAM.Find Or .HowMany = 0 Then
                rErr = "Rimorchio " + MaskEdBox5.Text + " non disponibile."
                GoTo ErrHandler
            End If
        End With
    End If
    
    ' Verifico la validità del primo autista
    If MaskEdBox6.Text = vbNullString Then
        rErr = "Occorre indicare il primo autista."
        GoTo ErrHandler
    Else
        With AAU.DS
            .Fields.ResetFields
            .Fields("0").FindValue = MaskEdBox6.Text    ' Codice
            .Fields("5").FindValue = DISPONIBILE        ' Disponibilità
            .Fields("8").FindValue = 0                  ' Annullato
            If Not AAU.Find Or .HowMany = 0 Then
                rErr = "Primo autista " + MaskEdBox6.Text + " non disponibile."
                GoTo ErrHandler
            End If
        End With
    End If
    
    ' Verifico la validità dei km viaggio
    If MaskEdBox18.Text <> vbNullString Then
        If Not IsNumeric(MaskEdBox18.Text) Then
            rErr = "Controllare i km viaggio"
            GoTo ErrHandler
        End If
    End If
    
    If Not VVT.CheckSequenza(XA.value(TDBGrid1.bookmark, 3), MaskEdBox11.Text, MaskEdBox4.Text, XA.value(TDBGrid1.bookmark, 4), XA.value(TDBGrid1.bookmark, 0)) Then
        rErr = "Il numero di sequenza è già in uso"
        GoTo ErrHandler
    End If
    
    ValidateTestataViaggio = True
    Exit Function
    
ErrHandler:
    ValidateTestataViaggio = False
End Function

Private Function UpdValidateConsegnaViaggio(ByRef rErr As String) As Boolean
    On Error GoTo ErrHandler
    
    
    ' Verifico la presenza del progressivo
    If MaskEdBox10.Text = vbNullString Then
        rErr = "Occorre indicare il progressivo consegna."
        GoTo ErrHandler
    Else
        If Not IsNumeric(MaskEdBox10.Text) Then
            rErr = "Occorre indicare un numero progressivo corretto."
            GoTo ErrHandler
        End If
    End If
    
    ' Verifico l'ora di consegna
    '-- DP 13/05/08
    '-- Data e ora di arrivo non sono obbligatorie se la consegna non è iniziata
    If Not Check1.value Then
        If MaskEdBox49.Text <> DATE_MASK Then
            If Not IsDate(Format(MaskEdBox49, "dd/mm/yy")) Then
                rErr = "Occorre indicare una data di consegna corretta."
                GoTo ErrHandler
            End If
            
            If MaskEdBox9.Text <> TIME_MASK Then
                If Not IsDate(Format(MaskEdBox9, "HH:MM")) Then
                    rErr = "Occorre indicare un'ora di consegna corretta."
                    GoTo ErrHandler
                End If
            Else
                rErr = "Occorre indicare un'ora di consegna corretta."
                GoTo ErrHandler
            End If
        End If
        
        If MaskEdBox9.Text <> TIME_MASK Then
            If Not IsDate(Format(MaskEdBox9, "HH:MM")) Then
                rErr = "Occorre indicare un' ora di consegna corretta."
                GoTo ErrHandler
            End If
            
            If MaskEdBox49.Text <> DATE_MASK Then
                If Not IsDate(Format(MaskEdBox49, "dd/mm/yy")) Then
                    rErr = "Occorre indicare una data di consegna corretta."
                    GoTo ErrHandler
                End If
            Else
                rErr = "Occorre indicare una data di consegna corretta."
                GoTo ErrHandler
            End If
        End If
    Else
        If Not IsDate(Format(MaskEdBox49, "dd/mm/yy")) Then
            rErr = "Occorre indicare una data di consegna corretta."
            GoTo ErrHandler
        End If
        
        If Not IsDate(Format(MaskEdBox9, "HH:MM")) Then
            rErr = "Occorre indicare un'ora di consegna corretta."
            GoTo ErrHandler
        End If
    End If
    
    ' Verifico la data prevista
    If Not IsDate(Format(MaskEdBox48, "dd/mm/yy")) Then
        rErr = "Occorre indicare una data prevista corretta."
        GoTo ErrHandler
    End If
    
    ' Verifico l'ora prevista
    If Not IsDate(Format(MaskEdBox30, "HH:MM")) Then
        rErr = "Occorre indicare un' ora prevista corretta."
        GoTo ErrHandler
    End If
    
    ' Verifico la data di consegna con la data di inizio viaggio
    If IsDate(MaskEdBox27) Then
        If DateDiff("d", MaskEdBox27, strDataViaggio) > 0 Then
            rErr = "La data di consegna non può essere precedente alla data del viaggio"
            GoTo ErrHandler
        Else
            If DateDiff("d", MaskEdBox27, strDataViaggio) = 0 Then
                If strOraInizioCaricoValore <> TIME_MASK And MaskEdBox49 <> DATE_MASK And MaskEdBox9 <> TIME_MASK Then
                    If DateDiff("n", MaskEdBox49 & " " & MaskEdBox9, MaskEdBox46.Text & " " & strOraInizioCaricoValore) > 0 Then
                        rErr = "L'ora di consegna non può essere precedente all'ora di inizio del carico"
                        GoTo ErrHandler
                    End If
                End If
            End If
        End If
    End If
    
    ' Verifico i Km
    If MaskEdBox17.Text <> vbNullString Then
        If Not IsNumeric(MaskEdBox17.Text) Then
            rErr = "Occore indicare un numero di km corretto."
            GoTo ErrHandler
        End If
    End If
    
    UpdValidateConsegnaViaggio = True
Exit Function
ErrHandler:
    UpdValidateConsegnaViaggio = False
End Function

Private Function ValidateProdottoViaggio(ByRef rErr As String) As Boolean
    On Error GoTo ErrHandler

    ' Verifico la quantità caricata
    If MaskEdBox28.Text = vbNullString Then
        rErr = "Occorre indicare la quantità caricata in litri."
        GoTo ErrHandler
    Else
        If Not IsNumeric(MaskEdBox28.Text) Then
            rErr = "Occorre indicare una quantità caricata significativa."
            GoTo ErrHandler
        End If
    End If
    
    ' Verifico la quantità scaricata
    If Check1.value = vbChecked Then
        If MaskEdBox29.Text = vbNullString Then
            rErr = "Occorre indicare la quantità scaricata in litri."
            GoTo ErrHandler
        Else
            If Not IsNumeric(MaskEdBox29.Text) Then
                rErr = "Occorre indicare una quantità scaricata significativa."
                GoTo ErrHandler
            End If
        End If
    End If
    ValidateProdottoViaggio = True
Exit Function
ErrHandler:
    ValidateProdottoViaggio = False
End Function

Private Function TestataViaggioUpdated(ByRef rResult As Boolean) As Boolean
    On Error GoTo ErrHandler
    
    rResult = False
    If (UCase(MaskEdBox1.Text) <> strOraInizioCaricoValore) Or _
       (UCase(MaskEdBox42.Text) <> strOraFineCaricoValore) Or _
       (UCase(MaskEdBox46.Text) <> strDataFineCarico) Or _
       (UCase(MaskEdBox4.Text) <> strUdTValoreValore) Or _
       (UCase(MaskEdBox3.Text) <> strMotriceValore) Or _
       (UCase(MaskEdBox5.Text) <> strRimorchioValore) Or _
       (UCase(MaskEdBox6.Text) <> strAutista1Valore) Or _
       (UCase(MaskEdBox32.Text) <> strVettore1Valore) Or _
       (UCase(MaskEdBox33.Text) <> strVettore2Valore) Or _
       (UCase(MaskEdBox18.Text) <> strKmViaggioValore) Or _
       (UCase(MaskEdBox19.Text) <> strPedaggioValore) Or _
       blnNoGPSDataValore <> IIf(chkNoGPSData.value = vbChecked, True, False) Or _
       MaskEdBox11.Text <> strSequenzaViaggioValore Or _
       Check2.value <> blnFineCaricoValore Then
        rResult = True
    End If
    TestataViaggioUpdated = True
Exit Function
ErrHandler:
    TestataViaggioUpdated = False
End Function

Private Function UpdTestataViaggio(ByVal pIdViaggio As String, ByRef rErr As String) As Boolean
    Const UPD_TESTATA_VIAGGIO = _
        "UPDATE VIAGVVT0 SET "
    
    ' -- DP 05/06/08
    ' -- Aggiunti come ultimo campo "vvtdtinv" e "vvtflagg" per consentire l'aggiornamento del campo "Ora Fine Carico" e "Flag ultimo aggiornamento"
    Const UPD_TESTATA_VIAGGIO_2 = _
        "  vvtidaut = ?, " & _
        "  vvtidam1 = ?, vvtidam2 = ?, " & _
        "  vvtidau1 = ?, vvtkmdef = ?, " & _
        "  vvtnripa = ?, vvtseqvi = ?, " & _
        "  vvtflst3 = ?, vvtidau2 = ?, " & _
        "  vvtidve1 = ?, vvtidve2 = ?, " & _
        "  vvtfludg = ?, vvtdtinv = ?, " & _
        "  vvthrinc = ?, vvthrinv = ?, " & _
        "  vvtseqin = ?, vvtflagg = 'SE' " & _
        "WHERE vvtidvvt = ? "
    
    Const UPD_TESTATA_VIAGGIO_3 = _
        "  vvtidaut = ?, " & _
        "  vvtidam1 = ?, vvtidam2 = ?, " & _
        "  vvtidau1 = ?, vvtkmdef = ?, " & _
        "  vvtnripa = ?, vvtseqvi = ?, " & _
        "  vvtflst3 = ?, vvtidau2 = ?, " & _
        "  vvtidve1 = ?, vvtidve2 = ?, " & _
        "  vvtfludg = ?, vvtdtinv = ?, " & _
        "  vvthrinc = ?, vvtseqin = ?, vvtflagg = 'SE' " & _
        "WHERE vvtidvvt = ? "
    
    Const UPD_FESTIVO = "UPDATE VIAGVVT0 SET VVTFLST2 = ? WHERE VVTIDVVT = ?"
    
    Dim strDataCarico As Variant
    Dim strOraCarico As Variant
    Dim strOraFineCarico As Variant
    Dim strUdt As String
    Dim strMotrice As String
    Dim strRimorchio As String
    Dim strAutista1 As String
    Dim strAutista2 As String
    Dim strVettore1 As String
    Dim strVettore2 As String
    Dim lngKmDef As Long
    Dim varImportoPedaggi As Variant
    Dim varSequenzaViaggio As Variant
    Dim strErr As String
    Dim strUsaDatiGPS As String
    
    Dim blnCambiaSequenzaInvio As Boolean
    Dim strSequenzaInvio As String
    
    On Error GoTo ErrHandler
    
    If Not ValidateTestataViaggio(strErr) Then
        GoTo ErrHandler
    End If
    
    strDataCarico = MaskEdBox46.Text
    strOraCarico = MaskEdBox1.Text                                                      ' Ora Inizio Carico
    strOraFineCarico = MaskEdBox42.Text                                                 ' Ora Fine Carico
    strUdt = UCase(MaskEdBox4.Text)                                                     ' Unità di Trasporto
    strMotrice = UCase(MaskEdBox3.Text)                                                 ' Motrice
    strRimorchio = UCase(MaskEdBox5.Text)                                               ' Rimorchio
    strAutista1 = UCase(MaskEdBox6.Text)                                                ' Autista 1
    strVettore1 = UCase(MaskEdBox32.Text)                                               ' Vettore
    strVettore2 = UCase(MaskEdBox33.Text)                                               ' Vettore Trasp.
    lngKmDef = IIf(Len(MaskEdBox18.Text) = 0, 0, MaskEdBox18.Text)                      ' Km Viaggio
    varSequenzaViaggio = MaskEdBox11.Text
    strSequenzaViaggioValore = MaskEdBox11.Text
    strSequenzaInvio = MaskEdBox54.Text                                                 ' Sequenza Invio
    
    If IsNumeric(MaskEdBox19.Text) Then
        varImportoPedaggi = MaskEdBox19.Text
    Else
        varImportoPedaggi = 0
    End If
    blnNoGPSDataValore = IIf(chkNoGPSData.value = vbChecked, True, False)
    
    If blnNoGPSDataValore Then
        strUsaDatiGPS = "N"
    Else
        strUsaDatiGPS = "S"
    End If
    
    
    '-- DP 27/08/08
    '-- Gestione sequenza di invio
    If Format(strDataCarico, "dd/mm/yy") <> Format(XA.value(TDBGrid1.bookmark, 3), "dd/mm/yy") And XA.value(TDBGrid1.bookmark, 75) <> "N" Then
        blnCambiaSequenzaInvio = True
    End If
    
    If Trim(UCase(strUdt)) <> Trim(UCase(XA.value(TDBGrid1.bookmark, 11))) And XA.value(TDBGrid1.bookmark, 75) <> "N" Then
        blnCambiaSequenzaInvio = True
    End If
    
    If CInt(varSequenzaViaggio) <> CInt(XA.value(TDBGrid1.bookmark, 8)) Then
        blnCambiaSequenzaInvio = True
    End If
        
    If blnCambiaSequenzaInvio Then
        If Not IsNull(XA.value(TDBGrid1.bookmark, 72)) Then
            If XA.value(TDBGrid1.bookmark, 72) <> vbNullString Then
                If Right(XA.value(TDBGrid1.bookmark, 72), 1) = Right(varSequenzaViaggio, 1) Then
                    strSequenzaInvio = Format(CInt(XA.value(TDBGrid1.bookmark, 72)) + 10, "000")
                Else
                    strSequenzaInvio = Format(varSequenzaViaggio, "000")
                End If
            Else
                strSequenzaInvio = Format(CInt(varSequenzaViaggio) + 10, "000")
            End If
        Else
            strSequenzaInvio = Format(CInt(varSequenzaViaggio) + 10, "000")
        End If
        
        If Not VVT.blnUpdateSEQOK("N", CVar(pIdViaggio)) Then
            GoTo ErrHandler
        End If
    End If
    '-- Fine gestione sequenza di invio
    
    
'    If strOraFineCarico <> TIME_MASK Then
    If Check2.value = vbChecked Then
        If Not DataClass.DS.DA.ExecuteCommand(UPD_TESTATA_VIAGGIO + UPD_TESTATA_VIAGGIO_2, _
                                              strUdt, strMotrice, strRimorchio, strAutista1, lngKmDef, varImportoPedaggi, varSequenzaViaggio, "S", strAutista2, strVettore1, strVettore2, strUsaDatiGPS, strDataCarico, "01/01/1900 " & strOraCarico, "01/01/1900 " & strOraFineCarico, strSequenzaInvio, pIdViaggio) Then
            strErr = "Errore durante l'aggiornamento della testata del viaggio " + pIdViaggio
            GoTo ErrHandler
        End If
    Else
        If Not DataClass.DS.DA.ExecuteCommand(UPD_TESTATA_VIAGGIO + UPD_TESTATA_VIAGGIO_3, _
                                              strUdt, strMotrice, strRimorchio, strAutista1, lngKmDef, varImportoPedaggi, varSequenzaViaggio, "N", strAutista2, strVettore1, strVettore2, strUsaDatiGPS, strDataCarico, "01/01/1900 " & strOraCarico, strSequenzaInvio, pIdViaggio) Then
            strErr = "Errore durante l'aggiornamento della testata del viaggio " + pIdViaggio
            GoTo ErrHandler
        End If
    End If
    
    If chkFestivo.value = vbChecked Then
        If Not DataClass.DS.DA.ExecuteCommand(UPD_FESTIVO, "F", pIdViaggio) Then
            GoTo ErrHandler
        End If
    Else
        If Not DataClass.DS.DA.ExecuteCommand(UPD_FESTIVO, "''", pIdViaggio) Then
            GoTo ErrHandler
        End If
    End If
    
    UpdTestataViaggio = True
    Exit Function

ErrHandler:
    UpdTestataViaggio = False
    If strErr = vbNullString Then
        strErr = "Errore sconosciuto durante l'aggiornamento della testata del viaggio " + pIdViaggio
    End If
    rErr = strErr
End Function

Private Function ConsegnaViaggioUpdated(ByRef rResult As Boolean) As Boolean
    On Error GoTo ErrHandler

    rResult = False
    If (UCase(MaskEdBox49.Text) <> strDataConsegnaValore) Or _
       (UCase(MaskEdBox9.Text) <> strOraConsegnaValore) Or _
       (UCase(MaskEdBox48.Text) <> strDataPrevistaValore) Or _
       (UCase(MaskEdBox30.Text) <> strOraPrevistaValore) Or _
       (UCase(MaskEdBox10.Text) <> strProgressivoValore) Or _
       (UCase(MaskEdBox17.Text) <> strKmValore) Or _
       (UCase(MaskEdBox25.Text) <> strNoteValore) Or _
       (UCase(TxtNoteEstese.Text) <> strNoteEsteseValore) Or _
       (MaskEdBox27.Text <> strDataConsegna) Or _
       (MaskEdBox43.Text <> strOraFineScaricoValore) Then
        rResult = True
    End If
        
    If Check1.value = vbChecked Then
        If strInizioScaricoValore <> "S" Then
            rResult = True
        End If
    Else
        If strInizioScaricoValore <> "N" Then
            rResult = True
        End If
    End If
    If Check4.value = vbChecked Then
        If strFineScaricoValore <> "S" Then
            rResult = True
        End If
    Else
        If strFineScaricoValore <> "N" Then
            rResult = True
        End If
    End If
    ConsegnaViaggioUpdated = True
Exit Function
ErrHandler:
    ConsegnaViaggioUpdated = False
End Function

Private Function UpdConsegnaViaggio(ByVal pIdViaggio As String, ByVal pIdConsegna As String, ByRef rErr As String) As Boolean
    Const UPD_DATA_CONSEGNA_VIAGGIO = "UPDATE VIAGVVC0 SET vvcflagg = 'SE', vvcdtarr = "
    Const UPD_ORA_CONSEGNA_VIAGGIO = ", vvchrarr = "
    
    '-- DP 13/05/08
    '-- L'ora di arrivo non è obbligatoria
    Const UPD_CONSEGNA_VIAGGIO_B = "UPDATE VIAGVVC0 SET vvcflagg = 'SE', "
    
    Const UPD_CONSEGNA_VIAGGIO_1 = " vvcnrseq = ?, vvcflst5 = ?, vvcflst6 = ?, vvcnote1 = ?, vvcnotee = ? "
    Const UPD_CONSEGNA_VIAGGIO_KM_PREVISTI = ", vvckmpap = ? "
    Const UPD_CONSEGNA_VIAGGIO_KM_DEFINITIVI = ", vvckmpad = ? "
    Const UPD_CONSEGNA_DATA = ", vvcdtcon = ? "
    Const UPD_CONSEGNA_ORA_PREVISTA_AGGIORNATA = ", vvcflst4 = ? "
    Const UPD_CONSEGNA_DATA_PREVISTA = ", vvcdtpre = "
    Const UPD_CONSEGNA_ORA_PREVISTA = ", vvchrpre = "
    Const UPD_CONSEGNA_VIAGGIO_WHERE = " WHERE vvcidvvt = ? AND vvcidvvc = ? "
    Const UPD_PRODOTTO_ORA_FINE_SCARICO = " UPDATE VIAGVVP0 SET VVPFLAGG = 'SE', VVPHRFIS = "
    Const UPD_PRODOTTO_ORA_FINE_SCARICO_WHERE = " WHERE VVPIDVVT = ? AND VVPIDVVC = ? AND VVPANNUL = 0 "
    
    Dim varDataArrivo As Variant
    Dim varOraConsegna As Variant
    Dim varDataPrevista As Variant
    Dim varOraPrevista As Variant
    Dim varOraFineScarico As Variant
    Dim varDataConsegna As Variant
    Dim strProgressivo As String
    Dim varKM As Variant
    Dim strNote As String
    Dim strNoteEstese As String
    Dim strConsegnaIniziata As String
    Dim strConsegnaTerminata As String
    Dim strFlagOraPrevistaAggiornata As String
    Dim strErr As String
    On Error GoTo ErrHandler
 
    If Not UpdValidateConsegnaViaggio(strErr) Then
        GoTo ErrHandler
    End If
    
    '-- DP 13/05/08
    '-- Data e ora di arrivo non sono obbligatorie
    If MaskEdBox9.Text <> TIME_MASK And MaskEdBox49.Text <> DATE_MASK Then
        varDataArrivo = MaskEdBox49.Text
        varOraConsegna = MaskEdBox9.Text
    End If
    varDataPrevista = MaskEdBox48.Text
    varOraPrevista = MaskEdBox30.Text
    varOraFineScarico = MaskEdBox43.Text
    varDataConsegna = MaskEdBox27.Text
    strProgressivo = UCase(MaskEdBox10.Text)
    If IsNumeric(MaskEdBox17) Then
        varKM = CLng(MaskEdBox17.Text)
    Else
        varKM = 0
    End If
    strNote = UCase(MaskEdBox25.Text)
    strNoteEstese = UCase(Replace(TxtNoteEstese.Text, vbCrLf, " "))
    If Check1.value = vbChecked Then
        strConsegnaIniziata = "S"
        ' Se non abbiamo impostato i km "reali" utilizziamo i km contrattuali
        If varKM = 0 Then
            If IsNumeric(CVar(MaskEdBox24.Text)) Then
                varKM = MaskEdBox24.Text
            End If
        End If
    Else
        strConsegnaIniziata = "N"
    End If
    
    If Check4.value = vbChecked Then
        strConsegnaTerminata = "S"
    Else
        strConsegnaTerminata = "N"
    End If
    
    If MaskEdBox30.Text <> strOraPrevistaValore Or MaskEdBox48.Text <> strDataPrevistaValore Then
        strFlagOraPrevistaAggiornata = "S"
    Else
        strFlagOraPrevistaAggiornata = "N"
    End If
    
    DataClass.DS.DA.BeginTransaction
    
    '-- DP 13/05/08
    '-- Aggiunto l'"IIf" per gestire l'ora di arrivo come non obbligatoria
    If Not DataClass.DS.DA.ExecuteCommand(IIf(MaskEdBox9.Text <> TIME_MASK And MaskEdBox49.Text <> DATE_MASK, UPD_DATA_CONSEGNA_VIAGGIO + DataClass.DS.DA.MakeSQL(varDataArrivo, 2) + UPD_ORA_CONSEGNA_VIAGGIO + DataClass.DS.DA.MakeSQL(varOraConsegna, 4) + ", ", UPD_CONSEGNA_VIAGGIO_B) + _
                                        UPD_CONSEGNA_VIAGGIO_1 + UPD_CONSEGNA_VIAGGIO_KM_DEFINITIVI + UPD_CONSEGNA_DATA + UPD_CONSEGNA_ORA_PREVISTA_AGGIORNATA + _
                                        UPD_CONSEGNA_DATA_PREVISTA + DataClass.DS.DA.MakeSQL(varDataPrevista, dtDate) + _
                                        UPD_CONSEGNA_ORA_PREVISTA + DataClass.DS.DA.MakeSQL(varOraPrevista, 4) + UPD_CONSEGNA_VIAGGIO_WHERE, _
                                        strProgressivo, strConsegnaIniziata, strConsegnaTerminata, strNote, strNoteEstese, varKM, varDataConsegna, strFlagOraPrevistaAggiornata, pIdViaggio, pIdConsegna) Then
        rErr = "Errore durante l'aggiornamento della consegna " + pIdConsegna + " per il viaggio " + pIdViaggio
        GoTo ErrHandler
    End If
    
    If (varOraFineScarico <> TIME_MASK) And (strConsegnaTerminata = "S") Then
        If Not DataClass.DS.DA.ExecuteCommand(UPD_PRODOTTO_ORA_FINE_SCARICO + DataClass.DS.DA.MakeSQL(varOraFineScarico, 4) + UPD_PRODOTTO_ORA_FINE_SCARICO_WHERE, _
                                              pIdViaggio, pIdConsegna) Then
            rErr = "Errore durante l'aggiornamento degli orari di fine scarico della consegna " + pIdConsegna + " per il viaggio " + pIdViaggio
            GoTo ErrHandler
        End If
    End If
    
    DataClass.DS.DA.EndTrans ETCOMMIT
    
    VVT.DS.Fields("0").FindValue = pIdViaggio
    VVT.Find
    If VVT.CheckViaggioCompletato(pIdViaggio) Then
        If Not VVT.SetStato("C") Then
            GoTo ErrHandler
        End If
    Else
        If Not VVT.SetStato("D") Then
            GoTo ErrHandler
        End If
    End If
    
    UpdConsegnaViaggio = True
Exit Function
ErrHandler:
    DataClass.DS.DA.EndTrans ETROLLBACK
    UpdConsegnaViaggio = False
    If strErr = vbNullString Then
        strErr = "Errore sconosciuto durante l'aggiornamento della consegna " + pIdConsegna + " del viaggio " + pIdViaggio
    End If
    rErr = strErr
End Function

Private Function ProdottoViaggioUpdated(ByRef rResult As Boolean) As Boolean
    On Error GoTo ErrHandler
    
    rResult = False
    If (CDbl(MaskEdBox28.Text) <> CDbl(strQtaCaricataValore)) Or _
       (CDbl(MaskEdBox29.Text) <> CDbl(strQtaScaricataValore)) Then
        rResult = True
    End If
    If Me.chkNoConforme.value = vbChecked Then
        If strFlagNoConforme <> "S" Then
            rResult = True
        End If
    Else
        If strFlagNoConforme <> "N" Then
            rResult = True
        End If
    End If
    If Me.chkDesigillato.value = vbChecked Then
        If strDesigillato <> "S" Then
            rResult = True
        End If
    Else
        If strDesigillato <> "N" Then
            rResult = True
        End If
    End If
    ProdottoViaggioUpdated = True
Exit Function
ErrHandler:
    ProdottoViaggioUpdated = False
End Function

Private Function UpdProdottoViaggio(ByVal pIdViaggio As String, ByVal pIdConsegna As String, ByVal pIdProdotto As String, ByRef rErr As String) As Boolean
    Const UPD_VIAGGIO_PRODOTTO = "UPDATE VIAGVVP0 SET vvpflagg = 'SE', vvpqtrlt = ?, vvpflncc = ?, vvpfldes = ?  "
    Const UPD_VIAGGIO_PRODOTTO_QT_SCARICATA = ", vvpqcolt = ? "
    Const UPD_VIAGGIO_PRODOTTO_WHERE = " WHERE vvpidvvt = ? and vvpidvvc = ? and vvpidvvp = ? "
    
    Dim varQtaCaricata As Variant
    Dim varQtaScaricata As Variant
    Dim strErr As String
    On Error GoTo ErrHandler
    
    If Not ValidateProdottoViaggio(strErr) Then
        GoTo ErrHandler
    End If
    
    varQtaCaricata = CDbl(MaskEdBox28.Text)
    varQtaScaricata = CDbl(MaskEdBox29.Text)
    
    If chkDesigillato.value = vbChecked Then
        strDesigillato = "S"
    Else
        strDesigillato = "N"
    End If
    If Me.chkNoConforme.value = vbChecked Then
        strFlagNoConforme = "S"
    Else
        strFlagNoConforme = "N"
    End If
    
    DataClass.DS.DA.BeginTransaction
    ' Verifico la quantità scaricata
    If Check1.value = vbChecked Then
        If Not DataClass.DS.DA.ExecuteCommand(UPD_VIAGGIO_PRODOTTO + UPD_VIAGGIO_PRODOTTO_QT_SCARICATA + UPD_VIAGGIO_PRODOTTO_WHERE, _
                                              CDbl(varQtaCaricata), strFlagNoConforme, strDesigillato, CDbl(varQtaScaricata), pIdViaggio, pIdConsegna, pIdProdotto) Then
            strErr = "Errore durante l'aggiornamento del prodotto " + pIdProdotto + " nella consegna " + pIdConsegna + " del viaggio " + pIdViaggio
            GoTo ErrHandler
        End If
    Else
        If Not DataClass.DS.DA.ExecuteCommand(UPD_VIAGGIO_PRODOTTO + UPD_VIAGGIO_PRODOTTO_WHERE, _
                                              CDbl(varQtaCaricata), strFlagNoConforme, strDesigillato, pIdViaggio, pIdConsegna, pIdProdotto) Then
            strErr = "Errore durante l'aggiornamento del prodotto " + pIdProdotto + " nella consegna " + pIdConsegna + " del viaggio " + pIdViaggio
            GoTo ErrHandler
        End If
    End If
    DataClass.DS.DA.EndTrans ETCOMMIT
    UpdProdottoViaggio = True
Exit Function
ErrHandler:
    DataClass.DS.DA.EndTrans ETROLLBACK
    UpdProdottoViaggio = False
    If strErr = vbNullString Then
        strErr = "Errore sconosciuto durante l'aggiornamento del prodotto " + pIdProdotto + " nella consegna " + pIdConsegna + " del viaggio " + pIdViaggio
    End If
    rErr = strErr
End Function

Private Function UpdModificato(ByVal pIdViaggio As String, ByVal pIdConsegna As String, ByVal pIdProdotto As String, _
                               ByVal pUpdatedTestata As Boolean, ByVal pUpdatedConsegna As Boolean, ByVal pUpdatedProdotto As Boolean, _
                               ByVal pNewValue As String) As Boolean
    Const UPD_MODIFICATO = _
        "UPDATE VIAGVVP0 SET VVPFLST5 = ? WHERE "
    Const UPD_MODIFICATO_PRODOTTO = _
        " AND VVPIDVVP = ? "
    Const UPD_MODIFICATO_CONSEGNA = _
        " AND VVPIDVVC = ?"
    Const UPD_MODIFICATO_TESTATA = _
        "VVPIDVVT = ? "
    Const UPD_MODIFICATO_NULL = _
        "UPDATE VIAGVVP0 SET VVPFLST5 = NULL WHERE VVPFLST5 = 'S'"
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    
    strSQL = vbNullString
    If pUpdatedTestata Then
        strSQL = UPD_MODIFICATO + UPD_MODIFICATO_TESTATA
        DataClass.DS.DA.BeginTransaction
        If Not DataClass.DS.DA.ExecuteCommand(strSQL, pNewValue, pIdViaggio) Then
            GoTo ErrHandler
        End If
        DataClass.DS.DA.EndTrans ETCOMMIT
    ElseIf pUpdatedConsegna Then
        strSQL = UPD_MODIFICATO + UPD_MODIFICATO_TESTATA + UPD_MODIFICATO_CONSEGNA
        DataClass.DS.DA.BeginTransaction
        If Not DataClass.DS.DA.ExecuteCommand(strSQL, pNewValue, pIdViaggio, pIdConsegna) Then
            GoTo ErrHandler
        End If
        DataClass.DS.DA.EndTrans ETCOMMIT
    ElseIf pUpdatedProdotto Then
        strSQL = UPD_MODIFICATO + UPD_MODIFICATO_TESTATA + UPD_MODIFICATO_CONSEGNA + UPD_MODIFICATO_PRODOTTO
        DataClass.DS.DA.BeginTransaction
        If Not DataClass.DS.DA.ExecuteCommand(strSQL, pNewValue, pIdViaggio, pIdConsegna, pIdProdotto) Then
            GoTo ErrHandler
        End If
        DataClass.DS.DA.EndTrans ETCOMMIT
    ElseIf pNewValue = "NULL" Then
        strSQL = UPD_MODIFICATO_NULL
        DataClass.DS.DA.BeginTransaction
        If Not DataClass.DS.DA.ExecuteCommand(strSQL) Then
            GoTo ErrHandler
        End If
        DataClass.DS.DA.EndTrans ETCOMMIT
    End If
    UpdModificato = True
Exit Function
ErrHandler:
    DataClass.DS.DA.EndTrans ETROLLBACK
    UpdModificato = False
End Function

Private Function UpdViaggio(ByVal pIdViaggio As String, ByVal pIdConsegna As String, ByVal pIdProdotto As String, ByRef rErr As String) As Boolean
    Dim strErr As String
    Dim blnUpdatedTestata As Boolean
    Dim blnUpdatedConsegna As Boolean
    Dim blnUpdatedProdotto As Boolean
    
    On Error GoTo ErrHandler
    
    DataClass.DS.DA.BeginTransaction
    
    If Not TestataViaggioUpdated(blnUpdatedTestata) Then
        strErr = "Errore durante la verifica della modifica dei dati di testata del viaggio."
        GoTo ErrHandler
    End If
    If blnUpdatedTestata Then
        If Not UpdTestataViaggio(pIdViaggio, strErr) Then
            GoTo ErrHandler
        End If
    End If
    If Not ConsegnaViaggioUpdated(blnUpdatedConsegna) Then
        strErr = "Errore durante la verifica della modifica dei dati di consegna del viaggio."
        GoTo ErrHandler
    End If
    If blnUpdatedConsegna Then
        If Not UpdConsegnaViaggio(pIdViaggio, pIdConsegna, strErr) Then
            GoTo ErrHandler
        End If
    End If
    If Not ProdottoViaggioUpdated(blnUpdatedProdotto) Then
        strErr = "Errore durante la verifica della modifica dei dati del prodotto del viaggio."
        GoTo ErrHandler
    End If
    If blnUpdatedProdotto Then
        If Not UpdProdottoViaggio(pIdViaggio, pIdConsegna, pIdProdotto, strErr) Then
            GoTo ErrHandler
        End If
    End If
    If Not UpdModificato(pIdViaggio, pIdConsegna, pIdProdotto, blnUpdatedTestata, blnUpdatedConsegna, blnUpdatedProdotto, "S") Then
        strErr = "Errore durante la modifica del flag di aggiornamento."
        GoTo ErrHandler
    End If
    
    DataClass.DS.DA.EndTrans ETCOMMIT
    UpdViaggio = True
    Exit Function
    
ErrHandler:
    DataClass.DS.DA.EndTrans ETROLLBACK
    If strErr <> vbNullString Then
        rErr = strErr
    End If
    UpdViaggio = False
End Function

Private Sub PulisciDatiModificatiViaggio()
    MaskEdBox1.Text = TIME_MASK         ' Ora Inizio Carico
    MaskEdBox4.Text = vbNullString      ' Unità di Trasporto
    MaskEdBox3.Text = vbNullString      ' Motrice
    MaskEdBox5.Text = vbNullString      ' Rimorchio
    MaskEdBox6.Text = vbNullString      ' Autista 1
    MaskEdBox32.Text = vbNullString     ' Vettore
    MaskEdBox33.Text = vbNullString     ' Vettore Trasportatore
    MaskEdBox8.Text = vbNullString      ' N.I.F.
    MaskEdBox49.Text = DATE_MASK        ' Data Consegna
    MaskEdBox9.Text = TIME_MASK         ' Ora Consegna
    MaskEdBox48.Text = DATE_MASK        ' Data Prevista
    MaskEdBox30.Text = TIME_MASK        ' Ora Prevista
    MaskEdBox10.Text = vbNullString     ' Progressivo Consegna
    Check1.value = vbGrayed             ' Flag Fine Carico
    MaskEdBox24.Text = vbNullString     ' Km CEPSA
    MaskEdBox17.Text = vbNullString     ' Km Definitivi
    MaskEdBox25.Text = vbNullString     ' Note
    TxtNoteEstese.Text = vbNullString   ' Note Estese
    MaskEdBox28.Text = vbNullString     ' Q.tà Caricata L.
    MaskEdBox29.Text = vbNullString     ' Q.tà Scaricata L.
    MaskEdBox43.Text = TIME_MASK
    MaskEdBox35.Text = TIME_MASK
    strOraInizioCaricoValore = vbNullString
    strOraFineCaricoValore = vbNullString
    strOraFineScaricoValore = vbNullString
    strUdTValoreValore = vbNullString
    strMotriceValore = vbNullString
    strKmViaggioValore = vbNullString
    strRimorchioValore = vbNullString
    strAutista1Valore = vbNullString
    strAutista2Valore = vbNullString
    strVettore1Valore = vbNullString
    strVettore2Valore = vbNullString
    strPedaggioValore = vbNullString
    strOraConsegnaValore = vbNullString
    strProgressivoValore = vbNullString
    strKmValore = vbNullString
    strNoteValore = vbNullString
    strNoteEsteseValore = vbNullString
    strInizioScaricoValore = vbNullString
    blnFineCaricoValore = False
    strQtaCaricataValore = vbNullString
    strQtaScaricataValore = vbNullString
    blnNoGPSDataValore = False
    strSequenzaViaggioValore = vbNullString
End Sub

Private Function SincronizzaDatiViaggio(ByVal pRichiesta As SincronizzazioneRichiesta) As Boolean
      
        
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim BKM As Variant
    Dim strErr As String
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim strIDProdotto As String
    Dim strNIF As String
    Dim lngRiga As Long
    Dim SQL_Eventos As String
    Dim Buffev As Variant
    Dim lngev As Long
    Dim Rowsev As Long
    Dim Codcli As String
    Dim Viaje As String
    Dim blnRigaGiàAggiornata As Boolean
    On Error GoTo ErrHandler
    
    BKM = TDBGrid1.bookmark
    strIDViaggio = XA.value(BKM, 0)
    strIDConsegna = XA.value(BKM, 1)
    strIDProdotto = XA.value(BKM, 2)
    Select Case pRichiesta
        Case srLetturaDati
            ' -- DP 08/05/08
            ' -- Aggiungo il campo MaskEdBox47 per la visualizzazione dell'ora programmata per la consegna (VVCTEMVG)
            MaskEdBox47 = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 41), 4), 2, 5)                                  ' Ora Consegna Programmata
        
            ' -- DP 06/05/08
            ' -- Aggiungo il campo MaskEdBox46 per la visualizzazione della data di fine carico (data inizio viaggio)
            MaskEdBox46 = Format(XA.value(BKM, 3), "dd/mm/yy")                                                      ' Data Fine Carico (Data Inizio Viaggio)
            strDataFineCarico = MaskEdBox46.Text
            
            MaskEdBox42 = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 9), 4), 2, 5)                                   ' Ora Fine Carico (VVTHRINV - Inizio viaggio)
            strOraFineCaricoValore = MaskEdBox42.Text
            MaskEdBox1 = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 10), 4), 2, 5)                                   ' Ora Inizio Carico
            strOraInizioCaricoValore = MaskEdBox1
            MaskEdBox4.Text = IIf(IsNull(XA.value(BKM, 11)), vbNullString, XA.value(BKM, 11))                       ' Unità di Trasporto
            strUdTValoreValore = MaskEdBox4.Text
            MaskEdBox3.Text = IIf(IsNull(XA.value(BKM, 12)), vbNullString, XA.value(BKM, 12))                       ' Motrice
            strMotriceValore = MaskEdBox3.Text
            MaskEdBox5.Text = IIf(IsNull(XA.value(BKM, 13)), vbNullString, XA.value(BKM, 13))                       ' Rimorchio
            strRimorchioValore = MaskEdBox5.Text
            MaskEdBox6.Text = IIf(IsNull(XA.value(BKM, 14)), vbNullString, XA.value(BKM, 14))                       ' Autista 1
            MaskEdBox7.Text = IIf(IsNull(XA.value(BKM, 17)), vbNullString, XA.value(BKM, 17))
            MaskEdBox18.Text = IIf(IsNull(XA.value(BKM, 15)), vbNullString, XA.value(BKM, 15))                      ' Km viaggio
            strKmViaggioValore = MaskEdBox18.Text
            strAutista1Valore = MaskEdBox6.Text
            MaskEdBox19.Text = IIf(IsNull(XA.value(BKM, 35)), 0, XA.value(BKM, 35))                                 ' Pedaggi
            strPedaggioValore = MaskEdBox19.Text
            
            If Not IsEmptyVariantString(XA.value(BKM, 44)) And XA.value(BKM, 44) = "N" Then
                blnNoGPSDataValore = True
                chkNoGPSData.value = vbChecked
            Else
                blnNoGPSDataValore = False
                chkNoGPSData.value = vbUnchecked
            End If
            MaskEdBox11.Text = XA.value(BKM, 8)
            strSequenzaViaggioValore = MaskEdBox11.Text
            
            Call DataClass.DS.ValidateClientField("34", strAutista1Valore)
            FieldUpdateDescription "34", Me, MaskEdBox6
            FieldSetDescription "34", Me, strAutista1Valore, MaskEdBox6
            If IsNull(XA.value(BKM, 16)) Then
                MaskEdBox8 = vbNullString
            Else
                MaskEdBox8.Text = XA.value(BKM, 16)                                                                            ' Autista 2
            End If

            ' Vettore
            MaskEdBox32.Text = XA.value(BKM, 42)
            strVettore1Valore = MaskEdBox32.Text
            Call DataClass.DS.ValidateClientField("32", strVettore1Valore)
            FieldUpdateDescription "32", Me, MaskEdBox32
            FieldSetDescription "32", Me, strVettore1Valore, MaskEdBox32

            ' Vettore Trasp.
            MaskEdBox33.Text = XA.value(BKM, 43)
            strVettore2Valore = MaskEdBox33.Text
            Call DataClass.DS.ValidateClientField("39", strVettore2Valore)
            FieldUpdateDescription "39", Me, MaskEdBox33
            FieldSetDescription "39", Me, strVettore2Valore, MaskEdBox33
            
            MaskEdBox34.Text = XA.value(BKM, 47)                                                                            ' Telefono Autista
            MaskEdBox35.Text = XA.value(BKM, 48)                                                                            ' Note Ordine
            MaskEdBox37.Text = Format(XA.value(BKM, 49), "hh") & ":" & Format(XA.value(BKM, 49), "nn")                                                           ' Dalle ore 1
            MaskEdBox36.Text = Format(XA.value(BKM, 50), "hh") & ":" & Format(XA.value(BKM, 50), "nn")                                                          ' Alle ore 1
            MaskEdBox39.Text = Format(XA.value(BKM, 51), "hh") & ":" & Format(XA.value(BKM, 51), "nn")                                                            ' Dalle ore 2
            MaskEdBox38.Text = Format(XA.value(BKM, 52), "hh") & ":" & Format(XA.value(BKM, 52), "nn")                                                           ' Alle ore 2
            
            
            '-- DP 30/09/08
            '-- Gestione data e ora prevista consegna
            '-- Se non sono state aggiornate non le visualizzo
            If Format(XA.value(BKM, 65), "dd/mm/yy") <> "01/01/00" Then
'                MaskEdBox48.Text = Format(IIf(Format(XA.value(BKM, 65), "dd/mm/yy") <> "01/01/00", XA.value(BKM, 65), XA.value(BKM, 66)), "dd/mm/yy")
                MaskEdBox48.Text = Format(XA.value(BKM, 65), "dd/mm/yy")
                
                MaskEdBox30.Text = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 22), 4), 2, 5)       ' Ora Prevista
'                If MaskEdBox30 = "00:00" Then
'                    '-- DP 15/05/08
'                    '-- Utilizzo il nuovo campo VVCHRARR al posto di VVCTEMVG
'                    MaskEdBox30 = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 64), 4), 2, 5)
'                End If
            End If
            strDataPrevistaValore = MaskEdBox48.Text
            strOraPrevistaValore = MaskEdBox30.Text
            
            
            MaskEdBox10.Text = IIf(IsNull(XA.value(BKM, 18)), vbNullString, XA.value(BKM, 18))                      ' Progressivo Consegna
            strProgressivoValore = MaskEdBox10.Text
            MaskEdBox11.Text = XA.value(BKM, 8)                                                                     ' Progressivo viaggio
            
            If (IsNull(XA.value(BKM, 25)) Or (XA.value(BKM, 25) = vbNullString) Or (XA.value(BKM, 25) = "N")) Then  ' Flag Fine Scarico
                Check1.value = vbUnchecked
                strInizioScaricoValore = "N"
                
                ' -- DP 06/05/08
                ' -- Se la consegna non è iniziata non devono essere visualizzate data e ora di arrivo, trascrivo il campo così com è nella variabile associata
                strDataConsegnaValore = MaskEdBox49.Text
                strOraConsegnaValore = MaskEdBox9.Text
            Else
                Check1.value = vbChecked
                strInizioScaricoValore = "S"
                
                '-- DP 16/05/08
                '-- Se la consegna è iniziata deve essere visualizzata la data di arrivo
                MaskEdBox49.Text = Format(XA.value(BKM, 66), "dd/mm/yy")
                strDataConsegnaValore = MaskEdBox49.Text
                
                ' -- DP 06/05/08
                ' -- Se la consegna è iniziata deve essere visualizzata l'ora di arrivo
                MaskEdBox9.Text = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 64), 4), 2, 5)       ' Ora Arrivo
                strOraConsegnaValore = MaskEdBox9.Text
            End If
            
            If XA.value(BKM, 37) = "S" Then
                Check2.value = vbChecked
                blnFineCaricoValore = Check2.value
            Else
                Check2.value = vbUnchecked
                blnFineCaricoValore = Check2.value
            End If
            
            If XA.value(BKM, 57) = "S" Then
                Check3.value = vbChecked
                strInizioCaricoValore = "S"
            Else
                Check3.value = vbUnchecked
                strInizioCaricoValore = "N"
            End If
            If XA.value(BKM, 59) = "S" Then
                Check4.value = vbChecked
                strFineScaricoValore = "S"
            Else
                Check4.value = vbUnchecked
                strFineScaricoValore = "N"
            End If
            If XA.value(BKM, 60) = "S" Then
                chkNoConforme.value = vbChecked
                strFlagNoConforme = "S"
            Else
                chkNoConforme.value = vbUnchecked
                strFlagNoConforme = "N"
            End If
            If XA.value(BKM, 61) = "S" Then
                chkDesigillato.value = vbChecked
                strDesigillato = "S"
            Else
                chkDesigillato.value = vbUnchecked
                strDesigillato = "N"
            End If
            
            MaskEdBox24.Text = IIf(IsNull(XA.value(BKM, 23)), vbNullString, XA.value(BKM, 23))                  ' Km Raggio (CEPSA)
            If strInizioScaricoValore = "S" Then
                MaskEdBox17.Text = IIf(IsNull(XA.value(BKM, 24)), vbNullString, XA.value(BKM, 24))                  ' Km Definitivi
            Else
                MaskEdBox17.Text = vbNullString
            End If
            strKmValore = MaskEdBox17.Text
            MaskEdBox25.Text = IIf(IsNull(XA.value(BKM, 26)), vbNullString, XA.value(BKM, 26))                      ' Note
            strNoteValore = MaskEdBox25.Text
            TxtNoteEstese.Text = IIf(IsNull(XA.value(BKM, 27)), vbNullString, XA.value(BKM, 27))                    ' Note Estese
            strNoteEsteseValore = TxtNoteEstese.Text
            MaskEdBox27.Text = Format(IIf(IsDate(XA.value(BKM, 36)), XA.value(BKM, 36), XA.value(BKM, 3)), "dd/mm/yy")                  ' Data consegna (se assente Data viaggio)
            strDataConsegna = MaskEdBox27.Text
            MaskEdBox28.Text = IIf(IsNull(XA.value(BKM, 30)), vbNullString, XA.value(BKM, 30))                      ' Q.tà Caricata L.
            strQtaCaricataValore = MaskEdBox28.Text
            MaskEdBox29.Text = IIf(IsNull(XA.value(BKM, 31)), vbNullString, XA.value(BKM, 31))                      ' Q.tà Scaricata L.
            strQtaScaricataValore = MaskEdBox29.Text
            If XA.value(BKM, 38) = "F" Then
                chkFestivo.value = vbChecked
            Else
                chkFestivo.value = vbUnchecked
            End If
            strDataViaggio = XA.value(BKM, 3)
            strSocietaViaggio = XA.value(BKM, 33)
            If Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 55), 4), 2, 5) = "00:00" Then
                MaskEdBox43.Text = TIME_MASK
                strOraFineScaricoValore = TIME_MASK
            Else
                MaskEdBox43.Text = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 55), 4), 2, 5)
                strOraFineScaricoValore = MaskEdBox43.Text
            End If
            If Not (IsNull(XA.value(BKM, 53))) Then MaskEdBox31 = XA.value(BKM, 53)                                                                         ' NIS Carico
            If Not (IsNull(XA.value(BKM, 54))) Then MaskEdBox40 = XA.value(BKM, 54)                                                                         ' NIS Scarico
            MaskEdBox41.Text = Left(XA.value(BKM, 19), 10)                                                          ' Numero Ordine
            With VPI.DS
                .Fields.ResetFields
                .Fields("0").FindValue = XA.value(BKM, 0)
                .Fields("1").FindValue = XA.value(BKM, 1)
                .Fields("2").FindValue = XA.value(BKM, 2)
            End With
            VPI.Find
            If VPI.DS.HowMany > 0 Then
                Me.chkFlagEventi.value = vbChecked
            Else
                Me.chkFlagEventi.value = vbUnchecked
            End If
            
            '-- DP 14/08/08
            MaskEdBox50.Text = IIf(IsNull(XA.value(BKM, 68)), vbNullString, XA.value(BKM, 68))
            MaskEdBox51.Text = IIf(IsNull(XA.value(BKM, 69)), vbNullString, XA.value(BKM, 69))
            MaskEdBox52.Text = IIf(IsNull(XA.value(BKM, 70)), vbNullString, XA.value(BKM, 70))
            MaskEdBox53.Text = IIf(IsNull(XA.value(BKM, 71)), vbNullString, XA.value(BKM, 71))
            MaskEdBox54.Text = IIf(IsNull(XA.value(BKM, 72)), vbNullString, XA.value(BKM, 72))
            
            '-- DP 20/08/08
            If Not IsNull(XA.value(BKM, 73)) Then
                MaskEdBox55.Text = Format(XA.value(BKM, 73), "dd/mm/yy")
                MaskEdBox56.Text = Mid(DataClass.DS.DA.MakeSQL(XA.value(BKM, 74), 4), 2, 5)
            End If
            Set XAEv = New XArrayDB
            SQL_Eventos = "SELECT     TOP 100 PERCENT GESCODEV, GESNUMLO, CONVERT(datetime, GESDYEV0, 113) AS Fecha, GESHREV0, GESRECEP, GESOBCEP " & _
                    "FROM         dbo.VIAGGES0 " & _
                    "GROUP BY GESCODEV, GESCODDE, IDVVT, GESNUMLO, GESDYEV0, GESHREV0, GESRECEP, GESOBCEP " & _
                    "HAVING      (IDVVT = ?) AND (GESCODDE = ?) " & _
                    "ORDER BY GESDYEV0,GESHREV0,GESCODEV "
            Codcli = Right(XA.value(BKM, 20), 8)
            Viaje = XA.value(BKM, 0)
            If DA.QueryCommand(SQL_Eventos, Buffev, lngev, Viaje, Codcli) Then
                For Rowsev = 1 To lngev
                
                    With XAEv
                                .ReDim 1, Rowsev, 0, 16
                                .value(Rowsev, 1) = Buffev(0, Rowsev - 1)
                                .value(Rowsev, 2) = Buffev(3, Rowsev - 1)
                                .value(Rowsev, 3) = Buffev(2, Rowsev - 1)
                                .value(Rowsev, 4) = Buffev(1, Rowsev - 1)
                                .value(Rowsev, 5) = Buffev(4, Rowsev - 1)
                                .value(Rowsev, 10) = Buffev(5, Rowsev - 1)
                    End With
                Next Rowsev
            End If
            With TDBGrid2
                .Array = XAEv
                .ReBind
                .bookmark = 1
            End With
        
        Case srScritturaDati
            ' Aggiorno i dati sul viaggio
            If Not UpdViaggio(strIDViaggio, strIDConsegna, strIDProdotto, strErr) Then
                GoTo ErrHandler
            End If
    End Select
    SincronizzaDatiViaggio = True
Exit Function
ErrHandler:
    SincronizzaDatiViaggio = False
    If strErr <> vbNullString Then
        Call MsgBox(strErr, vbOKOnly + vbExclamation, "Errore")
    End If
End Function

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim strErr As String
    
    ClientKeyUp KeyCode, Me
    Select Case KeyCode
        Case vbKeyF2
            ' Salva i dati eventualmente modificati sul viaggio
            If Not TDBGrid1.Enabled Then
                If SincronizzaDatiViaggio(srScritturaDati) Then
                    ' Chiude il frame di modifica viaggi e rinfresca il viaggio sulla griglia
                    Call RipristinaSelezione
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strIDParam As String
    ReDim ColTags(1 To 2) As Variant
    Dim Col As Column
    Dim strTimerMinutes As String
    Dim varDummy As Variant

    
    Loading = True
    mintCurFrame = 1
    ColTags(1) = "0"
    ColTags(2) = "1"
    mskMaxRecords = MAX_RECORDS
    
    Set mobjOutFile = New CTextFile
    Set mobjOutFileInc = New CTextFile
    
    Select Case mvarFunction
        Case fcEXPSNPV, fcEXPSNP2, fcEXPSNPA, fcEXPSNPA2
            If mvarFunction = fcEXPSNPA2 Then
                mskFileName = gmRegistry.GetSetting("Options", "TripChangedExport2", , True)
            Else
                mskFileName = gmRegistry.GetSetting("Options", "TripChangedExport", , True)
            End If
            
            For i = 1 To 4
                varDummy = gmRegistry.GetSetting("Options", "TrasmAllTime" & CStr(i), , True)
                If Len(varDummy) > 0 Then
                    Me.mskOrario(i).Text = varDummy
                Else
                    Me.mskOrario(i).Text = TIME_MASK
                End If
            Next i
            
            
            With mobjOutFile
                .Add "Ente Emittente", 4, "1"
                .Add "Progressivo Invio", 10, "2", , fftNumerico
                .Add "Data Sistema", 8, "3"
                .Add "Ora Sistema", 8, "4"
                .Add "Numero Record", 4, "5", , fftNumerico
                .Add "Rif. Esterno 1 Consegna", 13, "6"
                .Add "Partizione", 2, "7", "00"
                .Add "Destinazione", 8, "8"
                .Add "GPS", 2, "9", "01", fftNumerico
                .Add "X GPS", 9, "10"
                .Add "Y GPS", 9, "11"
                .Add "Prodotto", 7, "12"
                .Add "Deposito", 4, "13"
                .Add "Q.tà Concordata Ordine", 9, "14", , fftNumerico
                .Add "Unità di Misura", 1, "15", "L"
                .Add "Data Viaggio", 8, "16"
                .Add "Data Consegna aggiornata", 8, "17"            ' Inizia a Colonna 107
                .Add "Ora Consegna aggiornata", 5, "18"
                .Add "Q.tà Trasportata", 9, "19", , fftNumerico
                .Add "Q.tà Scaricata", 9, "20", , fftNumerico
                .Add "Data Viaggio 2", 8, "21"
                .Add "Ora Fine Carico", 5, "22"
                .Add "Data Consegna effettiva", 8, "23"             ' Inizia a Colonna 138
                .Add "Ora Consegna effettiva", 5, "24"
                .Add "Badge UdT", 11, "25"
                .Add "Nr. Sequenza", 2, "26"
                .Add "Nr. Carta Identità", 20, "27"
                .Add "Km Definitivi", 4, "28", , fftNumerico
                .Add "Km CEPSA", 4, "29", , fftNumerico
                .Add "Quantità 1", 6, "30", "000.00", fftAlfa
                .Add "Quantità 2", 6, "31", "000.00", fftAlfa
                .Add "Quantità 3", 6, "32", "000.00", fftAlfa
                .Add "Numero Consegna", 2, "33", , fftNumerico
                .Add "Note Consegna", 40, "34"
                .Add "Codice evento", 4, "35"
                .Add "Data Evento", 8, "36"                         ' Inizia a Colonna 269
                .Add "Ora Evento", 5, "37"
                .Add "Stato compartimento 1", 4, "38"
                .Add "Stato compartimento 2", 4, "39"
                .Add "Stato compartimento 3", 4, "40"
                .Add "Stato compartimento 4", 4, "41"
                .Add "Stato compartimento 5", 4, "42"
                .Add "Stato compartimento 6", 4, "43"
                .Add "Stato compartimento 7", 4, "44"
                .Add "Filler", 55, "45"
                .Add "Flag consegna", 1, "46"
                .Add "Accordi con cliente", 1, "47"
                .Add "Ora carico prevista", 5, "48"
            End With
            
            With mobjOutFileInc
                .Add "Ente Emittente", 4, "1", "0056"
                .Add "Progressivo Invio", 10, "2", , fftNumerico
                .Add "Data Generazione", 8, "3"
                .Add "Ora Generazione", 8, "4"
                .Add "Numero Records", 4, "5", , fftNumerico
                .Add "Riferimento Ordine", 13, "6", , fftNumerico               ' Ordine + riga
                .Add "Partizione Ordine", 2, "7", "00", fftNumerico
                .Add "Codice Destinazione", 8, "8"
                .Add "Tipo evento", 3, "9", , fftNumerico
                .Add "Motivo evento", 3, "10", , fftNumerico
                .Add "Data evento", 8, "11"
                .Add "Ora inizio evento", 5, "12"
                .Add "Ora termine evento", 5, "13"
                .Add "Ordine precedente responsabile evento", 13, "14", , fftNumerico
                .Add "Partizione ordine precedente", 2, "15", , fftNumerico
                .Add "Evento precedente", 3, "16", , fftNumerico
                
                .Add "Prodotto contaminante (CLH)", 7, "17"
                .Add "Quantità prodotto contaminante", 9, "18", , fftNumerico
                .Add "Unità di misura prodotto contaminante", 1, "19"
                .Add "Prodotto contaminato 1 (CLH)", 7, "20"
                .Add "Quantità prodotto contaminato 1", 9, "21", , fftNumerico
                .Add "Unità di misura prodotto contaminato 1", 1, "22"
                .Add "Prodotto contaminato 2 (CLH)", 7, "23"
                .Add "Quantità prodotto contaminato 2", 9, "24", , fftNumerico
                .Add "Unità di misura prodotto contaminato 2", 1, "25"
                .Add "Prodotto contaminato 3 (CLH)", 7, "26"
                .Add "Quantità prodotto contaminato 3", 9, "27", , fftNumerico
                .Add "Unità di misura prodotto contaminato 3", 1, "28"
                .Add "Prodotto contaminato 4 (CLH)", 7, "29"
                .Add "Quantità prodotto contaminato 4", 9, "30", , fftNumerico
                .Add "Unità di misura prodotto contaminato 4", 1, "31"
                
                .Add "Prodotto spandimento (CLH)", 7, "32"
                .Add "Quantità prodotto spandimento", 9, "33", , fftNumerico
                .Add "Unità di misura prodotto spandimento", 1, "34"
                .Add "Misure adottate", 50, "35"
                
                .Add "Prodotto rientrato (CLH)", 7, "36"
                .Add "Quantità prodotto rientrato", 9, "37", , fftNumerico
                .Add "Unità di misura prodotto rientrato", 1, "38"
                .Add "Deposito rientro (CLH)", 4, "39"
                .Add "Data rientro in deposito", 8, "40"
                
                .Add "Prodotto dirottato (CLH)", 7, "41"
                .Add "Quantità prodotto dirottato", 9, "42", , fftNumerico
                .Add "Unità di misura prodotto dirottato", 1, "43"
                .Add "Destinazione dirottamento", 8, "44"
                .Add "Data consegna prodotto dirottato", 8, "45"
                
                .Add "Ora inizio paralisi deposito", 5, "46"
                .Add "Ora termine paralisi deposito", 5, "47"
                .Add "Ora inizio paralisi destinazione", 5, "48"
                .Add "Ora termine paralisi destinazione", 5, "49"
                
                .Add "Data comunicazione evento", 8, "50"
                .Add "Responsabile comunicazione evento", 40, "51"
                .Add "Osservazioni", 200, "52"
                
            End With
    
        Case fcEXPSNPVA
    
            For i = 1 To 4
                Me.mskOrario(i).Visible = False
            Next i
            
            Caption = "TRASMISSIONE ANOMALIE AGIP"
            mskFileName = gmRegistry.GetSetting("Options", "TripChangedExportAgip", , True)
            
            With mobjOutFile
                .Add "Tipo Record", 1, "1", "1", fftAlfa
                .Add "Ente Emittente", 4, "2", , fftAlfa
                .Add "Progressivo Invio", 3, "3", , fftNumerico
                .Add "Data Invio", 8, "4", , fftAlfa
                .Add "Ora Invio", 5, "5", , fftAlfa
            End With
            
            With mobjOutFileInc
                .Add "Motivo Evento", 2, "1", , fftAlfa
                .Add "Osservazioni", 60, "2", , fftAlfa
                .Add "Data Evento", 8, "3", , fftAlfa
                .Add "Unità di Trasporto", 3, "4", , fftAlfa
                .Add "Badge", 3, "5", , fftAlfa
                .Add "Sequenza Viaggio", 2, "6", , fftNumerico
                .Add "Sequenza Consegna", 2, "7", , fftNumerico
                .Add "Sequenza Prodotto", 2, "8", , fftNumerico
                .Add "Area", 2, "9", , fftAlfa
                .Add "Deposito", 2, "10", , fftAlfa
                .Add "Raggruppamento deposito", 2, "11", , fftAlfa
                .Add "Numero de pedido", 8, "12", , fftAlfa
                .Add "Prodotto", 6, "13", , fftAlfa
                .Add "Destinazione", 8, "14", , fftAlfa
                .Add "Nome Destinazione", 30, "15", , fftAlfa
                .Add "Tipo vendita", 2, "16", , fftAlfa
                .Add "Momento del servizio", 1, "17", , fftAlfa
            End With

    End Select

    Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .Field("Ente Emittente").Required = True
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set VVC = New CViaggioConsegna
    With VVC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set VPI = New CViaggioEvento
    With VPI.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AME = New CMotivoEvento
    With AME.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AEE = New CEnteEmittente
    With AEE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set ARP = New CRagDeposito
    With ARP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set ATV = New CTipoVendita
    With ATV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    Set SEV = New CSatellitareEvento
    With SEV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = strIDParam
        AXX.Find
    End With
    ' ---
    For Each Col In TDBGrid1.Columns
        Col.ButtonHeader = True
    Next Col

    mlngTickTimer = 0
    mblnAutomaticSend = False
    TimerAutomaticSend.Enabled = False
    Select Case mvarFunction
        Case fcEXPSNPA, fcEXPSNPA2
            strTimerMinutes = gmRegistry.GetSetting("Options", "TimerMinutesCepsaTrasm", "10", True)
            If IsNumeric(strTimerMinutes) Then
                mlngTimerMinutes = CLng(strTimerMinutes)
                If mlngTimerMinutes > 99 Then
                    mlngTimerMinutes = 99
                End If
            Else
                mlngTimerMinutes = 10
            End If
            mskMinutes.Text = mlngTimerMinutes
            lbTimer.Caption = mlngTickTimer
            cmdStartStopAutomatic.Visible = True
            Picture1(8).Visible = True
            SetState fsparaminput
            SetState fsSelection
            mblnAutomaticSend = True
            TimerAutomaticSend.Enabled = True
            
        Case fcEXPSNPV ' Cepsa
            mskFileName = gmRegistry.GetSetting("Options", "TripChangedExport", , True)
            cmdStartStopAutomatic.Visible = False
            SetState fsparaminput
            
        Case fcEXPSNP2 ' Guarner
            mskFileName = gmRegistry.GetSetting("Options", "TripChangedExport2", , True)
            cmdStartStopAutomatic.Visible = False
            SetState fsparaminput

        Case Else
            cmdStartStopAutomatic.Visible = False
            SetState fsparaminput
            
    End Select

    CmdOpenTrip.Enabled = False
    cmdOpenFile.Enabled = False
    cmdExport.Enabled = False

    Loading = False
    
'Sm 11/05/2010 Cargar el ddl con los eventos de cepsa
    
   cmb_cges.AddItem "002-Entrada en Factoria"
   cmb_cges.ItemData(0) = "002"
   cmb_cges.AddItem "003-Salida de Factoria"
   cmb_cges.ItemData(1) = "003"
   cmb_cges.AddItem "005-Llegada a destino"
   cmb_cges.ItemData(2) = "005"
   cmb_cges.AddItem "006-Salida de destino"
   cmb_cges.ItemData(3) = "006"
   cmb_cges.AddItem "008-Desprecintado"
   cmb_cges.ItemData(4) = "008"
   cmb_cges.AddItem "009-Imposible descarga total"
   cmb_cges.ItemData(5) = "009"
   cmb_cges.AddItem "010-Imposible descarga parcial"
   cmb_cges.ItemData(6) = "010"
   cmb_cges.AddItem "011-Contaminacion"
   cmb_cges.ItemData(7) = "011"
   cmb_cges.AddItem "012-Derrame"
   cmb_cges.ItemData(8) = "012"
   cmb_cges.AddItem "013-Discrepancia cantidad"
   cmb_cges.ItemData(9) = "013"
   cmb_cges.AddItem "014-Accidente"
   cmb_cges.ItemData(10) = "014"
   cmb_cges.AddItem "015-Averia"
   cmb_cges.ItemData(11) = "015"
   cmb_cges.AddItem "016-Devolucion"
   cmb_cges.ItemData(12) = "016"
   cmb_cges.AddItem "017-Desvio"
   cmb_cges.ItemData(13) = "017"
   cmb_cges.AddItem "018-Degradación"
   cmb_cges.ItemData(14) = "018"
   cmb_cges.AddItem "019-Paralización en centro de carga"
   cmb_cges.ItemData(15) = "019"
   cmb_cges.AddItem "020-Paralización en destino"
   cmb_cges.ItemData(16) = "020"
   cmb_cges.AddItem "001-Salida De Base"
   cmb_cges.ItemData(17) = "001"
   cmb_cges.AddItem "007-Vuelta a Base"
   cmb_cges.ItemData(18) = "007"
    


    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Cancel = 0
'    Select Case mvarFunction
'
'        Case fcEXPSNPV, fcEXPSNP2
'            If lblLog(5).Caption = "0" Then
'                If Not UpdModificato(vbNullString, vbNullString, vbNullString, False, False, False, "NULL") Then
'                    Cancel = 1
'                End If
'            Else
'                If MsgBox("Esta saliendo sin transmitir los eventos. ¿Desea salir?", vbYesNo + vbExclamation, "Attenzione") <> vbYes Then
'                    Cancel = 1
'                Else
'                    If Not UpdModificato(vbNullString, vbNullString, vbNullString, False, False, False, "") Then
'                        Cancel = 1
'                    End If
'                End If
'            End If
'
'    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = 1 To 4
        gmRegistry.SaveSetting "Options", "TrasmAllTime" & CStr(i), Me.mskOrario(i).Text, True
    Next i
    
    Select Case mvarFunction
        Case fcEXPSNPA ' Cepsa Auto
            gmRegistry.SaveSetting "Options", "TimerMinutesCepsaTrasm", mskMinutes, True
            gmRegistry.SaveSetting "Options", "TripChangedExport", mskFileName, True
        Case fcEXPSNPA2 ' Cepsa Auto
            gmRegistry.SaveSetting "Options", "TimerMinutesCepsaTrasm", mskMinutes, True
            gmRegistry.SaveSetting "Options", "TripChangedExport2", mskFileName, True
        Case fcEXPSNPV ' Cepsa
            gmRegistry.SaveSetting "Options", "TripChangedExport", mskFileName, True

        Case fcEXPSNPVA
            gmRegistry.SaveSetting "Options", "TripChangedExportAgip", mskFileName, True

        Case fcEXPSNP2 ' Guarner
            gmRegistry.SaveSetting "Options", "TripChangedExport2", mskFileName, True
    End Select
    
    Set AUT = Nothing
    Set AAU = Nothing
    Set AAM = Nothing
    Set AAM = Nothing
    Set ADE = Nothing
    Set ADZ = Nothing
    Set AXX = Nothing
    Set VPI = Nothing
    Set AME = Nothing
    Set AZO = Nothing
    Set ARP = Nothing
    Set ATV = Nothing
    Set XA = Nothing
    Set XALog = Nothing
    Set mobjOutFile = Nothing
    Set mobjOutFileInc = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set frmVVTSnapshotExport = Nothing
End Sub

Private Sub MaskEdBox4_GotFocus()
    UnitàDiTrasportoAppoggio = MaskEdBox4.Text
End Sub

Private Sub MaskEdBox4_LostFocus()
    ' Se è cambiata l'unità di trasporto rinfresco anche i dati dell'automezzo e degli autisti
    If UnitàDiTrasportoAppoggio <> MaskEdBox4.Text Then
        If MaskEdBox4.Text = vbNullString Then
            MaskEdBox3.Text = vbNullString
            MaskEdBox5.Text = vbNullString
            MaskEdBox6.Text = vbNullString
            MaskEdBox8.Text = vbNullString
            MaskEdBox32.Text = vbNullString
            MaskEdBox33.Text = vbNullString
        Else
            With AUT.DS.Fields
                .ResetFields
                .Item("0").FindValue = MaskEdBox4.Text
                If Not AUT.Find Or AUT.DS.HowMany = 0 Then
                    Call MsgBox("Unità di trasporto non trovata", vbOKOnly + vbExclamation, "Attenzione")
                    MaskEdBox3.Text = vbNullString
                    MaskEdBox5.Text = vbNullString
                    MaskEdBox6.Text = vbNullString
                    MaskEdBox8.Text = vbNullString
                    MaskEdBox32.Text = vbNullString
                    MaskEdBox33.Text = vbNullString
                Else
                    MaskEdBox3.Text = .Item("2").value
                    MaskEdBox5.Text = .Item("3").value
                    MaskEdBox6.Text = .Item("4").value
                    MaskEdBox32.Text = .Item("6").value
                    MaskEdBox33.Text = .Item("7").value
                End If
            End With
        End If
        
        ' Autista 1
        Call DataClass.DS.ValidateClientField("34", MaskEdBox6)
        FieldUpdateDescription "34", Me, MaskEdBox6
        FieldSetDescription "34", Me, strAutista1Valore, MaskEdBox6
            
        ' Vettore
        Call DataClass.DS.ValidateClientField("32", MaskEdBox32)
        FieldUpdateDescription "32", Me, MaskEdBox32
        FieldSetDescription "32", Me, strVettore1Valore, MaskEdBox32

        ' Vettore Trasp.
        Call DataClass.DS.ValidateClientField("39", MaskEdBox33)
        FieldUpdateDescription "39", Me, MaskEdBox33
        FieldSetDescription "39", Me, strVettore2Valore, MaskEdBox33
    End If
End Sub

Private Sub MaskEdBox6_LostFocus()
    With AAU.DS.Fields
        .ResetFields
        .Item("0").FindValue = MaskEdBox6.Text
        If Not AAU.Find Or AAU.DS.HowMany = 0 Then
            Call MsgBox("Autista non trovato", vbOKOnly + vbExclamation, "Attenzione")
            MaskEdBox8.Text = vbNullString
        Else
            MaskEdBox8.Text = AAU.DS.Fields("9").value
        End If
    End With
End Sub

Private Sub mskMinutes_LostFocus()
    If IsNumeric(mskMinutes.Text) Then
        mlngTimerMinutes = CLng(mskMinutes.Text)
    End If
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurFrame Then
        Exit Sub
    End If
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    Frame1(mintCurFrame).Visible = False
    mintCurFrame = TabStrip1.SelectedItem.Index
    
    Select Case mvarFunction
        Case fcEXPSNPA, fcEXPSNPA2
        
        Case Else
            If mintCurFrame = 1 Then
                CmdOpenTrip.Enabled = False
                cmdOpenFile.Enabled = False
                cmdExport.Enabled = False
            Else
                CmdOpenTrip.Enabled = True
                cmdOpenFile.Enabled = True
                cmdExport.Enabled = True
            End If
    End Select

End Sub

Private Sub TDBGrid1_DblClick()
    DoCommand fcOpen
End Sub

' Dopo... La uso per colorare le CMP modificate
Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If (XA(bookmark, 25) = "S") Then
        RowStyle.BackColor = vbGreen
        RowStyle.ForeColor = vbBlack
    Else
        If (XA(bookmark, 32) = "S") Then
            RowStyle.BackColor = vbRed
            RowStyle.ForeColor = vbWhite
        End If
    End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim x As XTYPE
    Dim o As XORDER
    Static blnLastSortASC  As Boolean
    Static intLastCol As Integer
    
    If ColIndex <> intLastCol Then
        blnLastSortASC = False
    Else
        blnLastSortASC = Not blnLastSortASC
    End If
    If blnLastSortASC Then
        o = XORDER_DESCEND
    Else
        o = XORDER_ASCEND
    End If
        
    Select Case ColIndex
        Case 0, 3, 4, 16, 17, 18, 20, 21
            x = XTYPE_LONG
        Case 5, 6, 7, 15
            x = XTYPE_DATE
        Case Else
            x = XTYPE_STRING
    End Select
    XA.QuickSort 0, XA.UpperBound(1), ColIndex, o, x
    TDBGrid1.ReBind
    intLastCol = ColIndex
End Sub


Private Sub ExportRecords()
    On Error GoTo ErrHandler

    Const TEMP_FILE = "GES"
    
    Dim blnNeedTrasmission As Boolean
    Dim blnConsegnaIniziata As Boolean
    Dim blnUltimaConsegna As Boolean
    Dim blnPrimaConsegna As Boolean
    Dim blnConcErr As Boolean
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngNumViaggiModificati As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngProgrConsegna As Long
    Dim lngKMCepsa As Long
    Dim lngKm As Long
    Dim lngPortataVeicoloKG As Long
    Dim lngTotQuantitàKG As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim intRowsTrip As Integer
    Dim intFileFromNum As Integer
    Dim intFileToNum As Integer
    Dim intProgressivoViaggio As Integer
    Dim intProgressivoConsegna As Integer
    Dim intCountSospesi As Integer
    Dim strCodOperatore As String
    Dim strDestinazione As String
    Dim strCartaIdentitaAutista1 As String
    Dim strIDZona As String
    Dim strIDZonaPrec As String
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim strIDScarico As String
    Dim strIDViaggioPrec As String
    Dim strIDEnteEmittente As String
    Dim strIDDeposito As String
    Dim strIDSocietàDeposito As String
    Dim strProdotto As String
    Dim strConsegnaIniziata As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strDataViaggio As String
    Dim strDataInvio As String
    Dim strOraInvio As String
    Dim strProgrInvio As String
    Dim strOraConsegna As String
    Dim strOraPrevista As String
    Dim strOraInizioCarico As String
    Dim strOraFineCarico As String
    Dim strIDMotrice As String
    Dim varBadge As Variant
    Dim strCentroOperativo As String
    Dim strUnitàTrasporto As String
    Dim strDestinatario As String
    Dim strFileName As String
    Dim strRifAGIP As String
    Dim strGPSX As String
    Dim strGPSY As String
    Dim strStatoOrdine As String
    Dim strTipoVendita As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim strDummy As String
    Dim strDataEvento As String
    Dim strOraEvento As String
    Dim strLine As String
    Dim strTempFileName As String
    Dim intRowError As Errors
    Dim varCounter As Variant
    Dim varDensitàProdotto As Variant
    Dim varCUCarga As Variant
    Dim varCUReparto As Variant
    Dim varCUTotale As Variant
    Dim varEventi As Variant
    Dim varCoordX As Variant
    Dim varCoordY As Variant
    Dim varDistanzaDepositoDestinazione As Variant
    Dim ErrCode As Errors
    Dim C As CError
    Dim strPcName As String
    Dim strComName As String
    
    Select Case mvarFunction
        Case fcEXPSNPA, fcEXPSNPA2
            tboxLog.Text = Format(Now(), "dd/mm/yy Hh:Nn.Ss") & " === ExportRecords === " & vbCrLf
    End Select

    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    strCodOperatore = vbNullString
    
    lngNumRows = XA.UpperBound(1)
    lngNumViaggiModificati = ViaggiModificati
    lngCountOK = 0
    
    If GetComputerName(strPcName) Then
        strComName = UCase(strPcName)
    End If
    
    strTempFileName = TEMP_FILE & "_" & strComName & "_" & mskField10.Text & ".tmp"
    
    If (lngNumViaggiModificati > 0) Or (chkTransAll.value = vbChecked) Then
        If Len(mskFileName) > 0 Then
            strFileName = mskFileName
        Else
            ErrCode = FileNotFound
            GoTo ErrHandler
        End If
        
        If Len(Dir(strTempFileName)) > 0 Then
            If Len(Dir(TEMP_FILE)) > 0 Then
                Kill TEMP_FILE
            End If
        End If
        If Not mobjOutFile.OpenFile(strTempFileName, 1) Then
            ErrCode = FileNotFound
            GoTo ErrHandler
        End If
    End If
    
    strDataInvio = Format$(Date, "YYYYMMDD")
    strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
    
    If chkTransAll.value = vbChecked Then
        lngCountToGo = lngNumRows + 1
        lblLog(5).Caption = lngNumRows + 1
    Else
        lngCountToGo = lngNumViaggiModificati
    End If
    
    ' --- Ciclo di elaborazione dei viaggi
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        blnNeedTrasmission = True
        
        If (XA(lngRow, 32) = "S") Or (chkTransAll.value = vbChecked) Then
            strIDViaggio = XA(lngRow, 5)
            strIDConsegna = XA(lngRow, 1)
            strIDScarico = XA(lngRow, 2)
            If strIDViaggio <> strIDViaggioPrec Then
                strIDViaggioPrec = strIDViaggio
                
                lngProgrConsegna = 0
                lngPortataVeicoloKG = 0
                With VVT.DS
                    .Fields.ResetFields
                    .Fields("0").FindValue = strIDViaggio
                    VVT.Find
                End With
                
                strIDMotrice = VVT.DS.Field("Motrice").value
                With AAM.DS
                    .Fields.ResetFields
                    .Fields("0").FindValue = strIDMotrice
                    AAM.Find
                    If IsNumeric(.Fields("25").value) Then
                        lngPortataVeicoloKG = .Fields("25").value
                    End If
                End With
                If Len(VVT.DS.Fields("16").value) > 0 Then
                    With AAM.DS
                        .Fields.ResetFields
                        .Fields("0").FindValue = VVT.DS.Fields("16").value
                        AAM.Find
                        If IsNumeric(.Fields("25").value) Then
                            lngPortataVeicoloKG = lngPortataVeicoloKG + .Fields("25").value
                        End If
                    End With
                End If
                If Not VVT.ComputeCU(strIDViaggio, lngPortataVeicoloKG, varCUCarga, varCUReparto, varCUTotale) Then
                    ErrCode = FasciaTariffariaError
                    GoTo LogResult
                End If
            End If
            lngProgrConsegna = lngProgrConsegna + 1
            
            With AEE.DS
                .Fields.ResetFields
                .Fields(.KeyFields(1)).FindValue = XA(lngRow, 7)
                If Not AEE.Find Then
                    ErrCode = NoEnteEmittente
                    GoTo LogResult
                End If
                If .HowMany = 0 Then
                    ErrCode = NoEnteEmittente
                    GoTo LogResult
                End If
                strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(XA(lngRow, 33))
            End With
    
            With ADE.DS
                .Fields.ResetFields
                .Fields("0").FindValue = XA(lngRow, 4)
                If (Not ADE.Find) Or (.HowMany = 0) Then
                    intRowError = NoDeposito
                    GoTo LogResult
                End If
                strIDDeposito = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                strIDSocietàDeposito = .Field("Società").value
            End With
            
            With ADZ.DS
                .Fields.ResetFields
                .Fields("0").FindValue = XA(lngRow, 20)
                If (Not ADZ.Find) Or (.HowMany = 0) Then
                    intRowError = NoDestinazione
                    GoTo LogResult
                End If
                
                Select Case UCase(.Field("Tipo Destinazione").value)
                    Case "EA", "O"
                        blnNeedTrasmission = False
                        GoTo LogResult
                End Select
                
                strDestinazione = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                varDistanzaDepositoDestinazione = ADZ.GetDistanza(XA(lngRow, 20), XA(lngRow, 4))
                
                lngKMCepsa = 0
                If Not IsEmpty(varDistanzaDepositoDestinazione) Then
                    lngKMCepsa = varDistanzaDepositoDestinazione(0, 0)
                End If
            End With
            
            With AUT.DS
                .Fields.ResetFields
                .Fields("0").FindValue = XA(lngRow, 11)
                If (Not AUT.Find) Or (.HowMany = 0) Then
                    intRowError = noUT
                    GoTo LogResult
                End If
                varBadge = .MultiSoc.DS.GetFirmCode(strIDSocietàDeposito)
                If IsNull(varBadge) Or Len(varBadge) = 0 Then
                    intRowError = InvalidCode
                    GoTo LogResult
                End If
            End With
        
            With APR.DS
                .Fields.ResetFields
                .Fields("0").FindValue = XA(lngRow, 28)
                If (Not APR.Find) Or (.HowMany = 0) Then
                    intRowError = NoProdotto
                    GoTo LogResult
                End If
                strProdotto = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH, 3)
            End With
            
            With AAU.DS
                .Fields.ResetFields
                .Fields("0").FindValue = XA(lngRow, 14)
                If (Not AAU.Find) Or (.HowMany = 0) Then
                    intRowError = NoAutista
                    GoTo LogResult
                End If
                strCartaIdentitaAutista1 = Replace(Replace(.Fields("9").value, " ", ""), "-", "")
            End With
            
            strDataViaggio = Format(XA(lngRow, 3), "YYYYMMDD")
            
            '-- DP 15/05/08
            '-- Utilizzo il nuovo campo VVCHRARR al posto di VVCTEMVG
            strOraConsegna = Format(XA(lngRow, 64), "HH:MM")
            strOraConsegna = Left(strOraConsegna, 2) & ":" & Right(strOraConsegna, 2)
            
            strOraPrevista = Format(XA(lngRow, 22), "HH:MM")
            strOraPrevista = Left(strOraPrevista, 2) & ":" & Right(strOraPrevista, 2)
            strOraInizioCarico = Format(XA(lngRow, 10), "HH:MM")
            strOraInizioCarico = Left(strOraInizioCarico, 2) & ":" & Right(strOraInizioCarico, 2)
            strOraFineCarico = Format(XA(lngRow, 9), "HH:MM")
            strOraFineCarico = Left(strOraFineCarico, 2) & ":" & Right(strOraFineCarico, 2)
            strConsegnaIniziata = IIf(IsNull(XA(lngRow, 25)), vbNullString, XA(lngRow, 25))
            blnConsegnaIniziata = (strConsegnaIniziata = "S")
            If IsNumeric(XA(lngRow, 24)) And (XA(lngRow, 24) <> 0) Then
                lngKm = XA(lngRow, 24)
            ElseIf IsNumeric(XA(lngRow, 62)) And (XA(lngRow, 62) <> 0) Then
                lngKm = XA(lngRow, 62)
            Else
                lngKm = 0
            End If
            
            ' Se ci troviamo sull'ultima riga di un viaggio aggiungiamo ai km i km di rientro
            blnUltimaConsegna = False
            blnPrimaConsegna = False
            For k = lngRow To lngNumRows
                If (strIDConsegna <> XA(k, 1)) Or (strIDViaggio <> XA(k, 5)) Then
                    If strIDViaggio <> XA(k, 5) Then
                        blnUltimaConsegna = True
                    End If
                    
                    If k > 0 Then
                        If strIDViaggio <> XA(k - 1, 5) Then
                            blnPrimaConsegna = True
                        End If
                    Else
                        blnPrimaConsegna = True
                    End If
                    
                    Exit For
                End If
            Next k
            
            If k > lngNumRows Then
                blnUltimaConsegna = True
            End If
            
            If blnUltimaConsegna Then
                ' -- DP 10/07/2009
                ' -- Se li ho, sommo ai "Km parziali" dell'ultima consegna i "Km di rientro" del viaggio (se ce li ho quelli reali: VVTKMRIR, altrimenti quelli Mover VVTKMRIE)
                If IsNumeric(XA(lngRow, 78)) Then
                    lngKm = lngKm + XA(lngRow, 78)
                Else
                    If IsNumeric(XA(lngRow, 63)) Then
                        lngKm = lngKm + XA(lngRow, 63)
                    End If
                End If
            ElseIf blnPrimaConsegna Then
                ' -- DP 09/07/2009
                ' -- Se li ho, sommo ai "Km parziali" della prima consegna i "Km a vuoto" del viaggio
                If IsNumeric(XA(lngRow, 77)) Then
                    lngKm = lngKm + XA(lngRow, 77)
                End If
            End If

            With mobjOutFile
                .ClearFields
                .Item("1").value = strIDEnteEmittente
                .Item("2").value = strProgrInvio
                .Item("3").value = strDataInvio
                .Item("4").value = strOraInvio
                .Item("6").value = IIf(IsNull(XA(lngRow, 19)), vbNullString, XA(lngRow, 19))    'Riferimento esterno 1
                If Not IsNull(XA(lngRow, 34)) Then
                    .Item("7").value = XA(lngRow, 34)                                           'Partizione
                End If
                .Item("8").value = Mid(strDestinazione, 3)
                .Item("12").value = strProdotto
                .Item("13").value = strIDDeposito
                .Item("14").value = IIf(IsNull(XA(lngRow, 29)), 0, XA(lngRow, 29))              ' Q.tà concordata ordine
                If Len(XA(lngRow, 46)) > 0 Then
                    .Item("15").value = XA(lngRow, 46)
                End If
                .Item("16").value = strDataViaggio
                .Item("19").value = IIf(IsNull(XA(lngRow, 30)), 0, XA(lngRow, 30))              ' Q.tà caricata
                .Item("20").value = IIf(IsNull(XA(lngRow, 31)), 0, XA(lngRow, 31))              ' Q.tà scaricata
                .Item("21").value = strDataViaggio
                If XA(lngRow, 37) = "S" Then
                    .Item("22").value = strOraFineCarico
                End If
                
                '-- DP 21/05/08
                '-- Su indicazioni ricevute da L. Iacopozzi, nei campi data e ora consegna "aggiornata" devo scrivere solamente se c'è stato un cambio rispetto all'ora "programmata"
                If CDate(XA(lngRow, 65)) <> CDate("01/01/1900") Then
                    .Item("17").value = Format(XA(lngRow, 65), "yyyymmdd")                      ' Data consegna aggiornata (in questo caso ci scrivo VVCDTPRE)
                    .Item("18").value = strOraPrevista                                          ' Ora consegna aggiornata (ci scrivo il contenuto di VVCHRPRE, potrebbe essere a NULL)
                Else
                    '-- DP 27/05/08
                    '-- R. Chaparro mi indica che la "Data consegna aggiornata" è un campo obbligatorio
                    .Item("17").value = Format(XA(lngRow, 36), "yyyymmdd")                      ' Data consegna aggiornata (in questo caso ci scrivo VVCDTCON)
                    .Item("18").value = vbNullString                                            ' Ora consegna aggiornata
                End If
                
                '-- DP 19/05/08
                '-- Se la consegna non è iniziata lascio in bianco le ore effettive di consegna
                If XA(lngRow, 25) = "S" Then
                    If blnConsegnaIniziata Then
                        .Item("23").value = Format(XA(lngRow, 66), "yyyymmdd")                  ' Data consegna effettiva
                        .Item("24").value = strOraConsegna                                      ' Ora consegna effettiva
                    Else
                        .Item("23").value = vbNullString                                        ' Data consegna effettiva
                        .Item("24").value = vbNullString                                        ' Ora consegna effettiva
                    End If
                Else
                    .Item("23").value = vbNullString                                            ' Data consegna effettiva
                    .Item("24").value = vbNullString                                            ' Ora consegna effettiva
                End If
                
                .Item("25").value = varBadge
                
                
                '-- DP 14/08/08
                If Not IsNull(XA(lngRow, 72)) Then
                    .Item("26").value = Right(XA(lngRow, 72), 2)                                ' Nr. sequenza viaggio
                Else
                    .Item("26").value = IIf(IsNull(XA(lngRow, 8)), vbNullString, XA(lngRow, 8)) ' Nr. sequenza viaggio
                End If
                
                
                If IsNumeric(Left(strCartaIdentitaAutista1, 2)) Then
                    .Item("27").value = "ES" & strCartaIdentitaAutista1
                Else
                    .Item("27").value = strCartaIdentitaAutista1
                End If
                .Item("28").value = lngKm
                .Item("29").value = lngKMCepsa
                .Item("30").value = Replace(Format(varCUCarga, "000.00"), ",", ".")
                .Item("31").value = Replace(Format(varCUReparto, "000.00"), ",", ".")
                .Item("32").value = Replace(Format(varCUTotale, "000.00"), ",", ".")
                .Item("33").value = lngProgrConsegna                                            ' Nr. sequenza consegna
                .Item("34").value = IIf(IsNull(XA(lngRow, 26)), vbNullString, XA(lngRow, 26))   ' Note
                .Item("35").value = IIf(IsNull(XA(lngRow, 27)), vbNullString, XA(lngRow, 27))   ' Note estese
                
                VVC.DS.Fields.ResetFields
                VVC.DS.Field("Numero Viaggio").FindValue = strIDViaggio
                VVC.DS.Field("Numero Consegna").FindValue = strIDConsegna
                VVC.Find
                
                If XA.value(lngRow, 59) = "S" Then
                    .Item("35").value = "013"                               ' Salida de destino
                    '-- DP 21/05/08
                    '-- Utilizzo il nuovo campo VVCDTARR al posto di VVCDTCON
                    strDataEvento = XA(lngRow, 66)
                    strOraEvento = XA.value(lngRow, 55)
                ElseIf blnConsegnaIniziata Then 'XA.value(lngRow, 25) = "S" Then
                    .Item("35").value = "005"                               ' Inicio descarga JG 30/09/09 CAMBIO 009 a 005
                    '-- DP 21/05/08
                    '-- Utilizzo il nuovo campo VVCDTARR al posto di VVCDTCON
                    strDataEvento = XA(lngRow, 66)
                    
                    '-- DP 15/05/08
                    '-- Utilizzo il nuovo campo VVCHRARR al posto di VVCTEMVG
                    strOraEvento = strOraConsegna
                ElseIf XA(lngRow, 58) = "S" Then
                    .Item("35").value = "000"                               ' Actualizacion fecha prevista entrega
                    strDataEvento = Date
                    strOraEvento = Time
                    VVC.DS.Field("Flag stato 4").NewValue = "N"
                    VVC.DS.Field("Flag stato 4").Updated = True
                    VVC.BeginSave blnConcErr
                    VVC.DS.DA.EndTrans ETCOMMIT
                ElseIf XA.value(lngRow, 37) = "S" Then
                    .Item("35").value = "003"                               ' Salida del cargadero --JG  30/09/09 Cambio 006 A 003
                    strDataEvento = VVT.DS.Field("Data inizio viaggio").value
                    strOraEvento = VVT.DS.Field("Ora inizio viaggio").value
                ElseIf XA.value(lngRow, 57) = "S" Then
                    .Item("35").value = "002"                               ' Entrada en el cargadero
                    strDataEvento = VVT.DS.Field("Data inizio carico").value
                    strOraEvento = VVT.DS.Field("Ora inizio carico").value
                ElseIf XA.value(lngRow, 61) = "S" Then
                    .Item("35").value = "008"                               ' Desprecinctado
                    strDataEvento = Date
                    strOraEvento = Time
                Else
                    strDataEvento = Date
                    strOraEvento = Time
                End If
                                    
                If Len(.Item("35").value) > 0 Then
                    .Item("36").value = Format(strDataEvento, "yyyymmdd")
                    .Item("37").value = Left(Format(strOraEvento, "hh:mm"), 2) & ":" & Right(Format(strOraEvento, "hh:mm"), 2)
                End If
                                    
'                SEV.SAT_Get_Last_Pos strIDMotrice, strDataEvento, strOraEvento, varCoordX, varCoordY
'
'                If (Not IsEmpty(varCoordX)) And (Not IsEmpty(varCoordY)) Then
'                    .Item("10").value = Replace(Format(varCoordX, "00.00000"), ".", ",")
'                    .Item("11").value = Replace(Format(varCoordY, "00.000000"), ".", ",")
'                End If

                ' -- DP 10/07/2009
                ' -- Scrivo le coordinate fisse dall'anagrafica
                Select Case .Item("35").value
                    Case "003", "002" '--JG  30/09/09 Cambio 006 A 003
                        If Not IsEmpty(XA(lngRow, 81)) And Not IsEmpty(XA(lngRow, 82)) Then
                            .Item("10").value = Replace(Format(XA(lngRow, 81), "00.00000"), ".", ",")
                            .Item("11").value = Replace(Format(XA(lngRow, 82), "00.000000"), ".", ",")
                        
                        End If
                    
                    Case "005"
                    
                        If (Not IsEmpty(XA(lngRow, 79)) And Not IsEmpty(XA(lngRow, 80))) And (XA(lngRow, 79) <> "" And XA(lngRow, 80) <> "") And (Not IsNull(XA(lngRow, 79)) And Not IsNull(XA(lngRow, 80))) Then
                            .Item("10").value = Replace(Format(XA(lngRow, 79), "00.000000"), ".", ",")
                            .Item("11").value = Replace(Format(XA(lngRow, 80), "00.000000"), ".", ",")
 '                          JG SM Si el cliente no tiene coordenadas las tomamos de la ultima posición del vehículo.
                        Else
                        
                            'SEV.SAT_Get_Last_Pos strIDMotrice, strDataEvento, strOraEvento, varCoordX, varCoordY
                            
                            Dim rowbuff, Varbuff As Variant
                            Dim lnghm, lnghow As Long
                            Dim idvehi As String
                            Dim DA As CDataAccess
                            Set DA = DataConnect.DA
                            
                            'SM Cambio de la tabla donde sacar la ultima posición.
                            idvehi = AAM.FindIDVehiOes(strIDMotrice)
                            If Not DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud From PosicionesOesia WHERE (IDVEHIC = ?)", Varbuff, 1, idvehi) Then
                            End If
                            varCoordX = Varbuff(1, 0)
                            varCoordY = Varbuff(2, 0)
                            
                            .Item("10").value = Replace(Format(varCoordX, "00.00000"), ".", ",")
                            .Item("11").value = Replace(Format(varCoordY, "00.000000"), ".", ",")

                             With ADZ
                                .DS.Fields("0").FindValue = strDestinazione

                                If Not .Find Then
                                   GoTo ErrHandler
                                End If

                                .DS.Fields("31").NewValue = CVar(Replace(Format(varCoordX, "00.00000"), ".", ","))
                                .DS.Fields("31").Updated = True
                                .DS.Fields("32").NewValue = CVar(Replace(Format(varCoordY, "00.000000"), ".", ","))
                                .DS.Fields("32").Updated = True
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                     GoTo ErrHandler
                                End If

                                .DS.DA.EndTrans ETCOMMIT
                            End With
                        
                        End If
                    Case Else
                End Select
                
                If XA.value(lngRow, 60) = "S" Then
                    .Item("46").value = "1"
                End If
                
                If Not IsNumeric(XA(lngRow, 76)) Then
                    XA(lngRow, 76) = "0"
                End If
                .Item("47").value = XA(lngRow, 76)
                .Item("48").value = strOraInizioCarico
                    
                If Not .WriteLine(True) Then
                    intRowError = OutputError
                    GoTo LogResult
                End If
            
                '-- DP 24/07/08
                '-- Segnalo l'esportazione su GES dell'ora di scarico reale
  
                If .Item("35").value = "005" Then '--JG  30/09/09 Cambio 009 A 005
                    If Not blnSetDeliveryTransmitted(XA(lngRow, 0), XA(lngRow, 1), XA(lngRow, 2)) Then
                        intRowError = UpdateError
                        GoTo LogResult
                    End If
                End If
            End With
            
            Select Case mvarFunction
                Case fcEXPSNPA, fcEXPSNPA2
                    Call UpdModificato(XA(lngRow, 0), XA(lngRow, 1), XA(lngRow, 2), False, False, True, "")
                    
            End Select
        Else
            blnNeedTrasmission = False
        End If

LogResult:
        ' --- Scrivo il risultato dell'elaborazione
        If blnNeedTrasmission Then
            If intRowError = NoError Then
                lngCountOK = lngCountOK + 1
                
                '-- DP 09/09/08
                '-- Aggiorno il campo "Sequenza Invio Trasmessa"
                If Not VVT.blnUpdateSEQOK("S", CVar(XA(lngRow, 0))) Then
                    GoTo ErrHandler
                End If
            Else
                lngCountFail = lngCountFail + 1
            End If
            lngCountToGo = lngCountToGo - 1
            lblLog(2) = lngCountOK
            lblLog(4) = lngCountFail
            lblLog(5) = lngCountToGo
            Select Case intRowError
                Case NoError
                    strErrorMessage = "OK"
                Case NoAutista
                    strErrorMessage = "Primo autista non trovato"
                Case OutputError
                    strErrorMessage = "Errore durante la scrittura sul file"
                Case noUT
                    strErrorMessage = "Unità di trasporto non trovata"
                Case NoDestinazione
                    strErrorMessage = "Destinazione non trovata"
                Case RowsNotFound
                    strErrorMessage = "Errore durante la lettura delle consegne"
                Case NoSetState
                    strErrorMessage = "Impossibile cambiare lo stato del viaggio"
                Case NoOrderID
                    strErrorMessage = "Ordine non trovato"
                Case NoTipoVendita
                    strErrorMessage = "Codice Tipo Vendita non trovato"
                Case NoDestinazione
                    strErrorMessage = "Codice Destinazione non trovato"
                Case NoDeposito
                    strErrorMessage = "Codice Deposito non trovato"
                Case NoEnteEmittente
                    strErrorMessage = "Ente Emittente non trovato"
                Case InvalidCode
                    strErrorMessage = "Badge non trovato"
                Case UpdateError
                    strErrorMessage = "Errore durante l'aggiornamento dello stato di trasmissione"
            End Select
        Else
            strErrorMessage = "Record da non trasmettere"
        End If
        XALog.value(lngRow, 3) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
        DoEvents
    Next lngRow
    
    GoTo EndProc

ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & strFileName
        Case ParameterNotFound
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti: Prodotti"
        Case NoEnteEmittente
            strMessage = "Ente Emittente non trovato."
            strInfo = "Ente emittente: " & strIDEnteEmittente
        Case FasciaTariffariaError
            strMessage = "Errore nel calcolo dei coefficienti di utilizzo"
        Case InvalidCode
            strMessage = "Impossibile leggere il progressivo dell'invio"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    DataClass.DS.DA.EndTrans ETROLLBACK
    Set C = ErrH.SetError(ErrCode, "frmVVTSnapshotExport", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
                
    Select Case mvarFunction
        Case fcEXPSNPA, fcEXPSNPA2
            tboxLog.Text = " Codice: " & C.Code & vbCrLf & _
                    " Descrizione principale: " & C.MainDescription & vbCrLf & _
                    " Info: " & C.info & vbCrLf

        Case Else
            DisplayError ErrH.LastError
    End Select
    lngCountOK = 0
    
EndProc:





    Dim strNumeratore As String
    mobjOutFile.CloseFile
    
    Select Case mvarFunction
        Case fcEXPSNP2
            strNumeratore = "015"
        Case Else
            Select Case mskField10
                Case "624"
                    strNumeratore = "007"
                Case "004"
                    strNumeratore = "107"
                Case Else
                    lngCountOK = 0
            End Select
            
    End Select
    
    If lngCountOK > 0 Then
        With DataClass.DS.DA
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", strNumeratore, 0)
            If varCounter = -1 Then
                ErrCode = InvalidCode
                GoTo ErrHandler
            End If
            If varCounter > 9999999999# Then
                varCounter = 1
            End If
        End With
        
        strProgrInvio = Format(varCounter, String(10, "0"))
        strFileName = strFileName & "\GES_N" & strProgrInvio & ".DAT"
        intFileFromNum = FreeFile
        Open strTempFileName For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Mid(strLine, 31, 4) = Format(lngCountOK, "0000")
            Mid(strLine, 5, 10) = strProgrInvio
            Print #intFileToNum, strLine
        Loop
        Close #intFileFromNum
        Close #intFileToNum
    
        Select Case mvarFunction
            Case fcEXPSNPA, fcEXPSNPA2
                tboxLog.Text = tboxLog.Text & " File: " & strFileName & vbCrLf & " Records inviati: " & lngCountOK & vbCrLf & _
                    " Records in errore: " & lngCountFail & vbCrLf
            Case Else
                chkTransAll.value = vbUnchecked
        End Select
    Else
        Select Case mvarFunction
            Case fcEXPSNPA, fcEXPSNPA2
                tboxLog.Text = tboxLog.Text & " Nessun record inviato." & vbCrLf
            Case Else
                chkTransAll.value = vbUnchecked
        End Select
    End If
    '--SM -JG Detección de ejecución de proccesos
Dim ret As Boolean
Dim nombrefichero As String
Dim x As CTextFile
Set x = New CTextFile
    
    If mskField10 = "624" Then
    nombrefichero = "TPHGES.txt"
    ElseIf (mskField10 = "004") Then
    nombrefichero = "EUTGES.txt"
    End If
    
    ret = x.OpenFile("R:\Resoil2\Alarmas\" + nombrefichero, 1)
    ret = x.CloseFile()
    
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
End Sub

Private Sub TimerAutomaticSend_Timer()
    Dim i As Integer
    Dim blnTrasmAll As Boolean
    
    TimerAutomaticSend.Enabled = False

    mlngTickTimer = mlngTickTimer + 1
    lbTimer.Caption = mlngTickTimer
    If mlngTickTimer >= mlngTimerMinutes Then
        mlngTickTimer = 0
        lbTimer.Caption = mlngTickTimer
    
        MaskEdBox12 = Format(Date, "dd/mm/yy")
        MaskEdBox14 = Format(Date, "dd/mm/yy")
    
        For i = 1 To 4
            If (mskOrario(i) <> TIME_MASK) Then
                If (DateDiff("n", mskOrario(i), Time) > 0) And _
                   (Not mblnTrasmAll(i)) Then
                    mblnTrasmAll(i) = True
                    Me.chkTransAll = vbChecked
                    blnTrasmAll = True
               End If
            End If
        Next i
        
        DoCommand fcFind
        DoCommand fcExport
        
        If blnTrasmAll Then
            Me.chkTransAll = False
            For i = 4 To 1 Step -1
                If mskOrario(i) <> TIME_MASK Then
                    Exit For
                End If
            Next i
            If i > 0 Then
                If mblnTrasmAll(i) Then
                    For i = 1 To 4
                        mblnTrasmAll(i) = False
                    Next i
                End If
            End If
        End If
    End If

    TimerAutomaticSend.Enabled = True
End Sub

Private Sub ExportEventiAgip()
    On Error GoTo ErrHandler

    Const TEMP_FILE = "INC.tmp"

    Dim strFileName As String
    Dim lngCountOK As Long
    Dim lngRecordCount As Long
    Dim lngErrorCount As Long
    Dim lngNumRows As Long
    Dim lngCounter As Long
    Dim lngNumRowsEventi As Long
    
    Dim ErrCode As Errors
    Dim varBufferEventi As Variant
    Dim lngRowsCounter As Long
    Dim lngEventsCounter As Long
    Dim strIDViaggio As String
    Dim strIDViaggioPrec As String

    Dim strIDEnteEmittente As String
    Dim strIDProdotto As String
    Dim strIDDestinazione As String
    Dim strIDUnitaTrasporto As String
    Dim strIDDeposito As String
    Dim lngSequenzaConsegna As Long
    Dim lngSequenzaProdotto As Long
    Dim strIDVVC As String
    Dim strIDVVCPrec As String
    Dim strIDVVP As String
    Dim strIDVVPPrec As String
    Dim strProgrInvio As String
    Dim strDataInvio As String
    Dim strOraInvio As String
    Dim strHeader As String
    Dim intRowError As Errors
    Dim strErrorMessage As String
    Dim blnTrasmesso As Boolean
    Dim C As CError
    Dim varCounter As Variant
    Dim strCounterProgInvio As String
    Dim strIDMotivoEvento As String
    Dim strIDAsc As String
    Dim strIDTipoVendita As String

    strIDAsc = ID_SOC_AGIP
    strFileName = mskFileName

    If Not mobjOutFileInc.OpenFile(TEMP_FILE, 1) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If

    With AXX.DS
        .Fields(.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find() Or .HowMany = 0 Then
            ErrCode = ParameterNotFound
            GoTo ErrHandler
        End If
    End With

    With AEE.DS
        .Fields.ResetFields
        .Fields(.KeyFields(1)).FindValue = AXX.DS.Field("Ente Emittente").value
        If Not AEE.Find Or .HowMany = 0 Then
            ErrCode = NoEnteEmittente
            GoTo ErrHandler
        End If
        strIDEnteEmittente = .MultiSoc.DS.GetFirmCode(strIDAsc)
        If Len(strIDEnteEmittente) = 0 Then
            ErrCode = NoEnteEmittente
            GoTo ErrHandler
        End If
    End With
    
    lngCountOK = 0
    lngErrorCount = 0
    lngCounter = 0
    mobjOutFileInc.Delimiter = "|"

    ' Scrittura del file di dettaglio
    ' ------------------------------
    lngNumRows = XA.UpperBound(1)
    strIDViaggio = ""
    
    lngRecordCount = 0
    lngCountOK = 0
    lngErrorCount = 0
    
    lblLog(5).Caption = lngRecordCount
    lblLog(2).Caption = lngCountOK
    lblLog(4).Caption = lngErrorCount
    
    For lngRowsCounter = 0 To lngNumRows

        intRowError = NoError
        strIDViaggio = XA(lngRowsCounter, 0)
        
        If strIDViaggio <> strIDViaggioPrec Then
            strIDViaggioPrec = strIDViaggio
        
            With VPI.DS
                .Fields.ResetFields
                .Field("Numero Viaggio").FindValue = XA(lngRowsCounter, 0)
            End With

            If VPI.GetSelectionExportAgip(True) Then
                varBufferEventi = VPI.DS.List
            End If
            
            lngSequenzaConsegna = 0
            lngSequenzaProdotto = 0
            strIDVVCPrec = ""
            strIDVVPPrec = ""
            If Not IsEmpty(varBufferEventi) Then
                
                lngNumRowsEventi = UBound(varBufferEventi, 2)

                For lngEventsCounter = 0 To lngNumRowsEventi

                    intRowError = NoError

                    If Len(varBufferEventi(19, lngEventsCounter)) = 0 Or _
                        varBufferEventi(19, lngEventsCounter) = "N" Then
                        blnTrasmesso = False
                    Else
                        blnTrasmesso = True
                    End If
                    
                    If chkTransAll.value = 0 And blnTrasmesso Then
                        GoTo LogRecord
                    Else
                        blnTrasmesso = False
                    End If
                    
                    strIDVVC = varBufferEventi(15, lngEventsCounter)
                    If strIDVVC <> strIDVVCPrec Then
                        strIDVVCPrec = strIDVVC
                        lngSequenzaConsegna = lngSequenzaConsegna + 1
                        lngSequenzaProdotto = 0
                        strIDVVPPrec = ""
                    End If
                    
                    strIDVVP = varBufferEventi(16, lngEventsCounter)
                    If strIDVVP <> strIDVVPPrec Then
                        strIDVVPPrec = strIDVVP
                        lngSequenzaProdotto = lngSequenzaProdotto + 1
                    End If

                    With mobjOutFileInc
                        ' Motivo Evento
                        With AME.DS
                            .Fields.ResetFields
                            .Field("Codice").FindValue = varBufferEventi(0, lngEventsCounter)
                            If Not AME.Find() Or .HowMany = 0 Then
                                intRowError = NoMotivoEvento
                                GoTo LogRecord
                            End If
                            strIDMotivoEvento = .MultiSoc.DS.GetFirmCode(strIDAsc)
                            If Len(strIDMotivoEvento) = 0 Then
                                intRowError = NoMotivoEvento
                                GoTo LogRecord
                            End If
                        End With
                        .Item("1").value = strIDMotivoEvento

                        ' Osservazioni
                        If Len(varBufferEventi(1, lngEventsCounter)) > 60 Then
                            .Item("2").value = Left(varBufferEventi(1, lngEventsCounter), 60)
                        Else
                            .Item("2").value = varBufferEventi(1, lngEventsCounter)
                        End If

                        ' Data Evento
                        .Item("3").value = Format(varBufferEventi(2, lngEventsCounter), "DD/MM/YY")

                        ' Unità di Trasporto
                        With AUT.DS
                            .Fields.ResetFields
                            .Fields("0").FindValue = varBufferEventi(3, lngEventsCounter)
                            If (Not AUT.Find) Or (.HowMany = 0) Then
                                intRowError = noUT
                                GoTo LogRecord
                            End If
                            strIDUnitaTrasporto = .MultiSoc.DS.GetFirmCode(strIDAsc)
                            If Len(strIDUnitaTrasporto) = 0 Then
                                ErrCode = noUT
                                GoTo ErrHandler
                            End If
                        End With
                        .Item("4").value = strIDUnitaTrasporto
                        
                        ' Badge
                        .Item("5").value = varBufferEventi(4, lngEventsCounter)
                        
                        ' Sequenza Viaggio
                        .Item("6").value = CLng(varBufferEventi(5, lngEventsCounter))

                        ' Sequenza Consegna
                        .Item("7").value = lngSequenzaConsegna
                        
                        ' Sequenza Prodotto
                        .Item("8").value = lngSequenzaProdotto

                        ' Zona
                        With AZO.DS
                            .Fields.ResetFields
                            .Fields("Codice").FindValue = varBufferEventi(6, lngEventsCounter)
                            If Not AZO.Find Or .HowMany = 0 Then
                                intRowError = NoZona
                                GoTo LogRecord
                            End If
                        End With
                        .Item("9").value = varBufferEventi(6, lngEventsCounter)
                        
                        ' Deposito
                        With ADE.DS
                            .Fields.ResetFields
                            .Fields("0").FindValue = varBufferEventi(7, lngEventsCounter)
                            If (Not ADE.Find) Or (.HowMany = 0) Then
                                intRowError = NoDeposito
                                GoTo LogRecord
                            End If
                            strIDDeposito = .MultiSoc.DS.GetFirmCode(strIDAsc)
                            If Len(strIDDeposito) = 0 Then
                                ErrCode = NoDeposito
                                GoTo ErrHandler
                            End If
                        End With
                        .Item("10").value = strIDDeposito

                        ' Raggruppamento deposito
                        With ARP.DS
                            .Fields(.KeyFields(1)).FindValue = varBufferEventi(8, lngEventsCounter)
                            If Not ARP.Find Or .HowMany = 0 Then
                                intRowError = NoRagDeposito
                                GoTo LogRecord
                            End If
                        End With
                        .Item("11").value = varBufferEventi(8, lngEventsCounter)

                        ' Numero de pedido
                        .Item("12").value = varBufferEventi(9, lngEventsCounter)
                        
                        ' Prodotto
                        With APR.DS
                            .Fields.ResetFields
                            .Fields("0").FindValue = varBufferEventi(10, lngEventsCounter)
                            If (Not APR.Find) Or (.HowMany = 0) Then
                                intRowError = NoProdotto
                                GoTo LogRecord
                            End If
                            
                            With .MultiSoc.DS
                                .Fields.ResetFields
                                .Field("Codice").FindValue = varBufferEventi(10, lngEventsCounter)
                                .Field("Imballo").FindValue = varBufferEventi(20, lngEventsCounter)
                                .Field("Posizione Doganale").FindValue = varBufferEventi(21, lngEventsCounter)
                                .Field("Società").FindValue = strIDAsc
                                If Not APR.DS.MultiSoc.Find Or .HowMany = 0 Then
                                    intRowError = NoProdotto
                                    GoTo LogRecord
                                End If
                            End With
                        End With
                        .Item("13").value = APR.DS.MultiSoc.DS.Field("CodPerSoc").value
                        
                        ' Destinazione
                        With ADZ.DS
                            .Fields.ResetFields
                            .Fields("0").FindValue = varBufferEventi(11, lngEventsCounter)
                            If (Not ADZ.Find) Or (.HowMany = 0) Then
                                intRowError = NoDestinazione
                                GoTo LogRecord
                            End If
                            strIDDestinazione = .MultiSoc.DS.GetFirmCode(strIDAsc)
                            If Len(strIDDestinazione) = 0 Then
                                ErrCode = NoDestinazione
                                GoTo ErrHandler
                            End If
                        End With
                        .Item("14").value = strIDDestinazione
                        
                        ' Nome Destinazione
                        .Item("15").value = varBufferEventi(12, lngEventsCounter)
                        
                        ' Tipo de venta
                        With ATV.DS
                            .Fields.ResetFields
                            .Fields(.KeyFields(1)).FindValue = varBufferEventi(13, lngEventsCounter)
                            If Not ATV.Find Then
                                intRowError = NoTipoVendita
                                GoTo LogRecord
                            End If
                            strIDTipoVendita = .MultiSoc.DS.GetFirmCode(strIDAsc)
                            If Len(strIDTipoVendita) = 0 Then
                                ErrCode = NoTipoVendita
                                GoTo ErrHandler
                            End If
                        End With
                        .Item("16").value = varBufferEventi(13, lngEventsCounter)
                        
                        ' Momento del servizio
                        .Item("17").value = "M"

                        If Not .WriteLine(True) Then
                            intRowError = OutputError
                            GoTo ErrHandler
                        End If
                    End With

                    If Len(varBufferEventi(19, lngEventsCounter)) = 0 Or _
                        varBufferEventi(19, lngEventsCounter) = "N" Then
                        With VPI.DS
                            .Fields.ResetFields
                            .Fields("0").FindValue = varBufferEventi(14, lngEventsCounter)
                            .Fields("1").FindValue = varBufferEventi(15, lngEventsCounter)
                            .Fields("2").FindValue = varBufferEventi(16, lngEventsCounter)
                            .Fields("3").FindValue = varBufferEventi(17, lngEventsCounter)
                        End With

                        VPI.Find
                        If Not VPI.SetFlagTrasmissione("S") Then
                            intRowError = UpdateError
                            GoTo ErrHandler
                        End If
                    End If

LogRecord:
                    lngRecordCount = lngRecordCount + 1
                    
                    If intRowError = NoError Then
                        If Not blnTrasmesso Then
                            lngCountOK = lngCountOK + 1
                        End If
                    Else
                        lngErrorCount = lngErrorCount + 1
                    End If
                    
                    Select Case intRowError
                        Case NoError
                            If blnTrasmesso Then
                                strErrorMessage = "OK, già trasmesso."
                            Else
                                strErrorMessage = "OK"
                            End If
                        Case NoTipoVendita
                            strErrorMessage = "Codice Tipo Vendita non trovato"
                        Case NoMotivoEvento
                            strErrorMessage = "Motivo Evento non trovato"
                        Case noUT
                            strErrorMessage = "Unità di trasporto non trovata"
                        Case NoDeposito
                            strErrorMessage = "Codice Deposito non trovato"
                        Case NoZona
                            strErrorMessage = "Codice Zona non trovato"
                        Case NoProdotto
                            strErrorMessage = "Codice Prodotto non trovato"
                        Case NoRagDeposito
                            strErrorMessage = "Codice Raggruppamento Deposito non trovato"
                        Case NoDestinazione
                            strErrorMessage = "Destinazione non trovata"
                        Case Else
                            strErrorMessage = "Errore sconosciuto"
                    End Select
                    
                    XALog.ReDim 0, lngRecordCount - 1, 0, 3
                    XALog.value(lngRecordCount - 1, 0) = varBufferEventi(14, lngEventsCounter)
                    XALog.value(lngRecordCount - 1, 1) = varBufferEventi(15, lngEventsCounter)
                    XALog.value(lngRecordCount - 1, 2) = varBufferEventi(16, lngEventsCounter)
                    XALog.value(lngRecordCount - 1, 3) = strErrorMessage
                    With tdbLog
                        .Refresh
                        .bookmark = lngRecordCount - 1
                    End With
                    TDBGrid1.bookmark = lngNumRows
                    DoEvents
                    
                Next lngEventsCounter
            End If
        
        End If

    Next lngRowsCounter
    
    GoTo EndProc

ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & TEMP_FILE
        Case ParameterNotFound
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti: Prodotti"
        Case NoEnteEmittente
            strMessage = "Ente Emittente non trovato."
            strInfo = "Ente emittente: " & strIDEnteEmittente
        Case CounterNotFound
            strMessage = "Counter non trovato."
            strInfo = "Counter: " & strCounterProgInvio
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    lngCountOK = 0

EndProc:


'--SM -JG Detección de ejecución de proccesos
Dim ret As Boolean
Dim nombrefichero As String
Dim x As CTextFile
Set x = New CTextFile
    
    If (strIDEnteEmittente = "004") Then
    nombrefichero = "TPHGES.txt"
    ElseIf (strIDEnteEmittente = "624") Then
    nombrefichero = "EUTGES.txt"
    End If
    
    ret = x.OpenFile("R:\Resoil2\Alarmas\" + nombrefichero, 1)
    ret = x.CloseFile()



    Dim intFileFromNum As Integer
    Dim intFileToNum As Integer
    Dim strLine As String
    mobjOutFileInc.CloseFile
    
    lblLog(5).Caption = lngRecordCount
    lblLog(2).Caption = lngCountOK
    lblLog(4).Caption = lngErrorCount
    
    If lngCountOK > 0 Then
        With DataClass.DS.DA
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "010", 0)
            If varCounter = -1 Then
                ErrCode = InvalidCode
                GoTo ErrHandler
            End If
            If varCounter > 9999999999# Then
                varCounter = 1
            End If
        End With
    
        strDataInvio = Format$(Date, "dd/mm/yy")
        strOraInvio = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00")
        strProgrInvio = Format(varCounter Mod 999, "000")

        strHeader = "1|" & strIDEnteEmittente & "|" & _
            strProgrInvio & "|" & strDataInvio & "|" & _
            strOraInvio & "|"

        strFileName = strFileName & "\INC_" & strProgrInvio & ".DAT"
        intFileFromNum = FreeFile
        Open TEMP_FILE For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Print #intFileToNum, strHeader
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Print #intFileToNum, strLine
        Loop
        Close #intFileFromNum
        Close #intFileToNum
    End If

End Sub

Private Function VerificaAnomalie() As Boolean
On Error GoTo ErrHandler
    
    Dim varBufferEventi As Variant
    Dim VVP As CViaggioProdotto
    Dim OCT As COrdine
    Dim OCR As COrdineriga
    Dim blnOrdineCopiato As Boolean
    Dim strIDOCT As String
    Dim strIDOCTNew As String
    Dim lngIDOCR As Long
    Dim strIDVVT As String
    Dim strIDVVTPrec As String
    Dim strIDVVC As String
    Dim strIDVVCPrec As String
    Dim strIDVVP As String
    Dim strIDVVPPrec As String
    Dim blnSkipVVC As Boolean
    Dim strErrorMessage As String
    Dim lngNumRows As Long
    Dim lngRowsCounter As Long
    Dim lngNumRowsEventi As Long
    Dim lngEventsCounter As Long
    Dim intRowError As Errors
    Dim sngQta As Single
    Dim lngRecordCount As Long
    Dim lngCountOK As Long
    Dim lngErrorCount As Long
    Dim ErrCode As Errors
    Dim C As CError

    ErrCode = NoError

    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    Set OCT = New COrdine
    With OCT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    Set OCR = New COrdineriga
    With OCR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    lngRecordCount = 0
    lngCountOK = 0
    lngErrorCount = 0

    XALog.ReDim 0, 0, 0, 3
    XALog.DeleteRows 0
    With tdbLog
        .Refresh
    End With
    
    lngNumRows = XA.UpperBound(1)
    For lngRowsCounter = 0 To lngNumRows

        strIDVVT = XA(lngRowsCounter, 0)
        If strIDVVT <> strIDVVTPrec Then
            strIDVVTPrec = strIDVVT

            With VPI.DS
                .Fields.ResetFields
                .Field("Numero Viaggio").FindValue = strIDVVT
            End With

            If VPI.GetSelectionExportAgip(True) Then
                varBufferEventi = VPI.DS.List
            End If

            If Not IsEmpty(varBufferEventi) Then

                lngNumRowsEventi = UBound(varBufferEventi, 2)
                strIDVVCPrec = ""
                strIDVVPPrec = ""
                blnOrdineCopiato = False
                blnSkipVVC = False
                
                For lngEventsCounter = 0 To lngNumRowsEventi

                    intRowError = NoError
                    strErrorMessage = ""

                    strIDVVC = varBufferEventi(15, lngEventsCounter)
                    If strIDVVC <> strIDVVCPrec Then
                    
                        ' Close last transaction
                        If strIDVVCPrec <> vbNullString Then
                            If blnSkipVVC Then
                                DataClass.DS.DA.EndTrans 1
                            Else
                                DataClass.DS.DA.EndTrans 0
                            End If
                        End If
                        
                        strIDVVCPrec = strIDVVC
                        strIDVVPPrec = ""
                        blnSkipVVC = False

                        DataClass.DS.DA.BeginTransaction
                        
                        strIDOCT = CStr(varBufferEventi(22, lngEventsCounter))
                        
                        ' Verifica che non sia già stato copiato l'ordine
                        If Not OCT.IsOrdineCopiato(strIDOCT, blnOrdineCopiato) Then
                            intRowError = InsertError
                            strErrorMessage = "Errore nella selezione ordine copiato."
                            GoTo LogRecord
                        End If

                        If Not blnOrdineCopiato Then
                            ' Effettua la copia dell'ordine
                            If Not OCT.CopyOrder(strIDOCT, False, strIDOCTNew) Then
                                intRowError = InsertError
                                strErrorMessage = "Errore nella copia dell'ordine."
                                GoTo LogRecord
                            End If

                            ' Imposta lo stato dell'ordine in 'Acquisito'
                            OCT.DS.Field("Numero Ordine").value = strIDOCTNew
                            If Not OCT.SetAvanzamento("A") Then
                                intRowError = UpdateError
                                strErrorMessage = "Errore nell'aggiornamento dell'avanzamento ordine."
                                GoTo LogRecord
                            End If
                        Else
                            blnSkipVVC = True
                        End If
                    End If

                    strIDVVP = varBufferEventi(16, lngEventsCounter)
                    If strIDVVP <> strIDVVPPrec Then
                        strIDVVPPrec = strIDVVP

                        If Not blnSkipVVC Then
                            lngIDOCR = CLng(varBufferEventi(23, lngEventsCounter))
                            
                            If IsNumeric(varBufferEventi(24, lngEventsCounter)) Then
                                sngQta = CDbl(varBufferEventi(24, lngEventsCounter))
                            Else
                                sngQta = 0
                            End If
    
                            ' Imposta la qta scaricata a 0 sul VVP
                            If sngQta <> 0 Then
                                If Not VVP.UpdateQtyDelivered(strIDVVT, strIDVVC, strIDVVP, 0) Then
                                    intRowError = UpdateError
                                    strErrorMessage = "Errore nell'aggiornamento della qta scaricata."
                                    GoTo LogRecord
                                End If
                            End If
                        
                            ' Copia la riga dell'ordine
                            If Not OCT.CopyOrderRow(strIDOCT, lngIDOCR, strIDOCTNew) Then
                                intRowError = InsertError
                                strErrorMessage = "Errore nella copia della riga dell'ordine."
                                GoTo LogRecord
                            End If
                        End If

                    End If

LogRecord:
                    lngRecordCount = lngRecordCount + 1
                    
                    If intRowError = NoError Then
                        lngCountOK = lngCountOK + 1
                    Else
                        blnSkipVVC = True
                        lngErrorCount = lngErrorCount + 1
                    End If
                    
                    If strErrorMessage = vbNullString Then
                        Select Case intRowError
                            Case NoError
                                If blnOrdineCopiato Then
                                    strErrorMessage = "Scarico già elaborato."
                                Else
                                    strErrorMessage = "Ok, creato ordine """ & strIDOCTNew & """ riga """ & lngIDOCR & """"
                                End If
                            Case InsertError
                                strErrorMessage = "Errore nell'inserimento del record."
                            Case UpdateError
                                strErrorMessage = "Errore nell'aggiornamento del record."
                            Case Else
                                strErrorMessage = "Errore sconosciuto"
                        End Select
                    End If
                    
                    XALog.ReDim 0, lngRecordCount - 1, 0, 3
                    XALog.value(lngRecordCount - 1, 0) = varBufferEventi(14, lngEventsCounter)
                    XALog.value(lngRecordCount - 1, 1) = varBufferEventi(15, lngEventsCounter)
                    XALog.value(lngRecordCount - 1, 2) = varBufferEventi(16, lngEventsCounter)
                    XALog.value(lngRecordCount - 1, 3) = strErrorMessage
                    With tdbLog
                        .Refresh
                        .bookmark = lngRecordCount - 1
                    End With
                    TDBGrid1.bookmark = lngNumRows
                    DoEvents
                
                Next lngEventsCounter
                
                ' Close last transaction
                If blnSkipVVC Then
                    DataClass.DS.DA.EndTrans 1
                Else
                    DataClass.DS.DA.EndTrans 0
                End If
            
            End If

        End If

    Next lngRowsCounter

    VerificaAnomalie = True
    GoTo EndProc

ErrHandler:

    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmVVTSnapshotExport", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
    lngCountOK = 0
    
    VerificaAnomalie = False
    
EndProc:

    Set VVP = Nothing
    Set OCT = Nothing
    Set OCR = Nothing
End Function
