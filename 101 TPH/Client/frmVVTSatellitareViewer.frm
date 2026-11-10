VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmVVTSatellitareViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SITUAZIONE AUTOMEZZI IN TEMPO REALE"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11430
   Begin VB.CheckBox checkDisplayMessagesAllVehicles 
      Caption         =   "Visualizza messaggi di tuti gli automezzi"
      Height          =   195
      Left            =   4080
      TabIndex        =   79
      Top             =   5700
      Width           =   3075
   End
   Begin MSMask.MaskEdBox mskRefreshSeconds 
      Height          =   255
      Left            =   8760
      TabIndex        =   14
      Top             =   5700
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      ClipMode        =   1
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   10680
      TabIndex        =   16
      Top             =   5640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "REFRESH"
            Object.ToolTipText     =   "Refresh Griglia corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZOOM"
            Object.ToolTipText     =   "Zoom Griglia corrente"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9180
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":0000
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":0112
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":0224
            Key             =   "Editing"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":0E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":1404
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":199E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":1F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":24D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":2A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":3006
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":35A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":3B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":40D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":466E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":4C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":51A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   795
      Left            =   60
      TabIndex        =   6
      Top             =   5940
      Width           =   11295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   7380
         Picture         =   "frmVVTSatellitareViewer.frx":573C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Invia Viaggio"
         Height          =   375
         Left            =   7260
         Picture         =   "frmVVTSatellitareViewer.frx":583E
         TabIndex        =   4
         Tag             =   "CUSTOM1"
         Top             =   300
         Width           =   1875
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   5460
         Picture         =   "frmVVTSatellitareViewer.frx":5940
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Seleziona Tutti"
         Height          =   375
         Left            =   5340
         Picture         =   "frmVVTSatellitareViewer.frx":5A42
         TabIndex        =   3
         Tag             =   "SAVE"
         Top             =   300
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   3300
         Picture         =   "frmVVTSatellitareViewer.frx":5B44
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
      End
      Begin VB.CommandButton CmdOpenTrip 
         Caption         =   "Apri &Viaggio"
         Height          =   375
         Left            =   3240
         Picture         =   "frmVVTSatellitareViewer.frx":5C46
         TabIndex        =   2
         Tag             =   "OPEN"
         Top             =   300
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   9840
         Picture         =   "frmVVTSatellitareViewer.frx":5D48
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9600
         Picture         =   "frmVVTSatellitareViewer.frx":5E4A
         TabIndex        =   5
         Tag             =   "EXIT"
         Top             =   300
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1740
         Picture         =   "frmVVTSatellitareViewer.frx":5F4C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - &Annulla"
         Height          =   375
         Left            =   1680
         Picture         =   "frmVVTSatellitareViewer.frx":604E
         TabIndex        =   1
         Tag             =   "UNDO"
         Top             =   300
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
         Left            =   180
         Picture         =   "frmVVTSatellitareViewer.frx":6150
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   375
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - &Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTSatellitareViewer.frx":6252
         TabIndex        =   0
         Tag             =   "FIND"
         Top             =   300
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   10260
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":6354
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTSatellitareViewer.frx":64C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer timerRefresh 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   10800
      Top             =   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6750
      Width           =   11430
      _ExtentX        =   20161
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
            Object.Width           =   14711
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10.31"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "30/08/2005"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   11175
      Begin VB.Frame frameMessages 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3135
         Left            =   60
         TabIndex        =   18
         Top             =   2220
         Width           =   11055
         Begin VB.ComboBox comboAutomezzi 
            Height          =   315
            Left            =   8220
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   2340
            Width           =   1635
         End
         Begin TrueDBGrid60.TDBGrid TDBGrid4 
            Height          =   1635
            Left            =   60
            OleObjectBlob   =   "frmVVTSatellitareViewer.frx":6638
            TabIndex        =   27
            Top             =   360
            Width           =   10935
         End
         Begin MSComctlLib.TabStrip tabstripFilterMessagesDays 
            Height          =   375
            Left            =   8820
            TabIndex        =   19
            Top             =   60
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            MultiRow        =   -1  'True
            Separators      =   -1  'True
            TabMinWidth     =   600
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   5
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Oggi"
                  Object.Tag             =   "0"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "7"
                  Object.Tag             =   "7"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "15"
                  Object.Tag             =   "15"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "30"
                  Object.Tag             =   "30"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Tutti"
                  Object.Tag             =   "-1"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame3 
            Caption         =   " Invio nuovo messaggio "
            Height          =   1035
            Left            =   60
            TabIndex        =   20
            Top             =   2100
            Width           =   10935
            Begin VB.CheckBox checkForzaAutomezzo 
               Caption         =   "Motrice"
               Height          =   315
               Left            =   7260
               TabIndex        =   81
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton cmdClearMessage 
               Caption         =   "Pulisci"
               Height          =   315
               Left            =   9900
               TabIndex        =   24
               Top             =   240
               Width           =   915
            End
            Begin VB.CommandButton cmdSendMessage 
               Caption         =   "Invia"
               Height          =   315
               Left            =   9900
               TabIndex        =   23
               Top             =   600
               Width           =   915
            End
            Begin VB.TextBox textMessage 
               Height          =   315
               Left            =   1080
               MaxLength       =   150
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   600
               Width           =   8715
            End
            Begin VB.TextBox textObject 
               Height          =   315
               Left            =   1080
               MaxLength       =   30
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Top             =   240
               Width           =   6015
            End
            Begin VB.Label Label5 
               Caption         =   "Oggetto"
               Height          =   195
               Left            =   180
               TabIndex        =   26
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Testo"
               Height          =   195
               Left            =   180
               TabIndex        =   25
               Top             =   660
               Width           =   735
            End
         End
         Begin MSComctlLib.TabStrip tabstripFilterMessages 
            Height          =   375
            Left            =   60
            TabIndex        =   78
            Top             =   60
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   661
            MultiRow        =   -1  'True
            TabStyle        =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Inviati e Ricevuti"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Inviati"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ricevuti"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frameConsegne 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3135
         Left            =   120
         TabIndex        =   28
         Top             =   2340
         Width           =   10995
         Begin TrueDBGrid60.TDBGrid TDBGrid3 
            Height          =   1455
            Left            =   0
            OleObjectBlob   =   "frmVVTSatellitareViewer.frx":F1AA
            TabIndex        =   30
            Top             =   1620
            Width           =   10935
         End
         Begin TrueDBGrid60.TDBGrid TDBGrid2 
            Height          =   1575
            Left            =   0
            OleObjectBlob   =   "frmVVTSatellitareViewer.frx":18030
            TabIndex        =   29
            Top             =   0
            Width           =   10935
         End
      End
      Begin MSComctlLib.ImageList ImageList4 
         Left            =   8520
         Top             =   780
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSatellitareViewer.frx":252A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSatellitareViewer.frx":25840
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame frameLegenda 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3015
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   10935
         Begin VB.Frame Frame7 
            Caption         =   " Messaggi "
            Height          =   915
            Left            =   3840
            TabIndex        =   45
            Top             =   1920
            Width           =   3315
            Begin VB.Shape shapeSMELegenda1 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   300
               Width           =   195
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Non letto"
               Height          =   195
               Left            =   540
               TabIndex        =   46
               Top             =   300
               Width           =   645
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " Consegne "
            Height          =   1635
            Left            =   3840
            TabIndex        =   41
            Top             =   60
            Width           =   3315
            Begin VB.Shape shapeVVPLegenda1 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   300
               Width           =   195
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Pendiente"
               Height          =   195
               Left            =   540
               TabIndex        =   44
               Top             =   300
               Width           =   720
            End
            Begin VB.Shape shapeVVPLegenda2 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   660
               Width           =   195
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "En curso"
               Height          =   195
               Left            =   540
               TabIndex        =   43
               Top             =   660
               Width           =   630
            End
            Begin VB.Shape shapeVVPLegenda3 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   1020
               Width           =   195
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Finalizado"
               Height          =   195
               Left            =   540
               TabIndex        =   42
               Top             =   1020
               Width           =   705
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " Viaggio "
            Height          =   2835
            Left            =   180
            TabIndex        =   34
            Top             =   60
            Width           =   3315
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Cancelado"
               Height          =   195
               Left            =   540
               TabIndex        =   40
               Top             =   2100
               Width           =   765
            End
            Begin VB.Shape shapeVVTLegenda6 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   2100
               Width           =   195
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Finalizado"
               Height          =   195
               Left            =   540
               TabIndex        =   39
               Top             =   1740
               Width           =   705
            End
            Begin VB.Shape shapeVVTLegenda5 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   1740
               Width           =   195
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Llegada en factoria, En curso"
               Height          =   195
               Left            =   540
               TabIndex        =   38
               Top             =   1380
               Width           =   2085
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Inicio de viaje"
               Height          =   195
               Left            =   540
               TabIndex        =   37
               Top             =   1020
               Width           =   975
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Inviado, Recibido"
               Height          =   195
               Left            =   540
               TabIndex        =   36
               Top             =   660
               Width           =   1245
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Pendiente"
               Height          =   195
               Left            =   540
               TabIndex        =   35
               Top             =   300
               Width           =   720
            End
            Begin VB.Shape shapeVVTLegenda4 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   1380
               Width           =   195
            End
            Begin VB.Shape shapeVVTLegenda3 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   1020
               Width           =   195
            End
            Begin VB.Shape shapeVVTLegenda2 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   660
               Width           =   195
            End
            Begin VB.Shape shapeVVTLegenda1 
               FillStyle       =   0  'Solid
               Height          =   195
               Left            =   180
               Top             =   300
               Width           =   195
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   " Icone "
            Height          =   2715
            Left            =   7500
            TabIndex        =   32
            Top             =   60
            Width           =   3315
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Dato provvisorio"
               Height          =   195
               Left            =   540
               TabIndex        =   33
               Top             =   300
               Width           =   1155
            End
            Begin VB.Image Image1 
               Height          =   240
               Left            =   240
               Picture         =   "frmVVTSatellitareViewer.frx":25DDA
               Top             =   300
               Width           =   240
            End
         End
      End
      Begin VB.Frame frameFilters 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3135
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   10935
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   1
            Left            =   8220
            Picture         =   "frmVVTSatellitareViewer.frx":26364
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton cmdCalendar 
            Height          =   300
            Index           =   0
            Left            =   3240
            Picture         =   "frmVVTSatellitareViewer.frx":26466
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   120
            Width           =   375
         End
         Begin VB.CheckBox chkActiveOnly 
            Caption         =   "Solo viaggi non annullati"
            Height          =   255
            Left            =   2940
            TabIndex        =   48
            Top             =   2760
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin MSMask.MaskEdBox mskMaxRecords 
            Height          =   300
            Left            =   1740
            TabIndex        =   51
            Top             =   2700
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateFrom 
            Height          =   255
            Left            =   1920
            TabIndex        =   52
            Tag             =   "5"
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateTo 
            Height          =   255
            Left            =   6900
            TabIndex        =   53
            Tag             =   "6"
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   1920
            TabIndex        =   54
            Tag             =   "0"
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   6900
            TabIndex        =   55
            Tag             =   "1"
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   1920
            TabIndex        =   56
            Tag             =   "11"
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0000000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   6900
            TabIndex        =   57
            Tag             =   "12"
            Top             =   780
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
            Left            =   1920
            TabIndex        =   58
            Tag             =   "3"
            Top             =   1560
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
            Left            =   1920
            TabIndex        =   59
            Tag             =   "13"
            Top             =   1920
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
            Left            =   1920
            TabIndex        =   60
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
            Left            =   2520
            TabIndex        =   74
            Tag             =   "3"
            Top             =   1560
            Width           =   5595
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Società"
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   73
            Top             =   1560
            Width           =   615
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   8220
            Tag             =   "3"
            Top             =   1560
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
            Left            =   2520
            TabIndex        =   72
            Tag             =   "13"
            Top             =   1920
            Width           =   5595
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Unità di Trasporto"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   71
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   8220
            Tag             =   "13"
            Top             =   1920
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   18
            Left            =   8220
            Tag             =   "7"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Dep."
            Height          =   255
            Index           =   18
            Left            =   1320
            TabIndex        =   70
            Top             =   1200
            Width           =   495
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
            Left            =   2520
            TabIndex        =   69
            Tag             =   "7"
            Top             =   1200
            Width           =   5595
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   2
            Left            =   3240
            Tag             =   "11"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Dal Num. &Viaggio Def."
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   68
            Top             =   840
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   1
            Left            =   8220
            Tag             =   "12"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Al Num. Viaggio Def."
            Height          =   255
            Index           =   1
            Left            =   5220
            TabIndex        =   67
            Top             =   780
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   3240
            Tag             =   "11"
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Dal Num. &Viaggio"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   66
            Top             =   480
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   4
            Left            =   8220
            Tag             =   "12"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Al Num. Viaggio"
            Height          =   255
            Index           =   4
            Left            =   5220
            TabIndex        =   65
            Top             =   420
            Width           =   1575
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   0
            Left            =   8700
            Tag             =   "6"
            Top             =   60
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Alla Data Inizio Viaggio"
            Height          =   255
            Index           =   0
            Left            =   5100
            TabIndex        =   64
            Top             =   60
            Width           =   1695
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   3720
            Tag             =   "5"
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "&Dalla Data Inizio Viaggio"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   63
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Restituisci al massimo"
            Height          =   255
            Index           =   16
            Left            =   60
            TabIndex        =   62
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "viaggi"
            Height          =   255
            Left            =   2280
            TabIndex        =   61
            Top             =   2760
            Width           =   495
         End
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   1815
         Left            =   0
         OleObjectBlob   =   "frmVVTSatellitareViewer.frx":26568
         TabIndex        =   75
         Top             =   0
         Width           =   11175
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3675
         Left            =   0
         TabIndex        =   76
         Top             =   1860
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6482
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Condizioni"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Consegne"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Messaggi"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Legenda"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlToolbarIcons 
         Left            =   4995
         Top             =   3255
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSatellitareViewer.frx":341F6
               Key             =   "View Details"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVVTSatellitareViewer.frx":34308
               Key             =   "View List"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frameZoomGrid 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   5535
      Left            =   120
      TabIndex        =   77
      Top             =   60
      Width           =   11175
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   10200
      Picture         =   "frmVVTSatellitareViewer.frx":3441A
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Refresh ogni                  secondi"
      Height          =   255
      Left            =   7740
      TabIndex        =   15
      Top             =   5700
      Width           =   2355
   End
End
Attribute VB_Name = "frmVVTSatellitareViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private XA As XArrayDB
Private XAConsegne As XArrayDB
Private XAEventi As XArrayDB
Private XAMessaggi As XArrayDB

Private XALog As XArrayDB

Public WithEvents DataClass As CViaggioSelezioneSatellitare
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private AXX As CParametro
Private VVP As CViaggioProdotto
Private SEV As CSatellitareEvento
Private SME As CSatellitareMessaggio

Private strCurrentVVTId As String
Private strCurrentAMMId As String

Dim gridZoomGrid As TDBGrid
Dim frameZoomGridContainer As Frame
Dim sngZoomGridLeft As Single
Dim sngZoomGridTop As Single
Dim sngZoomGridWidth As Single
Dim sngZoomGridHeight As Single

Private lngRefreshTicks As Long
Private lngRefreshSeconds As Long

Private Const GRID_VVT_ROW_BACKCOLOR1 = &H9DF4D6
Private Const GRID_VVT_ROW_BACKCOLOR2 = &HBFEAC1
Private Const GRID_VVT_ROW_BACKCOLOR3 = &HFFCAAA
Private Const GRID_VVT_ROW_BACKCOLOR4 = &H84F9AA
Private Const GRID_VVT_ROW_BACKCOLOR5 = &HB4ACBF

Private Const GRID_VVP_ROW_BACKCOLOR1 = &HFFCAAA
Private Const GRID_VVP_ROW_BACKCOLOR2 = &H84F9AA

Private Const GRID_SME_ROW_BACKCOLOR1 = &HFFC0C0
Private Const GRID_SME_ROW_FORECOLOR1 = &H800000

Private Const GRID_VVT_COLUMN_VVTDTINV = 0
Private Const GRID_VVT_COLUMN_VVTIDVVT = 1
Private Const GRID_VVT_COLUMN_VVTNRSEQ = 3
Private Const GRID_VVT_COLUMN_MGSCOUNTS = 6
Private Const GRID_VVT_COLUMN_VVTIDAM1 = 8
Private Const GRID_VVT_COLUMN_VVTKMDEF = 12
Private Const GRID_VVT_COLUMN_VVTFLSTA_DESCR = 14
Private Const GRID_VVT_COLUMN_VVTHRINC = 16
Private Const GRID_VVT_COLUMN_VVTTMPCA = 17
Private Const GRID_VVT_COLUMN_VVTHRINV = 18
Private Const GRID_VVT_COLUMN_VVTTMPVI = 21
Private Const GRID_VVT_COLUMN_VVTFLSTA = 22
Private Const GRID_VVT_COLUMN_ESTADO = 23
Private Const GRID_VVT_COLUMN_ESTADO_DESCR = 24

Private Const GRID_VVP_COLUMN_VVPIDVVT = 0
Private Const GRID_VVP_COLUMN_VVPIDVVC = 1
Private Const GRID_VVP_COLUMN_VVPIDVVP = 2
Private Const GRID_VVP_COLUMN_VVCKMPAD = 7
Private Const GRID_VVP_COLUMN_VVPHRINS = 8
Private Const GRID_VVP_COLUMN_TEMPOSCAR = 9
Private Const GRID_VVP_COLUMN_VVPFLSTA = 22
Private Const GRID_VVP_COLUMN_VVCTEMVG = 23
Private Const GRID_VVP_COLUMN_VVPHRFIS = 24

Private Const GRID_SME_COLUMN_SMEIDSME = 0
Private Const GRID_SME_COLUMN_SMEDTINS = 1
Private Const GRID_SME_COLUMN_SMEFLPRODESC = 2
Private Const GRID_SME_COLUMN_SMEIDAMM = 3
Private Const GRID_SME_COLUMN_SMEFLREA = 8
Private Const GRID_SME_COLUMN_SMEFLPRO = 9

Private Const GRID_SEV_COLUMN_SEVIDSEV = 0
Private Const GRID_SEV_COLUMN_SEVDTEVE = 5

Private Const REGISTRY_KEY_NAME_REFRESH_SECONDS = "RefreshSecondsSatellitare"

Private Const SECONDS_PER_TICK = 2
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "View Details"
            'Da fare: Aggiunge il codice per il pulsante 'View Details'.
            MsgBox "Aggiunge il codice per il pulsante 'View Details'."
        Case "View List"
            'Da fare: Aggiunge il codice per il pulsante 'View List'.
            MsgBox "Aggiunge il codice per il pulsante 'View List'."
    End Select
End Sub


Private Sub DisplaySelection(List As Variant)
    Dim i As Long
    
    If IsEmpty(List) Then
        ClearXArray XA, TDBGrid1
        TDBGrid1.bookmark = Null
        
        ClearXArray XAConsegne, TDBGrid2
        TDBGrid2.bookmark = Null
        
        ClearXArray XAEventi, TDBGrid3
        TDBGrid3.bookmark = Null
        
'        ClearXArray XAMessaggi, TDBGrid4
'        TDBGrid4.Bookmark = Null
    Else
        RefreshXArray XA, TDBGrid1, List
        TDBGrid1.bookmark = 0
    End If
    
End Sub
Private Sub DisplaySelectionVVP(List As Variant)
    Dim i As Long
    
    If IsEmpty(List) Then
        ClearXArray XAConsegne, TDBGrid2
        TDBGrid2.bookmark = Null
    Else
        RefreshXArray XAConsegne, TDBGrid2, List
        TDBGrid2.bookmark = 0
    End If
End Sub
Private Sub DisplaySelectionSEV(List As Variant)
    Dim i As Long
    
    If IsEmpty(List) Then
        ClearXArray XAEventi, TDBGrid3
        TDBGrid3.bookmark = Null
    Else
        RefreshXArray XAEventi, TDBGrid3, List
        TDBGrid3.bookmark = 0
    End If
End Sub
Private Sub DisplaySelectionSME(List As Variant, bookmark As Variant)
    Dim i As Long
    
    If IsEmpty(List) Then
        ClearXArray XAMessaggi, TDBGrid4
        TDBGrid4.bookmark = Null
    Else
        RefreshXArray XAMessaggi, TDBGrid4, List
        TDBGrid4.bookmark = bookmark
        TDBGrid4.Enabled = True
    End If

    RefreshNewMessages
End Sub
Private Function IsXAEmpty(XA As XArrayDB) As Boolean

    If XA Is Nothing Then
        IsXAEmpty = True
        Exit Function
    Else
        If XA.Count(1) <= 1 And XA.Count(2) <= 1 Then
            IsXAEmpty = True
            Exit Function
        End If
    End If
    IsXAEmpty = False
End Function
Private Function GetSelectedID() As Variant
    If IsXAEmpty(XA) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTIDVVT)
    End If
End Function
Private Function GetSelectedAMMID() As Variant
    If IsXAEmpty(XA) Then
        GetSelectedAMMID = vbNullString
    Else
        GetSelectedAMMID = XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTIDAM1)
    End If
End Function
Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    Dim lngViaggiCounter As Long
    Dim strVVPId As String
    Dim lngGridRowIndex As Long

    lngGridRowIndex = -1
    If Not IsXAEmpty(XA) And Not IsNull(TDBGrid1.bookmark) Then
        strVVPId = XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTIDVVT)
    Else
        strVVPId = ""
    End If

    With DataClass.DS
        If .Fields.Validate Then
            If .ScanClientCond Then
                ReDim varOrderFields(3) As Variant
                Me.MousePointer = vbHourglass
                varOrderFields(0) = "37"
                varOrderFields(1) = "7"
                varOrderFields(2) = "13"
                varOrderFields(3) = "38"

                blnActiveOnly = (chkActiveOnly = vbChecked)
                DataClass.DS.Fields("10").FindValue(1) = "D"
                DataClass.DS.Fields("10").FindValue(2) = "C"
                If Not DataClass.GetSelection(blnActiveOnly, varOrderFields, CLng(mskMaxRecords.Text)) Then
                    GoTo ErrHandler
                End If
                varBuffer = DataClass.DS.List

                If Not IsEmpty(varBuffer) Then
                    For lngViaggiCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
                    
                        If (varBuffer(GRID_VVT_COLUMN_VVTFLSTA, lngViaggiCounter) = "0" Or IsNull(varBuffer(GRID_VVT_COLUMN_VVTFLSTA, lngViaggiCounter))) _
                            And Not IsNull(varBuffer(GRID_VVT_COLUMN_ESTADO_DESCR, lngViaggiCounter)) Then
                            varBuffer(GRID_VVT_COLUMN_VVTFLSTA_DESCR, lngViaggiCounter) = varBuffer(GRID_VVT_COLUMN_ESTADO_DESCR, lngViaggiCounter)
                        End If
                    
                        If strVVPId = varBuffer(GRID_VVT_COLUMN_VVTIDVVT, lngViaggiCounter) Then
                            lngGridRowIndex = lngViaggiCounter
                        End If
                    
                    Next lngViaggiCounter
                End If
                
                DisplaySelection varBuffer
                If IsEmpty(varBuffer) Then
                    GoTo ErrHandler
                End If
            End If
        Else
            ShowValidationError DataClass.DS
            GoTo ErrHandler
        End If
    End With
    
    If lngGridRowIndex <> -1 Then
        TDBGrid1.bookmark = lngGridRowIndex
    End If
    
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
    
    ReDim varReportsList(0 To 2, 0 To 3) As Variant
    
    Select Case Cmd
        Case fcSave
            Dim lngRowsCounter As Long
            With TDBGrid1.SelBookmarks
                For lngRowsCounter = XA.LowerBound(1) To XA.UpperBound(1)
                    .Add lngRowsCounter
                Next lngRowsCounter
            End With
            TDBGrid1.Refresh

        Case fcUndo
            If (State = fsParamInput) Or (State = fsSelection) Then
                ClearForm
                SetState fsParamInput
            End If
        Case fcFind
            
            If SelectRecords Then
                SetState fsSelection
                TabStrip1.Tabs(2).Selected = True
            End If
        Case fcPrint
'            If TDBGrid1.Bookmark <> vbNullString Then
'                Call SincronizzaDatiViaggio(srLetturaDati)
'                Call SetInterfaceForUpdTrip(arModificaViaggio)
'            End If
        
        Case fcOpen
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID()
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        
        Case fcExit
            Unload Me
        
        Case fcCustom1
            SendViaggi
            SelectRecords

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

Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        
        Case fsParamInput
'            If Not Loading Then
                TabStrip1.Tabs(1).Selected = True
'                mskDateFrom.SetFocus
'            End If
            mskDateFrom.Text = Format(Date, "dd/mm/yy")
            mskDateTo.Text = Format(Date, "dd/mm/yy")
        
        Case fsSelection
            StatusBar1.Panels(2).Text = StatusBar1.Panels(2).Text & " - Viaggi trovati: " & DataClass.DS.HowMany
            TabStrip1.Tabs(2).Selected = True
            If Not Loading Then
'                TDBGrid2.SetFocus
            End If
    End Select
End Sub

Private Sub checkDisplayMessagesAllVehicles_Click()
    SelectRecordsSME
End Sub

Private Sub checkForzaAutomezzo_Click()
    RefreshFrameMessageCaption
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox

    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Dalla data inizio viaggio"
        Case 1
            Set M = mskDateTo
            strTitle = "Alla data inizio viaggio"
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub cmdEventi_Click()
    Dim LinkForm As Form
    
    Set LinkForm = DisplayForm(fcVIAGVPI)
    If Not LinkForm Is Nothing Then
        With LinkForm
            Set .ParentForm = Me
            If .GetRecord(XA.value(TDBGrid1.bookmark, 0)) Then
                .SetState fsEdit
            Else
                .SetState fsInsert
            End If
            Me.Enabled = False
        End With
    End If
End Sub
Private Sub cmdClearMessage_Click()
    textMessage.Text = ""
    textObject.Text = ""
End Sub

Public Function GetSelectedAutomezzo() As String
    
    Dim strIDAutomezzo As String
    
    strIDAutomezzo = ""
    
    If checkForzaAutomezzo.value = 1 Then
        strIDAutomezzo = comboAutomezzi.List(comboAutomezzi.ListIndex)
    Else
        If checkDisplayMessagesAllVehicles.value = 0 And Not IsXAEmpty(XA) Then
            If Not IsXAEmpty(XA) And Not IsNull(TDBGrid1.bookmark) Then
                strIDAutomezzo = XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTIDAM1)
            End If
        Else
            If Not IsXAEmpty(XAMessaggi) And Not IsNull(TDBGrid4.bookmark) Then
                strIDAutomezzo = XAMessaggi.value(TDBGrid4.bookmark, GRID_SME_COLUMN_SMEIDAMM)
            End If
        End If
    End If
    
    GetSelectedAutomezzo = strIDAutomezzo

End Function

Private Sub cmdSendMessage_Click()

    Dim SME As CSatellitareMessaggio
    Dim strIDAutomezzo As String
    
    Set SME = New CSatellitareMessaggio
    
    With SME.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    strIDAutomezzo = GetSelectedAutomezzo()

    If strIDAutomezzo <> "" Then
        If SME.Insert_Message_ToSend(strIDAutomezzo, textMessage.Text, textObject.Text) Then
            If SelectRecordsSME() Then
            End If
        End If
    End If

    Set SME = Nothing
End Sub

Private Sub comboAutomezzi_Change()
    RefreshFrameMessageCaption
End Sub

Private Sub comboAutomezzi_Click()
    RefreshFrameMessageCaption
End Sub

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim strErr As String
    
    ClientKeyUp KeyCode, Me
    Select Case KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strIDParam As String
    ReDim ColTags(1 To 2) As Variant
    Dim Col As Column
    
    Loading = True
    ColTags(1) = "0"
    ColTags(2) = "1"
    mskMaxRecords = MAX_RECORDS

    Set DataClass = New CViaggioSelezioneSatellitare
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set SEV = New CSatellitareEvento
    With SEV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set SME = New CSatellitareMessaggio
    With SME.DS
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
    
    SetState fsParamInput
    DoCommand fcFind
    
    Loading = False
        
    strCurrentVVTId = ""
    strCurrentAMMId = ""
    
    ' Imposta i colori nella legenda
    shapeVVTLegenda1.FillColor = TDBGrid1.BackColor
    shapeVVTLegenda2.FillColor = GRID_VVT_ROW_BACKCOLOR1
    shapeVVTLegenda3.FillColor = GRID_VVT_ROW_BACKCOLOR2
    shapeVVTLegenda4.FillColor = GRID_VVT_ROW_BACKCOLOR3
    shapeVVTLegenda5.FillColor = GRID_VVT_ROW_BACKCOLOR4
    shapeVVTLegenda6.FillColor = GRID_VVT_ROW_BACKCOLOR5

    shapeVVPLegenda1.FillColor = TDBGrid2.BackColor
    shapeVVPLegenda2.FillColor = GRID_VVP_ROW_BACKCOLOR1
    shapeVVPLegenda3.FillColor = GRID_VVP_ROW_BACKCOLOR2

    shapeSMELegenda1.FillColor = GRID_SME_ROW_BACKCOLOR1
    
    ' Imposta intervallo timer
    SetRefreshInterval CLng(gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_REFRESH_SECONDS, "15", True)), True
    LoadAutomezzi

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AXX = Nothing
    Set VVP = Nothing
    Set XA = Nothing
    Set XAConsegne = Nothing
    Set XALog = Nothing
    Set SEV = Nothing
    Set SME = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set frmVVTSatellitareViewer = Nothing
End Sub
Private Function SelectRecordsVVP() As Boolean
   On Error GoTo ErrHandler

'    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim lngRecordsCounter As Long
    Dim strVVPId As String
    Dim lngGridRowIndex As Long
    
    ReDim varOrderFields(1) As Variant

    Me.MousePointer = vbHourglass
    
    If GetSelectedID() <> vbNullString Then
    
        lngGridRowIndex = -1
        If Not IsXAEmpty(XAConsegne) And Not IsNull(TDBGrid2.bookmark) Then
            strVVPId = strCurrentVVTId & "-" & XAConsegne.value(TDBGrid2.bookmark, GRID_VVP_COLUMN_VVPIDVVC) & "-" & XAConsegne.value(TDBGrid2.bookmark, GRID_VVP_COLUMN_VVPIDVVP)
        Else
            strVVPId = ""
        End If

'        varOrderFields(0) = "1"
'        varOrderFields(1) = "2"
        blnActiveOnly = (chkActiveOnly = vbChecked)
        lngHowMany = 0
        If Not VVP.GetSelectionViaggioProdottiEventi(GetSelectedID(), blnActiveOnly, lngHowMany) Then
            GoTo ErrHandler
        End If
        varBuffer = VVP.DS.List
        
        If IsEmpty(varBuffer) = False Then
            For lngRecordsCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
            
                ' Imposta la data di inizio scarico provvisoria o definitiva
                If varBuffer(GRID_VVP_COLUMN_VVPFLSTA, lngRecordsCounter) = "0" Or IsNull(varBuffer(GRID_VVT_COLUMN_VVTFLSTA, lngRecordsCounter)) Then
                    varBuffer(GRID_VVP_COLUMN_VVPHRINS, lngRecordsCounter) = varBuffer(GRID_VVP_COLUMN_VVCTEMVG, lngRecordsCounter)
                End If
                
                ' Imposta il tempo di scarico
                If varBuffer(GRID_VVP_COLUMN_VVPFLSTA, lngRecordsCounter) = "2" Then
                    varBuffer(GRID_VVP_COLUMN_TEMPOSCAR, lngRecordsCounter) = DateDiff("n", varBuffer(GRID_VVP_COLUMN_VVPHRINS, lngRecordsCounter), varBuffer(GRID_VVP_COLUMN_VVPHRFIS, lngRecordsCounter))
                End If
                
                If strVVPId = varBuffer(GRID_VVP_COLUMN_VVPIDVVT, lngRecordsCounter) & "-" & varBuffer(GRID_VVP_COLUMN_VVPIDVVC, lngRecordsCounter) & "-" & varBuffer(GRID_VVP_COLUMN_VVPIDVVP, lngRecordsCounter) Then
                    lngGridRowIndex = lngRecordsCounter
                End If
            Next lngRecordsCounter
        End If
    
    '    If IsEmpty(varBuffer) Then
    '        GoTo ErrHandler
    '    End If
        DisplaySelectionVVP varBuffer
        
        If lngGridRowIndex <> -1 Then
            TDBGrid2.bookmark = lngGridRowIndex
        End If
    End If
    
    SelectRecordsVVP = True
    GoTo EndProc
ErrHandler:
    SelectRecordsVVP = False
EndProc:
    Me.MousePointer = vbDefault
End Function

Private Function SelectRecordsSME() As Boolean
   On Error GoTo ErrHandler

    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    Dim lngMessagesCounter As Long
    Dim strSMEId As String
    Dim lngGridRowIndex As Long

'    ReDim varOrderFields(0) As Variant

    Me.MousePointer = vbHourglass
    
    SME.DS.Fields.ResetFields
    If tabstripFilterMessages.SelectedItem.Index = 2 Then
        SME.DS.Fields("3").FindValue = "D"
    Else
        If tabstripFilterMessages.SelectedItem.Index = 3 Then
            SME.DS.Fields("3").FindValue = "A"
        End If
    End If
    
    lngGridRowIndex = -1
    If Not IsXAEmpty(XAMessaggi) And Not IsNull(TDBGrid4.bookmark) Then
        strSMEId = XAMessaggi.value(TDBGrid4.bookmark, GRID_SME_COLUMN_SMEIDSME)
    Else
        strSMEId = ""
    End If
    
    ' Devono essere visualizzati solo i messaggi dell'autoveicolo
    If checkDisplayMessagesAllVehicles.value = 0 Then
        
        If GetSelectedAMMID() <> vbNullString Then
            blnActiveOnly = (chkActiveOnly = vbChecked)
                                
            If Not SME.GetSelectionMessaggiAutomezzoOrderDesc(GetSelectedAMMID(), CLng(tabstripFilterMessagesDays.Tabs(tabstripFilterMessagesDays.SelectedItem.Index).Tag), blnActiveOnly, 0) Then
                GoTo ErrHandler
            End If
            varBuffer = SME.DS.List
    
        End If
    
    ' Devono essere visualizzati i messaggi di tutti gli autoveicoli
    Else
        If Not SME.GetSelectionMessaggiOrderDesc(CLng(tabstripFilterMessagesDays.Tabs(tabstripFilterMessagesDays.SelectedItem.Index).Tag), blnActiveOnly, 0) Then
            GoTo ErrHandler
        End If
        varBuffer = SME.DS.List
    End If

    If Not IsEmpty(varBuffer) Then
        For lngMessagesCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
            
            ' Converte la data
            If VarType(varBuffer(GRID_SME_COLUMN_SMEDTINS, lngMessagesCounter)) = vbString Then
                varBuffer(GRID_SME_COLUMN_SMEDTINS, lngMessagesCounter) = _
                        CDate(Left(varBuffer(GRID_SME_COLUMN_SMEDTINS, lngMessagesCounter), _
                        Len(varBuffer(GRID_SME_COLUMN_SMEDTINS, lngMessagesCounter)) - 4))
            End If

            If strSMEId = varBuffer(GRID_SME_COLUMN_SMEIDSME, lngMessagesCounter) Then
                lngGridRowIndex = lngMessagesCounter
            End If

        Next lngMessagesCounter
    End If
    
    DisplaySelectionSME varBuffer, IIf(lngGridRowIndex = -1, 0, lngGridRowIndex)
    
    SelectRecordsSME = True
    GoTo EndProc
ErrHandler:
    SelectRecordsSME = False
EndProc:
    Me.MousePointer = vbDefault
End Function

Private Function SelectRecordsSEV() As Boolean
   On Error GoTo ErrHandler

    Dim varOrderFields As Variant
    Dim blnActiveOnly As Boolean
    Dim varBuffer As Variant
    Dim lngEventsCounter As Long
    Dim strSEVId As String
    Dim lngGridRowIndex As Long

'    ReDim varOrderFields(0) As Variant

    Me.MousePointer = vbHourglass
    
    If GetSelectedID() <> vbNullString Then
    
'        varOrderFields(0) = "5"
        
        lngGridRowIndex = -1
        If Not IsXAEmpty(XAEventi) And Not IsNull(TDBGrid3.bookmark) Then
            strSEVId = XAEventi.value(TDBGrid3.bookmark, GRID_SEV_COLUMN_SEVIDSEV)
        Else
            strSEVId = ""
        End If
        
        blnActiveOnly = (chkActiveOnly = vbChecked)
        If Not SEV.GetSelectionEvents(GetSelectedID(), blnActiveOnly, 0) Then
            GoTo ErrHandler
        End If
        varBuffer = SEV.DS.List
        
        If Not IsEmpty(varBuffer) Then
            For lngEventsCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
                
                ' Converte la data
                If VarType(varBuffer(GRID_SEV_COLUMN_SEVDTEVE, lngEventsCounter)) = vbString Then
                    
                    varBuffer(GRID_SEV_COLUMN_SEVDTEVE, lngEventsCounter) = _
                    CDate(Left(varBuffer(GRID_SEV_COLUMN_SEVDTEVE, lngEventsCounter), _
                    Len(varBuffer(GRID_SEV_COLUMN_SEVDTEVE, lngEventsCounter)) - 4))
                
                End If

                If strSEVId = varBuffer(GRID_SEV_COLUMN_SEVIDSEV, lngEventsCounter) Then
                    lngGridRowIndex = lngEventsCounter
                End If

            Next lngEventsCounter
        End If
        
        DisplaySelectionSEV varBuffer
    
    End If
    
    If lngGridRowIndex <> -1 Then
        TDBGrid3.bookmark = lngGridRowIndex
    End If

    SelectRecordsSEV = True
    GoTo EndProc
    
ErrHandler:
    SelectRecordsSEV = False
EndProc:
    Me.MousePointer = vbDefault
End Function
Private Sub SetRefreshInterval(lngSeconds As Long, blnRefreshControl As Boolean)
    StopRefresh
    lngRefreshSeconds = lngSeconds
    If lngRefreshSeconds > 0 And lngRefreshSeconds < 20 Then
        lngRefreshSeconds = 20
    End If
    If blnRefreshControl = True Then
        mskRefreshSeconds.Text = CStr(lngRefreshSeconds) + Space(mskRefreshSeconds.MaxLength - Len(CStr(lngRefreshSeconds)))
    End If
    StartRefresh
End Sub

Private Sub mskRefreshSeconds_Validate(Cancel As Boolean)
On Error GoTo ErrHandler
    Dim blnDummy As Boolean
    SetRefreshInterval CLng(mskRefreshSeconds.FormattedText), True
    If lngRefreshSeconds <> 0 Then
        blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_REFRESH_SECONDS, CStr(lngRefreshSeconds), True)
    End If
    Exit Sub
ErrHandler:
    Cancel = True
End Sub
Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Index

        Case 1
            frameFilters.Visible = True
            frameConsegne.Visible = False
            frameMessages.Visible = False
            frameLegenda.Visible = False

        Case 2
            frameFilters.Visible = False
            frameConsegne.Visible = True
            frameMessages.Visible = False
            frameLegenda.Visible = False

        Case 3
            frameFilters.Visible = False
            frameConsegne.Visible = False
            frameMessages.Visible = True
            frameLegenda.Visible = False

        Case 4
            frameFilters.Visible = False
            frameConsegne.Visible = False
            frameMessages.Visible = False
            frameLegenda.Visible = True

    End Select
End Sub

Private Sub tabstripFilterMessages_Click()
    SelectRecordsSME
End Sub

Private Sub tabstripFilterMessagesDays_Click()
    SelectRecordsSME
End Sub

Private Sub TDBGrid1_DblClick()
    DoCommand fcOpen
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal condition As Integer, ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid60.StyleDisp)
    If XA.Count(1) > 0 And Not IsEmpty(XA.value(0, 0)) Then
        
        Dim strStato As String
        Dim blnChangeStyle As Boolean
        strStato = IIf(IsNull(XA.value(bookmark, GRID_VVT_COLUMN_VVTFLSTA)), "0", XA.value(bookmark, GRID_VVT_COLUMN_VVTFLSTA))
        blnChangeStyle = False
        
        Select Case Col
            Case GRID_VVT_COLUMN_VVTHRINC
                If InStr("01", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
            
            Case GRID_VVT_COLUMN_VVTTMPCA
                If InStr("012", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
            
            Case GRID_VVT_COLUMN_VVTHRINV
                If InStr("012", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
            
            Case GRID_VVT_COLUMN_VVTTMPVI
                If InStr("0123", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
            
            Case GRID_VVT_COLUMN_VVTKMDEF
                If InStr("0123", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
        End Select
        
        If blnChangeStyle Then
            CellStyle.ForegroundPicture = ImageList2.ListImages(1).Picture
            CellStyle.TransparentForegroundPicture = True
            CellStyle.ForegroundPicturePosition = dbgFPLeft
        End If
    End If
End Sub
Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If Not IsXAEmpty(XA) Then
        Select Case IIf(IsNull(XA.value(bookmark, GRID_VVT_COLUMN_VVTFLSTA)) = True, "0", XA.value(bookmark, GRID_VVT_COLUMN_VVTFLSTA))
            Case "0"
                If XA.value(bookmark, GRID_VVT_COLUMN_ESTADO) <> "0" Then
                    RowStyle.BackColor = GRID_VVT_ROW_BACKCOLOR1
                End If

            Case "1"
                RowStyle.BackColor = GRID_VVT_ROW_BACKCOLOR2
                
            Case "2", "3"
                RowStyle.BackColor = GRID_VVT_ROW_BACKCOLOR3

            Case "4"
                RowStyle.BackColor = GRID_VVT_ROW_BACKCOLOR4
                
            Case "5"
                RowStyle.BackColor = GRID_VVT_ROW_BACKCOLOR5
        End Select
    End If
End Sub
Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim x As XTYPE
    Dim o As XORDER
    Static blnLastSortASC  As Boolean
    Static intLastCol As Integer
    
    If Not IsXAEmpty(XA) Then
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
            Case 1, 2, 3, 10, 11, 15, 16, 19, 20
                x = XTYPE_LONG
            Case 0, 14, 16, 17
                x = XTYPE_DATE
            Case Else
                x = XTYPE_STRING
        End Select
        XA.QuickSort 0, XA.UpperBound(1), ColIndex, o, x
        TDBGrid1.ReBind
        intLastCol = ColIndex
    End If
End Sub
Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim blnDummy As Boolean
    
    ' Selezionato un Viaggio, visualizza la Grid contenente gli scarichi
    ' con lo stato dell'evento e i messaggi
    If strCurrentVVTId <> GetSelectedID() Then
        blnDummy = SelectRecordsVVP()
        blnDummy = SelectRecordsSEV()
        If strCurrentAMMId <> GetSelectedAMMID() Then
            blnDummy = SelectRecordsSME()
        End If
    End If
    
    ' Aggiorna la caption della frame per invio messaggio
    RefreshFrameMessageCaption
    strCurrentVVTId = GetSelectedID()
    strCurrentAMMId = GetSelectedAMMID()
End Sub
Private Sub RefreshFrameMessageCaption()
    Dim strIDVeicolo As String
    
    strIDVeicolo = GetSelectedAutomezzo()
    
    If strIDVeicolo <> "" Then
        Frame3.Caption = "Invio nuovo messaggio all'automezzo """ & strIDVeicolo & """"
        cmdSendMessage.Enabled = True
    Else
        Frame3.Caption = "Invio nuovo messaggio all'automezzo"
        cmdSendMessage.Enabled = False
    End If

End Sub
Public Sub RefreshAll()
    
    Dim blnDummy As Boolean
    
    blnDummy = SelectRecords()
    blnDummy = SelectRecordsVVP()
    blnDummy = SelectRecordsSEV()
    blnDummy = SelectRecordsSME()
End Sub
Private Sub RefreshNewMessages()
On Error GoTo ErrHandler
    
    Dim SME As CSatellitareMessaggio
    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim lngRowCounter As Long
    Dim lngMessagesNotReadCounter As Long
    
    Set SME = New CSatellitareMessaggio
    
    With SME.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    If Not SME.GetAutomezziMessaggesNotRead(varBuffer) Then
        GoTo ErrHandler
    End If

    TabStrip1.Tabs(3).Caption = "Messaggi"
    
    If Not IsEmpty(varBuffer) Then
        For lngMessagesNotReadCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
            If GetSelectedAMMID() = varBuffer(0, lngMessagesNotReadCounter) Then
                If varBuffer(1, lngMessagesNotReadCounter) > 0 Then
                    TabStrip1.Tabs(3).Caption = "Messaggi (" & varBuffer(1, lngMessagesNotReadCounter) & ")"
                End If
                Exit For
            End If
        Next lngMessagesNotReadCounter
    End If
    
    ' Scansione della griglia
    If Not IsXAEmpty(XA) Then
        For lngRowCounter = XA.LowerBound(1) To XA.UpperBound(1)
            XA.value(lngRowCounter, GRID_VVT_COLUMN_MGSCOUNTS) = ""
            If Not IsEmpty(varBuffer) Then
                For lngMessagesNotReadCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
                    If XA.value(lngRowCounter, GRID_VVT_COLUMN_VVTIDAM1) = varBuffer(0, lngMessagesNotReadCounter) Then
                        XA.value(lngRowCounter, GRID_VVT_COLUMN_MGSCOUNTS) = varBuffer(1, lngMessagesNotReadCounter)
                        Exit For
                    End If
                Next lngMessagesNotReadCounter
            End If
        Next lngRowCounter
        TDBGrid1.RefetchCol (GRID_VVT_COLUMN_MGSCOUNTS)
        TDBGrid1.RefreshCol (GRID_VVT_COLUMN_MGSCOUNTS)
    End If
    
    Set SME = Nothing
    Exit Sub

ErrHandler:
    Set SME = Nothing

End Sub
Private Sub TDBGrid2_FetchCellStyle(ByVal condition As Integer, ByVal Split As Integer, bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid60.StyleDisp)
    If Not IsXAEmpty(XAConsegne) Then
        Dim strStato As String
        Dim strStatoViaggio As String
        Dim blnChangeStyle As Boolean
        
        strStato = IIf(IsNull(XAConsegne.value(bookmark, GRID_VVP_COLUMN_VVPFLSTA)), "0", XAConsegne.value(bookmark, GRID_VVP_COLUMN_VVPFLSTA))
        strStatoViaggio = IIf(IsNull(XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTFLSTA)), "0", XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTFLSTA))
        blnChangeStyle = False
        
        Select Case Col
            Case GRID_VVP_COLUMN_VVCKMPAD
'                If InStr("01", strStato) <> 0 Then
'                    blnChangeStyle = True
'                End If
                
                If InStr("0123", strStatoViaggio) <> 0 Then
'                    CellStyle.ForegroundPicture = ImageList2.ListImages(2).Picture
'                    CellStyle.TransparentForegroundPicture = True
'                    CellStyle.ForegroundPicturePosition = dbgFPLeft
                    blnChangeStyle = True
                End If

            Case GRID_VVP_COLUMN_VVPHRINS
                If InStr("0", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
                
            Case GRID_VVP_COLUMN_TEMPOSCAR
                If InStr("01", strStato) <> 0 Then
                    blnChangeStyle = True
                End If
        End Select
            
        If blnChangeStyle Then
            CellStyle.ForegroundPicture = ImageList2.ListImages(1).Picture
            CellStyle.TransparentForegroundPicture = True
            CellStyle.ForegroundPicturePosition = dbgFPLeft
        End If
    End If
End Sub

Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If Not IsXAEmpty(XAConsegne) Then
        Select Case IIf(IsNull(XAConsegne.value(bookmark, GRID_VVT_COLUMN_VVTFLSTA)) = True, "0", XAConsegne.value(bookmark, GRID_VVT_COLUMN_VVTFLSTA))
            Case "0"
                
            Case "1"
                RowStyle.BackColor = GRID_VVP_ROW_BACKCOLOR1
    
            Case "2"
                RowStyle.BackColor = GRID_VVP_ROW_BACKCOLOR2
        End Select
    End If
End Sub

Private Sub TDBGrid4_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If Not IsXAEmpty(XAMessaggi) Then
        If CStr(XAMessaggi.value(bookmark, GRID_SME_COLUMN_SMEFLREA)) = "N" And _
            XAMessaggi(bookmark, GRID_SME_COLUMN_SMEFLPRO) = "A" Then
            
            RowStyle.BackColor = GRID_SME_ROW_BACKCOLOR1
            RowStyle.ForeColor = GRID_SME_ROW_FORECOLOR1
        End If
    End If
End Sub

Private Sub TDBGrid4_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrHandler
    If LastRow <> TDBGrid4.Row And XAMessaggi.value(LastRow, GRID_SME_COLUMN_SMEFLREA) = "N" Then
        Dim SME As CSatellitareMessaggio
        
        Set SME = New CSatellitareMessaggio
        
        With SME.DS
            Set .DA = DataConnect.DA
            Set .ErrH = ErrH
        End With

        If SME.SetReadMessage(XAMessaggi.value(LastRow, GRID_SME_COLUMN_SMEIDSME)) Then
            SelectRecordsSME
        End If
        Set SME = Nothing
    End If
    
    RefreshFrameMessageCaption
    Exit Sub

ErrHandler:
    Set SME = Nothing
End Sub

Private Sub timerRefresh_Timer()
On Error GoTo ErrHandler

    Dim lngIconIndex As Long
    
    timerRefresh.Enabled = False
    lngRefreshTicks = lngRefreshTicks + 1

    lngIconIndex = (lngRefreshTicks / (lngRefreshSeconds / SECONDS_PER_TICK)) * ImageList3.ListImages.Count
    Image2.Picture = ImageList3.ListImages(lngIconIndex).Picture

    If lngRefreshTicks * SECONDS_PER_TICK >= lngRefreshSeconds Then
        If Not IsXAEmpty(XA) And (TabStrip1.SelectedItem.Index = 2 Or TabStrip1.SelectedItem.Index = 3) And TDBGrid1.SelBookmarks.Count = 0 Then
            RefreshAll
        End If
        StartRefresh
        Exit Sub
    End If

    timerRefresh.Enabled = True
    Exit Sub

ErrHandler:
    timerRefresh.Enabled = True
End Sub
Public Sub StartRefresh()
    If lngRefreshSeconds = 0 Then
        Exit Sub
    End If

    lngRefreshTicks = 0
    timerRefresh.Interval = SECONDS_PER_TICK * 1000
    timerRefresh.Enabled = True

End Sub
Public Sub StopRefresh()
    
    timerRefresh.Enabled = False
    lngRefreshTicks = -1

End Sub
Private Function SendViaggio(varViaggioId As Variant, _
                            ByRef blnViaggioCreated) As Boolean
On Error GoTo ErrHandler
    
    Dim VVT As CViaggio
    Dim objSatError As CError

    blnViaggioCreated = False
    Set VVT = New CViaggio
        
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With

    With VVT
        .DS.Fields("0").FindValue = varViaggioId
    
        If .Find Then

            Set objSatError = New CError
            ' Inserisce o aggiorna i dati di KNOSOS
            If .DS.Fields("47").value = "" Then
                If Not .Insert_Sat(objSatError) Then
                    GoTo ErrHandler
                End If
                blnViaggioCreated = True
            Else
                If Not .Update_Sat(objSatError) Then
                    GoTo ErrHandler
                End If
                blnViaggioCreated = False
            End If
            Set objSatError = Nothing
        End If
    End With
    
    Set VVT = Nothing
    SendViaggio = True
    Exit Function

ErrHandler:
    If Not objSatError Is Nothing Then
        DisplayError objSatError
        Set objSatError = Nothing
    End If
    Set VVT = Nothing
    SendViaggio = False
End Function
Private Sub SendViaggi()
On Error GoTo ErrHandler

    Dim lngRowsCounter As Long
    Dim objXAViaggi As XArrayDB
    Dim lngViaggiUpdatedCount As Long
    Dim lngViaggiCreatedCount As Long
    Dim lngViaggiErrorsCount As Long
    Dim blnViaggioCreated As Boolean

    If IsXAEmpty(XA) Then
        Exit Sub
    End If
    
    If TDBGrid1.SelBookmarks.Count = 0 Then
        If SendViaggio(XA.value(TDBGrid1.bookmark, GRID_VVT_COLUMN_VVTIDVVT), blnViaggioCreated) Then
            DataClass.DS.DisplayMessage "Viaggi satellitare creati: " & IIf(blnViaggioCreated = True, 1, 0) & ", viaggi satellitari aggiornati: " & IIf(blnViaggioCreated = True, 0, 1) & ", numero di errori: 0"
        Else
            DataClass.DS.DisplayMessage "Errore durante la creazione o aggiornamento del viaggio satellitare."
        End If
        Exit Sub
    End If
    
    Set objXAViaggi = New XArrayDB
    
    objXAViaggi.ReDim 0, TDBGrid1.SelBookmarks.Count - 1, 0, 3

    For lngRowsCounter = 0 To TDBGrid1.SelBookmarks.Count - 1
        objXAViaggi.value(lngRowsCounter, 0) = XA.value(lngRowsCounter, GRID_VVT_COLUMN_VVTIDVVT)
        objXAViaggi.value(lngRowsCounter, 1) = XA.value(lngRowsCounter, GRID_VVT_COLUMN_VVTDTINV)
        objXAViaggi.value(lngRowsCounter, 2) = XA.value(lngRowsCounter, GRID_VVT_COLUMN_VVTIDAM1)
        objXAViaggi.value(lngRowsCounter, 3) = XA.value(lngRowsCounter, GRID_VVT_COLUMN_VVTNRSEQ)
    Next lngRowsCounter

    objXAViaggi.QuickSort objXAViaggi.LowerBound(1), objXAViaggi.UpperBound(1), _
                            1, XORDER_ASCEND, XTYPE_DATE, _
                            2, XORDER_ASCEND, XTYPE_STRING, _
                            3, XORDER_ASCEND, XTYPE_STRING

    lngViaggiUpdatedCount = 0
    lngViaggiCreatedCount = 0
    lngViaggiErrorsCount = 0
    For lngRowsCounter = objXAViaggi.LowerBound(1) To objXAViaggi.UpperBound(1)
        If SendViaggio(objXAViaggi.value(lngRowsCounter, 0), blnViaggioCreated) Then
            If blnViaggioCreated Then
                lngViaggiCreatedCount = lngViaggiCreatedCount + 1
            Else
                lngViaggiUpdatedCount = lngViaggiUpdatedCount + 1
            End If
        Else
            lngViaggiErrorsCount = lngViaggiErrorsCount + 1
        End If
    Next lngRowsCounter
    
    DataClass.DS.DisplayMessage "Viaggi satellitare creati: " & lngViaggiCreatedCount & _
                                ", viaggi satellitari aggiornati: " & lngViaggiUpdatedCount & _
                                ", viaggi in errore: " & lngViaggiErrorsCount

    Set objXAViaggi = Nothing
    Exit Sub

ErrHandler:
    Set objXAViaggi = Nothing
End Sub
Private Function UnzoomGrid() As Boolean

    If Not gridZoomGrid Is Nothing Then
        If Not frameZoomGridContainer Is Nothing Then
            Set gridZoomGrid.Container = frameZoomGridContainer
            Set frameZoomGridContainer = Nothing
        Else
            gridZoomGrid.Container = Nothing
        End If
        gridZoomGrid.Left = sngZoomGridLeft
        gridZoomGrid.Top = sngZoomGridTop
        gridZoomGrid.Width = sngZoomGridWidth
        gridZoomGrid.Height = sngZoomGridHeight
        
        frameZoomGrid.Visible = False
        frameMain.Visible = True
        UnzoomGrid = True
        If gridZoomGrid.Visible Then
            gridZoomGrid.SetFocus
        End If
        Set gridZoomGrid = Nothing
    Else
        UnzoomGrid = False
    End If

End Function
Private Function ZoomGrid() As Boolean
    
    UnzoomGrid
    
    If Not TypeOf ActiveControl Is TDBGrid Then
        Exit Function
    End If
    
    If frmVVTSatellitareViewer.ActiveControl Is TDBGrid1 Then
        Set gridZoomGrid = TDBGrid1
    End If
    
    If frmVVTSatellitareViewer.ActiveControl Is TDBGrid2 Then
        Set gridZoomGrid = TDBGrid2
    End If
    
    If frmVVTSatellitareViewer.ActiveControl Is TDBGrid3 Then
        Set gridZoomGrid = TDBGrid3
    End If

    If frmVVTSatellitareViewer.ActiveControl Is TDBGrid4 Then
        Set gridZoomGrid = TDBGrid4
    End If

    If Not gridZoomGrid Is Nothing Then
        sngZoomGridLeft = gridZoomGrid.Left
        sngZoomGridTop = gridZoomGrid.Top
        sngZoomGridWidth = gridZoomGrid.Width
        sngZoomGridHeight = gridZoomGrid.Height
        Set frameZoomGridContainer = gridZoomGrid.Container
        Set gridZoomGrid.Container = frameZoomGrid

        frameMain.Visible = False
        frameZoomGrid.Visible = True
        gridZoomGrid.Visible = True
        gridZoomGrid.Left = 0
        gridZoomGrid.Top = 0
        gridZoomGrid.Width = frameZoomGrid.Width
        gridZoomGrid.Height = frameZoomGrid.Height
        
        ZoomGrid = True
    Else
        ZoomGrid = False
    End If
End Function
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "REFRESH"
            If Not TypeOf ActiveControl Is TDBGrid Then
                Exit Sub
            End If
                
            If frmVVTSatellitareViewer.ActiveControl = TDBGrid1 Then
                SelectRecords
                Exit Sub
            End If
            
            If frmVVTSatellitareViewer.ActiveControl = TDBGrid2 Then
                SelectRecordsVVP
                Exit Sub
            End If
            
            If frmVVTSatellitareViewer.ActiveControl = TDBGrid3 Then
                SelectRecordsSEV
                Exit Sub
            End If
        
            If frmVVTSatellitareViewer.ActiveControl = TDBGrid4 Then
                SelectRecordsSME
                Exit Sub
            End If
        
        Case "ZOOM"
            If gridZoomGrid Is Nothing Then
                ZoomGrid
            Else
                UnzoomGrid
            End If
                
            If gridZoomGrid Is Nothing Then
                Toolbar2.Buttons(2).Image = 5
            Else
                Toolbar2.Buttons(2).Image = 6
            End If
    End Select
End Sub

Sub LoadAutomezzi()
On Error GoTo ErrHandler
    
    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim lngAutomezziCounter As Long

    Const SQL_GET_VEHICLES = "SELECT AAMIDAAM FROM ANAGAAM0 WHERE AAMANNUL = 0 AND AAMFLSAT = 'K' ORDER BY AAMIDAAM"
    
    lngHowMany = 0
    If Not DataClass.DS.DA.QueryCommand(SQL_GET_VEHICLES, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then
        For lngAutomezziCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
            comboAutomezzi.AddItem (varBuffer(0, lngAutomezziCounter))
        Next lngAutomezziCounter
    End If

ErrHandler:
End Sub
