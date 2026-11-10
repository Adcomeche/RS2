VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmEU2KImpExp 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPORTAZIONE / ESPORTAZIONE ANAGRAFICHE EUROSYSTEM"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   9795
   Begin VB.Frame Frame1 
      Height          =   5655
      Index           =   1
      Left            =   120
      TabIndex        =   47
      Top             =   420
      Visible         =   0   'False
      Width           =   9555
      Begin VB.CommandButton cmdBrowseExport 
         Height          =   300
         Index           =   0
         Left            =   9060
         Picture         =   "frmEU2KImpExp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFileNameDestinazioni 
         Height          =   285
         Left            =   1860
         TabIndex        =   56
         Top             =   240
         Width           =   7155
      End
      Begin VB.Frame Frame3 
         Height          =   4935
         Left            =   60
         TabIndex        =   48
         Top             =   660
         Width           =   9435
         Begin VB.Frame frameGridExportData 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   3855
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   9195
            Begin TrueDBGrid60.TDBGrid TDBGridExport 
               Height          =   1455
               Index           =   0
               Left            =   1320
               OleObjectBlob   =   "frmEU2KImpExp.frx":0102
               TabIndex        =   52
               Tag             =   "SELECT"
               Top             =   600
               Width           =   4515
            End
         End
         Begin VB.TextBox txtExportErrorMessage 
            Height          =   285
            Left            =   4560
            TabIndex        =   50
            Top             =   4560
            Width           =   4755
         End
         Begin VB.TextBox txtExportRecordCount 
            Height          =   285
            Left            =   1800
            TabIndex        =   49
            Top             =   4560
            Width           =   975
         End
         Begin MSComctlLib.TabStrip TabStrip3 
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   180
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   661
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            TabMinWidth     =   88
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Destinazioni"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Messaggio di errore"
            Height          =   255
            Left            =   2940
            TabIndex        =   55
            Top             =   4560
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Records da esportare"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   4560
            Width           =   1755
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File Destinazioni"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1635
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
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
            Picture         =   "frmEU2KImpExp.frx":9140
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEU2KImpExp.frx":9252
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEU2KImpExp.frx":9364
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   420
      Visible         =   0   'False
      Width           =   9555
      Begin VB.CheckBox chkImportEnabled 
         Caption         =   "Unità di trasporto"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkImportEnabled 
         Caption         =   "Automezzi"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   960
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox chkImportEnabled 
         Caption         =   "Autisti"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox chkImportEnabled 
         Caption         =   "Vettori"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CommandButton cmdRefreshImport 
         Height          =   300
         Index           =   3
         Left            =   9060
         Picture         =   "frmEU2KImpExp.frx":9476
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Refresh file Unità di Trasporto"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdRefreshImport 
         Height          =   300
         Index           =   2
         Left            =   9060
         Picture         =   "frmEU2KImpExp.frx":9A00
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Refresh file Automezzi"
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdRefreshImport 
         Height          =   300
         Index           =   1
         Left            =   9060
         Picture         =   "frmEU2KImpExp.frx":9F8A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresh file Autisti"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdRefreshImport 
         Height          =   300
         Index           =   0
         Left            =   9060
         Picture         =   "frmEU2KImpExp.frx":A514
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Refresh file Vettori"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFileNameAutomezzi 
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txtFileNameUnitaTrasporto 
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1320
         Width           =   6255
      End
      Begin VB.TextBox txtFileNameAutisti 
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   6255
      End
      Begin VB.TextBox txtFileNameVettori 
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton cmdBrowseImport 
         Height          =   300
         Index           =   3
         Left            =   8700
         Picture         =   "frmEU2KImpExp.frx":AA9E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Selezione file Unità di Trasporto"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseImport 
         Height          =   300
         Index           =   2
         Left            =   8700
         Picture         =   "frmEU2KImpExp.frx":ABA0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Selezione file Automezzi"
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseImport 
         Height          =   300
         Index           =   1
         Left            =   8700
         Picture         =   "frmEU2KImpExp.frx":ACA2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selezione file Autisti"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseImport 
         Height          =   300
         Index           =   0
         Left            =   8700
         Picture         =   "frmEU2KImpExp.frx":ADA4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Selezione file Vettori"
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame4 
         Height          =   3975
         Left            =   60
         TabIndex        =   33
         Tag             =   "txtRecordsCount"
         Top             =   1620
         Width           =   9435
         Begin VB.TextBox txtImportErrorCount 
            Height          =   285
            Left            =   8340
            TabIndex        =   23
            Top             =   3180
            Width           =   975
         End
         Begin VB.TextBox txtImportElaboratedCount 
            Height          =   285
            Left            =   5040
            TabIndex        =   22
            Top             =   3180
            Width           =   975
         End
         Begin VB.TextBox txtImportRecordCount 
            Height          =   285
            Left            =   1620
            TabIndex        =   21
            Top             =   3240
            Width           =   975
         End
         Begin VB.TextBox txtImportErrorMessage 
            Height          =   285
            Left            =   1620
            TabIndex        =   24
            Top             =   3600
            Width           =   7695
         End
         Begin VB.Frame frameGridImportData 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   2475
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   9195
            Begin TrueDBGrid60.TDBGrid TDBGridImport 
               Height          =   1275
               Index           =   2
               Left            =   4320
               OleObjectBlob   =   "frmEU2KImpExp.frx":AEA6
               TabIndex        =   20
               Tag             =   "SELECT"
               Top             =   1920
               Visible         =   0   'False
               Width           =   4875
            End
            Begin TrueDBGrid60.TDBGrid TDBGridImport 
               Height          =   1395
               Index           =   1
               Left            =   0
               OleObjectBlob   =   "frmEU2KImpExp.frx":1D524
               TabIndex        =   19
               Tag             =   "SELECT"
               Top             =   1800
               Visible         =   0   'False
               Width           =   4155
            End
            Begin TrueDBGrid60.TDBGrid TDBGridImport 
               Height          =   1815
               Index           =   0
               Left            =   4320
               OleObjectBlob   =   "frmEU2KImpExp.frx":280BA
               TabIndex        =   18
               Tag             =   "SELECT"
               Top             =   0
               Visible         =   0   'False
               Width           =   4875
            End
            Begin TrueDBGrid60.TDBGrid TDBGridImport 
               Height          =   1695
               Index           =   3
               Left            =   0
               OleObjectBlob   =   "frmEU2KImpExp.frx":348E4
               TabIndex        =   17
               Tag             =   "SELECT"
               Top             =   0
               Visible         =   0   'False
               Width           =   4155
            End
         End
         Begin MSComctlLib.TabStrip TabStrip1 
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   180
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   661
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            TabMinWidth     =   88
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   4
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Vettori"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Autisti"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Automezzo"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Unità di Trasporto"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label Label5 
            Caption         =   "Messaggio di errore"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Records in errore"
            Height          =   255
            Left            =   6600
            TabIndex        =   45
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Records elaborati"
            Height          =   255
            Left            =   3180
            TabIndex        =   44
            Top             =   3240
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Records nel file"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3240
            Width           =   1275
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File"
         Height          =   255
         Index           =   3
         Left            =   1860
         TabIndex        =   42
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   41
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   40
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7860
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEU2KImpExp.frx":41BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEU2KImpExp.frx":4216C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   60
      TabIndex        =   35
      Top             =   6180
      Width           =   9675
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7980
         Picture         =   "frmEU2KImpExp.frx":42706
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmEU2KImpExp.frx":42808
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
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
         Index           =   0
         Left            =   1800
         Picture         =   "frmEU2KImpExp.frx":4290A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
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
         Index           =   6
         Left            =   3540
         Picture         =   "frmEU2KImpExp.frx":42A0C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7860
         Picture         =   "frmEU2KImpExp.frx":42B0E
         TabIndex        =   28
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Esporta"
         Height          =   375
         Left            =   3420
         Picture         =   "frmEU2KImpExp.frx":42C10
         TabIndex        =   27
         Tag             =   "EXPORT"
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Importa"
         Height          =   375
         Left            =   1680
         Picture         =   "frmEU2KImpExp.frx":42D12
         TabIndex        =   26
         Tag             =   "IMPORT"
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmEU2KImpExp.frx":42E14
         TabIndex        =   25
         Tag             =   "BROWSE"
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5460
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt"
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   6075
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   10716
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Importazione"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Esportazione"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   6960
      Width           =   9795
      _ExtentX        =   17277
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
            Object.Width           =   11827
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18.02"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23/03/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEU2KImpExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Public XAAVE As XArrayDB
Public XAAAU As XArrayDB
Public XAAAM As XArrayDB
Public XAAUT As XArrayDB
Public XAADZ As XArrayDB
'Public XAADD As XArrayDB
Public WithEvents DataClass As CDestinazione
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState

Private AVE As CVettore
Private AAU As CAutista
Private AUT As CUnitaDiTrasporto
Private AAM As CAutomezzo
Private ASC As CSocieta
Private PPD As CPrefatturaDettaglio

Private SVE As CVettoreSoc
Private SAU As CAutistaSoc
Private SMP As CPagamentoSoc
Private SIV As CIVASoc
Private SDE As CDepositoSoc
Private SEE As CEnteEmittenteSoc
Private STM As CTipoAutomezzoSoc
Private STS As CTipoCisternaSoc
Private SUT As CUnitaDiTrasportoSoc
Private SSC As CSocietaSoc
Private SCN As CContrattoSoc
Private SCL As CClienteSoc

Private Const EUROSYSTEM_ASC_CODE = "900"

Private lngADZRecordCount As Long

Private strSaveFileName As String
Private lngExportRecordCount(1) As Long
Private strExportErrorMessage(1) As String

Private lngImportRecordCount(4) As Long
Private lngImportErrorCount(4) As Long
Private lngImportElaboratedCount(4) As Long
Private strImportErrorMessage(4) As String
Public objTrasmAnagrafiche As CTrasmAnagrafiche
Public Loading As Boolean

Private Const REGISTRY_KEY_NAME_AVE_FILE_NAME = "ImportAVEFileName"
Private Const REGISTRY_KEY_NAME_AAU_FILE_NAME = "ImportAAUFileName"
Private Const REGISTRY_KEY_NAME_AAM_FILE_NAME = "ImportAAMFileName"
Private Const REGISTRY_KEY_NAME_AUT_FILE_NAME = "ImportAUTFileName"
Private Const REGISTRY_KEY_NAME_ADZ_FILE_NAME = "ExportADZFileName"

Private Const lngDestinazioniIndex = 0

Private Const lngVettoriIndex = 0
Private Const lngAutistiIndex = 1
Private Const lngAutomezziIndex = 2
Private Const lngUnitaTrasportoIndex = 3

'Public DataClass As CMovimentoMagazzinoSelezione
'Public State As FormState

Private Sub LoadRegistryVar()

    txtFileNameVettori.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_AVE_FILE_NAME, "", True)
    txtFileNameAutisti.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_AAU_FILE_NAME, "", True)
    txtFileNameAutomezzi.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_AAM_FILE_NAME, "", True)
    txtFileNameUnitaTrasporto.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_AUT_FILE_NAME, "", True)
    txtFileNameDestinazioni.Text = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_ADZ_FILE_NAME, "", True)

End Sub

Private Sub SaveRegistryVar()
    Dim blnDummy As Boolean
    
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_AVE_FILE_NAME, txtFileNameVettori.Text, True)
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_AAU_FILE_NAME, txtFileNameAutisti.Text, True)
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_AAM_FILE_NAME, txtFileNameAutomezzi.Text, True)
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_AUT_FILE_NAME, txtFileNameUnitaTrasporto.Text, True)
    blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_ADZ_FILE_NAME, txtFileNameDestinazioni.Text, True)

End Sub
Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub

Private Sub RefreshInfo()

'    TabStrip3.Tabs(lngDestinazioniIndex + 1).Image = 0
'    strExportErrorMessage(lngDestinazioniIndex) = ""
'    lngExportRecordCount(lngDestinazioniIndex) = 0
'
'    TabStrip3.Tabs(lngAddebitiIndex + 1).Image = 0
'    strExportErrorMessage(lngAddebitiIndex) = ""
'    lngExportRecordCount(lngAddebitiIndex) = 0
    
    Select Case TabStrip2.SelectedItem.Index
        Case 1  ' Importazioni
            If Len(txtFileNameVettori.Text) > 0 Then
                TabStrip1.Tabs(lngVettoriIndex + 1).Caption = "Vettori (" & XAAVE.Count(1) & ")"
            Else
                TabStrip1.Tabs(lngVettoriIndex + 1).Caption = "Vettori"
            End If
            
            If Len(txtFileNameAutisti.Text) > 0 Then
                TabStrip1.Tabs(lngAutistiIndex + 1).Caption = "Autisti (" & XAAAU.Count(1) & ")"
            Else
                TabStrip1.Tabs(lngAutistiIndex + 1).Caption = "Autisti"
            End If
        
            If Len(txtFileNameAutomezzi.Text) > 0 Then
                TabStrip1.Tabs(lngAutomezziIndex + 1).Caption = "Automezzi (" & XAAAM.Count(1) & ")"
            Else
                TabStrip1.Tabs(lngAutomezziIndex + 1).Caption = "Automezzi"
            End If
        
            If Len(txtFileNameUnitaTrasporto.Text) > 0 Then
                TabStrip1.Tabs(lngUnitaTrasportoIndex + 1).Caption = "Unità di Trasporto (" & XAAUT.Count(1) & ")"
            Else
                TabStrip1.Tabs(lngUnitaTrasportoIndex + 1).Caption = "Unità di Trasporto"
            End If
            
            txtImportRecordCount.Text = lngImportRecordCount(TabStrip1.SelectedItem.Index - 1)
            txtImportErrorCount.Text = lngImportErrorCount(TabStrip1.SelectedItem.Index - 1)
            txtImportElaboratedCount.Text = lngImportElaboratedCount(TabStrip1.SelectedItem.Index - 1)
            txtImportErrorMessage.Text = strImportErrorMessage(TabStrip1.SelectedItem.Index - 1)
        
        Case 2  ' Esportazione

            ' Destinazioni
            If Len(txtFileNameDestinazioni.Text) > 0 Then
                TabStrip3.Tabs(lngDestinazioniIndex + 1).Caption = "Destinazioni (" & XAADZ.Count(1) & ")"
            Else
                TabStrip3.Tabs(lngDestinazioniIndex + 1).Caption = "Destinazioni"
            End If
            
            txtExportRecordCount.Text = lngExportRecordCount(TabStrip3.SelectedItem.Index - 1)
            txtExportErrorMessage.Text = strExportErrorMessage(TabStrip3.SelectedItem.Index - 1)

    End Select

End Sub
Private Sub DisplaySelection()

    Select Case TabStrip2.SelectedItem.Index
        Case 1  ' Importazioni
        
        Case 2  ' Esportazione

            ' Destinazioni
            If lngADZRecordCount = 0 Then
                ResetGrid TDBGridExport(0), XAADZ
            Else
                RefreshXArray XAADZ, TDBGridExport(0), DataClass.DS.List
                TDBGridExport(0).Enabled = True
            End If
            
    End Select

    RefreshInfo
End Sub

Public Sub SetState(NewState As FormState)
    Dim lngHowMany As Long
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            
            
        Case fsSelection
            lngHowMany = DataClass.DS.HowMany
'            Frame1(2).Visible = True
'            Frame1(1).Visible = False

            ' FM 13/08/02 ' Aggiunta Visualizzazione Record Trovati
            Select Case lngHowMany
                Case 1
                    DataClass.DS.DisplayMessage lngHowMany & " Documento selezionato"
                Case Else
                    DataClass.DS.DisplayMessage lngHowMany & " Documenti selezionati"
            End Select
            
            If Not Loading Then
'                If TDBgrid1.Enabled Then
'                    TDBgrid1.SetFocus
'                End If
            End If
    End Select
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim strErrorMessage As String
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
'                SetState fsSelection
                'AutoSizeGrid TDBgrid1
            End If

        Case fcImport
            Import

        Case fcExport
            Export

        Case fcOpen
'            With CommonDialog1
'                .fileName = txtFileName
'                .DialogTitle = "Seleziona il file per la trasmissione... "
'                .InitDir = APP.Path
'                .Filter = "Tutti i files (*.*)|*.*| "
'                .FilterIndex = 1
'                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
'                .ShowSave
'                txtFileName = .fileName
'            End With
        
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub chkExportEnabled_Click(Index As Integer)
    SelectRecords
End Sub

Private Sub chkImportEnabled_Click(Index As Integer)
    RefreshImportFile CLng(Index)
End Sub

Private Sub cmdBrowseExport_Click(Index As Integer)
    With CommonDialog1
    
        Select Case Index
            Case 0
                CommonDialog1.DialogTitle = "Seleziona il file delle destinazioni."
                .fileName = txtFileNameDestinazioni.Text
            
        End Select
        
        .DefaultExt = ".txt"
        .Filter = "Text file (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
        .ShowOpen
        If Not .CancelError Then

            Select Case Index
                Case 0
                    txtFileNameDestinazioni.Text = .fileName
                
            End Select
        End If
    End With
End Sub

Private Function RefreshImportFile(lngIndex As Long) As Boolean
    
    Dim strErrorMessage As String
    
    Select Case lngIndex
        
        Case lngVettoriIndex
            RefreshImportFile = LoadFileVettori(txtFileNameVettori, strErrorMessage)
            TabStrip1.Tabs(lngIndex + 1).Selected = True
            TDBGridImport(lngVettoriIndex).Enabled = True
        
        Case lngAutistiIndex
            RefreshImportFile = LoadFileAutisti(txtFileNameAutisti, strErrorMessage)
            TabStrip1.Tabs(lngIndex + 1).Selected = True
            TDBGridImport(lngAutistiIndex).Enabled = True

        Case lngAutomezziIndex
            RefreshImportFile = LoadFileAutomezzi(txtFileNameAutomezzi, strErrorMessage)
            TabStrip1.Tabs(lngIndex + 1).Selected = True
            TDBGridImport(lngAutomezziIndex).Enabled = True

        Case lngUnitaTrasportoIndex
            RefreshImportFile = LoadFileUnitaTrasporto(txtFileNameUnitaTrasporto, strErrorMessage)
            TabStrip1.Tabs(lngIndex + 1).Selected = True
            TDBGridImport(lngUnitaTrasportoIndex).Enabled = True

    End Select
    TabStrip1.Tabs(lngIndex + 1).Image = 0
End Function

Private Sub cmdBrowseImport_Click(Index As Integer)

    Dim strErrorMessage As String

    With CommonDialog1
    
        Select Case Index
            Case lngVettoriIndex
                CommonDialog1.DialogTitle = "Seleziona il file dei vettori."
                .fileName = txtFileNameVettori.Text
            
            Case lngAutistiIndex
                CommonDialog1.DialogTitle = "Seleziona il file degli autisti."
                .fileName = txtFileNameAutisti.Text
            
            Case lngAutomezziIndex
                CommonDialog1.DialogTitle = "Seleziona il file degli automezzi."
                .fileName = txtFileNameAutomezzi.Text
            
            Case lngUnitaTrasportoIndex
                CommonDialog1.DialogTitle = "Seleziona il file delle unità di trasporto."
                .fileName = txtFileNameUnitaTrasporto.Text
    
        End Select
        
        .DefaultExt = ".txt"
        .Filter = "Text file (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
        .ShowOpen
        If Not .CancelError Then
    
            Select Case Index
                Case lngVettoriIndex
                    txtFileNameVettori.Text = .fileName
                
                Case lngAutistiIndex
                    txtFileNameAutisti.Text = .fileName
                
                Case lngAutomezziIndex
                    txtFileNameAutomezzi.Text = .fileName
                
                Case lngUnitaTrasportoIndex
                    txtFileNameUnitaTrasporto.Text = .fileName
        
            End Select
            
            RefreshImportFile CLng(Index)
        End If
    End With

End Sub

Private Sub cmdRefreshImport_Click(Index As Integer)
    RefreshImportFile CLng(Index)
End Sub

Private Sub ErrH_OnError()
'    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    
    Dim lngCounter As Long
    Dim strErrorMessage As String

    Loading = True
    
    Set DataClass = New CDestinazione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .Fields("20").Required = True
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set XAAVE = New XArrayDB
    Set XAAAU = New XArrayDB
    Set XAAAM = New XArrayDB
    Set XAAUT = New XArrayDB
    Set XAADZ = New XArrayDB
    Set objTrasmAnagrafiche = New CTrasmAnagrafiche
    
    Set AVE = New CVettore
    With AVE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    Set PPD = New CPrefatturaDettaglio
    With PPD.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    Set SCL = New CClienteSoc
    With SCL.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SVE = New CVettoreSoc
    With SVE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SAU = New CAutistaSoc
    With SAU.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SMP = New CPagamentoSoc
    With SMP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SIV = New CIVASoc
    With SIV.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SDE = New CDepositoSoc
    With SDE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    Set SEE = New CEnteEmittenteSoc
    With SEE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set STM = New CTipoAutomezzoSoc
    With STM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    Set STS = New CTipoCisternaSoc
    With STS.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SUT = New CUnitaDiTrasportoSoc
    With SUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set SSC = New CSocietaSoc
    With SSC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    Set SCN = New CContrattoSoc
    With SCN.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
    Set ASC = New CSocieta
    With ASC.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With

    TabStrip2.Tabs(1).Selected = True
    
    For lngCounter = 0 To TDBGridImport.Count - 1
        TDBGridImport(lngCounter).Left = 0
        TDBGridImport(lngCounter).Top = 0
        TDBGridImport(lngCounter).Height = frameGridImportData.Height
        TDBGridImport(lngCounter).Width = frameGridImportData.Width
    Next lngCounter
    TDBGridImport.Item(TabStrip1.SelectedItem.Index - 1).Visible = True
    
    For lngCounter = 0 To TDBGridExport.Count - 1
        TDBGridExport(lngCounter).Left = 0
        TDBGridExport(lngCounter).Top = 0
        TDBGridExport(lngCounter).Height = frameGridExportData.Height
        TDBGridExport(lngCounter).Width = frameGridExportData.Width
    Next lngCounter
    TDBGridExport.Item(TabStrip3.SelectedItem.Index - 1).Visible = True
    
    LoadRegistryVar
'    Combo1.ListIndex = 0
    
    Loading = False

    SetState fsparaminput

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegistryVar
    
    Set XAAVE = Nothing
    Set XAAAU = Nothing
    Set XAAAM = Nothing
    Set XAAUT = Nothing
    Set XAADZ = Nothing
    Set AVE = Nothing
    Set AAU = Nothing
    Set AUT = Nothing
    Set AAM = Nothing
    Set PPD = Nothing
    Set SCL = Nothing
    Set SVE = Nothing
    Set SAU = Nothing
    Set SMP = Nothing
    Set SIV = Nothing
    Set SDE = Nothing
    Set SEE = Nothing
    Set STM = Nothing
    Set STS = Nothing
    Set SUT = Nothing
    Set SSC = Nothing
    Set SCN = Nothing
    Set DataClass = Nothing
    Set objTrasmAnagrafiche = Nothing
    Set frmEU2KImpExp = Nothing
End Sub

Private Sub ResetGrid(Grid As TDBGrid, XA As XArrayDB)
    
    If XA.Count(1) > 0 Then
        XA.DeleteRows XA.LowerBound(1), XA.Count(1)
    End If
    Set Grid.Array = XA
    Grid.ReBind

End Sub

Private Sub ResetXA(XA As XArrayDB)
    
    If XA.Count(1) > 0 Then
        XA.DeleteRows XA.LowerBound(1), XA.Count(1)
    End If

End Sub

Private Function LoadFileVettori(strFileName As String, strErrorMessage As String)
    
    If chkImportEnabled(lngVettoriIndex).value And Len(strFileName) > 0 Then
        If LoadFile(strFileName, TDBGridImport(lngVettoriIndex), XAAVE, objTrasmAnagrafiche.TRVettore, strImportErrorMessage(lngVettoriIndex)) Then
            txtFileNameVettori.Text = strFileName
            LoadFileVettori = True
        Else
            LoadFileVettori = False
            ResetGrid TDBGridImport(lngVettoriIndex), XAAVE
            strErrorMessage = strImportErrorMessage(lngVettoriIndex)
        End If
    Else
        ResetGrid TDBGridImport(lngVettoriIndex), XAAVE
        LoadFileVettori = True
    End If
    lngImportRecordCount(lngVettoriIndex) = XAAVE.Count(1)
    lngImportErrorCount(lngVettoriIndex) = 0
    lngImportElaboratedCount(lngVettoriIndex) = 0
    RefreshInfo

End Function

Private Function LoadFileAutisti(strFileName As String, strErrorMessage As String)

    If chkImportEnabled(lngAutistiIndex).value And Len(strFileName) > 0 Then
        If LoadFile(strFileName, TDBGridImport(lngAutistiIndex), XAAAU, objTrasmAnagrafiche.TRAutista, strImportErrorMessage(lngAutistiIndex)) Then
            txtFileNameAutisti.Text = strFileName
            LoadFileAutisti = True
        Else
            strErrorMessage = strImportErrorMessage(lngAutistiIndex)
            ResetGrid TDBGridImport(lngAutistiIndex), XAAAU
            LoadFileAutisti = False
        End If
    Else
        ResetGrid TDBGridImport(lngAutistiIndex), XAAAU
        LoadFileAutisti = True
    End If
    RefreshInfo
    lngImportRecordCount(lngAutistiIndex) = XAAAU.Count(1)
    lngImportErrorCount(lngAutistiIndex) = 0
    lngImportElaboratedCount(lngAutistiIndex) = 0

End Function

Private Function LoadFileAutomezzi(strFileName As String, strErrorMessage As String)
        
    If chkImportEnabled(lngAutomezziIndex).value And Len(strFileName) > 0 Then
        If LoadFile(strFileName, TDBGridImport(lngAutomezziIndex), XAAAM, objTrasmAnagrafiche.TRAutomezzo, strImportErrorMessage(lngAutomezziIndex)) Then
            txtFileNameAutomezzi.Text = strFileName
            LoadFileAutomezzi = True
        Else
            strErrorMessage = strImportErrorMessage(lngAutomezziIndex)
            ResetGrid TDBGridImport(lngAutomezziIndex), XAAAM
            LoadFileAutomezzi = False
        End If
    Else
        ResetGrid TDBGridImport(lngAutomezziIndex), XAAAM
        LoadFileAutomezzi = True
    End If
    lngImportRecordCount(lngAutomezziIndex) = XAAAM.Count(1)
    lngImportErrorCount(lngAutomezziIndex) = 0
    lngImportElaboratedCount(lngAutomezziIndex) = 0
    RefreshInfo

End Function

Private Function LoadFileUnitaTrasporto(strFileName As String, strErrorMessage As String)
        
    If chkImportEnabled(lngUnitaTrasportoIndex).value And Len(strFileName) > 0 Then
        If LoadFile(strFileName, TDBGridImport(lngUnitaTrasportoIndex), XAAUT, objTrasmAnagrafiche.TRUnitaTrasporto, strImportErrorMessage(lngUnitaTrasportoIndex)) Then
            txtFileNameUnitaTrasporto.Text = strFileName
            LoadFileUnitaTrasporto = True
        Else
            strErrorMessage = strImportErrorMessage(lngUnitaTrasportoIndex)
            ResetGrid TDBGridImport(lngUnitaTrasportoIndex), XAAUT
            LoadFileUnitaTrasporto = False
        End If
    Else
        ResetGrid TDBGridImport(lngUnitaTrasportoIndex), XAAUT
        LoadFileUnitaTrasporto = True
    End If
    lngImportRecordCount(lngUnitaTrasportoIndex) = XAAUT.Count(1)
    lngImportErrorCount(lngUnitaTrasportoIndex) = 0
    lngImportElaboratedCount(lngUnitaTrasportoIndex) = 0
    RefreshInfo

End Function


Private Sub TabStrip3_Click()
    Dim lngCounter As Long
    For lngCounter = 0 To TDBGridExport.Count - 1
        TDBGridExport(lngCounter).Visible = (lngCounter = TabStrip3.SelectedItem.Index - 1)
    Next lngCounter
    RefreshInfo
End Sub

Private Sub TDBGridImport_FetchRowStyle(Index As Integer, ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
On Error GoTo ErrHandler

    Dim XA As XArrayDB

    Select Case Index

        Case lngVettoriIndex
            Set XA = XAAVE

        Case lngAutistiIndex
            Set XA = XAAAU

        Case lngAutomezziIndex
            Set XA = XAAAM

        Case lngUnitaTrasportoIndex
            Set XA = XAAUT

    End Select

    If Not IsNull(bookmark) And XA.Count(1) > 0 Then

        Select Case XA.value(bookmark, XA.Count(2) - 2)
            Case "E"
                RowStyle.BackColor = &HB9B9B9    ' &HE1C6D1

            Case "I", "U"
                RowStyle.BackColor = &HA3E4BB

        End Select

    End If

ErrHandler:
End Sub

Private Sub txtFileNameAutisti_GotFocus()
    strSaveFileName = txtFileNameAutisti.Text
End Sub

Private Sub txtFileNameAutisti_LostFocus()
    If strSaveFileName <> txtFileNameAutisti.Text Then
        RefreshImportFile lngAutistiIndex
        If Len(strImportErrorMessage(lngAutistiIndex)) > 0 Then
            MsgBox strImportErrorMessage(lngAutistiIndex), vbCritical Or vbOKOnly, "Errore caricamento Unità di Autisti!"
        End If
    End If
End Sub

Private Sub txtFileNameAutomezzi_GotFocus()
    strSaveFileName = txtFileNameAutomezzi.Text
End Sub

Private Sub txtFileNameAutomezzi_LostFocus()
    If strSaveFileName <> txtFileNameAutomezzi.Text Then
        RefreshImportFile lngAutomezziIndex
        If Len(strImportErrorMessage(lngAutomezziIndex)) > 0 Then
            MsgBox strImportErrorMessage(lngAutomezziIndex), vbCritical Or vbOKOnly, "Errore caricamento Unità di Automezzi!"
        End If
    End If
End Sub

Private Sub txtFileNameUnitaTrasporto_GotFocus()
    strSaveFileName = txtFileNameUnitaTrasporto.Text
End Sub

Private Sub txtFileNameUnitaTrasporto_LostFocus()
    If strSaveFileName <> txtFileNameUnitaTrasporto.Text Then
        RefreshImportFile lngUnitaTrasportoIndex
        If Len(strImportErrorMessage(lngUnitaTrasportoIndex)) > 0 Then
            MsgBox strImportErrorMessage(lngUnitaTrasportoIndex), vbCritical Or vbOKOnly, "Errore caricamento Unità di Trasporto!"
        End If
    End If
End Sub

Private Sub txtFileNameVettori_GotFocus()
    strSaveFileName = txtFileNameVettori.Text
End Sub

Private Sub txtFileNameVettori_LostFocus()
    If strSaveFileName <> txtFileNameVettori.Text Then
        RefreshImportFile lngVettoriIndex
        If Len(strImportErrorMessage(lngVettoriIndex)) > 0 Then
            MsgBox strImportErrorMessage(lngVettoriIndex), vbCritical Or vbOKOnly, "Errore caricamento Vettori!"
        End If
    End If
End Sub

Private Sub TabStrip1_Click()
    Dim lngCounter As Long
    For lngCounter = 0 To TDBGridImport.Count - 1
        TDBGridImport(lngCounter).Visible = (lngCounter = TabStrip1.SelectedItem.Index - 1)
    Next lngCounter
    RefreshInfo
End Sub

Private Function LoadFile(strFileName As String, _
                            Grid As TDBGrid, _
                            XA As XArrayDB, _
                            TextFile As CTextFile, _
                            Optional strErrorMessage As String) As Boolean
On Error GoTo ErrHandler

    Dim strTextLine As String
    Dim lngCurrentRow As Long
    Dim lngFieldCounter  As Long
    Dim strErrorMsg As String
    Dim lngMaxLength As Long
    Dim varBuffer As Variant

    strErrorMessage = ""
    
    lngMaxLength = 0
    For lngFieldCounter = 1 To TextFile.Count
        lngMaxLength = lngMaxLength + TextFile.Item(lngFieldCounter).Size
    Next lngFieldCounter
    
    If Not TextFile.OpenFile(strFileName, 0) Then
        strErrorMsg = "Impossibile aprire il file " & strFileName & "!"
        GoTo ErrHandler
    End If

    XA.Clear
    lngCurrentRow = 0
    Do
        If TextFile.EOF Then
            Exit Do
        End If
        strTextLine = TextFile.GetLine(False)

'        If Len(strTextLine) <> lngMaxLength Then
'            strErrorMsg = "Errore alla riga """ & lngCurrentRow & """: lunghezza record non valida!"
'            GoTo ErrHandler
'        End If

'        varBuffer = TextFile.ParseString(strTextLine, ";")
'        If IsEmpty(varBuffer) Then
'            strErrorMsg = "Errore alla riga """ & lngCurrentRow & """: errore nel parsing dei dati!"
'            GoTo ErrHandler
'        End If
        
        If Not TextFile.ParseFixedLenFields(strTextLine, 1, TextFile.Count) Then
            strErrorMsg = "Errore alla riga """ & lngCurrentRow & """: errore nel parsing dei dati!"
            GoTo ErrHandler
        End If

        XA.ReDim 0, lngCurrentRow, 0, TextFile.Count + 1
        
        For lngFieldCounter = 0 To TextFile.Count - 1
'            XA.value(lngCurrentRow, lngFieldCounter) = RTrim(varBuffer(lngFieldCounter))
            XA.value(lngCurrentRow, lngFieldCounter) = TextFile.Item(lngFieldCounter + 1).value
        Next lngFieldCounter
        XA.value(lngCurrentRow, TextFile.Count) = "*"
        XA.value(lngCurrentRow, TextFile.Count + 1) = "Da elaborare"
        lngCurrentRow = lngCurrentRow + 1
    Loop

    If lngCurrentRow > 0 Then
        Set Grid.Array = XA
    Else
        ResetGrid Grid, XA
    End If
    Grid.ReBind
    
    TextFile.CloseFile
    LoadFile = True
    Exit Function
ErrHandler:
    If Len(strErrorMsg) = 0 Then
        strErrorMsg = "Errore sconociuto durante il caricamento del file!"
    End If
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = strErrorMsg
    End If
    TextFile.CloseFile
    
    LoadFile = False
End Function

Private Sub Import()
    
    Dim strErrorMessage As String
    Dim lngCounter As Long

    TabStrip1.Tabs(1).Image = 0
    TabStrip1.Tabs(2).Image = 0
    TabStrip1.Tabs(3).Image = 0
    TabStrip1.Tabs(4).Image = 0
    
    If Not ImportVettori(strImportErrorMessage(lngVettoriIndex), _
        lngImportRecordCount(lngVettoriIndex), _
        lngImportErrorCount(lngVettoriIndex), _
        lngImportElaboratedCount(lngVettoriIndex)) Then
        
        TabStrip1.Tabs(lngVettoriIndex + 1).Image = 2
    Else
        TabStrip1.Tabs(lngVettoriIndex + 1).Image = IIf(lngImportErrorCount(lngVettoriIndex) > 0, 1, 0)
    End If

    If Not ImportAutisti(strImportErrorMessage(lngAutistiIndex), _
        lngImportRecordCount(lngAutistiIndex), _
        lngImportErrorCount(lngAutistiIndex), _
        lngImportElaboratedCount(lngAutistiIndex)) Then
        
        TabStrip1.Tabs(lngAutistiIndex + 1).Image = 2
    Else
        TabStrip1.Tabs(lngAutistiIndex + 1).Image = IIf(lngImportErrorCount(lngAutistiIndex) > 0, 1, 0)
    End If

    If Not ImportAutomezzi(strImportErrorMessage(lngAutomezziIndex), _
        lngImportRecordCount(lngAutomezziIndex), _
        lngImportErrorCount(lngAutomezziIndex), _
        lngImportElaboratedCount(lngAutomezziIndex)) Then
        
        TabStrip1.Tabs(lngAutomezziIndex + 1).Image = 2
    Else
        TabStrip1.Tabs(lngAutomezziIndex + 1).Image = IIf(lngImportErrorCount(lngAutomezziIndex) > 0, 1, 0)
    End If

    If Not ImportUnitaTrasporto(strImportErrorMessage(lngUnitaTrasportoIndex), _
        lngImportRecordCount(lngUnitaTrasportoIndex), _
        lngImportErrorCount(lngUnitaTrasportoIndex), _
        lngImportElaboratedCount(lngUnitaTrasportoIndex)) Then
        
        TabStrip1.Tabs(lngUnitaTrasportoIndex + 1).Image = 2
    Else
        TabStrip1.Tabs(lngUnitaTrasportoIndex + 1).Image = IIf(lngImportErrorCount(lngUnitaTrasportoIndex) > 0, 1, 0)
    End If
    
    RefreshInfo

End Sub

Private Function ImportVettori(ByRef strErrorMessage As String, _
        ByRef lngRecordCount As Long, _
        ByRef lngErrorCount As Long, _
        ByRef lngElaboratedCount As Long) As Boolean

On Error GoTo ErrHandler
    
    Dim lngCounter As Long
    Dim blnIsNewRecord As Boolean
    Dim strCodiceAIV As String
    Dim strCodiceAMP As String
    Dim strIDAVE As String
    Dim blnRecordError As Boolean
    Dim XA As XArrayDB

    Me.MousePointer = vbHourglass
    
    Set XA = XAAVE

    If Len(txtFileNameVettori.Text) = 0 Then
        GoTo OKHandler
    End If

    If Not LoadFileVettori(txtFileNameVettori.Text, strErrorMessage) Then
        GoTo ErrHandler
    End If

    strErrorMessage = ""
    lngRecordCount = XA.Count(1)
    lngErrorCount = 0
    lngElaboratedCount = 0
    
    If XA.Count(1) = 0 Then
        GoTo OKHandler
    End If

    ASC.DS.Fields("0").FindValue = ID_SOC
    ASC.Find
    
    For lngCounter = 0 To XA.Count(1) - 1

        blnRecordError = False
        DataConnect.DA.BeginTransaction
        
        ' Controllo Codici x societa
        Do
            ' Modalita' di Pagamento
            If Not TranslateMultiSoc(strCodiceAMP, XA.value(lngCounter, 10), EUROSYSTEM_ASC_CODE, SMP) Then
                SetXAMessage XA, lngCounter, "E", "Modalità di pagamento " & XA.value(lngCounter, 10) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
            
            ' IVA
            If Len(XA.value(lngCounter, 11)) = 0 Then
                strCodiceAIV = ASC.DS.Fields("9").value
            Else
                If Not TranslateMultiSoc(strCodiceAIV, XA.value(lngCounter, 11), EUROSYSTEM_ASC_CODE, SIV) Then
                    SetXAMessage XA, lngCounter, "E", "IVA " & XA.value(lngCounter, 11) & " non trovata!"
                    blnRecordError = True
                    Exit Do
                End If
            End If
            
            Exit Do
        Loop
        
        If Not blnRecordError Then
            
            blnIsNewRecord = Not TranslateMultiSoc(strIDAVE, XA.value(lngCounter, 0), EUROSYSTEM_ASC_CODE, SVE)

            With AVE
            
                .DS.Fields.ResetFields
                
                If Not blnIsNewRecord Then
                    .DS.Fields("0").FindValue = strIDAVE
        
                    If Not .Find Then
                        strErrorMessage = formatErrorMsg(.DS.DA.ErrH)
                        GoTo ErrHandler
                    End If
                    
                    If .DS.HowMany > 0 Then
                        blnIsNewRecord = False
                    Else
                        blnIsNewRecord = True
                    End If
                End If
                
                If blnIsNewRecord Then
                    strIDAVE = ""
                End If

                ' Ragione Sociale
                .DS.Fields("1").NewValue = XA.value(lngCounter, 1)
                .DS.Fields("1").Updated = True
    
                ' Ragione Sociale aggiuntiva
                .DS.Fields("2").NewValue = XA.value(lngCounter, 2)
                .DS.Fields("2").Updated = True
                
                ' Ragione Sociale abbreviata
                .DS.Fields("3").NewValue = XA.value(lngCounter, 3)
                .DS.Fields("3").Updated = True
        
                ' Indirizzo
                .DS.Fields("4").NewValue = XA.value(lngCounter, 4)
                .DS.Fields("4").Updated = True
        
                ' Località
                .DS.Fields("5").NewValue = XA.value(lngCounter, 5)
                .DS.Fields("5").Updated = True
        
                ' Provincia
                If Len(XA.value(lngCounter, 6)) > .DS.Fields("6").MaxLength Then
                    .DS.Fields("6").NewValue = Left(XA.value(lngCounter, 6), .DS.Fields("6").MaxLength)
                Else
                    .DS.Fields("6").NewValue = XA.value(lngCounter, 6)
                End If
                .DS.Fields("6").Updated = True
        
                ' CAP
                .DS.Fields("7").NewValue = XA.value(lngCounter, 7)
                .DS.Fields("7").Updated = True
        
                ' Codice Fiscale
                .DS.Fields("8").NewValue = XA.value(lngCounter, 8)
                .DS.Fields("8").Updated = True
        
                ' Partita IVA
                .DS.Fields("9").NewValue = XA.value(lngCounter, 9)
                .DS.Fields("9").Updated = True
    
                ' Modalità di Pagamento
                .DS.Fields("10").NewValue = strCodiceAMP
                .DS.Fields("10").Updated = True
    
                ' IVA
                .DS.Fields("11").NewValue = strCodiceAIV
                .DS.Fields("11").Updated = True
        
                ' Telefono
                .DS.Fields("12").NewValue = XA.value(lngCounter, 12)
                .DS.Fields("12").Updated = True
        
                ' Nota 1
                .DS.Fields("13").NewValue = XA.value(lngCounter, 13)
                .DS.Fields("13").Updated = True
        
                ' Nota 2
                .DS.Fields("14").NewValue = XA.value(lngCounter, 14)
                .DS.Fields("14").Updated = True
        
                ' Annullato
                If Len(XA.value(lngCounter, 15)) > 0 And XA.value(lngCounter, 15) = "A" Then
                    .DS.Fields("15").NewValue = 1
                Else
                    .DS.Fields("15").NewValue = 0
                End If
                .DS.Fields("15").Updated = True
                
                If blnIsNewRecord Then
                    
                    If Not .GetNewKey(strIDAVE) Then
                        SetXAMessage XA, lngCounter, "E", "Errore nella generazione della nuova chiave."
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        ' Codice
                        .DS.Fields("0").NewValue = strIDAVE
                        .DS.Fields("0").Updated = True
                    
                        If Not .Insert Then
                            SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                            lngErrorCount = lngErrorCount + 1
                            blnRecordError = True
                        Else
                            SetXAMessage XA, lngCounter, "I", "Inserito Vettore '" & strIDAVE & "'."
                            
                            With SVE
                                .DS.Fields.ResetFields
                                .DS.Fields("Codice").NewValue = strIDAVE
                                .DS.Fields("Codice").Updated = True
                                .DS.Fields("Società").NewValue = EUROSYSTEM_ASC_CODE
                                .DS.Fields("Società").Updated = True
                                .DS.Fields("CodPerSoc").NewValue = XA.value(lngCounter, 0)
                                .DS.Fields("CodPerSoc").Updated = True
                                .DS.Fields("Annullato").NewValue = 0
                                .DS.Fields("Annullato").Updated = True
                            
                                If Not .Insert() Then
                                    SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                                    lngErrorCount = lngErrorCount + 1
                                    blnRecordError = True
                                End If
                            End With
                        
                        End If
                    End If
                Else
                    If Not .BeginSave(True) Then
                        SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        SetXAMessage XA, lngCounter, "U", "Aggiornato Vettore '" & strIDAVE & "'."
                        .DS.DA.EndTrans 0
                    End If
                End If
            End With
        Else
            lngErrorCount = lngErrorCount + 1
            blnRecordError = True
        End If
        
        If blnRecordError Then
            DataConnect.DA.EndTrans 1
        Else
            DataConnect.DA.EndTrans 0
        End If
        
        lngElaboratedCount = lngElaboratedCount + 1
    
    Next lngCounter
    
OKHandler:
    Set XA = Nothing
    Me.MousePointer = vbDefault
    ImportVettori = True
    Exit Function

ErrHandler:
    If Len(strErrorMessage) = 0 Then
        strErrorMessage = "Errore sconosciuto durante l'importazione dei Vettori!"
    End If
    Set XA = Nothing
    ImportVettori = False
    Me.MousePointer = vbDefault
End Function

Private Function formatErrorMsg(objError As CError) As String
    
    Dim lngCounter As Long
    
    For lngCounter = 1 To objError.Descriptions.Count
        If Len(formatErrorMsg) > 0 Then
            formatErrorMsg = formatErrorMsg & " - "
        End If
        formatErrorMsg = formatErrorMsg & objError.Descriptions(lngCounter).Description
    Next
End Function

Private Function ImportAutisti(ByRef strErrorMessage As String, _
        ByRef lngRecordCount As Long, _
        ByRef lngErrorCount As Long, _
        ByRef lngElaboratedCount As Long) As Boolean
On Error GoTo ErrHandler

    Dim lngCounter As Long
    Dim blnIsNewRecord As Boolean
    Dim strIDAAU As String
    Dim strCodiceAVE As String
    Dim blnRecordError As Boolean
    Dim XA As XArrayDB

    Me.MousePointer = vbHourglass
    
    Set XA = XAAAU
    
    If Len(txtFileNameAutisti.Text) = 0 Then
        GoTo OKHandler
    End If
    
    If Not LoadFileAutisti(txtFileNameAutisti.Text, strErrorMessage) Then
        GoTo ErrHandler
    End If
    
    strErrorMessage = ""
    lngRecordCount = XA.Count(1)
    lngErrorCount = 0
    lngElaboratedCount = 0
    
    If XA.Count(1) = 0 Then
        GoTo OKHandler
    End If

    For lngCounter = 0 To XA.Count(1) - 1

        DataConnect.DA.BeginTransaction
        blnRecordError = False
        
        ' Controllo Codici x societa
        Do
            ' Vettore
            If IsNull(XA.value(lngCounter, 3)) Or Len(XA.value(lngCounter, 3)) = 0 Then
                strCodiceAVE = "001"
            Else
                If Not TranslateMultiSoc(strCodiceAVE, XA.value(lngCounter, 3), EUROSYSTEM_ASC_CODE, SVE) Then
                    SetXAMessage XA, lngCounter, "E", "Vettore " & XA.value(lngCounter, 3) & " non trovato!"
                    blnRecordError = True
                    Exit Do
                End If
            End If
            
            Exit Do
        Loop
        
        If Not blnRecordError Then
            
            blnIsNewRecord = Not TranslateMultiSoc(strIDAAU, XA.value(lngCounter, 0), EUROSYSTEM_ASC_CODE, SAU)
            
            With AAU
                
                .DS.Fields.ResetFields
                
                If Not blnIsNewRecord Then
                    .DS.Fields("0").FindValue = strIDAAU
        
                    If Not .Find Then
                        strErrorMessage = formatErrorMsg(.DS.DA.ErrH)
                        GoTo ErrHandler
                    End If
                    
                    If .DS.HowMany > 0 Then
                        blnIsNewRecord = False
                    Else
                        blnIsNewRecord = True
                    End If
                End If
                
                If blnIsNewRecord Then
                    strIDAAU = ""
                End If
                
                ' Nome
                .DS.Fields("1").NewValue = XA.value(lngCounter, 1)
                .DS.Fields("1").Updated = True
            
                ' Nome Abbreviato
                If Len(XA.value(lngCounter, 2)) > .DS.Fields("2").MaxLength Then
                    .DS.Fields("2").NewValue = Left(XA.value(lngCounter, 2), .DS.Fields("2").MaxLength)
                Else
                    .DS.Fields("2").NewValue = XA.value(lngCounter, 2)
                End If
                .DS.Fields("2").Updated = True
            
                ' Vettore
                .DS.Fields("3").NewValue = strCodiceAVE
                .DS.Fields("3").Updated = True
            
                ' Numero Patente
                If Len(XA.value(lngCounter, 4)) > .DS.Fields("4").MaxLength Then
                    .DS.Fields("4").NewValue = Left(XA.value(lngCounter, 4), .DS.Fields("4").MaxLength)
                Else
                    .DS.Fields("4").NewValue = XA.value(lngCounter, 4)
                End If
                .DS.Fields("4").Updated = True
            
                ' Numero NIF
                .DS.Fields("9").NewValue = LTrim(XA.value(lngCounter, 5))
                .DS.Fields("9").Updated = True
            
                ' Disponibilità
                If Len(XA.value(lngCounter, 6)) > 0 And XA.value(lngCounter, 6) = "1" Then
                    .DS.Fields("5").NewValue = "S"
                Else
                    .DS.Fields("5").NewValue = "N"
                End If
                .DS.Fields("5").Updated = True

                ' Nota 1
                .DS.Fields("6").NewValue = XA.value(lngCounter, 7)
                .DS.Fields("6").Updated = True
            
                ' Nota 2
                .DS.Fields("7").NewValue = XA.value(lngCounter, 8)
                .DS.Fields("7").Updated = True
            
                ' Annullato
                If Len(XA.value(lngCounter, 9)) > 0 And XA.value(lngCounter, 9) = "A" Then
                    .DS.Fields("8").NewValue = 1
                Else
                    .DS.Fields("8").NewValue = 0
                End If
                .DS.Fields("8").Updated = True
    
                If blnIsNewRecord Then
                    
                    ' Nuova Chiave
                    If Not .GetNewKey(strIDAAU) Then
                        SetXAMessage XA, lngCounter, "E", "Errore nella generazione della nuova chiave."
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        ' Codice
                        .DS.Fields("0").NewValue = strIDAAU
                        .DS.Fields("0").Updated = True
                        
                        If Not .Insert Then
                            SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                            lngErrorCount = lngErrorCount + 1
                            blnRecordError = True
                        Else
                            SetXAMessage XA, lngCounter, "I", "Inserito Autista '" & strIDAAU & "'."
                            
                            ' Inserisce il cod. multisoc
                            With SAU
                                .DS.Fields.ResetFields
                                .DS.Fields("Codice").NewValue = strIDAAU
                                .DS.Fields("Codice").Updated = True
                                .DS.Fields("Società").NewValue = EUROSYSTEM_ASC_CODE
                                .DS.Fields("Società").Updated = True
                                .DS.Fields("CodPerSoc").NewValue = XA.value(lngCounter, 0)
                                .DS.Fields("CodPerSoc").Updated = True
                                .DS.Fields("Annullato").NewValue = 0
                                .DS.Fields("Annullato").Updated = True
    
                                If Not .Insert() Then
                                    SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                                    lngErrorCount = lngErrorCount + 1
                                    blnRecordError = True
                                End If
                            End With
                        End If
                    
                    End If
                Else
                    If Not .BeginSave(True) Then
                        SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        SetXAMessage XA, lngCounter, "U", "Aggiornato Autista '" & strIDAAU & "'."
                    End If
                End If
                
            End With
        Else
            lngErrorCount = lngErrorCount + 1
            blnRecordError = True
        End If

        If blnRecordError Then
            DataConnect.DA.EndTrans 1
        Else
            DataConnect.DA.EndTrans 0
        End If

        lngElaboratedCount = lngElaboratedCount + 1

    Next lngCounter
    
OKHandler:
    Me.MousePointer = vbDefault
    Set XA = Nothing
    ImportAutisti = True
    Exit Function

ErrHandler:
    DataConnect.DA.EndTrans 1
    Set XA = Nothing
    If Len(strErrorMessage) = 0 Then
        strErrorMessage = "Errore sconosciuto durante l'importazione degli Autisti!"
    End If
    Me.MousePointer = vbDefault
    ImportAutisti = False
End Function

Function FieldGetNumericValue(varValue As Variant) As Double
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(CStr(varValue)) = 0 Then
        FieldGetNumericValue = 0
    Else
        FieldGetNumericValue = CDbl(varValue)
    End If
End Function

Private Function ImportAutomezzi(ByRef strErrorMessage As String, _
        ByRef lngRecordCount As Long, _
        ByRef lngErrorCount As Long, _
        ByRef lngElaboratedCount As Long) As Boolean
On Error GoTo ErrHandler

    Dim lngCounter As Long
    Dim blnIsNewRecord As Boolean
    Dim strCodiceATM As String
    Dim strCodiceATS As String
    Dim blnRecordError As Boolean
    Dim XA As XArrayDB

    Me.MousePointer = vbHourglass
    
    Set XA = XAAAM

    If Len(txtFileNameAutomezzi.Text) = 0 Then
        GoTo OKHandler
    End If
    
    If Not LoadFileAutomezzi(txtFileNameAutomezzi.Text, strErrorMessage) Then
        GoTo ErrHandler
    End If
    
    strErrorMessage = ""
    lngRecordCount = XA.Count(1)
    lngErrorCount = 0
    lngElaboratedCount = 0
    
    If XA.Count(1) = 0 Then
        GoTo OKHandler
    End If

    For lngCounter = 0 To XA.Count(1) - 1

        blnRecordError = False
        DataConnect.DA.BeginTransaction
        
        ' Controllo Codici x societa
        Do
            ' Tipo Automezzo
            If Not TranslateMultiSoc(strCodiceATM, XA.value(lngCounter, 3), EUROSYSTEM_ASC_CODE, STM) Then
                SetXAMessage XA, lngCounter, "E", "Tipo Automezzo " & XA.value(lngCounter, 3) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If

            ' Tipo Cisterna
            If Not TranslateMultiSoc(strCodiceATS, XA.value(lngCounter, 26), EUROSYSTEM_ASC_CODE, STS) Then
                SetXAMessage XA, lngCounter, "E", "Tipo Cisterna " & XA.value(lngCounter, 26) & " non trovata!"
                blnRecordError = True
                Exit Do
            End If
            
            Exit Do
        Loop

        If Not blnRecordError Then
            
            With AAM
                .DS.Fields.ResetFields
                .DS.Fields("0").FindValue = XA.value(lngCounter, 0)
    
                If Not .Find Then
                    strErrorMessage = formatErrorMsg(.DS.DA.ErrH)
                    GoTo ErrHandler
                End If
                
                If .DS.HowMany > 0 Then
                    blnIsNewRecord = False
                Else
                    blnIsNewRecord = True
                End If
    
                ' Targa
                .DS.Fields("0").NewValue = XA.value(lngCounter, 0)
                .DS.Fields("0").Updated = True
            
                ' Motrice/Rimorchio
                Select Case XA.value(lngCounter, 1)
                    Case "T"
                        .DS.Fields("1").NewValue = "M"
                        
                    Case Else
                        .DS.Fields("1").NewValue = XA.value(lngCounter, 1)
                End Select
                .DS.Fields("1").Updated = True
            
                ' Descrizione
                .DS.Fields("2").NewValue = XA.value(lngCounter, 2)
                .DS.Fields("2").Updated = True
            
                ' Tipo Automezzo
                .DS.Fields("3").NewValue = strCodiceATM
                .DS.Fields("3").Updated = True
            
                ' Disponibilità
                If Len(XA.value(lngCounter, 4)) > 0 And XA.value(lngCounter, 4) = "1" Then
                    .DS.Fields("4").NewValue = "S"
                Else
                    .DS.Fields("4").NewValue = "N"
                End If
                .DS.Fields("4").Updated = True

                ' Capacità Scomparto 1
                .DS.Fields("5").NewValue = FieldGetNumericValue(XA.value(lngCounter, 5))
                .DS.Fields("5").Updated = True

                ' Capacità Scomparto 2
                .DS.Fields("7").NewValue = FieldGetNumericValue(XA.value(lngCounter, 6))
                .DS.Fields("7").Updated = True
                
                ' Capacità Scomparto 3
                .DS.Fields("9").NewValue = FieldGetNumericValue(XA.value(lngCounter, 7))
                .DS.Fields("9").Updated = True
                
                ' Capacità Scomparto 4
                .DS.Fields("11").NewValue = FieldGetNumericValue(XA.value(lngCounter, 8))
                .DS.Fields("11").Updated = True
                
                ' Capacità Scomparto 5
                .DS.Fields("13").NewValue = FieldGetNumericValue(XA.value(lngCounter, 9))
                .DS.Fields("13").Updated = True
                
                ' Capacità Scomparto 6
                .DS.Fields("15").NewValue = FieldGetNumericValue(XA.value(lngCounter, 10))
                .DS.Fields("15").Updated = True
                
                ' Capacità Scomparto 7
                .DS.Fields("17").NewValue = FieldGetNumericValue(XA.value(lngCounter, 11))
                .DS.Fields("17").Updated = True
                
                ' Capacità Scomparto 8
                .DS.Fields("19").NewValue = FieldGetNumericValue(XA.value(lngCounter, 12))
                .DS.Fields("19").Updated = True
                
                ' Capacità Scomparto 9
                .DS.Fields("21").NewValue = FieldGetNumericValue(XA.value(lngCounter, 13))
                .DS.Fields("21").Updated = True
    
                ' Capacità Scomparto 10
                .DS.Fields("23").NewValue = FieldGetNumericValue(XA.value(lngCounter, 14))
                .DS.Fields("23").Updated = True
    
                ' Numero Scomparti 1
                .DS.Fields("6").NewValue = FieldGetNumericValue(XA.value(lngCounter, 15))
                .DS.Fields("6").Updated = True
                
                ' Numero Scomparti 2
                .DS.Fields("8").NewValue = FieldGetNumericValue(XA.value(lngCounter, 16))
                .DS.Fields("8").Updated = True
                
                ' Numero Scomparti 3
                .DS.Fields("10").NewValue = FieldGetNumericValue(XA.value(lngCounter, 17))
                .DS.Fields("10").Updated = True
                
                ' Numero Scomparti 4
                .DS.Fields("12").NewValue = FieldGetNumericValue(XA.value(lngCounter, 18))
                .DS.Fields("12").Updated = True
                
                ' Numero Scomparti 5
                .DS.Fields("14").NewValue = FieldGetNumericValue(XA.value(lngCounter, 19))
                .DS.Fields("14").Updated = True
                
                ' Numero Scomparti 6
                .DS.Fields("16").NewValue = FieldGetNumericValue(XA.value(lngCounter, 20))
                .DS.Fields("16").Updated = True
                
                ' Numero Scomparti 7
                .DS.Fields("18").NewValue = FieldGetNumericValue(XA.value(lngCounter, 21))
                .DS.Fields("18").Updated = True
                
                ' Numero Scomparti 8
                .DS.Fields("20").NewValue = FieldGetNumericValue(XA.value(lngCounter, 22))
                .DS.Fields("20").Updated = True
                
                ' Numero Scomparti 9
                .DS.Fields("22").NewValue = FieldGetNumericValue(XA.value(lngCounter, 23))
                .DS.Fields("22").Updated = True
                
                ' Numero Scomparti 10
                .DS.Fields("24").NewValue = FieldGetNumericValue(XA.value(lngCounter, 24))
                .DS.Fields("24").Updated = True
                
                ' Portata
                .DS.Fields("25").NewValue = FieldGetNumericValue(XA.value(lngCounter, 25))
                .DS.Fields("25").Updated = True
            
                ' Tipo Cisterna
                .DS.Fields("27").NewValue = strCodiceATS
                .DS.Fields("27").Updated = True
            
                ' Tara Pesata
                .DS.Fields("28").NewValue = FieldGetNumericValue(XA.value(lngCounter, 27))
                .DS.Fields("28").Updated = True
            
                ' Tara Libretto
                .DS.Fields("29").NewValue = FieldGetNumericValue(XA.value(lngCounter, 28))
                .DS.Fields("29").Updated = True
            
                ' Portata rimorchiabile
                .DS.Fields("30").NewValue = FieldGetNumericValue(XA.value(lngCounter, 29))
                .DS.Fields("30").Updated = True
            
                ' Cubatura costruttore
                .DS.Fields("31").NewValue = FieldGetNumericValue(XA.value(lngCounter, 30))
                .DS.Fields("31").Updated = True
            
                ' Tipo Carica
                .DS.Fields("32").NewValue = XA.value(lngCounter, 31)
                .DS.Fields("32").Updated = True

                ' Satellitare installato
                If XA.value(lngCounter, 32) = "1" Then
                    .DS.Fields("34").NewValue = "K"
                Else
                    .DS.Fields("34").NewValue = "N"
                End If
                .DS.Fields("34").Updated = True
            
                ' Annullato
                If Len(XA.value(lngCounter, 33)) > 0 And XA.value(lngCounter, 33) = "A" Then
                    .DS.Fields("26").NewValue = 1
                Else
                    .DS.Fields("26").NewValue = 0
                End If
                .DS.Fields("26").Updated = True
    
                If blnIsNewRecord Then
                    If Not .Insert Then
                        SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        SetXAMessage XA, lngCounter, "I", "Inserito."
                    End If
                Else
                    If Not .BeginSave(True) Then
                        SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        SetXAMessage XA, lngCounter, "U", "Aggiornato."
                    End If
                End If
            End With
        Else
            lngErrorCount = lngErrorCount + 1
            blnRecordError = True
        End If
        
        If blnRecordError Then
            DataConnect.DA.EndTrans 1
        Else
            DataConnect.DA.EndTrans 0
        End If
        
        lngElaboratedCount = lngElaboratedCount + 1

    Next lngCounter
    
OKHandler:
    Me.MousePointer = vbDefault
    Set XA = Nothing
    ImportAutomezzi = True
    Exit Function

ErrHandler:
    If Len(strErrorMessage) = 0 Then
        strErrorMessage = "Errore sconosciuto durante l'importazione degli Automezzi!"
    End If
    Me.MousePointer = vbDefault
    ImportAutomezzi = False
End Function

Private Function TranslateMultiSoc(ByRef strCodice As String, _
            strCodMultiSoc As String, _
            strIDSocieta As String, _
            objDataClass As Object) As Boolean
On Error GoTo ErrHandler

    Dim strSQL As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strTableName As String
    
    strCodice = ""
    If Len(strCodMultiSoc) > 0 Then
        With objDataClass
    
            strTableName = Mid(.DS.Fields("Codice").fieldName, 1, InStr(.DS.Fields("Codice").fieldName, ".") - 1)

            strSQL = "SELECT " & .DS.Fields("Codice").fieldName & " FROM " & strTableName & _
                " WHERE " & .DS.Fields("Società").fieldName & " = ?" & _
                " AND " & .DS.Fields("CodPerSoc").fieldName & " = ?"

            If Not .DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, strIDSocieta, strCodMultiSoc) Then
                GoTo ErrHandler
            End If

            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
            
            strCodice = varBuffer(0, 0)
        End With
    End If
    
    TranslateMultiSoc = True
    Exit Function
        
ErrHandler:
    TranslateMultiSoc = False
    strCodice = ""
End Function

Private Sub SetXAMessage(XA As XArrayDB, lngIndex As Long, strType As String, strMessage As String)
    XA.value(lngIndex, XA.Count(2) - 2) = strType
    XA.value(lngIndex, XA.Count(2) - 1) = strMessage
End Sub

Private Function ImportUnitaTrasporto(ByRef strErrorMessage As String, _
        ByRef lngRecordCount As Long, _
        ByRef lngErrorCount As Long, _
        ByRef lngElaboratedCount As Long) As Boolean
On Error GoTo ErrHandler

    Dim lngCounter As Long
    Dim blnIsNewRecord As Boolean
    Dim blnRecordError As Boolean
    Dim strIDAUT As String
    Dim strCodiceAAU1 As String
    Dim strCodiceAAU2 As String
    Dim strCodiceAVE As String
    Dim strCodiceAVET As String
    Dim strCodiceADE As String
    Dim strCodiceAEE As String
    Dim strCodiceASC As String
    Dim strCodiceACN As String
    Dim XA As XArrayDB

    Me.MousePointer = vbHourglass
    
    Set XA = XAAUT

    If Len(txtFileNameUnitaTrasporto.Text) = 0 Then
        GoTo OKHandler
    End If

    If Not LoadFileUnitaTrasporto(txtFileNameUnitaTrasporto.Text, strErrorMessage) Then
        GoTo ErrHandler
    End If

    strErrorMessage = ""
    lngRecordCount = XA.Count(1)
    lngErrorCount = 0
    lngElaboratedCount = 0
    
    If XA.Count(1) = 0 Then
        GoTo OKHandler
    End If

    For lngCounter = 0 To XA.Count(1) - 1

        blnRecordError = False
        DataConnect.DA.BeginTransaction

        ' Controllo Codici x societa
        Do
            ' Autista 1
            If Not TranslateMultiSoc(strCodiceAAU1, XA.value(lngCounter, 4), EUROSYSTEM_ASC_CODE, SAU) Then
                SetXAMessage XA, lngCounter, "E", "Autista " & XA.value(lngCounter, 4) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
            
            ' Autista 2
            If Not TranslateMultiSoc(strCodiceAAU2, XA.value(lngCounter, 5), EUROSYSTEM_ASC_CODE, SAU) Then
                SetXAMessage XA, lngCounter, "E", "Autista " & XA.value(lngCounter, 5) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
            
            ' Vettore
            If Not TranslateMultiSoc(strCodiceAVE, XA.value(lngCounter, 6), EUROSYSTEM_ASC_CODE, SVE) Then
                SetXAMessage XA, lngCounter, "E", "Vettore " & XA.value(lngCounter, 6) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
    
            ' Vettore Trasportatore
            If Len(XA.value(lngCounter, 7)) = 0 Then
                strCodiceAVET = strCodiceAVE
            Else
                If Not TranslateMultiSoc(strCodiceAVET, XA.value(lngCounter, 7), EUROSYSTEM_ASC_CODE, SVE) Then
                    SetXAMessage XA, lngCounter, "E", "Vettore " & XA.value(lngCounter, 7) & " non trovato!"
                    blnRecordError = True
                    Exit Do
                End If
            End If
            
            ' Deposito
            If Not TranslateMultiSoc(strCodiceADE, XA.value(lngCounter, 10), EUROSYSTEM_ASC_CODE, SDE) Then
                SetXAMessage XA, lngCounter, "E", "Deposito " & XA.value(lngCounter, 10) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
    
            ' Ente Emittente
            If Not TranslateMultiSoc(strCodiceAEE, XA.value(lngCounter, 11), EUROSYSTEM_ASC_CODE, SEE) Then
                SetXAMessage XA, lngCounter, "E", "Ente Emittente " & XA.value(lngCounter, 11) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
    
            ' Società
            If Not TranslateMultiSoc(strCodiceASC, XA.value(lngCounter, 12), EUROSYSTEM_ASC_CODE, SSC) Then
                SetXAMessage XA, lngCounter, "E", "Società " & XA.value(lngCounter, 12) & " non trovata!"
                blnRecordError = True
                Exit Do
            End If
            
            ' Contratto
            If Not TranslateMultiSoc(strCodiceACN, XA.value(lngCounter, 15), EUROSYSTEM_ASC_CODE, SCN) Then
                SetXAMessage XA, lngCounter, "E", "Contratto " & XA.value(lngCounter, 15) & " non trovato!"
                blnRecordError = True
                Exit Do
            End If
            
            Exit Do
        Loop
    
        If Not blnRecordError Then
            
            blnIsNewRecord = Not TranslateMultiSoc(strIDAUT, XA.value(lngCounter, 0), EUROSYSTEM_ASC_CODE, SUT)
            
            With AUT
            
                .DS.Fields.ResetFields
                
                If Not blnIsNewRecord Then
                    .DS.Fields("0").FindValue = strIDAUT
        
                    If Not .Find Then
                        strErrorMessage = formatErrorMsg(.DS.DA.ErrH.LastError)
                        GoTo ErrHandler
                    End If
                    
                    If .DS.HowMany > 0 Then
                        blnIsNewRecord = False
                    Else
                        blnIsNewRecord = True
                    End If
                End If
                
                If blnIsNewRecord Then
                    strIDAUT = ""
                End If
                
                ' Descrizione
                .DS.Fields("1").NewValue = XA.value(lngCounter, 1)
                .DS.Fields("1").Updated = True
            
                ' Motrice
                .DS.Fields("2").NewValue = XA.value(lngCounter, 2)
                .DS.Fields("2").Updated = True
            
                ' Rimorchio
                .DS.Fields("3").NewValue = XA.value(lngCounter, 3)
                .DS.Fields("3").Updated = True
    
                ' Autista 1
                .DS.Fields("4").NewValue = strCodiceAAU1
                .DS.Fields("4").Updated = True
            
                ' Autista 2
                .DS.Fields("5").NewValue = strCodiceAAU2
                .DS.Fields("5").Updated = True
            
                ' Vettore
                .DS.Fields("6").NewValue = strCodiceAVE
                .DS.Fields("6").Updated = True
            
                ' Vettore Trasportatore
                .DS.Fields("7").NewValue = strCodiceAVET
                .DS.Fields("7").Updated = True
            
                ' Portata Totale
                .DS.Fields("8").NewValue = FieldGetNumericValue(XA.value(lngCounter, 8))
                .DS.Fields("8").Updated = True
            
                ' Disponibilità
                If Len(XA.value(lngCounter, 9)) > 0 And XA.value(lngCounter, 9) = "1" Then
                    .DS.Fields("9").NewValue = "S"
                Else
                    .DS.Fields("9").NewValue = "N"
                End If
                .DS.Fields("9").Updated = True
            
                ' Deposito
                .DS.Fields("10").NewValue = strCodiceADE
                .DS.Fields("10").Updated = True
            
                ' Ente Emittente
                .DS.Fields("11").NewValue = XA.value(lngCounter, 11)
                .DS.Fields("11").Updated = True
            
                ' Società
                .DS.Fields("12").NewValue = strCodiceASC
                .DS.Fields("12").Updated = True
    
                ' Sociale/Terzo
                Select Case XA.value(lngCounter, 13)
                    Case "1"
                        .DS.Fields("13").NewValue = "S"
                        .DS.Fields("13").Updated = True
                        
                    Case "2"
                        .DS.Fields("13").NewValue = "T"
                        .DS.Fields("13").Updated = True
                End Select

                ' Badge
                .DS.Fields("14").NewValue = XA.value(lngCounter, 14)
                .DS.Fields("14").Updated = True
            
                ' Contratto
                .DS.Fields("15").NewValue = strCodiceACN
                .DS.Fields("15").Updated = True

                ' Badge
                .DS.Fields("17").NewValue = XA.value(lngCounter, 16)
                .DS.Fields("17").Updated = True
            
                ' Annullato
                If Len(XA.value(lngCounter, 17)) > 0 And XA.value(lngCounter, 17) = "A" Then
                    .DS.Fields("16").NewValue = 1
                Else
                    .DS.Fields("16").NewValue = 0
                End If
                .DS.Fields("16").Updated = True
    
                If blnIsNewRecord Then
                    
                    If Not .GetNewKey(strIDAUT) Then
                        SetXAMessage XA, lngCounter, "E", "Errore nella generazione della nuova chiave."
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        ' Codice
                        .DS.Fields("0").NewValue = strIDAUT
                        .DS.Fields("0").Updated = True
                        
                        If Not .Insert Then
                            SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                            lngErrorCount = lngErrorCount + 1
                            blnRecordError = True
                        Else
                            SetXAMessage XA, lngCounter, "I", "Inserita Unità di Trasporto '" & strIDAUT & "'."
                        
                            With SUT
                                .DS.Fields.ResetFields
                                .DS.Fields("Codice").NewValue = strIDAUT
                                .DS.Fields("Codice").Updated = True
                                .DS.Fields("Società").NewValue = EUROSYSTEM_ASC_CODE
                                .DS.Fields("Società").Updated = True
                                .DS.Fields("CodPerSoc").NewValue = XA.value(lngCounter, 0)
                                .DS.Fields("CodPerSoc").Updated = True
                                .DS.Fields("Annullato").NewValue = 0
                                .DS.Fields("Annullato").Updated = True
                            
                                If Not .Insert() Then
                                    SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                                    lngErrorCount = lngErrorCount + 1
                                    blnRecordError = True
                                End If
                            End With
                        End If
                    End If
                Else
                    If Not .BeginSave(True) Then
                        SetXAMessage XA, lngCounter, "E", formatErrorMsg(.DS.ErrH.LastError)
                        lngErrorCount = lngErrorCount + 1
                        blnRecordError = True
                    Else
                        SetXAMessage XA, lngCounter, "U", "Aggiornata Unità di Trasporto '" & strIDAUT & "'."
                        .DS.DA.EndTrans 0
                    End If
                End If
            End With
        Else
            lngErrorCount = lngErrorCount + 1
            blnRecordError = True
        End If
        
        If blnRecordError Then
            DataConnect.DA.EndTrans 1
        Else
            DataConnect.DA.EndTrans 0
        End If
        
        lngElaboratedCount = lngElaboratedCount + 1

    Next lngCounter
    
OKHandler:
    Set XA = Nothing
    Me.MousePointer = vbDefault
    ImportUnitaTrasporto = True
    Exit Function

ErrHandler:
    If Len(strErrorMessage) = 0 Then
        strErrorMessage = "Errore sconosciuto durante l'importazione delle Unità di Trasporto!"
    End If
    Set XA = Nothing
    Me.MousePointer = vbDefault
    ImportUnitaTrasporto = False
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    ReDim varOrderFields(0) As Variant

    Me.MousePointer = vbHourglass
    
    lngADZRecordCount = 0

    If Not DataClass.DS.Fields.Validate Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    If Not DataClass.DS.ScanClientCond Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If

    ' Carica Destinazioni
    varOrderFields(0) = "0"
    If Not DataClass.GetSelectionExport(True, varOrderFields) Then
        GoTo ErrHandler
    End If
    lngADZRecordCount = DataClass.DS.HowMany
    
    DisplaySelection
    GoTo EndProc

ErrHandler:
    Me.MousePointer = vbDefault
    SelectRecords = False
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function

Private Sub TabStrip2_Click()
    Dim lngCounter As Long
    For lngCounter = 0 To Frame1.Count - 1
        Frame1(lngCounter).Visible = (lngCounter = TabStrip2.SelectedItem.Index - 1)
    Next lngCounter
End Sub

Private Sub Export()

    Dim strErrorMessage As String
    
    If Len(txtFileNameDestinazioni.Text) > 0 Then
        
        If Not ExportDestinazioni(txtFileNameDestinazioni.Text, _
            lngExportRecordCount(lngDestinazioniIndex), _
            strExportErrorMessage(lngDestinazioniIndex)) Then
            TabStrip3.Tabs(lngDestinazioniIndex + 1).Image = 1
            MsgBox "Errore durante l'esportazione delle destinazioni.", vbOKOnly Or vbCritical, "Errore"
        Else
            TabStrip3.Tabs(lngDestinazioniIndex + 1).Image = 0
            MsgBox "L'esportazione delle destinazioni si è conclusa correttamente.", vbOKOnly, "Messaggio"
        End If
    
    End If

    RefreshInfo
End Sub

Private Function ExportDestinazioni(strFileName As String, _
                    ByRef lngRecordCount As Long, _
                    ByRef strErrorMessage As String) As Boolean
On Error GoTo ErrHandler

    Dim lngCounter As Long
    Dim TextFile As CTextFile

    Set TextFile = objTrasmAnagrafiche.TRDestinazione

    Me.MousePointer = vbHourglass
    
    strErrorMessage = ""
    lngRecordCount = 0

    If XAADZ.Count(1) = 0 Then
        GoTo OKHandler
    End If
    
    If Not TextFile.OpenFile(strFileName, 1) Then
        strErrorMessage = "Impossibile aprire il file " & strFileName & "!"
        GoTo ErrHandler
    End If
    
    For lngCounter = 0 To XAADZ.Count(1) - 1
        
        TextFile.Item("1").value = XAADZ.value(lngCounter, 0)
        TextFile.Item("2").value = XAADZ.value(lngCounter, 1)
        TextFile.Item("3").value = XAADZ.value(lngCounter, 2)
        TextFile.Item("4").value = XAADZ.value(lngCounter, 3)
        TextFile.Item("5").value = XAADZ.value(lngCounter, 4)
        TextFile.Item("6").value = XAADZ.value(lngCounter, 5)
    
        If Not TextFile.WriteLine() Then
            strErrorMessage = "Errore durante la scrittura del file " & strFileName & "!"
            GoTo ErrHandler
        End If

        lngRecordCount = lngRecordCount + 1
    Next lngCounter

OKHandler:
    TextFile.CloseFile
    
    Me.MousePointer = vbDefault
    Set TextFile = Nothing
    ExportDestinazioni = True
    Exit Function

ErrHandler:
    TextFile.CloseFile

    If Len(strErrorMessage) = 0 Then
        strErrorMessage = "Errore sconosciuto durante l'esportazione delle destinazioni!"
    End If
    
    Me.MousePointer = vbDefault
    Set TextFile = Nothing
    ExportDestinazioni = False
End Function

Private Function GetCodMultiSoc(ByRef strCodSoc As String, strID As String, strIDSoc As String, objDataClass As Object) As Boolean

    strCodSoc = ""

    If Len(strID) > 0 Then
        With objDataClass
            .DS.Fields(.DS.KeyFields(1)).FindValue = strID
            .DS.Fields(.DS.KeyFields(2)).FindValue = strIDSoc
            If Not .Find Then
                GetCodMultiSoc = False
                Exit Function
            End If
    
            If .DS.HowMany = 0 Then
                GetCodMultiSoc = False
                Exit Function
            End If
            
            strCodSoc = .DS.Fields("CodPerSoc").value
        End With
    End If

    GetCodMultiSoc = True

End Function

Private Function E2KFormatNumber(varValue As Variant, lngSize As Long, Optional lngDecimal As Long = 0) As String

    Dim dblValue As Double
    Dim strFormat As String
    Dim dblDecimal As Double

    dblDecimal = CDbl("1" & String(lngDecimal, "0"))
    If IsEmpty(varValue) Or IsNull(varValue) Then
        dblValue = 0
    Else
        If IsNumeric(varValue) Then
            dblValue = CDbl(varValue)
        Else
            dblValue = 0
        End If
    End If
    
    strFormat = String(lngSize, "0")

    E2KFormatNumber = Format(dblValue * dblDecimal, strFormat)

End Function

