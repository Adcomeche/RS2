VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2ACN0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTRATTI"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   9855
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   9675
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   5340
         Picture         =   "OIL2ACN0.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   780
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
         Picture         =   "OIL2ACN0.frx":0532
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
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
         Picture         =   "OIL2ACN0.frx":0634
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   795
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
         Left            =   7980
         Picture         =   "OIL2ACN0.frx":0736
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
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
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2ACN0.frx":0838
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
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
         Picture         =   "OIL2ACN0.frx":093A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
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
         Picture         =   "OIL2ACN0.frx":0A3C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7860
         Picture         =   "OIL2ACN0.frx":0B3E
         TabIndex        =   21
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteAddebito 
         Caption         =   "Elimina Riga Dettaglio"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2ACN0.frx":0C40
         TabIndex        =   20
         Tag             =   "CUSTOM1"
         Top             =   720
         Width           =   2475
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Riga Dettaglio"
         Height          =   375
         Left            =   2280
         Picture         =   "OIL2ACN0.frx":0D42
         TabIndex        =   19
         Tag             =   "EDITLINK"
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F6 - Aggiungi Riga"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ACN0.frx":0E44
         TabIndex        =   18
         Tag             =   "ADDLINK"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2ACN0.frx":0F46
         TabIndex        =   17
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2ACN0.frx":1048
         TabIndex        =   16
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ACN0.frx":114A
         TabIndex        =   15
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   6885
      Width           =   9855
      _ExtentX        =   17383
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
            Object.Width           =   11933
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11.45"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "07/03/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   4200
      TabIndex        =   22
      Top             =   0
      Width           =   5595
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   3900
         TabIndex        =   12
         Tag             =   "9"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "10"
         Top             =   3480
         Width           =   1935
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
               Picture         =   "OIL2ACN0.frx":124C
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ACN0.frx":135E
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ACN0.frx":1470
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1260
         TabIndex        =   2
         Tag             =   "0"
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1260
         TabIndex        =   5
         Tag             =   "3"
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1260
         TabIndex        =   3
         Tag             =   "1"
         Top             =   960
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Tag             =   "2"
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1260
         TabIndex        =   6
         Tag             =   "4"
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   1260
         TabIndex        =   7
         Tag             =   "5"
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   1260
         TabIndex        =   8
         Tag             =   "6"
         Top             =   2400
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
         Left            =   1260
         TabIndex        =   9
         Tag             =   "7"
         Top             =   2760
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
         Left            =   1260
         TabIndex        =   10
         Tag             =   "8"
         Top             =   3120
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
         Left            =   1260
         TabIndex        =   1
         Tag             =   "11"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   1260
         TabIndex        =   13
         Tag             =   "12"
         Top             =   3900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   4560
         Tag             =   "12"
         Top             =   3900
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
         Left            =   1860
         TabIndex        =   47
         Tag             =   "12"
         Top             =   3900
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mod. fattura"
         Height          =   255
         Index           =   10
         Left            =   300
         TabIndex        =   46
         Top             =   3900
         Width           =   855
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
         Left            =   1860
         TabIndex        =   43
         Tag             =   "11"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente"
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   4620
         Tag             =   "11"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   41
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   1980
         Tag             =   "0"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5220
         Tag             =   "3"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   2340
         Tag             =   "1"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Numero"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   39
         Top             =   960
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   5220
         Tag             =   "2"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N&umero Sec."
         Height          =   255
         Index           =   3
         Left            =   3060
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   5220
         Tag             =   "4"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N&ote"
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   37
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   2820
         Tag             =   "5"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Km a contratto"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   36
         Top             =   2040
         Width           =   1095
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
         Left            =   1860
         TabIndex        =   35
         Tag             =   "6"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Tipo contratto"
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   34
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   4620
         Tag             =   "6"
         Top             =   2400
         Width           =   240
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
         Left            =   1860
         TabIndex        =   33
         Tag             =   "7"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Società"
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   32
         Top             =   2760
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   4620
         Tag             =   "7"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   4620
         Tag             =   "8"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Vettore"
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   31
         Top             =   3120
         Width           =   855
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
         Left            =   1860
         TabIndex        =   30
         Tag             =   "8"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   34
         Left            =   3300
         Tag             =   "10"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Tari&ffa"
         Height          =   255
         Index           =   34
         Left            =   300
         TabIndex        =   29
         Top             =   3480
         Width           =   855
      End
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   4155
      Left            =   120
      OleObjectBlob   =   "OIL2ACN0.frx":1582
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin TrueDBGrid60.TDBGrid TDBGDettaglio 
      Height          =   1275
      Left            =   120
      OleObjectBlob   =   "OIL2ACN0.frx":4568
      TabIndex        =   14
      Top             =   4320
      Width           =   9675
   End
End
Attribute VB_Name = "OIL2ACN0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 11
Private XA As XArrayDB
Public XAACD As XArrayDB
Public WithEvents DataClass As CContratto
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CContrattoDettaglio
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private ColTags As Variant

Private Sub ClearGridDettaglio()
    ClearXArray XAACD, TDBGDettaglio
End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    DoEvents
    If Not IsNull(DBGrid.bookmark) Then
        If DBGrid.bookmark <= XA.UpperBound(1) Then
            GetSelectedID = XA.value(DBGrid.bookmark, 0)
        End If
    End If
End Function

Private Sub TrackLostFocus(Optional mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(2) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsInsert) Then
'                    .Fields("0").FindValue = MaskEdBox1.Text
'                    .Fields("1").FindValue = mskField.Text
                    If GetRecord(MaskEdBox7.Text, mskField1.Text) Then
                        SetState fsEdit
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ClearForm()
    DataClass.DS.ClearClient
    ClearGridDettaglio
End Sub

Private Sub DisplayDetail()
    
    Dim varOrderFields As Variant
    
    With DetailClass
        ReDim varOrderFields(2) As Variant
        varOrderFields(0) = .DS.KeyFields(1)
        varOrderFields(1) = .DS.KeyFields(2)
        varOrderFields(2) = .DS.KeyFields(3)
        
        .DS.Field("Cliente").FindValue = DataClass.DS.Field("Cliente").value
        .DS.Field("Contratto").FindValue = DataClass.DS.Field("Codice").value
        .GetSelectionExtended , varOrderFields
        If .DS.HowMany > 0 Then
            RefreshXArray XAACD, TDBGDettaglio, .DS.List
        Else
            ClearXArray XAACD, TDBGDettaglio
        End If
    End With
End Sub

Private Sub DisplayList()
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    Dim varOrderFields As Variant
    
    With DataClass
        .DS.Fields.ResetFields
        
        ReDim varOrderFields(1) As Variant
        varOrderFields(0) = .DS.KeyFields(1)
        varOrderFields(1) = .DS.KeyFields(2)
        If .GetSelectionList(, varOrderFields) Then
            With .DS
                If .HowMany > 0 Then
                    RefreshXArray XA, TDBGrid1, .List
                End If
            End With
        End If
    End With
End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    Dim strSelectionFormula As String
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcAddNew
            If (State = fsEdit) Or (State = fsNone) Then
                ClearXArray XA, TDBGrid1
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
                DisplayList
                If (State = fsInsert) Or (State = fsCopy) Then
                    SyncList DataClass.DS.Fields("1").value
                End If
                SetState fsEdit
            End If
        
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Or (State = fsCopy) Then
                SetState fsNone
            End If

        Case fcOpenLinkEdit
            If (State = fsEdit) Then
                Set LinkForm = DisplayForm(fcANAGACD)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .GetRecord XAACD.value(TDBGDettaglio.bookmark, 0), _
                                   XAACD.value(TDBGDettaglio.bookmark, 1), _
                                   XAACD.value(TDBGDettaglio.bookmark, 2)
                        .SetState fsEdit
                        .Command1.Enabled = False
                        .Command4.Enabled = False
                        Me.Enabled = False
                    End With
                End If
            End If
        
        Case fcOpenLinkAdd
            If (State = fsEdit) Then
                Set LinkForm = DisplayForm(fcANAGACD)
                If Not LinkForm Is Nothing Then
                    With LinkForm
                        Set .ParentForm = Me
                        .SetState fsInsert
                        
                        ' Cliente
                        .MaskEdBox7.Text = DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
                        .DataClass.DS.ValidateClientField .MaskEdBox7.Tag, MaskEdBox7.Text
                        FieldUpdateDescription .MaskEdBox7.Tag, LinkForm, MaskEdBox7
                        .MaskEdBox7.Enabled = False
                    
                        ' Contratto
                        .mskField1.Text = DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value
                        .DataClass.DS.ValidateClientField .mskField1.Tag, mskField1.Text
                        .mskField1.Enabled = False
                    
                        .Command1.Enabled = False
                        .Command4.Enabled = False
                        
                        .mskField4.SetFocus
                        Me.Enabled = False
                    End With
                End If
            End If

        Case fcLinkRefresh
            DisplayDetail

        Case fcCustom1
            If (State = fsEdit) And Not IsNull(TDBGDettaglio.bookmark) Then
                If MsgBox("Sei sicuro di voler eliminare la riga di dettaglio?", _
                           vbYesNo + vbExclamation + vbDefaultButton2, _
                          "Attenzione") = vbYes Then

                    If Not DetailClass.Delete(XAACD.value(TDBGDettaglio.bookmark, 0), _
                                   XAACD.value(TDBGDettaglio.bookmark, 1), _
                                   XAACD.value(TDBGDettaglio.bookmark, 2)) Then
                    End If
                    DisplayDetail
                End If
            End If

'        Case fcPrint
'            If (State = fsEdit) Then
'                varReportsList(0, 0) = "Stampa Tariffario"
'                varReportsList(1, 0) = "oil2tfb1.rpt"
'                varReportsList(2, 0) = 1
'                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
'                CRPrint varReportsList, strSelectionFormula, Me
'            End If
'        Case fcCustom1
'            If (State = fsEdit) Then
'                mstrFromIDDeposito = MaskEdBox1
'                mstrFromIDTariffa = mskField1
'                ClearXArray XA, TDBGrid1
'                DataClass.DS.Fields.ResetFields
'                ClearForm
'                SetState fsCopy
'            End If
        
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub SyncList(ByVal ID As String)
    Dim i As Integer
    Dim blnFound As Boolean
    Dim varList As Variant
    
    i = 0
    blnFound = False
    varList = DataClass.DS.List
    Do While i <= UBound(varList, 2) And (Not blnFound)
        If ID = varList(0, i) Then
            blnFound = True
        Else
            i = i + 1
        End If
    Loop
    If blnFound Then
        TDBGrid1.bookmark = i
    End If
End Sub

Public Function GetRecord(varIDACL As Variant, varIDACN As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = varIDACL
        .Fields(.KeyFields(2)).FindValue = varIDACN
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    SetState fsEdit
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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value
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
                ClearXArray XA, TDBGrid1
                TDBGrid1.bookmark = Null
                ClearGridDettaglio
                MaskEdBox7.SetFocus
            End If
            TDBGDettaglio.Enabled = False
'        Case fsCopy
'            TDBGDettaglio.Enabled = False
'            MaskEdBox7.SetFocus
        Case fsEdit
            MaskEdBox1.SetFocus
            TDBGDettaglio.Enabled = True
        Case fsInsert
            TDBGDettaglio.Enabled = False
            ClearGridDettaglio
            MaskEdBox7.SetFocus
    End Select
End Sub

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    ReDim ColTags(1 To MAX_COLS) As Variant
    
    Loading = True
    
    ColTags(1) = "0"
    ColTags(2) = "1"
    ColTags(3) = "2"
    ColTags(4) = ""
    ColTags(5) = "3"
    ColTags(6) = ""
    ColTags(7) = "4"
    ColTags(8) = "5"
    ColTags(9) = ""
    ColTags(10) = "6"
    ColTags(11) = ""
'
    Set DataClass = New CContratto
    Set DetailClass = New CContrattoDettaglio
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtFil1  'ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With

    With DetailClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set XAACD = New XArrayDB
    
    ClearGridDettaglio
    SetState fsNone
    DisplayList
    Loading = False
End Sub

Public Function InsertDetail() As Boolean
    On Error GoTo ErrHandler
    
    If Not DetailClass.Insert Then
        GoTo ErrHandler
    End If
    InsertDetail = True
    Exit Function
ErrHandler:
    InsertDetail = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set XAACD = Nothing
    Set DataClass = Nothing
    Set OIL2ACN0 = Nothing
End Sub

Private Sub TDBGDettaglio_DblClick()
    DoCommand fcOpenLinkEdit
End Sub

'Private Sub MaskEdBox1_LostFocus()
'    With DataClass.DS
'        If .ValidateClientField(MaskEdBox1.Tag, MaskEdBox1.Text) Then
'            If (State = fsNone) Or (State = fsInsert) Then
''                .Fields("0").FindValue = MaskEdBox1.Text
'                DisplayList
'            End If
'        End If
'    End With
'End Sub
'
'
'Private Sub mskField1_LostFocus()
'    TrackLostFocus mskField1
'End Sub
'

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strID As String
    
    If (State <> fsInsert) And (State <> fsCopy) Then
        strID = GetSelectedID(TDBGrid1)
        If Len(strID) > 0 Then
            GetRecord strID, XA.value(TDBGrid1.bookmark, 1)
        End If
    End If
End Sub


