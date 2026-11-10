VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2TFC0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TARIFFE CONTRATTO"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10020
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   4200
      TabIndex        =   28
      Tag             =   "9"
      Top             =   0
      Width           =   5775
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Tag             =   "8"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2520
         Picture         =   "OIL2TFC0.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1320
         Width           =   1635
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Tag             =   "2"
         Top             =   1320
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Tag             =   "1"
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Tag             =   "4"
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Tag             =   "3"
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "#,##0.00"
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
               Picture         =   "OIL2TFC0.frx":0102
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2TFC0.frx":0214
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2TFC0.frx":0326
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Tag             =   "7"
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Tag             =   "0"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Tag             =   "9"
         Top             =   960
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
         Index           =   7
         Left            =   5160
         Tag             =   "0"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   42
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2340
         Tag             =   "2"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3000
         Tag             =   "4"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data &Validità"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   41
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "C&ontratto"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   40
         Top             =   600
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
         Left            =   1800
         TabIndex        =   39
         Tag             =   "1"
         Top             =   600
         Width           =   3255
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
         Left            =   1800
         TabIndex        =   38
         Tag             =   "6"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Deposito"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   4560
         Tag             =   "6"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   4560
         Tag             =   "3"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Tariffa"
         Height          =   255
         Index           =   27
         Left            =   2640
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   5280
         Tag             =   "5"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5160
         Tag             =   "7"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dest&inazione"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   34
         Top             =   2760
         Width           =   1095
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
         TabIndex        =   33
         Tag             =   "7"
         Top             =   2760
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   32
         Tag             =   "0"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   735
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
         Left            =   1800
         TabIndex        =   30
         Tag             =   "9"
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Addebito"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5160
         Tag             =   "9"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   5160
         Tag             =   "1"
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   5220
      Width           =   9855
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1920
         Picture         =   "OIL2TFC0.frx":0438
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F11 - Copia"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2TFC0.frx":053A
         TabIndex        =   18
         Tag             =   "CUSTOM1"
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
         Picture         =   "OIL2TFC0.frx":063C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2TFC0.frx":073E
         TabIndex        =   17
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
         Index           =   2
         Left            =   8160
         Picture         =   "OIL2TFC0.frx":0840
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
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
         Picture         =   "OIL2TFC0.frx":0942
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
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
         Picture         =   "OIL2TFC0.frx":0A44
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
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
         Picture         =   "OIL2TFC0.frx":0B46
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2TFC0.frx":0C48
         TabIndex        =   16
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8040
         Picture         =   "OIL2TFC0.frx":0D4A
         TabIndex        =   19
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2TFC0.frx":0E4C
         TabIndex        =   15
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2TFC0.frx":0F4E
         TabIndex        =   14
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6480
      Width           =   10020
      _ExtentX        =   17674
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
            Object.Width           =   12224
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.47"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09/03/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "OIL2TFC0.frx":1050
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin TrueDBGrid60.TDBGrid TDBGDettaglio 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "OIL2TFC0.frx":4C22
      TabIndex        =   12
      Top             =   3600
      Width           =   4455
   End
   Begin TrueDBGrid60.TDBGrid TDBGScaglioni 
      Height          =   1575
      Left            =   4680
      OleObjectBlob   =   "OIL2TFC0.frx":9C21
      TabIndex        =   13
      Top             =   3600
      Width           =   5295
   End
End
Attribute VB_Name = "OIL2TFC0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 5
Private XA As XArrayDB
Private XAScaglioni As XArrayDB
Public GridArrayDetail As ArrayGridCls
Public WithEvents DataClass As CTariffaContratto
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CTariffaContrattoRiga
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private AVA As CValuta
Private ColTags As Variant
Private mstrIDACL As String
Private mstrIDACN As String
Private mstrIDADD As String
Private mobjParentForm As Object

Private mstrFromIDACL As String
Private mstrFromIDACN As String
Private mstrFromIDADD As String
Private mstrFromIDTFC As String

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Public Property Let IDACL(strIDACL As String)
    mstrIDACL = strIDACL
End Property

Public Property Get IDACL() As String
    IDACL = mstrIDACL
End Property

Public Property Let IDACN(strIDACN As String)
    mstrIDACN = strIDACN
End Property

Public Property Get IDACN() As String
    IDACN = mstrIDACN
End Property

Public Property Let IDADD(strIDADD As String)
    mstrIDADD = strIDADD
End Property

Public Property Get IDADD() As String
    IDADD = mstrIDADD
End Property

Public Sub ClearDetail1()
    Dim varList(1 To 5, 1 To 12) As Variant
    
    RefreshXArray XAScaglioni, TDBGScaglioni, varList
    TDBGScaglioni.bookmark = Null
End Sub

Private Sub ClearGridDettaglio()
    ClearUnboundGrid TDBGDettaglio, GridArrayDetail, MAX_COLS
    ClearDetail1
End Sub

Private Function Copy() As Boolean
    On Error GoTo ErrHandler
    Dim C As Control

    Dim blnChkValue As Boolean
    DataClass.DS.ScanClient 1

    If Not DataClass.DS.Fields.Validate Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If

    DataClass.DS.DA.BeginTransaction
    If Not DataClass.Insert Then
        GoTo ErrHandler
    End If
    If Not DataClass.CopyDetail(DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, _
                                DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value, _
                                DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value, _
                                DataClass.DS.Fields(DataClass.DS.KeyFields(4)).value, _
                                mstrFromIDACL, _
                                mstrFromIDACN, _
                                mstrFromIDADD, _
                                mstrFromIDTFC) Then
        GoTo ErrHandler
    End If
    GoTo EndProc
ErrHandler:
    DataClass.DS.DA.EndTrans ETROLLBACK
    Copy = False
    Exit Function
EndProc:
    DataClass.DS.DA.EndTrans ETCOMMIT
    Copy = True
End Function

Private Sub DisplayDetail1()
    On Error GoTo ErrHandler
    
    Dim varListScaglioni As Variant
'    Dim varList(1 To 5, 1 To 12) As Variant
    Dim varList(1 To 6, 1 To 15) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim intRow As Integer
    Dim intCount As Integer
    
    If State = fsEdit Then
        If IsNull(TDBGDettaglio.bookmark) Then
            GoTo ErrHandler
        End If
        Me.MousePointer = vbHourglass
        With DetailClass.DS
            .Fields(.KeyFields(1)).FindValue = MaskEdBox3.Text
            .Fields(.KeyFields(2)).FindValue = mskField10.Text
            .Fields(.KeyFields(3)).FindValue = MaskEdBox4.Text
            .Fields(.KeyFields(4)).FindValue = mskField1.Text
            .Fields(.KeyFields(5)).FindValue = GridArrayDetail.value(TDBGDettaglio.bookmark, 0)
        End With
        DetailClass.Find
        intCount = 5
        For i = 1 To 15
            For j = 1 To 6
                intCount = intCount + 1
                varList(j, i) = DetailClass.DS.Fields(CStr(intCount)).value ' varListScaglioni(intCount, intRow)
            Next j
        Next i
        RefreshXArray XAScaglioni, TDBGScaglioni, varList
        TDBGScaglioni.bookmark = 1
        Me.MousePointer = vbDefault
    End If
    Exit Sub
ErrHandler:

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    DoEvents
    If Not IsNull(DBGrid.bookmark) Then
        If DBGrid.bookmark <= XA.UpperBound(1) Then
            GetSelectedID = XA.value(DBGrid.bookmark, 0)
        End If
    End If
End Function

Private Sub GridToTFR(RowBookmark As Integer)
    Dim j As Integer
    Dim intCount As Integer
    
    With DetailClass.DS
        intCount = 5 + (RowBookmark - 1) * 6 '5
        For j = 1 To 6 ' 5
            intCount = intCount + 1
            .Fields(CStr(intCount)).NewValue = XAScaglioni(RowBookmark, j)
            .Fields(CStr(intCount)).Updated = True
        Next j
    End With
End Sub


Private Sub TrackLostFocus(Optional mskField As MaskEdBox)
    Dim strField As String
    Dim C As Control
    
    With DataClass.DS
        If mskField.Tag = .KeyFields(2) Then
            If .ValidateClientField(mskField.Tag, mskField.Text) Then
                If (State = fsInsert) Then
'                    .Fields("0").FindValue = MaskEdBox1.Text
'                    .Fields("1").FindValue = mskField.Text
                    If GetRecord(mskField1.Text, mskField10.Text, _
                            MaskEdBox4.Text, mskField1.Text) Then
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
    Dim DS As CDataStruct
    
    Set DS = DataClass.DS
    With DetailClass
        .DS.Fields(.DS.KeyFields(1)).FindValue = DS.Fields(DS.KeyFields(1)).value
        .DS.Fields(.DS.KeyFields(2)).FindValue = DS.Fields(DS.KeyFields(2)).value
        .DS.Fields(.DS.KeyFields(3)).FindValue = DS.Fields(DS.KeyFields(3)).value
        .DS.Fields(.DS.KeyFields(4)).FindValue = DS.Fields(DS.KeyFields(4)).value
        .GetList
        RefreshUnboundGrid TDBGDettaglio, .DS.List, GridArrayDetail, _
                           .DS.HowMany, MAX_COLS
    End With
End Sub


Public Sub DisplayList()
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    
    With DataClass
        .DS.Fields(.DS.KeyFields(1)).FindValue = IIf(Len(mstrIDACL) > 0, mstrIDACL, MaskEdBox3.Text)
        .DS.Fields(.DS.KeyFields(2)).FindValue = IIf(Len(mstrIDACN) > 0, mstrIDACN, mskField10.Text)
        .DS.Fields(.DS.KeyFields(3)).FindValue = IIf(Len(mstrIDADD) > 0, mstrIDADD, MaskEdBox4.Text)

        If .GetList Then
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
                Case fsCopy
                    blnResult = Copy
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
        Case fcPrint
            If (State = fsEdit) Then
                varReportsList(0, 0) = "Stampa Tariffario"
                varReportsList(1, 0) = "oil2tfb1.rpt"
                varReportsList(2, 0) = 1
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcCustom1
            If (State = fsEdit) Then
                mstrFromIDACL = MaskEdBox3.Text
                mstrFromIDACN = mskField10.Text
                mstrFromIDADD = MaskEdBox4.Text
                mstrFromIDTFC = mskField1.Text
                ClearXArray XA, TDBGrid1
                DataClass.DS.Fields.ResetFields
                ClearForm
                SetState fsCopy
            End If
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

Public Function GetRecord(varIDACL As Variant, _
    varIDACN As Variant, _
    varIDADD As Variant, _
    varIDTFC As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = varIDACL
        .Fields(.KeyFields(2)).FindValue = varIDACN
        .Fields(.KeyFields(3)).FindValue = varIDADD
        .Fields(.KeyFields(4)).FindValue = varIDTFC
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    DisplayDetail
                    SetState fsEdit
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
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(2)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(3)).value, _
                            DataClass.DS.Fields(DataClass.DS.KeyFields(4)).value
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
                If MaskEdBox1.Enabled And MaskEdBox1.Visible Then
                    MaskEdBox1.SetFocus
                End If
            End If
            TDBGDettaglio.Enabled = False
            TDBGScaglioni.Enabled = False
            Command8.Enabled = False
            Command5.Enabled = False
        
        Case fsCopy
            TDBGDettaglio.Enabled = False
            TDBGScaglioni.Enabled = False
            Command8.Enabled = False
            ' Set default value
            If Len(mstrIDACL) <> 0 Then
                MaskEdBox3.Text = mstrIDACL
                DataClass.DS.ValidateClientField MaskEdBox3.Tag, MaskEdBox3.Text
                FieldUpdateDescription MaskEdBox3.Tag, Me, MaskEdBox3
            End If
            
            If Len(mstrIDACN) <> 0 Then
                mskField10.Text = mstrIDACN
                DataClass.DS.ValidateClientField mskField10.Tag, mskField10.Text
                FieldUpdateDescription mskField10.Tag, Me, mskField10
            End If
            
            If Len(mstrIDADD) <> 0 Then
                MaskEdBox4.Text = mstrIDADD
                DataClass.DS.ValidateClientField MaskEdBox4.Tag, MaskEdBox4.Text
                FieldUpdateDescription MaskEdBox4.Tag, Me, MaskEdBox4
            End If
            
            mskField1.SetFocus
        
        Case fsEdit
            MaskEdBox2.SetFocus
            TDBGDettaglio.Enabled = True
            TDBGScaglioni.Enabled = True
            Command8.Enabled = True
            Command5.Enabled = True
        
        Case fsInsert
            TDBGDettaglio.Enabled = False
            TDBGScaglioni.Enabled = False
            Command8.Enabled = False
            ClearGridDettaglio
            
            ' Set default value
            If Len(mstrIDACL) <> 0 Then
                MaskEdBox3.Text = mstrIDACL
                DataClass.DS.ValidateClientField MaskEdBox3.Tag, MaskEdBox3.Text
                FieldUpdateDescription MaskEdBox3.Tag, Me, MaskEdBox3
                MaskEdBox3.Enabled = False
            End If
            
            If Len(mstrIDACN) <> 0 Then
                mskField10.Text = mstrIDACN
                DataClass.DS.ValidateClientField mskField10.Tag, mskField10.Text
                FieldUpdateDescription mskField10.Tag, Me, mskField10
                mskField10.Enabled = False
            End If
            
            If Len(mstrIDADD) <> 0 Then
                MaskEdBox4.Text = mstrIDADD
                DataClass.DS.ValidateClientField MaskEdBox4.Tag, MaskEdBox4.Text
                FieldUpdateDescription MaskEdBox4.Tag, Me, MaskEdBox4
                MaskEdBox4.Enabled = False
            End If
            
            mskField1.SetFocus
    End Select
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox16
            strTitle = "Data validità"
    End Select
    GetCalendarDate M, strTitle
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
    ReDim ColTags(1 To MAX_COLS) As Variant
    
    Loading = True
    
    ColTags(1) = "3"
    ColTags(2) = "4"
    ColTags(3) = "5"
    ColTags(4) = "96"
    ColTags(5) = "97"
'
    Set DataClass = New CTariffaContratto
    Set DetailClass = New CTariffaContrattoRiga
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtFil1  'ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me ', ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set XAScaglioni = New XArrayDB
    
    Set AVA = New CValuta
    With AVA.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    With DetailClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set GridArrayDetail = New ArrayGridCls
    ClearGridDettaglio
    SetValueItems TDBGScaglioni.Columns(4), AVA, False
    SetState fsNone
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

    If Not mobjParentForm Is Nothing Then
        mobjParentForm.Enabled = True
    End If
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2TFB0 = Nothing
End Sub





Private Sub MaskEdBox1_LostFocus()
    With DataClass.DS
        If .ValidateClientField(MaskEdBox1.Tag, MaskEdBox1.Text) Then
            If (State = fsNone) Or (State = fsInsert) Then
'                .Fields("0").FindValue = MaskEdBox1.Text
                DisplayList
            End If
        End If
    End With
End Sub


Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub


Private Sub TDBGDettaglio_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim blnDataOk As Boolean
    Dim i As Integer
    
    blnDataOk = True
    
    With TDBGDettaglio.Columns(ColIndex)
        Select Case ColIndex
            Case 0
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                    GoTo EndProc
                End If
            Case 1
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                End If
                If Len(TDBGDettaglio.Columns(2).Text) > 0 Then
                    If CInt(TDBGDettaglio.Columns(2).Text) < CInt(TDBGDettaglio.Columns(1).Text) Then
                        blnDataOk = False
                        GoTo EndProc
                    End If
                End If
            Case 2
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                End If
                If Len(TDBGDettaglio.Columns(2).Text) > 0 Then
                    If CLng(TDBGDettaglio.Columns(2).Text) < CLng(TDBGDettaglio.Columns(1).Text) Then
                        blnDataOk = False
                        GoTo EndProc
                    End If
                End If
            Case 3
                If (.value <> "C") And (.value <> "L") Then
                    blnDataOk = False
                End If
        End Select
    End With
EndProc:
    Cancel = Not blnDataOk
End Sub

Private Sub TDBGDettaglio_PostEvent(ByVal MsgId As Integer)
    On Error GoTo ErrHandler
    
    Select Case MsgId
        Case 0
            GridArrayDetail.DeleteRow 0
            With TDBGDettaglio
                .bookmark = 0
                .ReBind
                .Row = 1
            End With
    End Select
ErrHandler:
End Sub

Private Sub TDBGDettaglio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Select Case TDBGDettaglio.AddNewMode
        Case dbgNoAddNew
            ' --- Non modificare: LastRow può essere Null
            If LastRow = TDBGDettaglio.bookmark Then
            Else
                DisplayDetail1
            End If
            TDBGScaglioni.Enabled = True
        Case dbgAddNewCurrent
            ' --- Se ci siamo posizionati sulla riga di inserimento
            ' --- svuoto la grid di dettaglio collegata
            ClearDetail1
            TDBGScaglioni.Enabled = False
    End Select
End Sub


'*** APEX Migration Utility Code Change ***
'Private Sub TDBGDettaglio_UnboundAddData(ByVal RowBuf As TrueDBGrid50.RowBuffer, NewRowBookmark As Variant)
Private Sub TDBGDettaglio_UnboundAddData(ByVal RowBuf As TrueDBGrid60.RowBuffer, NewRowBookmark As Variant)
    If IsEmpty(GridArrayDetail.value(0, 0)) Then
        TDBGDettaglio.PostMsg 0
    End If
    GridArrayDetail.AddData RowBuf, NewRowBookmark, TDBGDettaglio.Columns
    RowToRecord NewRowBookmark, DetailClass.DS, GridArrayDetail, Me, ColTags
    If Not InsertDetail Then
        GridArrayDetail.Delete NewRowBookmark
    End If
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGDettaglio_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal OffSet As Long, ApproximatePosition As Long)
Private Sub TDBGDettaglio_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    Me.GridArrayDetail.ReadData RowBuf, StartLocation, offset, ApproximatePosition
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub TDBGDettaglio_UnboundWriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, WriteLocation As Variant)
Private Sub TDBGDettaglio_UnboundWriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, WriteLocation As Variant)
    Dim blnConcurError As Boolean
    
    GridArrayDetail.WriteData RowBuf, WriteLocation
    With DetailClass
        RowToRecord WriteLocation, .DS, GridArrayDetail, Me, ColTags
        .BeginSave blnConcurError
        If blnConcurError Then
            If MsgBox(MSG_CONCUR_ERR, _
                       vbOKCancel + vbExclamation + vbDefaultButton2, _
                      "Attenzione") = vbOK Then
                .DS.DA.EndTrans ETCOMMIT
                .DS.EndSave
            Else
                .DS.DA.EndTrans ETROLLBACK
                RowBuf.RowCount = 0
            End If
            blnConcurError = False
        Else
            .DS.DA.EndTrans ETCOMMIT
            .DS.EndSave
        End If
    End With
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strID As String
    
    If (State <> fsInsert) And (State <> fsCopy) Then
        strID = GetSelectedID(TDBGrid1)
        If Len(strID) > 0 Then
            GetRecord strID, _
                        XA.value(TDBGrid1.bookmark, 1), _
                        XA.value(TDBGrid1.bookmark, 2), _
                        XA.value(TDBGrid1.bookmark, 3)
        End If
    End If
End Sub




Private Sub TDBGScaglioni_AfterUpdate()
    Dim blnConcurError As Boolean
    
    GridToTFR TDBGScaglioni.bookmark
    DetailClass.BeginSave blnConcurError
    With DetailClass.DS
        If blnConcurError Then
            If MsgBox(MSG_CONCUR_ERR, _
                       vbOKCancel + vbExclamation + vbDefaultButton2, _
                      "Attenzione") = vbOK Then
                .DA.EndTrans ETCOMMIT
                .EndSave
            Else
                .DA.EndTrans ETROLLBACK
                DisplayDetail1
            End If
        Else
            .DA.EndTrans ETCOMMIT
            .EndSave
            DetailClass.RefreshTimeStamp
        End If
    End With
End Sub

Private Sub TDBGScaglioni_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim blnDataOk As Boolean
    Dim i As Integer
    
    blnDataOk = True
    
    With TDBGScaglioni.Columns(ColIndex)
        Select Case ColIndex
            Case 0, 1, 2, 3
                If Not IsNumeric(.Text) Then
                    blnDataOk = False
                End If
        End Select
    End With
EndProc:
    Cancel = Not blnDataOk
End Sub

