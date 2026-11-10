VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2TFA0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FASCE TARIFFARIE"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   9870
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   30
      Top             =   5280
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
         Picture         =   "OIL2TFA0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F11 - Copia"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2TFA0.frx":0102
         TabIndex        =   41
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
         Picture         =   "OIL2TFA0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2TFA0.frx":0306
         TabIndex        =   39
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
         Picture         =   "OIL2TFA0.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
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
         Picture         =   "OIL2TFA0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
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
         Picture         =   "OIL2TFA0.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
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
         Picture         =   "OIL2TFA0.frx":070E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2TFA0.frx":0810
         TabIndex        =   24
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8040
         Picture         =   "OIL2TFA0.frx":0912
         TabIndex        =   25
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2TFA0.frx":0A14
         TabIndex        =   23
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2TFA0.frx":0B16
         TabIndex        =   22
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6525
      Width           =   9870
      _ExtentX        =   17410
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
            Object.Width           =   11959
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11.43"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "07/08/2003"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   4200
      TabIndex        =   26
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "9"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2520
         Picture         =   "OIL2TFA0.frx":0C18
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   20
         Tag             =   "8"
         Top             =   2760
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Tag             =   "4"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField5 
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Tag             =   "6"
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Tag             =   "5"
         Top             =   2040
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
         Left            =   1200
         TabIndex        =   17
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
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Tag             =   "3"
         Top             =   1320
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Tag             =   "2"
         Top             =   960
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
               Picture         =   "OIL2TFA0.frx":0D1A
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2TFA0.frx":0E2C
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2TFA0.frx":0F3E
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Tag             =   "10"
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
         TabIndex        =   43
         Tag             =   "10"
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dest&inazione"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   18
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5160
         Tag             =   "10"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   27
         Left            =   5280
         Tag             =   "9"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Tariffa"
         Height          =   255
         Index           =   27
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   4560
         Tag             =   "2"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   4560
         Tag             =   "0"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Zona"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   0
         Top             =   240
         Width           =   615
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
         TabIndex        =   35
         Tag             =   "0"
         Top             =   240
         Width           =   2655
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
         TabIndex        =   29
         Tag             =   "7"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Val Min. 2"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   16
         Top             =   2760
         Width           =   735
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   3720
         Tag             =   "7"
         Top             =   2760
         Width           =   240
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
         Left            =   1800
         TabIndex        =   28
         Tag             =   "5"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Val. Min. 1"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   3720
         Tag             =   "5"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data &Validità"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3000
         Tag             =   "3"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Imp Min. 2"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   2520
         Tag             =   "6"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Imp. Min. 1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   2520
         Tag             =   "4"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   1800
         Tag             =   "1"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "OIL2TFA0.frx":1050
      TabIndex        =   21
      Top             =   120
      Width           =   3975
   End
   Begin TrueDBGrid60.TDBGrid TDBGDettaglio 
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "OIL2TFA0.frx":3A42
      TabIndex        =   36
      Top             =   3600
      Width           =   4455
   End
   Begin TrueDBGrid60.TDBGrid TDBGScaglioni 
      Height          =   1575
      Left            =   4680
      OleObjectBlob   =   "OIL2TFA0.frx":8A41
      TabIndex        =   37
      Top             =   3600
      Width           =   5055
   End
End
Attribute VB_Name = "OIL2TFA0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 5
Private XA As XArrayDB
Private XAScaglioni As XArrayDB
Public GridArrayDetail As ArrayGridCls
Public WithEvents DataClass As CFasciaTariffaria
Attribute DataClass.VB_VarHelpID = -1
Public DetailClass As CFasciaTariffariaRiga
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private AVA As CValuta
Private ColTags As Variant
Private mstrFromIDZona As String
Private mstrFromIDTariffa As String

Public Sub ClearDetail1()
    Dim varList(1 To 5, 1 To 12) As Variant
    
    RefreshXArray XAScaglioni, TDBGScaglioni, varList
    TDBGScaglioni.Bookmark = Null
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
    If Not DataClass.CopyDetail(DataClass.DS.Fields("0").value, _
                                DataClass.DS.Fields("1").value, _
                                mstrFromIDZona, _
                                mstrFromIDTariffa) Then
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
    Dim varList(1 To 6, 1 To 12) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim intRow As Integer
    Dim intCount As Integer
    
    If State = fsEdit Then
        If IsNull(TDBGDettaglio.Bookmark) Then
            GoTo ErrHandler
        End If
        Me.MousePointer = vbHourglass
        With DetailClass.DS
            .Fields(.KeyFields(1)).FindValue = MaskEdBox1
            .Fields(.KeyFields(2)).FindValue = mskField1
            .Fields(.KeyFields(3)).FindValue = GridArrayDetail.value(TDBGDettaglio.Bookmark, 0)
        End With
        DetailClass.Find
        intCount = 4
        For i = 1 To 12
            For j = 1 To 6
                intCount = intCount + 1
                varList(j, i) = DetailClass.DS.Fields(CStr(intCount)).value ' varListScaglioni(intCount, intRow)
            Next j
        Next i
        RefreshXArray XAScaglioni, TDBGScaglioni, varList
        TDBGScaglioni.Bookmark = 1
        Me.MousePointer = vbDefault
    End If
    Exit Sub
ErrHandler:

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    DoEvents
    If Not IsNull(DBGrid.Bookmark) Then
        If DBGrid.Bookmark <= XA.UpperBound(1) Then
            GetSelectedID = XA.value(DBGrid.Bookmark, 0)
        End If
    End If
End Function

Private Sub GridToTFR(RowBookmark As Integer)
    Dim j As Integer
    Dim intCount As Integer
    
    With DetailClass.DS
        intCount = 4 + (RowBookmark - 1) * 6 '5
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
                    If GetRecord(mskField1) Then
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
        .DS.Fields("0").FindValue = DS.Fields("0").value
        .DS.Fields("1").FindValue = DS.Fields("1").value
        .GetList
        RefreshUnboundGrid TDBGDettaglio, .DS.List, GridArrayDetail, _
                           .DS.HowMany, MAX_COLS
    End With
End Sub


Private Sub DisplayList()
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    
    With DataClass
        .DS.Fields("0").FindValue = MaskEdBox1.Text
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
                varReportsList(1, 0) = "oil2tfa1.rpt"
                varReportsList(2, 0) = 1
                strSelectionFormula = DataClass.DS.MakeKeyCond(True)
                CRPrint varReportsList, strSelectionFormula, Me
            End If
        Case fcCustom1
            If (State = fsEdit) Then
                mstrFromIDZona = MaskEdBox1
                mstrFromIDTariffa = mskField1
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
        TDBGrid1.Bookmark = i
    End If
End Sub

Public Function GetRecord(strID As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = MaskEdBox1
        .Fields(.KeyFields(2)).FindValue = strID
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
                ClearXArray XA, TDBGrid1
                TDBGrid1.Bookmark = Null
                ClearGridDettaglio
                MaskEdBox1.SetFocus
            End If
            TDBGDettaglio.Enabled = False
            TDBGScaglioni.Enabled = False
            Command8.Enabled = False
            Command5.Enabled = False
        Case fsCopy
            TDBGDettaglio.Enabled = False
            TDBGScaglioni.Enabled = False
            Command8.Enabled = False
            MaskEdBox1.SetFocus
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
            MaskEdBox1.SetFocus
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
    
    ColTags(1) = "2"
    ColTags(2) = "3"
    ColTags(3) = "4"
    ColTags(4) = "95"
    ColTags(5) = "96"
'
    Set DataClass = New CFasciaTariffaria
    Set DetailClass = New CFasciaTariffariaRiga
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
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2TFA0 = Nothing
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
                    If CInt(TDBGDettaglio.Columns(2).Text) < CInt(TDBGDettaglio.Columns(1).Text) Then
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
                .Bookmark = 0
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
            If LastRow = TDBGDettaglio.Bookmark Then
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
            GetRecord strID
        End If
    End If
End Sub




Private Sub TDBGScaglioni_AfterUpdate()
    Dim blnConcurError As Boolean
    
    GridToTFR TDBGScaglioni.Bookmark
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

