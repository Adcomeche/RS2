VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2AAM1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AUTOMEZZI"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   1215
      Left            =   0
      TabIndex        =   36
      Top             =   4920
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1920
         Picture         =   "OIL2AAM1.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2AAM1.frx":0102
         TabIndex        =   43
         Tag             =   "FIND"
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2AAM1.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   41
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
         Picture         =   "OIL2AAM1.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   40
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
         Picture         =   "OIL2AAM1.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   39
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
         Index           =   3
         Left            =   240
         Picture         =   "OIL2AAM1.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
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
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2AAM1.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2AAM1.frx":070E
         TabIndex        =   31
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2AAM1.frx":0810
         TabIndex        =   32
         Tag             =   "PRINT"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2AAM1.frx":0912
         TabIndex        =   33
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2AAM1.frx":0A14
         TabIndex        =   30
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2AAM1.frx":0B16
         TabIndex        =   29
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Top             =   6165
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
            TextSave        =   "15.39"
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
      Height          =   4815
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "33"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "32"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "S&compartatura"
         Height          =   2175
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   5415
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1575
            Left            =   240
            OleObjectBlob   =   "OIL2AAM1.frx":0C18
            TabIndex        =   17
            Tag             =   "ENABLE"
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "1"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "4"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   8040
         TabIndex        =   28
         Tag             =   "26"
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   7920
         TabIndex        =   19
         Tag             =   "25"
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Tag             =   "2"
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   7920
         TabIndex        =   21
         Tag             =   "28"
         Top             =   3240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   7920
         TabIndex        =   23
         Tag             =   "29"
         Top             =   3600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   7920
         TabIndex        =   25
         Tag             =   "30"
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   7920
         TabIndex        =   27
         Tag             =   "31"
         Top             =   4320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Tag             =   "3"
         Top             =   1200
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
         Left            =   1800
         TabIndex        =   7
         Tag             =   "27"
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   120
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
               Picture         =   "OIL2AAM1.frx":4C22
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2AAM1.frx":4D34
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2AAM1.frx":4E46
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   9000
         Tag             =   "33"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Scarico"
         Height          =   255
         Index           =   13
         Left            =   5880
         TabIndex        =   14
         Top             =   1920
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   5400
         Tag             =   "27"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tip&o Cisterna"
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   6
         Top             =   1560
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
         Left            =   2400
         TabIndex        =   47
         Tag             =   "27"
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5400
         Tag             =   "3"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ti&po Automezzo"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblField 
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
         Left            =   2400
         TabIndex        =   46
         Tag             =   "3"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   10
         Left            =   9000
         Tag             =   "31"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "C&ubatura costruttore in Lt"
         Height          =   255
         Index           =   10
         Left            =   5880
         TabIndex        =   26
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   9000
         Tag             =   "30"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Portata rimorc&hiabile in Kg"
         Height          =   255
         Index           =   9
         Left            =   5880
         TabIndex        =   24
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   9000
         Tag             =   "29"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tara li&bretto in Kg"
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   22
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   9000
         Tag             =   "28"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tara pe&sata in Kg"
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   20
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ca&rica"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   9000
         Tag             =   "32"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblPortata 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7920
         TabIndex        =   45
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Portata Total&e in Lt"
         Height          =   255
         Index           =   12
         Left            =   6360
         TabIndex        =   44
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   3960
         Tag             =   "1"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Motrice/Rimorchio"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   9000
         Tag             =   "4"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "D&isponibilità"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Portata Totale in &Kg"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   18
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   9000
         Tag             =   "25"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5400
         Tag             =   "2"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3960
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Targa"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2AAM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private XA As XArrayDB
Public WithEvents DataClass As CAutomezzo
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private Sub ClearForm()
    DataClass.DS.ClearClient
    lblPortata = vbNullString
End Sub



Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    
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
                varReportsList(0, 0) = "Stampa elenco Automezzi"
                varReportsList(1, 0) = "oil2aam1.rpt"
                varReportsList(2, 0) = 1
                CRPrint varReportsList, "", Me
            End If
        Case fcFind
            DoSearch "AAM", Me, "0", 1
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
                    GetTotaleScomparti
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




Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If Not IsNull(DBGrid.Bookmark) Then
        GetSelectedID = XA.value(DBGrid.Bookmark, 0)
    End If
End Function

Private Sub GetTotaleScomparti()
    Dim i As Integer
    Dim dblPortata As Double
    Dim intNumScomparti As Integer
    Dim dblPortataScomparto As Double
    
    For i = XA.LowerBound(1) To XA.UpperBound(1)
        If Len(XA(i, 3)) = 0 Then
            intNumScomparti = 0
        Else
            intNumScomparti = XA(i, 3)
        End If
        If Len(XA(i, 2)) = 0 Then
            dblPortataScomparto = 0
        Else
            dblPortataScomparto = XA(i, 2)
        End If
        dblPortata = dblPortata + (intNumScomparti * dblPortataScomparto)
    Next i
    lblPortata = dblPortata

End Sub

Private Function InsertRecord() As Boolean
    On Error GoTo ErrHandler
    Dim C As Control
    Dim i As Integer
    
    Dim blnChkValue As Boolean
    DataClass.DS.ScanClient 1
    
    For i = 1 To 10
        If XA.value(i, 2) <> DataClass.DS.Fields(CStr((2 * i - 1) + 4)).value Then
            DataClass.DS.Fields(CStr((2 * i - 1) + 4)).NewValue = XA.value(i, 2)
            DataClass.DS.Fields(CStr((2 * i - 1) + 4)).Updated = True
        End If
        If XA.value(i, 3) <> DataClass.DS.Fields(CStr((2 * i) + 4)).value Then
            DataClass.DS.Fields(CStr((2 * i) + 4)).NewValue = XA.value(i, 3)
            DataClass.DS.Fields(CStr((2 * i) + 4)).Updated = True
        End If
    Next i
    
'    If DataClass.DS.Fields.Validate Then
    If DataClass.Validate Then
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
    Dim i As Integer
    
    'Scansione della finestra
    DataClass.DS.ScanClient 0
    
    For i = 1 To 10
        If XA.value(i, 2) <> DataClass.DS.Fields(CStr((2 * i - 1) + 4)).value Then
            DataClass.DS.Fields(CStr((2 * i - 1) + 4)).NewValue = XA.value(i, 2)
            DataClass.DS.Fields(CStr((2 * i - 1) + 4)).Updated = True
        End If
        If XA.value(i, 3) <> DataClass.DS.Fields(CStr((2 * i) + 4)).value Then
            DataClass.DS.Fields(CStr((2 * i) + 4)).NewValue = XA.value(i, 3)
            DataClass.DS.Fields(CStr((2 * i) + 4)).Updated = True
        End If
    Next i
    
'    If DataClass.DS.Fields.Validate Then
    If DataClass.Validate Then
        DataClass.BeginSave blnConcurError
        If blnConcurError Then
            If MsgBox(MSG_CONCUR_ERR, _
                       vbOKCancel + vbExclamation + vbDefaultButton2, _
                      "Attenzione") = vbOK Then
                GoTo EndProc
            Else
                DataClass.DS.DA.EndTrans ETROLLBACK
                GetRecord DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
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
    TDBGrid1.Bookmark = 1
    Select Case NewState
        Case fsNone
            If Not Loading Then
                mskField1.SetFocus
                lblPortata = vbNullString
            End If
        Case fsEdit
            Combo1.SetFocus
        Case fsInsert
            mskField1.SetFocus
    End Select
End Sub

Private Sub DataClass_ClientClear()
    Dim i As Integer
    
    If Not Loading Then
        For i = 1 To 10
            XA(i, 2) = vbNullString
            XA(i, 3) = vbNullString
        Next i
        TDBGrid1.ReBind
    End If
End Sub

Private Sub DataClass_ClientRefresh()
    Dim i As Integer
    With DataClass.DS
        For i = 1 To 10
            XA.value(i, 2) = .Fields(CStr((i * 2 - 1) + 4)).value
            XA.value(i, 3) = .Fields(CStr((i * 2) + 4)).value
        Next i
    End With
    TDBGrid1.Array = XA
    TDBGrid1.ReBind
End Sub

'Private Sub DataClass_Command(Cmd As RESLIB05.FormCommand)
'    DoCommand Cmd
'End Sub
'
Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim varList(1 To 3, 1 To 10) As Variant
    
    Loading = True
    
    
    Set DataClass = New CAutomezzo
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set XA = New XArrayDB
    For i = 1 To 10
        varList(1, i) = i
    Next i
    
    RefreshXArray XA, TDBGrid1, varList
    
    SetState fsNone
    Loading = False
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set XA = Nothing
    Set ErrH = Nothing
    Set DataClass.DS.FormHandle = Nothing
    Set DataClass = Nothing
    Set OIL2AAM1 = Nothing
End Sub

Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub


Private Sub TDBGrid1_AfterInsert()
    GetTotaleScomparti
End Sub


Private Sub TDBGrid1_AfterUpdate()
    GetTotaleScomparti
End Sub


