VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2FPD0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FASCIA PREZZI DETTAGLIO"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   8700
   Begin VB.Frame Frame1 
      Height          =   3735
      Index           =   1
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8655
      Begin VB.CheckBox Check1 
         Caption         =   "Annullato"
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Tag             =   "15"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Margine"
         Height          =   1575
         Left            =   5760
         TabIndex        =   26
         Top             =   2040
         Width           =   2775
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   255
            Left            =   960
            TabIndex        =   8
            Tag             =   "12"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   255
            Left            =   960
            TabIndex        =   9
            Tag             =   "13"
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            Height          =   255
            Left            =   960
            TabIndex        =   10
            Tag             =   "14"
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   13
            Left            =   2280
            Tag             =   "14"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   12
            Left            =   2280
            Tag             =   "13"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   11
            Left            =   2280
            Tag             =   "12"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Margine 1"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Margine 2"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Margine 3"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sconto in Valuta"
         Height          =   1575
         Left            =   2880
         TabIndex        =   22
         Top             =   2040
         Width           =   2775
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   255
            Left            =   960
            TabIndex        =   5
            Tag             =   "9"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            Height          =   255
            Left            =   960
            TabIndex        =   6
            Tag             =   "10"
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10 
            Height          =   255
            Left            =   960
            TabIndex        =   7
            Tag             =   "11"
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.000"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   10
            Left            =   2280
            Tag             =   "11"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   9
            Left            =   2280
            Tag             =   "10"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   8
            Left            =   2280
            Tag             =   "9"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Valore 1"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Valore 2"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Valore 3"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Sconto Percentuale"
         Height          =   1575
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   2655
         Begin MSMask.MaskEdBox MaskEdBox20 
            Height          =   255
            Left            =   960
            TabIndex        =   2
            Tag             =   "6"
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Tag             =   "7"
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   960
            TabIndex        =   4
            Tag             =   "8"
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "0.00"
            PromptChar      =   "_"
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   7
            Left            =   2160
            Tag             =   "8"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   6
            Left            =   2160
            Tag             =   "7"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgState 
            Height          =   240
            Index           =   5
            Left            =   2160
            Tag             =   "6"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sconto 3"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sconto 2"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sconto 1"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Tag             =   "4"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   4680
         TabIndex        =   1
         Tag             =   "5"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Tag             =   "0"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1680
         TabIndex        =   38
         Tag             =   "1"
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Tag             =   "2"
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1680
         TabIndex        =   40
         Tag             =   "3"
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice Fascia"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Prodotto"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   46
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Imballo"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   45
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Posizione &Doganale"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   1455
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
         Left            =   2520
         TabIndex        =   43
         Tag             =   "3"
         Top             =   1320
         Width           =   3375
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
         Left            =   2520
         TabIndex        =   42
         Tag             =   "2"
         Top             =   960
         Width           =   3375
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
         Left            =   2880
         TabIndex        =   41
         Tag             =   "1"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   6000
         Tag             =   "1"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   6000
         Tag             =   "2"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   6000
         Tag             =   "3"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   6000
         Tag             =   "5"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   3000
         Tag             =   "4"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prezzo Vendita"
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Prezzo acquisto"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   30
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Width           =   8655
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2FPD0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2FPD0.frx":0102
         TabIndex        =   35
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
         Index           =   0
         Left            =   240
         Picture         =   "OIL2FPD0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   450
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
         Left            =   7080
         Picture         =   "OIL2FPD0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   6960
         Picture         =   "OIL2FPD0.frx":0408
         TabIndex        =   12
         Tag             =   "EXIT"
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
         Index           =   3
         Left            =   1920
         Picture         =   "OIL2FPD0.frx":050A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2FPD0.frx":060C
         TabIndex        =   11
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2FPD0.frx":070E
         TabIndex        =   33
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   4755
      Width           =   8700
      _ExtentX        =   15346
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
            Object.Width           =   9895
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12.40"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "22/03/00"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
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
            Picture         =   "OIL2FPD0.frx":0810
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2FPD0.frx":0922
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OIL2FPD0.frx":0A34
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OIL2FPD0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Nome Form      : OILFPD0
' Descrizione    : Gestione Fasce di Prezzi (Dettaglio)
' Data Creazione : 22/11/99
' Autore         : Fabio Mozzi

Option Explicit

Public WithEvents DataClass As CFasciaPrezziDettaglio
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object
Private StrFasciaPrezzi As String
Property Let SetIdFasciaPrezzi(FPZ As String)
    StrFasciaPrezzi = FPZ
End Property
Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcAddNew
                DataClass.DS.Fields.ResetFields
                ClearForm
                SetState fsInsert
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
            Select Case State
                Case fsEdit
                    SetState fsEdit
                Case fsInsert
                    SetState fsNone
            End Select
        Case fcExit
            Unload Me
    End Select
End Sub
Public Function GetRecord(strIDFascia As String, strIDProdotto As String, strIDImballo As String, strIDPDoganale As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strIDFascia
        .Fields(.KeyFields(2)).FindValue = strIDProdotto
        .Fields(.KeyFields(3)).FindValue = strIDImballo
        .Fields(.KeyFields(4)).FindValue = strIDPDoganale
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
    Dim strMsg As String
    Dim i As Integer
    Dim C As Control
    
    'Scansione della finestra
    If Not DataClass.DS.ScanClient(0) Then
        GoTo ErrHandler
    End If
    
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
    mskField1 = StrFasciaPrezzi
    Select Case NewState
        Case fsEdit
            mskField1.Enabled = False
        Case fsInsert
            mskField1.Enabled = False
            MaskEdBox1.SetFocus
        Case fsNone
            mskField1.Enabled = False
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
    Set DataClass = New CFasciaPrezziDettaglio
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    Loading = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not CheckExit(State)
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcLinkRefresh
        End With
    End If
    Set OIL2FPD0 = Nothing
End Sub
Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

