VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2APR0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODOTTI"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   9375
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   4320
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   5280
         Picture         =   "OIL2APR0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2APR0.frx":0102
         TabIndex        =   38
         Tag             =   "FIND"
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
         Index           =   4
         Left            =   3600
         Picture         =   "OIL2APR0.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
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
         Picture         =   "OIL2APR0.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   795
         Visible         =   0   'False
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
         Picture         =   "OIL2APR0.frx":0408
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
         Index           =   1
         Left            =   1920
         Picture         =   "OIL2APR0.frx":050A
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
         Picture         =   "OIL2APR0.frx":060C
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
         Picture         =   "OIL2APR0.frx":070E
         TabIndex        =   24
         Tag             =   "UNDO"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2APR0.frx":0810
         TabIndex        =   25
         Tag             =   "PRINT"
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2APR0.frx":0912
         TabIndex        =   26
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2APR0.frx":0A14
         TabIndex        =   23
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2APR0.frx":0B16
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
      TabIndex        =   28
      Top             =   5220
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
      Height          =   4215
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9375
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   20
         Tag             =   "8"
         Top             =   1800
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField2 
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Tag             =   "1"
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField3 
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Tag             =   "2"
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField5 
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Tag             =   "6"
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField6 
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Tag             =   "7"
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField4 
         Height          =   255
         Left            =   7680
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField9 
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Tag             =   "4"
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField7 
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Tag             =   "5"
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   1095
         Left            =   240
         OleObjectBlob   =   "OIL2APR0.frx":0C18
         TabIndex        =   21
         Tag             =   "MULTISOC"
         Top             =   2880
         Width           =   8895
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
               Picture         =   "OIL2APR0.frx":5A7A
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2APR0.frx":5B8C
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2APR0.frx":5C9E
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   7680
         TabIndex        =   13
         Tag             =   "9"
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Tag             =   "10"
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
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
         TabIndex        =   39
         Tag             =   "10"
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo&logia Prodotto"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   5640
         Tag             =   "10"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   8640
         Tag             =   "9"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Densità C&ontratto"
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   5640
         Tag             =   "5"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Raggr. &Fisc. Prodotto"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1695
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
         Left            =   2520
         TabIndex        =   36
         Tag             =   "5"
         Top             =   2160
         Width           =   3015
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
         TabIndex        =   35
         Tag             =   "4"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Raggr. Prodotto"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   5640
         Tag             =   "4"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descri&zione Abbr."
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   5
         Left            =   9000
         Tag             =   "3"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "De&nsità Max"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   4
         Left            =   5640
         Tag             =   "7"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Densità Min"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   2880
         Tag             =   "6"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "De&scrizione Fiscale"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   5640
         Tag             =   "2"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   7320
         Tag             =   "1"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "D&escrizione Comm."
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   3240
         Tag             =   "0"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Codice"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "OIL2APR0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GridArray As ArrayGridCls
Public WithEvents DataClass As CProdotto
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta
Private Imb As CImballo
Private Dog As CPosDoganale
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
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





Private Sub ClearForm()
    DataClass.DS.ClearClient
    ClearUnboundGrid TDBGrid2, GridArray, 5
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim Respok, Respuesta As String
    
    Select Case Cmd
    
        Case fcAddNew
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Solo se pueden modificar estos datos desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
            MsgBox "Password Correcta, datos guardados", vbOKOnly
        
                If (State = fsEdit) Or (State = fsNone) Then
                    DataClass.DS.Fields.ResetFields
                    ClearForm
                    SetState fsInsert
                End If
            End If
        Case fcSave
            Respok = CInt(Format(Date, "dd")) * CInt(Format(Date, "mm")) * CInt(Format(Time, "hh")) * 5
            Respuesta = InputBox("Solo se pueden modificar estos datos desde la Web", "RS3")
            If IsNumeric(Respuesta) And Not IsNull(Respuesta) Then Respuesta = CInt(Respuesta)
            If Respuesta = Respok Then
            MsgBox "Password Correcta, datos guardados", vbOKOnly
        
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
            End If
        Case fcUndo
            If (State = fsEdit) Or (State = fsInsert) Then
                SetState fsNone
            End If
        Case fcPrint
            If (State = fsEdit) Or (State = fsNone) Then
            End If
        Case fcFind
'            ShowSearchForm Me, "0", DataClass, 1
            DoSearch "APR", Me, "0", 1
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
                    DisplayMultiSoc
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




Private Sub DisplayMultiSoc()
    RefreshUnboundGrid TDBGrid2, DataClass.DS.MultiSoc.DS.List, GridArray, _
                       DataClass.DS.MultiSoc.DS.HowMany, 5
    
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
                mskField1.SetFocus
            End If
            ClearUnboundGrid TDBGrid2, GridArray, 5
        Case fsEdit
            mskField2.SetFocus
        Case fsInsert
            mskField1.SetFocus
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
    ReDim ColTags(1 To 5) As Variant
    
    Loading = True
    Set GridArray = New ArrayGridCls
    Set DataClass = New CProdotto
    
    ColTags(1) = "Società"
    ColTags(2) = "Imballo"
    ColTags(3) = "PosizioneDoganale"
    ColTags(4) = "CodPerSoc"
    ColTags(5) = "Annullato"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
        Set .DA = DataConnect.DA
        Set Soc = New CSocieta
        Set Soc.DS.DA = DataConnect.DA
        Soc.GetList True
        .ConfigureInterface
        ConfigureObjects Me, ColTags, True
        Set Imb = New CImballo
        Set Dog = New CPosDoganale
        'Set .MultiSoc.DS.DA = DA
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
'        Set Imb.DS.DA = DA
'        Set Dog.DS.DA = DA
        Set Imb.DS.DA = DataConnect.DA
        Set Dog.DS.DA = DataConnect.DA
    End With
    
    Imb.GetList True
    Dog.GetList True
    
    SetValueItems TDBGrid2.Columns(1), Imb
    SetValueItems TDBGrid2.Columns(2), Dog
    TDBGrid2.Columns(2).Alignment = dbgLeft
    TDBGrid2.Columns(4).Alignment = dbgCenter
    
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2APR0 = Nothing
End Sub





Private Sub mskField1_LostFocus()
    TrackLostFocus mskField1
End Sub



