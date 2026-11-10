VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OIL2ASZ0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STORICO PREZZI"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   7875
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   7815
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "OIL2ASZ0.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   380
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
         Left            =   6000
         Picture         =   "OIL2ASZ0.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ASZ0.frx":0204
         TabIndex        =   1
         Tag             =   "PRINT"
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   5880
         Picture         =   "OIL2ASZ0.frx":0306
         TabIndex        =   2
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4515
      Width           =   7875
      _ExtentX        =   13891
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
            Object.Width           =   8440
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.38"
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
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7815
      Begin TrueDBGrid60.TDBGrid TDBGrid1 
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "OIL2ASZ0.frx":0408
         TabIndex        =   0
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "OIL2ASZ0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Nome Form      : OIL2ASZ0
' Descrizione    : Visualizzazione Storico Prezzi
' Data Creazione : 18/11/99
' Autore         : Fabio Mozzi
Option Explicit



Public WithEvents DataClass As CStoricoPrezzi
Attribute DataClass.VB_VarHelpID = -1
Private XA As XArrayDB


Public State As FormState
Public Loading As Boolean

Public mblnConcurError As Boolean
Private mobjParentForm As Object
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private StrIDADZ As String
Private StrIDAPR As String
Private StrIDAIM As String
Private StrIDAPD As String
Public Property Let IDDestinazione(DEST As String)
     StrIDADZ = DEST
End Property
Public Property Let IDProdotto(PROD As String)
     StrIDAPR = PROD
End Property
Public Property Let IDImballo(IMBA As String)
     StrIDAIM = IMBA
End Property
Public Property Let IDPdoganale(PDOG As String)
     StrIDAPD = PDOG
End Property
Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
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
            End If
        Case fcFind
'            ShowSearchForm Me, "0", DataClass, 1
            DoSearch "APR", Me, "0", 1
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
    Select Case NewState
        Case fsNone
            If Not Loading Then
                DisplayList StrIDADZ, StrIDAPR, StrIDAIM, StrIDAPD
                TDBGrid1.SetFocus
            End If
            
        Case fsEdit
            'mskField2.SetFocus
        Case fsInsert
            'mskField1.SetFocus
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
     
    Set DataClass = New CStoricoPrezzi
    
    ColTags(1) = "4"
    ColTags(2) = "5"
    ColTags(3) = "6"
    ColTags(4) = "7"
    ColTags(5) = "8"
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1S
'        Set .DA = DataConnect.DA
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags, True
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    SetState fsNone
    Loading = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
    mobjParentForm.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set XA = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set OIL2ASZ0 = Nothing
End Sub
Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property
Private Sub DisplayList(Destinazione As String, Prodotto As String, Imballo As String, Pdoganale As String)
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    DataClass.DS.Fields("0").FindValue = Destinazione
    DataClass.DS.Fields("1").FindValue = Prodotto
    DataClass.DS.Fields("2").FindValue = Imballo
    DataClass.DS.Fields("3").FindValue = Pdoganale
    If DataClass.GetList Then
        If DataClass.DS.HowMany > 0 Then
            RefreshXArray XA, TDBGrid1, DataClass.DS.List
        End If
    End If
End Sub
