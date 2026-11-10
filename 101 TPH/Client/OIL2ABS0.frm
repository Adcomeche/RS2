VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form OIL2ABS0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BASI"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   9390
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "&Annullato"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Tag             =   "Annullato"
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Tag             =   "Codice"
         Top             =   360
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
         Left            =   1920
         TabIndex        =   3
         Tag             =   "Descrizione"
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4920
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
               Picture         =   "OIL2ABS0.frx":0000
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ABS0.frx":0112
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OIL2ABS0.frx":0224
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   1215
         Left            =   120
         OleObjectBlob   =   "OIL2ABS0.frx":0336
         TabIndex        =   17
         Top             =   1440
         Width           =   5355
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5280
         Tag             =   "Descrizione"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   2640
         Tag             =   "Codice"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Descrizione"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   855
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
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7680
         Picture         =   "OIL2ABS0.frx":30C0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "OIL2ABS0.frx":31C2
         TabIndex        =   15
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
         Height          =   225
         Index           =   6
         Left            =   5280
         Picture         =   "OIL2ABS0.frx":32C4
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Elimina Distanza"
         Height          =   375
         Left            =   5160
         Picture         =   "OIL2ABS0.frx":37F6
         TabIndex        =   19
         Tag             =   "CUSTOM1"
         Top             =   360
         Width           =   2415
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
         Picture         =   "OIL2ABS0.frx":38F8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   3480
         Picture         =   "OIL2ABS0.frx":39FA
         TabIndex        =   16
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
         Index           =   1
         Left            =   1920
         Picture         =   "OIL2ABS0.frx":3AFC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   1800
         Picture         =   "OIL2ABS0.frx":3BFE
         TabIndex        =   14
         Tag             =   "SAVE"
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
         Picture         =   "OIL2ABS0.frx":3D00
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1 - Inserisci"
         Height          =   375
         Left            =   120
         Picture         =   "OIL2ABS0.frx":3E02
         TabIndex        =   13
         Tag             =   "ADD"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3870
      Width           =   9390
      _ExtentX        =   16563
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
            Object.Width           =   11113
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15.28"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "26/02/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "OIL2ABS0.frx":3F04
      TabIndex        =   5
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "OIL2ABS0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3
Private Const MAX_COLS_DISTANZE = 2

Private XA As XArrayDB
Public WithEvents DataClass As CBase
Attribute DataClass.VB_VarHelpID = -1
Public GridArray As ArrayGridCls
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private GridArrayDistanze As ArrayGridCls
Private ADE As CDeposito

Private Sub ClearForm()
    DataClass.DS.ClearClient
    ClearUnboundGrid TDBGrid2, GridArrayDistanze, MAX_COLS_DISTANZE
End Sub

Private Sub DisplayList()
    Dim i As Long
    Dim j As Long
    Dim NumRows As Long
    
    If DataClass.GetList Then
        If DataClass.DS.HowMany > 0 Then
            RefreshXArray XA, TDBGrid1, DataClass.DS.List
        End If
    End If
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcAddNew
            If State = fsEdit Then
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
            End Select
            If blnResult Then
                DisplayList
                If State = fsInsert Then
                    SyncList DataClass.DS.Fields(DataClass.DS.KeyFields(1)).value
                End If
                SetState fsEdit
            End If
        Case fcUndo
            DisplayList
            SetState fsEdit
        Case fcPrint
            If State = fsEdit Then
            End If
        Case fcCustom1
            If (State = fsEdit) And Not IsNull(TDBGrid2.bookmark) And Not IsEmpty(TDBGrid2.Columns(0).value) Then
                If MsgBox("Sei sicuro di voler eliminare le distanze per il deposito " & TDBGrid2.Columns(0).Text & "?", _
                           vbYesNo + vbExclamation + vbDefaultButton2, _
                          "Attenzione") = vbYes Then

                    If Not DataClass.DeleteDistanze(DataClass.DS.Field("Codice").value, _
                                    TDBGrid2.Columns(0).value) Then
                        MsgBox "Impossibile eliminare le distanze per il deposito " & TDBGrid2.Columns(0).Text & "!", , vbOKOnly, "Errore"
                    End If
                    DisplayDistanze
                End If
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

Private Sub GetRecord(strID As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
            .RefreshClient
            DisplayDistanze
        Else
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub




Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If Not IsNull(DBGrid.bookmark) Then
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
    End If
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
                DisplayList
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
        Case fsEdit
            If Not Loading Then
                TDBGrid1.SetFocus
            End If
        Case fsInsert
            mskField1.Enabled = True
            mskField1.SetFocus
    End Select
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
    Dim i As Integer
    
    Loading = True
    
    Set GridArray = New ArrayGridCls
    Set GridArrayDistanze = New ArrayGridCls
    
    Set DataClass = New CBase
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtFil1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set ADE = New CDeposito
    Set ADE.DS.DA = DataConnect.DA
    SetValueItemsDeposito TDBGrid2.Columns(0), ADE
    ClearUnboundGrid TDBGrid2, GridArrayDistanze, MAX_COLS_DISTANZE

    DisplayList
    SetState fsEdit
    
    Loading = False
End Sub

Private Sub DisplayDistanze()
    Dim varList As Variant
    Dim lngHowMany As Long
    
    varList = DataClass.GetListDistanze(DataClass.DS.Fields("Codice").value)
    If Not IsEmpty(varList) Then
        lngHowMany = UBound(varList, 2) + 1
        RefreshUnboundGrid TDBGrid2, varList, GridArrayDistanze, _
                           lngHowMany, MAX_COLS_DISTANZE
    Else
        ClearUnboundGrid TDBGrid2, GridArrayDistanze, MAX_COLS_DISTANZE
    End If
    
End Sub

Private Sub SetValueItemsDeposito(Col As Column, Entity As Object, _
                         Optional Translate As Boolean = True)
    Dim vItem As ValueItem
    Dim DS As CDataStruct
    Dim i As Long
    Dim intRows As Integer
    
    With Col.ValueItems
        .Presentation = dbgSortedComboBox
        .Translate = Translate
        .Validate = True
        Set DS = Entity.DS
        Entity.GetList True
        If DS.HowMany > 0 Then
            .MaxComboItems = 20
            intRows = UBound(DS.List, 2)
            For i = 0 To intRows 'UBound(DS.List, 2)
                 Set vItem = New ValueItem
                 vItem.value = DS.ListItem(0, i)
                 vItem.DisplayValue = DS.ListItem(1, i) & " (" & DS.ListItem(0, i) & ")"
                .Add vItem
            Next i
        End If
    End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set XA = Nothing
    Set ErrH = Nothing
    Set GridArray = Nothing
    Set DataClass = Nothing
    Set OIL2ABS0 = Nothing
End Sub







Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strID As String
    
    strID = GetSelectedID(TDBGrid1)
    GetRecord strID
End Sub





Private Sub TDBGrid2_UnboundAddData(ByVal RowBuf As TrueDBGrid60.RowBuffer, NewRowBookmark As Variant)
    GridArrayDistanze.AddData RowBuf, NewRowBookmark, TDBGrid2.Columns
    
    If Not DataClass.InsertDistanze(DataClass.DS.Fields("Codice").value, _
                             TDBGrid2.Columns(0).value, _
                             TDBGrid2.Columns(1).value) Then
        DataClass.DS.DisplayMessage "L'inserimento delle distanze è fallito", 9
    End If
End Sub


Private Sub TDBGrid2_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    GridArrayDistanze.ReadData RowBuf, StartLocation, offset, ApproximatePosition
End Sub


Private Sub TDBGrid2_UnboundWriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, WriteLocation As Variant)
    GridArrayDistanze.WriteData RowBuf, WriteLocation
    
    If Not DataClass.UpdateDistanze(DataClass.DS.Fields("Codice").value, _
                             TDBGrid2.Columns(0).value, _
                             TDBGrid2.Columns(1).value) Then
        DataClass.DS.DisplayMessage "L'aggiornamento delle distanze è fallito", 9
    End If
End Sub


