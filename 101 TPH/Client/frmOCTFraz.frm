VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOCTFraz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRAZIONAMENTO ORDINE"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   2640
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
         Picture         =   "frmOCTFraz.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F6 - Imposta Ord. aperto"
         Height          =   375
         Left            =   1800
         Picture         =   "frmOCTFraz.frx":0102
         TabIndex        =   23
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2775
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
         Picture         =   "frmOCTFraz.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
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
         Left            =   240
         Picture         =   "frmOCTFraz.frx":0306
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7560
         Picture         =   "frmOCTFraz.frx":0408
         TabIndex        =   15
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Esegui"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTFraz.frx":050A
         TabIndex        =   14
         Tag             =   "SAVE"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3510
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
            TextSave        =   "17:04"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11/06/2013"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   6000
         Picture         =   "frmOCTFraz.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskIDOrdine 
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskNumFraz 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskQuantita 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         Format          =   "0"
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
               Picture         =   "frmOCTFraz.frx":070E
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTFraz.frx":0820
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTFraz.frx":0932
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskTotale 
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   25
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskIDRiga 
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   255
         Left            =   4680
         TabIndex        =   3
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
      Begin VB.Label Label1 
         Caption         =   "Data Richiesta"
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Riga"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Totale"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   2760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "="
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "x"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Quantità"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Numero Partizioni"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ordine"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmOCTFraz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private XA As XArrayDB
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Form
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public WithEvents DataClass As COrdine
Attribute DataClass.VB_VarHelpID = -1

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Private Sub ClearForm()
    DataClass.DS.ClearClient
End Sub



Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim blnConcErr As Boolean

    
    Select Case Cmd
        Case fcSave
            blnResult = InsertRecord
            If blnResult Then
                
            End If
        Case fcOpen
        
            With DataClass.DS
                .Field("Numero Ordine").FindValue = Me.mskIDOrdine
                DataClass.Find
                .Field("Raggruppamento Utente 5").NewValue = "S"
                .Field("Raggruppamento Utente 5").Updated = True
                DataClass.BeginSave blnConcErr
                .DA.EndTrans ETCOMMIT
                .EndSave
            End With
            DoCommand fcExit
        Case fcExit
            Unload Me
    End Select
End Sub






Private Function InsertRecord() As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim strError As String
    Dim strNewID As String
    
    If (Not IsNumeric(Me.mskNumFraz)) Or (Not IsNumeric(Me.mskQuantita)) Then
        strError = "Dati non validi"
        GoTo ErrHandler
    End If
    If Not IsDate(Me.mskData) Then
        strError = "Data non valida"
    End If
    
    DataClass.DS.DisplayMessage "Elaborazione in corso... "
    
    DataClass.DS.DA.BeginTransaction
    
    For i = 1 To Me.mskNumFraz
        If Not DataClass.CopyOrder(Me.mskIDOrdine, True, strNewID, True, Me.mskIDRiga, Me.mskQuantita, Me.mskData, strError) Then
            GoTo ErrHandler
        End If
    Next i
    
    DataClass.DS.DA.EndTrans ETCOMMIT
    DataClass.DS.DisplayMessage "Elaborazione terminata correttamente."
    
    GoTo EndProc
ErrHandler:
    DataClass.DS.DA.EndTrans ETROLLBACK
    DataClass.DS.DisplayMessage strError, Error
    InsertRecord = False
    Exit Function
EndProc:
    InsertRecord = True
End Function



Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsEdit
        Case fsInsert
    End Select
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskData
            strTitle = "Data consegna richiesta"
    End Select
    GetCalendarDate M, strTitle
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
    
    Set DataClass = New COrdine
    
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtFil1S
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    SetState fsEdit
    
    Loading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CheckExit(State)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set XA = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        mobjParentForm.Enabled = True
        mobjParentForm.DoCommand fcFind
    End If
    Set frmOCTFraz = Nothing
End Sub









Private Sub mskNumFraz_Change()
    If IsNumeric(mskNumFraz) And IsNumeric(mskQuantita) Then
        mskTotale = mskNumFraz * mskQuantita
    End If
End Sub


Private Sub mskQuantita_Change()
    If IsNumeric(mskNumFraz) And IsNumeric(mskQuantita) Then
        mskTotale = mskNumFraz * mskQuantita
    End If
End Sub


