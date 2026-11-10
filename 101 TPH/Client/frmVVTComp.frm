VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVVTComp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPLETAMENTO VIAGGIO"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   10470
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "frmVVTComp.frx":0000
      TabIndex        =   0
      Top             =   2280
      Width           =   10455
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid2 
      Height          =   1335
      Left            =   0
      OleObjectBlob   =   "frmVVTComp.frx":66A6
      TabIndex        =   9
      Top             =   840
      Width           =   10455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2400
         Picture         =   "frmVVTComp.frx":C3DC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F8 - Conferma Viaggio"
         Height          =   375
         Left            =   2280
         Picture         =   "frmVVTComp.frx":C4DE
         TabIndex        =   11
         Tag             =   "REFRESHLIST"
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
         Index           =   2
         Left            =   8760
         Picture         =   "frmVVTComp.frx":C5E0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
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
         Picture         =   "frmVVTComp.frx":C6E2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   440
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8640
         Picture         =   "frmVVTComp.frx":C7E4
         TabIndex        =   2
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 - Salva"
         Height          =   375
         Left            =   120
         Picture         =   "frmVVTComp.frx":C8E6
         TabIndex        =   1
         Tag             =   "SAVE"
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6015
      Width           =   10470
      _ExtentX        =   18468
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
            Object.Width           =   13018
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14.59"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "19/08/2003"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mskLink 
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1680
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
            Picture         =   "frmVVTComp.frx":C9E8
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTComp.frx":CAFA
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVVTComp.frx":CC0C
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   255
      Left            =   7080
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "5"
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:nn"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   255
      Left            =   9360
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "27"
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskField7 
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Tag             =   "14"
      Top             =   120
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
      Left            =   4680
      TabIndex        =   19
      Tag             =   "6"
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "35"
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tot. Scarico Lt"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   29
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblTotScarico 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   4800
      Width           =   975
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   2
      Left            =   6000
      Tag             =   "35"
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Rif. Est. Viaggio"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   27
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblDifferenza 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9000
      TabIndex        =   26
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Differenza Lt"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   25
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblTotConsegne 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tot. Consegne Lt"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   23
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblTotCarico 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tot. Carico Lt"
      Height          =   255
      Index           =   51
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Data"
      Height          =   255
      Index           =   36
      Left            =   4200
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Unità Trasp."
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   1
      Left            =   10200
      Tag             =   "27"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Km definitivi"
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   0
      Left            =   7920
      Tag             =   "5"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ora carico"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero Viaggio"
      Height          =   255
      Index           =   28
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmVVTComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3

Public WithEvents DataClass As CViaggio
Attribute DataClass.VB_VarHelpID = -1
Private VVP As CViaggioProdotto
Private VCA As CViaggioCarico

Private XA As XArrayDB
Private XALoad As XArrayDB
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private mobjParentForm As Object



Private Sub ComputeDelivery()
    Dim i As Long
    Dim lngTotDel As Long
    
    For i = 0 To XA.UpperBound(1)
        lngTotDel = lngTotDel + TDBGrid1.Columns(5).CellText(i)
    Next i
    lblTotConsegne = lngTotDel
    If Len(lblTotCarico) = 0 Then
        lblTotCarico = 0
    End If
    If Len(lblTotScarico) = 0 Then
        lblTotScarico = 0
    End If
    lblDifferenza = lblTotCarico - lblTotScarico - lblTotConsegne

End Sub

Private Sub ComputeLoad()
    Dim i As Long
    Dim lngTotLoad As Long
    Dim lngTotScarico As Long
    
    For i = 0 To XALoad.UpperBound(1)
        If IsNumeric(TDBGrid2.Columns(3).CellText(i)) Then
            lngTotLoad = lngTotLoad + TDBGrid2.Columns(3).CellText(i)
        End If
        If IsNumeric(TDBGrid2.Columns(9).CellText(i)) Then
            lngTotScarico = lngTotScarico + TDBGrid2.Columns(9).CellText(i)
        End If
    Next i
    lblTotCarico = lngTotLoad
    lblTotScarico = lngTotScarico
    If Len(lblTotConsegne) = 0 Then
        lblTotConsegne = 0
    End If
    lblDifferenza = lblTotCarico - lblTotScarico - lblTotConsegne

End Sub

Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcSave
            Me.MousePointer = vbHourglass
            TDBGrid1.Update
            TDBGrid2.Update
            If UpdateTrip Then
                Unload Me
            Else
                MousePointer = vbDefault
            End If
        Case fcLinkRefresh
            TDBGrid1.Update
            TDBGrid2.Update
            If MsgBox("Confermi il completamento di questo viaggio?", vbExclamation + vbYesNo) = vbYes Then
                DoEvents
                Me.MousePointer = vbHourglass
                If UpdateTrip(True) Then
                    Unload Me
                Else
                    MousePointer = vbDefault
                End If
            End If
        Case fcExit
            Unload Me
    End Select
End Sub


Public Function GetRecord(strIDViaggio As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim varTripList As Variant
    Dim varLoadList As Variant
    Dim lngHowMany As Long
    
    With DataClass.DS
        .Fields("0").FindValue = strIDViaggio
        .Fields("25").ReadOnly = True
        
        If DataClass.Find Then
            If .HowMany > 0 Then
                If .RefreshClient Then
                    ' --- Visualizziamo i dati di carico
                    varLoadList = VCA.GetLoadData(strIDViaggio, True)
                    If IsEmpty(varLoadList) Then
                        GoTo ErrHandler
                    End If
                    RefreshXArray XALoad, TDBGrid2, varLoadList
                    ' ---
                    
                    ' --- Visualizziamo le consegne
                    If Not .DA.QueryCommand(VVP.SQL_GET_TRIP_LIST, _
                                            varTripList, lngHowMany, strIDViaggio) Then
                        GoTo ErrHandler
                    End If
                    RefreshXArray XA, TDBGrid1, varTripList
                    ' ---
                    
                    ComputeLoad
                    ComputeDelivery
                    
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







Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsEdit
    End Select
End Sub


Public Function UpdateTrip(Optional ChangeState As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim strIDViaggio As String
    Dim strData As String
    Dim strOra As String
    Dim strKmRaggio As String
    Dim strDataViaggio As String
    Dim intError As Errors
    
    DataClass.DS.DA.BeginTransaction
    
    strIDViaggio = DataClass.DS.Fields("0").value
    strDataViaggio = DataClass.DS.Fields("6").value
    For i = 0 To XALoad.UpperBound(1)
        If Not VCA.UpdateLoadData(XALoad(i, 3), XALoad(i, 4), XALoad(i, 5), XALoad(i, 6), _
                                  XALoad(i, 7), XALoad(i, 8), strIDViaggio, XALoad(i, 0), _
                                  XALoad(i, 1), XALoad(i, 2), XALoad(i, 9)) Then
            intError = LoadError
            GoTo ErrHandler
        End If
    Next i
    
    For i = 0 To XA.UpperBound(1)
        
        If IsNull(XA(i, 6)) Then
            strData = strDataViaggio
        Else
            If Len(XA(i, 6)) > 0 Then
                strData = XA(i, 6)
            Else
                strData = strDataViaggio
            End If
        End If
        
        If IsNull(XA(i, 7)) Then
            strOra = "00.00"
        Else
            If Len(XA(i, 7)) > 0 Then
                strOra = XA(i, 7)
            Else
                strOra = "00.00"
            End If
        End If
        If Not IsDate(strOra) Then
            MsgBox "Controllare gli orari di consegna", vbCritical + vbOKOnly, "Errore"
            GoTo ErrHandler
        End If
        
        If IsNull(XA(i, 8)) Then
            strKmRaggio = "0"
        Else
            If Len(XA(i, 8)) = 0 Then
                strKmRaggio = "0"
            Else
                strKmRaggio = XA(i, 8)
            End If
        End If
        If Not VVP.CompleteRow(XA(i, 5), strData, strOra, _
                              strKmRaggio, XA(i, 9), XA(i, 10), strIDViaggio, _
                              XA(i, 0), XA(i, 1), XA(i, 11)) Then
            intError = DeliveryError
            GoTo ErrHandler
        End If
    Next i
    
    ' --- Imposto i km definitivi del viaggio
    Dim blnConcErr As Boolean
    DataClass.DS.ScanClient smUpdate
    DataClass.DS.Fields("6").Updated = False
    DataClass.BeginSave blnConcErr
    DataClass.DS.DA.EndTrans ETCOMMIT
    ' ---
    
    ' --- Imposto lo stato del viaggi a Completato
    If ChangeState Then
        If Not DataClass.SetStato("C") Then
            GoTo ErrHandler
        End If
    End If
    ' ---
    
    DataClass.DS.DA.EndTrans ETCOMMIT
    UpdateTrip = True
    Exit Function
ErrHandler:
    Select Case intError
        Case DeliveryError
            MsgBox "Si è verificato un errore durante il completamento. " & _
                   "Consegna: " & XA(i, 0) & " Scarico: " & XA(i, 1), vbCritical + vbOKOnly
        Case LoadError
            MsgBox "Si è verificato un errore durante l'inserimento del carico. " & _
                   "Prodotto: " & XALoad(i, 0), vbCritical + vbOKOnly
    End Select
    DataClass.DS.DA.EndTrans ETROLLBACK
    UpdateTrip = False
End Function

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
    
    Set DataClass = New CViaggio
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set VCA = New CViaggioCarico
    With VCA.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    DataClass.DS.Fields("14").ReadOnly = True
    DataClass.DS.Fields("6").ReadOnly = True
    SetState fsNone
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ErrH = Nothing
    Set DataClass = Nothing
    If Not mobjParentForm Is Nothing Then
        With mobjParentForm
            .Enabled = True
            .DoCommand fcFind
        End With
    End If
    Set XA = Nothing
    Set XALoad = Nothing
    Set VVP = Nothing
    Set VCA = Nothing
    Set frmVVTComp = Nothing
End Sub





Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property





Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 5 Then
        ComputeDelivery
    End If
End Sub

Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
    If (ColIndex = 3) Or (ColIndex = 9) Then
        ComputeLoad
    End If
End Sub

