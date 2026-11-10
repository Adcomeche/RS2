VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSetDatiStipendi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INSERIMENTO DATI STIPENDI"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   12930
   Begin VB.Frame Frame4 
      Caption         =   "Consumi"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   9240
      TabIndex        =   22
      Top             =   3960
      Width           =   3675
      Begin TrueDBGrid60.TDBGrid TDBGrid3 
         Height          =   1395
         Left            =   120
         OleObjectBlob   =   "frmSetDatiStipendi.frx":0000
         TabIndex        =   5
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Viaggi"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   0
      TabIndex        =   20
      Top             =   3960
      Width           =   9195
      Begin TrueDBGrid60.TDBGrid TDBGrid2 
         Height          =   1395
         Left            =   120
         OleObjectBlob   =   "frmSetDatiStipendi.frx":A316
         TabIndex        =   4
         Top             =   240
         Width           =   8955
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   12915
      Begin TrueDBGrid60.TDBGrid TDBgrid1 
         Height          =   2775
         Left            =   120
         OleObjectBlob   =   "frmSetDatiStipendi.frx":15F94
         TabIndex        =   3
         Top             =   240
         Width           =   12675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   795
      Left            =   0
      TabIndex        =   15
      Top             =   5820
      Width           =   12915
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmSetDatiStipendi.frx":1E806
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   315
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
         Left            =   3600
         Picture         =   "frmSetDatiStipendi.frx":1E908
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F6 - Apri Viaggio"
         Height          =   375
         Left            =   3480
         Picture         =   "frmSetDatiStipendi.frx":1EA0A
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "OPEN"
         Top             =   240
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1920
         Picture         =   "frmSetDatiStipendi.frx":1EB0C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 - Annulla"
         Height          =   375
         Left            =   1800
         Picture         =   "frmSetDatiStipendi.frx":1EC0E
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "UNDO"
         Top             =   240
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
         Left            =   10740
         Picture         =   "frmSetDatiStipendi.frx":1ED10
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   10620
         Picture         =   "frmSetDatiStipendi.frx":1EE12
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "EXIT"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "F9 - Ricerca"
         Height          =   375
         Left            =   120
         Picture         =   "frmSetDatiStipendi.frx":1EF14
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "FIND"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7320
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F016
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F128
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F23A
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F34C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F45E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F570
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F682
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F794
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F8A6
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1F9B8
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1FACA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1FBDC
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":1FCEE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":20230
            Key             =   "CreateDocs"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":2033A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":2044C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":209F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6615
      Width           =   12930
      _ExtentX        =   22807
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
            Object.Width           =   17357
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12.24"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "09/08/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
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
            Picture         =   "frmSetDatiStipendi.frx":20B04
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":20C16
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetDatiStipendi.frx":20D28
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di selezione"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12915
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2400
         Picture         =   "frmSetDatiStipendi.frx":20E3A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "CAL"
         Top             =   300
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskField10 
         Height          =   255
         Left            =   4620
         TabIndex        =   1
         Tag             =   "10"
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   1140
         TabIndex        =   0
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9600
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.txt"
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   0
         Left            =   8160
         Tag             =   "10"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "&Data"
         Height          =   255
         Index           =   12
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   2880
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Base"
         Height          =   255
         Index           =   9
         Left            =   3780
         TabIndex        =   14
         Top             =   300
         Width           =   855
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
         Left            =   5340
         TabIndex        =   13
         Tag             =   "10"
         Top             =   300
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmSetDatiStipendi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_COLS = 19
Private Const EDIT_MASK_TIME = "  :  "
Private Const EDIT_MASK_DATE = "  /  /  "

Public XAAutisti As XArrayDB
Public XAViaggi As XArrayDB
Public XAAutomezzi As XArrayDB
Public WithEvents DataClass As CAutista
Attribute DataClass.VB_VarHelpID = -1
Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mvarFunction As FunctionCodes

Private ColTags(1 To MAX_COLS) As Variant

Private VVT As CViaggio
Private OLV As CStipendioDettaglio


Public Property Let FunctionCode(NewValue As FunctionCodes)
    mvarFunction = NewValue
End Property
Private Sub SetButtonState()
    Dim lngCount As Long
    
'    lngCount = GridArraySelect.RowCount
    
'    Command1.Enabled = (lngCount > 0) And (State = fsSelection)
'    Command6.Enabled = (XAADD.Count(1) > 0) And (State = fsSelection)
'    Command8.Enabled = (lngCount > 0) And (State = fsSelection)
'    Command2.Enabled = (lngCount > 0) And (State = fsSelection)
End Sub

Private Sub ResetGrid(Grid As TDBGrid, XA As XArrayDB)
    
    If XA.Count(1) > 0 Then
        XA.DeleteRows XA.LowerBound(1), XA.Count(1)
    End If
    Set Grid.Array = XA
    Grid.ReBind

End Sub

Private Sub DisplaySelection(ListStipendi As Variant, Data As Variant)
    Dim i As Long
    
    For i = 0 To UBound(ListStipendi, 2)
        If ListStipendi(2, i) = "01/01/1900" Then
            ListStipendi(2, i) = Format(Data, "dd/mm/yy")
        End If
        If ListStipendi(4, i) = "01/01/1900" Then
            ListStipendi(4, i) = Format(ListStipendi(2, i), "dd/mm/yy")
        Else
            ListStipendi(4, i) = Format(ListStipendi(4, i), "dd/mm/yy")
        End If
        If ListStipendi(5, i) = "01/01/1900" Then
            ListStipendi(5, i) = vbNullString
        Else
            ListStipendi(5, i) = Replace(Format(ListStipendi(5, i), "hh:nn"), ":", ".")
        End If
        If ListStipendi(6, i) = "01/01/1900" Then
            ListStipendi(6, i) = Format(ListStipendi(2, i), "dd/mm/yy")
        Else
            ListStipendi(6, i) = Format(ListStipendi(6, i), "dd/mm/yy")
        End If
        If ListStipendi(7, i) = "01/01/1900" Then
            ListStipendi(7, i) = vbNullString
        Else
            ListStipendi(7, i) = Replace(Format(ListStipendi(7, i), "hh:nn"), ":", ".")
        End If
    Next i
    
    If Not IsEmpty(ListStipendi) Then
        RefreshXArray XAAutisti, TDBGrid1, ListStipendi
        TDBGrid1.bookmark = 0
        TDBGrid1.Enabled = True
    Else
        ResetGrid TDBGrid1, XAAutisti
    End If
End Sub
Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XAAutisti.value(DBGrid.bookmark, 0)
    End If
End Function

Private Function SelectRecords(Optional varSortIndex As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim blnActiveOnly As Boolean
    Dim varList As Variant
    Dim varData As Variant
    Dim strError As String
    
    Me.MousePointer = vbHourglass

    If Not DataClass.DS.Fields.Validate Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    If Not DataClass.DS.ScanClientCond Then
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    varData = mskDateFrom
    If Not IsDate(varData) Then
        strError = "Data non valida"
        GoTo ErrHandler
    End If
    
    If Len(Me.mskField10) > 0 Then
        If Not DataClass.GetListStipendi(varData, varList, mskField10, strError) Then
            GoTo ErrHandler
        End If
    Else
        If Not DataClass.GetListStipendi(varData, varList, , strError) Then
            GoTo ErrHandler
        End If
    End If
    
    DisplaySelection varList, varData
    GoTo EndProc
ErrHandler:
    SelectRecords = False
    Me.MousePointer = vbDefault
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function
Private Sub ClearForm()
    DataClass.DS.ClearClient
    DataClass.DS.DisplayMessage vbNullString
    ResetGrid TDBGrid1, XAAutisti
    ResetGrid TDBGrid2, XAViaggi
    ResetGrid TDBGrid3, XAAutomezzi
'    mskField15 = mstrIDADEDef
End Sub
Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim i As Long
    Dim Rows As Long
    Dim strSelectionFormula As String
    Dim strError As String
    Dim varDocsList As Variant
    ReDim varReportsList(0 To 4, 0 To 2) As Variant
    Dim DAM As CDocumentoAmministrativo
    Dim strErrorMessage As String
    
    Select Case Cmd
        Case fcUndo
            TDBGrid1.Update
            TDBGrid2.Update
            TDBGrid3.Update
            If (State = fsparaminput) Or (State = fsSelection) Then
                ClearForm
                SetState fsparaminput
            End If
        Case fcFind
            TDBGrid1.Update
            TDBGrid2.Update
            TDBGrid3.Update
            If SelectRecords Then
                SetState fsSelection
            End If
        

        Case fcOpen
            Set LinkForm = DisplayForm(fcVIAGVVT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord XAViaggi(TDBGrid2.bookmark, 0)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If

        Case fcExit
            TDBGrid1.Update
            TDBGrid2.Update
            TDBGrid3.Update
            Unload Me
    End Select
End Sub
Private Sub GetRecord(strID As String)
    On Error GoTo ErrHandler
    Dim i As Integer
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        If DataClass.Find Then
        Else
            GoTo ErrHandler
        End If
    End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub
Public Sub SetState(NewState As FormState)
    Dim lngHowMany As Long
    
    DataClass.DS.SetClientState NewState
    Select Case NewState
        Case fsparaminput
            
            SetButtonState
            
            Frame1(1).Enabled = True
            Frame1(2).Enabled = False
            Frame3.Enabled = False
            Frame4.Enabled = False
            
            If Not Loading Then
                mskField10.SetFocus
            End If
                
            DataClass.DS.ValidateClientField mskField10.Tag, mskField10.Text
            FieldUpdateDescription mskField10.Tag, Me, mskField10
            FieldSetDescription mskField10.Tag, Me, mskField10.Text, mskField10
            
            
        Case fsSelection
            lngHowMany = DataClass.DS.HowMany
            SetButtonState
            
            Frame1(1).Enabled = False
            Frame1(2).Enabled = True
            Frame3.Enabled = True
            Frame4.Enabled = True
            
            If Not Loading Then
                If TDBGrid1.Enabled Then
                    TDBGrid1.SetFocus
                End If
            End If
    End Select
End Sub


Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Data"
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
    
    Set DataClass = New CAutista
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    Set OLV = New CStipendioDettaglio
    With OLV.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With

    Set XAViaggi = New XArrayDB
    Set XAAutisti = New XArrayDB
    Set XAAutomezzi = New XArrayDB
    
    SetState fsparaminput

    Loading = False
    
    If mskDateFrom.Visible Then
        mskDateFrom.SetFocus
    End If

End Sub

Private Sub Form_Resize() '    Dim i As Integer
'
'    If Me.WindowState <> vbMinimized Then
'        For i = Frame1.LBound To Frame1.UBound
'            Frame1(i).Move 0, _
'                           Frame1(i).Top, _
'                           Me.Width - 105, _
'                           Me.Height - Frame1(i).Top - StatusBar1.Height - Frame2.Height - 360
'            With TDBgrid1
'                .Move .Left, _
'                      .Top, _
'                      Frame1(2).Width - 240, _
'                      Frame1(2).Height - .Top - 120
'            End With
'        Next i
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set XAAutisti = Nothing
    Set XAViaggi = Nothing
    Set XAAutomezzi = Nothing
    Set frmSetDatiStipendi = Nothing
End Sub


Private Sub TDBgrid1_AfterUpdate()
    On Error GoTo ErrHandler
    
    Dim strError As String
    Dim blnConcErr As Boolean
    Dim strDataInizio As String
    Dim strOraInizio As String
    Dim strDataFine As String
    Dim strOraFine As String
    Dim strFlagAssenza As String
    Dim lngKm As Long
    
    With TDBGrid1
        strDataInizio = .Columns(4).Text
        strOraInizio = Left(.Columns(5).Text, 2) & ":" & Right(.Columns(5).Text, 2)
        strDataFine = .Columns(6).Text
        strOraFine = Left(.Columns(7).Text, 2) & ":" & Right(.Columns(7).Text, 2)
        If Not IsNumeric(TDBGrid1.Columns(8).value) Then
            lngKm = 0
        Else
            lngKm = TDBGrid1.Columns(8).value
        End If
        If Len(.Columns(9).Text) > 0 Then
            strFlagAssenza = .Columns(9).value
        Else
            strFlagAssenza = vbNullString
        End If
    End With
    
    If strDataInizio = EDIT_MASK_DATE Then
        strDataInizio = "01/01/1900"
    End If
    If strDataFine = EDIT_MASK_DATE Then
        strDataFine = "01/01/1900"
    End If
    If strOraInizio = EDIT_MASK_TIME Then
        strOraInizio = "01/01/1900"
    End If
    If strOraFine = EDIT_MASK_TIME Then
        strOraFine = "01/01/1900"
    End If
    
    With OLV.DS
        .Fields.ResetFields
        .Field("Data Inizio Lavoro").NewValue = strDataInizio
        .Field("Data Inizio Lavoro").Updated = True
        .Field("Ora Inizio Lavoro").NewValue = strOraInizio
        .Field("Ora Inizio Lavoro").Updated = True
        .Field("Data Fine Lavoro").NewValue = strDataFine
        .Field("Data Fine Lavoro").Updated = True
        .Field("Ora Fine Lavoro").NewValue = strOraFine
        .Field("Ora Fine Lavoro").Updated = True
        .Field("Km percorsi").NewValue = lngKm
        .Field("Km percorsi").Updated = True
        .Field("Flag motivo assenza").NewValue = strFlagAssenza
        .Field("Flag motivo assenza").Updated = True
        .Field("Codice Autista").FindValue = GetSelectedID(TDBGrid1)
        .Field("Data").FindValue = mskDateFrom
        OLV.Find
        If .HowMany = 0 Then
            .Field("Codice Autista").NewValue = GetSelectedID(TDBGrid1)
            .Field("Codice Autista").Updated = True
            .Field("Data").NewValue = mskDateFrom
            .Field("Data").Updated = True
            OLV.Insert
        Else
            OLV.BeginSave blnConcErr
            OLV.DS.DA.EndTrans ETCOMMIT
        End If
    End With
    GoTo EndProc
ErrHandler:
    If Len(strError) = 0 Then
        strError = "Errore sconosciuto"
    End If
    DataClass.DS.DisplayMessage strError, Error
    
EndProc:
End Sub

Private Sub TDBgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim strDataInizio As String
    Dim strOraInizio As String
    Dim strDataFine As String
    Dim strOraFine As String
    Dim strError As String
    
    With TDBGrid1
        strDataInizio = .Columns(4).Text
        strOraInizio = Left(.Columns(5).Text, 2) & ":" & Right(.Columns(5).Text, 2)
        strDataFine = .Columns(6).Text
        strOraFine = Left(.Columns(7).Text, 2) & ":" & Right(.Columns(7).Text, 2)
    End With
    
    If Not IsDate(strDataInizio) And (strDataInizio <> EDIT_MASK_DATE) Then
        strError = "Data Inizio Lavoro non valida"
        Cancel = True
    ElseIf Not IsDate(strDataFine) And (strDataFine <> EDIT_MASK_DATE) Then
        strError = "Data Fine Lavoro non valida"
        Cancel = True
    ElseIf Not IsDate(strOraInizio) And (strOraInizio <> EDIT_MASK_TIME) Then
        strError = "Ora Inizio Lavoro non valida"
        Cancel = True
    ElseIf Not IsDate(strOraFine) And (strOraFine <> EDIT_MASK_TIME) Then
        strError = "Ora Fine Lavoro non valida"
        Cancel = True
    End If
    If Cancel Then
        DataClass.DS.DisplayMessage strError, Error
    End If
End Sub

Private Sub TDBgrid1_ButtonClick(ByVal ColIndex As Integer)
    Dim LinkForm As Form
    Dim strIDAutista As String
    Dim varData As Variant
    
    Select Case ColIndex
        Case 10
            Set LinkForm = DisplayForm(fcSTIPOLA)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    strIDAutista = GetSelectedID(TDBGrid1)
                    varData = mskDateFrom
                    .mskField10 = strIDAutista
                    .mskDateFrom = varData
                    .GetRecord strIDAutista, varData
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
    End Select
End Sub

Private Sub TDBGrid1_DblClick()
    DoCommand fcOpen
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
    If XAAutisti(bookmark, 9) = "A01" Or XAAutisti(bookmark, 9) = "A02" Then
        RowStyle.BackColor = &HC0FFC0
    Else
        RowStyle.BackColor = &HC0FFFF
    End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
'    Static blnDescending As Boolean
'    Static intColIndex As Integer
'
'    If intColIndex <> ColIndex Then
'        blnDescending = False
'    End If
'
'    If blnDescending Then
'        SelectRecords "D" & ColTags(ColIndex + 1)
'    Else
'        SelectRecords ColTags(ColIndex + 1)
'    End If
'
'    blnDescending = Not blnDescending
'    intColIndex = ColIndex
End Sub




Private Function IsEmptyVariantString(varValue As Variant) As Variant
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(varValue) = 0 Then
        IsEmptyVariantString = True
    Else
        IsEmptyVariantString = False
    End If
End Function

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim varListViaggi As Variant
    Dim varListAutomezzi As Variant
    Dim strIDAutista As String
        
    If (IsNull(LastRow)) Or (LastRow <> TDBGrid1.bookmark) Then
    
        Me.MousePointer = vbHourglass
        TDBGrid2.Update
        TDBGrid3.Update
        strIDAutista = GetSelectedID(TDBGrid1)
        VVT.GetListViaggiAutista Me.mskDateFrom, strIDAutista, varListViaggi
        If Not IsEmpty(varListViaggi) Then
            RefreshXArray XAViaggi, TDBGrid2, varListViaggi
            TDBGrid2.Enabled = True
            TDBGrid2.bookmark = 0
        Else
            ResetGrid TDBGrid2, XAViaggi
        End If
        
        If DataClass.GetListAutomezzi(mskDateFrom, varListAutomezzi, strIDAutista) Then
            If Not IsEmpty(varListAutomezzi) Then
                RefreshXArray XAAutomezzi, TDBGrid3, varListAutomezzi
                TDBGrid3.Enabled = True
                TDBGrid3.bookmark = 0
            Else
                ResetGrid TDBGrid3, XAAutomezzi
            End If
        End If
        
        Me.MousePointer = vbDefault
        
    End If
End Sub



Private Sub TDBGrid2_AfterUpdate()
    On Error GoTo ErrHandler
    
    Dim varKM As Variant
    Dim varPedaggio As Variant
    Dim blnFestivo As Boolean
    Dim strError As String
    Dim blnConcErr  As Boolean
    
    varKM = TDBGrid2.Columns(5).value
    varPedaggio = TDBGrid2.Columns(6).value
    blnFestivo = IIf((TDBGrid2.Columns(7).Text = "0") Or (TDBGrid2.Columns(7).Text = ""), False, True)
    
    With VVT.DS
        .Fields.ResetFields
        .Field("Km definitivi").NewValue = varKM
        .Field("Km definitivi").Updated = True
        .Field("Flag calcolo km definitivi").NewValue = "M"
        .Field("Flag calcolo km definitivi").Updated = True
        .Field("Importo pedaggi").NewValue = varPedaggio
        .Field("Importo pedaggi").Updated = True
        .Field("Flag stato 2").NewValue = IIf(blnFestivo, "F", "''")
        .Field("Flag stato 2").Updated = True
        
        .Field("Numero Viaggio").FindValue = TDBGrid2.Columns(0).value
        If Not VVT.Find Then
            GoTo ErrHandler
        End If
        If Not VVT.BeginSave(blnConcErr, False) Then
            strError = "Si è verificato un errore nell'aggiornamento del viaggio"
            GoTo ErrHandler
        End If
        If blnConcErr Then
            strError = "Si è verificato un errore nell'aggiornamento del viaggio"
            GoTo ErrHandler
        End If
        .DA.EndTrans ETCOMMIT
    End With
    
    GoTo EndProc
    
ErrHandler:
    VVT.DS.DA.EndTrans ETROLLBACK
    DataClass.DS.DisplayMessage strError, Error
EndProc:
End Sub

Private Sub TDBGrid3_AfterUpdate()
    On Error GoTo ErrHandler
    
    Dim varConsumoInterno As Variant
    Dim varConsumoEsterno As Variant
    Dim strError As String
    Dim AAM As CAutomezzo
    
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataClass.DS.DA
        Set .ErrH = DataClass.DS.ErrH
    End With
    
    varConsumoInterno = TDBGrid3.Columns(1).value
    varConsumoEsterno = TDBGrid3.Columns(2).value
    If IsNull(varConsumoInterno) Or (Len(varConsumoInterno) = 0) Then
        varConsumoInterno = 0
    End If
    If IsNull(varConsumoEsterno) Or (Len(varConsumoEsterno) = 0) Then
        varConsumoEsterno = 0
    End If
    If Not AAM.InsertConsumi(TDBGrid3.Columns(0).value, mskDateFrom, varConsumoInterno, varConsumoEsterno) Then
        strError = "Si è verificato un errore nell'inserimento dei consumi"
        GoTo ErrHandler
    End If
    GoTo EndProc
    
ErrHandler:
    DataClass.DS.DisplayMessage strError, Error
    
EndProc:
End Sub

