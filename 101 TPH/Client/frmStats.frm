VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funzioni statistiche"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      Caption         =   "Parametri"
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   1320
      Width           =   9255
      Begin VB.ListBox lstZone 
         Height          =   1410
         Left            =   5760
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   1
         Left            =   2880
         Picture         =   "frmStats.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "20"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2880
         Picture         =   "frmStats.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "19"
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Tag             =   "19"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Tag             =   "20"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "&Zone"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   3
         Left            =   3360
         Tag             =   "20"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   11
         Left            =   3360
         Tag             =   "19"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Dalla Data"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Alla Data"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri"
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Width           =   9255
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   3
         Left            =   3480
         Picture         =   "frmStats.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   2
         Left            =   7680
         Picture         =   "frmStats.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Tag             =   "13"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Tag             =   "15"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Tag             =   "14"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   3
         Format          =   "0000000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   6240
         TabIndex        =   21
         Tag             =   "16"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox12 
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Tag             =   "17"
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   327681
         BackColor       =   12648447
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Alla Data Ini&zio Viaggio"
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   28
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N&umero Viaggio"
         Height          =   255
         Index           =   8
         Left            =   4800
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   9
         Left            =   7680
         Tag             =   "14"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dalla Data &Inizio Viaggio"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   8
         Left            =   3960
         Tag             =   "15"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dal &Numero Viaggio"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   7
         Left            =   3480
         Tag             =   "13"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   6
         Left            =   6360
         Tag             =   "17"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Deposito"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   24
         Top             =   1080
         Width           =   735
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
         Left            =   2640
         TabIndex        =   23
         Tag             =   "17"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   8160
         Tag             =   "16"
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.CheckBox chkPrintGraph 
      Caption         =   "Stampa i grafici"
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.ListBox lstStats 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   3720
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmStats.frx":0408
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   120
         Picture         =   "frmStats.frx":050A
         TabIndex        =   8
         Tag             =   "PRINT"
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7560
         Picture         =   "frmStats.frx":060C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmStats.frx":070E
         TabIndex        =   9
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4635
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10901
            Key             =   "Msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10.00"
            Key             =   "Time"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16/12/99"
            Key             =   "Date"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmStats.frx":0810
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmStats.frx":0922
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmStats.frx":0A34
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCurFrame As Integer
Private mvarStats As Variant
Private mvarReports As Variant
Public State As FormState
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Public WithEvents DataClass As CStatistiche
Attribute DataClass.VB_VarHelpID = -1

Private AZO As CZona

Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = MaskEdBox2
            strTitle = "Dalla data inizio viaggio"
        Case 1
            Set M = MaskEdBox3
            strTitle = "Alla data inizio viaggio"
        Case 2
            Set M = MaskEdBox9
            strTitle = "Dalla data inizio viaggio"
        Case 3
            Set M = MaskEdBox7
            strTitle = "Alla data inizio viaggio"
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Long
    
    mvarStats = Array("Grado di servizio per data", _
                      "Grado di servizio per Unità di Trasporto")
    mvarReports = Array("OIL2ST01.RPT", _
                        "OIL2ST02.RPT")
    
    Set DataClass = New CStatistiche
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .ConfigureInterface
        ConfigureObjects Me
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
    End With
    
    With lstStats
        For i = LBound(mvarStats) To UBound(mvarStats)
            .AddItem mvarStats(i)
        Next i
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    Set AZO = New CZona
    With AZO.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    AZO.GetList True
    
    With AZO.DS
        For j = 0 To UBound(.List, 2)
            lstZone.AddItem .ListItem(0, j) & " " & .ListItem(1, j)
            If .ListItem(2, j) = 0 Then
                lstZone.Selected(lstZone.NewIndex) = True
            End If
        Next j
    End With
    
    If lstZone.ListCount > 0 Then
        lstZone.ListIndex = 0
    End If
    
End Sub


Public Sub DoCommand(Cmd As FormCommand)
    On Error GoTo ErrHandler
    
    Dim blnResult As Boolean
    Dim LinkForm As Form
    Dim i As Long
    Dim j As Integer
    Dim Rows As Long
    Dim intIndex As Integer
    Dim strSelectionFormula As String
    Dim strFindValue As String
    ReDim varReportsList(0 To 2, 0 To 0) As Variant
    ReDim varFormulas(0 To 1, 0 To 0) As Variant
    Dim strDescription As String
    
    Select Case Cmd
        Case fcPrint
'            If (State = fsSelection) Then
            If Not DataClass.DS.Fields.Validate Then
                ShowValidationError DataClass.DS
                GoTo ErrHandler
            End If
            intIndex = lstStats.ListIndex
            Select Case intIndex
                Case 0, 1
                    For i = 0 To lstZone.ListCount - 1
                        If lstZone.Selected(i) Then
                            j = j + 1
                            DataClass.DS.Fields("18").FindValue(j) = AZO.DS.ListItem(0, i)
                        End If
                    Next i
                    DataClass.DS.Fields("19").FindValue = MaskEdBox2.Text
                    DataClass.DS.Fields("20").FindValue = MaskEdBox3.Text
                    DataClass.DS.Fields("21").FindValue = DataClass.DS.Fields("19").FindValue
                    DataClass.DS.Fields("22").FindValue = DataClass.DS.Fields("20").FindValue
                Case Else
                    If Not DataClass.DS.ScanClientCond Then
                        GoTo ErrHandler
                    End If
            End Select
            
            strSelectionFormula = DataClass.DS.MakeKeyCond(True, strDescription)
            
            Select Case intIndex
                Case 0, 1
                    strSelectionFormula = strSelectionFormula & " AND ((" & _
                            "{" & _
                            DataClass.DS.Fields("19").fieldName & "}" & _
                            DataClass.DS.Fields("19").operator
                            strFindValue = DataClass.DS.DA.MakeSQL(DataClass.DS.Fields("19").FindValue, DataClass.DS.Fields("19").FldType)
                            strFindValue = Mid$(strFindValue, 2, Len(strFindValue) - 2)
                            strFindValue = "Date(" & Right$(strFindValue, 4) & "," & _
                                            Left$(strFindValue, 2) & "," & _
                                            Mid$(strFindValue, 4, 2) & ")"

                    strSelectionFormula = strSelectionFormula & " " & strFindValue & ")"
                    
                    strSelectionFormula = strSelectionFormula & " AND (" & _
                            "{" & _
                            DataClass.DS.Fields("20").fieldName & "}" & _
                            DataClass.DS.Fields("20").operator
                            strFindValue = DataClass.DS.DA.MakeSQL(DataClass.DS.Fields("20").FindValue, DataClass.DS.Fields("20").FldType)
                            strFindValue = Mid$(strFindValue, 2, Len(strFindValue) - 2)
                            strFindValue = "Date(" & Right$(strFindValue, 4) & "," & _
                                            Left$(strFindValue, 2) & "," & _
                                            Mid$(strFindValue, 4, 2) & ")"

                    strSelectionFormula = strSelectionFormula & " " & strFindValue & ") "
                    
                    strSelectionFormula = strSelectionFormula & " OR (" & _
                            "{" & _
                            DataClass.DS.Fields("21").fieldName & "}" & _
                            DataClass.DS.Fields("21").operator
                            strFindValue = DataClass.DS.DA.MakeSQL(DataClass.DS.Fields("21").FindValue, DataClass.DS.Fields("21").FldType)
                            strFindValue = Mid$(strFindValue, 2, Len(strFindValue) - 2)
                            strFindValue = "Date(" & Right$(strFindValue, 4) & "," & _
                                            Left$(strFindValue, 2) & "," & _
                                            Mid$(strFindValue, 4, 2) & ")"

                    strSelectionFormula = strSelectionFormula & " " & strFindValue & ")"
                    
                    strSelectionFormula = strSelectionFormula & " AND (" & _
                            "{" & _
                            DataClass.DS.Fields("22").fieldName & "}" & _
                            DataClass.DS.Fields("22").operator
                            strFindValue = DataClass.DS.DA.MakeSQL(DataClass.DS.Fields("22").FindValue, DataClass.DS.Fields("22").FldType)
                            strFindValue = Mid$(strFindValue, 2, Len(strFindValue) - 2)
                            strFindValue = "Date(" & Right$(strFindValue, 4) & "," & _
                                            Left$(strFindValue, 2) & "," & _
                                            Mid$(strFindValue, 4, 2) & ")"

                    strSelectionFormula = strSelectionFormula & " " & strFindValue & "))"
                    
                Case Else
            End Select
            
            varReportsList(0, 0) = mvarStats(intIndex)
            varReportsList(1, 0) = mvarReports(intIndex)
            varReportsList(2, 0) = 1
            varFormulas(0, 0) = "PrintGraph"
            varFormulas(1, 0) = IIf(chkPrintGraph.value = vbChecked, "True", "False")
            CRPrint varReportsList, strSelectionFormula, Me, , varFormulas, strDescription
'            End If
        Case fcExit
            Unload Me
    End Select
    Exit Sub
ErrHandler:
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set AZO = Nothing
    Set ErrH = Nothing
    Set DataClass = Nothing
    Set frmStats = Nothing
End Sub

Private Sub lstStats_Click()
    Dim intIndex As Integer
    
    intIndex = lstStats.ListIndex
    Select Case intIndex
        Case 1
            intIndex = 0
    End Select
    If intIndex = mintCurFrame _
            Then Exit Sub
    DataClass.DS.ClearClient
    Frame1(intIndex).Visible = True
    Frame1(mintCurFrame).Visible = False
    mintCurFrame = intIndex
End Sub


