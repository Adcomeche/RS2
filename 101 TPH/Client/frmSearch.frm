VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RICERCA"
   ClientHeight    =   5460
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   9285
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid60.TDBGrid tdbResult 
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frmSearch.frx":0442
      TabIndex        =   12
      Tag             =   "SEARCH"
      Top             =   1560
      Width           =   9015
   End
   Begin TrueDBGrid60.TDBGrid TDBCond 
      Height          =   585
      Index           =   1
      Left            =   120
      OleObjectBlob   =   "frmSearch.frx":2E29
      TabIndex        =   13
      Top             =   960
      Width           =   9015
   End
   Begin TrueDBGrid60.TDBGrid TDBCond 
      Height          =   3015
      Index           =   2
      Left            =   120
      OleObjectBlob   =   "frmSearch.frx":581D
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Risultato della ricerca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   5400
         Picture         =   "frmSearch.frx":9DC5
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 - Stampa"
         Height          =   375
         Left            =   5280
         Picture         =   "frmSearch.frx":9EC7
         TabIndex        =   11
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
         Index           =   0
         Left            =   3720
         Picture         =   "frmSearch.frx":9FC9
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "F4 - Azzera"
         Height          =   375
         Left            =   3600
         Picture         =   "frmSearch.frx":A0CB
         TabIndex        =   6
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
         Left            =   2040
         Picture         =   "frmSearch.frx":A1CD
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "F6 - Apri"
         Height          =   375
         Left            =   1920
         Picture         =   "frmSearch.frx":A2CF
         TabIndex        =   5
         Tag             =   "OPEN"
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
         Index           =   5
         Left            =   360
         Picture         =   "frmSearch.frx":A3D1
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "F9 - Cerca"
         Height          =   375
         Left            =   240
         Picture         =   "frmSearch.frx":A4D3
         TabIndex        =   4
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
         Index           =   2
         Left            =   7560
         OLEDropMode     =   1  'Manual
         Picture         =   "frmSearch.frx":A5D5
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmSearch.frx":A6D7
         TabIndex        =   7
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5085
      Width           =   9285
      _ExtentX        =   16378
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
            Object.Width           =   10927
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16.10"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11/12/01"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3615
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6376
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ricerca &semplice"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ricerca &avanzata"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mskMaxRecords 
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Restituisci al massimo"
      Height          =   255
      Left            =   5640
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "records"
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private XA As XArrayDB
Private XASimple As XArrayDB
Public GridArraySearch As ArrayGridCls
Private mobjParentForm As Form
'Private mobjDA As CDataAccess
'Private mobjDA As CConnect
Private mobjDC As Object
Private strOperators(1 To 8) As String
Private mintColNum As Integer
Private mintCurGrid As Integer
Private mintActiveControl As Integer
Private mintFindCondCount As Integer
Private mstrFieldKey As String
Private mblnSearchOK As Boolean
Private mblnShowPrint As Boolean

Public LinkControlIndex As Long
Public ReportsList As Variant
Private Sub MakeCondStrings(SearchType As Integer, _
                            ByRef strCond As String, _
                            ByRef strOrder As String, _
                            Optional CrystalFlag As Boolean = False, _
                            Optional SortFields As Variant)
    Dim Col As Integer
    Dim Row As Integer
    Dim strFieldName As String
    Dim strFieldIndex As String
    Dim strCondValue As String
    Dim strOperator As String
    Dim strDirection As String
    Dim strFindValue As String
    Dim strJollyChar As String
    Dim varConditions As Variant
    Dim blnOrderFlag As Boolean
    Dim blnCondFlag As Boolean
    Dim FieldType As Byte
    Dim i As Integer
    Dim j As Integer
    
    If Not CrystalFlag Then
        strCond = " WHERE "
        strOrder = " ORDER BY "
    End If
    
    If SearchType = 1 Then          ' Ricerca semplice
        blnOrderFlag = False
        For Row = XASimple.LowerBound(2) To XASimple.UpperBound(2) - 1
            strCondValue = XASimple.value(0, Row)
            strFieldIndex = mobjDC.SearchListFields(Row + 1)
            If strCondValue <> vbNullString Then
                varConditions = ParseString(strCondValue, ";")
                blnCondFlag = True
                strFieldName = mobjDC.DS.Fields(strFieldIndex).fieldName
                If CrystalFlag Then
                    strFieldName = "{" & strFieldName & "}"
                End If
                FieldType = mobjDC.DS.Fields(strFieldIndex).FldType
                strCond = strCond & "("
                For i = LBound(varConditions) To UBound(varConditions)
                    If FieldType = 0 Then
                        strCond = strCond & strFieldName & strOperators(7)
                    Else
                        strCond = strCond & strFieldName & " = "
                    End If
'                    strFindValue = DA.MakeSQL(varConditions(i), FieldType)
                    strFindValue = DataConnect.MakeSQL(varConditions(i), FieldType)
                    If CrystalFlag And FieldType = 2 Then
                        strFindValue = "Date(" & Right$(strFindValue, 4) & "," & _
                                        Left$(strFindValue, 2) & "," & _
                                        Mid$(strFindValue, 4, 2) & ")"
                    End If
                    If FieldType = 0 Then
                         strFindValue = Left(strFindValue, Len(strFindValue) - 1) & DataConnect.SQLJollyChar & "'"
                    End If
                    strCond = strCond & strFindValue
                    If i <> UBound(varConditions) Then
                        strCond = strCond & " OR "
                    End If
                Next i
                strCond = strCond & ")"
                
                strCond = strCond & " AND "
            End If
        Next Row
        If blnCondFlag Then
            If Right(strCond, 5) = " AND " Then
                strCond = Left$(strCond, Len(strCond) - 5)
            End If
        End If
    Else                            ' Ricerca avanzata
        For Row = XA.LowerBound(1) To XA.UpperBound(1)
            Me.TDBCond(2).Bookmark = Row
            strFieldIndex = TDBCond(2).Columns(0).CellText(Row)
            If strFieldIndex <> vbNullString Then
                strFieldName = mobjDC.DS.Fields(CInt(strFieldIndex)).fieldName
                If CrystalFlag Then
                    strFieldName = "{" & strFieldName & "}"
                End If
                strCondValue = TDBCond(2).Columns(2).CellValue(Row)
                strOperator = TDBCond(2).Columns(1).CellValue(Row)
                strDirection = TDBCond(2).Columns(3).CellValue(Row)
                
                If (strCondValue <> vbNullString) And (strOperator <> vbNullString) Then
                    blnCondFlag = True
                    varConditions = ParseString(strCondValue, ";")
                    
                    FieldType = mobjDC.DS.Fields(CInt(strFieldIndex)).FldType
                    strCond = strCond & "("
                    For i = LBound(varConditions) To UBound(varConditions)
                        strCond = strCond & strFieldName & strOperators(strOperator)
                        strCondValue = varConditions(i)
                        If CrystalFlag Then
                            strJollyChar = "*"
                        Else
'                            strJollyChar = DA.SqlJollyChar
                            strJollyChar = DataConnect.SQLJollyChar
                        End If
                        Select Case TDBCond(2).Columns(1).CellValue(Row)
                            Case Is < 7
                            Case 7
                                strCondValue = strJollyChar & strCondValue & strJollyChar
                            Case 8
                                strCondValue = strCondValue & strJollyChar
                        End Select
'                        strCondValue = DA.MakeSQL(strCondValue, FieldType)
                        strCondValue = DataConnect.MakeSQL(strCondValue, FieldType)
                        If CrystalFlag And FieldType = 2 Then
                            strCondValue = "Date(" & Right$(strCondValue, 4) & "," & _
                                            Left$(strCondValue, 2) & "," & _
                                            Mid$(strCondValue, 4, 2) & ")"
                        End If
                        strCond = strCond & strCondValue
                        If i <> UBound(varConditions) Then
                            strCond = strCond & " OR "
                        End If
                    Next i
                    strCond = strCond & ")"
                                        
                    If Row <> XA.UpperBound(1) Then
                        strCond = strCond & " AND "
                    End If
                End If
                
                Select Case UCase(strDirection)
                    Case "CRESCENTE"
                        If CrystalFlag Then
                            ReDim Preserve SortFields(0 To j) As Variant
                            SortFields(j) = strFieldName
                            j = j + 1
                        Else
                            strOrder = strOrder & strFieldName & " ASC, "
                            blnOrderFlag = True
                        End If
                    Case "DECRESCENTE"
                        If CrystalFlag Then
                            ReDim Preserve SortFields(0 To j) As Variant
                            SortFields(j) = strFieldName
                            j = j + 1
                        Else
                            strOrder = strOrder & strFieldName & " DESC, "
                            blnOrderFlag = True
                        End If
                End Select
            End If
        Next Row
    End If
    ' ---
    
    If Right$(strCond, 5) = " AND " Then
        strCond = Left$(strCond, Len(strCond) - 4)
    End If
    If Not blnCondFlag Then
        strCond = vbNullString
    End If
    If blnOrderFlag Then
        strOrder = Left(strOrder, Len(strOrder) - 2)
        strCond = strCond & strOrder
    End If

End Sub

Private Function ParseString(L As String, Delimiter As String) As Variant
    On Error GoTo ErrHandler
    
    ReDim varBuffer(0) As Variant
    Dim FromChar As Integer
    Dim ToChar As Integer
    Dim i As Integer
    
    i = 0
    FromChar = 1
    ToChar = InStr(FromChar, L, Delimiter)
    If ToChar = 0 Then
        varBuffer(0) = L
    Else
    varBuffer(i) = Mid(L, FromChar, ToChar - FromChar)
    Do While FromChar < Len(L)
        FromChar = ToChar + 1
        ToChar = InStr(FromChar, L, Delimiter)
        If ToChar > 0 Then
            i = i + 1
            ReDim Preserve varBuffer(i)
            varBuffer(UBound(varBuffer)) = Trim(Mid(L, FromChar, ToChar - FromChar))
        Else
            i = i + 1
            ReDim Preserve varBuffer(i)
            varBuffer(UBound(varBuffer)) = Trim(Mid(L, FromChar, Len(L) - FromChar + 1))
        End If
    Loop
    End If
    ParseString = varBuffer
    Exit Function
ErrHandler:
    ParseString = Empty
End Function

Public Property Let ActiveControlTabIndex(NewValue As Integer)
    mintActiveControl = NewValue
End Property

Public Property Set DataClass(DC As Object)
    If Not DC Is Nothing Then
        Set mobjDC = DC
    End If
End Property

Public Sub DoCommand(Cmd As FormCommand)
    On Error GoTo ErrHandler
    
    Dim varResult As Variant
    Dim i As Integer
    Dim strSelectionFormula As String
    Dim strOrder As String
    Dim strTag As String
    ReDim varSortFields(0) As Variant
    
    Select Case Cmd
        Case fcFind
            If Not cmdFind.Enabled Then
                GoTo ErrHandler
            End If
            PerformSearch mintCurGrid
            If Me.Visible Then
                If mintCurGrid = 2 Then
                    TabStrip1.Tabs(1).Selected = True
                Else
                    tdbResult.SetFocus
                End If
            End If
        Case fcOpen
            If Not cmdOpen.Enabled Then
                GoTo ErrHandler
            End If
            
            TDBCond(1).Bookmark = TDBCond(1).Bookmark
            TDBCond(2).Bookmark = TDBCond(2).Bookmark
            
            varResult = GridArraySearch.value(tdbResult.Bookmark, 0)
            
            If mstrFieldKey <> vbNullString Then
                With mobjParentForm.DataClass.DS
                    If .ValidateClientField(mstrFieldKey, varResult) Then
                        ' --- E' una ricerca per chiave primaria ? ---
                        If (mstrFieldKey = .KeyFields(1)) And (.KeyFields.Count = 1) Then
                            With mobjParentForm
                                If .GetRecord(varResult) Then
                                    .Enabled = True
                                    .SetState fsEdit
                                Else
                                    If .State = fsNone Then
                                        .Enabled = True
                                        .SetState fsNone
                                    End If
                                End If
                            End With
                        Else
                            ' --- Ricerca su tabella collegata ---
                            Dim C As Control
                            For Each C In mobjParentForm.Controls
                                strTag = C.Tag
                                If strTag = mstrFieldKey Then
                                    Select Case UCase(TypeName(C))
                                        Case "MASKEDBOX"
                                            C = varResult
                                            FieldUpdateDescription strTag, mobjParentForm, C
                                            FieldSetDescription strTag, mobjParentForm, CStr(varResult), C
                                    End Select
                                End If
                            Next C
                            SetParentFocus
                            ' ---
                        End If
                        mblnSearchOK = True
                    End If
                End With
            Else
                mobjParentForm.Controls(LinkControlIndex).SelText = varResult
            End If
            Unload Me
'            Set frmSearch = Nothing
        Case fcPrint
            If mblnShowPrint Then
                MakeCondStrings mintCurGrid, strSelectionFormula, strOrder, True, varSortFields
                CRPrint ReportsList, strSelectionFormula, Me, varSortFields
            End If
        Case fcExit
            Unload Me
        Case fcUndo
            ClearUnboundGrid tdbResult, GridArraySearch, mintColNum
            If mintCurGrid = 1 Then
                For i = XASimple.LowerBound(2) To XASimple.UpperBound(2)
                    XASimple.value(0, i) = vbNullString
                Next i
                SetFixedCond 1
                With TDBCond(1)
                    .ReBind
                    .FirstRow = 1
                    .SetFocus
                    .Col = 0
                End With
            Else
                TDBCond(2).Bookmark = 1
                If mintFindCondCount = 0 Then
                    XA.Clear
                Else
                    For i = mintFindCondCount + 1 To XA.UpperBound(1)
                        XA.Delete 1, i
                    Next i
                End If
                With TDBCond(2)
                    .ReBind
                    .SetFocus
                    .Col = 0
                End With
            End If
            StatusBar1.Panels(2).Text = vbNullString
    End Select
ErrHandler:
    Exit Sub
End Sub

'Public Property Set DataAccess(DC As CConnect) '(DA As CDataAccess)
''    If Not DA Is Nothing Then
''        Set mobjDA = DA
''    End If
'    If Not DC Is Nothing Then
'        Set mobjDA = DC
'    End If
'End Property

Public Property Let FieldKey(NewValue As String)
    mstrFieldKey = NewValue
End Property

Public Property Set ParentForm(F As Form)
    If Not F Is Nothing Then
        Set mobjParentForm = F
    End If
End Property

Private Sub PerformSearch(SearchType As Integer)
    On Error GoTo ErrHandler
    
    Dim strCond As String
    Dim strOrder As String
    Dim strQ As String
    Dim varResult As Variant
    Dim lngHowMany As Long
    
    ' --- Compongo la condizione per la ricerca ---
    
    MakeCondStrings SearchType, strCond, strOrder
    
    strQ = mobjDC.SQL_GET_SEARCH_LIST & strCond & mobjDC.SQL_GET_SEARCH_LIST_GROUPBY
    
    ' --- Ripulisco la grid dei risultati ---
    If GridArraySearch.RowCount > 0 Then
        GridArraySearch.ClearData
        
        tdbResult.ReBind
    End If
    ' ---
    
    ' --- Eseguo la ricerca ---
    lngHowMany = CLng(mskMaxRecords)
'    If DA.QueryData(strQ, varResult, lngHowMany) Then
    If mobjDC.DS.DA.QueryData(strQ, varResult, lngHowMany) Then
        If lngHowMany > 0 Then
            RefreshUnboundGrid tdbResult, varResult, GridArraySearch, _
                               lngHowMany, mintColNum
        End If
        StatusBar1.Panels(2).Text = "Records trovati: " & lngHowMany
    End If
    ' ---

ErrHandler:
    Select Case Err.Number
         
        Case 9
            Resume Next
    End Select
End Sub


Private Sub SetFixedCond(Index As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim strFieldIndex As String
    Dim strConditions As String
    
    If Index = 1 Then
        With mobjDC
            For i = 0 To mintColNum
                strFieldIndex = .SearchListFields(i + 1)
                If strFieldIndex <> .DS.KeyFields(.DS.KeyFields.Count) Then
                    strConditions = vbNullString
                    For j = 1 To .DS.Fields(strFieldIndex).FindValuesCount
                        If CStr(.DS.Fields(strFieldIndex).FindValue(j)) <> vbNullString Then
'                            XASimple(i + 1, 1) = .DS.Fields(strFieldIndex).FindValue
                            strConditions = strConditions & .DS.Fields(strFieldIndex).FindValue(j) & ";"
                        End If
                    Next j
                    If Len(strConditions) > 0 Then
                        strConditions = Left$(strConditions, Len(strConditions) - 1)
                        XASimple(0, i) = strConditions
                    End If
                End If
            Next i
        End With
    Else
        With mobjDC.DS
            mintFindCondCount = 0
            For i = 1 To .Fields.Count
                If CStr(.Fields(i).FindValue) <> vbNullString Then
                    strConditions = vbNullString
                    mintFindCondCount = mintFindCondCount + 1
                    XA.ReDim 1, mintFindCondCount, 0, 3
                    XA(mintFindCondCount, 0) = .Fields(i).Name
                    XA(mintFindCondCount, 1) = 1
                    For j = 1 To .Fields(i).FindValuesCount
                        strConditions = strConditions & .Fields(i).FindValue(j) & ";"
                    Next j
                    If Len(strConditions) > 0 Then
                        strConditions = Left$(strConditions, Len(strConditions) - 1)
                    End If
                    XA(mintFindCondCount, 2) = strConditions
                End If
            Next i
        End With
    End If
End Sub

Private Sub SetParentFocus()
    Dim C As Control
    
    mobjParentForm.Enabled = True
    On Error Resume Next
    For Each C In mobjParentForm.Controls
        If (C.TabIndex = mintActiveControl) Then
            If Err = 0 Then
                C.SetFocus
                SendKeys "{TAB}"
                Exit For
            End If
        End If
    Next C
End Sub

Public Property Let ShowPrint(NewValue As Boolean)
    mblnShowPrint = NewValue
    Picture1(3).Visible = NewValue
    Command5.Visible = NewValue
End Property

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim strFieldIndex As String
    Dim strFieldName As String
    Dim C As Column
    Dim CCond As Column
    Dim Col As Integer
    Dim vItem As ValueItem
    
    mintCurGrid = 1
    Set XA = New XArrayDB
    XA.ReDim 1, 1, 0, 3
    Set TDBCond(2).Array = XA
    
    Set XASimple = New XArrayDB
    Set TDBCond(1).Array = XASimple
    mskMaxRecords = MAX_RECORDS
    Set GridArraySearch = New ArrayGridCls
    
    ConfigureObjects Me
    
    strOperators(1) = " = "
    strOperators(2) = " <> "
    strOperators(3) = " > "
    strOperators(4) = " >= "
    strOperators(5) = " < "
    strOperators(6) = " <= "
    strOperators(7) = " LIKE "
    strOperators(8) = " LIKE "
    
    With mobjDC
        mintColNum = .SearchListFields.Count - 1
        For Col = 0 To mintColNum
            If Col > 1 Then
                Set C = tdbResult.Columns.Add(Col)
                Set CCond = TDBCond(1).Columns.Add(Col)
            Else
                Set C = tdbResult.Columns(Col)
                Set CCond = TDBCond(1).Columns(Col)
            End If
            C.Visible = True
            CCond.Width = C.Width
            CCond.Visible = True
            strFieldIndex = .SearchListFields(Col + 1)
            strFieldName = .DS.Fields(strFieldIndex).Name
            C.Caption = strFieldName
            CCond.Caption = strFieldName
            If Col <> mintColNum Then
                C.DividerStyle = dbgNoDividers
            Else
                C.DividerStyle = dbgBlackLine
            End If
            CCond.DividerStyle = dbgBlackLine
            CCond.Locked = False
            XASimple.ReDim 0, 0, 0, Col + 1
            'XASimple(0, Col + 1) = strFieldName
            
            If UCase(.DS.Fields(strFieldIndex).Name) = "ANNULLATO" Then
                With C.ValueItems
                    .Translate = True
                    .Validate = True
                    Set vItem = New ValueItem
                    vItem.value = "0"
                    vItem.DisplayValue = "NO"
                    .Add vItem
                    Set vItem = New ValueItem
                    vItem.value = "1"
                    vItem.DisplayValue = "SI"
                    .Add vItem
                    Set vItem = New ValueItem
                    vItem.value = "-1"
                    vItem.DisplayValue = "SI"
                    .Add vItem
                    vItem.value = "true"
                    vItem.DisplayValue = "SI"
                    .Add vItem
                    vItem.value = "vero"
                    vItem.DisplayValue = "SI"
                    .Add vItem
                    vItem.value = "false"
                    vItem.DisplayValue = "NO"
                    .Add vItem
                    vItem.value = "falso"
                    vItem.DisplayValue = "NO"
                    .Add vItem
                End With
            End If
        Next Col
        SetFixedCond 1
        SetFixedCond 2

        TDBCond(1).ReBind
        TDBCond(2).Refresh
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParentForm Is Nothing Then
        mobjParentForm.Enabled = True
        If mstrFieldKey <> vbNullString Then
            If mstrFieldKey <> mobjParentForm.DataClass.DS.KeyFields(1) Then
                If Not mblnSearchOK Then
                    Dim C As Control
                    For Each C In mobjParentForm.Controls
                        If C.Tag = mstrFieldKey Then
                            Select Case UCase(TypeName(C))
                                Case "MASKEDBOX"
                                    C = vbNullString
                                    C.SetFocus
                                Case "LABEL"
                                    C = vbNullString
                            End Select
                        End If
                    Next C
                End If
            End If
        Else
            mobjParentForm.Controls(LinkControlIndex).SetFocus
        End If
    End If
    Set frmSearch = Nothing
End Sub
Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = mintCurGrid _
            Then Exit Sub
        TDBCond(TabStrip1.SelectedItem.Index).Visible = True
        TDBCond(mintCurGrid).Visible = False
        mintCurGrid = TabStrip1.SelectedItem.Index
        If mintCurGrid = 1 Then
            tdbResult.Visible = True
        Else
            tdbResult.Visible = False
        End If
End Sub
Private Sub TDBCond_AfterInsert(Index As Integer)
    cmdFind.Enabled = True
    cmdOpen.Enabled = True
End Sub
Private Sub TDBCond_AfterUpdate(Index As Integer)
    cmdFind.Enabled = True
    cmdOpen.Enabled = True
End Sub
Private Sub TDBCond_BeforeColEdit(Index As Integer, ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Dim strFieldIndex As String
    Dim i As Integer
    
    cmdFind.Enabled = False
    cmdOpen.Enabled = False
    If Index = 1 Then
        With mobjDC
            strFieldIndex = .SearchListFields(ColIndex + 1)
            If Not IsEmpty(.DS.Fields(strFieldIndex).FindValue) Then
                Cancel = True
            End If
        End With
    End If
End Sub

Private Sub TDBCond_BeforeColUpdate(Index As Integer, ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Select Case Index
        Case 1
            If ColIndex = 1 Then
                TDBCond(Index).Columns(ColIndex).value = UCase(TDBCond(Index).Columns(ColIndex).value)
            End If
        Case 2
            If mintFindCondCount > 0 Then
                If (TDBCond(Index).Bookmark <= mintFindCondCount) Then
                    If TDBCond(Index).AddNewMode <> dbgAddNewPending Then
                        Cancel = True
                    End If
                End If
            End If
            If ColIndex = 2 Then
                TDBCond(Index).Columns(ColIndex).value = UCase(TDBCond(Index).Columns(ColIndex).value)
            End If
    End Select
End Sub
Private Sub TDBCond_LeftColChange(Index As Integer, ByVal SplitIndex As Integer)
    If Index = 1 Then
        tdbResult.LeftCol = TDBCond(Index).LeftCol
    End If
End Sub
Private Sub TDBCond_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If Index = 1 Then
    If LastCol >= 0 Then
        TDBCond(Index).Columns(LastCol).value = UCase(TDBCond(Index).Columns(LastCol).value)
        XASimple.value(0, LastCol) = (TDBCond(Index).Columns(LastCol).value)
        cmdFind.Enabled = True
        cmdOpen.Enabled = True
    End If
End If
End Sub
Private Sub tdbResult_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DoEvents
        DoCommand fcOpen
    End If
End Sub
Private Sub tdbResult_LeftColChange(ByVal SplitIndex As Integer)
    TDBCond(1).LeftCol = tdbResult.LeftCol
End Sub

