VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const DATE_MASK = "__/__/__"
Private Const TIME_MASK = "__:__"

Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private mobjDA As CDataAccess
Private mcolFields As ColFields
Private mobjErrH As CErrorHandler
Private mvarRecBuf As Variant
Private TimeStamp As Variant
Private mvarList As Variant
Private mobjMultiSoc As Object
Private mvarHowMany As Long
Private mstrCancelField As String
Private mcolKeyFields As Collection
Private mobjFormHandle As Form

Public FormType As FormTypes

Public Enum ErrCodes
    UnknownError = vbObjectError + 513
    UpdateFail = vbObjectError + 514
    InsertFail = vbObjectError + 515
    SelectFail = vbObjectError + 516
End Enum
Public Enum ScanMode
    smUpdate = 0
    smInsert = 1
End Enum
Public Enum FormState
    fsNone = 0
    fsEdit = 1
    fsInsert = 2
    fsParamInput = 3
    fsSelection = 4
    fsCopy = 5
End Enum
Public Enum FormTypes
    ftEdtRcd1 = 0
    ftEdtRcd1S = 1
    ftEdtFil1 = 2
    ftEdtFil1S = 3
    ftEdtFil2 = 4
    ftEdtTrn = 5
End Enum
Public Enum FormCommand
    fcAddNew = 0
    fcSave = 1
    fcPrint = 2
    fcExit = 3
    fcUndo = 4
    fcOpenLinkAdd = 5
    fcOpenLinkEdit = 6
    fcLinkRefresh = 7
    fcFind = 8
    fcOpen = 9
    fcBrowse = 10
    fcImport = 11
    fcExport = 12
    fcCustom1 = 13
    fcCustom2 = 14
End Enum

Event ClientRefresh()
Event ClientClear()
'local variable(s) to hold property value(s)
Private mvarOmonimySearch As Boolean 'local copy

Public Function AddTSUpdate() As String
    Dim s As String
    
    With mobjDA
        If .DataEngine = deJet Then
            s = " timestmp = " & .MakeSQL(.CreateGUIDKey, dtAlfa)
        Else
            s = vbNullString
        End If
    End With
    AddTSUpdate = s
End Function


Public Property Get DA() As CDataAccess
    Set DA = mobjDA
End Property

Public Property Get ErrH() As CErrorHandler
    Set ErrH = mobjErrH
End Property

Public Property Set FormHandle(objNewValue As Object)
    Set mobjFormHandle = objNewValue
End Property

Public Property Get FormHandle() As Object
    Set FormHandle = mobjFormHandle
End Property

Public Function GetFirmCode(FirmID As String, _
                            Optional CodeColumn As Integer = 1) As String
    Dim i As Integer
    
    For i = 0 To UBound(mvarList, 2)
        If mvarList(0, i) = FirmID Then
            GetFirmCode = mvarList(CodeColumn, i)
            Exit For
        End If
    Next i
End Function

Public Property Let OmonimySearch(ByVal vData As Boolean)
Attribute OmonimySearch.VB_Description = "Segnala che l'ultima ricerca per descrizione effettuata ha restituito degli omonimi"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OmonimySearch = 5
    mvarOmonimySearch = vData
End Property


Public Property Get OmonimySearch() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OmonimySearch
    OmonimySearch = mvarOmonimySearch
End Property











Private Function FindInForeignTable(Field As String, SearchValue As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngRows As Long
    
    strQ = "SELECT "
    With mcolFields(Field)
        If Trim$(SearchValue) = vbNullString Then
            GoTo ErrHandler
        Else
            If (.ForeignTable <> vbNullString) _
                And (.ForeignTableFld <> vbNullString) _
                And (.ForeignTableDeleteFld <> vbNullString) _
                And (.ForeignTableDescField <> vbNullString) Then
                strQ = strQ & .ForeignTableDescField
                strQ = strQ & " FROM "
                strQ = strQ & .ForeignTable
                strQ = strQ & " WHERE "
                strQ = strQ & .ForeignTableFld
                strQ = strQ & " = "
                strQ = strQ & mobjDA.MakeSQL(SearchValue, .FldType)
                strQ = strQ & " AND "
                strQ = strQ & .ForeignTableDeleteFld
                strQ = strQ & " = "
                strQ = strQ & mobjDA.MakeSQL(False, dtBit)
                
                mobjDA.QueryData strQ, varBuffer, lngRows
                
                'Se ho trovato un valore nella tabella collegata...
                If lngRows > 0 Then
                    '... lo restituisco ed esco ...
                    .ForeignTableDescValue = varBuffer(0, 0)
                    GoTo EndProc
                Else
                    '... altrimenti provo a cercare tra le descrizioni
                    strQ = "SELECT "
                    strQ = strQ & .ForeignTableDescField
                    strQ = strQ & ","
                    strQ = strQ & .ForeignTableFld
                    strQ = strQ & " FROM "
                    strQ = strQ & .ForeignTable
                    strQ = strQ & " WHERE "
                    strQ = strQ & .ForeignTableDescField
                    strQ = strQ & " LIKE "
                    strQ = strQ & mobjDA.MakeSQL(SearchValue & "%", .FldType)
                    strQ = strQ & " AND "
                    strQ = strQ & .ForeignTableDeleteFld
                    strQ = strQ & " = "
                    strQ = strQ & mobjDA.MakeSQL(False, dtBit)
                    
                    mobjDA.QueryData strQ, varBuffer, lngRows
                    
                    If lngRows > 0 Then
                        'Se ho trovato un solo valore ...
                        If lngRows = 1 Then
                            '... lo restituisco ed esco...
                            .ForeignTableDescValue = varBuffer(0, 0)
                            .ForeignTableFldValue = varBuffer(1, 0)
                            GoTo EndProc
                        Else
                        '... altrimenti faccio selezionare un valore all'utente
                            mvarOmonimySearch = True
                        End If
                    Else
                        .ForeignTableDescValue = "Codice non trovato"
                        FindInForeignTable = False
                        Exit Function
                    End If
                End If
            Else
                GoTo ErrHandler
            End If
        End If
    End With
    GoTo EndProc
ErrHandler:
    mcolFields(Field).ForeignTableDescValue = vbNullString
    FindInForeignTable = False
    Exit Function
EndProc:
    FindInForeignTable = True
End Function

Public Function BeginSave(SQL As String, SQL_TS As String, ByRef blnConcurErr As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim blnSave As Boolean
    Dim varNewTS As Variant
    Dim Fld As CField
    Dim i As Integer
    
    strQ = SQL
    
    For Each Fld In mcolFields
        With Fld
            If .Updated Then
                strQ = strQ & .FieldName & " = " & mobjDA.MakeSQL(.NewValue, .FldType) & ","
                blnSave = True
            End If
        End With
    Next Fld
    
    If blnSave Then
        With mobjDA
            strQ = strQ & AddTSUpdate
            If Right$(strQ, 1) = "," Then
                strQ = Left$(strQ, Len(strQ) - 1)
            End If
            
            strQ = strQ & " WHERE "
            For i = 1 To mcolKeyFields.Count
                If mcolFields(mcolKeyFields(i)).Value <> vbNullString Then
                    strQ = strQ & mcolFields(mcolKeyFields(i)).FieldName & _
                           mcolFields(mcolKeyFields(i)).Operator & _
                           .MakeSQL(mcolFields(mcolKeyFields(i)).Value, _
                           mcolFields(mcolKeyFields(i)).FldType)
                    strQ = strQ & " AND "
                End If
            Next i
            If Right$(strQ, 5) = " AND " Then
                strQ = Left$(strQ, Len(strQ) - 4)
            End If
            .BeginTransaction
            
            varNewTS = GetTimeStamp(SQL_TS)
            
            If Not .ExecuteQuery(strQ) Then
                .EndTrans etRollback
                GoTo ErrHandler
            End If
            
            If .CompareTimeStamps(varNewTS, TimeStamp) Then
                blnConcurErr = False
            Else
                blnConcurErr = True
            End If
        End With
    End If
    BeginSave = True
    Exit Function
ErrHandler:
    BeginSave = False
End Function

Public Function DecodeFirmCode(FirmID As String, FirmCode As String, SQL As String, Optional ReturnID = False) As String
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strID As String
    Dim varResult As Variant
    Dim lngHowMany As Long
    
    With MultiSoc.DS
        strQ = MultiSoc.SQL_FIND
        strQ = strQ & .Fields("Società").FieldName & " = " & mobjDA.MakeSQL(FirmID, 0)
        strQ = strQ & " AND " & .Fields("CodPerSoc").FieldName & " = " & mobjDA.MakeSQL(FirmCode, 0)
        mobjDA.QueryData strQ, varResult, lngHowMany
        If lngHowMany > 0 Then
            strID = varResult(0, 0)
            If Not ReturnID Then
                Fields(KeyFields(1)).FindValue = strID
                If Find(SQL) Then
                    DecodeFirmCode = Fields("1").Value
                End If
            Else
                DecodeFirmCode = strID
            End If
        End If
    End With
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function

Public Property Let CancelField(NewValue As String)
    mstrCancelField = NewValue
End Property

Public Property Get CancelField() As String
    CancelField = mstrCancelField
End Property


Public Sub ConfigureInterface()
    Dim C As Control
    Dim strQ As String
    Dim strTableName As String
    Dim strFieldName As String
    Dim strCompleteFieldName As String
    Dim varData As Variant
    Dim lngRowNum As Long
    Dim i As Integer
    
    If Not mobjFormHandle Is Nothing Then
        For Each C In mobjFormHandle
            If C.Tag <> vbNullString Then
                Select Case UCase(TypeName(C))
                    Case "COMMANDBUTTON"
                    Case "LABEL"
                    Case "MASKEDBOX"
                        C.MaxLength = mcolFields(C.Tag).MaxLength
                        mcolFields(C.Tag).FieldHwnd = C.hwnd
                    Case "CHECKBOX"
                        mcolFields(C.Tag).FieldHwnd = C.hwnd
                    Case "COMBOBOX"
                        strCompleteFieldName = mcolFields(C.Tag).FieldName
                        strTableName = Left$(strCompleteFieldName, InStr(1, strCompleteFieldName, ".") - 1)
                        strFieldName = Mid$(strCompleteFieldName, InStr(1, strCompleteFieldName, ".") + 1)
                        lngRowNum = 0
                        strQ = "SELECT avvvalue,avvdescr FROM " & _
                                            " ANAGAVV WHERE avvidtbl = " & mobjDA.MakeSQL(strTableName, dtAlfa) & _
                                            "           AND avvidfld = " & mobjDA.MakeSQL(strFieldName, dtAlfa)
                        If mobjDA.QueryData(strQ, varData, lngRowNum) Then
                            If lngRowNum > 0 Then
                                For i = 0 To UBound(varData, 2)
                                    C.AddItem varData(1, i)
                                    mcolFields(C.Tag).AllowedValues.Add varData(0, i)
                                Next i
                            End If
                        End If
                        mcolFields(C.Tag).FieldHwnd = C.hwnd
                End Select
            End If
        Next C
    End If
End Sub

Public Function FindControl(FieldKey As String, ControlType As String) As Object
    Dim C As Control
    
    For Each C In mobjFormHandle
        If UCase(TypeName(C)) = ControlType Then
            If UCase(C.Tag) = UCase(FieldKey) Then
                Set FindControl = C
                Exit For
            End If
        End If
    Next C
End Function


Public Function GetSelection(SQL As String, _
                             ActiveRecOnly As Boolean, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0, _
                             Optional AddKeyConditions As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim strCond As String
    Dim lngRows As Long
    Dim i As Integer
    
    strQ = SQL
    If AddKeyConditions Then
        strCond = MakeKeyCond
        strQ = strQ & " WHERE " & strCond 'MakeKeyCond
    End If
    
    If ActiveRecOnly And (CancelField <> vbNullString) Then
        
        If (Right$(strQ, 5) <> " AND ") And (Len(strCond) > 0) Then
            strQ = strQ & " AND "
        Else
            If Right$(strQ, 7) <> " WHERE " Then
                strQ = strQ & " WHERE "
            End If
        End If
        strQ = strQ & CancelField & " = " & mobjDA.MakeSQL(False, dtBit)
    End If
    
    If Right$(strQ, 5) = " AND " Then
        strQ = Left$(strQ, Len(strQ) - 5)
    Else
        If Right$(strQ, 7) = " WHERE " Then
            strQ = Left$(strQ, Len(strQ) - 7)
        End If
    End If
    
    If Not IsMissing(OrderFields) Then
        strQ = strQ & " ORDER BY "
        For i = 0 To UBound(OrderFields)
            strQ = strQ & mcolFields(OrderFields(i)).FieldName & ", "
        Next i
        strQ = Left$(strQ, Len(strQ) - 2)
    End If
    
    lngRows = MaxRecords
    If Not mobjDA.QueryData(strQ, mvarList, lngRows) Then
        GoTo ErrHandler
    End If
    GoTo EndFunc
ErrHandler:
    mvarHowMany = 0
    GetSelection = False
    Exit Function
EndFunc:
    mvarHowMany = lngRows
    GetSelection = True

End Function

Private Function MakeOrder() As String
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim strQ As String
    
    If mcolKeyFields.Count > 0 Then
        strQ = " ORDER BY "
        For i = 1 To mcolKeyFields.Count
            strQ = strQ & mcolFields(mcolKeyFields(i)).FieldName & ", "
        Next i
        strQ = Left$(strQ, Len(strQ) - 2)
    End If
    MakeOrder = strQ
    Exit Function
ErrHandler:
    MakeOrder = vbNullString
End Function

Public Function RefreshTimeStamp(SQL_TS As String) As Boolean
    TimeStamp = GetTimeStamp(SQL_TS)
End Function

Public Function ScanClientCond() As Boolean
    On Error GoTo ErrHandler
    Dim C As Control
    Dim blnChkValue As Boolean
    Dim blnScanField As Boolean
    Dim strText As String
    
    For Each C In mobjFormHandle
        If C.Tag <> vbNullString Then
            blnScanField = True
            If blnScanField Then
                Select Case UCase(TypeName(C))
                    Case "MASKEDBOX"
                        strText = Trim(C.FormattedText)
                        If (strText <> vbNullString) And (strText <> DATE_MASK) Then
                            mcolFields(C.Tag).FindValue = UCase(strText)
                        Else
                            mcolFields(C.Tag).FindValue = vbNullString
                        End If
                    Case "CHECKBOX"
                        blnChkValue = (C.Value = vbChecked)
                        mcolFields(C.Tag).FindValue = blnChkValue
                    Case "COMBOBOX"
                        With mcolFields(C.Tag)
                            If (C.ListIndex >= 0) Then
                                .FindValue = .AllowedValues(C.ListIndex + 1)
                            Else
                                .FindValue = vbNullString
                            End If
                        End With
                End Select
            End If
        End If
    Next C
    GoTo EndProc
    
ErrHandler:
    ScanClientCond = False
    Exit Function
EndProc:
    ScanClientCond = True

End Function

Public Sub SetClientFieldFocus(FieldKey As String)
    Dim i As Image
    Dim L As Label
    
    Set i = FindControl(FieldKey, "IMAGE")
    If Not i Is Nothing Then
        i.Visible = False 'True
'        i.Picture = mobjFormHandle.ImageList1.ListImages("Editing").Picture
    End If
    Set L = FindControl(FieldKey, "LABEL")
    If Not L Is Nothing Then
        L.MousePointer = vbArrowQuestion
    End If
End Sub

Public Function ScanClient(ScanMode As ScanMode) As Boolean
    On Error GoTo ErrHandler
    Dim C As Control
    Dim blnScanField As Boolean
    Dim blnChkValue As Boolean
    Dim strFormattedText As String
    
    For Each C In mobjFormHandle
        If C.Tag <> vbNullString Then
            blnScanField = True
            If (ScanMode = smUpdate) And (C.Tag = KeyFields(1)) Then
                blnScanField = False
                mcolFields(C.Tag).Updated = False
            End If
            If blnScanField Then
                Select Case UCase(TypeName(C))
                    Case "MASKEDBOX"
                        With mcolFields(C.Tag)
                            If (ScanMode = smInsert) And (.ReadOnly) Then
                                .Updated = False
                            Else
                                strFormattedText = UCase(Trim(C.FormattedText))
                                If (UCase(Trim(.Value)) <> strFormattedText) Or (strFormattedText = vbNullString) Then
                                    If (strFormattedText = DATE_MASK) Or (strFormattedText = TIME_MASK) Then
                                        .NewValue = vbNullString
                                    Else
                                        .NewValue = strFormattedText
                                    End If
                                    .Updated = True
                                Else
                                    .Updated = False
                                End If
                            End If
                        End With
                    Case "CHECKBOX"
                        With mcolFields(C.Tag)
                            If (ScanMode = smInsert) And (.ReadOnly) Then
                                .Updated = False
                            Else
                                blnChkValue = (C.Value = vbChecked)
                                If .Value <> blnChkValue Then
                                    .NewValue = blnChkValue
                                    .Updated = True
                                Else
                                    .Updated = False
                                End If
                            End If
                        End With
                    Case "COMBOBOX"
                        With mcolFields(C.Tag)
                            If (ScanMode = smInsert) And (.ReadOnly) Then
                                .Updated = False
                            Else
                                If C.ListIndex > -1 Then
                                    If .Value <> .AllowedValues(C.ListIndex + 1) Then
                                        .NewValue = .AllowedValues(C.ListIndex + 1)
                                        .Updated = True
                                    Else
                                        .Updated = False
                                    End If
                                Else
                                    If Trim$(.Value) <> vbNullString Then
                                        .NewValue = vbNullString
                                        .Updated = True
                                    Else
                                        .Updated = False
                                    End If
                                End If
                            End If
                        End With
                End Select
            End If
        End If
    Next C
    GoTo EndProc
    
ErrHandler:
    ScanClient = False
    Exit Function
EndProc:
    ScanClient = True
End Function

Public Function RefreshClient() As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim C As Control
    Dim blnFound As Boolean
    Dim strValue As String
    Dim strMask As String
    
    mobjFormHandle.MousePointer = vbHourglass
    For Each C In mobjFormHandle
        If C.Tag <> vbNullString Then
            Select Case UCase(TypeName(C))
                Case "MASKEDBOX"
                    Select Case Fields(C.Tag).FldType
                        Case dtAlfa
                            C.Text = Trim$(Fields(C.Tag).Value)
                        Case dtNum
                            If Left$(Fields(C.Tag).Value, 1) = mobjDA.SqlDecSeparator Then
                                C.Text = "0" & Fields(C.Tag).Value
                            Else
                                C.Text = Fields(C.Tag).Value
                            End If
                        Case dtDate
                            strValue = Trim$(Fields(C.Tag).Value)
                            If Format$(strValue, "dd/mm/yyyy") = "01/01/1900" Then
                                strMask = C.Mask
                                C.Mask = vbNullString
                                C.Text = vbNullString
                                C.Mask = strMask
                            Else
                                If strValue <> vbNullString Then
                                    C.Text = Format$(strValue, "dd/mm/yy")
                                End If
                            End If
                        Case dtTime
                            strValue = Trim$(Fields(C.Tag).Value)
                            C.Text = Format$(Hour(strValue), "00") & ":" & Format$(Minute(strValue), "00")
                        Case Else
                            C.Text = Trim$(Fields(C.Tag).Value)
                    End Select
                Case "IMAGE"
                    C.Visible = False
                Case "CHECKBOX"
                    C.Value = IIf(Fields(C.Tag).Value, vbChecked, vbUnchecked)
                Case "COMBOBOX"
                    blnFound = False
                    i = 1
                    Do While (Not blnFound) And i <= Fields(C.Tag).AllowedValues.Count
                        blnFound = (Fields(C.Tag).AllowedValues(i) = Fields(C.Tag).Value)
                        If blnFound Then
                            C.ListIndex = i - 1
                        Else
                            i = i + 1
                        End If
                    Loop
                    If Not blnFound Then
                        C.ListIndex = -1
                    End If
                Case "LABEL"
                    If C.Tag <> "LINK" Then
                        If FindInForeignTable(C.Tag, Fields(C.Tag).Value) Then
                            C = " " & Fields(C.Tag).ForeignTableDescValue
                        End If
                    Else
                        C = Fields(KeyFields(1)).Value
                    End If
            End Select
        End If
    Next C
    GoTo EndProc
    Exit Function
ErrHandler:
    RefreshClient = False
    Exit Function
EndProc:
    mobjFormHandle.MousePointer = vbDefault
    RaiseEvent ClientRefresh
    RefreshClient = True
End Function



Public Sub ClearClient()
    Dim C As Control
    Dim strMask As String
    
    DoEvents
    
    For Each C In mobjFormHandle
        If C.Tag <> vbNullString Then
            Select Case UCase(TypeName(C))
                Case "MASKEDBOX"
                    If Len(C.Mask) = 0 Then
                        C.Text = vbNullString
                    Else
                        strMask = C.Mask
                        C.Mask = vbNullString
                        C.Text = vbNullString
                        C.Mask = strMask
                    End If
                Case "IMAGE"
                    C.Visible = False
                Case "CHECKBOX"
                    C.Value = vbUnchecked
                Case "COMBOBOX"
                    C.ListIndex = -1
                Case "LABEL"
                    C.Caption = vbNullString
            End Select
        End If
    Next C
    RaiseEvent ClientClear
End Sub

Public Sub SetClientState(NewState As FormState)
    Dim C As Control
    Dim i As Integer
    Dim blnDummy As Boolean
    
    Select Case NewState
        Case fsNone
            ClearClient
            Fields.ResetFields
            For Each C In mobjFormHandle.Controls
                Select Case UCase(TypeName(C))
                    Case "TDBGRID"
                        If UCase$(C.Tag) <> "MULTISOC" Then
                            C.Enabled = True
                        Else
                            C.Enabled = False
                        End If
                    Case "COMMANDBUTTON"
                        Select Case UCase(C.Tag)
                            Case "ADD"
                                C.Enabled = True
                            Case "SAVE"
                                C.Enabled = False
                            Case "PRINT"
                                C.Enabled = True
                            Case "FIND"
                                C.Enabled = True
                            Case "UNDO"
                                C.Enabled = False
                            Case "EDITLINK"
                                C.Enabled = False
                        End Select
                    Case "MASKEDBOX"
                        If Not Fields(C.Tag).ReadOnly Then
                            If C.Tag = KeyFields(1) Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        Else
                            C.Enabled = False
                        End If
                    Case "CHECKBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = False
                            Else
                                C.Enabled = False
                            End If
                        End If
                       
                    Case "COMBOBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = False
                            Else
                                C.Enabled = False
                            End If
                        End If
                End Select
            Next C
            With mobjFormHandle
                .State = fsNone
                .StatusBar1.Panels("Msg") = "In attesa di comandi"
            End With
        Case fsEdit
            For Each C In mobjFormHandle.Controls
                Select Case UCase(TypeName(C))
                    Case "TDBGRID"
                        C.Enabled = True
                    Case "COMMANDBUTTON"
                        Select Case UCase(C.Tag)
                            Case "ADD"
                                C.Enabled = True
                            Case "SAVE"
                                C.Enabled = True
                            Case "PRINT"
                                C.Enabled = True
                            Case "FIND"
                                C.Enabled = True
                            Case "UNDO"
                                C.Enabled = True
                             Case "EDITLINK"
                                C.Enabled = True
                       End Select
                    Case "MASKEDBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                blnDummy = False
                                For i = 1 To KeyFields.Count
                                    If C.Tag = KeyFields(i) Then
                                        blnDummy = True
                                        C.Enabled = False
                                    End If
                                Next i
                                If Not blnDummy Then
                                    C.Enabled = True
                                End If
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "CHECKBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "COMBOBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                End Select
            Next C

            With mobjFormHandle
                If .State = fsEdit Then
                    RefreshClient
                End If
                .State = fsEdit
                .StatusBar1.Panels("Msg") = "Modalità Modifica"
            End With
        Case fsInsert
            For Each C In mobjFormHandle.Controls
                Select Case UCase(TypeName(C))
                    Case "TDBGRID"
                        Select Case FormType
                            Case ftEdtRcd1S
                                C.Enabled = False
                            Case Else
                                If UCase(C.Tag) <> "MULTISOC" And UCase(C.Tag) <> "INSDETAIL" And UCase(C.Tag) <> "ENABLE" Then
                                    C.Enabled = False
                                End If
                        End Select
                    Case "COMMANDBUTTON"
                        Select Case UCase(C.Tag)
                            Case "ADD"
                                C.Enabled = False
                            Case "SAVE"
                                C.Enabled = True
                            Case "PRINT"
                                C.Enabled = False
                            Case "FIND"
                                C.Enabled = False
                            Case "UNDO"
                                C.Enabled = True
                        End Select
                    Case "MASKEDBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "CHECKBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "COMBOBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "IMAGE"
                        If mcolFields(C.Tag).Required Then
                            C.Picture = mobjFormHandle.ImageList1.ListImages("RedCross").Picture
                            C.Visible = True
                        End If
                End Select
            Next C
            With mobjFormHandle
                .State = fsInsert
                .StatusBar1.Panels("Msg") = "Modalità Inserimento"
            End With
        Case fsCopy
            For Each C In mobjFormHandle.Controls
                Select Case UCase(TypeName(C))
                    Case "COMMANDBUTTON"
                        Select Case UCase(C.Tag)
                            Case "ADD"
                                C.Enabled = False
                            Case "SAVE"
                                C.Enabled = True
                            Case "PRINT"
                                C.Enabled = False
                            Case "FIND"
                                C.Enabled = False
                            Case "UNDO"
                                C.Enabled = True
                        End Select
                    Case "MASKEDBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "CHECKBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "COMBOBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "IMAGE"
                        If mcolFields(C.Tag).Required Then
                            C.Picture = mobjFormHandle.ImageList1.ListImages("RedCross").Picture
                            C.Visible = True
                        End If
                End Select
            Next C
            With mobjFormHandle
                .State = fsCopy
                .StatusBar1.Panels("Msg") = "Modalità Copia"
            End With
        Case fsParamInput
            For Each C In mobjFormHandle.Controls
                Select Case UCase(TypeName(C))
                    Case "TDBGRID"
                        C.Enabled = False
                    Case "COMMANDBUTTON"
                        Select Case UCase(C.Tag)
                            Case "PRINT"
                                C.Enabled = False
                            Case "FIND"
                                C.Enabled = True
                            Case "UNDO"
                                C.Enabled = True
                        End Select
                    Case "MASKEDBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "CHECKBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                    Case "COMBOBOX"
                        If C.Tag <> vbNullString Then
                            If Not Fields(C.Tag).ReadOnly Then
                                C.Enabled = True
                            Else
                                C.Enabled = False
                            End If
                        End If
                End Select
            Next C

            With mobjFormHandle
                .State = fsParamInput
                .StatusBar1.Panels("Msg") = "Modalità Immissione Parametri"
            End With
        Case fsSelection
            For Each C In mobjFormHandle.Controls
                Select Case UCase(TypeName(C))
                    Case "TDBGRID"
                        C.Enabled = True
                    Case "COMMANDBUTTON"
                        Select Case UCase(C.Tag)
                            Case "PRINT"
                                C.Enabled = True
                            Case "FIND"
                                C.Enabled = True
                            Case "UNDO"
                                C.Enabled = True
                        End Select
                End Select
            Next C

            With mobjFormHandle
                .State = fsSelection
                .StatusBar1.Panels("Msg") = "Modalità Selezione"
            End With
    End Select
End Sub

Public Property Set ErrH(NewValue As CErrorHandler)
    Set mobjErrH = NewValue
    Set mobjDA.ErrH = mobjErrH
End Property

Public Property Get Fields() As ColFields
    Set Fields = mcolFields
End Property

Public Function Find(SQL As String) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim lngNRows As Long
    Dim i As Integer
    Dim intCount As Integer
    
    strQ = SQL & MakeKeyCond()
    
    lngNRows = 1
    If mobjDA.QueryData(strQ, mvarRecBuf, lngNRows) Then
        If lngNRows > 0 Then
            mvarHowMany = lngNRows + 1
            intCount = UBound(mvarRecBuf, 1) - 1
            For i = 0 To intCount
                If IsNull(mvarRecBuf(i, 0)) Then
                    mcolFields(i + 1).Value = vbNullString
                Else
                    mcolFields(i + 1).Value = mvarRecBuf(i, 0)
                End If
            Next i
            TimeStamp = mvarRecBuf(i, 0)
            
            If Not mobjMultiSoc Is Nothing Then
                With mobjMultiSoc.DS
                    If .KeyFields.Count > 1 Then
                        .Fields(.KeyFields(1)).FindValue = mcolFields(1).Value
                        For i = 2 To .KeyFields.Count
                            .Fields(.KeyFields(i)).FindValue = Empty
                        Next i
                    End If
                End With
                mobjMultiSoc.GetList
            End If
        Else
            mvarHowMany = 0
        End If
    Else
        GoTo ErrHandler
    End If
    GoTo EndFunc
ErrHandler:
    Find = False
    Exit Function
EndFunc:
    Find = True
End Function

Public Function GetList(SQL As String, _
                        Optional ActiveRecordsOnly As Boolean = False) As Boolean
    
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim strDummy As String
    Dim lngRows As Long
    Dim i As Integer
    Dim j As Integer
    Dim intKeyFieldsCount As Integer
    
    If mobjDA Is Nothing Then
        GoTo ErrHandler
    End If
    strQ = SQL
    intKeyFieldsCount = KeyFields.Count
    If intKeyFieldsCount > 1 Then
        strQ = strQ & " WHERE "
        For i = 1 To intKeyFieldsCount - 1
            strDummy = vbNullString
            With mcolFields(mcolKeyFields(i))
                If .FindValuesCount > 0 Then
                    For j = 1 To .FindValuesCount
                        If Not IsEmpty(.FindValue(j)) Then
                            strDummy = strDummy & .FieldName & _
                                   .Operator & _
                                   mobjDA.MakeSQL(.FindValue(j), .FldType)
                            If j <> .FindValuesCount Then
                                strDummy = strDummy & " OR "
                            End If
                        End If
                    Next j
                    If Len(strDummy) > 0 Then
                        strQ = strQ & "(" & strDummy & ")" & " AND "
                    End If
                End If
            End With
        Next i
    End If
        
    If Right$(strQ, 5) = " AND " Then
        strQ = Left$(strQ, Len(strQ) - 5)
    End If
    If ActiveRecordsOnly And (CancelField <> vbNullString) Then
        If KeyFields.Count > 1 Then
            strQ = strQ & " AND "
        Else
            strQ = strQ & " WHERE "
        End If
        strQ = strQ & CancelField & " = " & mobjDA.MakeSQL(False, dtBit)
    End If
    
    strQ = strQ & MakeOrder
    
    If Not mobjDA.QueryData(strQ, mvarList, lngRows) Then
        GoTo ErrHandler
    End If
    GoTo EndFunc
ErrHandler:
    mvarHowMany = 0
    GetList = False
    Exit Function
EndFunc:
    mvarHowMany = lngRows
    GetList = True
End Function

Private Function GetTimeStamp(SQL As String) As Variant
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim varResult As Variant
    Dim i As Integer
    
    strQ = SQL & MakeKeyCond
    
    If mobjDA.QueryData(strQ, mvarRecBuf) Then
        varResult = mvarRecBuf(0, 0)
        GoTo EndFunc
    Else
        GoTo ErrHandler
    End If
    Exit Function
ErrHandler:
    GetTimeStamp = Empty
    Exit Function
EndFunc:
    GetTimeStamp = varResult
End Function

Public Property Get HowMany() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HowMany
    HowMany = mvarHowMany
End Property

Public Function Insert(SQL As String, _
                       Optional TableName As String, _
                       Optional FieldName As String, _
                       Optional KeyFieldName As String, _
                       Optional KeyFieldValue As String, _
                       Optional KeyFieldType As SQLDataType) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strFields As String
    Dim strValues As String
    Dim blnSave As Boolean
    Dim varCounter As Variant
    Dim i As Integer
    
    ' --- Se necessario, leggo il valore del contatore
    If TableName <> vbNullString Then
        varCounter = mobjDA.GetCounter(TableName, FieldName, KeyFieldName, KeyFieldValue, KeyFieldType)
        mcolFields(KeyFields(KeyFields.Count)).NewValue = varCounter
        mcolFields(KeyFields(KeyFields.Count)).Updated = True
    End If
    
    
    strQ = SQL & "("
    For i = 1 To mcolFields.Count
        With mcolFields(i)
            If .Updated Then
                strFields = strFields & Mid$(.FieldName, InStr(1, .FieldName, ".") + 1) & ","
                strValues = strValues & mobjDA.MakeSQL(.NewValue, .FldType) & ","
                blnSave = True
            End If
        End With
    Next i
    
    If blnSave Then
        strQ = strQ & Left$(strFields, Len(strFields) - 1)
        strQ = strQ & ") VALUES ("
        strQ = strQ & Left$(strValues, Len(strValues) - 1)
        strQ = strQ & ")"
        If Not mobjDA.ExecuteQuery(strQ) Then
            If TableName <> vbNullString Then
                mobjDA.UpdateCounter FieldName, varCounter, True
            End If
            GoTo ErrHandler
        End If
        If TableName <> vbNullString Then
            If Not mobjDA.UpdateCounter(FieldName, varCounter) Then
                GoTo ErrHandler
            End If
        End If
        For i = 1 To mcolFields.Count
            With mcolFields(i)
                If .Updated Then
                    .Updated = False
                    .Value = .NewValue
                End If
            End With
        Next i
    Else
    End If
    
    GoTo EndFunc
ErrHandler:
    Insert = False
    Exit Function
EndFunc:
    Insert = True

End Function
Public Function MakeKeyCond(Optional CrystalFlag As Boolean = False, _
                            Optional ByRef Description As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim strQ As String
    Dim strFieldName As String
    Dim strOperator As String
    Dim strFindValue As String
    Dim intFieldsCount As Integer
    Dim intFindValuesCount As Integer
    
    intFieldsCount = mcolKeyFields.Count
    For i = 1 To intFieldsCount
        With mcolFields(mcolKeyFields(i))
            If .FindValue <> vbNullString Then
                strQ = strQ & "("
                intFindValuesCount = .FindValuesCount
                For j = 1 To intFindValuesCount
                    strFieldName = .FieldName
                    strOperator = .Operator
                    strFindValue = mobjDA.MakeSQL(.FindValue(j), .FldType)
                    
                    ' --- Stiamo componendo una formula per Crystal Report?
                    If CrystalFlag Then
                        strFieldName = "{" & strFieldName & "}"
                        ' --- Converto le date
                        If .FldType = 2 Then
                            strFindValue = Mid$(strFindValue, 2, Len(strFindValue) - 2)
                            strFindValue = "Date(" & Right$(strFindValue, 4) & "," & _
                                            Left$(strFindValue, 2) & "," & _
                                            Mid$(strFindValue, 4, 2) & ")"
                        End If
                        If Not IsMissing(Description) Then
                            Description = Description & .Name & " " & strOperator & " " & .FindValue & vbCrLf
                        End If
                        ' ---
                    End If
                    ' ---
                    
                    strQ = strQ & strFieldName & strOperator & strFindValue
                    If j <> intFindValuesCount Then
                        strQ = strQ & " OR "
                    End If
                Next j
                strQ = strQ & ")"
                strQ = strQ & " AND "
            End If
        End With
    Next i
    If Right$(strQ, 5) = " AND " Then
        strQ = Left$(strQ, Len(strQ) - 4)
    End If
    MakeKeyCond = strQ
End Function

Public Property Set MultiSoc(NewValue As Object)
    Set mobjMultiSoc = NewValue
End Property

Public Property Get MultiSoc() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IVASoc
    Set MultiSoc = mobjMultiSoc
End Property

Public Property Get KeyFields() As Collection
    Set KeyFields = mcolKeyFields
End Property


Public Property Get List() As Variant
    List = mvarList
End Property

Public Property Get ListItem(Col As Integer, Row As Long) As Variant
    ListItem = mvarList(Col, Row)
End Property

Public Sub EndSave()
    mcolFields.UpdateFields
End Sub


Public Property Set DA(NewValue As CDataAccess)
    Set mobjDA = NewValue
End Property


Public Function ValidateClientField(Field As String, _
                                    ValidateValue As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim strField As String
    Dim i As Image
    Dim C As Control
    
    strField = Field
    If strField <> vbNullString Then
        With Fields(strField)
            Set i = FindControl(strField, "IMAGE")
            If Not i Is Nothing Then
                i.Visible = True
                If ValidateValue = DATE_MASK Or ValidateValue = TIME_MASK Then
                    ValidateValue = vbNullString
                End If
                If (ValidateValue <> .Value) Or (ValidateValue = vbNullString) Then
                    .NewValue = ValidateValue
                    .Updated = True
                                 
                    If .Validate = 0 Then
                        'Se il campo è usato in una relazione...
                        If (.ForeignTable <> vbNullString) Then
                            '... controllo che il valore esista nella tabella collegata
                            If FindInForeignTable(strField, ValidateValue) Then
                                GoTo EndProc
                            Else
                                If .Required Then
                                    GoTo ErrHandler
                                Else
                                    GoTo EndProc
                                End If
                            End If
                        Else
                            GoTo EndProc
                        End If
                    Else
                        'Se il campo è usato in una relazione...
                        If .ForeignTable <> vbNullString Then
                            '... azzero la descrizione del campo
                            .ForeignTableDescValue = vbNullString
                        End If
                        GoTo ErrHandler
                    End If
                Else
                    'Se il campo è usato in una relazione...
                    If .ForeignTable <> vbNullString Then
                        '... controllo che il valore esista nella tabella collegata
                        If FindInForeignTable(strField, ValidateValue) Then
                            GoTo EndProc
                        Else
                            If .Required Then
                                GoTo ErrHandler
                            Else
                                GoTo EndProc
                            End If
                        End If
                    Else
                        GoTo EndProc
                    End If
                End If
            End If
        End With
    End If
    GoTo EndProc
    
ErrHandler:
    If Not i Is Nothing Then
        i.Picture = mobjFormHandle.ImageList1.ListImages("RedCross").Picture
        i.Visible = True
    End If
    ValidateClientField = False
    Exit Function
EndProc:
    If Not i Is Nothing Then
        i.Picture = mobjFormHandle.ImageList1.ListImages("GreenCheck").Picture
        i.Visible = True
    End If
    ValidateClientField = True
End Function

Private Sub Class_Initialize()
    Set mcolFields = New ColFields
    Set mcolKeyFields = New Collection
End Sub


Public Sub CheckClientCombo(C As Object)
    Dim varValue As Variant
    
    If C.ListIndex = -1 Then
        varValue = vbNullString
    Else
        varValue = Fields(C.Tag).AllowedValues(C.ListIndex + 1)
    End If
    ValidateClientField C.Tag, varValue
End Sub












