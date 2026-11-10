VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public WithEvents cmbComboBox As ComboBox
Attribute cmbComboBox.VB_VarHelpID = -1
Public WithEvents mskMaskEdBox As MaskEdBox
Attribute mskMaskEdBox.VB_VarHelpID = -1
Public WithEvents CommandBtn As CommandButton
Attribute CommandBtn.VB_VarHelpID = -1
Public WithEvents DBGrid As TDBGrid
Attribute DBGrid.VB_VarHelpID = -1
Public WithEvents Frm As Form
Attribute Frm.VB_VarHelpID = -1
Public WithEvents Lbl As Label
Attribute Lbl.VB_VarHelpID = -1

Private GA As ArrayGridCls
Public ColTags As Variant
Public ColTags1 As Variant
Public ColTags2 As Variant
Private Enum GridEvents
    geRefresh = 0
    
End Enum

Private Sub GetRecordDetail1(strID As String, strID1 As String, GridArray As ArrayGridCls, DataClass As Object)
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim intFrom As Integer
    Dim intTo As Integer
    Dim intKeyFieldsCount As Integer
    Dim strKeyField As String
    Dim varBookmark As Variant
    
    With DataClass.DS
        .Fields(.KeyFields(1)).FindValue = strID
        .Fields(.KeyFields(2)).FindValue = strID1
        intFrom = LBound(ColTags1)
        intTo = UBound(ColTags1)
        intKeyFieldsCount = .KeyFields.Count
        varBookmark = DBGrid.bookmark
        For i = 3 To intKeyFieldsCount
            strKeyField = .KeyFields(i)
            For j = intFrom To intTo
                If ColTags1(j) = strKeyField Then
                    .Fields(strKeyField).FindValue = GridArray.value(varBookmark, j - 1)
                End If
            Next j
        Next i
    End With
    DataClass.Find
ErrHandler:
    Exit Sub

End Sub

Public Sub SetColTags1(ColNum As Integer)
    ReDim ColTags1(1 To ColNum) As String
End Sub

Public Sub SetMultiSocGrid(Soc As CSocieta, Optional SetTags As Boolean = True)
    Dim i As Long
    Dim vItem As ValueItem
    
    If SetTags Then
        SetColTags MAX_COLS
        ColTags(1) = "Società"
        ColTags(2) = "CodPerSoc"
        ColTags(3) = "Annullato"
    End If
    
    With DBGrid.Columns(0).ValueItems
        .Presentation = dbgComboBox
        .Translate = True
        .Validate = True
        
            If Soc.DS.howmany > 0 Then
            ' El combo no puede tener mas de 100
            If Soc.DS.howmany > 100 Then
            .MaxComboItems = 100
            Else
            .MaxComboItems = Soc.DS.howmany
            End If
            For i = 0 To UBound(Soc.DS.List, 2)
                Set vItem = New ValueItem
                With Soc.DS
                    vItem.value = .ListItem(0, i)
                    vItem.DisplayValue = .ListItem(1, i)
                End With
                .Add vItem
            Next i
            End If
    End With
    DBGrid.Columns(2).Alignment = dbgCenter
End Sub

Public Sub SetColTags2(ColNum As Integer)
    ReDim ColTags2(1 To ColNum) As String
End Sub
Public Sub SetColTags(ColNum As Integer)
    ReDim ColTags(1 To ColNum) As String
End Sub





Private Sub cmbComboBox_Click()
    Frm.DataClass.DS.CheckClientCombo cmbComboBox
End Sub


Private Sub GetRecordDetail(strID As String, GridArray As ArrayGridCls, DataClass As Object) 'strIDSoc As String)
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim intCount As Integer
    Dim intFrom As Integer
    Dim intTo As Integer
    Dim strKeyField As String
    
    If Not IsEmpty(ColTags) Then
        With DataClass.DS
            .Fields(.KeyFields(1)).FindValue = strID
            intCount = .KeyFields.Count
            intFrom = LBound(ColTags)
            intTo = UBound(ColTags)
            For i = 2 To intCount
                strKeyField = .KeyFields(i)
                For j = intFrom To intTo
                    If ColTags(j) = strKeyField Then
                        .Fields(strKeyField).FindValue = GridArray.value(DBGrid.bookmark, j - 1)
                    End If
                Next j
            Next i
        End With
        DataClass.Find
    End If
ErrHandler:
    Exit Sub
End Sub
Private Sub GetRecordDetail2(strID As String, GridArray As ArrayGridCls, DataClass As Object) 'strIDSoc As String)
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim j As Integer
    Dim intCount As Integer
    Dim intFrom As Integer
    Dim intTo As Integer
    Dim strKeyField As String
    
    If Not IsEmpty(ColTags2) Then
        With DataClass.DS
            .Fields(.KeyFields(1)).FindValue = strID
            intCount = .KeyFields.Count
            intFrom = LBound(ColTags2)
            intTo = UBound(ColTags2)
            For i = 2 To intCount
                strKeyField = .KeyFields(i)
                For j = intFrom To intTo
                    If ColTags2(j) = strKeyField Then
                        .Fields(strKeyField).FindValue = GridArray.value(DBGrid.bookmark, j - 1)
                    End If
                Next j
            Next i
        End With
        DataClass.Find
    End If
ErrHandler:
    Exit Sub
End Sub
Private Sub cmbComboBox_GotFocus()
    If Not cmbComboBox.Locked Then
        Frm.DataClass.DS.SetClientFieldFocus cmbComboBox.Tag
    End If
End Sub

Private Sub cmbComboBox_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            cmbComboBox.ListIndex = -1
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub cmbComboBox_LostFocus()
    If Not cmbComboBox.Locked Then
        Frm.DataClass.DS.CheckClientCombo cmbComboBox
    End If
End Sub


Private Sub CommandBtn_Click()
    Select Case UCase(CommandBtn.Tag)
        Case "ADD"
            Frm.DoCommand fcAddNew
        Case "SAVE"
            Frm.DoCommand fcSave
        Case "PRINT"
            Frm.DoCommand fcPrint
        Case "UNDO"
            Frm.DoCommand fcUndo
        Case "EXIT"
            Frm.DoCommand fcExit
        Case "FIND"
            Frm.DoCommand fcFind
        Case "OPEN"
            Frm.DoCommand fcOpen
        Case "ADDLINK"
            Frm.DoCommand fcOpenLinkAdd
        Case "EDITLINK"
            Frm.DoCommand fcOpenLinkEdit
        Case "REFRESHLIST"
            Frm.DoCommand fcLinkRefresh
        Case "BROWSE"
            Frm.DoCommand fcBrowse
        Case "IMPORT"
            Frm.DoCommand fcImport
        Case "EXPORT"
            Frm.DoCommand fcExport
        Case "CUSTOM1"
            Frm.DoCommand fcCustom1
        Case "CUSTOM2"
            Frm.DoCommand fccustom2
        Case "EXPPALM"
            Frm.DoCommand fcexppalm
        Case "DELLSERV"
             Frm.DoCommand fcdellserv
        Case "DELLTR"
             Frm.DoCommand fcdelltr
    End Select

End Sub









Private Sub DBGrid_DblClick()
    On Error GoTo ErrHandler
    
    If DBGrid.AddNewMode <> dbgNoAddNew Then
        GoTo ErrHandler
    End If
    Select Case DBGrid.Tag
        Case "SEARCH"
            Frm.DoCommand fcOpen
        Case "SELECT"
            Frm.DoCommand fcOpen
        Case "DETAIL"
            Frm.DoCommand fcOpenLinkEdit
         Case "DETAIL2"
            Frm.DoCommand fcOpenLinkEdit
        Case "DETAIL1"
            Frm.DoCommand fcFind
    End Select
ErrHandler:
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub DBGrid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid50.StyleDisp)
Private Sub DBGrid_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid60.StyleDisp)
    If FullyDisplayed Then
        CellTip = vbNullString
    End If
End Sub

Private Sub DBGrid_GotFocus()
    Select Case DBGrid.Tag
        Case "INSDETAIL"
            DBGrid.Col = 0
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub DBGrid_KeyPress(KeyAscii As Integer)
    Select Case DBGrid.Tag
        Case "SELECT"
            If KeyAscii = vbKeyReturn Then
                Frm.DoCommand fcOpenLinkEdit
            End If
    End Select

End Sub


Private Sub DBGrid_PostEvent(ByVal MsgId As Integer)
    On Error GoTo ErrHandler
    
    Select Case MsgId
        Case geRefresh
            GA.DeleteRow 0
            With DBGrid
                .bookmark = 0
                .ReBind
                .Row = 1
            End With
    End Select
ErrHandler:

End Sub

Private Sub DBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strID As String
    Dim strID1 As String
    Dim strIDSoc As String
    Dim strIDDetail As String
    Dim DC As Object
    Dim C As Control
    
    Select Case DBGrid.AddNewMode
        Case dbgNoAddNew
            ' --- Non modificare: LastRow può essere Null
            If LastRow = DBGrid.bookmark Then
            Else
                With Frm
                    Select Case DBGrid.Tag
                        Case "MULTISOC"
                            Set GA = .GridArray
                            Set DC = .DataClass.DS.MultiSoc
                        Case "DETAIL"
                            Set GA = .GridArrayDetail
                            Set DC = .DetailClass
                        Case "DETAIL2"
                            Set GA = .GridArrayDetail2
                            Set DC = .DetailClass2
                            
                        Case "DETAIL1"
                            Set GA = .GridArrayDetail1
                            Set DC = .DetailClass1
                        Case Else
                            Exit Sub
                    End Select
                    With .DataClass.DS
                        If Not IsEmpty(.Fields(.KeyFields(1)).value) Then
                            Frm.MousePointer = vbHourglass
                            Select Case DBGrid.Tag
                                Case "MULTISOC"
                                    strID = .Fields(.KeyFields(1)).value
                                    GetRecordDetail strID, GA, DC
                                Case "DETAIL"
                                    strID = .Fields(.KeyFields(1)).value
                                    GetRecordDetail strID, GA, DC
                                    With Frm
                                        For Each C In .Controls
                                            If UCase(TypeName(C)) = "TDBGRID" Then
                                                If UCase(C.Tag) = "DETAIL1" Then
                                                    If DC.DS.howmany > 0 Then
                                                        .DisplayDetail1
                                                    Else
                                                        .ClearDetail1
                                                    End If
                                                    Exit For
                                                End If
                                            End If
                                        Next C
                                    End With
                                    
                                Case "DETAIL2"
                                    strID = .Fields(.KeyFields(1)).value
                                    GetRecordDetail2 strID, GA, DC
                                    
                                Case "DETAIL1"
                                    strID = .Fields(.KeyFields(1)).value
                                    With Frm.DetailClass.DS
                                        strID1 = .Fields(.KeyFields(2)).value
                                    End With
                                    If (Len(strID1) > 0) And (Len(strID) > 0) Then
                                        GetRecordDetail1 strID, strID1, GA, DC
                                    End If
                            End Select
                            Frm.MousePointer = vbDefault
                        End If
                    End With
                End With
            End If
        Case dbgAddNewCurrent
            ' --- Se ci siamo posizionati sulla riga di inserimento
            ' --- svuoto la grid di dettaglio collegata
            Select Case DBGrid.Tag
                Case "DETAIL"
                    For Each C In Frm.Controls
                        If UCase(TypeName(C)) = "TDBGRID" Then
                            If UCase(C.Tag) = "DETAIL1" Then
                                Frm.ClearDetail1
                                Exit For
                            End If
                        End If
                    Next C
            End Select
            ' ---
    End Select
End Sub


'*** APEX Migration Utility Code Change ***
'Private Sub DBGrid_UnboundAddData(ByVal RowBuf As TrueDBGrid50.RowBuffer, NewRowBookmark As Variant)
Private Sub DBGrid_UnboundAddData(ByVal RowBuf As TrueDBGrid60.RowBuffer, NewRowBookmark As Variant)
    Select Case DBGrid.Tag
        Case "MULTISOC"
            Set GA = Frm.GridArray
            If IsEmpty(GA.value(0, 0)) Then
                DBGrid.PostMsg geRefresh
            End If
            GA.AddData RowBuf, NewRowBookmark, DBGrid.Columns
            RowToRecord NewRowBookmark, Frm.DataClass.DS.MultiSoc.DS, Frm.GridArray, Frm, ColTags
            Frm.DataClass.DS.MultiSoc.Insert
        Case "DETAIL"
            Set GA = Frm.GridArrayDetail
            If IsEmpty(GA.value(0, 0)) Then
                DBGrid.PostMsg geRefresh
            End If
            GA.AddData RowBuf, NewRowBookmark, DBGrid.Columns
            RowToRecord NewRowBookmark, Frm.DetailClass.DS, Frm.GridArrayDetail, Frm, ColTags
            If Not Frm.InsertDetail Then
                GA.Delete NewRowBookmark
            End If
            
        Case "DETAIL2"
            Set GA = Frm.GridArrayDetail2
            If IsEmpty(GA.value(0, 0)) Then
                DBGrid.PostMsg geRefresh
            End If
            GA.AddData RowBuf, NewRowBookmark, DBGrid.Columns
            RowToRecord NewRowBookmark, Frm.DetailClass2.DS, Frm.GridArrayDetail2, Frm, ColTags2
            If Not Frm.InsertDetail2 Then
                GA.Delete NewRowBookmark
            End If
            
        Case "INSDETAIL"
            Set GA = Frm.GridArrayDetail
            If IsEmpty(GA.value(0, 0)) Then
                DBGrid.PostMsg geRefresh
            End If
            GA.AddData RowBuf, NewRowBookmark, DBGrid.Columns
        Case Else
            Exit Sub
    End Select
End Sub



Private Sub DBGrid_UnboundDeleteRow(bookmark As Variant)
    Select Case DBGrid.Tag
        Case "INSDETAIL"
            Set GA = Frm.GridArrayDetail
            GA.Delete bookmark
        Case Else
            Exit Sub
    End Select
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub DBGrid_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal OffSet As Long, ApproximatePosition As Long)
Private Sub DBGrid_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid60.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
    Select Case DBGrid.Tag
        Case "MULTISOC"
            Set GA = Frm.GridArray
        Case "DETAIL", "INSDETAIL"
            Set GA = Frm.GridArrayDetail
        Case "DETAIL2"
            Set GA = Frm.GridArrayDetail2
        Case "DETAIL1", "INSDETAIL1"
            Set GA = Frm.GridArrayDetail1
        Case "SEARCH"
            Set GA = Frm.GridArraySearch
        Case "SELECT"
            Set GA = Frm.GridArraySelect
        Case Else
            Exit Sub
    End Select
    GA.ReadData RowBuf, StartLocation, offset, ApproximatePosition
End Sub

'*** APEX Migration Utility Code Change ***
'Private Sub DBGrid_UnboundWriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, WriteLocation As Variant)
Private Sub DBGrid_UnboundWriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, WriteLocation As Variant)
    Dim blnConcurError As Boolean
    Dim i As Integer
    
    Select Case DBGrid.Tag
        Case "MULTISOC"
            Set GA = Frm.GridArray
            GA.WriteData RowBuf, WriteLocation
'            RowToRecord WriteLocation, Frm.DataClass.DS.MultiSoc.DS, _
                        Frm.GridArray, Frm, ColTags
            With Frm.DataClass.DS.MultiSoc
'                ' --- Codice
'                .DS.Fields("Codice").FindValue = _
'                    Frm.DataClass.DS.Fields(Frm.DataClass.DS.KeyFields(1)).value
'                ' ---
'                ' ---
'                For i = 2 To .DS.KeyFields.Count
'                    .DS.Fields(.DS.KeyFields(i)).FindValue = _
'                        GA.value(WriteLocation, 0)
'                Next i
'                ' ---
'                .Find
                RowToRecord WriteLocation, .DS, _
                            Frm.GridArray, Frm, ColTags
                .BeginSave blnConcurError
                If blnConcurError Then
                    If MsgBox(MSG_CONCUR_ERR, _
                               vbOKCancel + vbExclamation + vbDefaultButton2, _
                              "Attenzione") = vbOK Then
                        .DS.DA.EndTrans ETCOMMIT
                        .DS.EndSave
                    Else
                        .DS.DA.EndTrans ETROLLBACK
'                        RowBuf.RowCount = 0
                    End If
                    blnConcurError = False
                Else
                    .DS.DA.EndTrans ETCOMMIT
                    .DS.EndSave
                End If
            End With
        Case "DETAIL"
            Set GA = Frm.GridArrayDetail
            GA.WriteData RowBuf, WriteLocation
'            RowToRecord WriteLocation, Frm.DetailClass.DS, Frm.GridArrayDetail, Frm, ColTags
            With Frm.DetailClass
                RowToRecord WriteLocation, .DS, Frm.GridArrayDetail, Frm, ColTags ', True
                .BeginSave blnConcurError
                If blnConcurError Then
                    If MsgBox(MSG_CONCUR_ERR, _
                               vbOKCancel + vbExclamation + vbDefaultButton2, _
                              "Attenzione") = vbOK Then
                        .DS.DA.EndTrans ETCOMMIT
                        .DS.EndSave
                    Else
                        .DS.DA.EndTrans ETROLLBACK
                        RowBuf.RowCount = 0
                    End If
                    blnConcurError = False
                Else
                    .DS.DA.EndTrans ETCOMMIT
                    .DS.EndSave
                End If
            End With
        Case "DETAIL2"
            Set GA = Frm.GridArrayDetail2
            GA.WriteData RowBuf, WriteLocation
            'RowToRecord WriteLocation, Frm.DetailClass.DS, Frm.GridArrayDetail, Frm, ColTags
            With Frm.DetailClass2
                RowToRecord WriteLocation, .DS, Frm.GridArrayDetail2, Frm, ColTags2 ', True
                .BeginSave blnConcurError
                If blnConcurError Then
                    If MsgBox(MSG_CONCUR_ERR, _
                               vbOKCancel + vbExclamation + vbDefaultButton2, _
                              "Attenzione") = vbOK Then
                        .DS.DA.EndTrans ETCOMMIT
                        .DS.EndSave
                    Else
                        .DS.DA.EndTrans ETROLLBACK
                        RowBuf.RowCount = 0
                    End If
                    blnConcurError = False
                Else
                    .DS.DA.EndTrans ETCOMMIT
                    .DS.EndSave
                End If
        End With
        Case "DETAIL1"
            Set GA = Frm.GridArrayDetail1
            GA.WriteData RowBuf, WriteLocation
            With Frm.DetailClass1
                RowToRecord WriteLocation, .DS, Frm.GridArrayDetail1, Frm, ColTags1 ', True
                .BeginSave blnConcurError
                If blnConcurError Then
                    If MsgBox(MSG_CONCUR_ERR, _
                               vbOKCancel + vbExclamation + vbDefaultButton2, _
                              "Attenzione") = vbOK Then
                        .DS.DA.EndTrans ETCOMMIT
                        .DS.EndSave
                    Else
                        .DS.DA.EndTrans ETROLLBACK
                        RowBuf.RowCount = 0
                    End If
                    blnConcurError = False
                Else
                    .DS.DA.EndTrans ETCOMMIT
                    .DS.EndSave
                End If
            End With
        Case "INSDETAIL"
            Set GA = Frm.GridArrayDetail
            GA.WriteData RowBuf, WriteLocation
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub Frm_Load()
    Frm.Enabled = True
End Sub

Private Sub Frm_Unload(Cancel As Integer)
    If Forms.Count = 2 Then
        SetMenuState True
        SetMenuConfiguration
    End If
End Sub


Private Sub Lbl_DblClick()
    If Lbl.MousePointer = vbArrowQuestion Then
        DoSearch Frm.DataClass.DS.Fields(Lbl.Tag).ForeignTableEntity, _
                 Frm, _
                 Lbl.Tag, _
                 Frm.DataClass.DS.FindControl(Lbl.Tag, "MASKEDBOX").TabIndex
    End If
End Sub


Private Sub mskMaskEdBox_Change()
    If mskMaskEdBox.Text = "?" Then
        DoSearch Frm.DataClass.DS.Fields(mskMaskEdBox.Tag).ForeignTableEntity, Frm, mskMaskEdBox.Tag, mskMaskEdBox.TabIndex
    End If
End Sub

Private Sub mskMaskEdBox_GotFocus()
    With mskMaskEdBox
        Frm.DataClass.DS.SetClientFieldFocus .Tag
        If Len(.Text) > 0 Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub


Private Sub mskMaskEdBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub mskMaskEdBox_LostFocus()
    Dim strText As String
    Dim blnValidate As Boolean
    Dim DC As Object
    Dim strFieldKey As String
    
    strFieldKey = mskMaskEdBox.Tag
    strText = mskMaskEdBox.Text
    If strText = DATE_MASK Then
        strText = vbNullString
    End If
    
    With Frm.DataClass.DS
        Select Case .FormType
            Case ftEdtFil1, ftEdtFil1S
                blnValidate = .ValidateClientField(strFieldKey, strText)
            Case ftEdtRcd1, ftEdtRcd1S
                If .KeyFields.Count = 0 Then
                    blnValidate = .ValidateClientField(strFieldKey, strText)
                Else
                    If strFieldKey <> .KeyFields(1) Then
                        blnValidate = .ValidateClientField(strFieldKey, strText)
                    End If
                End If
        End Select
        If .OmonimySearch Then
            .OmonimySearch = False
            DoSearch Frm.DataClass.DS.Fields(strFieldKey).ForeignTableEntity, Frm, strFieldKey, mskMaskEdBox.TabIndex
        End If
        
        If strText <> "?" Then
            ' --- Aggiorno l'eventuale descrizione collegata al campo ---
            FieldUpdateDescription strFieldKey, Frm, mskMaskEdBox
            strText = mskMaskEdBox.Text
            FieldSetDescription strFieldKey, Frm, strText, mskMaskEdBox
        End If
    End With
    mskMaskEdBox.Text = UCase(strText)

End Sub


