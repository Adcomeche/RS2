Attribute VB_Name = "COMGRIDUTILS"
Option Explicit

Public Const MAX_COLS = 3

Public Sub ClearUnboundGrid(ByRef TDBG As TDBGrid, _
                            ByRef GridArray As ArrayGridCls, ByVal ColNum As Integer)
    GridArray.SetDims 1, ColNum
    TDBG.bookmark = 1
End Sub

Public Sub ClearXArray(XA As XArrayDB, TDBG As TDBGrid)
    Set XA = New XArrayDB
    XA.ReDim 0, 0, 0, 0
    TDBG.Array = XA
    TDBG.ReBind
End Sub


Public Sub RefreshXArray(XA As XArrayDB, TDBG As TDBGrid, List As Variant)
    Dim i As Long
    Dim j As Long
    Dim lngFrom1 As Long
    Dim lngTo1 As Long
    Dim lngFrom2 As Long
    Dim lngTo2 As Long
    
    Set XA = New XArrayDB
    With XA
        .ReDim LBound(List, 2), UBound(List, 2), LBound(List, 1), UBound(List, 1)
        lngFrom1 = .LowerBound(1)
        lngTo1 = .UpperBound(1)
        lngFrom2 = .LowerBound(2)
        lngTo2 = .UpperBound(2)
        For i = lngFrom1 To lngTo1
            For j = lngFrom2 To lngTo2
                .value(i, j) = List(j, i)
            Next j
        Next i
    End With
    With TDBG
        .Array = XA
        .ReBind
    End With
End Sub

Public Sub RefreshUnboundGrid(ByRef TDBG As TDBGrid, ByRef List As Variant, _
                              ByRef GridArray As ArrayGridCls, ByVal HowManyRows As Long, ByVal ColNum As Integer)
    Dim lngRows As Long
    Dim intCols As Integer
    Dim R As Long
    Dim C As Integer
    
    If HowManyRows > 0 Then
        lngRows = UBound(List, 2) + 1
        intCols = UBound(List, 1) + 1
        
        If lngRows > 0 Then
            GridArray.SetDims lngRows, intCols
            
            For R = 0 To lngRows - 1
                For C = 0 To intCols - 1
                    GridArray.value(R, C) = List(C, R)
                Next C
            Next R
            With TDBG
                .ApproxCount = GridArray.RowCount
                .bookmark = 0
                .ReBind
            End With
        End If
    Else
        GridArray.SetDims 1, ColNum
        TDBG.bookmark = 1
    End If
End Sub

Public Sub MyExportGridToFile(strPathFile As String, _
                              MyGrid As TDBGrid, _
                              MyList As Variant, _
                              MyProgresBar As ProgressBar, _
                              intStartCol As Integer, _
                              Optional intAppend As Integer = 1, _
                              Optional blnShowHiddenColumns = True)
    Dim iCol As Integer
    Dim intNumSplits As Integer
    Dim intNumColSplit As Integer
    Dim lngCountRow As Long
    Dim lngCountCol As Long
    Dim intCountSplit As Integer
    
    Dim blnShowColumn As Boolean
    
    Dim MyFile As CTextFile
    
    Dim blnFileOpen As Boolean
    Set MyFile = New CTextFile
    
    blnFileOpen = False
     
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    If Not IsEmpty(MyList) Then
        MyProgresBar.Min = 0
        MyProgresBar.Max = UBound(MyList, 2) + 1
    
        ' Creo Dinamicamente la struttura del File, in base alla definizione delle colonne
        With MyFile
            iCol = -1
            For lngCountCol = intStartCol To MyGrid.Columns.Count - 1
                DoEvents
                ' Controllo se devo mostrare la colonna sul file
                If Not blnShowHiddenColumns Then
                    blnShowColumn = False
                    
                    For intCountSplit = 0 To MyGrid.Splits.Count - 1
                        If MyGrid.Splits(intCountSplit).Columns(lngCountCol).Visible Then
                            blnShowColumn = True
                        End If
                    Next intCountSplit
                Else
                    blnShowColumn = True
                End If
                
                If blnShowColumn Then
                    iCol = iCol + 1
                    
                    Select Case MyGrid.Columns(lngCountCol).Alignment
                        Case 0
                            .Add MyGrid.Columns(lngCountCol).Caption, 41, CStr(iCol), , fftAlfa
                            
                        Case Else
                            .Add MyGrid.Columns(lngCountCol).Caption, 41, CStr(iCol), , fftNumerico
                            
                    End Select
                End If
            Next lngCountCol
            .Delimiter = vbTab
        End With
        
        If MyFile.OpenFile(strPathFile, intAppend) Then
            blnFileOpen = True
            ' Scrivo sul file il contenuto delle righe di intestazione
            With MyFile
                iCol = 0
                ' Scrivo la riga con i nomi degli splits
                For intNumSplits = 0 To MyGrid.Splits.Count - 1
                    If MyGrid.Splits(intNumSplits).Caption <> vbNullString Then
                        For intNumColSplit = iCol To MyGrid.Splits(intNumSplits).Columns.Count - 1
                            If MyGrid.Splits(intNumSplits).Columns(intNumColSplit).Visible Then
                                .Item(CStr(intNumColSplit)).value = Trim(MyGrid.Splits(intNumSplits).Caption)
                                iCol = iCol + 1
                                Exit For
                            Else
                                If blnShowHiddenColumns Then
                                    .Item(CStr(intNumColSplit)).value = vbNullString
                                End If
                                iCol = iCol + 1
                            End If
                        Next intNumColSplit
                    End If
                Next intNumSplits
                If iCol > 0 Then
                    .WriteLine True
                End If
            
                iCol = -1
                For lngCountCol = intStartCol To MyGrid.Columns.Count - 1
                    DoEvents
                    ' Controllo se devo mostrare la colonna sul file
                    If Not blnShowHiddenColumns Then
                        blnShowColumn = False
                        
                        For intCountSplit = 0 To MyGrid.Splits.Count - 1
                            If MyGrid.Splits(intCountSplit).Columns(lngCountCol).Visible Then
                                blnShowColumn = True
                            End If
                        Next intCountSplit
                    Else
                        blnShowColumn = True
                    End If
                    
                    If blnShowColumn Then
                        iCol = iCol + 1
                        .Item(CStr(iCol)).value = Trim(MyGrid.Columns(lngCountCol).Caption)
                    End If
                Next lngCountCol
                
                .WriteLine True
            End With
    
            ' Scrivo sul file il contenuto delle restanti righe
            With MyFile
                For lngCountRow = 0 To UBound(MyList, 2)
                    DoEvents
                    iCol = -1
                    For lngCountCol = intStartCol To MyGrid.Columns.Count - 1
                       ' Controllo se devo mostrare la colonna sul file
                       If Not blnShowHiddenColumns Then
                           blnShowColumn = False
                           
                           For intCountSplit = 0 To MyGrid.Splits.Count - 1
                               If MyGrid.Splits(intCountSplit).Columns(lngCountCol).Visible Then
                                   blnShowColumn = True
                               End If
                           Next intCountSplit
                       Else
                           blnShowColumn = True
                       End If
                       
                       If blnShowColumn Then
                           iCol = iCol + 1
                           '.Item(CStr(iCol)).value = IIf(.Item(CStr(iCol)).FieldType = fftAlfa, "'", vbNullString) & Trim(MyList(lngCountCol, lngCountRow))
                           .Item(CStr(iCol)).value = Trim(MyList(lngCountCol, lngCountRow))
                       End If
                    Next lngCountCol
                    MyProgresBar.value = lngCountRow + 1
                    
                    .WriteLine True
                Next lngCountRow
            End With
        End If
    
        MyFile.CloseFile
        Screen.MousePointer = vbNormal
    End If
    
    Set MyList = Nothing
    Exit Sub
    
ErrHandler:
    If blnFileOpen = True Then
        MyFile.CloseFile
    End If
    Screen.MousePointer = vbNormal
End Sub

Public Sub MyExportGridToDelimitedFile(strPathFile As String, _
                              MyGrid As TDBGrid, _
                              MyList As Variant, _
                              MyProgresBar As ProgressBar, _
                              intStartCol As Integer, _
                              Optional intAppend As Integer = 1, _
                              Optional blnShowHiddenColumns = True, _
                              Optional strSeparator As String = ";")
    Dim iCol As Integer
    Dim intNumSplits As Integer
    Dim intNumColSplit As Integer
    Dim lngCountRow As Long
    Dim lngCountCol As Long
    Dim intCountSplit As Integer
    
    Dim blnShowColumn As Boolean
    
    Dim blnFileOpen As Boolean
    Dim intFileNum As Integer
    Dim strLine As String
    
    blnFileOpen = False
     
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    If Not IsEmpty(MyList) Then
        MyProgresBar.Min = 0
        MyProgresBar.Max = UBound(MyList, 2) + 1
    
        intFileNum = FreeFile
        If intAppend = 0 Then
            Open strPathFile For Input Access Read As #intFileNum
        Else
            Open strPathFile For Output Access Write As #intFileNum
        End If
        
        blnFileOpen = True
        ' Scrivo sul file il contenuto delle righe di intestazione
        iCol = 0
        ' Scrivo la riga con i nomi degli splits
        For intNumSplits = 0 To MyGrid.Splits.Count - 1
            If MyGrid.Splits(intNumSplits).Caption <> vbNullString Then
                For intNumColSplit = iCol To MyGrid.Splits(intNumSplits).Columns.Count - 1
                    If MyGrid.Splits(intNumSplits).Columns(intNumColSplit).Visible Then
                        strLine = strLine & Trim(MyGrid.Splits(intNumSplits).Caption) & strSeparator
                        iCol = iCol + 1
                        Exit For
                    Else
                        If blnShowHiddenColumns Then
                            strLine = vbNullString
                        End If
                        iCol = iCol + 1
                    End If
                Next intNumColSplit
            End If
        Next intNumSplits
        If iCol > 0 Then
            Print #intFileNum, strLine
            strLine = vbNullString
        End If
    
        iCol = -1
        For lngCountCol = intStartCol To MyGrid.Columns.Count - 1
            DoEvents
            ' Controllo se devo mostrare la colonna sul file
            If Not blnShowHiddenColumns Then
                blnShowColumn = False
                
                For intCountSplit = 0 To MyGrid.Splits.Count - 1
                    If MyGrid.Splits(intCountSplit).Columns(lngCountCol).Visible Then
                        blnShowColumn = True
                    End If
                Next intCountSplit
            Else
                blnShowColumn = True
            End If
            
            If blnShowColumn Then
                iCol = iCol + 1
                strLine = strLine & Trim(MyGrid.Columns(lngCountCol).Caption) & strSeparator
            End If
        Next lngCountCol
        Print #intFileNum, strLine
        strLine = vbNullString
        
        ' Scrivo sul file il contenuto delle restanti righe
        For lngCountRow = 0 To UBound(MyList, 2)
            DoEvents
            iCol = -1
            For lngCountCol = intStartCol To MyGrid.Columns.Count - 1
               ' Controllo se devo mostrare la colonna sul file
               If Not blnShowHiddenColumns Then
                   blnShowColumn = False
                   
                   For intCountSplit = 0 To MyGrid.Splits.Count - 1
                       If MyGrid.Splits(intCountSplit).Columns(lngCountCol).Visible Then
                           blnShowColumn = True
                       End If
                   Next intCountSplit
               Else
                   blnShowColumn = True
               End If
               
               If blnShowColumn Then
                   iCol = iCol + 1
                   strLine = strLine & IIf(MyGrid.Columns(lngCountCol).Alignment = 0, "=""", vbNullString) & Trim(MyList(lngCountCol, lngCountRow)) & IIf(MyGrid.Columns(lngCountCol).Alignment = 0, """", vbNullString) & strSeparator
               End If
            Next lngCountCol
            Print #intFileNum, strLine
            strLine = vbNullString
            
            MyProgresBar.value = lngCountRow + 1
        Next lngCountRow
        
        Close #intFileNum

        Screen.MousePointer = vbNormal
    End If
    
    Set MyList = Nothing
    Exit Sub
    
ErrHandler:
    If blnFileOpen = True Then
        Close #intFileNum
    End If
    Screen.MousePointer = vbNormal
End Sub
