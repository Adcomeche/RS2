VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayGridCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' This is the definition of the ArrayGridCls class.  In addition
' to the actual Unbound event code for all of the grids, there
' are some additional class members defined to
' Special Note:  Variables that store Row indices will be
' of datatype Long (It is reasonable to assume there may
' be larger than 32,767 rows).  Those storing Column
' indices will be of type Integer. (It is very unlikely
' there will be more than 32,767 columns.)
'
' We declare three private class members here:
'     GridArray --  The array which holds our data to display
'     MaxCol    --  Holds number of columns in the GridArray
'     MaxRow    --  Holds number of rows in the GridArray
'
' These members are controlled by the Property Get/Let access
' properties and functions below.  For example, you can only
' set MaxRow by calling the SetDims subroutine, and you can
' only read it by reading the RowCount property of the class.
' More information about these is given below
'
Private GridArray As Variant
Private MaxCol As Integer
Private MaxRow As Long

Public Sub ClearData()
    
    SetDims 1, MaxCol
'    GridArray = Empty
End Sub


Public Property Get NoData() As Boolean
    NoData = IsEmpty(GridArray)
End Property

Public Property Get RowCount()
    '
    ' The RowCount property of the grid class is an access
    ' function to the private MaxRow member which holds the
    ' number of rows in the array of the specified class.
    ' Since there is only a "Get" property function and no
    ' "Let" function, the RowCount "property" of the grid
    ' class is read-only.
    '
    RowCount = MaxRow
    
End Property

Public Property Get ColCount()
    '
    ' The ColCount property of the grid class is an access
    ' function to the private MaxCol member which holds the
    ' number of columns in the array of the specified class.
    ' Since there is only a "Get" property function and no
    ' "Let" function, the ColCount "property" of the grid
    ' class is read-only.
    '
    ColCount = MaxCol
    
End Property

Public Property Let Value(ByVal Rows As Long, ByVal Cols As Integer, ByVal Val As Variant)
    '
    ' The Value property of the grid class is really an access
    ' function to modify the GridArray data structure.  This is
    ' the only way to modify the array outside of the class
    ' itself.  This is the "Let" function which is called when
    ' code like:  MyClass.Value(I, J) = X  is used to change
    ' the GridArray outside of the class.  Although we don't do
    ' it here, you can do error checking or other operations in
    ' this procedure when the array is modified.  Since Value has
    ' both a "Let" and "Get" property, it is read/write.
    '
    GridArray(Cols, Rows) = Val
    
End Property

Public Property Get Value(ByVal Rows As Long, ByVal Cols As Integer)
    '
    ' The Value property of the grid class is really an access
    ' function to modify the GridArray data structure.  This is
    ' the only way to modify the array outside of the class
    ' itself.  This is the "Get" function which is called when
    ' code like:  X = MyClass.Value(I, J)  is used to fetch data
    ' from the GridArray outside of the class.  Although
    ' we don't do it here, you can do error checking or other
    ' operations in this procedure when the array is read.
    ' Since Value has both a "Let" and "Get" property, it is
    ' read/write.
    '
    Value = GridArray(Cols, Rows)
    
End Property

Public Sub SetDims(ByVal Rows As Long, ByVal Cols As Integer)
    '
    ' SetDims is a class member that is used to set the
    ' dimensions of the private GridArray member.  This is
    ' an access function to the private MaxRow and MaxCol
    ' properties that store the dimensions of the GridArray.
    ' This function is the ONLY way to change MaxRow and
    ' MaxCol from outside of the class.  We automatically
    ' re-dimension the GridArray here also.
    '
    ReDim GridArray(0 To Cols - 1, 0 To Rows - 1)
    MaxRow = Rows
    MaxCol = Cols
    
End Sub

Public Sub ChangeRowDim(ByVal Rows As Integer)
    '
    ' ChangeRowDim is a class member that is used to change the
    ' row dimension of the private GridArray member.  This is
    ' an access function to the private MaxRow property. Since
    ' you can only use the VB ReDim Preserve statement to
    ' re-dimension the LAST dimension of the array and still
    ' keep its contents.  So this function will allow you to
    ' resize the number of GridArray rows outside of the class --
    ' this lets you add or delete rows through code.  If you need
    ' to be able to Add or Remove columns at run-time, you should
    ' call SetDims and re-load the array with your data, then
    ' set the grid's Bookmark to Null, and ReBind the grid.
    ' Alternatively, you can write your own access function to
    ' change the number of rows and columns and re-load the array
    ' yourself.
    '
    ReDim Preserve GridArray(0 To MaxCol - 1, 0 To Rows - 1)
    MaxRow = Rows
    
End Sub

Private Function MakeBookmark(Index As Long) As Variant
    '
    ' This support function is not used directly by the
    ' unbound events.  It is a good idea to create a
    ' MakeBookmark function such that all bookmarks can
    ' be created in the same way.  Thus the method by
    ' which bookmarks are created is consistent and easy
    ' to modify.  This function creates a bookmark when
    ' given an array row index.
    '
    ' Since we have data stored in an array, we will just
    ' use the array index as our bookmark.  We will convert
    ' it to a string first, using the Str$ function.  Thus,
    ' if Index = 27, the Bookmark that is created is the
    ' string, " 27".  (Str$ always leaves a leading space
    ' for the sign of the number.)
    '
    MakeBookmark = Str$(Index)
    
End Function

Private Function IndexFromBookmark(Bookmark As Variant, _
        OffSet As Long) As Long
    '
    ' IndexFromBookmark computes the row index that
    ' corresponds to a row that is (Offset) rows from the
    ' row specified by the Bookmark parameter.  For example,
    ' if Bookmark refers to the index 50 of the dataset
    ' array and Offset = -10, then IndexFromBookmark will
    ' return 50 + (-10), or 40.  Thus to get the index of
    ' the row specified by the bookmark itself, call
    ' IndexFromBookmark with an Offset of 0.  If the given
    ' Bookmark is Null, it refers to BOF or EOF.  If
    ' Offset < 0, the grid is requesting rows before the
    ' row specified by Bookmark, and so we must be at EOF
    ' because prior rows do not exist for BOF.  Conversely,
    ' if Offset > 0, we are at BOF.
    '
    Dim Index As Long
    
    If IsNull(Bookmark) Then
        If OffSet < 0 Then
            ' Bookmark refers to EOF.  Since (MaxRow - 1)
            ' is the index of the last record, we can use
            ' an index of (MaxRow) to represent EOF.
            Index = MaxRow + OffSet
        Else
            ' Bookmark refers to BOF.  Since 0 is the index
            ' of the first record, we can use an index of
            ' -1 to represent BOF.
            Index = -1 + OffSet
        End If
    Else
        ' Convert string to long integer
        Index = Val(Bookmark) + OffSet
    End If
    
    ' Check to see if the row index is valid:
    '   (0 <= Index < MaxRow).
    ' If not, set it to a large negative number to
    ' indicate that it is invalid.
    If Index >= 0 And Index < MaxRow Then
       IndexFromBookmark = Index
    Else
       IndexFromBookmark = -9999
    End If
    
End Function

Private Function GetNewBookmark() As Variant
    '
    ' GetNewBookmark is called from the AddData grid class
    ' procedure when we need to create a bookmark for a
    ' newly added record (row).  Also, we increment the
    ' row count of the array (MaxRow) to reflect the new
    ' GridArray size, and re-dimension the array to fit
    ' the new size.
    '
    
    ' Reserve space for the new row in the array
    ReDim Preserve GridArray(0 To MaxCol - 1, 0 To MaxRow)
    
    ' Since the row index of the last record is
    ' (MaxRow - 1), the next available index for the row
    ' to be added is (MaxRow), so create a bookmark for
    ' the new row using MaxRow as an index.
    GetNewBookmark = MakeBookmark(MaxRow)
    
    ' Increment the number of rows in our dataset
    MaxRow = MaxRow + 1
    
End Function

Private Function GetRelativeBookmark(Bookmark As Variant, _
        OffSet As Long) As Variant
    '
    ' GetRelativeBookmark is used to get a bookmark for a
    ' row that is a specified number of rows away from the
    ' given row.  Offset specifies the number of rows to
    ' move.  A positive Offset indicates that the desired
    ' row is after the one referred to by Bookmark, and a
    ' negative Offset means it is before the one referred
    ' to by Bookmark.
    '
    Dim Index As Long
    
    ' Compute the row index for the desired row
    Index = IndexFromBookmark(Bookmark, OffSet)
    If Index < 0 Or Index >= MaxRow Then
        ' Index refers to a row before the first or after
        ' the last, so just return Null.
        GetRelativeBookmark = Null
    Else
        GetRelativeBookmark = MakeBookmark(Index)
    End If
    
End Function

Private Function GetUserData(Bookmark As Variant, _
        Col As Integer) As Variant
    '
    ' In this example, GetUserData is called by the ReadData
    ' grid class procedure to ask the user what data should be
    ' displayed in a specific cell in the grid.  The grid
    ' row the cell is in is the one referred to by the
    ' Bookmark parameter, and the column it is in it given
    ' by the Col parameter.  GetUserData is called on a
    ' cell-by-cell basis.
    '
    Dim Index As Long
        
    ' Figure out which row the bookmark refers to
    Index = IndexFromBookmark(Bookmark, 0)
    
    If Index < 0 Or Index >= MaxRow Or _
       Col < 0 Or Col >= MaxCol Then
        ' Cell position is invalid, so just return null to
        ' indicate failure
        GetUserData = Null
    Else
        GetUserData = GridArray(Col, Index)
    End If
    
End Function

Private Function StoreUserData(Bookmark As Variant, _
        Col As Integer, Userval As Variant) As Boolean
    '
    ' In this example, StoreUserData is called by the WriteData
    ' grid class procedure to ask the user what data should be
    ' stored from the grid into the array.  The grid row the cell
    ' is in is the one referred to by the Bookmark parameter, and
    ' the column it is in it given by the Col parameter.
    ' StoreUserData is called on a cell-by-cell basis.
    '
    Dim Index As Long
    
    ' Figure out which row the bookmark refers to
    Index = IndexFromBookmark(Bookmark, 0)
    
    If Index < 0 Or Index >= MaxRow Or Col < 0 Or _
       Col >= MaxCol Then
        ' Cell position is invalid, so just return null
        ' to indicate failure
        StoreUserData = False
    Else
        StoreUserData = True
        GridArray(Col, Index) = Userval
    End If
    
End Function

Public Function DeleteRow(Bookmark As Variant) As Boolean
    '
    ' DeleteRow is called by DeleteRow to logically delete a
    ' record (row) from the GridArray.  The row to be deleted is
    ' given by the Bookmark parameter.  DeleteRow does this by
    ' physically removing the row from the array, shifting the
    ' remaining rows up one, to fill in the "hole" left by the
    ' deleted record.  Then the value of MaxRow is decremented
    ' by one to reflect the new array size.  If the delete fails
    ' for any reason. we return False, otherwise we return true.
    ' This has been declared as a public function so that you may
    ' use it to delete a row from the grid through code, outside
    ' of the class.  If use DeleteRow to programmatically delete
    ' a row, afterwards you will want to:
    '    (1) Set a new current row for the grid
    '    (2) ReBind the grid
    '
    Dim i As Long, Index As Long
    Dim j As Integer

    ' Figure out which row index the bookmark refers to
    Index = IndexFromBookmark(Bookmark, 0)
    
    If Index < 0 Or Index >= MaxRow Then
        ' Row index is invalid, so exit the delete,
        ' indicating failure.
        DeleteRow = False
        Exit Function
    End If
    
    ' Decrement the number of rows in the dataset
    MaxRow = MaxRow - 1
    
    ' Shift the data in the array to fill the empty space
    ' vacated by the deleted row
    For i = Index To MaxRow - 1
        For j = 0 To MaxCol - 1
            GridArray(j, i) = GridArray(j, i + 1)
        Next j
    Next i
    
    ' Resize array to free storage space used by deleted
    ' row
    If MaxRow > 0 Then
        ' If MaxRow = 0, then (MaxRow - 1) equals -1.
        ' This causes an error in the statement below, so
        ' we handle this special case in the Else clause.
        ReDim Preserve GridArray(0 To MaxCol - 1, 0 _
              To MaxRow - 1)
    Else
        ReDim Preserve GridArray(0 To MaxCol - 1, 0)
    End If
    DeleteRow = True
End Function

' Read event
Public Sub ReadData(RowBuf As RowBuffer, StartLocation As Variant, _
        ByVal OffSet As Long, ApproximatePosition As Long)
    '
    ' This ReadData procedure is where the processing is done
    ' for all UnboundReadDataEx events fired by the grid(s).
    ' Whenver an UnboundReadDataEx event is fired for a grid,
    ' this class member is called to fetch the data the
    ' grid asks for.
    '
    ' UnboundReadData is fired by an unbound grid whenever
    ' it requires data for display.  This event will fire
    ' when the grid is first shown, when Refresh or ReBind
    ' is used, when the grid is scrolled, and after a
    ' record in the grid is modified and the user commits
    ' the change by moving off of the current row.  The
    ' grid fetches data in "chunks", and the number of rows
    ' the grid is asking for is given by RowBuf.RowCount.
    '
    ' RowBuf is the row buffer where you place the data
    ' the bookmarks for the rows that the grid is
    ' requesting to display.  It will also hold the number
    ' of rows that were successfully supplied to the grid.
    '
    ' StartLocation is a bookmark which, together with
    ' Offset, specifies the row for the programmer to start
    ' transferring data.  A StartLocation of Null indicates
    ' a request for data from BOF or EOF.
    '
    ' Offset specifies the relative position (from
    ' StartLocation) of the row for the programmer to start
    ' transferring data.  A positive number indicates a
    ' forward relative position while a negative number
    ' indicates a backward relative position.  Regardless
    ' of whether the rows to be read are before or after
    ' StartLocation, rows are always fetched going forward
    ' (this is why there is no ReadPriorRows parameter to
    ' the procedure).
    '
    ' If you page down on the grid, for instance, the new
    ' top row of the grid will have an index greater than
    ' the StartLocation (Offset > 0).  If you page up on
    ' the grid, the new index is less than that of
    ' StartLocation, so Offset < 0.  If StartLocation is
    ' a bookmark to row N, the grid always asks for row
    ' data in the following order:
    '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
    '
    ' ApproximatePosition is a value you can set to indicate
    ' the ordinal position of (StartLocation + Offset).
    ' Setting this variable will enhance the ability of the
    ' grid to display its vertical scroll bar accurately.
    ' If the exact ordinal position of the new location is
    ' not known, you can set it to a reasonable,
    ' approximate value, or just ignore this parameter.
    '
    Dim ColIndex As Integer, j As Integer
    Dim RowsFetched As Integer, i As Long
    Dim NewPosition As Long, Bookmark As Variant
    Dim varData As Variant
    
    If Not IsEmpty(GridArray) Then
    RowsFetched = 0
    
    For i = 0 To RowBuf.RowCount - 1
        ' Get the bookmark of the next available row
        Bookmark = GetRelativeBookmark(StartLocation, _
                   OffSet + i)
    
        ' If the next row is BOF or EOF, then stop fetching
        ' and return any rows fetched up to this point.
        If IsNull(Bookmark) Then Exit For
    
        ' Place the record data into the row buffer
        For j = 0 To RowBuf.ColumnCount - 1
            ColIndex = RowBuf.ColumnIndex(i, j)
            varData = GetUserData(Bookmark, _
                                 ColIndex)
            If Not IsNull(varData) Then
                RowBuf.Value(i, j) = CStr(varData)
            End If
        Next j
    
        ' Set the bookmark for the row
        RowBuf.Bookmark(i) = Bookmark
    
        ' Increment the count of fetched rows
        RowsFetched = RowsFetched + 1
    Next i
    
    ' Tell the grid how many rows were fetched
    RowBuf.RowCount = RowsFetched
    
    ' Set the approximate scroll bar position.  Only
    ' nonnegative values of IndexFromBookmark() are valid.
    NewPosition = IndexFromBookmark(StartLocation, OffSet)
    If NewPosition >= 0 Then _
        ApproximatePosition = NewPosition
        
    End If
End Sub

Public Sub WriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
    '
    ' This WriteData procedure is where the processing is done
    ' for all UnboundWriteData events fired by the grid(s).
    ' Whenver an UnboundWriteData event is fired for a grid,
    ' this class member is called to save the data in the
    ' current grid row to the GridArray.
    '
    ' UnboundWriteData event is fired when the user
    ' modifies an existing row within an unbound grid and
    ' attempts to commit the changes by moving to another
    ' row or calling the Update method of the grid.  This
    ' notifies the program that data has changed in the
    ' grid and it needs to be stored in the data source.
    '
    ' RowBuf is the row buffer from which you retrieve the
    ' data to be stored in the data source.  Since only
    ' one row can be updated at a time, RowBuf.RowCount
    ' always equals 1.  If a particular column in the row
    ' has not been changed, its element in the RowBuf.Value
    ' array will be null.
    '
    ' WriteLocation is a bookmark that identifies the row
    ' to be updated.
    '
    ' Assume that a function StoreUserData(bookmark, col,
    ' value) takes a row bookmark, a column index, and a
    ' variant with the appropriate data to be stored in
    ' an array or database.  The function returns True if
    ' the data is acceptable and can be stored, False
    ' otherwise.
    '
    Dim i As Integer
    
    ' Loop over all the columns of the row, storing
    ' non-Null values
    For i = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i)) Then
            If Not StoreUserData(WriteLocation, i, _
               RowBuf.Value(0, i)) Then
                RowBuf.RowCount = 0 ' Tell the grid the
                                    ' update failed, so
            End If                  ' exit the event.
        End If
    Next i
    
End Sub

Public Sub AddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant, Cols As Columns)
    '
    ' This AddData procedure is where the processing is done
    ' for all UnboundAddData events fired by the grid(s).
    ' Whenver an UnboundAddData event is fired for a grid,
    ' this class member is called to add the data for the
    ' new grid row to the GridArray.
    '
    ' UnboundAddData is fired when the user adds a new
    ' row of data to an unbound grid to alert your
    ' application that it must add a new row of data to
    ' its unbound dataset.  This event will NOT be fired
    ' if the grid's AllowAddNew property is set to False.
    '
    ' Assume that a function StoreUserData(bookmark, col,
    ' value) takes a row bookmark, a column index, and a
    ' variant with the appropriate data to be stored in an
    ' array.  The function returns True if the data is
    ' acceptable and can be stored, False otherwise.
    '
    ' First, get a bookmark for the new row.  Do this with
    ' a function, GetNewBookmark.  This function allocates
    ' a new row of data in the storage medium (GridArray)
    ' and returns a variant containing a bookmark for the
    ' newly added row.
    '
    Dim NewVal As Variant
    Dim i As Integer
    
    ' Create a new row in the array and get a bookmark for
    ' the new row
    NewRowBookmark = GetNewBookmark()

    ' Loop over all the columns of the row, storing
    ' non-Null values
    For i = 0 To RowBuf.ColumnCount - 1
        ' Note that since only one row of data may be
        ' added at a time, the first index of the row
        ' buffer Value array is always 0.
        NewVal = RowBuf.Value(0, i)
        If IsNull(NewVal) Then
            ' The row buffer does not contain a value for
            ' this column.  A default value should be set.
            ' A convenient value is the default value for
            ' the column.
            NewVal = Cols(i).DefaultValue
        End If
    
        ' Now store the new values.
        If Not StoreUserData(NewRowBookmark, i, NewVal) Then
            ' Storage of the data has failed.  Delete the
            ' added row using a function, DeleteRow, which
            ' takes a bookmark as an argument.  Also,
            ' indicate that the update has failed by setting
            ' RowCount to 0.
            DeleteRow NewRowBookmark
            RowBuf.RowCount = 0
            Exit Sub          ' It failed, so exit the event.
        End If
    Next i
    
End Sub

Public Sub Delete(Bookmark As Variant)
    '
    ' This Delete procedure is where the processing is done
    ' for all UnboundDeleteRow events fired by the grid(s).
    ' Whenver an UnboundDeleteRow event is fired for a grid,
    ' this class member is called to delete the specified row
    ' from the GridArray.
    '
    ' UnboundDeleteRow is fired when the user deletes an
    ' existing row in the Unbound grid.  This event alerts
    ' the user that the program must delete the row
    ' specified by the Bookmark parameter from the dataset.
    ' Here, an auxiliary function (DeleteRow) is used to do
    ' this.
    '
    If Not DeleteRow(Bookmark) Then Bookmark = Null
    
End Sub

