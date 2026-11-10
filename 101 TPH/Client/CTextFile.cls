VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Record di dati da importare"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CInputField"
Attribute VB_Ext_KEY = "Member0" ,"CInputField"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private intFileNum As Integer
Private mstrDelimiter As String


Public Sub ClearFields()
    Dim C As CTextFileField
    
    For Each C In mCol
        C.value = vbNullString
    Next C
    
End Sub


Public Property Let Delimiter(NewValue As String)
    mstrDelimiter = NewValue
End Property

Public Function EOF(Optional ByLine As Boolean = False, _
                    Optional Position As Long) As Boolean
    If Not ByLine Then
        EOF = VBA.EOF(intFileNum)
    Else
        EOF = Position >= LOF(intFileNum)
    End If
End Function


Public Property Let FileNum(NewValue As Integer)
    intFileNum = NewValue
End Property

Public Property Get FileNum() As Integer
    FileNum = intFileNum
End Property

Public Function GetLineUNIX(Optional Delimited As Boolean = True, _
                            Optional StringLen As Integer = 0, _
                            Optional ByRef Position As Long, _
                            Optional ByRef strLine As String, _
                            Optional ByRef blnEOF As Boolean) As Variant
    Dim lngLenGetLine As Long
    Dim lngCurChar As Long
    Dim strChar As String
        
    ' Ciclo di scorrimento fino a carattere di fine riga UNIX
    lngCurChar = Position
    If Position = 1 Then
        Line Input #intFileNum, strLine
    End If
    lngLenGetLine = 0
    strChar = Mid(strLine, lngCurChar, 1)
    Do While strChar <> Chr(10) And lngCurChar <= LOF(intFileNum)
        lngCurChar = lngCurChar + 1
        lngLenGetLine = lngLenGetLine + 1
        strChar = Mid(strLine, lngCurChar, 1)
        DoEvents
    Loop
    
    If lngCurChar > LOF(intFileNum) Then
        blnEOF = True
    End If
    
    GetLineUNIX = Mid(strLine, Position, lngLenGetLine)
    Position = lngCurChar + 1
End Function

Public Function GetLine(Optional Delimited As Boolean = True, _
                        Optional StringLen As Integer = 0, _
                        Optional ByRef Position As Long) As Variant
    Dim strLine As String
    
    If StringLen = 0 Then
        Line Input #intFileNum, strLine
    Else
        strLine = Input(StringLen, intFileNum)
        Position = Position + StringLen
    End If
    If Delimited Then
        GetLine = ParseString(strLine, FLD_DELIMITER)
    Else
        GetLine = strLine
    End If
End Function

Public Function CloseFile() As Boolean
    On Error GoTo ErrHandler
    Close #intFileNum
    CloseFile = True
    Exit Function
ErrHandler:
    CloseFile = False
End Function

Public Function OpenFile(strFileName As String, Mode As Integer) As Boolean
    On Error GoTo ErrHandler
    
    intFileNum = FreeFile
    If Mode = 0 Then
        Open strFileName For Input Access Read As #intFileNum
    Else
        Open strFileName For Output Access Write As #intFileNum
    End If
    OpenFile = True
    Exit Function
ErrHandler:
    OpenFile = False
End Function

Public Function ParseFixedLenFields(L As String, _
                                    FromField As Integer, _
                                    ToField As Integer) As Boolean
    On Error GoTo ErrHandler
    
    Dim FromChar As Integer
    Dim ToChar As Integer
    Dim i As Integer
    
    i = 1
    FromChar = 1
    For i = FromField To ToField
        ToChar = FromChar + mCol(i).Size
'        If ToChar > 0 Then
'            i = i + 1
'            ReDim Preserve varBuffer(i)
'            varBuffer(UBound(varBuffer)) = Trim(Mid(L, FromChar, ToChar - FromChar))
            mCol(i).value = Trim(Mid$(L, FromChar, ToChar - FromChar))
'        End If
        FromChar = ToChar '+ 1
    Next i
    ParseFixedLenFields = True
    Exit Function
ErrHandler:
    ParseFixedLenFields = False

End Function

Public Function ParseString(L As String, Delimiter As String) As Variant
    On Error GoTo ErrHandler
    
    ReDim varBuffer(0) As Variant
    Dim FromChar As Integer
    Dim ToChar As Integer
    Dim i As Integer
    
    i = 0
    FromChar = 1
    ToChar = InStr(FromChar, L, Delimiter)
    varBuffer(i) = Mid(L, FromChar, ToChar - FromChar)
    Do While FromChar <= Len(L)
        FromChar = ToChar + 1
        ToChar = InStr(FromChar, L, Delimiter)
        If ToChar > 0 Then
            i = i + 1
            ReDim Preserve varBuffer(i)
            varBuffer(UBound(varBuffer)) = Trim(Mid(L, FromChar, ToChar - FromChar))
        End If
    Loop
    ParseString = varBuffer
    Exit Function
ErrHandler:
    ParseString = Empty
End Function

Public Function Add(Name As String, Size As Integer, _
                    Optional sKey As String, _
                    Optional Default As Variant, _
                    Optional FieldType As TextFileFieldType = fftAlfa) As CTextFileField
    
    'create a new object
    Dim objNewMember As CTextFileField
    Set objNewMember = New CTextFileField
    
    With objNewMember
        .Name = Name
        .Size = Size
        If Not IsMissing(Default) Then
            .Default = Default
        End If
        .FieldType = FieldType
    End With

    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CTextFileField
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Function WriteLine(Optional Delimited As Boolean = False, Optional Espacios As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim strLine As String
    Dim strField As String
    Dim strValue As String
    
    For i = 1 To mCol.Count
        strField = Space(mCol(i).Size)
        If Espacios = True Then
        strValue = mCol(i).value
        Else
        strValue = Trim$("" & mCol(i).value)
        End If
        If Len(strValue) = 0 And Not IsEmpty(mCol(i).Default) Then
            strValue = mCol(i).Default
        End If
        
        If mCol(i).FieldType = 0 Then
            LSet strField = Left$(strValue, mCol(i).Size)
        Else
            RSet strField = Format$(strValue, String(Len(strField), "0"))
        End If
        
        If Delimited Then
            strField = strField & mstrDelimiter
        End If
        strLine = strLine & strField
    Next i
    Print #intFileNum, strLine
    WriteLine = True
    Exit Function
ErrHandler:
    WriteLine = False
End Function

Public Function WriteLineNofij(Optional Delimited As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim strLine As String
    Dim strField As String
    Dim strValue As String
    
    For i = 1 To mCol.Count

        strValue = Trim$("" & mCol(i).value)
        strField = Space(Len(strValue))
        If Len(strValue) = 0 And Not IsEmpty(mCol(i).Default) Then
            strValue = mCol(i).Default
        End If
        
        If mCol(i).FieldType = 0 Then
            LSet strField = Left$(strValue, mCol(i).Size)
        Else
            RSet strField = Format$(strValue, String(Len(strField), "0"))
        End If
        
        If Delimited Then
            strField = strField & mstrDelimiter
        End If
        strLine = strLine & strField
    Next i
    Print #intFileNum, strLine
    WriteLineNofij = True
    Exit Function
ErrHandler:
    WriteLineNofij = False
End Function

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

