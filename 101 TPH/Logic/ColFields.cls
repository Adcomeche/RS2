VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Raccolta di campi che compongono un'entità"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CField"
Attribute VB_Ext_KEY = "Member0" ,"CField"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mcolValidationErrors As Collection

Public Function AddValidationError(ErrorMsg As String) As Boolean
    On Error GoTo ErrHandler
    mcolValidationErrors.Add ErrorMsg
        
    AddValidationError = True
    Exit Function
ErrHandler:
    AddValidationError = False
End Function


Public Sub ResetFields(Optional UpdatedState As Boolean = False)
    Dim Fld As CField
    
    For Each Fld In mCol
        With Fld
            .Updated = UpdatedState
            .Value = Empty
            .NewValue = Empty
            .FindValue = Empty
        End With
    Next Fld
    Set mcolValidationErrors = Nothing
End Sub

Public Function Add(Name As String, _
                    FieldName As String, _
                    FldType As Byte, _
                    Optional Required = False, _
                    Optional MaxLength = 0, _
                    Optional sKey As String, _
                    Optional Operator As String = " = ", _
                    Optional ReadOnly As Boolean = False) As CField
    
    'create a new object
    Dim objNewMember As CField
    Set objNewMember = New CField

    'set the properties passed into the method
    With objNewMember
        .Name = Name
        .FieldName = FieldName
        .FldType = FldType
        .Required = Required
        .MaxLength = MaxLength
        .Operator = Operator
        .ReadOnly = ReadOnly
    End With
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function
Public Function Validate() As Boolean
    Dim blnDataOK As Boolean
    Dim i As Integer
    Dim ValidError As ValidateError
    Dim strErrMsg As String
    
    blnDataOK = True
    Set mcolValidationErrors = New Collection
    For i = 1 To mCol.Count
        ValidError = mCol(i).Validate
        If ValidError <> veNoError Then
            blnDataOK = False
            strErrMsg = "Il campo " & mCol(i).Name
            Select Case ValidError
                Case veRequired
                    strErrMsg = strErrMsg & " non può essere vuoto"
                Case veLength
                    strErrMsg = strErrMsg & " non può essere più lungo di " & mCol(i).MaxLength & " caratteri"
                Case veFormat
                    strErrMsg = strErrMsg & " "
                Case veRange
                    strErrMsg = strErrMsg & " "
                Case veType
                    strErrMsg = strErrMsg & " contiene un dato non valido"
            End Select
            mcolValidationErrors.Add strErrMsg
        End If
    Next i
    Validate = blnDataOK
End Function

Public Property Get Item(vntIndexKey As Variant) As CField
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


Public Sub UpdateFields()
    Dim Fld As CField
    
    For Each Fld In mCol
        With Fld
            If .Updated Then
                .Updated = False
                .Value = .NewValue
            End If
        End With
    Next Fld

End Sub


Public Property Get ValidationErrors() As Collection
    Set ValidationErrors = mcolValidationErrors
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

