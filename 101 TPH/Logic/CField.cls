VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarName As String 'local copy
Private mvarValue As Variant 'local copy
Private mvarNewValue As Variant 'local copy
Private mvarUpdated As Boolean 'local copy
Private mvarFieldName As String 'local copy
Private mvarType As Byte 'local copy
Private mblnRequired As Boolean ' local copy
Private mintMaxLength As Integer ' local copy
'Private mvarFindValue As Variant 'local copy
Private mcolFindValue As Collection 'local copy
Private mvarForeignTable As String 'local copy
Private mvarForeignTableFld As String 'local copy
Private mvarForeignTableDeleteFld As String 'local copy
Private mvarForeignTableDescField As String 'local copy
Private mvarForeignTableDescValue As Variant 'local copy
Private mvarForeignTableFldValue As Variant 'local copy
Private mvarForeignTableEntity As String 'local copy
Private mvarOperator As String 'local copy
Private mvarFieldHwnd As Long 'local copy
Public AllowedValues As Collection 'local copy

Public Enum ValidateError
    veNoError = 0
    veRequired = 1
    veLength = 2
    veFormat = 3
    veRange = 4
    veType = 5
End Enum
'local variable(s) to hold property value(s)
Private mvarReadOnly As Boolean 'local copy
Public Property Let ReadOnly(ByVal vData As Boolean)
Attribute ReadOnly.VB_Description = "Indica se il campo può subire modifiche o meno"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReadOnly = 5
    mvarReadOnly = vData
End Property


Public Property Get ReadOnly() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadOnly
    ReadOnly = mvarReadOnly
End Property




Public Property Let FieldHwnd(ByVal vData As Long)
Attribute FieldHwnd.VB_Description = "Handle del controllo associato al campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldHwnd = 5
    mvarFieldHwnd = vData
End Property


Public Property Get FieldHwnd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldHwnd
    FieldHwnd = mvarFieldHwnd
End Property




Public Property Get FindValuesCount() As Integer
    FindValuesCount = mcolFindValue.Count
End Property

Public Property Let Operator(ByVal vData As String)
Attribute Operator.VB_Description = "Operatore da usare nei confronti operati sul campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Operator = 5
    mvarOperator = vData
End Property


Public Property Get Operator() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Operator
    Operator = mvarOperator
End Property





Public Property Let ForeignTableEntity(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTableEntity = 5
    mvarForeignTableEntity = vData
End Property


Public Property Get ForeignTableEntity() As String
Attribute ForeignTableEntity.VB_Description = "Codice dell'entità collegata"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTableEntity
    ForeignTableEntity = mvarForeignTableEntity
End Property



Public Property Let ForeignTableFldValue(ByVal vData As Variant)
Attribute ForeignTableFldValue.VB_Description = "Valore della chiave della tabella collegata (ForeignTable). Viene usato dopo una ricerca per descrizione nella tabella collegata."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTableFldValue = 5
    mvarForeignTableFldValue = vData
End Property




Public Property Get ForeignTableFldValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTableFldValue
    ForeignTableFldValue = mvarForeignTableFldValue
End Property



Public Property Let ForeignTableDescValue(ByVal vData As Variant)
Attribute ForeignTableDescValue.VB_Description = "Valore del campo decodificato nella tabella collegata (ForeignTable)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTableDescValue = 5
    mvarForeignTableDescValue = vData
End Property




Public Property Get ForeignTableDescValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTableDescValue
    ForeignTableDescValue = mvarForeignTableDescValue
End Property



Public Property Let ForeignTableDescField(ByVal vData As String)
Attribute ForeignTableDescField.VB_Description = "Campo descrizione nella tabella collegata (Foreign Field)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTableDescField = 5
    mvarForeignTableDescField = vData
End Property


Public Property Get ForeignTableDescField() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTableDescField
    ForeignTableDescField = mvarForeignTableDescField
End Property



Public Property Let ForeignTableDeleteFld(ByVal vData As String)
Attribute ForeignTableDeleteFld.VB_Description = "Campo di annullamento della tabella collegata (ForeignTable)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTableDeleteFld = 5
    mvarForeignTableDeleteFld = vData
End Property


Public Property Get ForeignTableDeleteFld() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTableDeleteFld
    ForeignTableDeleteFld = mvarForeignTableDeleteFld
End Property




Public Property Let ForeignTableFld(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTableFld = 5
    mvarForeignTableFld = vData
End Property


Public Property Get ForeignTableFld() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTableFld
    ForeignTableFld = mvarForeignTableFld
End Property



Public Property Let ForeignTable(ByVal vData As String)
Attribute ForeignTable.VB_Description = "Nome della tabella collegata attraverso il campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeignTable = 5
    mvarForeignTable = vData
End Property


Public Property Get ForeignTable() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeignTable
    ForeignTable = mvarForeignTable
End Property









Public Property Let FindValue(Optional Index As Integer = 1, ByVal vData As Variant)
Attribute FindValue.VB_Description = "Valore da utilizzare nelle ricerche per il campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FindValue = 5
    If Index = 1 Then
'        mcolFindValue.Clear
        Set mcolFindValue = New Collection
    End If
    mcolFindValue.Add vData
    'mvarFindValue = vData
End Property




Public Property Get FindValue(Optional Index As Integer = 1) As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FindValue
    If mcolFindValue.Count > 0 Then
        FindValue = mcolFindValue(Index)
    Else
        FindValue = Empty
    End If
End Property



Public Property Let FldType(ByVal vData As Byte)
Attribute FldType.VB_Description = "Tipo del campo: 0=Alfanumerico ; 1 = Numerico ; 2 = Data/Ora ; 3 = Bit"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type = 5
    mvarType = vData
End Property


Public Property Get FldType() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Type
    FldType = mvarType
End Property



Public Property Let FieldName(ByVal vData As String)
Attribute FieldName.VB_Description = "Nome fisico del campo nel database"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property



Public Property Let MaxLength(NewValue As Integer)
    mintMaxLength = NewValue
End Property

Public Property Get MaxLength() As Integer
    MaxLength = mintMaxLength
End Property

Public Property Let Required(NewValue As Boolean)
    mblnRequired = NewValue
End Property

Public Property Get Required() As Boolean
    Required = mblnRequired
End Property

Public Property Let Updated(ByVal vData As Boolean)
Attribute Updated.VB_Description = "Segnala che in NewValue è presente un nuovo valore da attribuire al campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Updated = 5
    mvarUpdated = vData
End Property


Public Property Get Updated() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Updated
    Updated = mvarUpdated
End Property



Public Property Let NewValue(ByVal vData As Variant)
Attribute NewValue.VB_Description = "Buffer per il nuovo valore da attribuire al campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewValue = 5
    mvarNewValue = vData
End Property




Public Property Get NewValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewValue
    If IsNull(mvarNewValue) Then
        NewValue = vbNullString
    Else
        NewValue = mvarNewValue
    End If
End Property



Public Function Validate() As ValidateError
    Dim Result As ValidateError
    
    Result = veNoError
    If Updated Then
        If mblnRequired Then
            If (IsEmpty(Me.NewValue)) Or (Me.NewValue = vbNullString) Then
                Result = veRequired
                GoTo EndProc
            End If
        End If
        If mintMaxLength > 0 Then
            If Len(Me.NewValue) > mintMaxLength Then
                Result = veLength
                GoTo EndProc
            End If
        End If
        Select Case mvarType
            Case 0  'Alfanumerico
            Case 1  'Numerico
                If Len(Me.NewValue) > 0 Then
                    If Not IsNumeric(Me.NewValue) Then
                        Result = veType
                        GoTo EndProc
                    End If
                End If
            Case 2  'Data
                If (Me.NewValue <> vbNullString) Then
                    If Not IsDate(Me.NewValue) Then
                        Result = veType
                        GoTo EndProc
                    End If
                End If
            Case 3  'Bit
            Case 4  'Ora
                If (Me.NewValue <> vbNullString) Then
                    If Not IsDate(Me.NewValue) Then
                        Result = veType
                        GoTo EndProc
                    End If
                End If
        End Select
    End If
EndProc:
    Validate = Result
End Function


Public Property Let Value(ByVal vData As Variant)
Attribute Value.VB_Description = "Valore attuale del campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property



Public Property Get Value() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property



Public Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = vData
End Property


Public Property Get Name() As String
Attribute Name.VB_Description = "Descrizione del campo"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
End Property



Private Sub Class_Initialize()
    Set AllowedValues = New Collection
    Set mcolFindValue = New Collection
End Sub


Private Sub Class_Terminate()
    Set AllowedValues = Nothing
    Set mcolFindValue = Nothing
End Sub


