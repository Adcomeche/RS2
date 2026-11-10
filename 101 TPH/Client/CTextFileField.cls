VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextFileField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Campo di un record di Input"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarName As String 'local copy
Private mvarSize As Integer 'local copy
Private mvarValue As Variant 'local copy
Private mvarDefault As Variant 'local copy
Private mvarType As TextFileFieldType

Public Enum TextFileFieldType
    fftAlfa = 0
    fftNumerico = 1
End Enum

Public Property Get FieldType() As TextFileFieldType
    FieldType = mvarType
End Property

Public Property Let FieldType(NewValue As TextFileFieldType)
    mvarType = NewValue
End Property

Public Property Let Default(ByVal vData As Variant)
Attribute Default.VB_Description = "Valore di default per il campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Default = 5
    mvarDefault = vData
End Property




Public Property Get Default() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Default
    Default = mvarDefault
End Property





Public Property Let Value(ByVal vData As Variant)
Attribute Value.VB_Description = "Valore contenuto nel campo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property




Public Property Get Value() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property



Public Property Let Size(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Size = 5
    mvarSize = vData
End Property


Public Property Get Size() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Size
    Size = mvarSize
End Property



Public Property Let Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = vData
End Property


Public Property Get Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
End Property



