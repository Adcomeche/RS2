VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolControls As New Collection

Public Function Add(ByRef Obj As Object, ByRef F As Form, Optional ByVal Key As String = "") As CControl
    On Error GoTo ErrHandler
    
    Dim clsControl As New CControl
    
    Select Case UCase(TypeName(Obj))
        Case "COMMANDBUTTON"
            Set clsControl.CommandBtn = Obj
        Case "COMBOBOX"
            Set clsControl.cmbComboBox = Obj
        Case "MASKEDBOX"
            Set clsControl.mskMaskEdBox = Obj
        Case "TDBGRID"
            Set clsControl.DBGrid = Obj
        Case "LABEL"
            Set clsControl.Lbl = Obj
        Case Else
            GoTo ErrHandler
    End Select
    Set clsControl.Frm = F
    If Key = vbNullString Then
        mcolControls.Add clsControl
    Else
        mcolControls.Add clsControl, Key
    End If
    GoTo EndProc
ErrHandler:
    Set Add = Nothing
    Exit Function
EndProc:
    Set Add = clsControl

End Function


Private Sub Class_Terminate()
    Set mcolControls = Nothing
End Sub


