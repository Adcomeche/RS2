Attribute VB_Name = "OIL2BLG0_Global"
Public Function IsEmptyVariantString(varValue As Variant) As Variant
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(varValue) = 0 Then
        IsEmptyVariantString = True
    Else
        IsEmptyVariantString = False
    End If
End Function

