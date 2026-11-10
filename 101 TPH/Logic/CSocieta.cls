VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSocieta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const SQL_GETLIST = "SELECT ascidasc, ascrasoc FROM anagasc0 "
Private Const SQL_FIND = "SELECT ascidasc, ascrasoc, ascindir, asclocal, " & _
                         "       ascprovi, ascnrcap, asccofis, ascpaiva, " & _
                         "       ascidamp, ascidaiv, ascannul " & _
                         "  FROM anagasc0 WHERE ascidasc = "
                         
Private Const SQL_UPDATE = "UPDATE anagasc0 SET "
Private Const SQL_INSERT = "INSERT anagasc0 "
Private Const SQL_GET_TS = "SELECT timestmp FROM anagasc0 WHERE ascidasc = "

Private mobjDA As CDataAccess
Private mobjErrH As CErrorHandler
Private mvarList As Variant
Private mcolFields As ColFields
Private mcolKeyFields As Collection
Private mvarRecBuf As Variant
Private mvarHowMany As Long
Private TimeStamp As Variant

Public Property Get KeyFields() As Collection
    Set KeyFields = mcolKeyFields
End Property

Public Property Get HowMany() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HowMany
    HowMany = mvarHowMany
End Property

Public Property Set ErrH(NewValue As CErrorHandler)
    Set mobjErrH = NewValue
    Set mobjDA.ErrH = mobjErrH
End Property

Public Function Insert() As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim strFields As String
    Dim strValues As String
    Dim blnSave As Boolean
    Dim I As Integer
    
    strQ = SQL_INSERT & "("
    For I = 1 To mcolFields.Count
        With mcolFields(I)
            If .Updated Then
                strFields = strFields & .FieldName & ","
                strValues = strValues & mobjDA.MakeSQL(.NewValue, .FldType) & ","
                blnSave = True
            End If
        End With
    Next I
    
    If blnSave Then
        strQ = strQ & Left$(strFields, Len(strFields) - 1)
        strQ = strQ & ") VALUES ("
        strQ = strQ & Left$(strValues, Len(strValues) - 1)
        strQ = strQ & ")"
        If Not mobjDA.ExecuteQuery(strQ) Then
            GoTo ErrHandler
        End If
        For I = 1 To mcolFields.Count
            With mcolFields(I)
                If .Updated Then
                    .Updated = False
                    .Value = .NewValue
                End If
            End With
    Next I
    Else
    End If
    
    GoTo EndFunc
ErrHandler:
    Insert = False
    Exit Function
EndFunc:
    Insert = True

End Function

Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim Fld As CField
    Dim blnSave As Boolean
    Dim varNewTS As Variant
    
    strQ = SQL_UPDATE
    For Each Fld In mcolFields
        With Fld
            If .Updated Then
                strQ = strQ & .FieldName & " = " & mobjDA.MakeSQL(.NewValue, .FldType) & ","
                blnSave = True
            End If
        End With
    Next Fld
    If blnSave Then
        strQ = Left$(strQ, Len(strQ) - 1)
        strQ = strQ & " WHERE ascidasc = " & mobjDA.MakeSQL(mcolFields("Codice").Value, mcolFields("Codice").FldType)
        mobjDA.BeginTransaction
        varNewTS = GetTimeStamp
        
        If Not mobjDA.ExecuteQuery(strQ) Then
            mobjDA.Rollback
            GoTo ErrHandler
        End If
        If mobjDA.CompareTimeStamps(varNewTS, TimeStamp) Then
            blnConcurErr = False
        Else
            blnConcurErr = True
        End If
        
    Else
    End If
    
    GoTo EndProc
ErrHandler:
    If Err.Source <> "CDataAccess" Then
        Err.Raise UpdateFail, "CSocieta"
    End If
    Exit Sub
EndProc:
End Sub


Private Function GetTimeStamp() As Variant
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim varResult As Variant
    
    strQ = SQL_GET_TS & mobjDA.MakeSQL(mcolFields("Codice").Value, dtAlfa)
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

Public Sub EndSave()
    mcolFields.UpdateFields
End Sub

Public Sub EndTrans(ByVal TranFlag As EndTran)
    mobjDA.EndTrans TranFlag
End Sub

Public Property Get List() As Variant
    List = mvarList
End Property

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim lngRows As Long
    
    strQ = SQL_GETLIST
    If ActiveRecordsOnly Then
        strQ = strQ & " WHERE ascannul = " & mobjDA.MakeSQL(False, dtBit)
    End If
    
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

Public Property Set DA(NewValue As CDataAccess)
    Set mobjDA = NewValue
End Property

Public Property Get Fields() As ColFields
    Set Fields = mcolFields
End Property

Public Function Find(Code As String) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim lngNRows As Long
    Dim I As Integer
    
    strQ = SQL_FIND & mobjDA.MakeSQL(Code, dtAlfa)
    lngNRows = 1
    If mobjDA.QueryData(strQ, mvarRecBuf, lngNRows) Then
        If lngNRows > 0 Then
            For I = 0 To mcolFields.Count
                mcolFields(I + 1).Value = mvarRecBuf(I, 0)
            Next I
            TimeStamp = mvarRecBuf(I, 0)
        Else
        
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


Private Sub Class_Initialize()
    Set mcolFields = New ColFields
    Set mcolKeyFields = New Collection
    
    With mcolFields
        .Add "Codice", "anagasc0.ascidasc", 0, True, 3, "Codice"
        .Add "RagioneSociale", "anagasc0.ascrasoc", 0, True, 25, "RagioneSociale"
        .Add "Indirizzo", "anagasc0.ascindir", 0, , 40, "Indirizzo"
        .Add "Localita", "anagasc0.asclocal", 0, , 25, "Localita"
        .Add "Provincia", "anagasc0.ascprovi", 0, , 2, "Provincia"
        .Add "CAP", "anagasc0.ascnrcap", 0, , 10, "CAP"
        .Add "CodiceFiscale", "anagasc0.asccofis", 0, , 16, "CodiceFiscale"
        .Add "PartitaIVA", "anagasc0.ascpaiva", 0, , 20, "PartitaIVA"
        .Add "IDPagamento", "anagasc0.ascidamp", 0, True, 3, "IDPagamento"
        .Add "IDIVA", "anagasc0.ascidaiv", 0, True, 3, "IDIVA"
        .Add "Annullato", "anagasc0.ascannul", 3, , , "Annullato"
    End With
    With mcolKeyFields
        .Add "Codice", "1"
    End With
End Sub


Private Sub Class_Terminate()
    Set mcolFields = Nothing
    Set mcolKeyFields = Nothing
End Sub


