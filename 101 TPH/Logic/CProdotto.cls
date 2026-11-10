VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProdotto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()

Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
    
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function

Public Function GetDensity(Data As Variant) As Variant
    On Error GoTo ErrHandler
    
    If Not IsNumeric(DS.Fields("9").Value) Then
        GoTo ErrHandler
    End If
    
    GetDensity = DS.Fields("9").Value
    
    Exit Function
ErrHandler:
    GetDensity = Empty
End Function

Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT apridapr, aprdesco, aprdesfi, aprdesab, " & _
               "       apridarr, apridafp, aprdemin, aprdemax, " & _
               "       aprdecon, apridttp, aprannul, timestmp FROM anagapr0 WHERE "
'    SQL_GETLIST = "SELECT apridapr, aprdesco FROM anagapr0 "
    SQL_GETLIST = "SELECT apridapr, aprdesab FROM anagapr0 "
    SQL_UPDATE = "UPDATE anagapr0 SET "
    SQL_INSERT = "INSERT INTO anagapr0 "
    SQL_GET_TS = "SELECT timestmp FROM anagapr0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT apridapr, aprdesco, aprdesfi, aprannul FROM anagapr0 "
    
    mvarFieldsIndexes = Array("0", "1", "2", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CProdottoSoc
    
    With DS.Fields
        .Add "Codice", "anagapr0.apridapr", 0, True, 6, "0"
        .Add "Descrizione Commerciale", "anagapr0.aprdesco", 0, True, 25, "1"
        .Add "Descrizione Fiscale", "anagapr0.aprdesfi", 0, True, 35, "2"
        .Add "Descrizione Abbreviata", "anagapr0.aprdesab", 0, True, 6, "3"
        .Add "Raggruppamento Prodotto", "anagapr0.apridarr", 0, True, 3, "4"
        With .Item("4")
            .ForeignTableEntity = "ARR"
            .ForeignTable = "anagarr0"
            .ForeignTableFld = "arridarr"
            .ForeignTableDescField = "arrdescr"
            .ForeignTableDeleteFld = "arrannul"
        End With
        .Add "Raggruppamento Fiscale Prodotto", "anagapr0.apridafp", 0, True, 3, "5"
        With .Item("5")
            .ForeignTableEntity = "AFP"
            .ForeignTable = "anagafp0"
            .ForeignTableFld = "afpidafp"
            .ForeignTableDescField = "afpdescr"
            .ForeignTableDeleteFld = "afpannul"
        End With
        .Add "Densit‡ Minima", "anagapr0.aprdemin", 1, False, 6, "6"
        .Add "Densit‡ Massima", "anagapr0.aprdemax", 1, False, 6, "7"
        .Add "Densit‡ Contrattuale", "anagapr0.aprdecon", 1, False, 6, "9"
        .Add "Tipologia Prodotto", "anagapr0.apridttp", 0, True, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "TTP"
            .ForeignTable = "anagttp0"
            .ForeignTableFld = "ttpidttp"
            .ForeignTableDescField = "ttpdescr"
            .ForeignTableDeleteFld = "ttpannul"
        End With
        .Add "Annullato", "anagapr0.aprannul", 3, , , "8"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagapr0.aprannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


Private Sub DS_ClientClear()
    RaiseEvent ClientClear

End Sub

Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh

End Sub

Public Function CalcolaLitri(QtaKg As Variant, _
                             Densit‡Ambiente As Variant, _
                             Optional ByRef DescErrore As Variant) As Variant
    On Error GoTo ErrHandler
    
    If (Not IsNumeric(QtaKg)) Or (Not IsNumeric(Densit‡Ambiente)) Then
        GoTo ErrHandler
    End If
    
    If Not CheckDensit‡(0, Densit‡Ambiente, DescErrore) Then
        GoTo ErrHandler
    End If
    
    If Densit‡Ambiente <= 0 Then
        GoTo ErrHandler
    End If
    
    CalcolaLitri = CDbl(QtaKg) / CDbl(Format(Densit‡Ambiente, "0.00000"))
    
    Exit Function
ErrHandler:
    CalcolaLitri = Empty
End Function

Public Function CalcolaKg(QtaL As Variant, _
                          Optional Densit‡Ambiente As Variant, _
                          Optional ByRef DescErrore As Variant) As Variant
    On Error GoTo ErrHandler
    
    If (Not IsNumeric(QtaL)) Or (Not IsNumeric(Densit‡Ambiente)) Then
        GoTo ErrHandler
    End If
    
    If Not CheckDensit‡(0, Densit‡Ambiente, DescErrore) Then
        GoTo ErrHandler
    End If
    
'    If CDbl(Densit‡Ambiente) > DS.Field("Densit‡ Massima").Value Then
'        DescErrore = "La densit‡ del prodotto " & DS.Field("Descrizione Commerciale").Value & _
'                     " non puÚ essere superiore a " & DS.Field("Densit‡ Massima").Value
'        GoTo ErrHandler
'    End If
'
    CalcolaKg = Round(CDbl(QtaL) * CDbl(Format(Densit‡Ambiente, "0.00000")))
    
    Exit Function
ErrHandler:
    CalcolaKg = Empty
End Function

Public Function CheckDensit‡(Optional Den15 As Variant = 0, _
                             Optional DenAm As Variant = 0, _
                             Optional ByRef DescErrore As Variant) As Boolean
    On Error GoTo ErrHandler
    
    If (Not IsNumeric(Den15)) Or (Not IsNumeric(DenAm)) Then
        GoTo ErrHandler
    End If
    
    If Den15 = 0 And DenAm = 0 Then
        DescErrore = "Valore non consentito per la densit‡"
        GoTo ErrHandler
    End If
    
    If DenAm > 0 Then
        If CDbl(DenAm) < DS.Fields("6").Value Then
            DescErrore = "La densit‡ ambiente del prodotto " & DS.Fields("1").Value & _
                         " non puÚ essere inferiore a " & DS.Fields("6").Value
            GoTo ErrHandler
        End If
        If CDbl(DenAm) > DS.Fields("7").Value Then
            DescErrore = "La densit‡ ambiente del prodotto " & DS.Fields("1").Value & _
                         " non puÚ essere superiore a " & DS.Fields("7").Value
            GoTo ErrHandler
        End If
    End If
    
    If Den15 > 0 Then
        If CDbl(Den15) < DS.Fields("6").Value Then
            DescErrore = "La densit‡ a 15∞ del prodotto " & DS.Fields("1").Value & _
                         " non puÚ essere inferiore a " & DS.Fields("6").Value
            GoTo ErrHandler
        End If
        If CDbl(Den15) > DS.Fields("7").Value Then
            DescErrore = "La densit‡ a 15∞ del prodotto " & DS.Fields("1").Value & _
                         " non puÚ essere superiore a " & DS.Fields("7").Value
            GoTo ErrHandler
        End If
    End If
    
    CheckDensit‡ = True
    
    Exit Function
ErrHandler:
    CheckDensit‡ = False
End Function

