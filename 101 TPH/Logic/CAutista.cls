VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' DB 01/09/2003 Aggiunto il campo numero documento di identità
Option Explicit


Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Private SQL_GET_SELECT_KEY As String
Private SQL_FIND_ORARIO_LAVORO As String
Private SQL_INSERT_ORARIO_LAVORO As String
Private SQL_UPDATE_ORARIO_LAVORO As String
Private SQL_UPDATE_ORARIO_LAVORO_2 As String
Private SQL_GET_LIST_STIPENDI As String
Private SQL_GET_LIST_STIPENDI_2 As String
Private SQL_GET_LIST_STIPENDI_3 As String
Private SQL_GET_LIST_STIPENDI_4 As String
Private SQL_GET_LIST_STIPENDI_5 As String

Private SQL_GET_LIST_AUTISTI_STIPENDI As String
Private SQL_GET_LIST_AUTISTI_STIPENDI_2 As String

Private SQL_GET_LIST_CONSUMI As String
Private SQL_GET_LIST_CONSUMI_2 As String
Private SQL_GET_LIST_CONSUMI_3 As String
Private SQL_GET_LIST_CONSUMI_4 As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)


' -- Invoca la funzione di controllo sulla possibilità di utilizzo in una data

Public Function CheckScadenze(varIDKEY As Variant, dtDateToCheck As Date, ByRef varResult As Variant) As Boolean

    On Error GoTo ErrHandler

 
    Dim SCASelect  As CScadenzeSelezione
    Set SCASelect = New CScadenzeSelezione

    With SCASelect.DS

        Set .DA = DS.DA

        Set .ErrH = DS.ErrH

    End With

    

    CheckScadenze = SCASelect.blnCheckScadenze(varIDKEY, dtDateToCheck, "AU", varResult)

 

    Set SCASelect = Nothing

    Exit Function

 

ErrHandler:

    Set SCASelect = Nothing

    CheckScadenze = False

End Function

Public Function FindByNrTelefono(NumeroTelefono As Variant) As Variant
    On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    strSQL = "SELECT AAUIDAAU " & _
             "  FROM ANAGAAU0 " & _
             " WHERE (AAUNOTE1 = ? OR AAUNOTE2 = ?) AND (AAUANNUL = 0)"
    
    If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, NumeroTelefono, NumeroTelefono) Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        FindByNrTelefono = Empty
    Else
        FindByNrTelefono = varBuffer(0, 0)
    End If
    
    Exit Function
ErrHandler:
    FindByNrTelefono = Empty
End Function

Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
    
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function

Public Function GetListStipendi(Data As Variant, ByRef ListStipendi As Variant, Optional IDBase As Variant, Optional ByRef ErrDesc As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim varBufferAutista As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim j As Long
    Dim strSQL As String
    Dim strIDAutista As String
    Dim strError As String
    
    If Not IsDate(Data) Then
        strError = "Data non valida"
        GoTo ErrHandler
    End If
    
    If Not IsMissing(IDBase) Then
        strSQL = SQL_GET_LIST_AUTISTI_STIPENDI
        If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, IDBase) Then
            GoTo ErrHandler
        End If
    Else
        strSQL = SQL_GET_LIST_AUTISTI_STIPENDI_2
        If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany) Then
            GoTo ErrHandler
        End If
        
'        strSQL = SQL_GET_LIST_STIPENDI & SQL_GET_LIST_STIPENDI_2 & DS.DA.MakeSQL(Data, dtDate) & SQL_GET_LIST_STIPENDI_4 & SQL_GET_LIST_STIPENDI_5
'        If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany) Then
'            GoTo ErrHandler
'        End If
    End If
        
    
    strSQL = SQL_GET_LIST_STIPENDI
    For i = 0 To UBound(varBuffer, 2)
        lngHowMany = 0
        strIDAutista = varBuffer(0, i)
        If Not DS.DA.QueryCommand(strSQL, varBufferAutista, lngHowMany, strIDAutista, Data) Then
            GoTo ErrHandler
        End If
        If Not IsEmpty(varBufferAutista) Then
            For j = 2 To 9
                varBuffer(j, i) = varBufferAutista(j, 0)
            Next j
        End If
    Next i
    
'        strSQL = SQL_GET_LIST_STIPENDI & SQL_GET_LIST_STIPENDI_3 & DS.DA.MakeSQL(Data, dtDate) & SQL_GET_LIST_STIPENDI_4 & SQL_GET_LIST_STIPENDI_5
'        If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, IDBase) Then
'            GoTo ErrHandler
'        End If
    ListStipendi = varBuffer
    GetListStipendi = True
    Exit Function
ErrHandler:
    ListStipendi = Empty
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    GetListStipendi = False
End Function

Public Function GetListAutomezzi(Data As Variant, ByRef ListAutomezzi As Variant, IDAutista As Variant, Optional ByRef ErrDesc As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strSQL As String
    Dim strError As String
    
    If Not IsDate(Data) Then
        strError = "Data non valida"
        GoTo ErrHandler
    End If
    
    strSQL = SQL_GET_LIST_CONSUMI & SQL_GET_LIST_CONSUMI_2 & DS.DA.MakeSQL(Data, dtDate) & SQL_GET_LIST_CONSUMI_3 & SQL_GET_LIST_CONSUMI_4
    If Not DS.DA.QueryCommand(strSQL, varBuffer, lngHowMany, IDAutista) Then
        GoTo ErrHandler
    End If
    
    ListAutomezzi = varBuffer
    GetListAutomezzi = True
    Exit Function
ErrHandler:
    ListAutomezzi = Empty
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    GetListAutomezzi = False
End Function


Public Function InsertOrarioInizioLavoro(IDAutista As Variant, Data As Variant, OraInizioLavoro As Variant, DataInizioLavoro As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not IsDate(Data) Then
        GoTo ErrHandler
    End If
    If Not IsDate(OraInizioLavoro) Then
        GoTo ErrHandler
    End If
    
    DS.DA.BeginTransaction
    
    DS.Field("Codice").FindValue = IDAutista
    If Not Me.Find Then
        GoTo ErrHandler
    End If
    If DS.HowMany = 0 Then
        GoTo ErrHandler
    End If
    
    If Not DS.DA.QueryCommand(SQL_FIND_ORARIO_LAVORO & DS.DA.MakeSQL(Data, dtDate), varBuffer, lngHowMany, IDAutista) Then
        GoTo ErrHandler
    End If
    
    If varBuffer(0, 0) = 0 Then
        If Not DS.DA.ExecuteCommand(SQL_INSERT_ORARIO_LAVORO & DS.DA.MakeSQL(Data, dtDate) & ", " & DS.DA.MakeSQL(OraInizioLavoro, dtTime) & ", " & DS.DA.MakeSQL(DataInizioLavoro, dtDate) & ")", IDAutista) Then
            GoTo ErrHandler
        End If
    Else
        If Not DS.DA.ExecuteCommand(SQL_UPDATE_ORARIO_LAVORO & DS.DA.MakeSQL(OraInizioLavoro, dtTime) & ", OLVDTINL = " & DS.DA.MakeSQL(DataInizioLavoro, dtDate) & SQL_UPDATE_ORARIO_LAVORO_2, IDAutista, Data) Then
            GoTo ErrHandler
        End If
    End If
    
    DS.DA.EndTrans etCommit
    InsertOrarioInizioLavoro = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    InsertOrarioInizioLavoro = False
End Function

Public Function ComputeOraInizioLavoro(IDAutista As String, Data As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim varOraPrimoCarico As Variant
    Dim varMinutiDistanza As Variant
    Dim varOraInizioLavoro As Variant
    Dim varDataInizioLavoro As Variant
    Dim strIDVVT As String
    Dim strIDADE As String
    Dim strIDABS As String
    Dim lngHowMany As Long
    Dim i As Long
    
    Const SQL_GET_MIN_ORA_CARICO = "SELECT VVTIDVVT, VVTHRINC, VVTIDADE, AAUIDABS " & _
                                   "  FROM VIAGVVT0 INNER JOIN ANAGAAU0 ON VVTIDAU1 = AAUIDAAU " & _
                                   " WHERE VVTIDAU1 = ? AND VVTANNUL = 0 AND VVTDTINC = "
    Const SQL_GET_MIN_ORA_CARICO_ORDERBY = " ORDER BY VVTHRINC"
    Const SQL_GET_DISTANZA_BASE_DEPOSITO = "SELECT BXDNRMIN FROM ANAGBXD0 WHERE BXDIDABS = ? AND BXDIDADE = ? AND BXDANNUL = 0 "
    
    If Not DS.DA.QueryCommand(SQL_GET_MIN_ORA_CARICO & DS.DA.MakeSQL(Data, dtDate) & SQL_GET_MIN_ORA_CARICO_ORDERBY, varBuffer, lngHowMany, IDAutista) Then
        GoTo ErrHandler
    End If
    
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    
    varOraPrimoCarico = varBuffer(1, 0)
    strIDVVT = varBuffer(0, 0)
    strIDADE = varBuffer(2, 0)
    strIDABS = varBuffer(3, 0)
    
    lngHowMany = 0
    varBuffer = Empty
    If Not DS.DA.QueryCommand(SQL_GET_DISTANZA_BASE_DEPOSITO, varBuffer, lngHowMany, strIDABS, strIDADE) Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    
    varMinutiDistanza = varBuffer(0, 0)
    
    varOraInizioLavoro = Format(DateAdd("n", -varMinutiDistanza, varOraPrimoCarico), "hh:nn")
    If varOraInizioLavoro > Format(varOraPrimoCarico, "hh:nn") Then
        varDataInizioLavoro = DateAdd("d", -1, Data)
    Else
        varDataInizioLavoro = Data
    End If
    If Not InsertOrarioInizioLavoro(IDAutista, Data, varOraInizioLavoro, varDataInizioLavoro) Then
        GoTo ErrHandler
    End If
    
    ComputeOraInizioLavoro = True
    Exit Function
ErrHandler:
    ComputeOraInizioLavoro = False
End Function
Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Private Sub DS_ClientClear()
    RaiseEvent ClientClear
End Sub

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function


Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    BeginSave = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
End Function
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT aauidaau, aaunomea, aaunomab, aauidave, " & _
               "       aaunumpa, aaunrcid, aaufldis, aaunote1, " & _
               "       aaunote2, aauidabs, aauannul, timestmp FROM anagaau0 WHERE "
    SQL_GETLIST = "SELECT aauidaau, aaunomea FROM anagaau0 "
    SQL_UPDATE = "UPDATE anagaau0 SET "
    SQL_INSERT = "INSERT INTO anagaau0 "
    SQL_GET_TS = "SELECT timestmp FROM anagaau0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aauidaau, aaunomea, aauidave,aaufldis, aauannul FROM anagaau0 "
    
    SQL_FIND_ORARIO_LAVORO = "SELECT COUNT(*) FROM STIPOLV0 WHERE OLVIDAAU = ? AND OLVANNUL = 0 AND OLVDTOLV = "
    SQL_INSERT_ORARIO_LAVORO = "INSERT INTO STIPOLV0 (OLVIDAAU, OLVDTOLV, OLVHRINL, OLVDTINL) VALUES (?, "
    SQL_UPDATE_ORARIO_LAVORO = "UPDATE STIPOLV0 SET OLVHRINL = "
    SQL_UPDATE_ORARIO_LAVORO_2 = " WHERE OLVIDAAU = ? AND OLVDTOLV = ?"

'    SQL_GET_LIST_STIPENDI = "SELECT AAUIDAAU, AAUNOMEA, OLVDTINL, OLVHRINL, " & _
'                            "       OLVDTILR, OLVHRILR, OLVDTFIL, OLVHRFIL, " & _
'                            "       OLVNRKMS, AAUIDAVE " & _
'                            "  FROM ANAGAAU0 LEFT OUTER JOIN STIPOLV0 ON AAUIDAAU = OLVIDAAU "
'    SQL_GET_LIST_STIPENDI_2 = " WHERE ((OLVDTOLV = "
'    SQL_GET_LIST_STIPENDI_3 = " WHERE AAUIDABS = ? AND ((OLVDTOLV = "
'    SQL_GET_LIST_STIPENDI_4 = " ) OR (OLVDTOLV IS NULL)) "
'    SQL_GET_LIST_STIPENDI_5 = " ORDER BY AAUIDAAU"
    
    SQL_GET_LIST_STIPENDI = "SELECT '', '', OLVDTINL, OLVHRINL, " & _
                            "       OLVDTILR, OLVHRILR, OLVDTFIL, OLVHRFIL, " & _
                            "       OLVNRKMS, OLVFLMAS, '' " & _
                            "  FROM STIPOLV0 " & _
                            " WHERE OLVIDAAU = ? AND OLVDTOLV = ? "
    
    SQL_GET_LIST_AUTISTI_STIPENDI = "SELECT AAUIDAAU, AAUNOMEA, '', '', " & _
                                    "       '', '', '', '', " & _
                                    "       '', '', AAUIDAVE " & _
                                    "  FROM ANAGAAU0 " & _
                                    " WHERE AAUIDABS = ? AND AAUANNUL = 0 " & _
                                    " ORDER BY AAUIDAAU "
    SQL_GET_LIST_AUTISTI_STIPENDI_2 = "SELECT AAUIDAAU, AAUNOMEA, '', '', " & _
                                    "       '', '', '', '', " & _
                                    "       '', '', AAUIDAVE " & _
                                    "  FROM ANAGAAU0 " & _
                                    " WHERE AAUANNUL = 0 " & _
                                    " ORDER BY AAUIDAAU "
    
    SQL_GET_LIST_CONSUMI = "SELECT DISTINCT VVTIDAM1, CAMNRCIN, CAMNRCES " & _
                           "  FROM VIAGVVT0 LEFT OUTER JOIN STIPCAM0 ON VVTIDAM1 = CAMIDAAM AND VVTDTINV = CAMDTCAM "
    SQL_GET_LIST_CONSUMI_2 = " WHERE VVTIDAU1 = ? AND (VVTANNUL = 0) AND (VVTFLSTV <> 'P') AND ((VVTDTINV = "
    SQL_GET_LIST_CONSUMI_3 = " )) "
    SQL_GET_LIST_CONSUMI_4 = " ORDER BY VVTIDAM1"
    
    mvarFieldsIndexes = Array("0", "1", "3", "5", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CAutistaSoc
    
    With DS.Fields
        .Add "Codice", "anagaau0.aauidaau", 0, True, 3, "0"
        .Add "Nome", "anagaau0.aaunomea", 0, True, 30, "1"
        .Add "Nome Abbreviato", "anagaau0.aaunomab", 0, , 6, "2"
        .Add "Vettore", "anagaau0.aauidave", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Numero Patente", "anagaau0.aaunumpa", 0, , 15, "4"
        .Add "Numero NIF", "anagaau0.aaunrcid", 0, True, 20, "9"
        .Add "Disponibilità", "anagaau0.aaufldis", 0, True, 1, "5"
        .Add "Note 1", "anagaau0.aaunote1", 0, , 40, "6"
        .Add "Note 2", "anagaau0.aaunote2", 0, , 40, "7"
        .Add "Base", "anagaau0.aauidabs", 0, True, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "ABS"
            .ForeignTable = "anagabs0"
            .ForeignTableFld = "absidabs"
            .ForeignTableDescField = "absdescr"
            .ForeignTableDeleteFld = "absannul"
        End With
        .Add "Annullato", "anagaau0.aauannul", 3, , , "8"
    End With
    With DS.KeyFields
        .Add "0", "1"
'        .Add "7", "2"
'        .Add "6", "3"
    End With
    DS.CancelField = "anagaau0.aauannul"
End Sub

Public Function GetNewKey(ByRef strKey As String) As Boolean
On Error GoTo ErrHandler

    Dim AAU As CAutista
    Dim varCounter As Variant
    Dim lngKeyPrefix As Long

    Set AAU = New CAutista
    With AAU.DS
        Set .DA = DS.DA
        Set .ErrH = DS.DA.ErrH
    End With

    Do
        varCounter = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "AAU", 0)
        If varCounter = -1 Then
            GoTo ErrHandler
        End If
    
        lngKeyPrefix = CLng(Mid(varCounter, Len(varCounter) - 3, 2))
        If lngKeyPrefix <= 9 Then
            strKey = Chr(48 + lngKeyPrefix) & Right(varCounter, 2)
        Else
            strKey = Chr(65 + (lngKeyPrefix - 9)) & Right(varCounter, 2)
        End If
        
        With AAU
            .DS.Fields.ResetFields
            .DS.Fields("0").FindValue = strKey
            If Not .Find Then
                GoTo ErrHandler
            End If

            If .DS.HowMany = 0 Then
                Exit Do
            End If
        End With
    Loop

    Set AAU = Nothing
    GetNewKey = True
    Exit Function

ErrHandler:
    Set AAU = Nothing
    GetNewKey = False
End Function

Private Sub Class_Terminate()
    Set DS = Nothing
End Sub


Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub




