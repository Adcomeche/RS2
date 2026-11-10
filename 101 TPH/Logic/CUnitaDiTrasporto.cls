VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUnitaDiTrasporto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_FIND_BY_BADGE As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_EXPORT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_EXPORT_VEICOLI As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)


' -- Invoca la funzione di controllo sulla possibilità di utilizzo in una data
Public Function CheckScadenze(varIDKEY As Variant, dtDateToCheck As Date, varIDASC As Variant, ByRef varResult As Variant) As Boolean

    On Error GoTo ErrHandler

 

    Dim SCASelect  As CScadenzeSelezione
    Set SCASelect = New CScadenzeSelezione
    With SCASelect.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    CheckScadenze = SCASelect.blnCheckScadenze(varIDKEY, dtDateToCheck, "UT", varResult, varIDASC)
    Set SCASelect = Nothing

    Exit Function

 

ErrHandler:

    Set SCASelect = Nothing

    CheckScadenze = False

End Function

Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
    
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function

Public Function GetExportData(ByRef ExportData As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    strQ = SQL_EXPORT
    blnResult = DS.DA.QueryCommand(strQ, ExportData, lngHowMany, DS.Fields("0").Value)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    GetExportData = True
    Exit Function
ErrHandler:
    GetExportData = False
    ExportData = Empty
    
End Function

Public Function GetIDByBadge(Badge As String) As String
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    blnResult = DS.DA.QueryCommand(SQL_FIND_BY_BADGE, varBuffer, lngHowMany, Badge)
    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If
    GetIDByBadge = varBuffer(0, 0)
    Exit Function
ErrHandler:
    GetIDByBadge = vbNullString
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


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    BeginSave = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
End Function
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT autidaut, autdescr, autidam1, autidam2, " & _
                      "autidau1, autidau2, autidav1, autidav2, autporkg, " & _
                      "autfldis, autidade, autidaee, autidasc, autflsoc, " & _
                      "autbadge, autidacn, autpresi, autannul, timestmp FROM anagaut0 WHERE "

    SQL_FIND_BY_BADGE = "SELECT autidaut FROM anagaut0 WHERE autbadge = ? AND autannul = 0"
    SQL_GETLIST = "SELECT autidaut, autdescr FROM anagaut0 "
    SQL_UPDATE = "UPDATE anagaut0 SET "
    SQL_INSERT = "INSERT INTO anagaut0 "
    SQL_GET_TS = "SELECT timestmp FROM anagaut0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT autidaut, autdescr, autpresi, autidade, " & _
                          "       autidaee, autidam1, autidam2, autidau1, autidau2, " & _
                          "       autidav1, autidav2 FROM anagaut0 "
    SQL_EXPORT = "SELECT ACNNUMCO, ACNNUMSE, ACNNOTEC, ACNKMCON, ACNIDAVE, AUTIDASC, " & _
                 "       AUTIDAEE, AVERASOC, AVEINDIR, AVELOCAL, AVEPROVI, AVENRCAP, " & _
                 "       AVENRTEL, AUTIDAM1, " & _
                 "       ANAGAAM0.AAMIDATM, ANAGAAM0.AAMIDATS, ANAGAAM0.AAMPORKG, " & _
                 "       ANAGAAM0.AAMTARAP, ANAGAAM0.AAMTARAL, ANAGAAM0.AAMPORRI, " & _
                 "       ANAGAAM0.AAMCUBCO, ANAGAAM0.AAMCAP01, ANAGAAM0.AAMNUM01, " & _
                 "       ANAGAAM0.AAMCAP02, ANAGAAM0.AAMNUM02, ANAGAAM0.AAMCAP03, " & _
                 "       ANAGAAM0.AAMNUM03, ANAGAAM0.AAMCAP04, ANAGAAM0.AAMNUM04, " & _
                 "       ANAGAAM0.AAMCAP05, ANAGAAM0.AAMNUM05, ANAGAAM0.AAMCAP06, " & _
                 "       ANAGAAM0.AAMNUM06, ANAGAAM0.AAMCAP07, ANAGAAM0.AAMNUM07, " & _
                 "       ANAGAAM0.AAMCAP08, ANAGAAM0.AAMNUM08, ANAGAAM0.AAMCAP09, " & _
                 "       ANAGAAM0.AAMNUM09, ANAGAAM0.AAMCAP10, ANAGAAM0.AAMNUM10, " & _
                 "       ANAGAAM0.AAMCARIC, ANAGAUT0.AUTIDAM2, "
   SQL_EXPORT = SQL_EXPORT & _
                 "       ANAGAAM0_1.AAMIDATM, ANAGAAM0_1.AAMIDATS, ANAGAAM0_1.AAMPORKG, " & _
                 "       ANAGAAM0_1.AAMTARAP, ANAGAAM0_1.AAMTARAL, ANAGAAM0_1.AAMPORRI, " & _
                 "       ANAGAAM0_1.AAMCUBCO, ANAGAAM0_1.AAMCAP01, ANAGAAM0_1.AAMNUM01, " & _
                 "       ANAGAAM0_1.AAMCAP02, ANAGAAM0_1.AAMNUM02, ANAGAAM0_1.AAMCAP03, " & _
                 "       ANAGAAM0_1.AAMNUM03, ANAGAAM0_1.AAMCAP04, ANAGAAM0_1.AAMNUM04, " & _
                 "       ANAGAAM0_1.AAMCAP05, ANAGAAM0_1.AAMNUM05, ANAGAAM0_1.AAMCAP06, " & _
                 "       ANAGAAM0_1.AAMNUM06, ANAGAAM0_1.AAMCAP07, ANAGAAM0_1.AAMNUM07, " & _
                 "       ANAGAAM0_1.AAMCAP08, ANAGAAM0_1.AAMNUM08, ANAGAAM0_1.AAMCAP09, " & _
                 "       ANAGAAM0_1.AAMNUM09, ANAGAAM0_1.AAMCAP10, ANAGAAM0_1.AAMNUM10, " & _
                 "       ANAGAAM0_1.AAMCARIC, ANAGAUT0.AUTIDAU1, ANAGAAU0.AAUNOMEA, ANAGAAU0.AAUIDAVE, " & _
                 "       ANAGAUT0.AUTIDAU2, ANAGAAU0_1.AAUNOMEA, ANAGAAU0_1.AAUIDAVE, ANAGACN0.ACNIDATN "
    SQL_EXPORT = SQL_EXPORT & _
                 "  FROM ANAGAAM0 RIGHT JOIN (((((ANAGAUT0 LEFT JOIN ANAGAAM0 AS ANAGAAM0_1 " & _
                 "       ON ANAGAUT0.AUTIDAM2 = ANAGAAM0_1.AAMIDAAM) LEFT JOIN ANAGAAU0 ON " & _
                 "       ANAGAUT0.AUTIDAU1 = ANAGAAU0.AAUIDAAU) LEFT JOIN ANAGAAU0 AS ANAGAAU0_1 ON " & _
                 "       ANAGAUT0.AUTIDAU2 = ANAGAAU0_1.AAUIDAAU) LEFT JOIN ANAGACN0 ON " & _
                 "       ANAGAUT0.AUTIDACN = ANAGACN0.ACNIDACN) LEFT JOIN ANAGAVE0 ON " & _
                 "       ANAGACN0.ACNIDAVE = ANAGAVE0.AVEIDAVE) ON ANAGAAM0.AAMIDAAM = ANAGAUT0.AUTIDAM1" & _
                 " WHERE (((ANAGAUT0.AUTIDAUT)=?))"
    
    SQL_EXPORT_VEICOLI = " SELECT ANAGAUT0.AUTIDAUT, ANAGAAM0.AAMIDAAM, ANAGAAM0_1.AAMIDAAM, ANAGAUT0.AUTIDADE, ANAGAUT0.AUTIDAEE, " & _
                         " ANAGAEE0.AEEDESCR, (ANAGAAM0.AAMPORKG+ANAGAAM0_1.AAMPORKG) AS PESOTOT, (anagaam0.aamcubco+ANAGAAM0_1.aamcubco) AS VOLTOTALE, " & _
                         " anagaam0.AAMFLSCA,ANAGAAM0_1.AAMFLSCA, ANAGAUT0.AUTIDACN, ANAGAUT0.AUTIDAU1, ANAGAAU0.AAUNOMEA, ANAGAUT0.AUTIDAU2, ANAGAAU0_1.AAUNOMEA, ANAGAAM0.AAMPORKG AS PESOMOTRICE, " & _
                         " ANAGAAM0.aamcubco AS VOLUMEMOTRICE, ANAGAAM0_1.AAMPORKG AS PESORIMORCHIO, ANAGAAM0_1.aamcubco AS VOLUMERIMORCHIO, "
                         
    SQL_EXPORT_VEICOLI = SQL_EXPORT_VEICOLI & " ANAGAAM0.AAMCAP01,ANAGAAM0.AAMNUM01, " & _
                         " ANAGAAM0.AAMCAP02,ANAGAAM0.AAMNUM02, ANAGAAM0.AAMCAP03,ANAGAAM0.AAMNUM03, " & _
                         " ANAGAAM0.AAMCAP04,ANAGAAM0.AAMNUM04,ANAGAAM0.AAMCAP05,ANAGAAM0.AAMNUM05,ANAGAAM0.AAMCAP06,ANAGAAM0.AAMNUM06, " & _
                         " ANAGAAM0.AAMCAP07,ANAGAAM0.AAMNUM07,ANAGAAM0.AAMCAP08,ANAGAAM0.AAMNUM08, " & _
                         " ANAGAAM0.AAMCAP09,ANAGAAM0.AAMNUM09,ANAGAAM0.AAMCAP10,ANAGAAM0.AAMNUM10,  "
                         
    SQL_EXPORT_VEICOLI = SQL_EXPORT_VEICOLI & " ANAGAAM0_1.AAMCAP01,ANAGAAM0_1.AAMNUM01,ANAGAAM0_1.AAMCAP02,ANAGAAM0_1.AAMNUM02,ANAGAAM0_1.AAMCAP03,ANAGAAM0_1.AAMNUM03, " & _
                         " ANAGAAM0_1.AAMCAP04,ANAGAAM0_1.AAMNUM04,ANAGAAM0_1.AAMCAP05,ANAGAAM0_1.AAMNUM05,ANAGAAM0_1.AAMCAP06,ANAGAAM0_1.AAMNUM06, " & _
                         " ANAGAAM0_1.AAMCAP07,ANAGAAM0_1.AAMNUM07,ANAGAAM0_1.AAMCAP08,ANAGAAM0_1.AAMNUM08,ANAGAAM0_1.AAMCAP09,ANAGAAM0_1.AAMNUM09, " & _
                         " ANAGAAM0_1.AAMCAP10,ANAGAAM0_1.AAMNUM10, ANAGAAM0.AAMTARAP, ANAGAAM0_1.AAMTARAP, AUTIDAV1, AUTIDAV2 "
                         
    SQL_EXPORT_VEICOLI = SQL_EXPORT_VEICOLI & " FROM ANAGAEE0 INNER JOIN (((ANAGAAM0 AS ANAGAAM0_1 RIGHT JOIN (ANAGAAM0 RIGHT JOIN ANAGAUT0 ON ANAGAAM0.AAMIDAAM = ANAGAUT0.AUTIDAM1) ON ANAGAAM0_1.AAMIDAAM = ANAGAUT0.AUTIDAM2) LEFT JOIN ANAGAAU0 ON ANAGAUT0.AUTIDAU1 = ANAGAAU0.AAUIDAAU) " & _
                         " LEFT JOIN ANAGAAU0 AS ANAGAAU0_1 ON ANAGAUT0.AUTIDAU2 = ANAGAAU0_1.AAUIDAAU) ON ANAGAEE0.AEEIDAEE = ANAGAUT0.AUTIDAEE " & _
                         " WHERE ANAGAUT0.AUTPRESI=? AND ANAGAUT0.AUTANNUL = 0  "
    
    
    mvarFieldsIndexes = Array("0", "1", "17", "10", "11", "2", "3", "4", "5", "6", "7")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CUnitaDiTrasportoSoc
    
    With DS.Fields
        .Add "Codice", "anagaut0.autidaut", 0, True, 3, "0"
        .Add "Descrizione", "anagaut0.autdescr", 0, True, 30, "1"
        .Add "Motrice", "anagaut0.autidam1", 0, False, 10, "2"
        With .Item("2")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Rimorchio", "anagaut0.autidam2", 0, False, 10, "3"
        With .Item("3")
            .ForeignTableEntity = "AAM"
            .ForeignTable = "anagaam0"
            .ForeignTableFld = "aamidaam"
            .ForeignTableDescField = "aamdescr"
            .ForeignTableDeleteFld = "aamannul"
        End With
        .Add "Autista 1", "anagaut0.autidau1", 0, False, 3, "4"
        With .Item("4")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Autista 2", "anagaut0.autidau2", 0, False, 3, "5"
        With .Item("5")
            .ForeignTableEntity = "AAU"
            .ForeignTable = "anagaau0"
            .ForeignTableFld = "aauidaau"
            .ForeignTableDescField = "aaunomea"
            .ForeignTableDeleteFld = "aauannul"
        End With
        .Add "Vettore", "anagaut0.autidav1", 0, True, 3, "6"
        With .Item("6")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Vettore Trasportatore", "anagaut0.autidav2", 0, True, 3, "7"
        With .Item("7")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Portata Totale", "anagaut0.autporkg", 1, True, 6, "8"
        .Add "Disponibilità", "anagaut0.autfldis", 0, True, 1, "9"
        .Add "Deposito", "anagaut0.autidade", 0, True, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Ente Emittente", "anagaut0.autidaee", 0, True, 3, "11"
        With .Item("11")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Società", "anagaut0.autidasc", 0, True, 3, "12"
        With .Item("12")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Sociale/Terzo", "anagaut0.autflsoc", 0, True, 1, "13"
        .Add "Badge", "anagaut0.autbadge", 0, False, 3, "14"
        .Add "Contratto", "anagaut0.autidacn", 0, False, 3, "15"
        With .Item("15")
            .ForeignTableEntity = "ACN"
            .ForeignTable = "anagacn0"
            .ForeignTableFld = "acnidacn"
            .ForeignTableDescField = "acndescr"
            .ForeignTableDeleteFld = "acnannul"
        End With
        .Add "Presidio", "anagaut0.autpresi", 0, False, 5, "17"
        .Add "Annullato", "anagaut0.autannul", 3, , , "16"
'        .Add "Portata Lt", "autporlt", 1, , , "15", , True
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagaut0.autannul"
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
Public Function GetExportDataVeicoli(strPresidio As String, ByRef ExportDataVeicoli As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    strQ = SQL_EXPORT_VEICOLI
    blnResult = DS.DA.QueryCommand(strQ, ExportDataVeicoli, lngHowMany, strPresidio)
    
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If lngHowMany <= 0 Then
         GoTo ErrHandler
    End If
    GetExportDataVeicoli = True
    Exit Function
    
ErrHandler:

    GetExportDataVeicoli = False
    ExportDataVeicoli = Empty
    
End Function

Public Function GetNewKey(ByRef strKey As String) As Boolean
On Error GoTo ErrHandler

    Dim AUT As CUnitaDiTrasporto
    Dim varCounter As Variant
    Dim lngKeyPrefix As Long

    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.DA.ErrH
    End With

    Do
        varCounter = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "AUT", 0)
        If varCounter = -1 Then
            GoTo ErrHandler
        End If
    
        lngKeyPrefix = CLng(Mid(varCounter, Len(varCounter) - 3, 2))
        If lngKeyPrefix <= 9 Then
            strKey = Chr(48 + lngKeyPrefix) & Right(varCounter, 2)
        Else
            strKey = Chr(65 + (lngKeyPrefix - 9)) & Right(varCounter, 2)
        End If
        
        With AUT
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

    Set AUT = Nothing
    GetNewKey = True
    Exit Function

ErrHandler:
    Set AUT = Nothing
    GetNewKey = False
End Function

