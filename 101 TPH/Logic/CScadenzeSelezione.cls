VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScadenzeSelezione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private MultEST As CEstintoriAutomezzo
Private AXX As CParametro

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Private strFLDIG As String

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private Const SQL_GET_SELECTION = "SELECT ANAGSCA0.SCAIDENT, ANAGAVV.AVVDESCR, ANAGSCA0.SCAIDATA, ANAGATA0.ATADESCR, " & _
                                "           ANAGSCA0.SCADTINV, ANAGSCA0.SCADTFIV, ANAGSCA0.SCAANNUL, ANAGATA0.ATAFLENT, " & _
                                "           ANAGAAM0.AAMIDAVE " & _
                                "  FROM ANAGATA0 INNER JOIN ANAGAVV ON " & _
                                "           ANAGATA0.ATAFLENT = ANAGAVV.AVVVALUE AND " & _
                                "           'ATAFLENT' = ANAGAVV.AVVIDFLD " & _
                                "  INNER JOIN ANAGSCA0 ON " & _
                                "           ANAGATA0.ATAIDATA = ANAGSCA0.SCAIDATA " & _
                                "  LEFT OUTER JOIN ANAGAAM0 ON " & _
                                "           ANAGATA0.ATAIDATA = ANAGAAM0.AAMIDAAM "
                                
Private Const SQL_GET_TIPI_SCADENZA_PER_ENTITA = "SELECT ANAGATA0.ATAIDATA, ANAGATA0.ATAFLVIN, ANAGATA0.ATADESCR " & _
                                " FROM ANAGATA0 WHERE ANAGATA0.ATAFLENT = ? AND ANAGATA0.ATAFLVIN <> 0 AND " & _
                                "           ANAGATA0.ATAANNUL = 0 " & _
                                " ORDER BY ANAGATA0.ATAFLVIN DESC "

Private Const SQL_CHECK_SCADENZE = "SELECT ANAGATA0.ATADESCR, ANAGSCA0.SCADTINV, ANAGSCA0.SCADTFIV " & _
                                "   FROM ANAGSCA0 INNER JOIN ANAGATA0 ON " & _
                                "           ANAGSCA0.SCAIDATA = ANAGATA0.ATAIDATA " & _
                                "   WHERE ANAGSCA0.SCAIDENT = ? AND ANAGSCA0.SCAIDATA = ? AND ANAGSCA0.SCAANNUL = 0 " & _
                                "   ORDER BY ANAGATA0.ATAFLVIN DESC"
                                
Private Const SQL_GET_AAU_AAM_PER_AUT = "SELECT AUTIDAU1, AUTIDAU2, AUTIDAM1, AUTIDAM2 FROM ANAGAUT0 WHERE AUTIDAUT = ?"

Private Const SQL_GET_ASU_PER_AAM = "SELECT ASUIDASU FROM ANAGASU0 WHERE ASUIDAAM = ? AND ASUANNUL = 0"

Private Const SQL_GET_ACC_PER_AAM = "SELECT ACCPROGR FROM ANAGACC0 WHERE ACCIDAAM = ? AND ACCANNUL = 0"

Private Const SQL_CHECK_ANT_ATM = "SELECT ANAGAAM0.AAMANNOI, ANAGANT0.ANTNRANN " & _
                            "  FROM ANAGANT0 INNER JOIN ANAGAAM0 ON " & _
                            "       ANAGANT0.ANTIDATM = ANAGAAM0.AAMIDATM " & _
                            "  WHERE ANAGAAM0.AAMIDAAM = ? AND ANAGANT0.ANTIDASC = ? AND ANAGANT0.ANTANNUL = 0"
 
Private Const SQL_CHECK_ANT_ATS = "SELECT ANAGAAM0.AAMANNOI, ANAGANT0.ANTNRANN " & _
                            "  FROM ANAGANT0 INNER JOIN ANAGAAM0 ON " & _
                            "       ANAGANT0.ANTIDATS = ANAGAAM0.AAMIDATS " & _
                            "  WHERE ANAGAAM0.AAMIDAAM = ? AND ANAGANT0.ANTIDASC = ? AND ANAGANT0.ANTANNUL = 0"
                            
Private Const SQL_GET_DATA_AAM = "SELECT ANAGATM0.ATMFLTI2, ANAGAAM0.AAMIDATS, ANAGAAM0.AAMFLSCA, ANAGAMT0.AMTFLDIG " & _
                                " FROM ANAGAAM0 INNER JOIN ANAGATM0 ON " & _
                                "       ANAGAAM0.AAMIDATM = ANAGATM0.ATMIDATM " & _
                                " LEFT OUTER JOIN ANAGAMT0 ON " & _
                                "       ANAGAAM0.AAMIDAMT = ANAGAMT0.AMTIDAMT " & _
                                " WHERE ANAGAAM0.AAMIDAAM = ? AND ANAGAAM0.AAMANNUL = 0"
                                
Private Const SQL_GET_AVERASOC = "SELECT AVERASOC FROM ANAGAVE0 WHERE AVEIDAVE = ?"

' -- Restituisce True se un autista o un automezzo o una u.d.t. è cessato
Public Function blnIsCessato(varIDENT As Variant, varFLENT As Variant) As Boolean
    Dim strSQL As String
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    Select Case CStr(varFLENT)
        Case "AM"
            strSQL = "SELECT 1 FROM ANAGAAM0 WHERE AAMIDAAM = ? AND AAMFLCES = 0"
        
        Case "AU"
            strSQL = "SELECT 1 FROM ANAGAAU0 WHERE AAUIDAAU = ? AND AAUFLCES = 0"
        
        Case "UT"
            strSQL = "SELECT 1 FROM ANAGAUT0 WHERE AUTIDAUT = ? AND AUTFLCES = 0"
        
        Case "AS"
            strSQL = "SELECT 1 FROM ANAGASU0 INNER JOIN ANAGAAM0 ON " & _
                    "       ANAGASU0.ASUIDAAM = ANAGAAM0.AAMIDAAM " & _
                    " WHERE ANAGASU0.ASUIDASU = ? AND ANAGAAM0.AAMFLCES = 0"
        
        Case "CL"
            strSQL = "SELECT 1 FROM ANAGAAM0 INNER JOIN ANAGACC0 ON " & _
                    "       ANAGAAM0.AAMIDAAM = ANAGACC0.ACCIDAAM " & _
                    " WHERE ANAGACC0.ACCPROGR = ? AND ANAGAAM0.AAMFLCES = 0"
                    
        Case "ES"
            strSQL = "SELECT 1 FROM MULTEAM0 INNER JOIN ANAGAAM0 ON " & _
                    "       MULTEAM0.EAMIDAAM = ANAGAAM0.AAMIDAAM " & _
                    " WHERE MULTEAM0.EAMIDEST = ? AND ANAGAAM0.AAMFLCES = 0"
                    
    End Select
    
    If DS.DA.QueryCommand(strSQL, varBuffer, 0, varIDENT) Then
        If Not IsEmpty(varBuffer) Then
            GoTo ErrHandler
        End If
    Else
        GoTo ErrHandler
    End If
    
    blnIsCessato = True
    Exit Function
    
ErrHandler:
    blnIsCessato = False
End Function

Public Function GetSelection(ByRef varBufResults As Variant, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional intTipoAutomezzi As Integer, _
                             Optional strIDAVE As String, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    Dim i As Long
    Dim c As Integer
    Dim blnDoneFirst As Boolean
    Dim blnOk As Boolean
    
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, , MaxRecords)
    End If
    
    If DS.HowMany > 0 Then
        ReDim varBufResults(8, 0)
    
        ' -- Depuro la DS.List dai records cessati e dai tipi veicoli non richiesti
        For i = 0 To DS.HowMany - 1
            If Not blnIsCessato(DS.List(0, i), DS.List(7, i)) Then
                If CStr(DS.List(7, i)) = "AM" And Not IsNull(DS.List(8, i)) Then
                    If DS.List(8, i) = strIDAVE Or strIDAVE = vbNullString Then
                        If intTipoAutomezzi <> 0 Then
                            If intTipoAutomezzi = 1 Then
                                ' -- Solo PROPRIOS
                                If Not IsNull(DS.List(8, i)) Then
                                    If CStr(DS.List(8, i)) = "A01" Or CStr(DS.List(8, i)) = "A02" Then
                                        blnOk = True
                                    Else
                                        blnOk = False
                                    End If
                                End If
                            ElseIf intTipoAutomezzi = 2 Then
                                ' -- Solo SUBCONTRATADOS
                                If Not IsNull(DS.List(8, i)) Then
                                    If CStr(DS.List(8, i)) <> "A01" And CStr(DS.List(8, i)) <> "A02" Then
                                        blnOk = True
                                    Else
                                        blnOk = False
                                    End If
                                End If
                            End If
                        Else
                            blnOk = True
                        End If
                    Else
                        blnOk = False
                    End If
                Else
                    blnOk = True
                End If
                
                If blnOk Then
                    If blnDoneFirst Then
                        ReDim Preserve varBufResults(8, 0 To UBound(varBufResults, 2) + 1)
                    End If
                    blnDoneFirst = True
                    
                    For c = 0 To 8
                        If c <> 8 Then
                            varBufResults(c, UBound(varBufResults, 2)) = DS.List(c, i)
                        Else
                            If Not IsNull(DS.List(c, i)) Then
                                varBufResults(c, UBound(varBufResults, 2)) = strGetAVERASOC(DS.List(c, i))
                            Else
                                varBufResults(c, UBound(varBufResults, 2)) = DS.List(c, i)
                            End If
                        End If
                    Next c
                End If
            End If
        Next i
        
        If Not blnDoneFirst Then
            varBufResults = Empty
        End If
    End If
End Function

' -- Restituisce la ragione sociale di una società
Private Function strGetAVERASOC(varIDASC As Variant) As String
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    If DS.DA.QueryCommand(SQL_GET_AVERASOC, varBuffer, 0, CVar(varIDASC)) Then
        If Not IsEmpty(varBuffer) Then
            strGetAVERASOC = varBuffer(0, 0)
        Else
            GoTo ErrHandler
        End If
    Else
        GoTo ErrHandler
    End If

    Exit Function

ErrHandler:
    strGetAVERASOC = vbNullString
End Function

' -- Controlla la possibilità di utilizzo di un'autista o un automezzo o una U.D.T. in una data
'       Parametri:
'       varIDKEY        : Chiave Primaria dell'entità da controllare
'       dtDateToCheck   : Data per la quale controllare la possibilità di utilizzo
'       strFLENT        : Codice dell'entità (AU = Autisti; AM = Automezzi; UT = U.D.T.)
'       varIDASC        : Codice Società (necessario per controllo ANTIGUEDADES su Automezzi)

' -- Risultato:
'    Viene restituito un array di 6 elementi:
'       0: Integer  - Esito: 0 = Possibilità di utilizzo; 1 = Avviso non bloccante; 2 = Impossibilità di utilizzo; 3 = Impossibilità di utilizzo causa ANTIGUEDAD
'       1: String   - Scadenza rilevata in caso l'elemento 0 fosse diverso da 0
'       2: Date     - Data di scadenza
'       3: String   - Chiave primaria dell'entità sulla quale è stata rilevata la scadenza (es. AAMIDAAM)
'       4: String   - Entità sulla quale è stata rilevata la scadenza ("AU": Autista, "ES": Estintore, "CL": Contalitri, "AS": Assicurazione, "AM": Automezzo, "UT": U.d.T.)
'       5: String   - Chiave tipologia di scadenza rilevata
Public Function blnCheckScadenze(varIDKEY As Variant, dtDateToCheck As Date, strFLENT As String, ByRef varResult As Variant, Optional varIDASC As Variant) As Boolean
    Dim ID_PARAM As String
    Dim strCheckScadenze As String
    
    Dim varBufIDATA As Variant
    Dim iATA As Integer
    
    Dim varBuffer As Variant
    Dim i As Integer
    
    Dim varBufAAM As Variant
    Dim strFLTI2 As String
    Dim strIDATS As String
    Dim strFLSCA As String
    
    Dim blnResultAutista1 As Boolean
    Dim varResultAutista1 As Variant
    
    Dim blnResultAutista2 As Boolean
    Dim varResultAutista2 As Variant
    
    Dim blnResultMotrice As Boolean
    Dim varResultMotrice As Variant
    
    Dim blnResultRimorchio As Boolean
    Dim varResultRimorchio As Variant
    
    Dim blnResultEstintore As Boolean
    Dim varResultEstintore As Variant
    
    Dim blnResultAssicurazione As Boolean
    Dim varResultAssicurazione As Variant
    
    Dim blnResultContalitri As Boolean
    Dim varResultContalitri As Variant
    
    Dim intAnniDiff As Integer
    
    On Error GoTo ErrHandler
    
    ' -- Controllo in ANAGAXX0 se il controllo sulle scadenze è abilitato
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    ID_PARAM = "001"
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If AXX.Find Then
            strCheckScadenze = Trim(.Field("Controllo Scadenze").Value)
        End If
    End With
    
    Set AXX = Nothing

    ReDim varResult(5)
    
    If Trim(strCheckScadenze) = "S" Then
        strFLDIG = "N"
        
        Select Case strFLENT
            Case "AU", "ES", "CL", "AS"
                ' -- Recupero l'elenco delle tipologie di scadenze per l'entità
                If DS.DA.QueryCommand(SQL_GET_TIPI_SCADENZA_PER_ENTITA, varBufIDATA, 0, strFLENT) Then
                    If Not IsEmpty(varBufIDATA) Then
                        For iATA = 0 To UBound(varBufIDATA, 2)
                            If CStr(varBufIDATA(0, iATA)) <> "013" Or (CStr(varBufIDATA(0, iATA)) = "013" And strFLDIG = "S") Then
                                ' -- Ricerco la data di scadenza per la tipologia di scadenza corrente
                                If DS.DA.QueryCommand(SQL_CHECK_SCADENZE, varBuffer, 0, varIDKEY, CVar(varBufIDATA(0, iATA))) Then
                                    If Not IsEmpty(varBuffer) Then
                                        If CDate(varBuffer(2, 0)) <= dtDateToCheck Then
                                            ' -- Scadenza, esco segnalando il livello di vincolo
                                            varResult(0) = CInt(varBufIDATA(1, iATA))
                                            varResult(1) = CStr(varBufIDATA(2, iATA))
                                            varResult(2) = CDate(varBuffer(2, 0))
                                            varResult(3) = CStr(varIDKEY)
                                            varResult(4) = strFLENT
                                            varResult(5) = CStr(varBufIDATA(0, iATA))
                                            blnCheckScadenze = True
                                            Exit Function
                                        End If
                                    Else
                                        ' -- Scadenza non impostata, esco segnalando il livello di vincolo
                                        varResult(0) = CInt(varBufIDATA(1, iATA))
                                        varResult(1) = CStr(varBufIDATA(2, iATA))
                                        varResult(2) = CDate("01/01/1900")
                                        varResult(3) = CStr(varIDKEY)
                                        varResult(4) = strFLENT
                                        varResult(5) = CStr(varBufIDATA(0, iATA))
                                        blnCheckScadenze = True
                                        Exit Function
                                    End If
                                Else
                                    GoTo ErrHandler
                                End If
                            End If
                        Next iATA
                        
                        varResult(0) = 0
                    Else
                        varResult(0) = 0
                    End If
                Else
                    GoTo ErrHandler
                End If
            
            Case "AM"
                ' -- Recupero ATMFLTI2, AAMIDATS, AAMFLSCA e AMTFLDIG
                ' le scadenze TACOGRAFO (002) e TARJETA DE TRANSPORTE (006) devono essere impostate solo per le Tractoras e i Rigidos (ATMFLTI2 = 'T' Or ATMFLTI2 = 'R')
                ' la scadenza CALIBRACION CISTERNA (007) deve essere impostata solo per automezzi le Cisternas e i Rigidos (ATMFLTI2 = 'C' Or ATMFLTI2 = 'R') e con cisterna FUEL/IFO (AAMIDATS = "007")
                ' la scadenza CALIBRACION CONTADOR (008) deve essere impostata solo per automezzi con contalitri AAMFLSCA = "S"
                ' la scadenza PRUEBA HIDRAULICA ADR (012) deve essere impostata solo per le Cisternas e i Rigidos (ATMFLTI2 = 'C' Or ATMFLTI2 = 'R')
                ' se l'automezzo ha il tachigrafo digitale (AMTFLDIG = 'S') per l'autista si deve controllare la scadenza TARJETA TAKOGRAFO DIGITAL (013)
                If DS.DA.QueryCommand(SQL_GET_DATA_AAM, varBufAAM, 0, varIDKEY) Then
                    If Not IsEmpty(varBufAAM) Then
                        strFLTI2 = CStr(varBufAAM(0, 0))
                        
                        If Not IsNull(varBufAAM(1, 0)) Then
                            strIDATS = CStr(varBufAAM(1, 0))
                        Else
                            strIDATS = ""
                        End If
                                            
                        If Not IsNull(varBufAAM(2, 0)) Then
                            strFLSCA = CStr(varBufAAM(2, 0))
                        Else
                            strFLSCA = "S"
                        End If
                        
                        If Not IsNull(varBufAAM(3, 0)) And strFLDIG <> "S" Then
                            strFLDIG = varBufAAM(3, 0)
                        End If
                    Else
                        GoTo ErrHandler
                    End If
                Else
                    GoTo ErrHandler
                End If
            
                ' -- Recupero l'elenco delle scadenze per l'entità
                If DS.DA.QueryCommand(SQL_GET_TIPI_SCADENZA_PER_ENTITA, varBufIDATA, 0, strFLENT) Then
                    If Not IsEmpty(varBufIDATA) Then
                        For iATA = 0 To UBound(varBufIDATA, 2)
                            If (CStr(varBufIDATA(0, iATA)) <> "002" And CStr(varBufIDATA(0, iATA)) <> "006" And CStr(varBufIDATA(0, iATA)) <> "007" And CStr(varBufIDATA(0, iATA)) <> "008" And CStr(varBufIDATA(0, iATA)) <> "012") Or _
                                ((CStr(varBufIDATA(0, iATA)) = "002" Or CStr(varBufIDATA(0, iATA)) = "006") And (strFLTI2 = "T" Or strFLTI2 = "R")) Or _
                                (CStr(varBufIDATA(0, iATA)) = "007" And (strFLTI2 = "C" Or strFLTI2 = "R") And strIDATS = "007") Or _
                                (CStr(varBufIDATA(0, iATA)) = "008" And strFLSCA = "S") Or _
                                (CStr(varBufIDATA(0, iATA)) = "012" And (strFLTI2 = "C" Or strFLTI2 = "R")) Then
                                ' -- Ricerco la data di scadenza per la tipologia di scadenza corrente
                                If DS.DA.QueryCommand(SQL_CHECK_SCADENZE, varBuffer, 0, varIDKEY, CVar(varBufIDATA(0, iATA))) Then
                                    If Not IsEmpty(varBuffer) Then
                                        If CDate(varBuffer(2, 0)) <= dtDateToCheck Then
                                            ' -- Scadenza, esco segnalando il livello di vincolo
                                            varResult(0) = CInt(varBufIDATA(1, iATA))
                                            varResult(1) = CStr(varBufIDATA(2, iATA))
                                            varResult(2) = CDate(varBuffer(2, 0))
                                            varResult(3) = CStr(varIDKEY)
                                            varResult(4) = strFLENT
                                            varResult(5) = CStr(varBufIDATA(0, iATA))
                                            blnCheckScadenze = True
                                            Exit Function
                                        End If
                                    Else
                                        ' -- Scadenza non impostata, esco segnalando il livello di vincolo
                                        varResult(0) = CInt(varBufIDATA(1, iATA))
                                        varResult(1) = CStr(varBufIDATA(2, iATA))
                                        varResult(2) = CDate("01/01/1900")
                                        varResult(3) = CStr(varIDKEY)
                                        varResult(4) = strFLENT
                                        varResult(5) = CStr(varBufIDATA(0, iATA))
                                        blnCheckScadenze = True
                                        Exit Function
                                    End If
                                Else
                                    GoTo ErrHandler
                                End If
                            End If
                        Next iATA
                    Else
                        varResult(0) = 0
                    End If
                Else
                    GoTo ErrHandler
                End If
                        
                ' -- Non ci sono vincoli da segnalare per l'automezzo, controllo quindi i relativi estintori, contalitri ed assicurazione
                ' -- Estintori
                Set MultEST = New CEstintoriAutomezzo
                With MultEST.DS
                    Set .DA = DS.DA
                    Set .ErrH = DS.ErrH
                End With
                
                varBuffer = Empty
                If MultEST.blnGetEstintori(varBuffer, varIDKEY) Then
                    If Not IsEmpty(varBuffer) Then
                        For i = 0 To UBound(varBuffer, 2)
                            blnResultEstintore = blnCheckScadenze(varBuffer(0, i), dtDateToCheck, "ES", varResultEstintore)
                        
                            If blnResultEstintore Then
                                If varResultEstintore(0) > 0 Then
                                    varResult(0) = varResultEstintore(0)
                                    varResult(1) = varResultEstintore(1)
                                    varResult(2) = varResultEstintore(2)
                                    varResult(3) = varResultEstintore(3)
                                    varResult(4) = varResultEstintore(4)
                                    varResult(5) = varResultEstintore(5)
                                    blnCheckScadenze = True
                                    Exit Function
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        Next i
                    Else
                        varResult(0) = 2
                        varResult(1) = "FALTA EXTINTOR"
                        varResult(2) = CDate("01/01/1900")
                        varResult(3) = CStr(varIDKEY)
                        varResult(4) = "AM"
                        varResult(5) = ""
                        blnCheckScadenze = True
                        Exit Function
                    End If
                Else
                    GoTo ErrHandler
                End If
                
                Set MultEST = Nothing
                
                
                ' -- Assicurazione
                varBuffer = Empty
                If DS.DA.QueryCommand(SQL_GET_ASU_PER_AAM, varBuffer, 0, varIDKEY) Then
                    If Not IsEmpty(varBuffer) Then
                        For i = 0 To UBound(varBuffer, 2)
                            blnResultAssicurazione = blnCheckScadenze(varBuffer(0, i), dtDateToCheck, "AS", varResultAssicurazione)
                        
                            If blnResultAssicurazione Then
                                If varResultAssicurazione(0) > 0 Then
                                    varResult(0) = varResultAssicurazione(0)
                                    varResult(1) = varResultAssicurazione(1)
                                    varResult(2) = varResultAssicurazione(2)
                                    varResult(3) = varResultAssicurazione(3)
                                    varResult(4) = varResultAssicurazione(4)
                                    varResult(5) = varResultAssicurazione(5)
                                    blnCheckScadenze = True
                                    Exit Function
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        Next
                    Else
                        varResult(0) = 2
                        varResult(1) = "FALTA SEGURO"
                        varResult(2) = CDate("01/01/1900")
                        varResult(3) = CStr(varIDKEY)
                        varResult(4) = "AM"
                        varResult(5) = ""
                        blnCheckScadenze = True
                        Exit Function
                    End If
                Else
                    GoTo ErrHandler
                End If
                
                
    ' -- La scadenza dei contalitri è stata spostata sugli automezzi
    '            If strFLSCA = "S" Then
    '                ' -- Contalitri
    '                varBuffer = Empty
    '                If DS.DA.QueryCommand(SQL_GET_ACC_PER_AAM, varBuffer, 0, varIDKEY) Then
    '                    If Not IsEmpty(varBuffer) Then
    '                        For i = 0 To UBound(varBuffer, 2)
    '                            blnResultContalitri = blnCheckScadenze(varBuffer(0, i), dtDateToCheck, "CL", varResultContalitri)
    '
    '                            If blnResultContalitri Then
    '                                If varResultContalitri(0) > 0 Then
    '                                    varResult(0) = varResultContalitri(0)
    '                                    varResult(1) = varResultContalitri(1)
    '                                    varResult(2) = varResultContalitri(2)
    '                                    varResult(3) = varResultContalitri(3)
    '                                    varResult(4) = varResultContalitri(4)
    '                                    varResult(5) = varResultContalitri(5)
    '                                    blnCheckScadenze = True
    '                                    Exit Function
    '                                End If
    '                            Else
    '                                GoTo ErrHandler
    '                            End If
    '                        Next
    '                    Else
    '                        varResult(0) = 2
    '                        varResult(1) = "FALTA CONTADOR"
    '                        varResult(2) = CDate("01/01/1900")
    '                        varResult(3) = CStr(varIDKEY)
    '                        varResult(4) = "AM"
    '                        varResult(5) = ""
    '                        blnCheckScadenze = True
    '                        Exit Function
    '                    End If
    '                Else
    '                    GoTo ErrHandler
    '                End If
    '            End If
                
                
                ' -- Controllo ANTIGUEDAD
                If Not IsMissing(varIDASC) Then
                    ' -- Controllo per "Tipo Automezzo"
                    varBuffer = Empty
                    If DS.DA.QueryCommand(SQL_CHECK_ANT_ATM, varBuffer, 0, varIDKEY, varIDASC) Then
                        If Not IsEmpty(varBuffer) Then
                            If IsDate(varBuffer(0, 0)) Then
                                intAnniDiff = DateDiff("yyyy", varBuffer(0, 0), Date)
                                
                                If intAnniDiff >= varBuffer(1, 0) Then
                                    If intAnniDiff > varBuffer(1, 0) Then
                                        ' -- Impossibilità di utilizzo causa ANTIGUEDAD
                                        varResult(0) = 3
                                        varResult(1) = "ERROR ANTIGUEDAD"
                                        varResult(2) = CDate("01/01/1900")
                                        varResult(3) = CStr(varIDKEY)
                                        varResult(4) = "AM"
                                        varResult(5) = ""
                                        blnCheckScadenze = True
                                        Exit Function
                                    Else
                                        If CDate(Left(varBuffer(0, 0), 6) & Year(Date)) <= Date Then
                                            ' -- Impossibilità di utilizzo causa ANTIGUEDAD
                                            varResult(0) = 3
                                            varResult(1) = "ERROR ANTIGUEDAD"
                                            varResult(2) = CDate("01/01/1900")
                                            varResult(3) = CStr(varIDKEY)
                                            varResult(4) = "AM"
                                            varResult(5) = ""
                                            blnCheckScadenze = True
                                            Exit Function
                                        Else
                                            ' -- Non ci sono vincoli da segnalare nemmeno per ANTIGUEDAD di questo automezzo
                                            varResult(0) = 0
                                        End If
                                    End If
                                Else
                                    ' -- Non ci sono vincoli da segnalare nemmeno per ANTIGUEDAD di questo automezzo
                                    varResult(0) = 0
                                End If
                            Else
                                ' -- Impossibilità di utilizzo causa ANTIGUEDAD
                                varResult(0) = 3
                                varResult(1) = "ERROR ANTIGUEDAD"
                                varResult(2) = CDate("01/01/1900")
                                varResult(3) = CStr(varIDKEY)
                                varResult(4) = "AM"
                                varResult(5) = ""
                                blnCheckScadenze = True
                                Exit Function
                            End If
                        End If
                    Else
                        varResult(0) = 3
                        varResult(1) = "ERROR ANTIGUEDAD"
                        varResult(2) = CDate("01/01/1900")
                        varResult(3) = CStr(varIDKEY)
                        varResult(4) = "AM"
                        varResult(5) = ""
                        blnCheckScadenze = True
                        Exit Function
                    End If
                    
                    ' -- Controllo per "Tipo Cisterna"
                    varBuffer = Empty
                    If DS.DA.QueryCommand(SQL_CHECK_ANT_ATS, varBuffer, 0, varIDKEY, varIDASC) Then
                        If Not IsEmpty(varBuffer) Then
                            If IsDate(varBuffer(0, 0)) Then
                                intAnniDiff = DateDiff("yyyy", varBuffer(0, 0), Date)
                                
                                If intAnniDiff >= varBuffer(1, 0) Then
                                    If intAnniDiff > varBuffer(1, 0) Then
                                        ' -- Impossibilità di utilizzo causa ANTIGUEDAD
                                        varResult(0) = 3
                                        varResult(1) = "ERROR ANTIGUEDAD"
                                        varResult(2) = CDate("01/01/1900")
                                        varResult(3) = CStr(varIDKEY)
                                        varResult(4) = "AM"
                                        varResult(5) = ""
                                        blnCheckScadenze = True
                                        Exit Function
                                    Else
                                        If Left(varBuffer(0, 0), 6) & Year(Date) >= Date Then
                                            ' -- Impossibilità di utilizzo causa ANTIGUEDAD
                                            varResult(0) = 3
                                            varResult(1) = "ERROR ANTIGUEDAD"
                                            varResult(2) = CDate("01/01/1900")
                                            varResult(3) = CStr(varIDKEY)
                                            varResult(4) = "AM"
                                            varResult(5) = ""
                                            blnCheckScadenze = True
                                            Exit Function
                                        Else
                                            ' -- Non ci sono vincoli da segnalare nemmeno per ANTIGUEDAD di questo automezzo
                                            varResult(0) = 0
                                        End If
                                    End If
                                Else
                                    ' -- Non ci sono vincoli da segnalare nemmeno per ANTIGUEDAD di questo automezzo
                                    varResult(0) = 0
                                End If
                            Else
                                ' -- Impossibilità di utilizzo causa ANTIGUEDAD
                                varResult(0) = 3
                                varResult(1) = "ERROR ANTIGUEDAD"
                                varResult(2) = CDate("01/01/1900")
                                varResult(3) = CStr(varIDKEY)
                                varResult(4) = "AM"
                                varResult(5) = ""
                                blnCheckScadenze = True
                                Exit Function
                            End If
                        End If
                    Else
                        varResult(0) = 3
                        varResult(1) = "ERROR ANTIGUEDAD"
                        varResult(2) = CDate("01/01/1900")
                        varResult(3) = CStr(varIDKEY)
                        varResult(4) = "AM"
                        varResult(5) = ""
                        blnCheckScadenze = True
                        Exit Function
                    End If
                Else
                    varResult(0) = 3
                    varResult(1) = "ERROR ANTIGUEDAD"
                    varResult(2) = CDate("01/01/1900")
                    varResult(3) = CStr(varIDKEY)
                    varResult(4) = "AM"
                    varResult(5) = ""
                    blnCheckScadenze = True
                    Exit Function
                End If
                
                varResult(0) = 0
            
            Case "UT"
                ' -- Recupero l'elenco delle scadenze per l'entità
                If DS.DA.QueryCommand(SQL_GET_TIPI_SCADENZA_PER_ENTITA, varBufIDATA, 0, strFLENT) Then
                    If Not IsEmpty(varBufIDATA) Then
                        For iATA = 0 To UBound(varBufIDATA, 2)
                            ' -- Ricerco la data di scadenza per la tipologia di scadenza corrente
                            If DS.DA.QueryCommand(SQL_CHECK_SCADENZE, varBuffer, 0, varIDKEY, CVar(varBufIDATA(0, iATA))) Then
                                If Not IsEmpty(varBuffer) Then
                                    If CDate(varBuffer(2, 0)) <= dtDateToCheck Then
                                        ' -- Scadenza, esco segnalando il livello di vincolo
                                        varResult(0) = CInt(varBufIDATA(1, iATA))
                                        varResult(1) = CStr(varBufIDATA(2, iATA))
                                        varResult(2) = CDate(varBuffer(2, 0))
                                        varResult(3) = CStr(varIDKEY)
                                        varResult(4) = strFLENT
                                        varResult(5) = CStr(varBufIDATA(0, iATA))
                                    End If
                                Else
                                    ' -- Scadenza non impostata, esco segnalando il livello di vincolo
                                    varResult(0) = CInt(varBufIDATA(1, iATA))
                                    varResult(1) = CStr(varBufIDATA(2, iATA))
                                    varResult(2) = CDate("01/01/1900")
                                    varResult(3) = CStr(varIDKEY)
                                    varResult(4) = strFLENT
                                    varResult(5) = CStr(varBufIDATA(0, iATA))
                                    Exit For
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        Next iATA
                    End If
                Else
                    GoTo ErrHandler
                End If
                        
                ' -- Non ci sono vincoli da segnalare per la U.D.T., controllo quindi i relativi autisti ed automezzi
                varBuffer = Empty
                If DS.DA.QueryCommand(SQL_GET_AAU_AAM_PER_AUT, varBuffer, 0, varIDKEY) Then
                    If Not IsEmpty(varBuffer) Then
                        ' -- Motrice
                        If Not IsNull(varBuffer(2, 0)) Then
                            blnResultMotrice = blnCheckScadenze(varBuffer(2, 0), dtDateToCheck, "AM", varResultMotrice, varIDASC)
                        
                            If blnResultMotrice Then
                                If varResultMotrice(0) > 0 Then
                                    varResult(0) = varResultMotrice(0)
                                    varResult(1) = varResultMotrice(1)
                                    varResult(2) = varResultMotrice(2)
                                    varResult(3) = varResultMotrice(3)
                                    varResult(4) = varResultMotrice(4)
                                    varResult(5) = varResultMotrice(5)
                                    blnCheckScadenze = True
                                    Exit Function
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        End If
                        
                        ' -- Rimorchio
                        If Not IsNull(varBuffer(3, 0)) Then
                            blnResultRimorchio = blnCheckScadenze(varBuffer(3, 0), dtDateToCheck, "AM", varResultRimorchio, varIDASC)
                        
                            If blnResultRimorchio Then
                                If varResultRimorchio(0) > 0 Then
                                    varResult(0) = varResultRimorchio(0)
                                    varResult(1) = varResultRimorchio(1)
                                    varResult(2) = varResultRimorchio(2)
                                    varResult(3) = varResultRimorchio(3)
                                    varResult(4) = varResultRimorchio(4)
                                    varResult(5) = varResultRimorchio(5)
                                    blnCheckScadenze = True
                                    Exit Function
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        End If
                        
                        ' -- Autista 1
                        If Not IsNull(varBuffer(0, 0)) Then
                            blnResultAutista1 = blnCheckScadenze(varBuffer(0, 0), dtDateToCheck, "AU", varResultAutista1)
                        
                            If blnResultAutista1 Then
                                If varResultAutista1(0) > 0 Then
                                    varResult(0) = varResultAutista1(0)
                                    varResult(1) = varResultAutista1(1)
                                    varResult(2) = varResultAutista1(2)
                                    varResult(3) = varResultAutista1(3)
                                    varResult(4) = varResultAutista1(4)
                                    varResult(5) = varResultAutista1(5)
                                    blnCheckScadenze = True
                                    Exit Function
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        End If
                        
                        ' -- Autista 2
                        If Not IsNull(varBuffer(1, 0)) Then
                            blnResultAutista2 = blnCheckScadenze(varBuffer(1, 0), dtDateToCheck, "AU", varResultAutista2)
                        
                            If blnResultAutista2 Then
                                If varResultAutista2(0) > 0 Then
                                    varResult(0) = varResultAutista2(0)
                                    varResult(1) = varResultAutista2(1)
                                    varResult(2) = varResultAutista2(2)
                                    varResult(3) = varResultAutista2(3)
                                    varResult(4) = varResultAutista2(4)
                                    varResult(5) = varResultAutista2(5)
                                    blnCheckScadenze = True
                                    Exit Function
                                End If
                            Else
                                GoTo ErrHandler
                            End If
                        End If
                        
                        ' -- Non ci sono vincoli da segnalare nemmeno per autisti ed automezzi di questa U.D.T.
                        varResult(0) = 0
                    Else
                        GoTo ErrHandler
                    End If
                Else
                    GoTo ErrHandler
                End If
            
            Case Else
                GoTo ErrHandler
        
        End Select
    Else
        ' -- Flag "Controllo Scadenze" in ANAGAXX non abilitato
        varResult(0) = 0
    End If

    blnCheckScadenze = True
    Exit Function

ErrHandler:
    blnCheckScadenze = False
    varResult = Empty
End Function

Private Sub Class_Initialize()
    Set DS = New CDataStruct
    With DS.Fields
        .Add "Codice Ente", "ANAGSCA0.SCAIDENT", 0, False, 10, "0"
        .Add "Tipo Scadenza", "ANAGSCA0.SCAIDATA", 0, False, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ATA"
            .ForeignTable = "anagata0"
            .ForeignTableFld = "ataidata"
            .ForeignTableDescField = "atadescr"
            .ForeignTableDeleteFld = "ataannul"
        End With
        .Add "Data Da", "ANAGSCA0.SCADTFIV", 2, False, 10, "2", ">="
        .Add "Data A", "ANAGSCA0.SCADTFIV", 2, False, 10, "3", "<="
        .Add "Annullato", "ANAGSCA0.SCAANNUL", 3, True, , "4"
        .Add "Entità", "ANAGATA0.ATAFLENT", 0, False, 3, "5"
        .Add "Tipo Automezzi", "ANAGAAM0.AAMIDAVE", 0, False, 3, "6"
        .Add "Vettore Automezzi", "ANAGAAM0.AAMIDAVE", 0, False, 3, "7"
        With .Item("7")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
    End With
    
    DS.CancelField = "ANAGSCA0.SCAANNUL"
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
