VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentoSelezione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_SET_AVANZAMENTO As String
Private Const SQL_GET_SELECTION = _
        " SELECT DAMIDAEE, DAMIDAND, DAMDTDOC, DAMIDDAM, " & _
        "        DAMNRDAM, DAMFLESI, DAMIDACL, DAMRASOC, " & _
        "        AMPDESCR, DAMFLINC, DAMNRIMP, DAMNRIMI, " & _
        "        (DAMNRIMI - DAMNRIMP) AS DIFFERENZA, DAMIDVVT, DAMIDVVC " & _
        "   FROM DOCUDAM0 INNER JOIN  ANAGAMP0 ON DAMIDAMP = AMPIDAMP "
        
Private Const SQL_GET_INCASSO_DETAIL_FOR_DOC = " SELECT INCAICT0.ICTDTICT, INCAICT0.ICTIDAEE,INCAICT0.ICTIDICT, INCAICT0.ICTDESCR, INCAICT0.ICTFLSTC, " & _
                                               " INCATDI0.TDIDESCR, INCAIIR0.IIRIMPIN,INCAABA0.ABARASAB, INCAIIR0.IIRRIFBA,INCAIIR0.IIRRIFIN " & _
                                               " FROM INCAIIT0 INNER JOIN " & _
                                               " INCAIIR0 ON INCAIIT0.IITIDAEE = INCAIIR0.IIRIDAEE AND " & _
                                               " INCAIIT0.IITIDICT = INCAIIR0.IIRIDICT AND " & _
                                               " INCAIIT0.IITIDEED = INCAIIR0.IIRIDEED AND " & _
                                               " INCAIIT0.IITIDAND = INCAIIR0.IIRIDAND AND " & _
                                               " INCAIIT0.IITIDDAM = INCAIIR0.IIRIDDAM INNER JOIN " & _
                                               " INCATDI0 ON " & _
                                               " INCAIIR0.IIRIDTDI = INCATDI0.TDIIDTDI INNER JOIN " & _
                                               " INCAICT0 ON INCAIIT0.IITIDAEE = INCAICT0.ICTIDAEE AND " & _
                                               " INCAIIT0.IITIDICT = INCAICT0.ICTIDICT LEFT JOIN " & _
                                               " INCAABA0 ON INCAIIR0.IIRIDABA = INCAABA0.ABAIDABA " & _
                                               " WHERE (INCAIIR0.IIRIDEED = ?) AND (INCAIIR0.IIRIDAND = ?) AND " & _
                                               " (INCAIIR0.IIRIDDAM = ?) ORDER BY INCAICT0.ICTIDAEE,INCAICT0.ICTDTICT DESC,INCAICT0.ICTIDICT"

Private Const SQL_GET_SELECTION_DAM1 = " SELECT DOCUDAM0.DAMIDAEE, DOCUDAM0.DAMIDAND, DOCUDAM0.DAMIDDAM, DOCUDAM0.DAMIDASC, " & _
                                      " DOCUATO0.ATODESCR, DOCUDAM0.DAMNRDAM, DOCUDAM0.DAMDTDOC, DOCUDAM0.DAMIDACL, " & _
                                      " ANAGACL0.ACLRASOC, DOCUDAM0.DAMIDADE, DOCUDAM0.DAMIDAMD, ANAGAMP0.AMPDESCR, " & _
                                      " DOCUDAM0.DAMNRIMP , DOCUDAM0.DAMNRIMI, DOCUDAM0.DAMNRIVA,DOCUDAM0.damannul " & _
                                      " FROM DOCUDAM0 LEFT OUTER JOIN ANAGAMP0 ON DOCUDAM0.DAMIDAMP = ANAGAMP0.AMPIDAMP LEFT OUTER JOIN " & _
                                      " DOCUATO0 ON DOCUDAM0.DAMIDATO = DOCUATO0.ATOIDATO LEFT OUTER JOIN " & _
                                      " ANAGACL0 ON DOCUDAM0.DAMIDACL = ANAGACL0.ACLIDACL"
                                      
                                      
Private Const SQL_GET_SELECTION_DAM = " SELECT DAMIDAEE, DAMIDAND, DAMIDDAM, DAMIDASC, " & _
                                       "       ATODESCR, DAMNRDAM, DAMDTDOC, DAMIDACL, " & _
                                       "       ACLRASOC, DAMIDADE, DAMIDAMD, AMPDESCR, " & _
                                       "       DAMNRIMP, DAMNRIMI, DAMNRIVA, DAMIDOCT, " & _
                                       "       DAMIDVVT, DAMIDVVC, DAMANNUL " & _
                                       "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON " & _
                                       "       VVTIDVVT = VVCIDVVT RIGHT OUTER JOIN DOCUDAM0 ON " & _
                                       "       VVCIDVVT = DAMIDVVT AND VVCIDVVC = DAMIDVVC LEFT OUTER JOIN ANAGACL0 ON " & _
                                       "       DAMIDACL = ACLIDACL LEFT OUTER JOIN ANAGAMP0 ON " & _
                                       "       DAMIDAMP = AMPIDAMP LEFT OUTER JOIN DOCUATO0 ON " & _
                                       "       DAMIDATO = ATOIDATO "


Private Const SQL_GET_SELECTION_DAM_PAD = "SELECT DISTINCT DAMIDAEE, DAMIDAND, DAMIDDAM, DAMIDASC, ATODESCR, DAMNRDAM," & _
                            " DAMDTDOC, DAMIDACL, ACLRASOC, DAMIDADE, DAMIDAMD, AMPDESCR," & _
                            " DAMNRIMP, DAMNRIMI, DAMNRIVA, DAMIDOCT, DAMIDVVT, DAMIDVVC," & _
                            " DAMANNUL, PPTIDPPT, DAMFLESI" & _
                            " FROM PREFPAD0 INNER JOIN" & _
                            " DOCUDAM0 ON PREFPAD0.PADIDAEE = DOCUDAM0.DAMIDAEE AND PREFPAD0.PADIDAND = DOCUDAM0.DAMIDAND AND" & _
                            " PREFPAD0.PADIDDAM = DOCUDAM0.DAMIDDAM INNER JOIN" & _
                            " PREFPPD0 ON PREFPAD0.PADIDPPT = PREFPPD0.PPDIDPPT AND PREFPAD0.PADIDPPD = PREFPPD0.PPDIDPPD INNER JOIN" & _
                            " PREFPPT0 ON PREFPPD0.PPDIDPPT = PREFPPT0.PPTIDPPT LEFT OUTER JOIN" & _
                            " VIAGVVT0 INNER JOIN" & _
                            " VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT ON DOCUDAM0.DAMIDVVT = VIAGVVC0.VVCIDVVT AND" & _
                            " DOCUDAM0.DAMIDVVC = VIAGVVC0.VVCIDVVC LEFT OUTER JOIN" & _
                            " ANAGACL0 ON DOCUDAM0.DAMIDACL = ANAGACL0.ACLIDACL LEFT OUTER JOIN" & _
                            " ANAGAMP0 ON DOCUDAM0.DAMIDAMP = ANAGAMP0.AMPIDAMP LEFT OUTER JOIN" & _
                            " DOCUATO0 ON DOCUDAM0.DAMIDATO = DOCUATO0.ATOIDATO"

Private Const SQL_GET_SELECTION_PREFATTURE = "SELECT DISTINCT PREFPAD0.PADIDPPT, PREFPPT0.PPTIDPPT" & _
                            " FROM DOCUDAM0 INNER JOIN" & _
                            " PREFPAD0 ON DOCUDAM0.DAMIDAEE = PREFPAD0.PADIDAEE AND DOCUDAM0.DAMIDAND = PREFPAD0.PADIDAND AND" & _
                            " DOCUDAM0.DAMIDDAM = PREFPAD0.PADIDDAM INNER JOIN" & _
                            " PREFPPD0 ON PREFPAD0.PADIDPPT = PREFPPD0.PPDIDPPT AND PREFPAD0.PADIDPPD = PREFPPD0.PPDIDPPD INNER JOIN" & _
                            " PREFPPT0 ON PREFPPD0.PPDIDPPT = PREFPPT0.PPTIDPPT"

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Private mvarTotaliQta As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, , MaxRecords)
    End If
End Function


Public Function GetSelectionDAM(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean

    If Not IsMissing(OrderFields) Then
        GetSelectionDAM = DS.GetSelection(SQL_GET_SELECTION_DAM, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionDAM = DS.GetSelection(SQL_GET_SELECTION_DAM, ActiveRecordsOnly, , MaxRecords)
    End If
End Function
Public Function GetSelectionDAM_PAD(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean

    If Not IsMissing(OrderFields) Then
        GetSelectionDAM_PAD = DS.GetSelection(SQL_GET_SELECTION_DAM_PAD, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionDAM_PAD = DS.GetSelection(SQL_GET_SELECTION_DAM_PAD, ActiveRecordsOnly, , MaxRecords)
    End If
End Function

Public Function GetSelectionPrefatture() As Boolean
    GetSelectionPrefatture = DS.GetSelection(SQL_GET_SELECTION_PREFATTURE, True)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
'    On Error GoTo ErrHandler
'
'    Dim strAvanzamento As String
'    Dim strIDDestinazione As String
'    Dim strDataInserimento As String
'    Dim blnUpdatePrices As Boolean
'    Dim varRows As Variant
'    Dim i As Integer
'
'    If (DS.Fields("9").Updated) Or (DS.Field("Data Inserimento").Updated) Then
'        blnUpdatePrices = True
'        strIDDestinazione = DS.Field("Destinazione").NewValue
'        strDataInserimento = DS.Field("Data Inserimento").NewValue
'    End If
'
'    ' --- Impedisco il salvataggio di ordini il cui stato avanzamento
'    ' --- sia "V" o "F"
'    strAvanzamento = DS.Fields("23").Value
'    If (strAvanzamento = "A") Or (strAvanzamento = "M") Or (strAvanzamento = "V") Then
'        DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
'    Else
'        BeginSave = False
'        Exit Function
'    End If
'
'    BeginSave = True
'    GoTo EndFunction
'    ' ---
'ErrHandler:
    DS.DA.EndTrans etRollback
    BeginSave = False
EndFunction:
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = " SELECT DAMIDAEE, DAMIDAND, DAMIDDAM, DAMIDASC, " & _
               " DAMIDADZ, DAMIDACL, DAMDTDOC, DAMDTCON, DAMIDATO, " & _
               " DAMNRDAM, DAMFLESI, DAMNRIMP, DAMNRIMI, DAMNRIVA, " & _
               " DAMIDAVA, DAMIDADE, DAMIDAMD, DAMIDAMP, DAMNRGDL, " & _
               " DAMIDATV, DAMRASOC, DAMINDIR, DAMLOCAL, DAMPROVI, " & _
               " DAMNRCAP, DAMRASDE, DAMINDDE, DAMLOCDE, " & _
               " DAMPRODE, DAMNRCDE, DAMRASDZ, DAMINDDZ, " & _
               " DAMLOCDZ, DAMPRODZ, DAMNRCDZ, DAMCOFIS, DAMPAIVA, " & _
               " DAMIDAVE, DAMIDOCT, DAMIDVVT, DAMIDVVC, DAMFLGES, " & _
               " DAMFLINC, DAMFLSLD, DAMNOTES, DAMNOPIE, DAMANNUL, " & _
               " TIMESTMP " & _
               " FROM DOCUDAM0 WHERE "
               
    SQL_GETLIST = "SELECT DAMIDAEE, DAMIDAND, DAMIDDAM,DAMIDADZ, DAMIDACL, DAMDTDOC FROM DOCUDAM0 "
    SQL_UPDATE = "UPDATE DOCUDAM0 SET "
    SQL_INSERT = "INSERT INTO DOCUDAM0 "
    SQL_GET_TS = "SELECT timestmp FROM DOCUDAM0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT OMRIDADE, OMRANNOM, OMRIDOMR" & _
                          "  FROM DOCUDAM0 "
    
    mvarFieldsIndexes = Array("0", "1", "2", "3")
    
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Anno Documento", "docudam0.damidand", 1, False, 4, "1"
        .Add "Dalla Data Documento", "DOCUDAM0.DAMDTDOC", 2, False, 10, "2", " >= "
        .Add "Alla Data Documento", "DOCUDAM0.DAMDTDOC", 2, False, 10, "3", " <= "
        .Add "Dal Progressivo Documento", "DOCUDAM0.DAMIDDAM", 0, False, 7, "4", " >= "
        .Add "Al Progressivo Documento", "DOCUDAM0.DAMIDDAM", 0, False, 7, "5", " <= "
        .Add "Dal Numero Documento", "DOCUDAM0.DAMNRDAM", 0, False, 7, "6", " >= "
        .Add "Al Numero Documento", "DOCUDAM0.DAMNRDAM", 0, False, 7, "7", " <= "
        .Add "Descrizione Modalita Pagamento", "ANAGAMP0.AMPDESCR", 0, False, 25, "8"
        .Add "Destinazione", "DOCUDAM0.damidadz", 0, False, 10, "9"
        With .Item("9")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Cliente", "DOCUDAM0.damidacl", 0, False, 10, "10"
        With .Item("10")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "anagacl0.aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Dalla Data Consegna", "DOCUDAM0.DAMDTDOC", 2, False, 10, "11", " >= "
        .Add "Alla Data Consegna", "DOCUDAM0.DAMDTDOC", 2, False, 10, "12", " <= "
        .Add "Importo Documento", "docudam0.damnrimp", 1, True, , "13"
        .Add "Importo Residuo Incasso", "(DOCUDAM0.DAMNRIMI - DOCUDAM0.DAMNRIMP)", 1, True, , "14"
        .Add "Numero Viaggio", "docudam0.damidvvt", 0, False, 7, "15"
        .Add "Numero Consegna", "docudam0.damidvvc", 0, False, 3, "16"
        .Add "Ente Emittente", "docudam0.damidaee", 0, False, 3, "17", " = "
         With .Item("17") ' Tabella Ente Emittente
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Esito Documento", "docudam0.damflesi", 0, False, 1, "18", " = "
        .Add "Rag.Soc. Cliente", "docudam0.damrasoc", 0, False, 40, "19", " = "
        .Add "Flag Incasso", "docudam0.damflinc", 0, False, , "20", " = "
        
        
        .Add "Importo Incassato", "docudam0.damnrimi", 1, False, 15, "21", " = "
        .Add "Società", "docudam0.damidasc", 0, False, 3, "22", " = "
         With .Item("22") ' Tabella Società
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Tipo Documento Amministrativo", "docudam0.damidato", 0, False, 3, "23", " = "
         With .Item("23") 'Tabella Tipo Documento
            .ForeignTableEntity = "ATO"
            .ForeignTable = "docuato0"
            .ForeignTableFld = "atoidato"
            .ForeignTableDescField = "atodescr"
            .ForeignTableDeleteFld = "atoannul"
        End With
        .Add "Deposito", "docudam0.damidade", 0, False, 3, "24", " = "
         With .Item("24") ' Tabella Depositi
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Magazzino di Deposito", "docudam0.damidamd", 0, False, 3, "25", " = "
         With .Item("25") ' Tabella Magazzini di Deposito
            .ForeignTableEntity = "AMD"
            .ForeignTable = "anagamd0"
            .ForeignTableFld = "amdidamd"
            .ForeignTableDescField = "amddescr"
            .ForeignTableDeleteFld = "amdannul"
        End With
        .Add "Modalità di Pagamento", "docudam0.damidamp", 0, False, 3, "26", " = "
         With .Item("26") 'Tabella Modalità di pagamento
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        
        .Add "Tipo Vendita", "docudam0.damidatv", 0, False, 3, "27", " = "
        With .Item("27") 'Tabella Tipo Vendita
            .ForeignTableEntity = "ATV"
            .ForeignTable = "anagatv0"
            .ForeignTableFld = "atvidatv"
            .ForeignTableDescField = "atvdescr"
            .ForeignTableDeleteFld = "atvannul"
        End With
        .Add "Flag Gestione", "docudam0.damflges", 0, False, 1, "28", " = "
        .Add "Flag Saldo", "docudam0.damflsld", 0, False, , "29", " = "
        
        
        .Add "IVA Documento", "docudam0.damnriva", 1, False, 15, "30", " = "
        .Add "Annullato", "docudam0.damannul", 0, False, , "31", " = "
        .Add "Ragione Sociale", "anagacl0.aclrasoc", 0, False, 40, "55"
        
        .Add "Numero Ordine", "docudam0.damidoct", 0, False, 7, "56"
        .Add "Numero Viaggio", "DOCUDAM0.DAMIDVVT", 0, False, 7, "57"
         
         .Add "DAM Annullato", "docudam0.damannul", 3, False, , "58", "="
         .Add "DAR Annullato", "docudar0.darannul", 3, False, , "59", "="
         .Add "VVP Annullato", "viagvvp0.vvpannul", 3, , , "60", "="
         .Add "Numero Viaggio Definitivo", "VIAGVVT0.VVTIDDEF", 0, False, 7, "61"
         .Add "Data", "PREFPPD0.PPDDTVIA", 2, False, 10, "62"

        .Add "Pref. Ente Emittente", "prefppd0.ppdidaee", 0, , 3, "63"
        With .Item("63")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Pref. Società", "prefppd0.ppdidasc", 0, , 3, "64"
        With .Item("64")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Pref. Data inizio periodo", "prefppt0.pptdtinp", 2, , 10, "65", " >= "
        .Add "Pref. Data fine periodo", "prefppt0.pptdtfip", 2, , 10, "66", " <= "
        .Add "Numero Prefattura", "prefppt0.pptidppt", 0, , 7, "67", " = "
        With .Item("67")
            .ForeignTableEntity = "PPT"
            .ForeignTable = "prefppt0"
            .ForeignTableFld = "pptidppt"
            .ForeignTableDescField = "pptdescr"
            .ForeignTableDeleteFld = ""
        End With
    End With
    
    With DS.KeyFields
        .Add "1", "1"
        .Add "2", "2"
        .Add "4", "3"
        .Add "6", "4"
        .Add "8", "5"
        .Add "9", "6"
        .Add "10", "7"
        .Add "11", "8"
        .Add "13", "9"
        .Add "14", "10"
        .Add "15", "11"
        .Add "16", "12"
        .Add "17", "13"
        .Add "3", "14"
        .Add "24", "15"
        .Add "63", "16"
        .Add "64", "17"
        .Add "65", "18"
        .Add "66", "19"
        .Add "67", "20"
        
    End With
    
    DS.CancelField = "DOCUDAM0.damannul"
    
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

Public Function ShowIncassoDetailForDoc(EnteEmittenteDoc As Variant, AnnoDoc As Variant, ProgDocumento As Variant) As Variant
' Questa Funzione Restituisce, Per un certo Documento I dettagli degli Incassi Inseriti
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    ' --- Leggo I dettagli Incasso della Cassa
    'DS.DA.MakeSQL
    If Not DS.DA.QueryCommand(SQL_GET_INCASSO_DETAIL_FOR_DOC, varBuffer, lngHowMany, EnteEmittenteDoc, AnnoDoc, ProgDocumento) Then
        GoTo ErrHandler
    End If
    ' ---
    ShowIncassoDetailForDoc = varBuffer
    Exit Function
ErrHandler:
    ShowIncassoDetailForDoc = Empty
End Function

