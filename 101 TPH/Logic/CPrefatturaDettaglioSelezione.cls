VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaDettaglioSelezione"
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
Public SQL_GET_SELECTION As String
Public SQL_GET_SELECTION_MINIMO As String
Public SQL_GET_SELECTION_MINIMO_2 As String
Public SQL_GET_SELECTION_EXPORT As String
Public SQL_GET_SELECTION_EXPORT_GP2 As String

'Private Const SQL_GET_SELECTION_E2K_EXPORT = "SELECT YEAR(PREFPPD0.PPDDTVIA) AS AD001, '01' AS AD002, PREFPPD0.PPDIDVVT AS AD003, VIAGVVC0.VVCFLST1 AS AD004," & _
'                                        " DOCUDAD0.DADIDADD AS AD005, ANAGADD0.ADDDESCR AS AD006, '01' AS AD007, DOCUDAM0.DAMIDACL AS AD008, PREFPAD0.PADIDACN AS AD009," & _
'                                        " '' AS AD010, 0 AS AD011, PREFPAD0.PADIMTAR AS AD012, PREFPAD0.PADIMIMP AS AD013, DOCUDAR0.DARIDAIV AS AD014," & _
'                                        " DOCUDAM0.DAMIDAND AS AD015, '01' AS AD016, DOCUDAM0.DAMNRDAM AS AD017, DOCUDAM0.DAMDTDOC AS AD018, 0 AS AD019, '' AS AD020," & _
'                                        " 0 AS AD021, DOCUDAM0.DAMIDATO AS AD022, DOCUDAM0.DAMIDAMP AS AD023, DOCUDAM0.DAMIDAVA AS AD024, 0 AS AD025," & _
'                                        " PREFPPD0.PPDDTVIA AS AD026, '01' AS AD027, 'D' AS AD028, '' AS AD029, PREFPPD0.PPDIDPPT AS AD030, PREFPPD0.PPDIDPPD AS AD031, '' AS AD032" & _
'                                        " FROM ANAGADD0 INNER JOIN" & _
'                                        " DOCUDAM0 INNER JOIN" & _
'                                        " DOCUDAR0 ON DOCUDAM0.DAMIDAEE = DOCUDAR0.DARIDAEE AND DOCUDAM0.DAMIDAND = DOCUDAR0.DARIDAND AND" & _
'                                        " DOCUDAM0.DAMIDDAM = DOCUDAR0.DARIDDAM INNER JOIN" & _
'                                        " DOCUDAD0 ON DOCUDAR0.DARIDAEE = DOCUDAD0.DADIDAEE AND DOCUDAR0.DARIDAND = DOCUDAD0.DADIDAND AND" & _
'                                        " DOCUDAR0.DARIDDAM = DOCUDAD0.DADIDDAM AND DOCUDAR0.DARIDDAR = DOCUDAD0.DADIDDAR INNER JOIN" & _
'                                        " PREFPAD0 INNER JOIN" & _
'                                        " PREFPPD0 ON PREFPAD0.PADIDPPT = PREFPPD0.PPDIDPPT AND PREFPAD0.PADIDPPD = PREFPPD0.PPDIDPPD ON " & _
'                                        " DOCUDAD0.DADIDAUT = PREFPPD0.PPDIDAUT AND DOCUDAD0.DADIDPPT = PREFPPD0.PPDIDPPT ON " & _
'                                        " ANAGADD0.ADDIDADD = PREFPAD0.PADIDADD INNER JOIN" & _
'                                        " VIAGVVC0 ON PREFPPD0.PPDIDVVT = VIAGVVC0.VVCIDVVT AND PREFPPD0.PPDIDVVC = VIAGVVC0.VVCIDVVC"

'Private Const SQL_GET_SELECTION_E2K_EXPORT = "SELECT YEAR(PREFPPD0.PPDDTVIA) AS AD001, '01' AS AD002, PREFPPD0.PPDIDVVT AS AD003, VIAGVVC0.VVCFLST1 AS AD004," & _
'                      " PREFPAD0.PADIDADD AS AD005, ANAGADD0.ADDDESCR AS AD006, '01' AS AD007, PREFPAD0.PADIDACL AS AD008, PREFPAD0.PADIDACN AS AD009," & _
'                      " '' AS AD010, 0 AS AD011, PREFPAD0.PADIMTAR AS AD012, PREFPAD0.PADIMIMP AS AD013, ANAGADD0.ADDIDAIV AS AD014," & _
'                      " DOCUDAM0.DAMIDAND AS AD015, '01' AS AD016, DOCUDAM0.DAMNRDAM AS AD017, DOCUDAM0.DAMDTDOC AS AD018, 0 AS AD019, '' AS AD020," & _
'                      " 0 AS AD021, DOCUDAM0.DAMIDATO AS AD022, DOCUDAM0.DAMIDAMP AS AD023, DOCUDAM0.DAMIDAVA AS AD024, 0 AS AD025," & _
'                      " PREFPPD0.PPDDTVIA AS AD026, '01' AS AD027, 'D' AS AD028, '' AS AD029, PREFPPD0.PPDIDPPT AS AD030, PREFPPD0.PPDIDPPD AS AD031," & _
'                      " '' AS AD032" & _
'                      " FROM VIAGVVC0 INNER JOIN" & _
'                      " ANAGADD0 INNER JOIN" & _
'                      " PREFPAD0 INNER JOIN" & _
'                      " PREFPPD0 ON PREFPAD0.PADIDPPT = PREFPPD0.PPDIDPPT AND PREFPAD0.PADIDPPD = PREFPPD0.PPDIDPPD ON" & _
'                      " ANAGADD0.ADDIDADD = PREFPAD0.PADIDADD ON VIAGVVC0.VVCIDVVT = PREFPPD0.PPDIDVVT AND" & _
'                      " VIAGVVC0.VVCIDVVC = PREFPPD0.PPDIDVVC INNER JOIN" & _
'                      " PREFPPT0 ON PREFPPD0.PPDIDPPT = PREFPPT0.PPTIDPPT LEFT OUTER JOIN" & _
'                      " DOCUDAM0 INNER JOIN" & _
'                      " DOCUDAR0 ON DOCUDAM0.DAMIDAEE = DOCUDAR0.DARIDAEE AND DOCUDAM0.DAMIDAND = DOCUDAR0.DARIDAND AND" & _
'                      " DOCUDAM0.DAMIDDAM = DOCUDAR0.DARIDDAM INNER JOIN" & _
'                      " DOCUDAD0 ON DOCUDAR0.DARIDAEE = DOCUDAD0.DADIDAEE AND DOCUDAR0.DARIDAND = DOCUDAD0.DADIDAND AND" & _
'                      " DOCUDAR0.DARIDDAM = DOCUDAD0.DADIDDAM AND DOCUDAR0.DARIDDAR = DOCUDAD0.DADIDDAR ON" & _
'                      " PREFPPD0.PPDIDAUT = DOCUDAD0.DADIDAUT And PREFPPD0.PPDIDPPT = DOCUDAD0.DADIDPPT"


Private Const SQL_GET_SELECTION_E2K_EXPORT = "SELECT YEAR(PREFPPD0.PPDDTVIA) AS AD001, '01' AS AD002, PREFPPD0.PPDIDVVT AS AD003, VIAGVVC0.VVCFLST1 AS AD004," & _
                      " PREFPAD0.PADIDADD AS AD005, ANAGADD0.ADDDESCR AS AD006, '01' AS AD007, PREFPAD0.PADIDACL AS AD008, PREFPAD0.PADIDACN AS AD009," & _
                      " '' AS AD010, ACDMINTA AS AD011, PREFPAD0.PADIMTAR AS AD012, PREFPAD0.PADIMIMP AS AD013, ANAGADD0.ADDIDAIV AS AD014," & _
                      " DOCUDAM0.DAMIDAND AS AD015, '01' AS AD016, DOCUDAM0.DAMNRDAM AS AD017, DOCUDAM0.DAMDTDOC AS AD018, 0 AS AD019, '' AS AD020," & _
                      " 0 AS AD021, DOCUDAM0.DAMIDATO AS AD022, DOCUDAM0.DAMIDAMP AS AD023, DOCUDAM0.DAMIDAVA AS AD024, 0 AS AD025," & _
                      " PREFPPD0.PPDDTVIA AS AD026, '01' AS AD027, 'D' AS AD028, '' AS AD029, PREFPPD0.PPDIDPPT AS AD030, PREFPPD0.PPDIDPPD AS AD031," & _
                      " '' AS AD032" & _
                      " FROM VIAGVVC0 INNER JOIN" & _
                      " ANAGADD0 INNER JOIN" & _
                      " PREFPAD0 INNER JOIN" & _
                      " PREFPPD0 ON PREFPAD0.PADIDPPT = PREFPPD0.PPDIDPPT AND PREFPAD0.PADIDPPD = PREFPPD0.PPDIDPPD ON" & _
                      " ANAGADD0.ADDIDADD = PREFPAD0.PADIDADD ON VIAGVVC0.VVCIDVVT = PREFPPD0.PPDIDVVT AND" & _
                      " VIAGVVC0.VVCIDVVC = PREFPPD0.PPDIDVVC INNER JOIN" & _
                      " PREFPPT0 ON PREFPPD0.PPDIDPPT = PREFPPT0.PPTIDPPT LEFT OUTER JOIN" & _
                      " DOCUDAM0 INNER JOIN" & _
                      " DOCUDAR0 ON DOCUDAM0.DAMIDAEE = DOCUDAR0.DARIDAEE AND DOCUDAM0.DAMIDAND = DOCUDAR0.DARIDAND AND" & _
                      " DOCUDAM0.DAMIDDAM = DOCUDAR0.DARIDDAM INNER JOIN" & _
                      " DOCUDAD0 ON DOCUDAR0.DARIDAEE = DOCUDAD0.DADIDAEE AND DOCUDAR0.DARIDAND = DOCUDAD0.DADIDAND AND" & _
                      " DOCUDAR0.DARIDDAM = DOCUDAD0.DADIDDAM AND DOCUDAR0.DARIDDAR = DOCUDAD0.DADIDDAR ON" & _
                      " PREFPAD0.PADIDADD = DOCUDAD0.DADIDADD AND PREFPAD0.PADIDDAM = DOCUDAD0.DADIDDAM AND" & _
                      " PREFPPD0.PPDIDAUT = DOCUDAD0.DADIDAUT AND PREFPPD0.PPDIDPPT = DOCUDAD0.DADIDPPT" & _
                      " INNER JOIN ANAGACD0 ON ACDIDACL = PADIDACL AND ACDIDACN = PADIDACN AND ACDIDADD = PADIDADD"

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, _
                                       OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, _
                                       , MaxRecords)
    End If
End Function

Public Function GetSelectionExport(Optional ActiveRecordsOnly As Boolean = False, _
                                   Optional OrderFields As Variant, _
                                   Optional MaxRecords As Long = 0) As Boolean
    ReDim varOrderFields(0 To 2) As Variant
    ' Selezioniamo solo le righe per cui sono segnalate anomalie
    DS.Fields("50").Operator = " <> "
    DS.Fields("50").FindValue = 0
    
    varOrderFields(0) = "0"
    varOrderFields(1) = "12"
    varOrderFields(2) = "1"
    
    If Not IsMissing(OrderFields) Then
        GetSelectionExport = DS.GetSelection(SQL_GET_SELECTION_EXPORT, ActiveRecordsOnly, _
                                       OrderFields, MaxRecords)
    Else
        GetSelectionExport = DS.GetSelection(SQL_GET_SELECTION_EXPORT, ActiveRecordsOnly, _
                                       varOrderFields, MaxRecords)
    End If
    DS.Fields("50").Operator = " = "
    DS.Fields("50").FindValue = Empty
End Function

Public Function GetSelectionExportGP2(Optional ActiveRecordsOnly As Boolean = False, _
                                   Optional OrderFields As Variant, _
                                   Optional MaxRecords As Long = 0) As Boolean
    ReDim varOrderFields(0 To 2) As Variant
'    ' Selezioniamo solo le righe per cui sono segnalate anomalie
'    DS.Fields("50").Operator = " <> "
'    DS.Fields("50").FindValue = 0
    
    varOrderFields(0) = "0"
    varOrderFields(1) = "12"
    varOrderFields(2) = "1"
    
    If Not IsMissing(OrderFields) Then
        GetSelectionExportGP2 = DS.GetSelection(SQL_GET_SELECTION_EXPORT_GP2, ActiveRecordsOnly, _
                                       OrderFields, MaxRecords)
    Else
        GetSelectionExportGP2 = DS.GetSelection(SQL_GET_SELECTION_EXPORT_GP2, ActiveRecordsOnly, _
                                       varOrderFields, MaxRecords)
    End If
'    DS.Fields("50").Operator = " = "
'    DS.Fields("50").FindValue = Empty
End Function

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function GetSelectionMinimo() As Boolean
    Dim strSQL As String
    
    DS.Fields("49").FindValue = "V"
    strSQL = SQL_GET_SELECTION_MINIMO
    strSQL = strSQL & DS.MakeKeyCond
    strSQL = strSQL & SQL_GET_SELECTION_MINIMO_2
    GetSelectionMinimo = DS.GetSelection(strSQL, True, , , False)
End Function



Public Function Insert() As Boolean
    DS.Fields("49").NewValue = "V"
    DS.Fields("49").Updated = True
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    SQL_FIND = "SELECT ppdidppt, ppdidppd, ppdidoct, ppdidocr, " & _
               "       ppddtior, ppddtrco, ppdrife1, ppdrife2, " & _
               "       ppdqtikg, ppdqtilt, ppdqtipz, ppdqticn, " & _
               "       ppdidvvt, ppdidvvc, ppdidvvp, ppddtvia, " & _
               "       ppdqttkg, ppdqttlt, ppdqttpz, ppdqttcn, " & _
               "       ppdidasc, ppdidaee, ppdidade, ppdidacl, " & _
               "       ppdidadz, ppdrifde, ppdidapr, ppdidaim, " & _
               "       ppdidapd, ppdqtakg, ppdqtalt, ppdqtapz, " & _
               "       ppdqtacn, ppdidaze, ppdidtfe, ppdvalta, " & _
               "       ppdimpva, ppdflimf, ppdidave, ppdqkmra, " & _
               "       ppdfltpe, ppdidazi, ppdidtfi, ppdvalti, " & _
               "       ppdimpvi, ppdflifi, ppdidavi, ppdqkmri, " & _
               "       ppdfltpi, ppdflstr, ppdflano, ppdflst2, " & _
               "       ppdflst3, ppdflst4, ppdflst5, ppdidaut  " & _
               "       timestmp FROM prefppd0 WHERE "
    SQL_GETLIST = "SELECT ppdidppt, ppdidppd FROM prefppd0 "
    SQL_UPDATE = "UPDATE prefppd0 SET "
    SQL_INSERT = "INSERT INTO prefppd0 "
    SQL_GET_TS = "SELECT timestmp FROM prefppd0 WHERE "
    SQL_GET_SELECTION = "SELECT ppdidppt, ppdidppd, ppdidoct, ppdidocr, " & _
                        "       ppddtior, ppddtrco, ppdrife1, ppdrife2 " & _
                        "  FROM prefppt0 INNER JOIN prefppd0 ON pptidppt = ppdidppt "
    
    SQL_GET_SELECTION_MINIMO = "SELECT PPDIDAUT, PPDIDAZI, PPDDTVIA, " & _
                               "       Sum(PPDIMPVI) AS TotImporti, " & _
                               "       Max(PREFPPD0.PPDFLTPI) AS TipoPercorso, " & _
                               "       Max(PREFPPD0.PPDIDPPD) AS UltimaRiga " & _
                               "  From PREFPPD0 " & _
                               " WHERE "
    SQL_GET_SELECTION_MINIMO_2 = "GROUP BY PREFPPD0.PPDIDAUT, PREFPPD0.PPDIDAZI, PREFPPD0.PPDDTVIA " & _
                                 "ORDER BY PREFPPD0.PPDIDAUT, PREFPPD0.PPDIDAZI, PREFPPD0.PPDDTVIA "
    
    '-- DP 20/08/08
    '-- Aggiunta la colonna ppdimsei per l'importo servizi esenti
    SQL_GET_SELECTION_EXPORT = "SELECT ppdidppt, ppdidppd, ppdflano, ppdidoct, " & _
                               "       ppdrife1, ppdidvvt, ppddtvia, ppdidaut, " & _
                               "       adzrasoc, ppdidapr, ppdqtalt, ppdvalti, " & _
                               "       ppdimpvi, ppdidadz, ppdidade, ppdqkmri, " & _
                               "       vvtidam1, vvtidam2, vvtidau1, vvtseqvi, " & _
                               "       ppdqtakg, octnota2, ppdrife1, octnota3, " & _
                               "       vvtflst5, ppdimppi, vvtdtinc, '', " & _
                               "       ppdflano, ppdimpoi, ppdimpbi, ppdimphi, " & _
                               "       ppdimpsi, octflres, octnonav, octfltor, " & _
                               "       ocrnrtck, ppdimsei, ppdidtfi " & _
                               "  FROM prefppd0 INNER JOIN anagadz0 ON ppdidadz = adzidadz  " & _
                               "  INNER JOIN viagvvt0 ON ppdidvvt = vvtidvvt " & _
                               "  INNER JOIN ordioct0 ON ppdidoct = octidoct " & _
                               "  LEFT OUTER JOIN ordiocr0 ON ppdidoct = ocridoct AND ppdidocr = ocridocr "

    SQL_GET_SELECTION_EXPORT_GP2 = "SELECT ppdidppt, ppdidppd, ppdflano, ppdidoct, " & _
                                   "       ppdrife1, ppdidvvt, ppddtvia, ppdidaut, " & _
                                   "       adzrasoc, ppdidapr, ppdqtalt, ppdvalti, " & _
                                   "       ppdimpvi, ppdidadz, ppdidade, ppdqkmri, " & _
                                   "       vvtidam1, vvtidam2, vvtidau1, vvtseqvi, " & _
                                   "       ppdqtakg, octnota2, ppdrife1, octnota3, " & _
                                   "       vvtflst5, ppdimppi, vvtdtinc, autbadge, " & _
                                   "       ppdflano, ppdidvvc, ppdidvvp, ppdidazi,  " & _
                                   "       adeidarp, octrife1, octidatv, ppddtgrs, " & _
                                   "       ppdidtfi " & _
                                   "  FROM prefppd0 INNER JOIN anagadz0 ON ppdidadz = adzidadz  " & _
                                   "       INNER JOIN viagvvt0 ON ppdidvvt = vvtidvvt INNER JOIN ordioct0 ON ppdidoct = octidoct " & _
                                   "       INNER JOIN anagaut0 ON ppdidaut = autidaut " & _
                                   "       INNER JOIN anagade0 ON ppdidade = adeidade "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Prefatturazione", "prefppd0.ppdidppt", 0, , 7, "0"
        .Add "Numero Dettaglio", "prefppd0.ppdidppd", 1, , 7, "1"
        .Add "Ordine cliente", "prefppd0.ppdidoct", 0, , 7, "2"
        .Add "Riferimento esterno 1", "prefppd0.ppdrife1", 0, , 15, "3"
        .Add "Riferimento esterno 2", "prefppd0.ppdrife2", 0, , 15, "7"
        .Add "Numero Viaggio", "prefppd0.ppdidvvt", 0, , 7, "12"
        .Add "Dal Numero Viaggio", "prefppd0.ppdidvvt", 0, , 7, "14", " >= "
        .Add "Al Numero Viaggio", "prefppd0.ppdidvvt", 0, , 7, "15", " <= "
        .Add "Dalla Data viaggio", "prefppd0.ppddtvia", 2, , 10, "5", " >= "
        .Add "Alla Data viaggio", "prefppd0.ppddtvia", 2, , 10, "6", " <= "
        .Add "Società", "prefppd0.ppdidasc", 0, , 3, "20"
        With .Item("20")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "prefppd0.ppdidaee", 0, , 3, "21"
        With .Item("21")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Deposito", "prefppd0.ppdidade", 0, , 3, "22"
        With .Item("22")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Cliente", "prefppd0.ppdidacl", 0, , 10, "23"
        With .Item("23")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Destinazione", "prefppd0.ppdidadz", 0, , 10, "24"
        With .Item("24")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Rif. Esterno Destinazione", "prefppd0.ppdrifde", 0, , 15, "25"
        .Add "Prodotto", "prefppd0.ppdidapr", 0, , 6, "26"
        With .Item("26")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Zona Fascia Tariffaria Interna", "prefppd0.ppdidazi", 0, , 3, "17"
        .Add "Fascia Tariffaria Interna", "prefppd0.ppdidtfi", 0, , 3, "42"
        .Add "Stato Riga", "prefppd0.ppdflstr", 0, , 1, "49"
        .Add "Anomalia", "prefppd0.ppdflano", 1, , 5, "50"
        .Add "Unità di trasporto", "prefppd0.ppdidaut", 0, , 3, "13"
        With .Item("13")
            .ForeignTableEntity = "AUT"
            .ForeignTable = "anagaut0"
            .ForeignTableFld = "autidaut"
            .ForeignTableDescField = "autdescr"
            .ForeignTableDeleteFld = "autannul"
        End With
        .Add "Dalla Data Documento", "DOCUDAM0.DAMDTDOC", 2, False, 10, "51", " >= "
        .Add "Alla Data Documento", "DOCUDAM0.DAMDTDOC", 2, False, 10, "52", " <= "
        .Add "Trasmesso", "prefppd0.ppdfltrc", 0, , 2, "65"
        .Add "Data viaggio", "prefppd0.ppddtvia", 2, , 10, "66"
        .Add "Dalla Data inizio periodo", "prefppt0.pptdtinp", 2, , 10, "67", " >= "
        .Add "Alla Data fine periodo", "prefppt0.pptdtfip", 2, , 10, "68", " <= "
        .Add "Prefattura Stato Diverso", "prefppt0.pptflsta", 0, , 2, "69", " <> "
        .Add "Dalla Data Inizio Carico", "VIAGVVT0.VVTDTINC", 2, False, 10, "70", " >= "
        .Add "Alla Data Inizio Carico", "VIAGVVT0.VVTDTINC", 2, False, 10, "71", " <= "
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "13", "3"
        .Add "17", "4"
        .Add "5", "5"
        .Add "6", "6"
        .Add "49", "7"
        .Add "20", "8"
        .Add "14", "9"
        .Add "15", "10"
        .Add "50", "11"
        .Add "51", "12"
        .Add "52", "13"
        .Add "65", "14"
        .Add "22", "15"
        .Add "67", "16"
        .Add "68", "17"
        .Add "69", "18"
        .Add "70", "19"
        .Add "71", "20"
        .Add "21", "21"
    End With
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


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property


Public Function GetSelectionE2K(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean

    If Not IsMissing(OrderFields) Then
        GetSelectionE2K = DS.GetSelection(SQL_GET_SELECTION_E2K_EXPORT, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionE2K = DS.GetSelection(SQL_GET_SELECTION_E2K_EXPORT, ActiveRecordsOnly, , MaxRecords)
    End If

End Function




