VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaDettaglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public SQL_GETLIST As String
Public SQL_GETLIST_CHECK As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_AZZERA_CONSEGNE As String
Public SQL_APPLICA_MINIMO As String
Public SQL_DELETE As String
Public SQL_DELETE_DETAIL As String
Public SQL_ADD_EXT_DATA As String
Public SQL_ADD_EXT_DATA_2 As String
Public SQL_ADD_EXT_DATA_3 As String
Public SQL_SET_ANOMALIA As String
Public SQL_GET_ANOMALIA As String
Public SQL_GET_PEDAGGI_SOC As String
Private SQL_SET_TRASM As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

Public Enum AnomaliePrefattura
    apNoAnomalie = 0
    apDiffQuantita = 1
    apDiffKm = 2
    apTariffaMinima = 4
    apDatiDiscordanti = 8
    apPropostaNonTrovata = 16
    apDiffImporto = 32
    apVettore = 64
    apData = 128
    apTipoServizio = 256
    apUnit‡DiTrasporto = 512
    apVeicoloPerDeposito = 1024
    apTargaMotrice = 2048
    apTargaRimorchio = 4096
    apCodiceAutista = 8192
    apCodFiscAutista = 16384
    apProgressivoViaggio = 32768
    apCodiceDestinatario = 65536
    apNomeImbarcazione = 131072
    apTipoTariffa = 262144
    apTariffa = 524288
    apPedaggio = 1048576
    apDataGradoServizio = 2097152
    apServOperarioBunker = 4194304
    apServBuquesPuerto = 8388608
    apServOreServSpeciali = 16777216
    apServImpServSpeciali = 33554432
    apServImpServEsenti = 67108864
End Enum


Public Function AddExtData(Zona, Tariffa, Importo, ImportoForzato, KmRaggio, _
                           NomeFile, Anno, Valuta, IDPrefattura, IDRiga, _
                           RifEst2, _
                           Optional Anomalia As Variant = 0, _
                           Optional ForzaImporti As Boolean = False, _
                           Optional Quantita As Variant = 0, _
                           Optional QuantitaKG As Variant = 0, _
                           Optional Pedaggio As Variant = 0, _
                           Optional ServOperarioBunker As Variant = 0, _
                           Optional ServBuquesPuerto As Variant = 0, _
                           Optional ServOreServEspeciales As Variant = 0, _
                           Optional ServImportoServEspeciales As Variant = 0, _
                           Optional ServImportoServExentos As Variant = 0) As Boolean
    On Error GoTo ErrHandler
    
    With DS.DA
        .BeginTransaction
        
        If (Anomalia And apPropostaNonTrovata) <> apPropostaNonTrovata Then
            If Not ForzaImporti Then
                If Not .ExecuteCommand(SQL_ADD_EXT_DATA, Zona, Tariffa, Importo, ImportoForzato, _
                                                        KmRaggio, NomeFile, Anno, Valuta, RifEst2, _
                                                        QuantitaKG, Quantita, Pedaggio, _
                                                        ServOperarioBunker, ServBuquesPuerto, _
                                                        ServOreServEspeciales, ServImportoServEspeciales, _
                                                        ServImportoServExentos, IDPrefattura, IDRiga) Then
                    GoTo ErrHandler
                End If
            Else
                If Not .ExecuteCommand(SQL_ADD_EXT_DATA_2, Zona, Tariffa, Importo, ImportoForzato, _
                                                        KmRaggio, NomeFile, Anno, Valuta, _
                                                        Tariffa, Importo, KmRaggio, Quantita, RifEst2, _
                                                        QuantitaKG, Quantita, Pedaggio, _
                                                        ServOperarioBunker, ServBuquesPuerto, _
                                                        ServOreServEspeciales, ServImportoServEspeciales, _
                                                        ServImportoServExentos, IDPrefattura, IDRiga) Then
                    GoTo ErrHandler
                End If
            End If
        End If
        
        If Not .ExecuteCommand(SQL_SET_ANOMALIA, Anomalia, IDPrefattura, IDRiga) Then
            GoTo ErrHandler
        End If
        
        .EndTrans etCommit
    End With
    AddExtData = True
    Exit Function

ErrHandler:
    DS.DA.EndTrans etRollback
    AddExtData = False
End Function


Public Function Delete(Optional IDTestata As String, _
                       Optional IDRiga As String) As Boolean
    On Error GoTo ErrHandler
    Dim strIDTestata As String
    Dim strIDRiga As String
    
    If Len(IDTestata) = 0 Then
        strIDTestata = DS.Fields("0").Value
    Else
        strIDTestata = IDTestata
    End If
    
    If Len(IDRiga) = 0 Then
        strIDRiga = DS.Fields("1").Value
    Else
        strIDRiga = IDRiga
    End If
    
    With DS.DA
        If Not .ExecuteCommand(Me.SQL_DELETE_DETAIL, strIDTestata, strIDRiga) Then
            GoTo ErrHandler
        End If
        If Not .ExecuteCommand(Me.SQL_DELETE, strIDTestata, strIDRiga) Then
            GoTo ErrHandler
        End If
        
    End With
    
    Delete = True
    Exit Function
ErrHandler:
    Delete = False
End Function

Public Function AzzeraConsegne(IDTestata As String, _
                               IDUT As String, _
                               IDZona As String, _
                               DataViaggio As String, _
                               UltimaRiga As Long) As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    
    blnResult = DS.DA.ExecuteCommand(SQL_AZZERA_CONSEGNE, IDTestata, IDUT, _
                                IDZona, DataViaggio, UltimaRiga)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    AzzeraConsegne = True
    Exit Function
ErrHandler:
    AzzeraConsegne = False
End Function

Public Function ApplicaMinimo(ImportoMinimo As Single, _
                              IDTestata As String, _
                              IDUT As String, _
                              IDZona As String, _
                              DataViaggio As String, _
                              UltimaRiga As Long) As Boolean
    On Error GoTo ErrHandler
    
    If Not DS.DA.ExecuteCommand(SQL_APPLICA_MINIMO, ImportoMinimo, ImportoMinimo, _
                                IDTestata, IDUT, IDZona, DataViaggio, UltimaRiga) Then
        GoTo ErrHandler
    End If
    
    ApplicaMinimo = True
    Exit Function
ErrHandler:
    ApplicaMinimo = False
End Function

Public Function Find() As Boolean
    Dim blnResult As Boolean
    
    blnResult = DS.Find(Me.SQL_FIND)
    Find = blnResult
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function GetListCheck(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetListCheck = DS.GetList(SQL_GETLIST_CHECK, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    DS.Fields("49").NewValue = "V"
    DS.Fields("49").Updated = True
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean

    If Not DS.Fields("65").Updated Then
        DS.Fields("65").NewValue = "N"
        DS.Fields("65").Updated = True
    End If
    
    If DS.Fields("44").NewValue <> DS.Fields("44").Value Then
        If DS.Fields("45").Value <> "S" Then
            DS.Fields("45").NewValue = "S"
            DS.Fields("45").Updated = True
        End If
        If CDbl(DS.Fields("44").NewValue) = CDbl(DS.Fields("36").Value) Then
            ' Se gli importi interni ed esterni sono uguali resettiamo le anomalie per importo e per tariffa discordanti
            DS.Fields("50").NewValue = CLng(DS.Fields("50").Value) And 33030111 '33554399 '65503
            DS.Fields("50").Updated = True
        Else
            DS.Fields("50").NewValue = CLng(DS.Fields("50").Value) Or 32
            DS.Fields("50").Updated = True
        End If
    End If
    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    BeginSave = True
    ' ---
End Function
Private Sub Class_Initialize()
    SQL_FIND = "SELECT ppdidppt, ppdidppd, ppdidoct, ppdidocr, " & _
               "       ppddtior, ppddtrco, ppdrife1, ppdrife2, " & _
               "       ppdqtikg, ppdqtilt, ppdqtipz, ppdqticn, " & _
               "       ppdidvvt, ppdidvvc, ppdidvvp, ppddtvia, " & _
               "       ppdqttkg, ppdqttlt, ppdqttpz, ppdqttcn, " & _
               "       ppdidasc, ppdidaee, ppdidade, ppdidacl, " & _
               "       ppdidadz, ppdrifde, ppdidapr, ppdidaim, " & _
               "       ppdidapd, ppdqtakg, ppdqtalt, ppdqtapz, " & _
               "       ppdqtacn, ppdqtekg, ppdqtelt, ppdqtepz, " & _
               "       ppdqtecn, ppdidaze, ppdidtfe, ppdvalta, " & _
               "       ppdimpva, ppdflimf, ppdidave, ppdqkmra, " & _
               "       ppdfltpe, ppdimppe, ppdimpoe, ppdimpbe, " & _
               "       ppdimphe, ppdimpse, ppdidazi, ppdidtfi, " & _
               "       ppdvalti, ppdimpvi, ppdflifi, ppdidavi, " & _
               "       ppdqkmri, ppdfltpi, ppdimppi, ppdimpoi, " & _
               "       ppdimpbi, ppdimphi, ppdimpsi, ppdflstr, " & _
               "       ppdflano, ppdflst2, ppdflst3, ppdflst4, " & _
               "       ppdflst5, ppdidaut, ppdrife3, ppdannoe, " & _
               "       ppdidacn, ppdidtfc, ppdfltrc, ppddtgrs, " & _
               "       ppdimsei, ppdimsee, timestmp FROM prefppd0 WHERE "
               
    SQL_GETLIST = "SELECT ppdidppd, ppdflano, ppdidoct, ppdrife1, " & _
                  "       ppdidvvt, ppddtvia, ppdidaut, adzrasoc, " & _
                  "       ppdidapr, ppdqtalt, ppdvalti, ppdimpvi, " & _
                  "       ppdidadz, ppdidade, ppdidazi, ppdqkmri, '', '', '', '', '', ppdimpva, " & _
                  "       ppdidvvc, ppdidvvp, ppddtgrs " & _
                  "  FROM prefppd0 INNER JOIN anagadz0 ON prefppd0.ppdidadz = " & _
                  "       anagadz0.adzidadz  "
                  
    SQL_GETLIST_CHECK = "SELECT ppdidppd, ppdflano, ppdidoct, ppdrife1, " & _
                        "       ppdidvvt, ppddtvia, ppdidaut, adzrasoc, " & _
                        "       ppdidapr, ppdqtalt, ppdvalti, ppdimpvi, " & _
                        "       ppdidadz, ppdidade, ppdidazi, ppdqkmri, " & _
                        "       vvtidam1, vvtidam2, vvtidau1, aaunrcid, " & _
                        "       vvtseqvi, ppdimpva, ppdimppi, ppdimpoi, " & _
                        "       ppdimpbi, ppdimphi, ppdimpsi, adzidatd, " & _
                        "       ppdimsei" & _
                        "  FROM prefppd0 INNER JOIN anagadz0 ON ppdidadz = anagadz0.adzidadz " & _
                        "       INNER JOIN viagvvt0 ON ppdidvvt = vvtidvvt INNER JOIN anagaau0 ON vvtidau1 = aauidaau "
    SQL_UPDATE = "UPDATE prefppd0 SET "
    SQL_INSERT = "INSERT INTO prefppd0 "
    SQL_GET_TS = "SELECT timestmp FROM prefppd0 WHERE "
    SQL_AZZERA_CONSEGNE = "UPDATE prefppd0 SET ppdidtfi = '999', ppdvalti = 0, ppdimpvi = 0, " & _
                          "                    ppdflano = ppdflano + 4 " & _
                          " WHERE ppdidppt = ? AND ppdidaut = ? AND ppdidazi = ? AND ppddtvia = ? " & _
                          "                    AND ppdflstr = 'V' AND ppdidppd <> ? "
    SQL_APPLICA_MINIMO = "UPDATE prefppd0 SET ppdvalti = ?, ppdimpvi = ?, " & _
                         "                    ppdflano = ppdflano + 4 " & _
                         " WHERE ppdidppt = ? AND ppdidaut = ? AND ppdidazi = ? AND ppddtvia = ? " & _
                         "                    AND ppdflstr = 'V' AND ppdidppd = ? "
    SQL_DELETE = "DELETE FROM prefppd0 WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_DELETE_DETAIL = "DELETE FROM prefpad0 WHERE padidppt = ? AND padidppd = ?"
    SQL_ADD_EXT_DATA = "UPDATE prefppd0 SET ppdidaze = ? ,ppdvalta = ?, ppdimpva = ?, " & _
                       "                    ppdflimf = ? ,ppdqkmra = ?, ppdrife3 = ?, " & _
                       "                    ppdannoe = ? ,ppdidave = ?, ppdidtfe = '999', " & _
                       "                    ppdrife2 = ? ,ppdqtekg = ?, ppdqtelt = ?, " & _
                       "                    ppdimppe = ? ,ppdimpoe = ?, ppdimpbe = ?, " & _
                       "                    ppdimphe = ? ,ppdimpse = ?, ppdimsee = ? " & _
                       " WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_ADD_EXT_DATA_2 = "UPDATE prefppd0 SET ppdidaze = ? ,ppdvalta = ?, ppdimpva = ?, " & _
                         "                    ppdflimf = ? ,ppdqkmra = ?, ppdrife3 = ?, " & _
                         "                    ppdannoe = ? ,ppdidave = ?, ppdidtfe = '999', " & _
                         "                    ppdvalti = ? ,ppdimpvi = ?, ppdflifi = 'S', " & _
                         "                    ppdqkmri = ? ,ppdqtalt = ?, ppdflano = 0, " & _
                         "                    ppdrife2 = ? ,ppdqtekg = ?, ppdqtelt = ?, " & _
                         "                    ppdimppe = ? ,ppdimpoe = ?, ppdimpbe = ?, " & _
                         "                    ppdimphe = ? ,ppdimpse = ?, ppdimsee = ? " & _
                         " WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_ADD_EXT_DATA_3 = "UPDATE prefppd0 SET ppdidaze = ? ,ppdvalta = ?, ppdimpva = ?, " & _
                       "                    ppdflimf = ? ,ppdqkmra = ?, ppdrife3 = ?, " & _
                       "                    ppdannoe = ? ,ppdidave = ?, ppdidtfe = '999', " & _
                       "                    ppdrife2 = ? ,ppdqtekg = ?, ppdqtelt = ?" & _
                       " WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_GET_ANOMALIA = "SELECT ppdflano FROM prefppd0 " & _
                       " WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_SET_ANOMALIA = "UPDATE prefppd0 SET ppdflano = ? " & _
                       " WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_SET_TRASM = "UPDATE prefppd0 SET ppdfltrc = ? " & _
                       " WHERE ppdidppt = ? AND ppdidppd = ?"
    SQL_GET_PEDAGGI_SOC = "SELECT SUM(gp1imppe) " & _
                          "  FROM prefppd0 INNER JOIN prefgp10 ON ppdidppt = gp1id"
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Prefatturazione", "prefppd0.ppdidppt", 0, True, 7, "0", , True
        .Add "Numero Dettaglio", "prefppd0.ppdidppd", 1, True, 7, "1", , True
        .Add "Ordine cliente", "prefppd0.ppdidoct", 0, True, 7, "2", , True
        .Add "Riga Ordine cliente", "prefppd0.ppdidocr", 1, True, 3, "3", , True
        .Add "Data invio ordine", "prefppd0.ppddtior", 2, False, 10, "4", , True
        .Add "Data richiesta consegna", "prefppd0.ppddtrco", 2, False, 10, "5", , True
        .Add "Riferimento esterno 1", "prefppd0.ppdrife1", 0, False, 15, "6", , True
        .Add "Riferimento esterno 2", "prefppd0.ppdrife2", 0, False, 15, "7", , True
        .Add "Q.ta Ordine iniziale Kg", "prefppd0.ppdqtikg", 1, False, 13, "8", , True
        .Add "Q.ta Ordine iniziale Lt", "prefppd0.ppdqtilt", 1, False, 13, "9", , True
        .Add "Q.ta Ordine iniziale Pz", "prefppd0.ppdqtipz", 1, False, 13, "10", , True
        .Add "Q.ta Ordine iniziale Cnf", "prefppd0.ppdqticn", 1, False, 13, "11", , True
        .Add "Numero Viaggio", "prefppd0.ppdidvvt", 0, True, 7, "12", , True
        .Add "Numero Consegna", "prefppd0.ppdidvvc", 0, True, 3, "13", , True
        .Add "Numero Scarico", "prefppd0.ppdidvvp", 0, True, 3, "14", , True
        .Add "Data viaggio", "prefppd0.ppddtvia", 2, False, 10, "15", , True
        .Add "Q.ta Trasportata Kg", "prefppd0.ppdqttkg", 1, False, 13, "16", , True
        .Add "Q.ta Trasportata Lt", "prefppd0.ppdqttlt", 1, False, 13, "17", , True
        .Add "Q.ta Trasportata Pz", "prefppd0.ppdqttpz", 1, False, 13, "18", , True
        .Add "Q.ta Trasportata Cnf", "prefppd0.ppdqttcn", 1, False, 13, "19", , True
        .Add "Societ‡", "prefppd0.ppdidasc", 0, True, 3, "20", , True
        With .Item("20")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "prefppd0.ppdidaee", 0, True, 3, "21", , True
        With .Item("21")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Deposito", "prefppd0.ppdidade", 0, True, 3, "22", , True
        With .Item("22")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Cliente", "prefppd0.ppdidacl", 0, True, 10, "23", , True
        With .Item("23")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Destinazione", "prefppd0.ppdidadz", 0, True, 10, "24", , True
        With .Item("24")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Rif. Esterno Destinazione", "prefppd0.ppdrifde", 0, False, 15, "25", , True
        .Add "Prodotto", "prefppd0.ppdidapr", 0, True, 6, "26", , True
        With .Item("26")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "prefppd0.ppdidaim", 0, True, 3, "27", , True
        With .Item("27")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "prefppd0.ppdidapd", 0, True, 3, "28", , True
        With .Item("28")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Q.ta Scaricata Kg", "prefppd0.ppdqtakg", 1, False, 13, "29", , True
        .Add "Q.ta Scaricata Lt", "prefppd0.ppdqtalt", 1, False, 13, "30", , False
        .Add "Q.ta Scaricata Pz", "prefppd0.ppdqtapz", 1, False, 13, "31", , True
        .Add "Q.ta Scaricata Cnf", "prefppd0.ppdqtacn", 1, False, 13, "32", , True
        .Add "Q.ta Esterna Kg", "prefppd0.ppdqtekg", 1, False, 13, "59", , True
        .Add "Q.ta Esterna Lt", "prefppd0.ppdqtelt", 1, False, 13, "60", , True
        .Add "Q.ta Esterna Pz", "prefppd0.ppdqtepz", 1, False, 13, "61", , True
        .Add "Q.ta Esterna Cnf", "prefppd0.ppdqtecn", 1, False, 13, "62", , True
        .Add "Zona Fascia Tariffaria Esterna", "prefppd0.ppdidaze", 0, True, 3, "33", , True
        With .Item("33")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Fascia Tariffaria Esterna", "prefppd0.ppdidtfe", 0, True, 3, "34", , True
        .Add "Valore Tariffa Esterna", "prefppd0.ppdvalta", 1, True, 11, "35", , True
        .Add "Importo valorizzato Esterno", "prefppd0.ppdimpva", 1, True, 11, "36", , True
        .Add "Flag Importo forzato Esterno", "prefppd0.ppdflimf", 0, False, 1, "37", , True
        .Add "Valuta Importo Esterno", "prefppd0.ppdidave", 0, True, 3, "38", , True
        With .Item("38")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Km Raggio Esterno", "prefppd0.ppdqkmra", 1, True, 6, "39", , True
        .Add "Tipo Percorso Esterno", "prefppd0.ppdfltpe", 0, False, 1, "40", , True
        .Add "Importo Pedaggio Esterno", "prefppd0.ppdimppe", 1, False, 11, "66", , True
        .Add "Importo Servizio Operario Esterno", "prefppd0.ppdimpoe", 1, False, 11, "73", , True
        .Add "Importo Servizio Buques Esterno", "prefppd0.ppdimpbe", 1, False, 11, "74", , True
        .Add "Ore Servizi Speciali Esterno", "prefppd0.ppdimphe", 1, False, 6, "75", , True
        .Add "Importo Servizi Speciali Esterno", "prefppd0.ppdimpse", 1, False, 11, "76", , True
        
        .Add "Zona Fascia Tariffaria Interna", "prefppd0.ppdidazi", 0, True, 3, "41"
        With .Item("41")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Fascia Tariffaria Interna", "prefppd0.ppdidtfi", 0, True, 3, "42"
        .Add "Valore Tariffa Interna", "prefppd0.ppdvalti", 1, True, 11, "43"
        .Add "Importo valorizzato Interno", "prefppd0.ppdimpvi", 1, True, 11, "44"
        .Add "Flag Importo forzato Interno", "prefppd0.ppdflifi", 0, False, 1, "45", , True
        .Add "Valuta Importo Interno", "prefppd0.ppdidavi", 0, True, 3, "46"
        With .Item("46")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Km Raggio Interno", "prefppd0.ppdqkmri", 1, True, 6, "47"
        .Add "Tipo Percorso Interno", "prefppd0.ppdfltpi", 0, True, 1, "48"
        .Add "Importo Pedaggio Interno", "prefppd0.ppdimppi", 1, False, 11, "67" ', , True
        .Add "Importo Servizio Operario Interno", "prefppd0.ppdimpoi", 1, False, 11, "69" ', , True
        .Add "Importo Servizio Buques Interno", "prefppd0.ppdimpbi", 1, False, 11, "70" ', , True
        .Add "Ore Servizi Speciali Interno", "prefppd0.ppdimphi", 1, False, 6, "71" ', , True
        .Add "Importo Servizi Speciali Interno", "prefppd0.ppdimpsi", 1, False, 11, "72" ', , True
        
        .Add "Stato Riga", "prefppd0.ppdflstr", 0, True, 1, "49", , True
        .Add "Anomalia", "prefppd0.ppdflano", 1, False, 5, "50"
        .Add "Flag Stato 2", "prefppd0.ppdflst2", 0, False, 1, "51"
        .Add "Flag Stato 3", "prefppd0.ppdflst3", 0, False, 1, "52"
        .Add "Flag Stato 4", "prefppd0.ppdflst4", 0, False, 1, "53"
        .Add "Flag Stato 5", "prefppd0.ppdflst5", 0, False, 1, "54"
        .Add "Unit‡ di trasporto", "prefppd0.ppdidaut", 0, True, 3, "55", , False
        With .Item("55")
            .ForeignTableEntity = "AUT"
            .ForeignTable = "anagaut0"
            .ForeignTableFld = "autidaut"
            .ForeignTableDescField = "autdescr"
            .ForeignTableDeleteFld = "autannul"
        End With
        .Add "Riferimento esterno 3", "prefppd0.ppdrife3", 0, False, 15, "57", , True
        .Add "Anno prefattura", "prefppd0.ppdannoe", 1, False, 4, "58", , True
        .Add "Contratto", "prefppd0.ppdidacn", 0, False, 3, "63"
        .Add "Fascia tariffaria contratto", "prefppd0.ppdidtfc", 0, False, 7, "64"
        ' (PM) 08/02/06
        .Add "Trasmesso", "prefppd0.ppdfltrc", 0, , 2, "65"
        
        .Add "Data Grado Servizio", "prefppd0.ppddtgrs", 2, , 10, "68", , True
        
        '-- DP 20/08/08
        .Add "Servicios Exentos Interno", "prefppd0.ppdimsei", 1, False, 13, "77"
        .Add "Servicios Exentos Esterno", "prefppd0.ppdimsee", 1, False, 13, "78"
        
        .Add "Ragione sociale destinazione", "anagadz0.adzrasoc", 0, False, 35, "56", , True
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
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


Public Function SetTrasmesso(varIDPrefattura As Variant, _
            varIDRiga As Variant, _
            strFlagTrasmessso As String) As Boolean
On Error GoTo ErrHandler
    
    If Not DS.DA.ExecuteCommand(SQL_SET_TRASM, strFlagTrasmessso, varIDPrefattura, varIDRiga) Then
        GoTo ErrHandler
    End If
    
    SetTrasmesso = True
    Exit Function

ErrHandler:
    SetTrasmesso = False
End Function
