VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefattura"
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
Public SQL_DELETE As String
Public SQL_DELETE_DETAIL As String
Public SQL_DELETE_DETAIL_ADDEBITI As String
Public SQL_RESET_TRIP_STATE As String
Public SQL_SET_IMP As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Public Function Delete(Optional IDTestata As String) As Boolean
    On Error GoTo ErrHandler
    Dim strIDTestata As String
    Dim strTipoPrefattura As String
    
    If Len(IDTestata) = 0 Then
        strIDTestata = DS.Fields("0").Value
        strTipoPrefattura = UCase(DS.Field("Tipo prefattura").Value)
    Else
        strIDTestata = IDTestata
    End If
    
    With DS.DA
        .BeginTransaction
        
        If strTipoPrefattura = "A" Then
            If Not .ExecuteCommand(Me.SQL_RESET_TRIP_STATE & DS.DA.MakeSQL(strIDTestata, dtAlfa) & ")") Then
                GoTo ErrHandler
            End If
        End If
        
        If Not .ExecuteCommand(Me.SQL_DELETE_DETAIL_ADDEBITI, strIDTestata) Then
            GoTo ErrHandler
        End If
        
        If Not .ExecuteCommand(Me.SQL_DELETE_DETAIL, strIDTestata) Then
            GoTo ErrHandler
        End If
        
        If Not .ExecuteCommand(Me.SQL_DELETE, strIDTestata) Then
            GoTo ErrHandler
        End If
        
        .EndTrans etCommit
    End With
    
    Delete = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    Delete = False
End Function


Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    DS.Field("Stato").NewValue = "G"
    DS.Field("Stato").Updated = True
    Insert = DS.Insert(Me.SQL_INSERT, "anaganu0", "anuprogr", "anuidanu", "005", 0)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean)
    
    BeginSave = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
End Function
Public Function SetImporti() As Boolean
    On Error GoTo ErrHandler
    
    If Not DS.DA.ExecuteCommand(SQL_SET_IMP, DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If
    
    SetImporti = True
    Exit Function
ErrHandler:
    SetImporti = False
End Function

Public Function CreateDocumento(DataDoc As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strQGetContratti As String
    Dim strQGetAddebiti_Deposito As String
    Dim strQGetAddebiti_Voce As String
    Dim strQSetRifAddebito As String
    Dim strQSetRifDocumento As String
    Dim strQSetRifDocumento_Voce As String
    Dim strQSetRifDocumento_Deposito As String
    
    Dim strIDPrefattura As String
    Dim strIDCliente As String
    Dim strIDContratto As String
    Dim strIDSocieta As String
    Dim strIDEnteEmittente As String
    Dim strIDDocumento As String
    Dim strIDIVA As String
    Dim strIDAddebito As String
    Dim strIDDeposito As String
    Dim strDescrDeposito As String
    Dim strIDDepositoPrec As String
    Dim strIDAddebitoPrec As String
    Dim strDescrizioneRiga As String
    Dim strFlagRotturaFattura As String
    
    Dim varBufContratti As Variant
    Dim varBufAddebiti As Variant
    Dim varIDIVACliente As Variant
    Dim varAnnoDoc As Variant
    Dim varNumeroDoc As Variant
    Dim varQuantita As Variant
    Dim varPrezzo As Variant
    Dim varTotConsegne As Variant
    Dim varTotKm As Variant
    Dim varTotLitri As Variant
    Dim varIDAUT As Variant
    Dim varIDAAM1 As Variant
    Dim varIDAAM2 As Variant
    Dim varDallaData As Variant
    Dim varAllaData As Variant
    
    Dim lngProgrRigaDoc As Long
    Dim lngHowMany As Long
    
    Dim blnRotturaFattura As Boolean
    
    
    Dim i As Long
    Dim j As Long
    
    Dim DAM As CDocumentoAmministrativo
    
    strQGetContratti = "SELECT DISTINCT PADIDACL, PADIDACN, AMFFLROF, AMFFLOR1, " & _
                       "       AMFFLOR2, AMFFLOR3, PPTIDASC " & _
                       "  FROM PREFPAD0 INNER JOIN ANAGACN0 ON PADIDACL = ACNIDACL AND PADIDACN = ACNIDACN " & _
                       "       INNER JOIN ANAGAMF0 ON ACNIDAMF = AMFIDAMF INNER JOIN PREFPPT0 ON PPTIDPPT = PADIDPPT" & _
                       " WHERE PADIDPPT = ? " & _
                       " ORDER BY PADIDACL, PADIDACN "
    
    strQGetAddebiti_Deposito = "SELECT PPDIDAUT, PADIDADD, ADDDESCR, ADDIDAIV, " & _
                      "       SUM(PADIMIMP), SUM(PADNRCON), SUM(PADNRKMS), SUM(PADNRLTS), " & _
                      "       AUTIDAM1, AUTIDAM2, PPDIDADE, ADEDESCR, PPTDTINP, PPTDTFIP " & _
                      "  FROM PREFPAD0 INNER JOIN ANAGADD0 ON PADIDADD = ADDIDADD " & _
                      "       INNER JOIN PREFPPD0 ON PADIDPPT = PPDIDPPT AND PADIDPPD = PPDIDPPD " & _
                      "       INNER JOIN ANAGAUT0 ON AUTIDAUT = PPDIDAUT " & _
                      "       INNER JOIN ANAGADE0 ON PPDIDADE = ADEIDADE INNER JOIN PREFPPT0 ON PPTIDPPT = PPDIDPPT " & _
                      " WHERE PADIDPPT = ? AND PADIDACL = ? AND PADIDACN = ? " & _
                      " GROUP BY PPDIDADE, ADEDESCR, PPDIDAUT, PPTDTINP, PPTDTFIP, PADIDADD, ADDDESCR, ADDIDAIV, AUTIDAM1, AUTIDAM2 ORDER BY PPDIDADE "
                          
    strQGetAddebiti_Voce = "SELECT PPDIDAUT, PADIDADD, ADDDESCR, ADDIDAIV, " & _
                      "       SUM(PADIMIMP), SUM(PADNRCON), SUM(PADNRKMS), SUM(PADNRLTS), " & _
                      "       AUTIDAM1, AUTIDAM2, '', '', PPTDTINP, PPTDTFIP " & _
                      "  FROM PREFPAD0 INNER JOIN ANAGADD0 ON PADIDADD = ADDIDADD " & _
                      "       INNER JOIN PREFPPD0 ON PADIDPPT = PPDIDPPT AND PADIDPPD = PPDIDPPD " & _
                      "       INNER JOIN ANAGAUT0 ON AUTIDAUT = PPDIDAUT " & _
                      "       INNER JOIN ANAGADE0 ON PPDIDADE = ADEIDADE INNER JOIN PREFPPT0 ON PPTIDPPT = PPDIDPPT " & _
                      " WHERE PADIDPPT = ? AND PADIDACL = ? AND PADIDACN = ? " & _
                      " GROUP BY PADIDADD, ADDDESCR, PPDIDAUT, PPTDTINP, PPTDTFIP, ADDIDAIV, AUTIDAM1, AUTIDAM2 ORDER BY PADIDADD "
    
    strQSetRifAddebito = "INSERT INTO DOCUDAD0 (DADIDAEE, DADIDAND, DADIDDAM, DADIDDAR, " & _
                         "                      DADIDADD, DADIDAUT, DADIDAM1, DADIDAM2, " & _
                         "                      DADNRCON, DADNRKMS, DADNRLTS, DADIDPPT, " & _
                         "                      DADIDACL, DADIDACN) " & _
                         "     VALUES (? , ? , ? , ?, " & _
                         "       ? , ? , ? , ?, " & _
                         "       ? , ? , ? , ?, " & _
                         "       ? , ?) "
    strQSetRifDocumento_Voce = "UPDATE PREFPAD0 " & _
                               "   SET PADIDAEE = ? , PADIDAND = ? , PADIDDAM = ?  " & _
                               " WHERE PADIDPPT = ? AND PADIDACL = ? AND PADIDACN = ? AND PADIDADD = ? "
    strQSetRifDocumento_Deposito = "UPDATE PREFPAD0 " & _
                                   "   SET PADIDAEE = ? , PADIDAND = ? , PADIDDAM = ?  " & _
                                   " WHERE PADIDPPT = ? AND PADIDACL = ? AND PADIDACN = ? AND PADIDPPD IN " & _
                                   " (SELECT PPDIDPPD FROM PREFPPD0 WHERE PPDIDPPT = "
                                   
    strIDPrefattura = DS.Fields("0").Value
    
    ' Estraggo i contratti presenti nelle righe di prefattura
    If Not DS.DA.QueryCommand(strQGetContratti, varBufContratti, lngHowMany, _
                              strIDPrefattura) Then
        GoTo ErrHandler
    End If
    
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    DS.DA.BeginTransaction
    ' Elaboro ogni contratto
    For i = 0 To UBound(varBufContratti, 2)
        strIDCliente = varBufContratti(0, i)
        strIDContratto = varBufContratti(1, i)
        strFlagRotturaFattura = varBufContratti(2, i)
        strIDSocieta = varBufContratti(6, i)
        
        ' Seleziono le righe di addebito per il contratto che sto elaborando,
        ' raggruppando secondo le impostazioni del contratto
        lngHowMany = 0
        
        Select Case strFlagRotturaFattura
            Case "00"
                strQ = strQGetAddebiti_Deposito
            Case "01"
                strQ = strQGetAddebiti_Voce
            Case "02"
                strQ = strQGetAddebiti_Deposito
        End Select
        
        If Not DS.DA.QueryCommand(strQ, varBufAddebiti, lngHowMany, _
                                  strIDPrefattura, _
                                  strIDCliente, _
                                  strIDContratto) Then
            GoTo ErrHandler
        End If
        
        ' Inserisco la testata di una nuova fattura
        If lngHowMany > 0 Then
            
            lngProgrRigaDoc = 1
            strIDDepositoPrec = vbNullString
            
            For j = 0 To UBound(varBufAddebiti, 2)
                strIDAddebito = varBufAddebiti(1, j)
                strIDDeposito = varBufAddebiti(10, j)
                strDescrDeposito = varBufAddebiti(11, j)
                varDallaData = varBufAddebiti(12, j)
                varAllaData = varBufAddebiti(13, j)
                
                ' Inseriamo un nuovo documento se siamo all'inizio dell'elaborazione oppure
                ' se il contratto prevede la rottura della fattura
                
                blnRotturaFattura = (j = 0)
                Select Case strFlagRotturaFattura
                    Case "01"
                        blnRotturaFattura = blnRotturaFattura Or (strIDAddebito <> strIDAddebitoPrec)
                        strQSetRifDocumento = strQSetRifDocumento_Voce
                    Case "00", "02"
                        blnRotturaFattura = blnRotturaFattura Or (strIDDeposito <> strIDDepositoPrec)
                        strQSetRifDocumento = strQSetRifDocumento_Deposito & DS.DA.MakeSQL(strIDPrefattura, dtAlfa) & " AND ppdidade = " & DS.DA.MakeSQL(strIDDeposito, dtAlfa) & ")"
                    Case Else
                End Select
                
                If blnRotturaFattura Then
                    
                    ' Completo la fattura precedente
                    If (j <> 0) Then
                        If Not DAM.ComputeTotali Then
                            GoTo ErrHandler
                        End If
                    End If
                    
                    If Not CreateDocumentoTestata(strIDCliente, Format(DataDoc, "dd/mm/yy"), DAM, varIDIVACliente, strIDSocieta) Then
                        GoTo ErrHandler
                    End If
                    
                    strIDEnteEmittente = DAM.DS.Field("Ente Emittente").Value
                    varAnnoDoc = DAM.DS.Field("Anno Documento").Value
                    strIDDocumento = DAM.DS.Field("Progressivo Documento").Value
                    
                    ' Lego le righe di addebito alla riga di fattura appena creata
                    Select Case strFlagRotturaFattura
                        Case "01"
                            If Not DS.DA.ExecuteCommand(strQSetRifDocumento, _
                                                        strIDEnteEmittente, varAnnoDoc, strIDDocumento, _
                                                        strIDPrefattura, strIDCliente, strIDContratto, strIDAddebito) Then
                                GoTo ErrHandler
                            End If
                        Case "00", "02"
                            If Not DS.DA.ExecuteCommand(strQSetRifDocumento, _
                                                        strIDEnteEmittente, varAnnoDoc, strIDDocumento, _
                                                        strIDPrefattura, strIDCliente, strIDContratto) Then
                                GoTo ErrHandler
                            End If
                    End Select
                    
                    If strFlagRotturaFattura = "02" Then
                        strDescrizioneRiga = "Fact. " & strDescrDeposito & Space(5) & "Desde el " & Format(varDallaData, "dd/mm/yy") & " Hasta el " & Format(varAllaData, "dd/mm/yy")
                        If Not CreateDocumentoDettaglio(strIDEnteEmittente, varAnnoDoc, strIDDocumento, _
                                                        lngProgrRigaDoc, strDescrizioneRiga, "T") Then
                            GoTo ErrHandler
                        End If
                        lngProgrRigaDoc = lngProgrRigaDoc + 1
                    End If
                    
                    strIDDepositoPrec = strIDDeposito
                    strIDAddebitoPrec = strIDAddebito
                End If
                
                varIDAUT = varBufAddebiti(0, j)
                strDescrizioneRiga = varIDAUT & " - " & varBufAddebiti(2, j)
                ' Gestisco l'eventuale esenzione IVA del cliente
                If Len(varIDIVACliente) > 0 Then
                    strIDIVA = varIDIVACliente
                Else
                    strIDIVA = varBufAddebiti(3, j)
                End If
                
                varQuantita = varBufAddebiti(4, j)
                varTotConsegne = varBufAddebiti(5, j)
                varTotKm = varBufAddebiti(6, j)
                varTotLitri = varBufAddebiti(7, j)
                varIDAAM1 = varBufAddebiti(8, j)
                varIDAAM2 = varBufAddebiti(9, j)
                
                varPrezzo = 1
                
                ' Inserisco una nuova riga di dettaglio in fattura
                If Not CreateDocumentoDettaglio(strIDEnteEmittente, varAnnoDoc, strIDDocumento, _
                                                lngProgrRigaDoc, strDescrizioneRiga, "N", _
                                                varQuantita, varPrezzo, strIDIVA, strIDAddebito) Then
                    GoTo ErrHandler
                End If
                
                ' Inserisco sulle righe di addebito il riferimento alla riga di fattura
                ' appena creata
                If Not DS.DA.ExecuteCommand(strQSetRifAddebito, _
                                            strIDEnteEmittente, varAnnoDoc, strIDDocumento, lngProgrRigaDoc, _
                                            strIDAddebito, varIDAUT, varIDAAM1, varIDAAM2, _
                                            varTotConsegne, varTotKm, varTotLitri, strIDPrefattura, _
                                            strIDCliente, strIDContratto) Then
                    GoTo ErrHandler
                End If
                
                lngProgrRigaDoc = lngProgrRigaDoc + 1
                
            Next j
            
            If Not DAM.ComputeTotali Then
                GoTo ErrHandler
            End If
        End If
        
    Next i
    
    ' (PM) 15/03/06 - Lo stato della prefattura diventa "E", Elaborato
    If Not SetStato("E") Then
        GoTo ErrHandler
    End If

    DS.DA.EndTrans etCommit
    CreateDocumento = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    CreateDocumento = False
End Function


Public Function CreateDocumentoTestata(IDCliente As String, DataDoc As Variant, _
                                       ByRef DAM As CDocumentoAmministrativo, _
                                       ByRef IDIVACliente, strIDSocieta As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDAMPCliente As String
    Dim strIDAVADef As String
    Dim strErrDesc As String
    Dim varNumeroDoc As Variant
    
    Dim ACL As CCliente
    Dim AXX As CParametro
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With

    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = "001"
        AXX.Find
        strIDAVADef = Trim$(.Fields("17").Value)
    End With
    
    ' Ricerco il cliente
    With ACL.DS
        .Fields.ResetFields
        .Fields("0").FindValue = IDCliente
        If Not ACL.Find Then
            strErrDesc = "Errore nella ricerca del cliente"
            GoTo ErrHandler
        End If
        If ACL.DS.HowMany = 0 Then
            strErrDesc = "Il cliente non è stato trovato"
            GoTo ErrHandler
        End If
        IDIVACliente = .Fields("15").Value
        strIDAMPCliente = .Field("Modalità di Pagamento").Value
    End With
    
    
    DS.DA.BeginTransaction
    
    ' Iniziamo l'inserimento di una nuova fattura
    ' (PM) 08/03/06 - Il numero della fattura viene valorizzato quando diventa definitiva
'    varNumeroDoc = DS.DA.GetCounter("docunda0", "ndanrprg", _
'                                    "ndaidasc", "001", 0, _
'                                    "ndaidaee", "056", 0, _
'                                    "ndaidato", "FAT", 0)
'    If varNumeroDoc = -1 Then
'        strErrDesc = "Impossibile leggere il numeratore per il documento."
'        GoTo ErrHandler
'    End If
    
    With DAM.DS
        .Fields.ResetFields
        .Field("Ente Emittente").NewValue = "001"
        .Field("Anno Documento").NewValue = Year(DataDoc)
        .Field("Società").NewValue = strIDSocieta
        .Field("Cliente").NewValue = ACL.DS.Field("Codice").Value
        .Field("Data Documento").NewValue = DataDoc
        .Field("Tipo Documento Amministrativo").NewValue = "FAT"
        ' (PM) 08/03/06 - Il numero della fattura viene valorizzato quando diventa definitiva
'        .Field("Numero Documento Amministrativo").NewValue = varNumeroDoc
        .Field("Numero Documento Amministrativo").NewValue = "PROV"
        .Field("Valuta").NewValue = strIDAVADef
        .Field("Modalità di pagamento").NewValue = strIDAMPCliente
        .Field("Rag.Soc. Cliente").NewValue = ACL.DS.Fields("1").Value
        .Field("Rag.Soc. Agg. Cliente").NewValue = ACL.DS.Fields("2").Value
        .Field("Indir. Cliente").NewValue = ACL.DS.Fields("4").Value
        .Field("Località Cliente").NewValue = ACL.DS.Fields("5").Value
        .Field("CAP Cliente").NewValue = ACL.DS.Fields("7").Value
        .Field("Provincia Cliente").NewValue = ACL.DS.Fields("6").Value
        .Field("Cod. Fis. Cliente").NewValue = ACL.DS.Fields("10").Value
        .Field("P. IVA Cliente").NewValue = ACL.DS.Fields("11").Value
        .Field("Flag Gestione").NewValue = "E"
        .Field("Flag tipo fatturazione").NewValue = "A"
        .Field("IVA Documento").NewValue = 0
        .Field("Importo Documento").NewValue = 0
        ' (PM) 08/03/06 - Stato della fattura e' Generata (Provvisorio)
        .Field("Esito Documento").NewValue = "G"
        
        .Field("Ente Emittente").Updated = True
        .Field("Anno Documento").Updated = True
        .Field("Società").Updated = True
        .Field("Cliente").Updated = True
        .Field("Data Documento").Updated = True
        .Field("Tipo Documento Amministrativo").Updated = True
        .Field("Numero Documento Amministrativo").Updated = True
        .Field("Valuta").Updated = True
        .Field("Modalità di pagamento").Updated = True
        .Field("Rag.Soc. Cliente").Updated = True
        .Field("Rag.Soc. Agg. Cliente").Updated = True
        .Field("Indir. Cliente").Updated = True
        .Field("Località Cliente").Updated = True
        .Field("CAP Cliente").Updated = True
        .Field("Provincia Cliente").Updated = True
        .Field("Cod. Fis. Cliente").Updated = True
        .Field("P. IVA Cliente").Updated = True
        .Field("Flag Gestione").Updated = True
        .Field("Flag tipo fatturazione").Updated = True
        .Field("IVA Documento").Updated = True
        .Field("Importo Documento").Updated = True
        
        ' (PM) 08/03/06
        .Field("Esito Documento").Updated = True
    End With
    
    If Not DAM.Insert Then
        strErrDesc = "Errore nell'inserimento del documento"
        GoTo ErrHandler
    End If
        
    DS.DA.EndTrans etCommit
    CreateDocumentoTestata = True
    Exit Function

ErrHandler:
    DS.DA.EndTrans etRollback
    CreateDocumentoTestata = False
End Function
Public Function CreateDocumentoIVA(IDCliente As String, DataDoc As Variant, _
                                       ByRef IDEnteEmittente As Variant, _
                                       ByRef AnnoDocumento As Variant, _
                                       ByRef IDDocumento As Variant, _
                                       ByRef NumeroDocumento As Variant, _
                                       ByRef IDIVACliente As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDAMPCliente As String
    Dim strIDAVADef As String
    Dim strErrDesc As String
    Dim varNumeroDoc As Variant
    
    Dim ACL As CCliente
    Dim DAM As CDocumentoAmministrativo
    Dim AXX As CParametro
    
    Set ACL = New CCliente
    With ACL.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    DS.DA.EndTrans etCommit
    CreateDocumentoIVA = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    CreateDocumentoIVA = False
End Function

Public Function CreateDocumentoDettaglio(IDEnteEmittente As String, AnnoDocumento As Variant, IDDocumento As String, _
                                         ByRef ProgrRigaDoc As Long, _
                                         Descrizione As String, _
                                         TipoRiga As String, _
                                         Optional Quantita As Variant, _
                                         Optional PrezzoUnitario As Variant, _
                                         Optional IDIVA As String, _
                                         Optional IDAddebito As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim DAR As CRigaDocumentoAmministrativo
    Dim dblPrezzoUnitario As Double
    Dim dblImponibile As Double
    
    Dim varQta As Variant
    Dim strErrDesc As String
    
    Set DAR = New CRigaDocumentoAmministrativo
    With DAR.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    With DAR.DS
        .Fields.ResetFields
        .Field("Ente Emittente").NewValue = IDEnteEmittente
        .Field("Anno Documento").NewValue = AnnoDocumento
        .Field("Progressivo Documento").NewValue = IDDocumento
        .Field("Numero Riga Documento").NewValue = Format(ProgrRigaDoc, "000")
        .Field("Flag Tipo Riga").NewValue = TipoRiga
        .Field("Ente Emittente").Updated = True
        .Field("Anno Documento").Updated = True
        .Field("Progressivo Documento").Updated = True
        .Field("Numero Riga Documento").Updated = True
        .Field("Flag Tipo Riga").Updated = True
        
        Select Case TipoRiga
            Case "N"
                .Field("Quantità Pz").NewValue = Quantita
                .Field("Unità di Misura").NewValue = "PZ"
                .Field("Prezzo Unitario").NewValue = PrezzoUnitario
                .Field("Aliquota IVA").NewValue = IDIVA
                
                dblPrezzoUnitario = PrezzoUnitario
                Select Case .Field("Unità di Misura").NewValue
                    Case "KG"
                        varQta = .Field("Quantità Kg").NewValue
                    Case "L"
                        varQta = .Field("Quantità L").NewValue
                    Case "PZ"
                        varQta = .Field("Quantità Pz").NewValue
                    Case "CNF"
                        varQta = .Field("Quantità Cnf").NewValue
                End Select
                dblImponibile = Round(dblPrezzoUnitario * (varQta), 2)
                
                .Field("Imponibile").NewValue = dblImponibile
                
                .Field("Addebito").NewValue = IDAddebito
                .Field("Descrizione Riga Documento").NewValue = Left(Descrizione, .Field("Descrizione Riga Documento").MaxLength)
                
                .Field("Quantità Pz").Updated = True
                .Field("Unità di Misura").Updated = True
                .Field("Prezzo Unitario").Updated = True
                .Field("Aliquota IVA").Updated = True
                .Field("Imponibile").Updated = True
                .Field("Addebito").Updated = True
                .Field("Descrizione Riga Documento").Updated = True
            Case "T"
                .Field("Descrizione Riga Documento").NewValue = Left(Descrizione, .Field("Descrizione Riga Documento").MaxLength)
                .Field("Descrizione Riga Documento").Updated = True
        End Select
    End With
    If Not DAR.Insert Then
        strErrDesc = "Errore nell'inserimento della riga"
        GoTo ErrHandler
    End If
    
    CreateDocumentoDettaglio = True
    Exit Function
ErrHandler:
    CreateDocumentoDettaglio = False
    
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT pptidppt, pptdtinp, pptdtfip, pptdescr, pptfltip, pptidasc, pptflsta, timestmp FROM prefppt0 WHERE "
    SQL_GETLIST = "SELECT pptidppt, pptdescr FROM prefppt0 "
    SQL_UPDATE = "UPDATE prefppt0 SET "
    SQL_INSERT = "INSERT INTO prefppt0 "
    SQL_GET_TS = "SELECT timestmp FROM prefppt0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT pptidppt, pptdtinp, pptdtfip, pptdescr " & _
                          "  FROM prefppt0 "
    
    SQL_DELETE = "DELETE FROM prefppt0 WHERE pptidppt = ? "
    SQL_DELETE_DETAIL = "DELETE FROM prefppd0 WHERE ppdidppt = ? "
    SQL_DELETE_DETAIL_ADDEBITI = "DELETE FROM prefpad0 WHERE padidppt = ? "
'    SQL_RESET_TRIP_STATE = "UPDATE PREFPPD0, VIAGVVT0 SET VIAGVVT0.VVTFLSTV = 'C' " & _
'                           " WHERE (((PREFPPD0.PPDIDPPT)=?) AND ((PREFPPD0.PPDIDVVT)=[VVTIDVVT]))"
    SQL_RESET_TRIP_STATE = "UPDATE VIAGVVT0 SET VIAGVVT0.VVTFLSTV = 'C' " & _
                           " WHERE VVTIDVVT IN (SELECT PPDIDVVT FROM PREFPPD0 WHERE PPDIDPPT="
    SQL_SET_IMP = "UPDATE PREFPPD0 SET ppdimpvi = ppdimpva, ppdflifi = 'S', ppdflano = ppdflano - 32 " & _
                  " WHERE ppdidppt = ? AND ppdimpva > ppdimpvi"
    
    mvarFieldsIndexes = Array("0", "1", "2", "3")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "prefppt0.pptidppt", 0, True, 7, "0"
        .Add "Data inizio periodo", "prefppt0.pptdtinp", 2, True, 10, "1", , True
        .Add "Data fine periodo", "prefppt0.pptdtfip", 2, True, 10, "2", , True
        .Add "Descrizione", "prefppt0.pptdescr", 0, True, 35, "3"
        .Add "Tipo prefattura", "prefppt0.pptfltip", 0, True, 1, "4"
        .Add "Società", "prefppt0.pptidasc", 0, True, 3, "5"
        With .Item("5")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With

        ' (PM) 08/03/06
        .Add "Stato", "prefppt0.pptflsta", 0, , 2, "6"
    End With
    With DS.KeyFields
        .Add "0", "1"
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

Public Function SetStato(varStato As Variant, Optional ByRef strErrorMessage As String) As Boolean
On Error GoTo ErrHandler

    Dim blnConcurErr As Boolean
    Dim blnResult As Boolean
    Dim strErrMessage As String

    DS.DA.BeginTransaction

    SetStato = True
    With DS
        Select Case varStato
            Case "G"    ' Generato
                If DS.Field("Stato").Value <> "E" Then
                    strErrMessage = "Non è possibile modificare lo stato della prefattura a Generato."
                    GoTo ErrHandler
                    Exit Function
                End If

                ' Annullo tutte le fatture appartenenti a questa prefattura a patto che siano
                ' tutte fatture provvisorie
                If Not AnnullaFattureProvvisorie(strErrMessage) Then
                    GoTo ErrHandler
                End If

            Case "E"    ' Elaborato
            
            Case "T"
                If DS.Field("Stato").Value <> "E" Then
                    strErrMessage = "Non è possibile modificare lo stato della prefattura a Trasmesso."
                    GoTo ErrHandler
                    Exit Function
                End If
                
                If Not SetFattureStampate(strErrMessage) Then
                    GoTo ErrHandler
                End If
            
            Case Else
                strErrMessage = "Stato del documento amministrativo non gestito."
                GoTo ErrHandler
        End Select
        
        .Field("Stato").NewValue = varStato
        .Field("Stato").Updated = True
    End With

    blnResult = BeginSave(blnConcurErr)
    
    If blnResult Then
        DS.DA.EndTrans etCommit
    Else
        strErrMessage = "Errore durante la modifica dello stato della prefattura."
        GoTo ErrHandler
    End If
    
    DS.DA.EndTrans etCommit
    Exit Function

ErrHandler:
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = IIf(strErrMessage = "", "Errore sconosciuto.", strErrMessage)
    End If
    DS.DA.EndTrans etRollback
    SetStato = False

End Function

Private Function SetFattureStampate(Optional ByRef strErrorMessage As String)
On Error GoTo ErrHandler
    
    Dim SQL_SELECT_FATTURE As String
    Dim varBuffer As Variant
    Dim DAM As CDocumentoAmministrativo
    Dim blnConcurErr As Boolean
    Dim strErrMessage As String
    Dim lngHowMany As Long
    Dim lngCounter As Long
    
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    SQL_SELECT_FATTURE = "SELECT DISTINCT DOCUDAM0.DAMIDAEE, DOCUDAM0.DAMIDAND, DOCUDAM0.DAMIDDAM" & _
                            " FROM PREFPAD0 INNER JOIN" & _
                            " DOCUDAM0 ON PREFPAD0.PADIDAEE = DOCUDAM0.DAMIDAEE AND" & _
                            " PREFPAD0.PADIDAND = DOCUDAM0.DAMIDAND And PREFPAD0.PADIDDAM = DOCUDAM0.DAMIDDAM" & _
                            " WHERE DAMANNUL = " & DS.DA.MakeSQL(False, dtBit) & " AND PADIDPPT = ?"
    
    If Not DS.DA.QueryCommand(SQL_SELECT_FATTURE, varBuffer, lngHowMany, _
        DS.Field("Codice").Value) Then
        
        strErrMessage = "Errore nella selezione delle fatture."
        GoTo ErrHandler
    End If

    DS.DA.BeginTransaction
    
    For lngCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
    
        With DAM.DS
        
            .Fields.ResetFields
            .Field("Ente Emittente").FindValue = varBuffer(0, lngCounter)
            .Field("Anno Documento").FindValue = varBuffer(1, lngCounter)
            .Field("Progressivo Documento").FindValue = varBuffer(2, lngCounter)

            If Not DAM.Find() Or .HowMany = 0 Then
                strErrMessage = "Errore nella ricerca del documento amministrativo."
                GoTo ErrHandler
            End If
            
            If Not DAM.SetStato("T", strErrMessage) Then
                GoTo ErrHandler
            End If
            
        End With

    Next lngCounter
    
    DS.DA.EndTrans etCommit

    SetFattureStampate = True
    Set DAM = Nothing
    Exit Function

ErrHandler:
    DS.DA.EndTrans etRollback
    Set DAM = Nothing
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = IIf(strErrMessage = "", "Errore sconosciuto.", strErrMessage)
    End If
    SetFattureStampate = False
End Function


Private Function AnnullaFattureProvvisorie(Optional ByRef strErrorMessage As String)
On Error GoTo ErrHandler
    
    Dim SQL_SELECT_FATTURE As String
    Dim varBuffer As Variant
    Dim DAM As CDocumentoAmministrativo
    Dim blnConcurErr As Boolean
    Dim strErrMessage As String
    Dim lngHowMany As Long
    Dim lngCounter As Long
    
    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    SQL_SELECT_FATTURE = "SELECT DISTINCT DOCUDAM0.DAMIDAEE, DOCUDAM0.DAMIDAND, DOCUDAM0.DAMIDDAM" & _
                            " FROM PREFPAD0 INNER JOIN" & _
                            " DOCUDAM0 ON PREFPAD0.PADIDAEE = DOCUDAM0.DAMIDAEE AND" & _
                            " PREFPAD0.PADIDAND = DOCUDAM0.DAMIDAND And PREFPAD0.PADIDDAM = DOCUDAM0.DAMIDDAM" & _
                            " WHERE DAMANNUL = " & DS.DA.MakeSQL(False, dtBit) & " AND PADIDPPT = ?"
    
    If Not DS.DA.QueryCommand(SQL_SELECT_FATTURE, varBuffer, lngHowMany, _
        DS.Field("Codice").Value) Then
        
        strErrMessage = "Errore nella selezione delle fatture."
        GoTo ErrHandler
    End If

    DS.DA.BeginTransaction
    
    For lngCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)
    
        With DAM.DS
        
            .Fields.ResetFields
            .Field("Ente Emittente").FindValue = varBuffer(0, lngCounter)
            .Field("Anno Documento").FindValue = varBuffer(1, lngCounter)
            .Field("Progressivo Documento").FindValue = varBuffer(2, lngCounter)

            If Not DAM.Find() Or .HowMany = 0 Then
                strErrMessage = "Errore nella ricerca del documento amministrativo."
                GoTo ErrHandler
            End If
            
            If .Field("Esito Documento").Value <> "G" Then
                strErrMessage = "Impossibile annullare una fattura con lo stato diverso da Provvisorio."
                GoTo ErrHandler
            End If

            .Field("Annullato").NewValue = True
            .Field("Annullato").Updated = True

            If Not DAM.BeginSave(blnConcurErr) Or blnConcurErr Then
                strErrMessage = "Errore durante il salvataggio della prefattura."
                GoTo ErrHandler
            End If

            DS.DA.EndTrans etCommit

        End With

    Next lngCounter
    
    DS.DA.EndTrans etCommit
    Set DAM = Nothing
    AnnullaFattureProvvisorie = True
    Exit Function

ErrHandler:
    DS.DA.EndTrans etRollback
    Set DAM = Nothing
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = IIf(strErrMessage = "", "Errore sconosciuto.", strErrMessage)
    End If
    AnnullaFattureProvvisorie = False
End Function


