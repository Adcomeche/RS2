VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentoAmministrativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Classe CDocumentoAmministrativo
'Data Creazione 4/11/99
Option Explicit

Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_INSERT_IVA As String

Private Const SQL_GET_RIGHE = _
    "   SELECT DARIDAIV, SUM(DARNRIMP), AIVALIQU " & _
    "     FROM DOCUDAR0 INNER JOIN ANAGAIV0 ON DARIDAIV = AIVIDAIV " & _
    "    WHERE DARIDAND = ? AND DARIDAEE = ? AND DARIDDAM = ? AND DARANNUL = 0 " & _
    " GROUP BY DARIDAIV, AIVALIQU " & _
    " ORDER BY DARIDAIV "
    
Private Const SQL_DELETE_IVA = _
    "  DELETE FROM DOCUDAI0 " & _
    "    WHERE DAIIDAND = ? AND DAIIDAEE = ? AND DAIIDDAM = ? "
    
Private Const SQL_DAI_DETAIL = " SELECT DAIIDAEE, DAIIDAND, DAIIDDAM, DAIIDAIV, AIVDESCR, AIVALIQU, DAINRIMP, " & _
                               "       DAINRIMI FROM DOCUDAI0 INNER JOIN " & _
                               " ANAGAIV0 ON " & _
                               " DOCUDAI0.DAIIDAIV = ANAGAIV0.AIVIDAIV " & _
                               " WHERE (DOCUDAI0.DAIIDAEE = ?) AND (DOCUDAI0.DAIIDAND = ?) AND (DOCUDAI0.DAIIDDAM = ?) "

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function AnnullaDocumento(EnteEmittente As Variant, Anno As Variant, Progressivo As Variant, Optional ByRef ErrDescr As String) As Boolean
    On Error GoTo ErrHandler
    Dim strErrorMessage As String
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim blnConcErr As Boolean

' (PM) 03/01/06 - <
'    Dim DFI As CDocumentoFiscale
'
'    Set DFI = New CDocumentoFiscale
'    With DFI.DS
'        Set .DA = DS.DA
'        Set .ErrH = DS.ErrH
'    End With

    DS.Fields("0").FindValue = EnteEmittente
    DS.Fields("1").FindValue = Anno
    DS.Fields("2").FindValue = Progressivo
    If Not Me.Find Then
        strErrorMessage = "Errore durante la ricerca del documento"
        GoTo ErrHandler
    End If
    If DS.HowMany = 0 Then
        strErrorMessage = "Documento non trovato"
        GoTo ErrHandler
    End If
    varIDViaggio = DS.Fields("45").Value
    varIDConsegna = DS.Fields("46").Value
'    DS.Field("Annullato").NewValue = True
'    BeginSave blnConcErr
'    DS.DA.EndTrans etCommit
'    DS.EndSave
    DS.DA.BeginTransaction
    If Not CreateNCRFromFAT(strErrorMessage) Then
        GoTo ErrHandler
    End If
' (PM) 03/02/06 - <
'    If Not DFI.SetStatoFatturazione("S", varIDViaggio, varIDConsegna) Then
'        strErrorMessage = "Errore nell'aggiornamento del DAS"
'        GoTo ErrHandler
'    End If
    DS.DA.EndTrans etCommit
    AnnullaDocumento = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    If Not IsMissing(ErrDescr) Then
        ErrDescr = strErrorMessage
    End If
    AnnullaDocumento = False
End Function


Public Function SetStato(NewValue As String, Optional ByRef strErrorMessage As String) As Boolean
    On Error GoTo ErrHandler
    Dim blnConcurErr As Boolean
    Dim blnResult As Boolean
    Dim varNumeroDoc As Variant
    Dim strErrMessage As String
    
    SetStato = True
    With DS.Fields
        Select Case NewValue
            Case "G"    ' Generato
                If DS.Fields("10").Value <> "S" Then
                    strErrMessage = "Impossibile impostare lo stato del Doc. Amministrativo a Provvisorio."
                    GoTo ErrHandler
                End If

            Case "S"    ' Stampato
                If DS.Fields("10").Value <> "G" Then
                    strErrMessage = "Impossibile impostare lo stato del Doc. Amministrativo a Definitivo."
                    GoTo ErrHandler
                End If
            
                If DS.Field("Numero Documento Amministrativo").Value = "PROV" Or _
                    (DS.Field("Numero Documento Amministrativo").Updated And DS.Field("Numero Documento Amministrativo").NewValue = "PROV") Then
                
                    ' (PM) 08/03/06 - Il viene valorizzato il numero della fattura
                    varNumeroDoc = DS.DA.GetCounter("docunda0", "ndanrprg", _
                                                    "ndaidasc", "001", 0, _
                                                    "ndaidaee", "056", 0, _
                                                    "ndaidato", "FAT", 0)
                    If varNumeroDoc = -1 Then
                        strErrMessage = "Errore nella valorizzazione del numero della fattura."
                        GoTo ErrHandler
                    End If
                    DS.Field("Numero Documento Amministrativo").NewValue = varNumeroDoc
                    DS.Field("Numero Documento Amministrativo").Updated = True
                End If
            Case "T"    ' Trasmesso
                If DS.Fields("10").Value <> "S" Then
                    strErrMessage = "Impossibile impostare lo stato del Doc. Amministrativo a Trasmesso."
                    GoTo ErrHandler
                End If
            
            Case Else
                GoTo ErrHandler
        End Select
    End With
    
    DS.Fields("10").NewValue = NewValue
    DS.Fields("10").Updated = True
    blnResult = BeginSave(blnConcurErr)
    
    If blnResult Then
        DS.DA.EndTrans etCommit
    Else
        GoTo ErrHandler
    End If
    
    SetStato = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    If Not IsMissing(strErrorMessage) Then
        strErrorMessage = IIf(strErrMessage = "", "Errore sconosciuto.", strErrMessage)
    End If
    SetStato = False
End Function



Public Function ComputeTotali(Optional ErrDescr As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDAIV As String
    Dim strErrDesc As String
    
    Dim varRigheDocumento As Variant
    
    Dim blnConcErr As Boolean
    
    Dim dblImportoIVA As Double
    Dim dblAliquotaIVA As Double
    Dim dblImportoIVAAliquota As Double
    Dim dblImponibileAliquota As Double
    Dim dblImponibileScarico As Double
    Dim dblImportoDoc As Double
    Dim dblImponibileDoc As Double
    
    Dim lngHowMany As Long
    Dim i As Long
    
    Dim AIV As CIVA
    
    dblImportoDoc = 0
    dblImponibileDoc = 0
    dblImportoIVA = 0
    
    If Not DS.DA.QueryCommand(SQL_GET_RIGHE, varRigheDocumento, lngHowMany, _
                            DS.Fields("1").Value, _
                            DS.Fields("0").Value, _
                            DS.Fields("2").Value) Then
        strErrDesc = "Errore nella lettura delle righe del documento"
        GoTo ErrHandler
    End If
    If lngHowMany = 0 Then
        strErrDesc = "Nessuna riga trovata per il documento"
        GoTo ErrHandler
    End If
    
    DS.DA.BeginTransaction
    
    If Not DS.DA.ExecuteCommand(SQL_DELETE_IVA, _
                            DS.Fields("1").Value, _
                            DS.Fields("0").Value, _
                            DS.Fields("2").Value) Then
        strErrDesc = "Errore nell'eliminazione delle righe di IVA"
        GoTo ErrHandler
    End If
    
    For i = 0 To lngHowMany - 1
        
        strIDAIV = varRigheDocumento(0, i)
        dblImponibileAliquota = varRigheDocumento(1, i)
        dblAliquotaIVA = varRigheDocumento(2, i) / 100
        
        If Len(strIDAIV) = 0 Then
            strErrDesc = "Codice IVA assente."
            GoTo ErrHandler
        End If
        
        
        dblImportoIVAAliquota = dblImponibileAliquota * dblAliquotaIVA
        
        dblImportoIVAAliquota = Round(dblImportoIVAAliquota, 2)
        ' ---
        
        If Not InsertImportiIVA(DS.Fields("0").Value, _
                                DS.Fields("1").Value, _
                                DS.Fields("2").Value, _
                                strIDAIV, _
                                dblImponibileAliquota, _
                                dblImportoIVAAliquota) Then
            strErrDesc = "Errore nell'inserimento dell'IVA"
            GoTo ErrHandler
        End If
        
        ' --- Sommiamo l'importo dell'aliquota precedente al totale
        ' --- dell'IVA del documento
        dblImportoIVA = dblImportoIVA + dblImportoIVAAliquota
        ' ---
        
        ' --- Sommiamo l'imponibile relativo al prodotto corrente al totale
        ' --- dell'imponibile per documento
        dblImponibileDoc = dblImponibileDoc + dblImponibileAliquota
        ' ---
        
    Next i
    
    dblImportoDoc = dblImponibileDoc + dblImportoIVA
    
    DS.Fields("11").NewValue = dblImportoDoc
    DS.Fields("11").Updated = True
    DS.Fields("12").NewValue = dblImportoIVA
    DS.Fields("12").Updated = True
    
    BeginSave blnConcErr
    DS.DA.EndTrans etCommit
    DS.EndSave
    
    DS.DA.EndTrans etCommit
    ComputeTotali = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    ComputeTotali = False
End Function

Public Function IsUpdatable() As Boolean
    Dim strAvanzamento As String
    Dim blnTrasmesso As Boolean
    Dim blnAutomatico As Boolean
    
    ' --- Impedisco il salvataggio del documento se
    ' --- Documento proveniente da viaggio o trasmesso
'    blnAutomatico = Len(DS.Field("Numero Viaggio").Value) > 0
    blnTrasmesso = (UCase(DS.Fields("10").Value) = "T")
    
    IsUpdatable = True
    
    If blnAutomatico Then
        IsUpdatable = False
        Exit Function
    End If
    
    If blnTrasmesso Then
        IsUpdatable = False
        Exit Function
    End If
    
    ' ---
End Function

Public Function InsertImportiIVA(IDEnteEmittente As Variant, _
                                 AnnoDocumento As Variant, _
                                 ProgrDocumento As Variant, _
                                 IDIVA As Variant, _
                                 Imponibile As Variant, _
                                 Importo As Variant) As Boolean
    On Error GoTo ErrHandler
    
    If Not DS.DA.ExecuteCommand(SQL_INSERT_IVA, _
                                IDEnteEmittente, _
                                AnnoDocumento, _
                                ProgrDocumento, _
                                IDIVA, _
                                Imponibile, _
                                Importo) Then
        GoTo ErrHandler
    End If
    
    InsertImportiIVA = True
    Exit Function
ErrHandler:
    InsertImportiIVA = False
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

    ' Esito Documento - default value
    If Not DS.Fields("10").Updated Then
        DS.Fields("10").NewValue = "G"
        DS.Fields("10").Updated = True
    End If
    
    ' Importo Documento - default value
    If Not DS.Fields("11").Updated Then
        DS.Fields("11").NewValue = 0
        DS.Fields("11").Updated = True
    End If
    
    ' Importo IVA - default value
    If Not DS.Fields("12").Updated Then
        DS.Fields("12").NewValue = 0
        DS.Fields("12").Updated = True
    End If
    
    ' Importo Incassato - default value
    If Not DS.Fields("44").Updated Then
        DS.Fields("44").NewValue = 0
        DS.Fields("44").Updated = True
    End If
    
    ' Giorni Dilazione - default value
    If Not DS.Fields("17").Updated Then
        DS.Fields("17").NewValue = 0
        DS.Fields("17").Updated = True
    End If
    
    ' Flag Incasso - default value
    If Not DS.Fields("29").Updated Then
        DS.Fields("29").NewValue = "S"
        DS.Fields("29").Updated = True
    End If
    
    ' Flag Saldo - default value
    If Not DS.Fields("30").Updated Then
        DS.Fields("30").NewValue = "N"
        DS.Fields("30").Updated = True
    End If
    
    ' Flag Tipo Fatturazione - default value"
    If Not DS.Fields("48").Updated Then
        DS.Fields("48").NewValue = "D"
        DS.Fields("48").Updated = True
    End If
    
    ' Flag Trasmesso in Contabilità - default value
    If Not DS.Fields("54").Updated Then
        DS.Fields("54").NewValue = "N"
        DS.Fields("54").Updated = True
    End If
    
    Insert = DS.Insert(Me.SQL_INSERT, "anaganu0", "anuprogr", "anuidanu", "9" & Right(DS.Fields("1").NewValue, 2), 0)
End Function
Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    BeginSave = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
End Function
Private Sub Class_Initialize()
    Dim i As Integer
    SQL_FIND = "    SELECT DAMIDAEE, DAMIDAND, DAMIDDAM, DAMIDASC, " & _
               " DAMIDADZ, DAMIDACL, DAMDTDOC, DAMDTCON, DAMIDATO, " & _
               " DAMNRDAM, DAMFLESI, DAMNRIMP, DAMNRIMI, DAMNRIVA, DAMIDAVA, " & _
               " DAMIDADE, DAMIDAMD, DAMIDAMP, DAMNRGDL, DAMIDATV, " & _
               " DAMRASOC, DAMINDIR, DAMLOCAL, DAMPROVI, DAMNRCAP, " & _
               " DAMCOFIS, DAMPAIVA, DAMRASDE, DAMINDDE, DAMLOCDE, " & _
               " DAMPRODE, DAMNRCDE, DAMRASDZ, DAMINDDZ, DAMLOCDZ, " & _
               " DAMPRODZ, DAMNRCDZ, DAMIDAVE, DAMIDOCT, DAMIDVVT, " & _
               " DAMIDVVC, DAMFLGES, DAMFLINC, DAMFLSLD, DAMNOTES, " & _
               " DAMNOPIE, DAMIDABA, DAMFLFAT, DAMRASAC, DAMRASAD, " & _
               " DAMFLTRC, DAMANNUL, TIMESTMP FROM DOCUDAM0 WHERE "
               
    SQL_GETLIST = "SELECT DAMIDAEE, DAMIDAND, DAMIDDAM FROM DOCUDAM0 "
    
    SQL_UPDATE = "UPDATE docudam0 SET "
    
    SQL_INSERT = "INSERT INTO docudam0 "
    
    SQL_GET_TS = "SELECT timestmp FROM docudam0 WHERE "
    
    SQL_GET_SEARCH_LIST = "SELECT DAMIDAEE, DAMIDAND, DAMIDDAM, DAMIDASC,DAMIDADZ, DAMIDACL, DAMIDATO, DAMANNUL FROM DOCUDAM0 "
    
    SQL_INSERT_IVA = "INSERT INTO DOCUDAI0 " & _
                     "            (DAIIDAEE, DAIIDAND, DAIIDDAM, DAIIDAIV, DAINRIMP, DAINRIMI) " & _
                     "     VALUES (?,?,?,?,?,?)"
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "5", "8", "33")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i

    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Ente Emittente", "docudam0.damidaee", 0, True, 3, "0"
         With .Item("0") ' Tabella Ente Emittente
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
'        .Add "Anno Documento", "docudam0.damidand", 0, True, 4, "1"
        .Add "Anno Documento", "docudam0.damidand", 1, True, 4, "1"
        .Add "Progressivo Documento", "docudam0.damiddam", 0, False, 7, "2", , True
        .Add "Società", "docudam0.damidasc", 0, True, 3, "3"
         With .Item("3") ' Tabella Società
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Destinazione", "docudam0.damidadz", 0, False, 10, "4"
         With .Item("4") ' Tabella Destinazioni
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Cliente", "docudam0.damidacl", 0, True, 10, "5"
         With .Item("5") ' Tabella Clienti
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With

        .Add "Data Documento", "docudam0.damdtdoc", 2, True, 10, "6"
        .Add "Data Consegna", "docudam0.damdtcon", 2, False, 10, "7"
        .Add "Tipo Documento Amministrativo", "docudam0.damidato", 0, True, 3, "8"
         With .Item("8") 'Tabella Tipo Documento
            .ForeignTableEntity = "ATO"
            .ForeignTable = "docuato0"
            .ForeignTableFld = "atoidato"
            .ForeignTableDescField = "atodescr"
            .ForeignTableDeleteFld = "atoannul"
        End With
        .Add "Numero Documento Amministrativo", "docudam0.damnrdam", 0, False, 15, "9"
        .Add "Esito Documento", "docudam0.damflesi", 0, False, 1, "10"
        .Add "Importo Documento", "docudam0.damnrimp", 1, False, 15, "11"
        .Add "Importo Incassato", "docudam0.damnrimi", 1, False, 15, "44"
        .Add "IVA Documento", "docudam0.damnriva", 1, False, 15, "12"
        .Add "Valuta", "docudam0.damidava", 0, True, 3, "13"
         With .Item("13") ' Tabella Valute
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Deposito", "docudam0.damidade", 0, False, 3, "14"
         With .Item("14") ' Tabella Depositi
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Magazzino di Deposito", "docudam0.damidamd", 0, False, 3, "15"
         With .Item("15") ' Tabella Magazzini di Deposito
            .ForeignTableEntity = "AMD"
            .ForeignTable = "anagamd0"
            .ForeignTableFld = "amdidamd"
            .ForeignTableDescField = "amddescr"
            .ForeignTableDeleteFld = "amdannul"
        End With
        .Add "Modalità di Pagamento", "docudam0.damidamp", 0, True, 3, "16"
         With .Item("16") 'Tabella Modalità di pagamento
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "Giorni Dilazione", "docudam0.damnrgdl", 1, True, 2, "17"

        .Add "Tipo Vendita", "docudam0.damidatv", 0, False, 3, "18"
        With .Item("18") 'Tabella Tipo Vendita
            .ForeignTableEntity = "ATV"
            .ForeignTable = "anagatv0"
            .ForeignTableFld = "atvidatv"
            .ForeignTableDescField = "atvdescr"
            .ForeignTableDeleteFld = "atvannul"
        End With

        .Add "Rag.Soc. Cliente", "docudam0.damrasoc", 0, True, 40, "19"
        .Add "Indir. Cliente", "docudam0.damindir", 0, True, 40, "20"
        .Add "Località Cliente", "docudam0.damlocal", 0, True, 25, "21"
        .Add "Provincia Cliente", "docudam0.damprovi", 0, True, 2, "22"
        .Add "CAP Cliente", "docudam0.damnrcap", 0, True, 5, "23"
        .Add "Cod. Fis. Cliente", "docudam0.damcofis", 0, , 16, "24"
        .Add "P. Iva Cliente", "docudam0.dampaiva", 0, , 11, "25"
        .Add "Rag.Soc. Destinazione", "docudam0.damrasde", 0, , 40, "34"
        .Add "Indir. Destinazione", "docudam0.damindde", 0, , 40, "35"
        .Add "Località Destinazione", "docudam0.damlocde", 0, , 25, "36"
        .Add "Provincia Destinazione", "docudam0.damprode", 0, , 2, "37"
        .Add "CAP Destinazione", "docudam0.damnrcde", 0, , 5, "38"
        .Add "Rag.Soc. Destinatario", "docudam0.damrasdz", 0, , 40, "39"
        .Add "Indir. Destinatario", "docudam0.daminddz", 0, , 40, "40"
        .Add "Località Destinatario", "docudam0.damlocdz", 0, , 25, "41"
        .Add "Provincia Destinatario", "docudam0.damprodz", 0, , 2, "42"
        .Add "CAP Destinatario", "docudam0.damnrcdz", 0, , 5, "43"
        .Add "Vettore", "docudam0.damidave", 0, , 3, "26"
         With .Item("26") ' Tabella Vettori
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Ordine Cliente", "docudam0.damidoct", 0, False, 7, "27"
        .Add "Numero Viaggio", "docudam0.damidvvt", 0, False, 7, "45"
        .Add "Numero Consegna", "docudam0.damidvvc", 0, False, 3, "46"
        .Add "Flag Gestione", "docudam0.damflges", 0, True, 1, "28"
        .Add "Flag Incasso", "docudam0.damflinc", 0, True, , "29"
        .Add "Flag Saldo", "docudam0.damflsld", 0, True, , "30"
        .Add "Note Testata", "docudam0.damnotes", 0, , 50, "31"
        .Add "Note piè Pagina", "docudam0.damnopie", 0, , 50, "32"
        .Add "Banca", "docudam0.damidaba", 0, , 3, "47"
        With .Item("47")
            .ForeignTableEntity = "ABA"
            .ForeignTable = "INCAABA0"
            .ForeignTableFld = "abaidaba"
            .ForeignTableDescField = "abarasoc"
            .ForeignTableDeleteFld = "abaannul"
        End With
        .Add "Flag Tipo Fatturazione", "docudam0.damflfat", 0, True, , "48"
        .Add "Rag.Soc. Agg. Cliente", "docudam0.damrasac", 0, , 40, "52"
        .Add "Rag.Soc. Agg. Destinazione", "docudam0.damrasad", 0, , 40, "53"

        ' (PM) 08/02/06
        .Add "Stato", "docudam0.damfltrc", 0, , 2, "54"

        .Add "Annullato", "docudam0.damannul", 3, True, , "33"
    End With
    With DS.KeyFields    ' CAMPI CHIAVE
        .Add "0", "1"    ' Ente Emittente
        .Add "1", "2"    ' Anno documento
        .Add "2", "3"    ' Progressivo documento
    End With
    DS.CancelField = "docudam0.damannul"
End Sub

Public Function CreateNCRFromFAT(Optional strErrorMessage As String) As Boolean
On Error GoTo ErrHandler

    Dim DAR As CRigaDocumentoAmministrativo
    Dim DAM As CDocumentoAmministrativo
    Dim lngFieldsCounter As Long
    Dim lngDARCounter As Long
    Dim blnSkipField As Boolean
    Dim Field As CField
    Dim strEnteEmittente As String
    Dim strAnnoDocumento As String
    Dim strProgressivoDocumento As String
    Dim lngHowMany As Long
    Dim varNumeroDoc As Variant
    Dim strErrMessage As String
    Dim varDAI As Variant
    Dim lngDAIHowMany As Long
    Dim lngDAICounter As Long

    Const SQL_SELECT_DAI = _
    "SELECT DAIIDAEE, DAIIDAND, DAIIDDAM, DAIIDAIV, DAINRIMP, DAINRIMI" & _
    " FROM DOCUDAI0" & _
    " WHERE DAIIDAEE = ? AND DAIIDAND = ? AND DAIIDDAM = ? "

    DS.DA.BeginTransaction

    Set DAR = New CRigaDocumentoAmministrativo
    With DAR.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With

    Set DAM = New CDocumentoAmministrativo
    With DAM.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With

    If DS.Fields("8").Value <> "FAT" Then
        strErrMessage = "Il documento non e' una fattura."
        GoTo ErrHandler
    End If

    ' Inserisce la testata
    For lngFieldsCounter = 0 To DS.Fields.Count - 1

        Set Field = DS.Fields.Item(CStr(lngFieldsCounter))
        blnSkipField = False

        If Field.Name = "Progressivo Documento" Or _
            Field.Name = "Annullato" Or _
            Field.Name = "Tipo Documento Amministrativo" Or _
            Field.Name = "Numero Documento Amministrativo" Or _
            Not UCase(Field.FieldName) Like "DOCUDAM0.*" Or _
            IsNull(Field.Value) Then
            blnSkipField = True
        End If
'            Field.Name = "Numero Viaggio" Or _
'            Field.Name = "Numero Consegna" Or _

        If Not blnSkipField Then
            DAM.DS.Fields.Item(CStr(lngFieldsCounter)).NewValue = Field.Value
            DAM.DS.Fields.Item(CStr(lngFieldsCounter)).Updated = True
        End If

    Next

    varNumeroDoc = DS.DA.GetCounter("docunda0", "ndanrprg", _
                                    "ndaidasc", DS.Fields("3").Value, dtAlfa, _
                                    "ndaidaee", DS.Fields("0").Value, dtAlfa, _
                                    "ndaidato", "FAT", dtAlfa)

    If varNumeroDoc = -1 Then
        strErrMessage = "Impossibile leggere il numeratore per il documento." & vbCrLf & "Società: " & DS.Fields("3").Value & vbCrLf & "Tipo Documento: NCR"
        GoTo ErrHandler
    End If

    DAM.DS.Fields("8").NewValue = "NCR"
    DAM.DS.Fields("33").NewValue = False
    DAM.DS.Fields("9").NewValue = varNumeroDoc
    DAM.DS.Fields("6").NewValue = CDate(Now)
    DAM.DS.Fields("8").Updated = True
    DAM.DS.Fields("33").Updated = True
    DAM.DS.Fields("9").Updated = True
    DAM.DS.Fields("6").Updated = True

    If Not DAM.Insert() Then
        GoTo ErrHandler
    End If

    DAR.DS.Fields("0").FindValue = DS.Fields("0").Value
    DAR.DS.Fields("1").FindValue = DS.Fields("1").Value
    DAR.DS.Fields("2").FindValue = DS.Fields("2").Value
    DAR.DS.Fields("24").FindValue = False
    DAR.GetList
    lngHowMany = DAR.DS.HowMany
    If lngHowMany > 0 Then

        For lngDARCounter = 0 To lngHowMany - 1

            DAR.DS.Fields.ResetFields

            DAR.DS.Fields("0").FindValue = DAR.DS.List(0, lngDARCounter)
            DAR.DS.Fields("1").FindValue = DAR.DS.List(1, lngDARCounter)
            DAR.DS.Fields("2").FindValue = DAR.DS.List(2, lngDARCounter)
            DAR.DS.Fields("3").FindValue = DAR.DS.List(3, lngDARCounter)

            DAR.Find

            ' Inserisce le righe
            For lngFieldsCounter = 0 To DAR.DS.Fields.Count - 1

                Set Field = DAR.DS.Fields.Item(CStr(lngFieldsCounter))
                blnSkipField = False

                If Field.Name = "Progressivo Documento" Or _
                    Field.Name = "Annullato" Or _
                    Not UCase(Field.FieldName) Like "DOCUDAR0.*" Or _
                    IsNull(Field.Value) Then
                    blnSkipField = True
                End If
'                    Field.Name = "Viaggio" Or _
'                    Field.Name = "Consegna" Or _
'                    Field.Name = "Scarico" Or _

                If Not blnSkipField Then
                    DAR.DS.Fields.Item(CStr(lngFieldsCounter)).NewValue = Field.Value
                    DAR.DS.Fields.Item(CStr(lngFieldsCounter)).Updated = True
                End If
            Next

            DAR.DS.Fields("2").NewValue = DAM.DS.Fields("2").Value
            DAR.DS.Fields("24").NewValue = False
            DAR.DS.Fields("2").Updated = True
            DAR.DS.Fields("24").Updated = True
            If Not DAR.Insert() Then
                GoTo ErrHandler
            End If
        Next
    End If

    lngDAIHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SELECT_DAI, varDAI, lngDAIHowMany, _
            DS.Fields("0").Value, _
            DS.Fields("1").Value, _
            DS.Fields("2").Value) Then
        GoTo ErrHandler
    End If

    If lngDAIHowMany > 0 Then
        For lngDAICounter = LBound(varDAI, 2) To UBound(varDAI, 2)
            If Not InsertImportiIVA(DAM.DS.Fields("0").Value, _
                DAM.DS.Fields("1").Value, _
                DAM.DS.Fields("2").Value, _
                varDAI(3, lngDAICounter), _
                varDAI(4, lngDAICounter), _
                varDAI(5, lngDAICounter)) Then
                GoTo ErrHandler
            End If
        Next lngDAICounter
    End If

    DS.DA.EndTrans 0
    CreateNCRFromFAT = True
    Exit Function

ErrHandler:
    DS.DA.EndTrans 1
    CreateNCRFromFAT = False
End Function

Public Function GetIvaDetail(varIdAEEDoc As Variant, varIDAnnoDoc As Variant, varIDprogDoc As Variant) As Variant
On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If DS.DA.QueryCommand(SQL_DAI_DETAIL, varBuffer, lngHowMany, varIdAEEDoc, varIDAnnoDoc, varIDprogDoc) Then
        GetIvaDetail = varBuffer
    Else
       GoTo ErrHandler
    End If
    
    Exit Function
    
ErrHandler:
    
    GetIvaDetail = Empty
    
End Function

Private Sub Class_Terminate()
    Set DS = Nothing
End Sub
Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub



