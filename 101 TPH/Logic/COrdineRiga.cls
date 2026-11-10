VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrdineRiga"
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
Public SQL_SET_ORD_MODIFIED As String
Public SQL_GET_ORD_STATE As String
Public SQL_GET_ROW_FROM_PRODUCT As String
Private SQL_GET_QTA_TO_EXPORT As String
Private SQL_GET_NEW_ROW As String
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert(Optional MoverFlag As Boolean = False) As Boolean
    Dim blnInsertOk As Boolean
    Dim strQ As String
    Dim strAvanzamento As String
    Dim varResult As Variant
    
    strQ = SQL_GET_ORD_STATE
    strQ = strQ & DS.DA.MakeSQL(DS.Fields("0").NewValue, DS.Fields("0").FldType)
    DS.DA.QueryData strQ, varResult
    If IsEmpty(varResult) Then
        strAvanzamento = "A"
    Else
        strAvanzamento = varResult(0, 0)
    End If
    If Not MoverFlag Then
        If (strAvanzamento = "A") Or (strAvanzamento = "M") Then
            
            If (DS.Fields("7").Updated) And (Not DS.Fields("11").Updated) Then
                DS.Fields("11").NewValue = DS.Fields("7").NewValue
                DS.Fields("11").Updated = True
            End If
            blnInsertOk = DS.Insert(Me.SQL_INSERT)
            Insert = blnInsertOk
        Else
            Insert = False
            Exit Function
        End If
    Else
        If strAvanzamento = "V" Then
            blnInsertOk = DS.Insert(Me.SQL_INSERT)
            Insert = blnInsertOk
        Else
            Insert = False
            Exit Function
        End If
    End If
    ' ---
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim i As Integer
    Dim blnQtyChanged As Boolean
    Dim blnSaveOk As Boolean
    Dim strQ As String
    Dim strAvanzamento As String
    Dim strIDOrdine As String
    Dim varResult As Variant
    Dim varQtaDiff As Variant
    Dim OCT As COrdine
    Dim OCR As COrdineRiga
    
    strIDOrdine = DS.Fields("0").Value
    strQ = SQL_GET_ORD_STATE
    strQ = strQ & DS.DA.MakeSQL(strIDOrdine, DS.Fields("0").FldType)
    DS.DA.QueryData strQ, varResult
    strAvanzamento = varResult(0, 0)
    If (strAvanzamento = "A") Or (strAvanzamento = "M") Or (strAvanzamento = "V") Then
        For i = 6 To 13
            If DS.Fields(CStr(i)).Updated Then
                blnQtyChanged = True
                varQtaDiff = DS.Fields(CStr(i)).NewValue - DS.Fields(CStr(i)).Value
            End If
        Next i
        blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
        If blnSaveOk And blnQtyChanged Then
            strQ = SQL_SET_ORD_MODIFIED & DS.DA.MakeSQL("M", 0)
            strQ = strQ & " WHERE octidoct = " & DS.DA.MakeSQL(strIDOrdine, DS.Fields("0").FldType)
            DS.DA.ExecuteQuery strQ
            
'            Set OCT = New COrdine
'            With OCT.DS
'                Set .DA = DS.DA
'                Set .ErrH = DS.ErrH
'                .Field("Numero Ordine").FindValue = strIDOrdine
'                OCT.Find
'                ' Stiamo modificando la riga di un ordine proveniente da un ordine aperto?
'                If (Len(.Field("Ordine di Riferimento").Value) > 0) And (Len(.Field("Raggruppamento Utente 4").Value) > 0) Then
'                    ' Aggiorniamo la riga dell'ordine aperto con la differenza tra quantità originale e quantità modificata
'                    strIDOCTRif = DS.Field("Ordine di Riferimento").Value
'                    Set OCR = New COrdineRiga
'                    With OCR.DS
'                        Set .DA = DS.DA
'                        Set .ErrH = DS.ErrH
'                    End With
'                    OCR.DS.Field("Numero Ordine").FindValue = DS.Field("Ordine di Riferimento").Value
'                    OCR.DS.Field("Numero Riga").FindValue = DS.Field("Riga Ordine di Riferimento").Value
'                    OCR.Find
'                    OCR.DS.Field("Quantita Richiesta L").NewValue = OCR.DS.Field("Quantita Richiesta L").Value - varQtaDiff
'                    OCR.BeginSave blnConcErr
'                    OCR.DS.DA.EndTrans etCommit
'                End If
'            End With
            
        End If
    Else
        BeginSave = False
        Exit Function
    End If
    BeginSave = True
    ' ---
End Function

Public Function GetRowOrder(strIDOCT As String, strRiferimentoEsterno As String, strIDSoc As String, strAprSoc As String, ByRef strErrorMessage As String) As Long
On Error GoTo ErrHandler

    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim APR As CProdotto
    Dim strProdottoSoc1 As String
    Dim strProdottoSoc2 As String

    Dim SQL_SelectOrderRow As String
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    SQL_SelectOrderRow = "SELECT OCRIDOCR, OCRIDAPR, OCRFLAVR" & _
        " FROM ORDIOCR0" & _
        " INNER JOIN ORDIOCT0 ON OCTIDOCT = OCRIDOCT" & _
        " WHERE OCRIDOCT = " & DS.DA.MakeSQL(strIDOCT, dtAlfa) & _
        " AND OCRRIFE2 = " & DS.DA.MakeSQL(strRiferimentoEsterno, dtAlfa) & _
        " AND OCRANNUL = " & DS.DA.MakeSQL(False, dtBit) & _
        " AND OCTANNUL = " & DS.DA.MakeSQL(False, dtBit)

    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SelectOrderRow, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If

    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If

    If Not IsEmptyVariantString(varBuffer(2, 0)) And varBuffer(2, 0) <> "A" Then
        strErrorMessage = "Riga Ordine già assegnata ad un viaggio."
    End If

    ' Possibile prodotto con additivo
    strProdottoSoc1 = strAprSoc
        
    If Len(strProdottoSoc1) = 0 Then
        strErrorMessage = "Codice Prodotto non valido."
        GoTo ErrHandler
    End If
    
    ' Ricerca prodotto 2 x soc.
    With APR.DS
        .Fields.ResetFields
        .Field("Codice").FindValue = varBuffer(1, 0)
        If Not APR.Find() Or .HowMany = 0 Then
            GoTo ErrHandler
        End If
        strProdottoSoc2 = .MultiSoc.DS.GetFirmCode(strIDSoc, 3)
        
        If Len(strProdottoSoc2) = 0 Then
            strErrorMessage = "Codice Prodotto """ & varBuffer(1, 0) & """ per società """ & strIDSoc & """ non trovato."
            GoTo ErrHandler
        End If
    End With

        
    If strProdottoSoc1 <> strProdottoSoc2 Then

        If Len(strProdottoSoc1) = Len(strProdottoSoc2) Then
            strErrorMessage = "Riga ordine non trovata."
            GoTo ErrHandler
        End If
        
        If Len(strProdottoSoc1) < Len(strProdottoSoc2) Then
            If strProdottoSoc1 <> Left(strProdottoSoc2, Len(strProdottoSoc1)) Then
                strErrorMessage = "Riga ordine non trovata."
                GoTo ErrHandler
            End If
        Else
            If strProdottoSoc2 <> Left(strProdottoSoc1, Len(strProdottoSoc2)) Then
                strErrorMessage = "Riga ordine non trovata."
                GoTo ErrHandler
            End If
        End If
    
    End If

    GetRowOrder = varBuffer(0, 0)

    Set APR = Nothing
    Exit Function

ErrHandler:
    GetRowOrder = -1
    Set APR = Nothing
    Exit Function
End Function

Public Function RowFromProduct(IDOrdine As String, IDProdotto As String, Quantita As Variant, Optional ByRef ErrDesc) As String
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strError As String
    Dim varResult As Variant
    Dim HowMany As Long
    
    With DS.DA
        strQ = SQL_GET_ROW_FROM_PRODUCT & .MakeSQL(IDOrdine, 0)
        strQ = strQ & " AND ocridapr = " & .MakeSQL(IDProdotto, 0)
'        strQ = strQ & " AND ocrqrilt = " & .MakeSQL(Quantita, dtNum)
        strQ = strQ & " AND ocrannul = " & .MakeSQL(False, 3)
        If Not .QueryData(strQ, varResult, HowMany) Then
            GoTo ErrHandler
        End If
        
        ' Non prevedo di gestire ordini con più righe con lo stesso prodotto
        ' Eventuali ordini con più righe per lo stesso prodotto devono venire partizionati
        If HowMany <> 1 Then
            strError = "Esistono più righe ordine con lo stesso prodotto"
            GoTo ErrHandler
        End If
    End With
    RowFromProduct = varResult(0, 0)
    Exit Function
ErrHandler:
    If Not IsMissing(ErrDesc) Then
        ErrDesc = strError
    End If
    RowFromProduct = vbNullString
End Function

Public Function SetQtaConsegnate(varNewValues As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim j As Integer
    Dim blnConcurErr As Boolean
    
    For j = 0 To 3
        DS.Fields(CStr(j + 10)).NewValue = varNewValues(j)
        DS.Fields(CStr(j + 10)).Updated = True
    Next j
    If Not DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr) Then
        GoTo ErrHandler
    End If
    DS.DA.EndTrans 0
    SetQtaConsegnate = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    SetQtaConsegnate = False
End Function

Private Sub Class_Initialize()
    SQL_FIND = "SELECT ocridoct, ocridocr, ocridapr, ocridaim, " & _
               "       ocridapd, ocrdesri, ocrqrikg, ocrqrilt, " & _
               "       ocrqripz, ocrqricn, ocrqcokg, ocrqcolt, " & _
               "       ocrqcopz, ocrqcocn, ocridaiv, ocridamp, " & _
               "       ocrprezz, ocrscper, ocrscval, ocrfltir, " & _
               "       ocrflstr, ocrflavr, ocridot1, ocridor1, " & _
               "       ocrrife1, ocrrife2, ocrragg1, ocrragg2, " & _
               "       ocrragg3, ocrnotbk, ocrnrtck, ocridaum, " & _
               "       ocrannul, timestmp FROM ordiocr0 WHERE "
    SQL_GETLIST = "SELECT ocridocr, ocridapr, ocridapd, ocridaim, ocrqrilt, ocrqcolt, ocridaiv, ocridamp, ocrannul, ocrflstr FROM ordiocr0 "
    SQL_UPDATE = "UPDATE ordiocr0 SET "
    SQL_INSERT = "INSERT INTO ordiocr0 "
    SQL_GET_TS = "SELECT timestmp FROM ordiocr0 WHERE "
    SQL_SET_ORD_MODIFIED = "UPDATE ordioct0 SET octflorm = "
    SQL_GET_ORD_STATE = "SELECT octflaor FROM ordioct0 WHERE octidoct = "
    SQL_GET_ROW_FROM_PRODUCT = "SELECT ocridocr FROM ordiocr0 WHERE ocridoct = "
    SQL_GET_QTA_TO_EXPORT = "SELECT ordiocr0.ocrqcolt,ordiocr0.ocrqrilt, ordiocr0.ocridapd FROM ordiocr0 WHERE ocridoct=? AND ocridapr=?"
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Ordine", "ordiocr0.ocridoct", 0, True, 7, "0", , False
        .Add "Numero Riga", "ordiocr0.ocridocr", 1, True, 3, "1"
        .Add "Prodotto", "ordiocr0.ocridapr", 0, True, 6, "2"
        With .Item("2")
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "ordiocr0.ocridaim", 0, True, 3, "3"
        With .Item("3")
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "ordiocr0.ocridapd", 0, True, 3, "4"
        With .Item("4")
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Descrizione", "ordiocr0.ocrdesri", 0, False, 30, "5"
        .Add "Quantita Richiesta Kg", "ordiocr0.ocrqrikg", 1, False, 15, "6"
        .Add "Quantita Richiesta L", "ordiocr0.ocrqrilt", 1, False, 15, "7"
        .Add "Quantita Richiesta Pz", "ordiocr0.ocrqripz", 1, False, 15, "8"
        .Add "Quantita Richiesta Cnf", "ordiocr0.ocrqricn", 1, False, 15, "9"
        .Add "Quantita Iniziale Kg", "ordiocr0.ocrqcokg", 1, False, 15, "10", , True
        .Add "Quantita Iniziale L", "ordiocr0.ocrqcolt", 1, False, 15, "11", , True
        .Add "Quantita Iniziale Pz", "ordiocr0.ocrqcopz", 1, False, 15, "12", , True
        .Add "Quantita Iniziale Cnf", "ordiocr0.ocrqcocn", 1, False, 15, "13", , True
        .Add "IVA", "ordiocr0.ocridaiv", 0, False, 3, "14"
        With .Item("14")
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Modalità di Pagamento", "ordiocr0.ocridamp", 0, False, 3, "15"
        With .Item("15")
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "Prezzo", "ordiocr0.ocrprezz", 1, False, 15, "16"
        .Add "Sconto Percentuale", "ordiocr0.ocrscper", 1, False, 6, "17"
        .Add "Sconto Valore", "ordiocr0.ocrscval", 1, False, 15, "18"
        .Add "Tipo Riga", "ordiocr0.ocrfltir", 0, False, 1, "19"
        .Add "Stato Riga", "ordiocr0.ocrflstr", 0, True, 1, "20"
        .Add "Avanzamento Riga", "ordiocr0.ocrflavr", 0, False, 1, "21"
        .Add "Ordine di Riferimento", "ordiocr0.ocridot1", 0, False, 7, "22"
        .Add "Riga Ordine di Riferimento", "ordiocr0.ocridor1", 1, False, 3, "23"
        .Add "Riferimento Esterno 1", "ordiocr0.ocrrife1", 0, False, 15, "24"
        .Add "Riferimento Esterno 2", "ordiocr0.ocrrife2", 0, False, 15, "25"
        .Add "Raggruppamento Utente 1", "ordiocr0.ocrragg1", 0, False, 2, "26"
        .Add "Raggruppamento Utente 2", "ordiocr0.ocrragg2", 0, False, 2, "27"
        .Add "Raggruppamento Utente 3", "ordiocr0.ocrragg3", 0, False, 2, "28"
'        .Add "Note Bunker", "ordiocr0.ocrnotbk", 0, False, 150, "30"
        .Add "Note Bunker", "ordiocr0.ocrnotbk", 0, False, 64, "30"
        .Add "Numero Ticket", "ordiocr0.ocrnrtck", 0, False, 15, "31"
        .Add "UM", "ordiocr0.ocridaum", 0, False, 3, "32"
        With .Item("32")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Annullato", "ordiocr0.ocrannul", 3, , , "29"
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
  End With
    DS.CancelField = "ordiocr0.ocrannul"
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

Public Function GetQtaFromOrderRow(CodOCT As String, CodAPR As String, varBuffer As Variant) As Integer
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
 
    
    blnResult = DS.DA.QueryCommand(SQL_GET_QTA_TO_EXPORT, varBuffer, 0, CodOCT, CodAPR)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
    GetQtaFromOrderRow = 1
    Exit Function
ErrHandler:
    GetQtaFromOrderRow = -1
End Function
Public Function GetNewRowOrdine(CodOCT As String, ByRef StrIDNewRow As String) As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    SQL_GET_NEW_ROW = "SELECT MAX(OCRIDOCR) FROM ORDIOCR0 WHERE OCRIDOCT=?"
    
    blnResult = DS.DA.QueryCommand(SQL_GET_NEW_ROW, varBuffer, 0, CodOCT)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If IsNull(varBuffer(0, 0)) Then
        StrIDNewRow = "0"
    Else
        StrIDNewRow = varBuffer(0, 0)
    End If
    GetNewRowOrdine = True
    Exit Function
ErrHandler:
    
    GetNewRowOrdine = False
End Function
Public Function UpdateQtaConcordata(CodOCT As String, CodOCR As String, LngQtaConc As Long) As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim SQL_UPDATE_QTA_CONCMOVER As String
    
    SQL_UPDATE_QTA_CONCMOVER = "UPDATE ORDIOCR0 SET OCRQRILT=" & LngQtaConc & " WHERE OCRIDOCT=? AND OCRIDOCR=? "
    blnResult = DS.DA.ExecuteCommand(SQL_UPDATE_QTA_CONCMOVER, CodOCT, CodOCR)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    UpdateQtaConcordata = True
    Exit Function
ErrHandler:
    
    UpdateQtaConcordata = False
End Function
Public Function UpdateNrTicket(IDOrdine As String, IDProdotto As String, Quantita As Variant, NrTicket As String, _
                               Optional NumPartizione As String = "00", Optional StrRife2 As String = "") As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim strSQL As String
    
    If StrRife2 = "" Then
        strSQL = "UPDATE ORDIOCR0 SET OCRNRTCK= ? WHERE OCRIDOCT=? AND OCRIDAPR=? " 'AND OCRQRILT=? "
        
        blnResult = DS.DA.ExecuteCommand(strSQL, NrTicket, IDOrdine, IDProdotto) ', CLng(Quantita))
        If Not blnResult Then
            GoTo ErrHandler
        End If
    Else
        strSQL = "UPDATE ORDIOCR0 SET OCRNRTCK= ? WHERE OCRIDOCT=? AND OCRIDAPR=? AND OCRRIFE2=? " 'AND OCRQRILT=? "
        
        blnResult = DS.DA.ExecuteCommand(strSQL, NrTicket, IDOrdine, IDProdotto, StrRife2) ', CLng(Quantita))
        If Not blnResult Then
            GoTo ErrHandler
        End If
    
    End If
    
    UpdateNrTicket = True
    Exit Function
ErrHandler:
    
    UpdateNrTicket = False
End Function

Public Function GetOrderList(CodADZ, varBuffer As Variant, HowManyRows As Integer) As Integer
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim SQL_GET_ORDER_LIST As String
    HowManyRows = 0
    SQL_GET_ORDER_LIST = "SELECT APZIDAPR, (PREZAPZ0.APZQTCIS * PREZAPZ0.APZCACIS) AS QTADAPROPORRE, " & _
                         "       APZIDAPD, APZIDAIM    " & _
                         "  FROM PREZAPZ0 WHERE APZIDADZ=? AND APZANNUL = 0"
    
    blnResult = DS.DA.QueryCommand(SQL_GET_ORDER_LIST, varBuffer, 0, CodADZ)
    
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Or IsNull(varBuffer) Then
        GoTo ErrHandler
    End If
    HowManyRows = UBound(varBuffer, 2)
    GetOrderList = 1
    Exit Function
ErrHandler:
    GetOrderList = -1
End Function

Public Function SetAvanzamento(strAvanzamento As String) As Boolean
On Error GoTo ErrHandler
    
    Dim SQL_OrderUpdateAvanzamento As String
    
    SQL_OrderUpdateAvanzamento = "UPDATE ORDIOCR0 SET OCRFLAVR = " & DS.DA.MakeSQL(strAvanzamento, dtAlfa) & _
        " WHERE OCRIDOCT = " & DS.DA.MakeSQL(DS.Field("Numero Ordine").Value, dtAlfa) & _
        " AND OCRIDOCR = " & DS.DA.MakeSQL(DS.Field("Numero Riga").Value, dtNum) & _
        " AND OCRANNUL = " & DS.DA.MakeSQL(False, dtBit)

    If Not DS.DA.ExecuteCommand(SQL_OrderUpdateAvanzamento) Then
        GoTo ErrHandler
    End If

    SetAvanzamento = True
    Exit Function
ErrHandler:
    SetAvanzamento = False
End Function

Public Function SetAnnullato(varIDOCT As Variant, varIDOCR As Variant, blnAnnullato As Boolean) As Boolean
On Error GoTo ErrHandler
    
    Dim blnResult As Boolean
    Dim strSQLUpdate As String

    strSQLUpdate = "UPDATE ORDIOCR0 SET OCRANNUL = " & DS.DA.MakeSQL(blnAnnullato, dtBit) & _
        " WHERE OCRIDOCT = ? AND OCRIDOCR = ?"
    
    If IsEmptyVariantString(varIDOCT) Or IsEmptyVariantString(varIDOCR) Then
        GoTo ErrHandler
    End If

    If Not DS.DA.ExecuteCommand(strSQLUpdate, varIDOCT, varIDOCR) Then
        GoTo ErrHandler
    End If

    SetAnnullato = True
    Exit Function

ErrHandler:
    SetAnnullato = False
End Function

