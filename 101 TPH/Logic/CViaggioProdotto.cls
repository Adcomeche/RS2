VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioProdotto"
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
Public SQL_GET_TRIP_STATE As String
Public SQL_GET_ORDINE_RIGHE As String
Public SQL_UPDATE_QTY As String
Public SQL_UPDATE_QTY_DELIVERED As String
Public SQL_UPDATE_QTY_DELIVERED_2 As String
Public SQL_UPDATE_QTY_LD As String
Public SQL_UPDATE_QTY_LD_2 As String
Public SQL_GET_TRIP_LIST As String
Public SQL_COMPLETE_ROW As String
Public SQL_SET_DEL_DATE As String
Public SQL_SET_ORD_DATE As String
Private SQL_GET_SELECTION_PRODOTTI_EVENTI As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private Const SQL_GET_DATA_INVIO_GES = "SELECT VVPDTGES, VVPHRGES FROM VIAGVVP0 WHERE VVPIDVVT = ? AND VVPIDVVC = ? AND VVPIDVVP = ?"

Private Const SQL_GET_TRASMITTED = "SELECT 1 FROM VIAGVVP0 WHERE VVPIDVVT = ? AND VVPIDVVC = ? AND NOT VVPDTGES IS NULL AND VVPDTGES <> '01/01/1900' AND VVPANNUL = 0"

Public Function CompleteRow(Qta As Variant, _
                            Data As Variant, _
                            Ora As Variant, _
                            KmRaggio As Variant, _
                            RifEst As Variant, _
                            InizioCont As Variant, _
                            IDViaggio As Variant, _
                            IDConsegna As Variant, _
                            IDScarico As Variant, _
                            IDOrdine As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    DS.DA.BeginTransaction
    
    If Not DS.DA.ExecuteCommand(SQL_COMPLETE_ROW, Qta, RifEst, InizioCont, _
                                IDViaggio, IDConsegna, IDScarico) Then
        GoTo ErrHandler
    End If
    Ora = Trim("01/01/1900 " & Ora)
    If Not DS.DA.ExecuteCommand(SQL_SET_DEL_DATE, DS.DA.MakeSQL(Data, dtDate), Ora, KmRaggio, IDViaggio, IDConsegna) Then
        GoTo ErrHandler
    End If
    If Not DS.DA.ExecuteCommand(SQL_SET_ORD_DATE, Data, IDOrdine) Then
        GoTo ErrHandler
    End If
    
    VVT.DS.Fields("0").FindValue = IDViaggio
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
   
    DS.DA.EndTrans etCommit
    CompleteRow = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    CompleteRow = False
End Function

Public Function OrdineToProdotti(IDOrdine As String, _
                                 IDViaggio As String, _
                                 IDConsegna As String) As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    Dim strIDVVP As String
    Dim lngHowMany As Long
    Dim varRigheOrdine As Variant
    Dim i As Long
    
    strQ = SQL_GET_ORDINE_RIGHE & DS.DA.MakeSQL(IDOrdine, 0)
    strQ = strQ & " AND ocrannul = " & DS.DA.MakeSQL(False, 3)
    strQ = strQ & " AND ocrflstr = " & DS.DA.MakeSQL("A", 0)
    If Not DS.DA.QueryData(strQ, varRigheOrdine, lngHowMany) Then
        GoTo ErrHandler
    End If
    
    strIDVVP = "10"
    For i = LBound(varRigheOrdine, 2) To UBound(varRigheOrdine, 2)
        With DS.Fields
            .Item("0").NewValue = IDViaggio
            .Item("0").Updated = True
            .Item("1").NewValue = IDConsegna
            .Item("1").Updated = True
            .Item("2").NewValue = strIDVVP
            .Item("2").Updated = True
            .Item("3").NewValue = varRigheOrdine(0, i)
            .Item("3").Updated = True
            .Item("4").NewValue = varRigheOrdine(1, i)
            .Item("4").Updated = True
            .Item("5").NewValue = varRigheOrdine(2, i)
            .Item("5").Updated = True
            .Item("6").NewValue = varRigheOrdine(3, i)
            .Item("6").Updated = True
            .Item("7").NewValue = varRigheOrdine(4, i)
            .Item("7").Updated = True
            .Item("8").NewValue = varRigheOrdine(5, i)
            .Item("8").Updated = True
            .Item("9").NewValue = varRigheOrdine(6, i)
            .Item("9").Updated = True
            .Item("10").NewValue = varRigheOrdine(3, i)
            .Item("10").Updated = True
            .Item("11").NewValue = varRigheOrdine(4, i)
            .Item("11").Updated = True
            .Item("12").NewValue = varRigheOrdine(5, i)
            .Item("12").Updated = True
            .Item("13").NewValue = varRigheOrdine(6, i)
            .Item("13").Updated = True
'            .Item("19").NewValue = varRigheOrdine(7, i)
'            .Item("19").Updated = True
            .Item("20").NewValue = varRigheOrdine(8, i)
            .Item("20").Updated = True
            .Item("21").NewValue = varRigheOrdine(9, i)
            .Item("21").Updated = True
            .Item("22").NewValue = varRigheOrdine(10, i)
            .Item("22").Updated = True
            .Item("23").NewValue = varRigheOrdine(11, i)
            .Item("23").Updated = True
            .Item("24").NewValue = varRigheOrdine(12, i)
            .Item("24").Updated = True
            .Item("25").NewValue = varRigheOrdine(13, i)
            .Item("25").Updated = True
            .Item("26").NewValue = varRigheOrdine(14, i)
            .Item("26").Updated = True
            .Item("32").NewValue = varRigheOrdine(15, i)
            .Item("32").Updated = True
            .Item("33").NewValue = varRigheOrdine(16, i)
            .Item("33").Updated = True
        End With
        If Not Insert Then
            GoTo ErrHandler
        End If
        strIDVVP = CStr(CInt(strIDVVP) + 10)
    Next i
    
    OrdineToProdotti = True
    Exit Function
ErrHandler:
    OrdineToProdotti = False
End Function

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    Dim blnSaveOk As Boolean
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    If Not IsTripUpdatable(DS.Fields("0").NewValue) Then
        Insert = False
        Exit Function
    End If
    Insert = DS.Insert(Me.SQL_INSERT)
    VVT.DS.Fields("0").FindValue = DS.Fields("0").NewValue
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
End Function

Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    If Not IsTripUpdatable(DS.Fields("0").Value) Then
        BeginSave = False
        Exit Function
    End If

    ' Imposto il default dello stato satellitare
    If Not DS.Fields("37").Updated And DS.Fields("37").Value = "" Then
        DS.Fields("37").NewValue = "0"
        DS.Fields("37").Updated = True
    End If

    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    VVT.DS.Fields("0").FindValue = DS.Fields("0").Value
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
    BeginSave = True
    
    ' ---
End Function

Public Function IsTripUpdatable(IDTrip As String) As Boolean
    Dim strQ As String
    Dim blnTrasmAGIP As Boolean
    Dim blnDefinitivo As Boolean
    Dim varResult As Variant
    Dim strNumeroViaggioDef As String
    
    strQ = SQL_GET_TRIP_STATE
    strQ = strQ & DS.DA.MakeSQL(IDTrip, DS.Fields("0").FldType)
    DS.DA.QueryData strQ, varResult
    
    If varResult(2, 0) = "E" Then
        IsTripUpdatable = False
        Exit Function
    End If
    
    IsTripUpdatable = True
'    If IsNull(varResult(0, 0)) Then
'        strNumeroViaggioDef = "0"
'    Else
'        strNumeroViaggioDef = varResult(0, 0)
'    End If
'
'    If Trim$(strNumeroViaggioDef) <> vbNullString Then
'        If CLng(strNumeroViaggioDef) > 0 Then
'            blnDefinitivo = True
'        End If
'    End If
'
'    If IsNull(varResult(1, 0)) Then
'        blnTrasmAGIP = False
'    Else
'        blnTrasmAGIP = (varResult(1, 0) = "T")
'    End If
'
'    If (Not blnTrasmAGIP) And blnDefinitivo Then
'        IsTripUpdatable = False
'    Else
'        IsTripUpdatable = True
'    End If
End Function

Public Function UpdateQty(IDViaggio As String, _
                          IDConsegna As String, _
                          IDProdotto As String, _
                          Qta As Single, _
                          RifBolletta As String) As Boolean
    On Error GoTo ErrHandler
    Dim strSQL As String
    Dim blnResult As Boolean
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    strSQL = Me.SQL_UPDATE_QTY
    
    blnResult = DS.DA.ExecuteCommand(strSQL, Qta, RifBolletta, IDViaggio, IDConsegna, IDProdotto)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    VVT.DS.Fields("0").FindValue = IDViaggio
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
    UpdateQty = True
    Exit Function
ErrHandler:
    UpdateQty = False
End Function
Public Function UpdateQtyDelivered(IDViaggio As String, _
                                    IDConsegna As String, _
                                    IDProdotto As String, _
                                    Qta As Single, _
                                    Optional QtaKG As Single, _
                                    Optional AggiornaQtaCarico As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim strSQL As String
    Dim blnResult As Boolean
    
    If Not IsMissing(QtaKG) Then
        If Not AggiornaQtaCarico Then
            strSQL = Me.SQL_UPDATE_QTY_DELIVERED_2
            blnResult = DS.DA.ExecuteCommand(strSQL, Qta, QtaKG, IDViaggio, IDConsegna, IDProdotto)
        Else
            strSQL = Me.SQL_UPDATE_QTY_LD_2
            blnResult = DS.DA.ExecuteCommand(strSQL, Qta, QtaKG, Qta, QtaKG, IDViaggio, IDConsegna, IDProdotto)
        End If
    Else
        If Not AggiornaQtaCarico Then
            strSQL = Me.SQL_UPDATE_QTY_DELIVERED
            blnResult = DS.DA.ExecuteCommand(strSQL, Qta, IDViaggio, IDConsegna, IDProdotto)
        Else
            strSQL = Me.SQL_UPDATE_QTY_LD
            blnResult = DS.DA.ExecuteCommand(strSQL, Qta, Qta, IDViaggio, IDConsegna, IDProdotto)
        End If
    End If
    
    If Not blnResult Then
        GoTo ErrHandler
    End If
    
    UpdateQtyDelivered = True
    Exit Function
ErrHandler:
    UpdateQtyDelivered = False
End Function
Private Sub Class_Initialize()
    SQL_FIND = "SELECT vvpidvvt, vvpidvvc, vvpidvvp, vvpidapr, " & _
               "       vvpidaim, vvpidapd, vvpqtrkg, vvpqtrlt, " & _
               "       vvpqtrpz, vvpqtrcn, vvpqcokg, vvpqcolt, " & _
               "       vvpqcopz, vvpqcocn, vvpprinc, vvpprfic, " & _
               "       vvpidvt1, vvpidvc1, vvpidvp1, vvprife1, " & _
               "       vvprife2, vvpidaiv, vvpidamp, vvpprezz, " & _
               "       vvpscper, vvpscval, vvpfltir, vvpflst1, " & _
               "       vvpflst2, vvpflst3, vvpflst4, vvpflst5, " & _
               "       vvpidoct, vvpidocr, vvpidppd, vvpidkde, " & _
               "       vvpflsta, vvphrins, vvphrfis, vvpflncc, " & _
               "       vvpfldes, vvpnrver, vvpflagg, vvpdtges, " & _
               "       vvphrges, vvpannul, timestmp FROM viagvvp0 WHERE "
'    SQL_GETLIST = "SELECT vvpidvvp, vvpidapr, vvpqcolt, vvpqcokg, vvpprinc, vvpprfic, vvpprezz, vvpannul FROM viagvvp0 "
'    SQL_GETLIST = "SELECT vvpidvvp, vvpidapr, vvpqtrlt, vvpqcolt, vvpprinc, vvpprfic, vvpprezz, vvpannul, vvpidapd FROM viagvvp0 "
    SQL_GETLIST = "SELECT vvpidvvp, vvpidapr, vvpqtrlt, vvpqcolt, vvprife1, vvpidoct, vvpannul , vvpidapd FROM viagvvp0 "
    SQL_UPDATE = "UPDATE viagvvp0 SET "
    SQL_INSERT = "INSERT INTO viagvvp0 "
    SQL_GET_TS = "SELECT timestmp FROM viagvvp0 WHERE "
    SQL_GET_TRIP_STATE = "SELECT vvtiddef, vvtflst5, vvtflstv FROM viagvvt0 WHERE vvtidvvt = "
'    SQL_GET_ORDINE_RIGHE = "SELECT ocridapr, ocridaim, ocridapd, ocrrife1, " & _
'                           "       ocrrife2, ocridaiv, ocridamp, ocrqrikg, " & _
'                           "       ocrqrilt, ocrqricn, ocrqripz, ocrprezz, " & _
'                           "       ocrscper, ocrscval, ocrfltir, ocridoct, " & _
'                           "       ocridocr FROM ordiocr0 " & _
'                           "       WHERE ocridoct =  "
    SQL_GET_ORDINE_RIGHE = "SELECT ocridapr, ocridaim, ocridapd, ocrqrikg, " & _
                           "       ocrqrilt, ocrqricn, ocrqripz, ocrrife1, " & _
                           "       ocrrife2, ocridaiv, ocridamp, ocrprezz, " & _
                           "       ocrscper, ocrscval, ocrfltir, ocridoct, " & _
                           "       ocridocr FROM ordiocr0 " & _
                           "       WHERE ocridoct =  "
    SQL_UPDATE_QTY = "UPDATE viagvvp0 SET vvpqcolt = ?, vvprife1 = ? WHERE " & _
                     "       vvpidvvt = ? AND vvpidvvc = ? AND vvpidvvp = ?"
    SQL_UPDATE_QTY_DELIVERED = "UPDATE viagvvp0 SET vvpqcolt = ? WHERE " & _
                     "       vvpidvvt = ? AND vvpidvvc = ? AND vvpidvvp = ?"
    SQL_UPDATE_QTY_DELIVERED_2 = "UPDATE viagvvp0 SET vvpqcolt = ?, vvpqcokg = ? " & _
                                 " WHERE vvpidvvt = ? AND vvpidvvc = ? AND vvpidvvp = ?"
    SQL_UPDATE_QTY_LD = "UPDATE viagvvp0 SET vvpqcolt = ?, vvpqtrlt = ? WHERE " & _
                     "       vvpidvvt = ? AND vvpidvvc = ? AND vvpidvvp = ?"
    SQL_UPDATE_QTY_LD_2 = "UPDATE viagvvp0 SET vvpqcolt = ?, vvpqcokg = ?, vvpqtrlt = ?, vvpqtrkg = ? " & _
                                 " WHERE vvpidvvt = ? AND vvpidvvc = ? AND vvpidvvp = ?"
    SQL_COMPLETE_ROW = "UPDATE viagvvp0 SET vvpqcolt = ?, vvprife1 = ?, " & _
                       "       vvpprinc = ? WHERE " & _
                       "       vvpidvvt = ? AND vvpidvvc = ? AND vvpidvvp = ?"
    SQL_SET_DEL_DATE = "UPDATE viagvvc0 SET vvcdtcon = ?, vvctemvg = ?, vvckmrag = ? WHERE " & _
                       "       vvcidvvt = ? AND vvcidvvc = ?"
    SQL_SET_ORD_DATE = "UPDATE ordioct0 SET octdtcef = ? WHERE octidoct = ?"
    SQL_GET_TRIP_LIST = "SELECT vvpidvvc, vvpidvvp, vvcidadz, vvcderas, " & _
                        "       vvpidapr, vvpqcolt, vvcdtcon, CONVERT(CHAR, VVCTEMVG, 108), " & _
                        "       vvckmrag, vvprife1, vvpprinc, vvpidoct " & _
                        "  FROM viagvvc0 INNER JOIN viagvvp0 ON " & _
                        "       viagvvc0.vvcidvvt = viagvvp0.vvpidvvt AND " & _
                        "       viagvvc0.vvcidvvc = viagvvp0.vvpidvvc WHERE " & _
                        "       vvpidvvt = ? ORDER BY vvpidvvc, vvpidvvp"
    
    SQL_GET_SELECTION_PRODOTTI_EVENTI = "SELECT  VVCIDVVT, VVCIDVVC, VVPIDVVP, VVCNRSEQ, VVCIDADZ, VVCDERAS," & _
                " VVCDELOC, VVCKMPAD, VVPHRINS, NULL, VVPIDAPR," & _
                " VVCDTCON, OCRQRILT, VVPQCOLT, OCTRIFE1 + ISNULL(OCRRIFE2,'')," & _
                " VVPQTRLT, VVPFLST5, OCTRAUT4, VVCNOTE1, VVCNOTEE, ADZKMRAG," & _
                " AVVDESCR, VVPFLSTA, VVCTEMVG, VVPHRFIS" & _
                " FROM VIAGVVP0" & _
                " INNER JOIN VIAGVVC0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC" & _
                " INNER JOIN ORDIOCR0 ON VIAGVVP0.VVPIDOCT = ORDIOCR0.OCRIDOCT AND VIAGVVP0.VVPIDOCR = ORDIOCR0.OCRIDOCR" & _
                " INNER JOIN ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT AND ORDIOCR0.OCRIDOCT = ORDIOCT0.OCTIDOCT" & _
                " INNER JOIN ANAGADZ0 ON VIAGVVC0.VVCIDADZ = ANAGADZ0.ADZIDADZ" & _
                " LEFT JOIN ANAGAVV ON ANAGAVV.AVVIDTBL = 'VIAGVVP0' AND ANAGAVV.AVVIDFLD = 'VVPFLSTA' AND ANAGAVV.AVVVALUE = VIAGVVP0.VVPFLSTA"

    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Viaggio", "viagvvp0.vvpidvvt", 0, True, 7, "0", , True
        .Add "Numero Consegna", "viagvvp0.vvpidvvc", 0, True, 3, "1", , True
        .Add "Numero Riga", "viagvvp0.vvpidvvp", 0, True, 3, "2"
        .Add "Prodotto", "viagvvp0.vvpidapr", 0, True, 6, "3", , True
        With .Item("3")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "viagvvp0.vvpidaim", 0, True, 3, "4", , True
        With .Item("4")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "viagvvp0.vvpidapd", 0, True, 3, "5", , True
        With .Item("5")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Quantita Trasportata Kg", "viagvvp0.vvpqtrkg", 1, False, 15, "6" ', , True
        .Add "Quantita Trasportata L", "viagvvp0.vvpqtrlt", 1, False, 15, "7" ', , True
        .Add "Quantita Trasportata Pz", "viagvvp0.vvpqtrpz", 1, False, 15, "8" ', , True
        .Add "Quantita Trasportata Cnf", "viagvvp0.vvpqtrcn", 1, False, 15, "9" ', , True
        .Add "Quantita Consegnata Kg", "viagvvp0.vvpqcokg", 1, False, 15, "10"
        .Add "Quantita Consegnata L", "viagvvp0.vvpqcolt", 1, False, 15, "11"
        .Add "Quantita Consegnata Pz", "viagvvp0.vvpqcopz", 1, False, 15, "12"
        .Add "Quantita Consegnata Cnf", "viagvvp0.vvpqcocn", 1, False, 15, "13"
        .Add "Progressivo inizio Contalitri", "viagvvp0.vvpprinc", 1, False, 6, "14"
        .Add "Progressivo fine Contalitri", "viagvvp0.vvpprfic", 1, False, 6, "15"
        .Add "Numero Viaggio Riferimento", "viagvvp0.vvpidvt1", 0, , 7, "16", , True
        .Add "Numero Consegna Riferimento", "viagvvp0.vvpidvc1", 0, , 3, "17", , True
        .Add "Numero Riga Riferimento", "viagvvp0.vvpidvp1", 0, , 3, "18", , True
        .Add "Riferimento Esterno 1", "viagvvp0.vvprife1", 0, False, 15, "19"
        .Add "Riferimento Esterno 2", "viagvvp0.vvprife2", 0, False, 15, "20"
        .Add "IVA", "viagvvp0.vvpidaiv", 0, False, 3, "21", , True
        With .Item("21")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Modalità di Pagamento", "viagvvp0.vvpidamp", 0, False, 3, "22", , True
        With .Item("22")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "Prezzo", "viagvvp0.vvpprezz", 1, False, 15, "23", , True
        .Add "Sconto Percentuale", "viagvvp0.vvpscper", 1, False, 6, "24", , True
        .Add "Sconto Valore", "viagvvp0.vvpscval", 1, False, 15, "25", , True
        .Add "Tipo Riga", "viagvvp0.vvpfltir", 0, False, 1, "26", , True
        .Add "Stato Riga 1", "viagvvp0.vvpflst1", 0, False, 1, "27"
        .Add "Stato Riga 2", "viagvvp0.vvpflst2", 0, False, 1, "28"
        .Add "Stato Riga 3", "viagvvp0.vvpflst3", 0, False, 1, "29"
        .Add "Stato Riga 4", "viagvvp0.vvpflst4", 0, False, 1, "30"
        .Add "Stato Riga 5", "viagvvp0.vvpflst5", 0, False, 1, "31"
        .Add "Ordine Cliente", "viagvvp0.vvpidoct", 0, True, 7, "32", , True
        .Add "Riga Ordine Cliente", "viagvvp0.vvpidocr", 1, True, 3, "33", , True
        .Add "Dettaglio Prefattura", "viagvvp0.vvpidppd", 0, False, 7, "34", , True
        .Add "Knosos Id scarico", "viagvvp0.vvpidkde", 0, False, 3, "36", , True
        .Add "Stato satellitare", "viagvvp0.vvpflsta", dtAlfa, False, 1, "37", , True
        .Add "Ora inizio", "viagvvp0.vvphrins", 4, False, 5, "38", , True
        .Add "Ora fine", "viagvvp0.vvphrfis", 4, False, 5, "39", , True
        .Add "Flag Non Conformità Cliente", "viagvvp0.vvpflncc", 0, False, 2, "40"
        .Add "Flag Desigillato", "viagvvp0.vvpfldes", 0, False, 2, "41"
        .Add "Numero Venta En Ruta", "viagvvp0.vvpnrver", 0, , 7, "42", , True
        
        '-- DP 05/06/08
        '-- Aggiunto "VVPFLAGG" per l'indicazione del processo che ha aggiornato il record l'ultima volta
        .Add "Flag ultimo aggiornamento", "viagvvp0.vvpflagg", 0, False, 2, "43", , True
        
        '-- DP 24/07/08
        '-- Aggiungo 2 campi per l'indicazione di data ed ora dell'ultima esportazione su GES
        .Add "Data esportazione su GES", "viagvvp0.vvpdtges", 2, False, 10, "44"
        .Add "Ora esportazione su GES", "viagvvp0.vvphrges", 4, False, 5, "45"
        
        .Add "Annullato", "viagvvp0.vvpannul", 3, , , "35"
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
  End With
    DS.CancelField = "viagvvp0.vvpannul"
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
Public Function GetSelectionViaggioProdottiEventi(ViaggioId As String, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional MaxRecords As Long = 0) As Boolean
    Dim strStatementSQL As String

    strStatementSQL = SQL_GET_SELECTION_PRODOTTI_EVENTI
    strStatementSQL = strStatementSQL & " WHERE VVCIDVVT = " & ViaggioId
    
    If ActiveRecordsOnly Then
        strStatementSQL = strStatementSQL & " AND VVPANNUL = " & DS.DA.MakeSQL(False, dtBit)
        strStatementSQL = strStatementSQL & " AND VVCANNUL = " & DS.DA.MakeSQL(False, dtBit)
    End If
    
    strStatementSQL = strStatementSQL & " ORDER BY VVCNRSEQ"

'    DS.Fields("0").FindValue = ViaggioId

    GetSelectionViaggioProdottiEventi = DS.GetSelection(strStatementSQL, False, , MaxRecords)

End Function

'-- DP 24/07/08
'-- Controlla se c'è almeno un prodotto di una consegna già trasmesso su GES
Public Function blnGESAlreadyTransmitted(strIDVVT As String, strIDVVC As String) As Boolean
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    If DS.DA.QueryCommand(SQL_GET_TRASMITTED, varBuffer, 0, strIDVVT, strIDVVC) Then
        If IsEmpty(varBuffer) Then
            blnGESAlreadyTransmitted = False
        Else
            GoTo ErrHandler
        End If
    Else
        GoTo ErrHandler
    End If
    
    Exit Function
    
ErrHandler:
    blnGESAlreadyTransmitted = True
End Function

'-- DP 24/07/08
'-- Restituisce la data di esportazione su file GES di una riga di VVP
Public Function dtGetDataInvioGES(strIDVVT As String, strIDVVC As String, strIDVVP As String) As Date
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    If DS.DA.QueryCommand(SQL_GET_DATA_INVIO_GES, varBuffer, 0, strIDVVT, strIDVVC, strIDVVP) Then
        If Not IsEmpty(varBuffer) Then
            If Not IsNull(varBuffer(0, 0)) Then
                dtGetDataInvioGES = CDate(varBuffer(0, 0) & " " & Format(varBuffer(1, 0), "HH:MM"))
            Else
                GoTo ErrHandler
            End If
        Else
            GoTo ErrHandler
        End If
    Else
        GoTo ErrHandler
    End If
    
    Exit Function
    
ErrHandler:
    dtGetDataInvioGES = CDate("01/01/1900")
End Function
