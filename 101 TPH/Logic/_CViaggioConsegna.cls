VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioConsegna"
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
Public SQL_DELETE As String
Public SQL_DELETE_PRODOTTI As String
Public SQL_GET_TS As String
Public SQL_GET_TRIP_STATE As String
Public SQL_GET_TRIP_DATE As String
Public SQL_SET_AVANZAMENTO_ORDINE As String
Public SQL_SET_AVANZAMENTO_ORDINE_RIGHE As String
Public SQL_GET_TOTALE_QTA As String
Public SQL_GET_ORDINE As String
Public SQL_SET_KM_DEF As String
Public SQL_SET_DATE_CONSEGNE_EFF As String
Public SQL_GET_ORDERS_LIST As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarTotaliQta As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function Delete() As Boolean
    On Error GoTo ErrHandler
    Dim strQ As String
    
    
    DS.DA.BeginTransaction
    
    If Not IsTripUpdatable(DS.Fields("0").Value) Then
        GoTo ErrHandler
    End If
    
    strQ = SQL_DELETE_PRODOTTI & DS.DA.MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
    strQ = strQ & " AND vvpidvvc = " & DS.DA.MakeSQL(DS.Fields("1").Value, DS.Fields("1").FldType)
    If Not DS.DA.ExecuteQuery(strQ) Then
        GoTo ErrHandler
    End If
    
    strQ = SQL_DELETE & DS.DA.MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
    strQ = strQ & " AND vvcidvvc = " & DS.DA.MakeSQL(DS.Fields("1").Value, DS.Fields("1").FldType)
    
    If Not SetAvanzamentoOrdine("A") Then
        GoTo ErrHandler
    End If
    If Not DS.DA.ExecuteQuery(strQ) Then
        GoTo ErrHandler
    End If
    
    DS.DA.EndTrans 0
    Delete = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    Delete = False
End Function


Public Function GetSequentialList() As Boolean
    ReDim varOrderFields(1) As Variant
    
    varOrderFields(0) = "26"
    varOrderFields(1) = "2"
    GetSequentialList = DS.GetSelection(SQL_GETLIST, True, varOrderFields, _
                                        , True)
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

Public Function OrdineToConsegna(IDOrdine As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    With DS.DA
        strQ = SQL_GET_ORDINE & .MakeSQL(IDOrdine, 0)
        strQ = strQ & " AND octannul = " & .MakeSQL(False, 3)
        If Not .QueryData(strQ, varBuffer, lngHowMany) Then
            GoTo ErrHandler
        End If
    End With
    
    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If
    
    With DS.Fields
        .Item("3").NewValue = varBuffer(0, 0)
        .Item("3").Updated = True
        .Item("4").NewValue = varBuffer(1, 0)
        .Item("4").Updated = True
        .Item("5").NewValue = varBuffer(2, 0)
        .Item("5").Updated = True
        .Item("6").NewValue = varBuffer(3, 0)
        .Item("6").Updated = True
        .Item("7").NewValue = varBuffer(4, 0)
        .Item("7").Updated = True
        .Item("8").NewValue = varBuffer(5, 0)
        .Item("8").Updated = True
        .Item("9").NewValue = varBuffer(6, 0)
        .Item("9").Updated = True
        .Item("10").NewValue = varBuffer(7, 0)
        .Item("10").Updated = True
        .Item("11").NewValue = varBuffer(8, 0)
        .Item("11").Updated = True
        .Item("12").NewValue = varBuffer(9, 0)
        .Item("12").Updated = True
        .Item("13").NewValue = varBuffer(10, 0)
        .Item("13").Updated = True
        .Item("14").NewValue = varBuffer(11, 0)
        .Item("14").Updated = True
        .Item("15").NewValue = varBuffer(12, 0)
        .Item("15").Updated = True
        .Item("16").NewValue = varBuffer(13, 0)
        .Item("16").Updated = True
        .Item("17").NewValue = varBuffer(14, 0)
        .Item("17").Updated = True
        .Item("18").NewValue = varBuffer(15, 0)
        .Item("18").Updated = True
        .Item("19").NewValue = varBuffer(16, 0)
        .Item("19").Updated = True
        .Item("21").NewValue = "M"
        .Item("21").Updated = True
        .Item("23").NewValue = "M"
        .Item("23").Updated = True
        .Item("26").NewValue = varBuffer(17, 0)
        .Item("26").Updated = True
        .Item("27").NewValue = varBuffer(18, 0)
        .Item("27").Updated = True
        .Item("37").NewValue = varBuffer(19, 0)
        .Item("37").Updated = True
        .Item("30").NewValue = varBuffer(20, 0)
        .Item("30").Updated = True
        .Item("31").NewValue = varBuffer(21, 0)
        .Item("31").Updated = True
        ' --- Km Raggio
        .Item("39").NewValue = varBuffer(22, 0)
        .Item("39").Updated = True
        ' --- Consegna per ordine programmato
        .Item("34").NewValue = varBuffer(23, 0)
        .Item("34").Updated = True
        ' ---
    End With
    OrdineToConsegna = True
    Exit Function
ErrHandler:
    OrdineToConsegna = False
End Function


Public Function SetAvanzamentoOrdine(Avanzamento As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strTS As String
    Dim strTripDate As String
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    
    With DS.DA
        .BeginTransaction
        strQ = SQL_SET_AVANZAMENTO_ORDINE & .MakeSQL(Avanzamento, 0)
        
        If Avanzamento = "V" Then
            blnResult = .QueryCommand(Me.SQL_GET_TRIP_DATE, varBuffer, lngHowMany, DS.Fields("0").NewValue)
            strTripDate = .MakeSQL(varBuffer(0, 0), dtDate)
            strQ = strQ & ", octdtcpi = " & strTripDate
            strQ = strQ & ", octdtcef = " & strTripDate
        End If
        
        strTS = DS.AddTSUpdate
        If strTS <> vbNullString Then
            strQ = strQ & ", " & strTS
        End If
        strQ = strQ & " WHERE octidoct = " & .MakeSQL(DS.Fields("37").Value, DS.Fields("37").FldType)
        If Not .ExecuteQuery(strQ) Then
            GoTo ErrHandler
        End If
        If Avanzamento = "V" Or Avanzamento = "F" Then
            strQ = SQL_SET_AVANZAMENTO_ORDINE_RIGHE & .MakeSQL(Avanzamento, 0)
            strTS = DS.AddTSUpdate
            If strTS <> vbNullString Then
                strQ = strQ & ", " & strTS
            End If
            strQ = strQ & " WHERE ocridoct = " & .MakeSQL(DS.Fields("37").Value, DS.Fields("37").FldType)
            strQ = strQ & "   AND ocrflstr = " & .MakeSQL("A", 0)
            strQ = strQ & "   AND ocrannul = " & .MakeSQL(False, 3)
            
            If Not .ExecuteQuery(strQ) Then
                GoTo ErrHandler
            End If
        End If
        .EndTrans 0
    End With
    SetAvanzamentoOrdine = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    SetAvanzamentoOrdine = False
End Function

Public Sub GetTotali()
    Dim strQ As String
    Dim varResult As Variant
    Dim blnOk As Boolean
    Dim lngHowMany As Long
    Dim dblTot As Double
    
    With DS.DA
        strQ = SQL_GET_TOTALE_QTA & .MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
'        strQ = strQ & " AND vvcidadz = " & DS.DA.MakeSQL(DS.Fields("12").Value, DS.Fields("12").FldType)
        strQ = strQ & " AND vvpidvvc = " & DS.DA.MakeSQL(DS.Fields("1").Value, DS.Fields("1").FldType)
        blnOk = .QueryData(strQ, varResult, lngHowMany)
        If lngHowMany > 0 Then
            mvarTotaliQta = varResult
        End If
    End With
End Sub


Public Sub GetTotaliDestinazione()
    Dim strQ As String
    Dim varResult As Variant
    Dim blnOk As Boolean
    Dim lngHowMany As Long
    Dim dblTot As Double
    
    With DS.DA
        strQ = SQL_GET_TOTALE_QTA & .MakeSQL(DS.Fields("0").Value, DS.Fields("0").FldType)
        strQ = strQ & " AND vvcidadz = " & DS.DA.MakeSQL(DS.Fields("12").Value, DS.Fields("12").FldType)
        blnOk = .QueryData(strQ, varResult, lngHowMany)
        If lngHowMany > 0 Then
            mvarTotaliQta = varResult
        End If
    End With
End Sub


Public Property Get TotaliQta(Index As Integer) As Variant
    If Not IsEmpty(mvarTotaliQta) Then
        If (Index >= LBound(mvarTotaliQta, 1)) And _
            (Index <= UBound(mvarTotaliQta, 1)) Then
            TotaliQta = mvarTotaliQta(Index, 0)
        End If
    End If
End Property

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
    
    blnSaveOk = DS.Insert(Me.SQL_INSERT)
    If Not blnSaveOk Then
        Insert = False
        Exit Function
    End If
    
    VVT.DS.Fields("0").FindValue = DS.Fields("0").NewValue
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
    blnSaveOk = SetAvanzamentoOrdine("V")
    
    Insert = blnSaveOk
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean
    Dim strDataConsegna As String
    Dim blnSetDateEff As Boolean
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
    
    ' --- Se modifico i Km calcolati automaticamente
    ' --- setto il Flag calcolo Km
    If DS.Fields("20").Updated And DS.Fields("21").Value = "A" Then
        DS.Fields("21").NewValue = "M"
        DS.Fields("21").Updated = True
    End If
    If DS.Fields("22").Updated And DS.Fields("23").Value = "A" Then
        DS.Fields("23").NewValue = "M"
        DS.Fields("23").Updated = True
    End If
    ' ---
    
    #If VersionID = 1 Then
        If DS.Fields("41").Updated Then
            strDataConsegna = DS.Fields("41").NewValue
            blnSetDateEff = True
        End If
    #End If

    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    VVT.DS.Fields("0").FindValue = DS.Fields("0").Value
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
    #If VersionID = 1 Then
        ' --- Se la data consegna è cambiata modifico la data consegna
        ' --- effettiva dell'ordine
        If blnSetDateEff Then
            blnSaveOk = DS.DA.ExecuteCommand(SQL_SET_DATE_CONSEGNE_EFF, _
                                             strDataConsegna, _
                                             DS.Fields("0").Value, _
                                             DS.Fields("1").Value)
        End If
        ' ---
    #End If
    
    
    BeginSave = True
End Function

Public Function UpdateKmDef(strIDVVT As String, _
                            strIDADZ As String, _
                            lngKm As Long) As Boolean
                            
    On Error GoTo ErrHandler
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    If Not DS.DA.ExecuteCommand(Me.SQL_SET_KM_DEF, lngKm, strIDVVT, strIDADZ) Then
        GoTo ErrHandler
    End If
    
    VVT.DS.Fields("0").FindValue = strIDVVT
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
    UpdateKmDef = True
    
    Exit Function
ErrHandler:
    UpdateKmDef = False
    
End Function

Private Sub Class_Initialize()

    #If VersionID = 1 Then
        SQL_FIND = "SELECT vvcidvvt, vvcidvvc, vvcnrseq, vvcidasc, " & _
                   "       vvcidacl, vvcrasoc, vvcindir, vvclocal, " & _
                   "       vvcprovi, vvcnrcap, vvccofis, vvcpaiva, " & _
                   "       vvcidadz, vvcderas, vvcdeind, vvcdeloc, " & _
                   "       vvcdepro, vvcdecap, vvcdecfi, vvcdepiv, " & _
                   "       vvckmpap, vvcflcpp, vvckmpad, vvcflcpd, " & _
                   "       vvctemvg, vvctemsc, vvcidazo, vvcidatv, " & _
                   "       vvcidvt1, vvcidvc1, vvcries1, vvcries2, " & _
                   "       vvcflst1, vvcflst2, vvcflst3, vvcflst4, " & _
                   "       vvcflst5, vvcidoct, vvckmrag, vvckmras, " & _
                   "       vvcdtcon, vvcannul, timestmp FROM viagvvc0 WHERE "
    #Else
        SQL_FIND = "SELECT vvcidvvt, vvcidvvc, vvcnrseq, vvcidasc, " & _
                   "       vvcidacl, vvcrasoc, vvcindir, vvclocal, " & _
                   "       vvcprovi, vvcnrcap, vvccofis, vvcpaiva, " & _
                   "       vvcidadz, vvcderas, vvcdeind, vvcdeloc, " & _
                   "       vvcdepro, vvcdecap, vvcdecfi, vvcdepiv, " & _
                   "       vvckmpap, vvcflcpp, vvckmpad, vvcflcpd, " & _
                   "       vvctemvg, vvctemsc, vvcidazo, vvcidatv, " & _
                   "       vvcidvt1, vvcidvc1, vvcries1, vvcries2, " & _
                   "       vvcflst1, vvcflst2, vvcflst3, vvcflst4, " & _
                   "       vvcflst5, vvcidoct, vvcannul, timestmp FROM viagvvc0 WHERE "
    #End If
    
    #If VersionID = 4 Then
        SQL_GETLIST = "SELECT vvcidvvc, vvcnrseq, vvcidoct, vvcderas, vvcannul, vvctemvg, vvcidadz, vvcidatv, vvcidacl, vvckmpap,vvcidoct FROM viagvvc0 "
    #Else
        SQL_GETLIST = "SELECT vvcidvvc, vvcnrseq, vvcidoct, vvcderas, vvcannul, vvctemvg, vvcidadz, vvcidatv, vvcidacl, vvckmpad,vvcidoct FROM viagvvc0 "
    #End If
    
    SQL_UPDATE = "UPDATE viagvvc0 SET "
    SQL_INSERT = "INSERT INTO viagvvc0 "
    SQL_DELETE = "DELETE FROM viagvvc0 WHERE vvcidvvt = "
    SQL_DELETE_PRODOTTI = "DELETE FROM viagvvp0 WHERE vvpidvvt = "
    SQL_GET_TS = "SELECT timestmp FROM viagvvc0 WHERE "
    SQL_GET_TRIP_STATE = "SELECT vvtiddef, vvtflst5, vvtflstv FROM viagvvt0 WHERE vvtidvvt = "
    SQL_GET_TRIP_DATE = "SELECT vvtdtinv FROM viagvvt0 WHERE vvtidvvt = ? "
    SQL_SET_AVANZAMENTO_ORDINE = "UPDATE ordioct0 SET octflaor = "
    SQL_SET_AVANZAMENTO_ORDINE_RIGHE = "UPDATE ordiocr0 SET ocrflavr = "
'    SQL_GET_TOTALE_QTA = "SELECT SUM(vvpqtrkg),SUM(vvpqtrlt),SUM(vvpqtrpz),SUM(vvpqtrcn), " & _
'                         "       SUM(vvpqcokg),SUM(vvpqcolt),SUM(vvpqcopz),SUM(vvpqcocn) " & _
'                         "  FROM viagvvp0 WHERE vvpidvvt = "
    SQL_GET_TOTALE_QTA = "SELECT SUM(vvpqtrkg),SUM(vvpqtrlt),SUM(vvpqtrpz),SUM(vvpqtrcn), " & _
                         "       SUM(vvpqcokg),SUM(vvpqcolt),SUM(vvpqcopz),SUM(vvpqcocn) " & _
                         "  FROM VIAGVVC0 INNER JOIN viagvvp0 ON (VIAGVVC0.VVCIDVVC = viagvvp0.VVPIDVVC) " & _
                         "  AND (VIAGVVC0.VVCIDVVT = viagvvp0.VVPIDVVT) WHERE vvpidvvt = "
    SQL_GET_ORDINE = "SELECT octidasc, octidacl, aclrasoc, aclindir, " & _
                     "       acllocal, aclprovi, aclnrcap, aclcofis, " & _
                     "       aclpaiva, octidadz, adzrasoc, adzindir, " & _
                     "       adzlocal, adzprovi, adznrcap, adzcofis, " & _
                     "       adzpaiva, octidazo, octidatv, octidoct, " & _
                     "       octrife1, octrife2, octkmrag, octraut3 " & _
                     "       FROM (ordioct0 INNER JOIN anagacl0 ON ordioct0.octidacl = anagacl0.aclidacl) " & _
                     "       INNER JOIN anagadz0 ON ordioct0.octidadz = anagadz0.adzidadz " & _
                     "       WHERE octidoct = "
    SQL_SET_KM_DEF = "UPDATE viagvvc0 SET vvckmpad = ?, vvcflcpd = 'A' WHERE vvcidvvt = ? AND vvcidadz = ?"
    SQL_SET_DATE_CONSEGNE_EFF = " UPDATE (VIAGVVC0 INNER JOIN VIAGVVP0 ON " & _
                                " VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC) INNER JOIN " & _
                                " ORDIOCT0 ON VIAGVVP0.VVPIDOCT = ORDIOCT0.OCTIDOCT " & _
                                " SET ORDIOCT0.OCTDTCEF = ? WHERE ( (VIAGVVP0.VVPIDVVT=?) AND (VIAGVVP0.VVPIDVVC=?) )"
    SQL_GET_ORDERS_LIST = " SELECT OCTIDOCT,OCTIDADZ,ADZRASOC,ADZLOCAL,OCTDTCRI, Sum(OCRQRILT) AS SommaDiOCRQRILT " & _
                          " FROM (ANAGADZ0 INNER JOIN ORDIOCT0 ON ADZIDADZ = OCTIDADZ) INNER JOIN ORDIOCR0 ON OCTIDOCT = OCRIDOCT " & _
                          " WHERE OCTFLAOR='A' OR OCTFLAOR='M'     " & _
                          " GROUP BY OCTIDOCT,OCTIDADZ,ADZRASOC,ADZLOCAL,OCTDTCRI " & _
                          " ORDER BY OCTDTCRI,ADZLOCAL "

    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Viaggio", "viagvvc0.vvcidvvt", 0, True, 7, "0"
        .Add "Numero Consegna", "viagvvc0.vvcidvvc", 0, True, 3, "1"
        .Add "Numero Sequenza", "viagvvc0.vvcnrseq", 0, True, 3, "2"
        .Add "Società", "viagvvc0.vvcidasc", 0, True, 3, "3", , True
        With .Item("3")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Cliente", "viagvvc0.vvcidacl", 0, True, 10, "4", , True
        With .Item("4")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Ragione Sociale", "viagvvc0.vvcrasoc", 0, False, 40, "5"
        .Add "Indirizzo", "viagvvc0.vvcindir", 0, False, 40, "6"
        .Add "Località", "viagvvc0.vvclocal", 0, False, 25, "7"
        .Add "Provincia", "viagvvc0.vvcprovi", 0, False, 2, "8"
        .Add "CAP", "viagvvc0.vvcnrcap", 0, False, 10, "9"
        .Add "Codice Fiscale", "viagvvc0.vvccofis", 0, False, 16, "10"
        .Add "Partita IVA", "viagvvc0.vvcpaiva", 0, False, 20, "11"
        .Add "Destinazione", "viagvvc0.vvcidadz", 0, True, 10, "12", , True
        With .Item("12")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Ragione Sociale Destinazione", "viagvvc0.vvcderas", 0, False, 40, "13"
        .Add "Indirizzo Destinazione", "viagvvc0.vvcdeind", 0, False, 40, "14"
        .Add "Località Destinazione", "viagvvc0.vvcdeloc", 0, False, 25, "15"
        .Add "Provincia Destinazione", "viagvvc0.vvcdepro", 0, False, 2, "16"
        .Add "CAP Destinazione", "viagvvc0.vvcdecap", 0, False, 10, "17"
        .Add "Codice Fiscale Destinazione", "viagvvc0.vvcdecfi", 0, False, 16, "18"
        .Add "Partita IVA Destinazione", "viagvvc0.vvcdepiv", 0, False, 20, "19"
        .Add "Km previsti", "viagvvc0.vvckmpap", 1, False, 6, "20"
        .Add "Flag calcolo km previsti", "viagvvc0.vvcflcpp", 0, True, 1, "21"
        .Add "Km definitivi", "viagvvc0.vvckmpad", 1, False, 6, "22"
        .Add "Flag calcolo km definitivi", "viagvvc0.vvcflcpd", 0, True, 1, "23"
'        .Add "Tempo viaggio", "viagvvc0.vvctemvg", 1, False, 6, "24"
        .Add "Orario arrivo", "viagvvc0.vvctemvg", 4, False, 5, "24"
        .Add "Tempo scarico", "viagvvc0.vvctemsc", 1, False, 6, "25"
        .Add "Zona", "viagvvc0.vvcidazo", 0, True, 3, "26", , True
        With .Item("26")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Tipo Vendita", "viagvvc0.vvcidatv", 0, False, 3, "27", , True
        With .Item("27")
            .ForeignTableEntity = "ATV"
            .ForeignTable = "anagatv0"
            .ForeignTableFld = "atvidatv"
            .ForeignTableDescField = "atvdescr"
            .ForeignTableDeleteFld = "atvannul"
        End With
        .Add "Viaggio di Riferimento", "viagvvc0.vvcidvt1", 0, False, 7, "28", , True
        .Add "Consegna di Riferimento", "viagvvc0.vvcidvc1", 0, False, 3, "29", , True
        .Add "Riferimento Esterno 1", "viagvvc0.vvcries1", 0, False, 15, "30"
        .Add "Riferimento Esterno 2", "viagvvc0.vvcries2", 0, False, 15, "31"
        .Add "Flag stato 1", "viagvvc0.vvcflst1", 0, False, 2, "32"
        .Add "Flag stato 2", "viagvvc0.vvcflst2", 0, False, 2, "33"
        ' --- Consegna per ordine programmato
        .Add "Flag stato 3", "viagvvc0.vvcflst3", 0, False, 2, "34", , True
        ' ---
        .Add "Flag stato 4", "viagvvc0.vvcflst4", 0, False, 2, "35"
        .Add "Flag stato 5", "viagvvc0.vvcflst5", 0, False, 2, "36"
        .Add "Ordine Cliente", "viagvvc0.vvcidoct", 0, True, 7, "37", , True
        #If VersionID = 1 Then
            .Add "Km raggio", "viagvvc0.vvckmrag", 1, False, 6, "39"
            .Add "Km raggio stradali", "viagvvc0.vvckmras", 1, False, 6, "40"
            .Add "Data consegna", "viagvvc0.vvcdtcon", 2, True, 10, "41"
        #End If
        .Add "Annullato", "viagvvc0.vvcannul", 3, , , "38"
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
  End With
    DS.CancelField = "viagvvc0.vvcannul"
End Sub

Public Function GetOrderDescription(varBuffer As Variant, HowManyRows As Long) As Integer
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    
    HowManyRows = 0
    blnResult = DS.DA.QueryData(SQL_GET_ORDERS_LIST, varBuffer, HowManyRows)
    
    If Not blnResult Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Or IsNull(varBuffer) Then
        GoTo ErrHandler
    End If
    GetOrderDescription = 1
    Exit Function
ErrHandler:
    GetOrderDescription = -1
End Function


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

