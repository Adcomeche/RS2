VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTariffaBinomica"
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
Public SQL_FIND_BY_DATE As String
Public SQL_FIND_BY_DATE_DEST As String
Public SQL_FIND_MAX_DATE As String
Public SQL_FIND_MAX_DATE_VETTORE As String
Public SQL_FIND_MAX_DATE_DEST As String
Public SQL_FIND_MAX_DATE_DEST_VETTORE As String
Public SQL_FIND_BY_DATE_2 As String
Public SQL_FIND_BY_DATE_ATS As String
Public SQL_FIND_MAX_DATE_2 As String
Public SQL_FIND_MAX_DATE_2_VETTORE As String
Public SQL_FIND_MAX_DATE_ATS As String
Public SQL_FIND_MAX_DATE_ATS_VETTORE As String
Public SQL_COPY_DETAIL As String
Public SQL_COPY_DETAIL_2 As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Public Function CopyDetail(NewIDDeposito As String, _
                           NewIDTariffa As String, _
                           FromIDDeposito As String, _
                           FromIDTariffa As String) As Boolean
    
    On Error GoTo ErrHandler
    Dim strSQL As String
    
    strSQL = SQL_COPY_DETAIL & " SELECT " & DS.DA.MakeSQL(NewIDDeposito, dtAlfa) & _
                               ", " & DS.DA.MakeSQL(NewIDTariffa, dtAlfa) & _
                               ", " & SQL_COPY_DETAIL_2
    If Not DS.DA.ExecuteCommand(strSQL, _
                                FromIDDeposito, FromIDTariffa) Then
        GoTo ErrHandler
    End If
    CopyDetail = True
    Exit Function
ErrHandler:
    CopyDetail = False
End Function


Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function FindByDate(IDDeposito As String, _
                           DataViaggio As String, _
                           TipoTariffa As String, _
                           IDContratto As String, _
                           IDSocietà As String, _
                           IDEnteEmittente As String, _
                           IDTipologiaProdotto As String, _
                           Optional IDDestinazione As String = vbNullString, _
                           Optional TipoCisterna As String = vbNullString, _
                           Optional IDVettore As String = vbNullString, _
                           Optional FlagVeicolo As String = vbNullString) As String
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim blnVettoreOK As Boolean
    Dim lngHowMany As Long
    Dim i As Long
    Dim j As Long
    Dim varMaxData As Variant
    
    If Len(TipoCisterna) = 0 Then
        If Len(IDDestinazione) > 0 Then
            If Len(IDVettore) > 0 And (TipoTariffa = "P") Then
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_DEST_VETTORE, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, IDDestinazione, IDSocietà, IDEnteEmittente, IDTipologiaProdotto, IDVettore)
            Else
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_DEST, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, IDDestinazione, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            End If
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
            If IsNull(varBuffer(0, 0)) Then
                If Len(IDVettore) > 0 And (TipoTariffa = "P") Then
                    blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_VETTORE, varBuffer, lngHowMany, _
                                                   IDDeposito, TipoTariffa, DataViaggio, IDSocietà, IDEnteEmittente, IDTipologiaProdotto, IDVettore)
                Else
                    blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE, varBuffer, lngHowMany, _
                                                   IDDeposito, TipoTariffa, DataViaggio, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
                End If
                If Not blnResult Then
                    GoTo ErrHandler
                End If
                If lngHowMany = 0 Then
                    GoTo ErrHandler
                End If
                If IsNull(varBuffer(0, 0)) Then
                    GoTo ErrHandler
                End If
                varMaxData = varBuffer(0, 0)
            Else
                varMaxData = varBuffer(0, 0)
            End If
        Else
            If Len(IDVettore) > 0 And (TipoTariffa = "P") Then
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_VETTORE, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, IDSocietà, IDEnteEmittente, IDTipologiaProdotto, IDVettore)
            Else
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            End If
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
            varMaxData = varBuffer(0, 0)
        End If
        If Len(IDDestinazione) > 0 Then
            lngHowMany = 0
            blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE_DEST, varBuffer, lngHowMany, _
                                           varMaxData, IDDeposito, TipoTariffa, IDDestinazione, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            If Not blnResult Then
                GoTo ErrHandler
            End If
        End If
        If (Len(IDDestinazione) = 0) Or (lngHowMany = 0) Then
            blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE, varBuffer, lngHowMany, _
                                           varMaxData, IDDeposito, TipoTariffa, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
        End If
    Else
        If Len(IDDestinazione) > 0 Then
            If Len(IDVettore) > 0 And (TipoTariffa = "P") Then
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_ATS_VETTORE, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, IDDestinazione, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto, IDVettore)
            Else
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_ATS, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, IDDestinazione, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            End If
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
            If IsNull(varBuffer(0, 0)) Then
                If Len(IDVettore) > 0 And (TipoTariffa = "P") Then
                    blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_2_VETTORE, varBuffer, lngHowMany, _
                                                   IDDeposito, TipoTariffa, DataViaggio, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto, IDVettore)
                Else
                    blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_2, varBuffer, lngHowMany, _
                                                   IDDeposito, TipoTariffa, DataViaggio, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
                End If
                If Not blnResult Then
                    GoTo ErrHandler
                End If
                If lngHowMany = 0 Then
                    GoTo ErrHandler
                End If
                If IsNull(varBuffer(0, 0)) Then
                    GoTo ErrHandler
                End If
                varMaxData = varBuffer(0, 0)
            Else
                varMaxData = varBuffer(0, 0)
            End If
        Else
            If Len(IDVettore) > 0 And (TipoTariffa = "P") Then
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_2_VETTORE, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto, IDVettore)
            Else
                blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_2, varBuffer, lngHowMany, _
                                               IDDeposito, TipoTariffa, DataViaggio, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            End If
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
            varMaxData = varBuffer(0, 0)
        End If
        If Len(IDDestinazione) > 0 Then
            lngHowMany = 0
            blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE_ATS, varBuffer, lngHowMany, _
                                           varMaxData, IDDeposito, TipoTariffa, IDDestinazione, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            If Not blnResult Then
                GoTo ErrHandler
            End If
        End If
        If (Len(IDDestinazione) = 0) Or (lngHowMany = 0) Then
            blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE_2, varBuffer, lngHowMany, _
                                           varMaxData, IDDeposito, TipoTariffa, TipoCisterna, IDSocietà, IDEnteEmittente, IDTipologiaProdotto)
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
        End If
    End If
    
    If (TipoTariffa = "P") Then 'And (UBound(varBuffer, 2) > 0) Then
        If (IDVettore = vbNullString) And (FlagVeicolo = vbNullString) Then
            FindByDate = varBuffer(0, 0)
        Else
            If IDVettore = vbNullString Then
                blnVettoreOK = True
            Else
                blnVettoreOK = False
            End If
            For i = 0 To UBound(varBuffer, 2)
                If varBuffer(2, i) = IDVettore Then
                    If IDVettore <> vbNullString Then
                        blnVettoreOK = True
                    End If
                    If FlagVeicolo <> vbNullString Then
                        For j = i To UBound(varBuffer, 2)
                            If varBuffer(3, j) = FlagVeicolo Then
                                Exit For
                            End If
                        Next j
                        Exit For
                    Else
                        j = i
                        Exit For
                    End If
                End If
            Next i
            If blnVettoreOK Then
                FindByDate = varBuffer(0, j)
            Else
                GoTo ErrHandler
            End If
        End If
    Else
        FindByDate = varBuffer(0, 0)
    End If
    Exit Function
ErrHandler:
    FindByDate = vbNullString
End Function


Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT tfbidade, tfbidtfb, tfbdescr, tfbdtvin, " & _
               "       tfbfltit, tfbidadz, tfbidacn, tfbidats, " & _
               "       tfbfludm, tfbidasc, tfbidaee, tfbidttp, " & _
               "       tfbidave, tfbflmot, tfbannul, timestmp FROM taritfb0 WHERE "
    SQL_GETLIST = "SELECT tfbidtfb, tfbdescr FROM taritfb0 "
    SQL_UPDATE = "UPDATE taritfb0 SET "
    SQL_INSERT = "INSERT INTO taritfb0 "
    SQL_GET_TS = "SELECT timestmp FROM taritfb0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT tfbidade, tfbidtfb, tfbdescr, tfbdtvin, tfbannul FROM taritfb0 "
    SQL_FIND_BY_DATE = "SELECT tfbidtfb, tfbdtvin, tfbidave, tfbflmot " & _
                       "  FROM taritfb0 " & _
                       " WHERE ( tfbdtvin=? AND tfbannul=0 AND tfbidade= ? AND tfbfltit= ? and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?)  ) and ((tfbidadz is null) or (LEN(tfbidadz)=0)  and ((tfbidats is null) or (LEN(tfbidats)=0) ) ) " & _
                       " ORDER BY tfbidave, tfbflmot "
    SQL_FIND_BY_DATE_DEST = "SELECT tfbidtfb, tfbdtvin, tfbidave, tfbflmot " & _
                       "  FROM taritfb0 " & _
                       " WHERE ( tfbdtvin=? AND tfbannul=0 AND tfbidade= ? AND tfbfltit= ? AND tfbidadz = ? and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) and ((tfbidats is null) or (LEN(tfbidats)=0) ) ) " & _
                       " ORDER BY tfbidave, tfbflmot "
                       
    SQL_FIND_MAX_DATE = "SELECT max(tfbdtvin) " & _
                        "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                        "   AND tfbfltit= ? AND (tfbdtvin<=?) and (tfbidasc = ?)  and (tfbidaee = ?) and (tfbidttp = ?) and ((tfbidadz is null) or (LEN(tfbidadz)=0) ) and ((tfbidats is null) or (LEN(tfbidats)=0) ) "
    SQL_FIND_MAX_DATE_VETTORE = "SELECT max(tfbdtvin) " & _
                        "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                        "   AND tfbfltit= ? AND (tfbdtvin<=?) and (tfbidasc = ?)  and (tfbidaee = ?) and (tfbidttp = ?)  and (tfbidave = ?) and ((tfbidadz is null) or (LEN(tfbidadz)=0) ) and ((tfbidats is null) or (LEN(tfbidats)=0) ) "
    
    SQL_FIND_MAX_DATE_DEST = "SELECT max(tfbdtvin) " & _
                             "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                             "   AND tfbfltit= ? AND (tfbdtvin<=?) AND (tfbidadz = ?) and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) and ((tfbidats is null) or (LEN(tfbidats)=0) ) "
    SQL_FIND_MAX_DATE_DEST_VETTORE = "SELECT max(tfbdtvin) " & _
                                     "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                                     "   AND tfbfltit= ? AND (tfbdtvin<=?) AND (tfbidadz = ?) and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) and (tfbidave = ?) and ((tfbidats is null) or (LEN(tfbidats)=0) ) "
    
    SQL_FIND_BY_DATE_2 = "SELECT tfbidtfb, tfbdtvin, tfbidave, tfbflmot " & _
                       "  FROM taritfb0 " & _
                       " WHERE ( tfbdtvin=? AND tfbannul=0 AND tfbidade= ? AND tfbfltit= ?  ) and ((tfbidadz is null) or (LEN(tfbidadz)=0)  and (tfbidats = ?)  and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) ) " & _
                       " ORDER BY tfbidave, tfbflmot "
    SQL_FIND_BY_DATE_ATS = "SELECT tfbidtfb, tfbdtvin, tfbidave, tfbflmot " & _
                       "  FROM taritfb0 " & _
                       " WHERE ( tfbdtvin=? AND tfbannul=0 AND tfbidade= ? AND tfbfltit= ? AND tfbidadz = ?  and tfbidats = ?  and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) ) " & _
                       " ORDER BY tfbidave, tfbflmot "
                       
    SQL_FIND_MAX_DATE_2 = "SELECT max(tfbdtvin) " & _
                        "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                        "   AND tfbfltit= ? AND (tfbdtvin<=?) and ((tfbidadz is null) or (LEN(tfbidadz)=0) ) and (tfbidats = ?)  and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) "
    SQL_FIND_MAX_DATE_2_VETTORE = "SELECT max(tfbdtvin) " & _
                        "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                        "   AND tfbfltit= ? AND (tfbdtvin<=?) and ((tfbidadz is null) or (LEN(tfbidadz)=0) ) and (tfbidats = ?)  and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?)  and (tfbidave = ?) "
    SQL_FIND_MAX_DATE_ATS = "SELECT max(tfbdtvin) " & _
                             "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                             "   AND tfbfltit= ? AND (tfbdtvin<=?) AND (tfbidadz = ?) and (tfbidats = ?)  and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) "
    SQL_FIND_MAX_DATE_ATS_VETTORE = "SELECT max(tfbdtvin) " & _
                             "  FROM taritfb0 WHERE (tfbidade = ?) and (tfbannul=0)  " & _
                             "   AND tfbfltit= ? AND (tfbdtvin<=?) AND (tfbidadz = ?) and (tfbidats = ?)  and (tfbidasc = ?) and (tfbidaee = ?) and (tfbidttp = ?) and (tfbidave = ?) "
                       
    SQL_COPY_DETAIL = _
    " INSERT INTO TARITBR0 (TBRIDADE, TBRIDTFB, TBRPROGD, TBRINFKM, TBRSUPKM, " & _
    "                       TBRLIN01, TBRLSU01, TBRTOL01, TBRIDV01, TBRIMB01, " & _
    "                       TBRLIN02, TBRLSU02, TBRTOL02, TBRIDV02, TBRIMB02, " & _
    "                       TBRLIN03, TBRLSU03, TBRTOL03, TBRIDV03, TBRIMB03, " & _
    "                       TBRLIN04, TBRLSU04, TBRTOL04, TBRIDV04, TBRIMB04, " & _
    "                       TBRLIN05, TBRLSU05, TBRTOL05, TBRIDV05, TBRIMB05, " & _
    "                       TBRLIN06, TBRLSU06, TBRTOL06, TBRIDV06, TBRIMB06, " & _
    "                       TBRLIN07, TBRLSU07, TBRTOL07, TBRIDV07, TBRIMB07, " & _
    "                       TBRLIN08, TBRLSU08, TBRTOL08, TBRIDV08, TBRIMB08, " & _
    "                       TBRLIN09, TBRLSU09, TBRTOL09, TBRIDV09, TBRIMB09, " & _
    "                       TBRLIN10, TBRLSU10, TBRTOL10, TBRIDV10, TBRIMB10, " & _
    "                       TBRLIN11, TBRLSU11, TBRTOL11, TBRIDV11, TBRIMB11, " & _
    "                       TBRLIN12, TBRLSU12, TBRTOL12, TBRIDV12, TBRIMB12, " & _
    "                       TBRLIN13, TBRLSU13, TBRTOL13, TBRIDV13, TBRIMB13, " & _
    "                       TBRLIN14, TBRLSU14, TBRTOL14, TBRIDV14, TBRIMB14, " & _
    "                       TBRLIN15, TBRLSU15, TBRTOL15, TBRIDV15, TBRIMB15, " & _
    "                       TBRTIPOP, "
    SQL_COPY_DETAIL = SQL_COPY_DETAIL & _
        "TBRIMP01, TBRIMP02, TBRIMP03, TBRIMP04, TBRIMP05, TBRIMP06, TBRIMP07, TBRIMP08, " & _
        "TBRIMP09, TBRIMP10, TBRIMP11, TBRIMP12, TBRIMP13, TBRIMP14, TBRIMP15)"
    
    SQL_COPY_DETAIL_2 = _
    "                            TBRPROGD, TBRINFKM, TBRSUPKM, " & _
    "                       TBRLIN01, TBRLSU01, TBRTOL01, TBRIDV01, TBRIMB01, " & _
    "                       TBRLIN02, TBRLSU02, TBRTOL02, TBRIDV02, TBRIMB02, " & _
    "                       TBRLIN03, TBRLSU03, TBRTOL03, TBRIDV03, TBRIMB03, " & _
    "                       TBRLIN04, TBRLSU04, TBRTOL04, TBRIDV04, TBRIMB04, " & _
    "                       TBRLIN05, TBRLSU05, TBRTOL05, TBRIDV05, TBRIMB05, " & _
    "                       TBRLIN06, TBRLSU06, TBRTOL06, TBRIDV06, TBRIMB06, " & _
    "                       TBRLIN07, TBRLSU07, TBRTOL07, TBRIDV07, TBRIMB07, " & _
    "                       TBRLIN08, TBRLSU08, TBRTOL08, TBRIDV08, TBRIMB08, " & _
    "                       TBRLIN09, TBRLSU09, TBRTOL09, TBRIDV09, TBRIMB09, " & _
    "                       TBRLIN10, TBRLSU10, TBRTOL10, TBRIDV10, TBRIMB10, " & _
    "                       TBRLIN11, TBRLSU11, TBRTOL11, TBRIDV11, TBRIMB11, " & _
    "                       TBRLIN12, TBRLSU12, TBRTOL12, TBRIDV12, TBRIMB12, " & _
    "                       TBRLIN13, TBRLSU13, TBRTOL13, TBRIDV13, TBRIMB13, " & _
    "                       TBRLIN14, TBRLSU14, TBRTOL14, TBRIDV14, TBRIMB14, " & _
    "                       TBRLIN15, TBRLSU15, TBRTOL15, TBRIDV15, TBRIMB15, " & _
    "        TBRTIPOP, "
    
    SQL_COPY_DETAIL_2 = SQL_COPY_DETAIL_2 & _
        "TBRIMP01, TBRIMP02, TBRIMP03, TBRIMP04, TBRIMP05, TBRIMP06, TBRIMP07, TBRIMP08, " & _
        "TBRIMP09, TBRIMP10, TBRIMP11, TBRIMP12, TBRIMP13, TBRIMP14, TBRIMP15" & _
        "   FROM TARITBR0 " & _
        "  WHERE (((TARITBR0.TBRIDADE)=?) AND ((TARITBR0.TBRIDTFB)=?))"

    mvarFieldsIndexes = Array("0", "1", "2", "3", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Deposito", "taritfb0.tfbidade", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Codice", "taritfb0.tfbidtfb", 0, True, 3, "1"
        .Add "Descrizione", "taritfb0.tfbdescr", 0, True, 25, "2"
        .Add "Data validità", "taritfb0.tfbdtvin", 2, True, 10, "3"
        .Add "Flag tipo tariffa", "taritfb0.tfbfltit", 0, True, 1, "9"
        .Add "Destinazione", "taritfb0.tfbidadz", 0, False, 10, "10"
        With .Item("10")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Contratto", "taritfb0.tfbidacn", 0, False, 3, "11"
        With .Item("11")
            .ForeignTableEntity = "ACN"
            .ForeignTable = "anagacn0"
            .ForeignTableFld = "acnidacn"
            .ForeignTableDescField = "acndescr"
            .ForeignTableDeleteFld = "acnannul"
        End With
        .Add "Tipo Cisterna", "taritfb0.tfbidats", 0, , 3, "12"
        With .Item("12")
            .ForeignTableEntity = "ATS"
            .ForeignTable = "anagats0"
            .ForeignTableFld = "atsidats"
            .ForeignTableDescField = "atsdescr"
            .ForeignTableDeleteFld = "atsannul"
        End With
        .Add "Flag unità di misura", "taritfb0.tfbfludm", 0, False, 1, "13"
        .Add "Società", "taritfb0.tfbidasc", 0, True, 3, "14"
        With .Item("14")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "taritfb0.tfbidaee", 0, True, 3, "15"
        With .Item("15")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Tipologia Prodotto", "taritfb0.tfbidttp", 0, True, 3, "16"
        With .Item("16")
            .ForeignTableEntity = "TTP"
            .ForeignTable = "anagttp0"
            .ForeignTableFld = "ttpidttp"
            .ForeignTableDescField = "ttpdescr"
            .ForeignTableDeleteFld = "ttpannul"
        End With
        .Add "Vettore", "taritfb0.tfbidave", 0, False, 3, "17"
        With .Item("17")
            .ForeignTableEntity = "AVE"
            .ForeignTable = "anagave0"
            .ForeignTableFld = "aveidave"
            .ForeignTableDescField = "averasoc"
            .ForeignTableDeleteFld = "aveannul"
        End With
        .Add "Flag Veicolo", "taritfb0.tfbflmot", 0, False, 2, "18"
        .Add "Annullato", "taritfb0.tfbannul", 3, , , "8"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
    End With
    
    DS.CancelField = "taritfb0.tfbannul"
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


