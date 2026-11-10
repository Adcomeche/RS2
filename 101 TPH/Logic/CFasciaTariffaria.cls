VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFasciaTariffaria"
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
Public SQL_FIND_MAX_DATE_DEST As String
Public SQL_COPY_DETAIL As String
Public SQL_COPY_DETAIL_2 As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Public Function CopyDetail(NewIDZona As String, _
                           NewIDTariffa As String, _
                           FromIDZona As String, _
                           FromIDTariffa As String) As Boolean
    
    On Error GoTo ErrHandler
    Dim strSQL As String
    
    strSQL = SQL_COPY_DETAIL & " SELECT " & DS.DA.MakeSQL(NewIDZona, dtAlfa) & _
                               ", " & DS.DA.MakeSQL(NewIDTariffa, dtAlfa) & _
                               ", " & SQL_COPY_DETAIL_2
    If Not DS.DA.ExecuteCommand(strSQL, _
                                FromIDZona, FromIDTariffa) Then
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

Public Function FindByDate(IDZona As String, _
                           DataViaggio As String, _
                           TipoTariffa As String, _
                           Optional IDDestinazione As String = vbNullString) As String
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    Dim varMaxData As Variant
    
    If Len(IDDestinazione) > 0 Then
        blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE_DEST, varBuffer, lngHowMany, _
                                       IDZona, TipoTariffa, DataViaggio, IDDestinazione)
        
        If Not blnResult Then
            GoTo ErrHandler
        End If
        If lngHowMany = 0 Then
            GoTo ErrHandler
        End If
        If IsNull(varBuffer(0, 0)) Then
            blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE, varBuffer, lngHowMany, _
                                           IDZona, TipoTariffa, DataViaggio)
            If Not blnResult Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
            varMaxData = varBuffer(0, 0)
        Else
            varMaxData = varBuffer(0, 0)
        End If
    Else
    
        blnResult = DS.DA.QueryCommand(SQL_FIND_MAX_DATE, varBuffer, lngHowMany, _
                                       IDZona, TipoTariffa, DataViaggio)
        If Not blnResult Then
            GoTo ErrHandler
        End If
        If lngHowMany = 0 Then
            GoTo ErrHandler
        End If
        varMaxData = varBuffer(0, 0)
    End If
    
'    If Not blnResult Then
'        GoTo ErrHandler
'    End If
'    If lngHowMany = 0 Then
'        GoTo ErrHandler
'    End If
'    varMaxData = varBuffer(0, 0)
    If Len(IDDestinazione) > 0 Then
        lngHowMany = 0
        blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE_DEST, varBuffer, lngHowMany, _
                                       varMaxData, IDZona, TipoTariffa, IDDestinazione)
        If Not blnResult Then
            GoTo ErrHandler
        End If
    End If
    If (Len(IDDestinazione) = 0) Or (lngHowMany = 0) Then
        blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE, varBuffer, lngHowMany, _
                                       varMaxData, IDZona, TipoTariffa)
        If Not blnResult Then
            GoTo ErrHandler
        End If
        If lngHowMany = 0 Then
            GoTo ErrHandler
        End If
    End If
    FindByDate = varBuffer(0, 0)
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
    
    SQL_FIND = "SELECT tfaidazo, tfaidtfa, tfadescr, tfadtvin, " & _
               "       tfaimin1, tfaidav1, tfaimin2, tfaidav2, " & _
               "       tfafltit, tfaidadz, tfaannul, timestmp FROM taritfa0 WHERE "
    SQL_GETLIST = "SELECT tfaidtfa, tfadescr FROM taritfa0 "
    SQL_UPDATE = "UPDATE taritfa0 SET "
    SQL_INSERT = "INSERT INTO taritfa0 "
    SQL_GET_TS = "SELECT timestmp FROM taritfa0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT tfaidazo, tfaidtfa, tfadescr, tfadtvin, tfaannul FROM taritfa0 "
'    SQL_FIND_BY_DATE = "SELECT TFAIDTFA, TFADTVIN " & _
'                       "  FROM TARITFA0 " & _
'                       " WHERE (((TFADTVIN)=(select max(tfadtvin) " & _
'                       "  FROM taritfa0 where (tfaidazo = ?) and (tfaannul=0)  " & _
'                       "   AND (tfadtvin<=?) )) " & _
'                       "   AND ((TFAANNUL)=0) AND ((TFAIDAZO)= ? )  AND ((TFAFLTIT)= ? ))"
    SQL_FIND_BY_DATE = "SELECT TFAIDTFA, TFADTVIN " & _
                       "  FROM TARITFA0 " & _
                       " WHERE ( TFADTVIN=? AND TFAANNUL=0 AND TFAIDAZO= ? AND TFAFLTIT= ? ) and ((tfaidadz is null) or (LEN(tfaidadz)=0) )"
    SQL_FIND_BY_DATE_DEST = "SELECT TFAIDTFA, TFADTVIN " & _
                       "  FROM TARITFA0 " & _
                       " WHERE ( TFADTVIN=? AND TFAANNUL=0 AND TFAIDAZO= ? AND TFAFLTIT= ? AND TFAIDADZ = ?)"
                       
    SQL_FIND_MAX_DATE = "SELECT max(tfadtvin) " & _
                        "  FROM taritfa0 WHERE (tfaidazo = ?) and (tfaannul=0)  " & _
                        "   AND TFAFLTIT= ? AND (tfadtvin<=?) and ((tfaidadz is null) or (LEN(tfaidadz)=0) )"
    SQL_FIND_MAX_DATE_DEST = "SELECT max(tfadtvin) " & _
                             "  FROM taritfa0 WHERE (tfaidazo = ?) and (tfaannul=0)  " & _
                             "   AND TFAFLTIT= ? AND (tfadtvin<=?) AND (tfaidadz = ?) "
                       
    SQL_COPY_DETAIL = _
    " INSERT INTO TARITFR0 (TFRIDAZO, TFRIDTFA, TFRPROGD, " & _
    "                       TFRINFKM, TFRSUPKM, TFRLIN01, TFRLSU01, TFRTOL01, " & _
    "                       TFRIDV01, TFRLIN02, TFRLSU02, TFRTOL02, TFRIDV02, " & _
    "                       TFRLIN03, TFRLSU03, TFRTOL03, TFRIDV03, TFRLIN04, " & _
    "                       TFRLSU04, TFRTOL04, TFRIDV04, TFRLIN05, TFRLSU05, " & _
    "                       TFRTOL05, TFRIDV05, TFRLIN06, TFRLSU06, TFRTOL06, " & _
    "                       TFRIDV06, TFRLIN07, TFRLSU07, TFRTOL07, TFRIDV07, " & _
    "                       TFRLIN08, TFRLSU08, TFRTOL08, TFRIDV08, TFRLIN09, " & _
    "                       TFRLSU09, TFRTOL09, TFRIDV09, TFRLIN10, TFRLSU10, " & _
    "                       TFRTOL10, TFRIDV10, TFRLIN11, TFRLSU11, TFRTOL11, " & _
    "                       TFRIDV11, TFRLIN12, TFRLSU12, TFRTOL12, TFRIDV12, " & _
    "                       TFRLIN13, TFRLSU13, TFRTOL13, TFRIDV13, TFRLIN14, " & _
    "                       TFRLSU14, TFRTOL14, TFRIDV14, TFRLIN15, TFRLSU15, " & _
    "                       TFRTOL15, TFRIDV15, TFRTIPOP, "
    SQL_COPY_DETAIL = SQL_COPY_DETAIL & _
        "TFRIMP01, TFRIMP02, TFRIMP03, TFRIMP04, TFRIMP05, TFRIMP06, TFRIMP07, TFRIMP08, " & _
        "TFRIMP09, TFRIMP10, TFRIMP11, TFRIMP12, TFRIMP13, TFRIMP14, TFRIMP15)"
    
'    " SELECT ? AS NuovaZona, ? AS NuovoCodice, TFRPROGD, TFRINFKM, " & _

    SQL_COPY_DETAIL_2 = _
    "                            TFRPROGD, TFRINFKM, " & _
    "        TFRSUPKM, TFRLIN01, TFRLSU01, TFRTOL01, TFRIDV01, TFRLIN02, " & _
    "        TFRLSU02, TFRTOL02, TFRIDV02, TFRLIN03, TFRLSU03, TFRTOL03, " & _
    "        TFRIDV03, TFRLIN04, TFRLSU04, TFRTOL04, TFRIDV04, TFRLIN05, " & _
    "        TFRLSU05, TFRTOL05, TFRIDV05, TFRLIN06, TFRLSU06, TFRTOL06, " & _
    "        TFRIDV06, TFRLIN07, TFRLSU07, TFRTOL07, TFRIDV07, TFRLIN08, " & _
    "        TFRLSU08, TFRTOL08, TFRIDV08, TFRLIN09, TFRLSU09, TFRTOL09, " & _
    "        TFRIDV09, TFRLIN10, TFRLSU10, TFRTOL10, TFRIDV10, TFRLIN11, " & _
    "        TFRLSU11, TFRTOL11, TFRIDV11, TFRLIN12, TFRLSU12, TFRTOL12, " & _
    "        TFRIDV12, TFRLIN13, TFRLSU13, TFRTOL13, TFRIDV13, TFRLIN14, " & _
    "        TFRLSU14, TFRTOL14, TFRIDV14, TFRLIN15, TFRLSU15, TFRTOL15, " & _
    "        TFRIDV15, TFRTIPOP, "
    
    SQL_COPY_DETAIL_2 = SQL_COPY_DETAIL_2 & _
        "TFRIMP01, TFRIMP02, TFRIMP03, TFRIMP04, TFRIMP05, TFRIMP06, TFRIMP07, TFRIMP08, " & _
        "TFRIMP09, TFRIMP10, TFRIMP11, TFRIMP12, TFRIMP13, TFRIMP14, TFRIMP15" & _
        "   FROM TARITFR0 " & _
        "  WHERE (((TARITFR0.TFRIDAZO)=?) AND ((TARITFR0.TFRIDTFA)=?))"

    mvarFieldsIndexes = Array("0", "1", "2", "3", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Zona", "taritfa0.tfaidazo", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Codice", "taritfa0.tfaidtfa", 0, True, 3, "1"
        .Add "Descrizione", "taritfa0.tfadescr", 0, True, 25, "2"
        .Add "Data validità", "taritfa0.tfadtvin", 2, True, 10, "3"
        .Add "Importo Tariffa minima 1", "taritfa0.tfaimin1", 1, True, 13, "4"
        .Add "Valuta Tariffa minima 1", "taritfa0.tfaidav1", 0, True, 3, "5"
        With .Item("5")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Importo Tariffa minima 2", "taritfa0.tfaimin2", 1, , 13, "6"
        .Add "Valuta Tariffa minima 2", "taritfa0.tfaidav2", 0, , 3, "7"
        With .Item("7")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Flag tipo tariffa", "taritfa0.tfafltit", 0, True, 1, "9"
        .Add "Destinazione", "taritfa0.tfaidadz", 0, False, 10, "10"
        With .Item("10")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Annullato", "taritfa0.tfaannul", 3, , , "8"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
    End With
    
    DS.CancelField = "taritfa0.tfaannul"
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


