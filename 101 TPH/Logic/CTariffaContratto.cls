VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTariffaContratto"
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
Public SQL_FIND_MAX_DATE_DEST_DEP As String
'Public SQL_FIND_MAX_DATE As String
'Public SQL_FIND_MAX_DATE_DEST As String
'Public SQL_FIND_BY_DATE_2 As String
'Public SQL_FIND_BY_DATE_ATS As String
'Public SQL_FIND_MAX_DATE_2 As String
'Public SQL_FIND_MAX_DATE_ATS As String
Public SQL_COPY_DETAIL As String
Public SQL_COPY_DETAIL_2 As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Public Function CopyDetail(NewIDACL As String, _
                           NewIDACN As String, _
                           NewIDADD As String, _
                           NewIDTFC As String, _
                           FromIDACL As String, _
                           FromIDACN As String, _
                           FromIDADD As String, _
                           FromIDTFC As String) As Boolean
    
    On Error GoTo ErrHandler
    Dim strSQL As String
    
    strSQL = SQL_COPY_DETAIL & " SELECT " & DS.DA.MakeSQL(NewIDACL, dtAlfa) & _
                               ", " & DS.DA.MakeSQL(NewIDACN, dtAlfa) & _
                               ", " & DS.DA.MakeSQL(NewIDADD, dtAlfa) & _
                               ", " & DS.DA.MakeSQL(NewIDTFC, dtAlfa) & _
                               ", " & SQL_COPY_DETAIL_2
    If Not DS.DA.ExecuteCommand(strSQL, _
                                FromIDACL, FromIDACN, FromIDADD, FromIDTFC) Then
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

Public Function FindByDate(IDCliente As String, _
                           IDContratto As String, _
                           IDAddebito As String, _
                           DataViaggio As String, _
                           ParametriTassazione As String, _
                           Optional IDDeposito As String = vbNullString, _
                           Optional IDDestinazione As String = vbNullString) As String
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim varBuffer2 As Variant
    Dim blnResult As Boolean
    Dim lngHowMany As Long
    Dim varMaxData As Variant
    Dim strQ As String
    Dim strQ2 As String
    
    strQ = SQL_FIND_MAX_DATE_DEST_DEP
    strQ2 = SQL_FIND_BY_DATE
    
    If IDDeposito <> vbNullString Then
        strQ = strQ & " and (tfcidade = " & DS.DA.MakeSQL(IDDeposito, dtAlfa) & ")"
        strQ2 = strQ2 & " and (tfcidade = " & DS.DA.MakeSQL(IDDeposito, dtAlfa) & ")"
    End If
    
    If IDDestinazione <> vbNullString Then
        strQ = strQ & " and (tfcidadz = " & DS.DA.MakeSQL(IDDestinazione, dtAlfa) & ")"
        strQ2 = strQ2 & " and (tfcidadz = " & DS.DA.MakeSQL(IDDestinazione, dtAlfa) & ")"
    End If
    
'    Select Case UCase(ParametriTassazione)
'        ' Deposito + destinazione
'        Case "DD"
            
            ' Determino la data di validità corrispondente alla data del viaggio
            If Not DS.DA.QueryCommand(strQ, varBuffer, lngHowMany, _
                                           IDCliente, IDContratto, IDAddebito, DataViaggio) Then
                GoTo ErrHandler
            End If
            
            If (lngHowMany = 0) Or IsNull(varBuffer(0, 0)) Then
                ' Se non trovo nulla con destinazione e deposito, cerco con la sola destinazione
                If IDDestinazione <> vbNullString Then
                    strQ = SQL_FIND_MAX_DATE_DEST_DEP & " and (tfcidadz = " & DS.DA.MakeSQL(IDDestinazione, dtAlfa) & ") and (tfcidade is null or tfcidade = '') "
                    strQ2 = SQL_FIND_BY_DATE & " and (tfcidadz = " & DS.DA.MakeSQL(IDDestinazione, dtAlfa) & ") and (tfcidade is null or tfcidade = '') "
                    If Not DS.DA.QueryCommand(strQ, varBuffer, lngHowMany, _
                                                   IDCliente, IDContratto, IDAddebito, DataViaggio) Then
                        GoTo ErrHandler
                    End If
                End If
                
                If (lngHowMany = 0) Or IsNull(varBuffer(0, 0)) Then
                    ' Se non trovo nulla con la destinazione, cerco con il solo deposito
                    If IDDeposito <> vbNullString Then
                        strQ = SQL_FIND_MAX_DATE_DEST_DEP & " and (tfcidade = " & DS.DA.MakeSQL(IDDeposito, dtAlfa) & ") and (tfcidadz is null or tfcidadz = '') "
                        strQ2 = SQL_FIND_BY_DATE & " and (tfcidade = " & DS.DA.MakeSQL(IDDeposito, dtAlfa) & ") and (tfcidadz is null or tfcidadz = '') "
                        If Not DS.DA.QueryCommand(strQ, varBuffer, lngHowMany, _
                                                       IDCliente, IDContratto, IDAddebito, DataViaggio) Then
                            GoTo ErrHandler
                        End If
                    End If
                    
                    If (lngHowMany = 0) Or IsNull(varBuffer(0, 0)) Then
                        ' Se non trovo nulla con il deposito, cerco senza parametri
                        ' evitando di ripetere la ricerca se non mi sono stati passati ne il deposito ne la destinazione
                        If (IDDeposito <> vbNullString) Or (IDDestinazione <> vbNullString) Then
                            strQ = SQL_FIND_MAX_DATE_DEST_DEP & " and (tfcidadz is null or tfcidadz = '') and (tfcidade is null or tfcidade = '') "
                            strQ2 = SQL_FIND_BY_DATE & " and (tfcidadz is null or tfcidadz = '') and (tfcidade is null or tfcidade = '') "
                            If Not DS.DA.QueryCommand(strQ, varBuffer, lngHowMany, _
                                                           IDCliente, IDContratto, IDAddebito, DataViaggio) Then
                                GoTo ErrHandler
                            End If
                            If (lngHowMany = 0) Or IsNull(varBuffer(0, 0)) Then
'                                varMaxData = varBuffer(0, 0)
'
'                                ' Ricerco la fascia tariffaria corrispondente alla data del viaggio
'                                blnResult = DS.DA.QueryCommand(strQ2, varBuffer2, lngHowMany, _
'                                                               IDCliente, IDContratto, IDAddebito, varMaxData)
'                                If Not blnResult Then
'                                    GoTo ErrHandler
'                                End If
                                GoTo ErrHandler
                            Else
                                varMaxData = varBuffer(0, 0)
                                
                                ' Ricerco la fascia tariffaria corrispondente alla data del viaggio
                                blnResult = DS.DA.QueryCommand(strQ2, varBuffer2, lngHowMany, _
                                                               IDCliente, IDContratto, IDAddebito, varMaxData)
                                If Not blnResult Then
                                    GoTo ErrHandler
                                End If
                            End If
                        End If
                    Else
                        varMaxData = varBuffer(0, 0)
                        
                        ' Ricerco la fascia tariffaria corrispondente alla data del viaggio
                        blnResult = DS.DA.QueryCommand(strQ2, varBuffer2, lngHowMany, _
                                                       IDCliente, IDContratto, IDAddebito, varMaxData)
                        If Not blnResult Then
                            GoTo ErrHandler
                        End If
                    End If
                Else
                    varMaxData = varBuffer(0, 0)
                    
                    ' Ricerco la fascia tariffaria corrispondente alla data del viaggio
                    blnResult = DS.DA.QueryCommand(strQ2, varBuffer2, lngHowMany, _
                                                   IDCliente, IDContratto, IDAddebito, varMaxData)
                    If Not blnResult Then
                        GoTo ErrHandler
                    End If
                End If
            Else
                varMaxData = varBuffer(0, 0)
                
                ' Ricerco la fascia tariffaria corrispondente alla data del viaggio
                blnResult = DS.DA.QueryCommand(strQ2, varBuffer2, lngHowMany, _
                                               IDCliente, IDContratto, IDAddebito, varMaxData)
                If Not blnResult Then
                    GoTo ErrHandler
                End If
            
            End If
            
            
            
            If IsNull(varBuffer2(0, 0)) Then
                GoTo ErrHandler
            End If
            
'            ' Ricerco la fascia tariffaria corrispondente alla data del viaggio
'            blnResult = DS.DA.QueryCommand(SQL_FIND_BY_DATE, varBuffer, lngHowMany, _
'                                           IDCliente, IDContratto, IDAddebito, varMaxData)
'            If Not blnResult Then
'                GoTo ErrHandler
'            End If
'            If lngHowMany = 0 Then
'                GoTo ErrHandler
'            End If
'    End Select
    
    FindByDate = varBuffer2(0, 0)
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
    
    SQL_FIND = "SELECT tfcidacl, tfcidacn, tfcidadd, tfcidtfc, " & _
               "       tfcdescr, tfcdtvin, tfcfltit, tfcidade, " & _
               "       tfcidadz, tfcannul, timestmp FROM taritfc0 WHERE "
    SQL_GETLIST = "SELECT tfcidacl, tfcidacn, tfcidadd, tfcidtfc, tfcdescr FROM taritfc0 "
    SQL_UPDATE = "UPDATE taritfc0 SET "
    SQL_INSERT = "INSERT INTO taritfc0 "
    SQL_GET_TS = "SELECT timestmp FROM taritfc0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT tfcidacl, tfcidacn, tfcidtfc, tfcdescr, tfcdtvin, tfcannul FROM taritfc0 "
    
    SQL_FIND_BY_DATE = "SELECT tfcidtfc " & _
                       "  FROM taritfc0 " & _
                       " WHERE (tfcidacl = ? ) and (tfcidacn = ?) and (tfcidadd = ?) AND (tfcdtvin=?) AND (tfcannul=0) "
                       
    SQL_FIND_MAX_DATE_DEST_DEP = "SELECT max(tfcdtvin) " & _
                                 "  FROM taritfc0 " & _
                                 " WHERE (tfcidacl = ? ) and (tfcidacn = ?)  and (tfcidadd = ?) AND (tfcdtvin<=?) and (tfcannul=0)  "
    
'    SQL_FIND_MAX_DATE = "SELECT max(tfcdtvin) " & _
'                        "  FROM taritfc0 WHERE (tfcidacl = ?) and (tfcidacn = ?) and (tfcidtfc = ?) and (tfcannul=0)  " & _
'                        "   AND tfcfltit= ? AND (tfcdtvin<=?) and ((tfcidadz is null) or (LEN(tfcidadz)=0) )  ) "
'    SQL_FIND_MAX_DATE_DEST = "SELECT max(tfcdtvin) " & _
'                             "  FROM taritfc0 WHERE (tfcidade = ?) and (tfcannul=0)  " & _
'                             "   AND tfcfltit= ? AND (tfcdtvin<=?) AND (tfcidadz = ?)  ) "
'
'    SQL_FIND_BY_DATE_2 = "SELECT tfcidtfc, tfcdtvin " & _
'                       "  FROM taritfc0 " & _
'                       " WHERE ( tfcdtvin=? AND tfcannul=0 AND (tfcidacl = ?) and (tfcidacn = ?) and (tfcidtfc = ?) AND tfcfltit= ?  ) and ((tfcidadz is null) or (LEN(tfcidadz)=0)  )"
'    SQL_FIND_BY_DATE_ATS = "SELECT tfcidtfc, tfcdtvin " & _
'                       "  FROM taritfc0 " & _
'                       " WHERE ( tfcdtvin=? AND tfcannul=0 AND (tfcidacl = ?) and (tfcidacn = ?) and (tfcidtfc = ?)  AND tfcfltit= ? AND tfcidadz = ?  )"
'
'    SQL_FIND_MAX_DATE_2 = "SELECT max(tfcdtvin) " & _
'                        "  FROM taritfc0 WHERE (tfcidacl = ?) and (tfcidacn = ?) and (tfcidtfc = ?) and (tfcannul=0)  " & _
'                        "   AND tfcfltit= ? AND (tfcdtvin<=?) and ((tfcidadz is null) or (LEN(tfcidadz)=0) ) "
'    SQL_FIND_MAX_DATE_ATS = "SELECT max(tfcdtvin) " & _
'                             "  FROM taritfc0 WHERE (tfcidacl = ?) and (tfcidacn = ?) and (tfcidtfc = ?) and (tfcannul=0)  " & _
'                             "   AND tfcfltit= ? AND (tfcdtvin<=?) AND (tfcidadz = ?)  "
                       
    SQL_COPY_DETAIL = _
    " INSERT INTO TARITCR0 (TCRIDACL, TCRIDACN, TCRIDADD, TCRIDTFC, TCRPROGD, " & _
    "                       TCRINFKM, TCRSUPKM, " & _
    "                       TCRLIN01, TCRLSU01, TCRTOL01, TCRIDV01, TCRIMB01, " & _
    "                       TCRLIN02, TCRLSU02, TCRTOL02, TCRIDV02, TCRIMB02, " & _
    "                       TCRLIN03, TCRLSU03, TCRTOL03, TCRIDV03, TCRIMB03, " & _
    "                       TCRLIN04, TCRLSU04, TCRTOL04, TCRIDV04, TCRIMB04, " & _
    "                       TCRLIN05, TCRLSU05, TCRTOL05, TCRIDV05, TCRIMB05, " & _
    "                       TCRLIN06, TCRLSU06, TCRTOL06, TCRIDV06, TCRIMB06, " & _
    "                       TCRLIN07, TCRLSU07, TCRTOL07, TCRIDV07, TCRIMB07, " & _
    "                       TCRLIN08, TCRLSU08, TCRTOL08, TCRIDV08, TCRIMB08, " & _
    "                       TCRLIN09, TCRLSU09, TCRTOL09, TCRIDV09, TCRIMB09, " & _
    "                       TCRLIN10, TCRLSU10, TCRTOL10, TCRIDV10, TCRIMB10, " & _
    "                       TCRLIN11, TCRLSU11, TCRTOL11, TCRIDV11, TCRIMB11, " & _
    "                       TCRLIN12, TCRLSU12, TCRTOL12, TCRIDV12, TCRIMB12, " & _
    "                       TCRLIN13, TCRLSU13, TCRTOL13, TCRIDV13, TCRIMB13, " & _
    "                       TCRLIN14, TCRLSU14, TCRTOL14, TCRIDV14, TCRIMB14, " & _
    "                       TCRLIN15, TCRLSU15, TCRTOL15, TCRIDV15, TCRIMB15, " & _
    "                       TCRTIPOP, "
    SQL_COPY_DETAIL = SQL_COPY_DETAIL & _
        "TCRIMP01, TCRIMP02, TCRIMP03, TCRIMP04, TCRIMP05, TCRIMP06, TCRIMP07, TCRIMP08, " & _
        "TCRIMP09, TCRIMP10, TCRIMP11, TCRIMP12, TCRIMP13, TCRIMP14, TCRIMP15, TCRANNUL)"
    
    SQL_COPY_DETAIL_2 = _
    "                            TCRPROGD, TCRINFKM, TCRSUPKM, " & _
    "                       TCRLIN01, TCRLSU01, TCRTOL01, TCRIDV01, TCRIMB01, " & _
    "                       TCRLIN02, TCRLSU02, TCRTOL02, TCRIDV02, TCRIMB02, " & _
    "                       TCRLIN03, TCRLSU03, TCRTOL03, TCRIDV03, TCRIMB03, " & _
    "                       TCRLIN04, TCRLSU04, TCRTOL04, TCRIDV04, TCRIMB04, " & _
    "                       TCRLIN05, TCRLSU05, TCRTOL05, TCRIDV05, TCRIMB05, " & _
    "                       TCRLIN06, TCRLSU06, TCRTOL06, TCRIDV06, TCRIMB06, " & _
    "                       TCRLIN07, TCRLSU07, TCRTOL07, TCRIDV07, TCRIMB07, " & _
    "                       TCRLIN08, TCRLSU08, TCRTOL08, TCRIDV08, TCRIMB08, " & _
    "                       TCRLIN09, TCRLSU09, TCRTOL09, TCRIDV09, TCRIMB09, " & _
    "                       TCRLIN10, TCRLSU10, TCRTOL10, TCRIDV10, TCRIMB10, " & _
    "                       TCRLIN11, TCRLSU11, TCRTOL11, TCRIDV11, TCRIMB11, " & _
    "                       TCRLIN12, TCRLSU12, TCRTOL12, TCRIDV12, TCRIMB12, " & _
    "                       TCRLIN13, TCRLSU13, TCRTOL13, TCRIDV13, TCRIMB13, " & _
    "                       TCRLIN14, TCRLSU14, TCRTOL14, TCRIDV14, TCRIMB14, " & _
    "                       TCRLIN15, TCRLSU15, TCRTOL15, TCRIDV15, TCRIMB15, " & _
    "        TCRTIPOP, "
    
    SQL_COPY_DETAIL_2 = SQL_COPY_DETAIL_2 & _
        "TCRIMP01, TCRIMP02, TCRIMP03, TCRIMP04, TCRIMP05, TCRIMP06, TCRIMP07, TCRIMP08, " & _
        "TCRIMP09, TCRIMP10, TCRIMP11, TCRIMP12, TCRIMP13, TCRIMP14, TCRIMP15, 0" & _
        "   FROM TARITCR0 " & _
        "  WHERE ((tcridacl = ?) and (tcridacn = ?) and (tcridadd = ?) and (tcridtfc = ?))"

    mvarFieldsIndexes = Array("0", "1", "2", "3", "8")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Cliente", "taritfc0.tfcidacl", 0, True, 10, "0"
        With .Item("0")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Contratto", "taritfc0.tfcidacn", 0, True, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ACN"
            .ForeignTable = "anagacn0"
            .ForeignTableFld = "acnidacn"
            .ForeignTableDescField = "acndescr"
            .ForeignTableDeleteFld = "acnannul"
        End With
        .Add "Addebito", "taritfc0.tfcidadd", 0, True, 3, "9"
        With .Item("9")
            .ForeignTableEntity = "ADD"
            .ForeignTable = "anagadd0"
            .ForeignTableFld = "addidadd"
            .ForeignTableDescField = "adddescr"
            .ForeignTableDeleteFld = "addannul"
        End With
        .Add "Codice", "taritfc0.tfcidtfc", 0, True, 7, "2"
        .Add "Descrizione", "taritfc0.tfcdescr", 0, True, 25, "3"
        .Add "Data validità", "taritfc0.tfcdtvin", 2, True, 10, "4"
        .Add "Flag tipo tariffa", "taritfc0.tfcfltit", 0, True, 1, "5"
        .Add "Deposito", "taritfc0.tfcidade", 0, False, 3, "6"
        With .Item("6")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Destinazione", "taritfc0.tfcidadz", 0, False, 10, "7"
        With .Item("7")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Annullato", "taritfc0.tfcannul", 3, , , "8"
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "9", "3"
        .Add "2", "4"
    End With
    
    DS.CancelField = "taritfc0.tfcannul"
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

